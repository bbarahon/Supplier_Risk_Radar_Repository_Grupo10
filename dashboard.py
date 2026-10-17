#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Supplier Risk Radar - Dashboard diario (localhost)
==================================================

    pip install -r requirements.txt
    python dashboard.py                 # abre http://localhost:5000

Cada Excel de noticias que se sube es un "día" más del histórico. El radar
reprocesa todo el histórico en orden cronológico (así funciona la memoria entre
días) y el dashboard muestra, para la fecha elegida:
  * un resumen instantáneo: ¿hay problemas hoy?
  * materiales y servicios: si siguen cubiertos, el impacto económico estimado
    por los días que se espera que dure el problema, y un botón "Solucionar"
  * estado de los 40 proveedores, tendencia y noticias del día
  * avisos por correo cuando hay una alerta de máxima importancia
    (proveedor de fuente única interrumpido o material sin respaldo)
"""
from __future__ import annotations

import argparse
import glob
import io
import json
import os
import shutil
import smtplib
import ssl
import threading
import webbrowser
from datetime import datetime, timedelta
from email.message import EmailMessage

import pandas as pd
from flask import Flask, Response, jsonify, render_template, request

import radar_riesgo_proveedores as rr

BASE = os.path.dirname(os.path.abspath(__file__))
DATA = os.path.join(BASE, "datos_dashboard")
UPLOADS = os.path.join(DATA, "noticias")
OUTBOX = os.path.join(DATA, "correos_no_enviados")
STATE_FILE = os.path.join(DATA, "estado.json")
SMTP_FILE = os.path.join(DATA, "smtp.json")
EXAMPLES = os.path.join(BASE, "ejemplos")
DEFAULT_MASTER = os.path.join(BASE, "ejemplos", "Supplier_Risk_1_Training.xlsx")

# Supuestos del impacto económico (editables desde la web)
DEFAULT_ASSUMPTIONS = {
    "multiplicador_produccion": 8.0,   # € de venta perdida por cada € de material que no llega
    "sobrecoste_alternativo": 0.15,    # +15 % al comprar de urgencia al alternativo
}

ACTION_ES = {
    "Monitor": "Vigilar",
    "Increase safety stock": "Aumentar stock de seguridad",
    "Shift volume to approved alternate": "Pasar volumen al alternativo aprobado",
    "Start dual sourcing": "Iniciar doble abastecimiento",
    "Diversify geographically": "Diversificar geográficamente",
    "No action": "Sin acción",
}
EVENT_ES = {
    "Fire / plant incident": "Incendio / accidente en planta",
    "Quality / contamination": "Calidad / contaminación",
    "Regulatory action": "Acción regulatoria",
    "Export restriction": "Restricción a la exportación",
    "Financial distress": "Dificultades financieras",
    "Cyber attack": "Ciberataque",
    "Natural disaster": "Desastre natural",
    "Labour strike": "Huelga",
    "Utility outage": "Corte de suministro eléctrico",
    "Logistics disruption": "Disrupción logística",
    "Supply shortage / delay": "Escasez / retraso",
    "Acquisition by competitor": "Adquisición por competidor",
    "Leadership change": "Cambio de dirección",
    "Other adverse news (ML)": "Otra noticia adversa (ML)",
    "Resolution / good news": "Resolución / vuelta a la normalidad",
    "Positive / neutral news": "Noticia positiva / neutra",
    "Not relevant": "No relevante",
}
# eventos que paran físicamente la producción o la entrega
STOPPING_EVENTS = {"Fire / plant incident", "Quality / contamination", "Regulatory action",
                   "Export restriction", "Cyber attack", "Natural disaster", "Labour strike",
                   "Supply shortage / delay", "Utility outage"}
LEVEL_ORDER = {"critical": 0, "warning": 1, "watch": 2, "good": 3}

app = Flask(__name__, template_folder=os.path.join(BASE, "templates"))
_lock = threading.Lock()
_cache: dict = {}


# ---------------------------------------------------------------------------
# Persistencia mínima (carpeta + JSON)
# ---------------------------------------------------------------------------

def load_meta():
    meta = {}
    if os.path.exists(STATE_FILE):
        with open(STATE_FILE, encoding="utf-8") as f:
            meta = json.load(f)
    meta.setdefault("master", None)
    meta.setdefault("files", [])
    meta.setdefault("manual_closures", {})
    meta.setdefault("assumptions", dict(DEFAULT_ASSUMPTIONS))
    meta.setdefault("email", {"to": "", "enabled": False})
    meta.setdefault("email_log", [])
    meta.setdefault("notified", [])
    return meta


def save_meta(meta):
    os.makedirs(DATA, exist_ok=True)
    with open(STATE_FILE, "w", encoding="utf-8") as f:
        json.dump(meta, f, ensure_ascii=False, indent=2)


def load_smtp():
    cfg = {"host": "", "port": 587, "user": "", "password": "", "security": "starttls", "sender": ""}
    if os.path.exists(SMTP_FILE):
        with open(SMTP_FILE, encoding="utf-8") as f:
            cfg.update(json.load(f))
    if os.environ.get("RADAR_SMTP_PASSWORD"):
        cfg["password"] = os.environ["RADAR_SMTP_PASSWORD"]
    return cfg


def master_path(meta):
    p = meta.get("master")
    if p and os.path.exists(os.path.join(DATA, p)):
        return os.path.join(DATA, p)
    return DEFAULT_MASTER if os.path.exists(DEFAULT_MASTER) else None


# ---------------------------------------------------------------------------
# Motor: reprocesa el histórico
# ---------------------------------------------------------------------------

def run_engine(meta):
    key = json.dumps([meta.get("master"), [f["stored"] for f in meta["files"]]])
    if _cache.get("key") == key:
        return _cache["val"]
    mpath = master_path(meta)
    if not mpath:
        raise RuntimeError("Falta el fichero maestro (Suppliers + Materials).")
    sups, prods = rr.load_master(mpath)
    tm = rr.train_text_model(mpath)
    mat = pd.read_excel(mpath, "Materials")

    frames = []
    for i, f in enumerate(meta["files"]):
        _, df = rr.find_news_sheet(os.path.join(UPLOADS, f["stored"]))
        df = df.copy()
        df["_file"] = f["name"]
        df["_order"] = i
        frames.append(df)
    if frames:
        news = pd.concat(frames, ignore_index=True).drop_duplicates("News_ID", keep="last")
        keep = [c for c in ["News_ID", "Published_Date", "Source", "Headline", "Summary", "_file", "_order"] if c in news]
        news = news[keep].reset_index(drop=True)
        for c in ("Source", "Headline", "Summary"):
            news[c] = news[c].fillna("").astype(str) if c in news else ""
        news["Published_Date"] = pd.to_datetime(news["Published_Date"], errors="coerce")
        news["Published_Date"] = news["Published_Date"].fillna(pd.Timestamp(datetime.now().date()))
        news = news.sort_values(["Published_Date", "_order"]).reset_index(drop=True)
        radar = rr.RiskRadar(sups, prods, tm)
        pred = radar.run(news.drop(columns=["_order"]))
        res = news.merge(pred, on="News_ID")
        insights = radar.insights
    else:
        res, insights = pd.DataFrame(), []
    val = dict(sups=sups, prods=prods, mat=mat, res=res, insights=insights)
    _cache.update(key=key, val=val)
    return val


def alert_window(row) -> int:
    """Días que se espera que dure el problema (según la noticia o, si no lo dice, por gravedad)."""
    d = row.get("Duration_Days")
    if d is not None and not pd.isna(d) and d > 0:
        return int(max(7, min(120, d)))
    if row.get("Urgent") == "Y":
        return 21
    return 7 if row.get("Severity", 0) <= 2 else 14


# ---------------------------------------------------------------------------
# Impacto económico
# ---------------------------------------------------------------------------

def material_impact(row, sup_out, day_ts, assumptions):
    """
    Impacto estimado mientras dure el problema (desde hoy hasta la fecha prevista de fin):
      * sin respaldo -> días que el inventario NO cubre × gasto diario del material × multiplicador
      * cubrir con alternativo -> días × gasto diario del primario × sobrecoste de compra urgente
      * resto -> 0 (riesgo vigilado, sin coste directo)
    """
    mult = float(assumptions.get("multiplicador_produccion", 8))
    prem = float(assumptions.get("sobrecoste_alternativo", 0.15))
    p, a = sup_out.get(row["primary"]), sup_out.get(row["alternate"]) if row["alternate"] else None
    down = [s for s in (p, a) if s and s["status"] == "critical"]
    alerts = [al for s in (p, a) if s for al in s["active"]]
    out = dict(cost=0.0, days=0, uncovered_days=0, start=None, end=None, news=[], basis="")
    if not alerts:
        return out
    main = sorted(alerts, key=lambda al: (al["urgent"], al["severity"], al["until"]), reverse=True)[0]
    out["news"] = sorted({al["news_id"] for al in alerts})
    start = min(pd.Timestamp(al["date"]) for al in (down and [x for s in down for x in s["active"]] or alerts))
    end = max(pd.Timestamp(al["until"]) for al in (down and [x for s in down for x in s["active"]] or alerts))
    days_left = max(0, (end - day_ts).days)
    out.update(start=start.strftime("%Y-%m-%d"), end=end.strftime("%Y-%m-%d"),
               days=days_left, total_days=(end - start).days, main_news=main["news_id"],
               main_event=main["event"])
    if row["level"] == "critical":
        spend_day = sum(s["spend"] for s in (p, a) if s) / 365
        cover = (p["cover_days"] or 0) if p else 0
        stock_out = start + timedelta(days=cover)          # día en que se acaba el inventario
        uncovered = max(0, (end - max(day_ts, stock_out)).days)
        out["uncovered_days"] = uncovered
        out["cost"] = round(uncovered * spend_day * mult)
        out["basis"] = (f"El inventario ({int(cover)} d) se agota el {stock_out:%d/%m}; "
                        f"{uncovered} días sin material × {spend_day:,.0f} €/día de compra × {mult:g} "
                        f"(valor de producción por € de material)").replace(",", ".")
        if uncovered == 0:
            out["basis"] = (f"El inventario ({int(cover)} d) cubre el problema hasta el {end:%d/%m}: "
                            f"sin pérdida de producción si se resuelve a tiempo")
    elif row["level"] == "warning" and p and p["status"] == "critical":
        spend_day = p["spend"] / 365
        out["cost"] = round(days_left * spend_day * prem)
        out["basis"] = (f"{days_left} días comprando al alternativo × {spend_day:,.0f} €/día × "
                        f"{prem:.0%} de sobrecoste por compra urgente").replace(",", ".")
    else:
        out["basis"] = "Sin coste directo previsto: el material sigue llegando. Vigilar."
    return out


# ---------------------------------------------------------------------------
# Estado a una fecha
# ---------------------------------------------------------------------------

def build_state(as_of: str | None, meta=None):
    meta = meta or load_meta()
    eng = run_engine(meta)
    sups, prods, mat, res = eng["sups"], eng["prods"], eng["mat"], eng["res"]

    dates = sorted(res["Published_Date"].dt.strftime("%Y-%m-%d").unique().tolist()) if len(res) else []
    day = as_of if as_of in dates else (dates[-1] if dates else None)
    day_ts = pd.Timestamp(day) if day else None
    upto = res[res["Published_Date"] <= day_ts] if day else res

    # ---- alertas por proveedor
    alerts = {sid: [] for sid in sups}
    closures = {}
    for _, r in upto.iterrows():
        d = r["Published_Date"]
        for sid in rr.ids_set(r.get("Closed_Supplier_IDs")):
            closures[sid] = max(closures.get(sid, d), d)
        if r["Is_Risk"] != "Y":
            continue
        for sid in rr.ids_set(r["Affected_Supplier_IDs"]):
            if sid in alerts:
                alerts[sid].append(r)
    manual = {}
    for sid, d in meta.get("manual_closures", {}).items():
        d = pd.Timestamp(d)
        if day_ts is None or d <= day_ts:
            closures[sid] = max(closures.get(sid, d), d)
            manual[sid] = d.strftime("%Y-%m-%d")

    sup_out = {}
    for sid, s in sups.items():
        hist, active = [], []
        for r in alerts[sid]:
            d = r["Published_Date"]
            until = d + timedelta(days=alert_window(r))
            closed = sid in closures and closures[sid] >= d
            is_active = (not closed) and day_ts is not None and until >= day_ts
            item = dict(news_id=r["News_ID"], date=d.strftime("%Y-%m-%d"), headline=r["Headline"],
                        event=EVENT_ES.get(r["Event_Type"], r["Event_Type"]), urgent=r["Urgent"] == "Y",
                        severity=int(r["Severity"]), score=int(r["Urgency_Score"]),
                        action=ACTION_ES.get(r["Mitigation_Action"], r["Mitigation_Action"]),
                        explanation=r["Explanation"], until=until.strftime("%Y-%m-%d"),
                        stops=r["Event_Type"] in STOPPING_EVENTS, active=is_active, closed=closed)
            hist.append(item)
            if is_active:
                active.append(item)
        if any(a["urgent"] or (a["severity"] >= 4 and a["stops"]) for a in active):
            status = "critical"
        elif active:
            status = "warning"
        else:
            status = "good"
        sup_out[sid] = dict(
            id=sid, name=s.name, category=s.category, material=s.material, country=s.country,
            city=s.city, single_source=s.single_source, alternate=s.alternate,
            cover_days=s.cover_days, criticality=s.criticality, spend=s.spend,
            products=sorted(prods.get(sid, [])), status=status, active=active,
            history=sorted(hist, key=lambda x: x["date"], reverse=True),
            manual_closure=manual.get(sid))

    # ---- filas de materiales (+ servicios críticos que no están en la hoja Materials)
    rows = []
    covered = set()
    for _, m in mat.iterrows():
        alt = m.get("Alternate_Supplier_ID")
        rows.append(dict(id=str(m["Material_ID"]), material=str(m["Material"]), category=str(m["Category"]),
                         products=[x.strip() for x in str(m["Used_In_Products"]).split(";") if x.strip()],
                         primary=str(m["Primary_Supplier_ID"]),
                         alternate=str(alt) if isinstance(alt, str) else None))
        covered.update({str(m["Primary_Supplier_ID"]), str(alt)})
    n = 0
    for sid, s in sups.items():
        if sid in covered:
            continue
        n += 1
        rows.append(dict(id=f"SV{n:02d}", material=s.material, category=s.category,
                         products=["Todos los productos"] if s.category in ("Services", "Logistics") else [],
                         primary=sid, alternate=s.alternate))
        covered.update({sid, s.alternate})

    materials = []
    assumptions = meta.get("assumptions", DEFAULT_ASSUMPTIONS)
    for r in rows:
        p, a = sup_out.get(r["primary"]), sup_out.get(r["alternate"]) if r["alternate"] else None
        p_down = p is not None and p["status"] == "critical"
        a_down = a is None or a["status"] == "critical"
        if p_down and a_down:
            cov, lvl = ("Sin respaldo: primario y alternativo caídos" if a else
                        "Sin respaldo: fuente única caída"), "critical"
        elif p_down:
            cov, lvl = "Cubrir con el alternativo", "warning"
        elif a is not None and a["status"] == "critical":
            cov, lvl = "Plan B caído", "warning"
        elif (p and p["status"] == "warning") or (a is not None and a["status"] == "warning"):
            cov, lvl = "Vigilar", "watch"
        else:
            cov, lvl = "Operativo", "good"
        r.update(primary_name=p["name"] if p else "", primary_status=p["status"] if p else "good",
                 alternate_name=a["name"] if a else "", alternate_status=a["status"] if a else None,
                 single_source=a is None, cover_days=p["cover_days"] if p else None,
                 criticality=p["criticality"] if p else "", coverage=cov, level=lvl,
                 affected=[s["id"] for s in (p, a) if s and s["status"] != "good"])
        r["impact"] = material_impact(r, sup_out, day_ts, assumptions) if day_ts is not None else {"cost": 0}
        alerts_here = [al for s in (p, a) if s for al in s["active"]]
        best = sorted(alerts_here, key=lambda al: (al["urgent"], al["score"]), reverse=True)
        r["action"] = best[0]["action"] if best else ""
        r["explanation"] = best[0]["explanation"] if best else ""
        r["headline"] = best[0]["headline"] if best else ""
        materials.append(r)
    materials.sort(key=lambda x: (LEVEL_ORDER[x["level"]], -x["impact"].get("cost", 0), x["id"]))

    # ---- noticias del día y serie temporal
    today = res[res["Published_Date"] == day_ts] if day else res.iloc[0:0]
    news_today = [dict(news_id=r["News_ID"], source=r.get("Source", ""), headline=r["Headline"],
                       is_risk=r["Is_Risk"], urgent=r["Urgent"], suppliers=r["Affected_Supplier_IDs"],
                       event=EVENT_ES.get(r["Event_Type"], r["Event_Type"]),
                       action=ACTION_ES.get(r["Mitigation_Action"], r["Mitigation_Action"]),
                       explanation=r["Explanation"])
                  for _, r in (today.sort_values(["Urgent", "Is_Risk"], ascending=False).iterrows()
                               if len(today) else [])]
    timeline = []
    for d in dates:
        dd = res[res["Published_Date"] == pd.Timestamp(d)]
        timeline.append(dict(date=d, total=len(dd), risk=int(((dd.Is_Risk == "Y") & (dd.Urgent == "N")).sum()),
                             urgent=int((dd.Urgent == "Y").sum())))

    summary = build_summary(sup_out, materials, today)
    insights = [i for i in dict.fromkeys(eng["insights"]) if i.split(":")[0] in set(today.get("News_ID", []))]
    return dict(
        day=day, dates=dates, summary=summary, suppliers=list(sup_out.values()), materials=materials,
        news=news_today, timeline=timeline, insights=insights,
        categories=sorted({m["category"] for m in materials}),
        products=sorted({p for m in materials for p in m["products"]}),
        files=[dict(name=f["name"], uploaded=f["uploaded"]) for f in meta["files"]],
        total_news=len(res), master=os.path.basename(master_path(meta) or ""),
        manual_closures=manual, assumptions=assumptions,
        examples=bool(glob.glob(os.path.join(EXAMPLES, "*.xlsx"))))


def max_alerts(sup_out, materials):
    """Alertas de máxima importancia: fuente única interrumpida o material sin respaldo."""
    out = []
    for m in materials:
        if m["level"] == "critical":
            news = m["impact"].get("main_news", "")
            out.append(dict(key=f"{m['id']}|{news}", material=m["material"], products=m["products"],
                            suppliers=m["affected"], news=news, event=m["impact"].get("main_event", ""),
                            headline=m["headline"], cost=m["impact"].get("cost", 0),
                            end=m["impact"].get("end"), action=m["action"],
                            why=m["coverage"]))
    return out


def build_summary(sup_out, materials, today):
    crit_mat = [m for m in materials if m["level"] == "critical"]
    warn_mat = [m for m in materials if m["level"] == "warning"]
    watch_mat = [m for m in materials if m["level"] == "watch"]
    down = [s for s in sup_out.values() if s["status"] == "critical"]
    urgent_today = int((today.Urgent == "Y").sum()) if len(today) else 0
    cost = sum(m["impact"].get("cost", 0) for m in materials)
    products = sorted({p for m in crit_mat for p in m["products"]})
    if crit_mat:
        level = "critical"
        title = (f"{len(crit_mat)} material{'es' if len(crit_mat) > 1 else ''} sin respaldo hoy")
        sub = "Hay que actuar hoy: " + ", ".join(m["material"] for m in crit_mat[:3]) + ("…" if len(crit_mat) > 3 else "")
    elif warn_mat:
        level = "warning"
        title = f"{len(warn_mat)} material{'es' if len(warn_mat) > 1 else ''} dependen del plan B"
        sub = "Sin roturas de suministro, pero con proveedores caídos."
    elif watch_mat:
        level = "watch"
        title = "Sin problemas críticos"
        sub = f"{len(watch_mat)} material{'es' if len(watch_mat) > 1 else ''} a vigilar."
    else:
        level = "good"
        title = "Todo en orden"
        sub = "Todos los materiales tienen proveedor operativo."
    top = [dict(id=m["id"], material=m["material"], level=m["level"], coverage=m["coverage"],
                cost=m["impact"].get("cost", 0), end=m["impact"].get("end"),
                event=m["impact"].get("main_event", ""), products=m["products"])
           for m in (crit_mat + warn_mat)[:4]]
    return dict(level=level, title=title, sub=sub, cost=cost, no_backup=len(crit_mat),
                plan_b=len(warn_mat), down=len(down), urgent_today=urgent_today,
                products=products, top=top)


# ---------------------------------------------------------------------------
# Correo
# ---------------------------------------------------------------------------

def _eur(x):
    return f"{x:,.0f} €".replace(",", ".")


def compose_email(to, day, items, smtp):
    msg = EmailMessage()
    n = len(items)
    msg["Subject"] = f"[Supplier Risk Radar] ALERTA MÁXIMA {day}: {n} material{'es' if n > 1 else ''} sin respaldo"
    msg["From"] = smtp.get("sender") or smtp.get("user") or "radar@localhost"
    msg["To"] = to
    lines = [f"Supplier Risk Radar · {day}", "",
             "Se han detectado alertas de máxima importancia (sin proveedor operativo):", ""]
    html = [f"<h2 style='color:#a42727;margin:0 0 8px'>Alerta máxima · {day}</h2>",
            "<p>Se han detectado materiales <b>sin proveedor operativo</b>:</p><ul>"]
    for it in items:
        cost = _eur(it["cost"]) if it["cost"] else "sin pérdida si se resuelve antes de agotar el inventario"
        lines += [f"- {it['material']} ({it['why']})",
                  f"  Proveedores: {', '.join(it['suppliers'])} · Noticia {it['news']}: {it['headline']}",
                  f"  Productos: {', '.join(it['products']) or '—'}",
                  f"  Fin previsto: {it['end']} · Impacto estimado: {cost}",
                  f"  Acción recomendada: {it['action']}", ""]
        html.append(f"<li style='margin-bottom:10px'><b>{it['material']}</b> — {it['why']}<br>"
                    f"Proveedores: {', '.join(it['suppliers'])} · {it['news']}: {it['headline']}<br>"
                    f"Productos: {', '.join(it['products']) or '—'}<br>"
                    f"Fin previsto: {it['end']} · Impacto estimado: <b>{cost}</b><br>"
                    f"Acción recomendada: <b>{it['action']}</b></li>")
    lines.append("Abre el dashboard para ver el detalle: http://localhost:5000")
    html.append("</ul><p>Abre el dashboard para ver el detalle: "
                "<a href='http://localhost:5000'>http://localhost:5000</a></p>")
    msg.set_content("\n".join(lines))
    msg.add_alternative("<div style='font-family:Arial,sans-serif;font-size:14px'>" + "".join(html) + "</div>",
                        subtype="html")
    return msg


def send_email(msg, smtp):
    if not smtp.get("host"):
        raise RuntimeError("Servidor SMTP sin configurar")
    port = int(smtp.get("port") or 587)
    sec = smtp.get("security", "starttls")
    if sec == "ssl":
        server = smtplib.SMTP_SSL(smtp["host"], port, timeout=15, context=ssl.create_default_context())
    else:
        server = smtplib.SMTP(smtp["host"], port, timeout=15)
    with server:
        if sec == "starttls":
            server.starttls(context=ssl.create_default_context())
        if smtp.get("user"):
            server.login(smtp["user"], smtp.get("password", ""))
        server.send_message(msg)


def deliver(meta, to, day, items, kind):
    """Envía (o guarda en la bandeja de salida si falla) y registra el resultado."""
    smtp = load_smtp()
    msg = compose_email(to, day, items, smtp)
    entry = dict(when=f"{datetime.now():%Y-%m-%d %H:%M}", to=to, day=day, kind=kind,
                 items=[it["material"] for it in items])
    try:
        send_email(msg, smtp)
        entry["status"] = "enviado"
    except Exception as e:  # noqa: BLE001
        os.makedirs(OUTBOX, exist_ok=True)
        path = os.path.join(OUTBOX, f"{datetime.now():%Y%m%d_%H%M%S}_alerta_{day}.eml")
        with open(path, "wb") as f:
            f.write(bytes(msg))
        entry.update(status="no enviado", error=str(e), saved=os.path.relpath(path, BASE))
    meta["email_log"] = ([entry] + meta.get("email_log", []))[:30]
    return entry


def notify_if_needed(meta):
    """Tras subir noticias: si hay alertas máximas nuevas en el último día, avisa por correo."""
    em = meta.get("email", {})
    if not (em.get("enabled") and em.get("to")):
        return None
    st = build_state(None, meta)
    sup_index = {s["id"]: s for s in st["suppliers"]}
    items = [it for it in max_alerts(sup_index, st["materials"]) if it["key"] not in meta["notified"]]
    if not items:
        return None
    entry = deliver(meta, em["to"], st["day"], items, "automático")
    if entry["status"] == "enviado":
        meta["notified"] = (meta["notified"] + [it["key"] for it in items])[-500:]
    return entry


# ---------------------------------------------------------------------------
# Rutas
# ---------------------------------------------------------------------------

def _clean(o):
    """NaN/inf -> null, para que el navegador siempre reciba JSON válido."""
    if isinstance(o, float) and (o != o or o in (float("inf"), float("-inf"))):
        return None
    if isinstance(o, dict):
        return {k: _clean(v) for k, v in o.items()}
    if isinstance(o, (list, tuple)):
        return [_clean(v) for v in o]
    return o


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/api/state")
def api_state():
    try:
        with _lock:
            st = build_state(request.args.get("dia"))
        return jsonify(_clean(st))
    except Exception as e:  # noqa: BLE001
        return jsonify(error=str(e)), 500


def _store(fs, meta):
    os.makedirs(UPLOADS, exist_ok=True)
    stored = f"{datetime.now():%Y%m%d_%H%M%S_%f}_{os.path.basename(fs.filename)}"
    path = os.path.join(UPLOADS, stored)
    fs.save(path)
    try:
        rr.find_news_sheet(path)
    except Exception as e:  # noqa: BLE001
        os.remove(path)
        raise ValueError(f"{fs.filename}: {e}")
    meta["files"].append(dict(name=os.path.basename(fs.filename), stored=stored,
                              uploaded=f"{datetime.now():%Y-%m-%d %H:%M}"))


@app.route("/api/upload", methods=["POST"])
def api_upload():
    files = request.files.getlist("files")
    if not files:
        return jsonify(error="No se ha recibido ningún fichero."), 400
    with _lock:
        meta = load_meta()
        errors = []
        for fs in files:
            if not fs.filename.lower().endswith((".xlsx", ".xlsm")):
                errors.append(f"{fs.filename}: solo se aceptan ficheros .xlsx")
                continue
            try:
                _store(fs, meta)
            except ValueError as e:
                errors.append(str(e))
        email = notify_if_needed(meta)
        save_meta(meta)
    if errors and len(errors) == len(files):
        return jsonify(error=" · ".join(errors)), 400
    return jsonify(ok=True, warnings=errors, email=email)


@app.route("/api/examples", methods=["POST"])
def api_examples():
    """Carga los Excel de la carpeta ejemplos/ como histórico (train, val, test, live)."""
    paths = sorted(glob.glob(os.path.join(EXAMPLES, "*.xlsx")))
    with _lock:
        meta = load_meta()
        os.makedirs(UPLOADS, exist_ok=True)
        for p in paths:
            if "answer" in os.path.basename(p).lower():
                continue   # nunca usamos las respuestas de los moderadores
            try:
                rr.find_news_sheet(p)
            except Exception:  # noqa: BLE001
                continue
            stored = f"{datetime.now():%Y%m%d_%H%M%S_%f}_{os.path.basename(p)}"
            shutil.copy(p, os.path.join(UPLOADS, stored))
            meta["files"].append(dict(name=os.path.basename(p), stored=stored,
                                      uploaded=f"{datetime.now():%Y-%m-%d %H:%M}"))
        email = notify_if_needed(meta)
        save_meta(meta)
    return jsonify(ok=True, email=email)


@app.route("/api/master", methods=["POST"])
def api_master():
    fs = request.files.get("file")
    if not fs:
        return jsonify(error="No se ha recibido el fichero maestro."), 400
    os.makedirs(DATA, exist_ok=True)
    path = os.path.join(DATA, "maestro.xlsx")
    fs.save(path)
    try:
        rr.load_master(path)
    except Exception as e:  # noqa: BLE001
        return jsonify(error=f"El maestro debe tener las hojas Suppliers y Materials ({e})."), 400
    with _lock:
        meta = load_meta()
        meta["master"] = "maestro.xlsx"
        save_meta(meta)
        _cache.clear()
    return jsonify(ok=True)


@app.route("/api/solve", methods=["POST"])
def api_solve():
    """'Solucionar': cierra a mano las alertas de los proveedores indicados en la fecha dada."""
    body = request.get_json(force=True)
    sids, day = body.get("suppliers", []), body.get("dia")
    with _lock:
        meta = load_meta()
        for sid in sids:
            if body.get("undo"):
                meta["manual_closures"].pop(sid, None)
            else:
                meta["manual_closures"][sid] = day
        save_meta(meta)
    return jsonify(ok=True)


@app.route("/api/settings", methods=["GET", "POST"])
def api_settings():
    with _lock:
        meta = load_meta()
        smtp = load_smtp()
        if request.method == "POST":
            b = request.get_json(force=True)
            if "email" in b:
                meta["email"] = dict(to=str(b["email"].get("to", "")).strip(),
                                     enabled=bool(b["email"].get("enabled")))
            if "assumptions" in b:
                for k in DEFAULT_ASSUMPTIONS:
                    if k in b["assumptions"]:
                        meta["assumptions"][k] = float(b["assumptions"][k])
            if "smtp" in b:
                new = {k: b["smtp"].get(k, smtp.get(k)) for k in ("host", "port", "user", "security", "sender")}
                if b["smtp"].get("password"):
                    new["password"] = b["smtp"]["password"]
                elif os.path.exists(SMTP_FILE):
                    new["password"] = json.load(open(SMTP_FILE, encoding="utf-8")).get("password", "")
                os.makedirs(DATA, exist_ok=True)
                with open(SMTP_FILE, "w", encoding="utf-8") as f:
                    json.dump(new, f, indent=2)
                smtp = load_smtp()
            save_meta(meta)
    safe = {k: v for k, v in smtp.items() if k != "password"}
    safe["has_password"] = bool(smtp.get("password"))
    return jsonify(email=meta["email"], assumptions=meta["assumptions"], smtp=safe,
                   log=meta.get("email_log", []))


@app.route("/api/email/send", methods=["POST"])
def api_email_send():
    """Envía ahora el resumen de alertas máximas del día elegido (o un correo de prueba)."""
    body = request.get_json(force=True) or {}
    with _lock:
        meta = load_meta()
        to = (body.get("to") or meta["email"].get("to") or "").strip()
        if not to:
            return jsonify(error="Escribe primero la dirección de correo."), 400
        st = build_state(body.get("dia"), meta)
        items = max_alerts({s["id"]: s for s in st["suppliers"]}, st["materials"])
        if body.get("test") or not items:
            items = items or [dict(material="(correo de prueba)", why="no hay alertas máximas este día",
                                   suppliers=[], news="-", headline="-", products=[], end=st["day"],
                                   cost=0, action="-")]
        entry = deliver(meta, to, st["day"], items, "prueba" if body.get("test") else "manual")
        save_meta(meta)
    return jsonify(entry)


@app.route("/api/reset", methods=["POST"])
def api_reset():
    with _lock:
        meta = load_meta()
        shutil.rmtree(UPLOADS, ignore_errors=True)
        meta.update(files=[], manual_closures={}, notified=[])
        save_meta(meta)
        _cache.clear()
    return jsonify(ok=True)


@app.route("/api/export")
def api_export():
    """CSV del día (formato de entrega: News_ID, Is_Risk, Urgent, Affected_Supplier_IDs)."""
    with _lock:
        st = build_state(request.args.get("dia"))
    df = pd.DataFrame([dict(News_ID=n["news_id"], Is_Risk=n["is_risk"], Urgent=n["urgent"],
                            Affected_Supplier_IDs=n["suppliers"], Event_Type=n["event"],
                            Accion=n["action"], Explicacion=n["explanation"]) for n in st["news"]])
    buf = io.StringIO()
    df.to_csv(buf, index=False)
    return Response("﻿" + buf.getvalue(), mimetype="text/csv",
                    headers={"Content-Disposition": f"attachment; filename=radar_{st['day']}.csv"})


def main():
    ap = argparse.ArgumentParser(description="Dashboard diario del Supplier Risk Radar")
    ap.add_argument("--puerto", type=int, default=5000)
    ap.add_argument("--no-abrir", action="store_true", help="no abrir el navegador")
    a = ap.parse_args()
    url = f"http://localhost:{a.puerto}"
    print(f"\n  Supplier Risk Radar → {url}\n  (Ctrl+C para parar)\n")
    if not a.no_abrir:
        threading.Timer(1.2, lambda: webbrowser.open(url)).start()
    app.run(host="127.0.0.1", port=a.puerto, debug=False)


if __name__ == "__main__":
    main()
