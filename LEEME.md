# Supplier Risk Radar: cómo ejecutarlo

## Dashboard diario (localhost)

- **Windows:** doble clic en `iniciar_dashboard.bat`.
- **Mac/Linux:** `./iniciar_dashboard.sh`.
- **A mano:** `pip install -r requirements.txt` y después `python dashboard.py`.

Se abre en http://localhost:5000. Para usar otro puerto: `python dashboard.py --puerto 8080`.

Cada día, el equipo de compras entra en el dashboard y pulsa **Subir noticias** con el Excel del día. El Excel necesita una hoja con las columnas News_ID, Published_Date, Headline y Summary.

- **Histórico:** cada fichero se suma al histórico guardado en `datos_dashboard/`. El radar recuerda las alertas de días anteriores: un alternativo que ya ha caído deja de proteger, y las noticias de "vuelta a la normalidad" cierran las alertas.
- **Selector de fecha:** permite ver el estado de cualquier día pasado.
- **Datos de ejemplo:** en el menú ⋯ → *Cargar datos de ejemplo* se cargan train, validación, test y Live Day desde la carpeta `ejemplos/`. El Answer Key nunca se usa.
- **Fichero maestro:** por defecto se usa `ejemplos/Supplier_Risk_1_Training.xlsx`, con las hojas Suppliers y Materials. Se puede cambiar desde el menú ⋯.

Qué muestra el dashboard:

| Sección | Para qué sirve |
|---|---|
| Resumen instantáneo (arriba) | En 5 segundos: rojo, ámbar o verde; materiales sin respaldo, impacto estimado en €, proveedores interrumpidos y noticias urgentes. Al pulsar un problema se salta a su fila |
| Filtros (barra lateral) | Día, mostrar "con problemas" o "todo", situación del material, categoría, material, producto, solo fuente única, buscador y rango de la tendencia |
| Materiales y servicios | Si cada material sigue cubierto, la noticia que lo causa, los días que quedan, el inventario y el **impacto económico**. Al pulsar una fila se ve el detalle y el cálculo. El botón **Solucionar** cierra la alerta (se puede deshacer en la barra lateral) |
| Proveedores | Los proveedores con alertas activas (o todos), agrupados por categoría. Al pulsar uno se abre su ficha con el historial |
| Riesgos por día y noticias del día | Tendencia diaria y cada noticia con su explicación |
| ⚙ Ajustes | Avisos por correo, supuestos del impacto económico y datos (CSV del día, ejemplos, maestro, tema, borrar histórico) |

Estados de un proveedor:

- **Interrumpido:** tiene una alerta urgente, o un evento que para la producción o la entrega (incendio, huelga, desastre, regulador, ciberataque, escasez…) con severidad 4 o más.
- **En riesgo:** tiene una alerta activa no grave. Sigue entregando, pero hay que vigilarlo.
- **Operativo:** no tiene alertas activas.

Estados de un material (o servicio):

- **Sin respaldo:** el primario y el alternativo están interrumpidos, o la fuente única está interrumpida. Es la **alerta de máxima importancia**.
- **Cubrir con el alternativo:** el primario está caído y el alternativo funciona.
- **Plan B caído:** el primario funciona, pero el alternativo está caído.
- **Vigilar:** algún proveedor tiene una alerta leve.
- **Operativo.**

Cuánto dura una alerta:

- Si la noticia da una duración, se usa esa (entre 7 y 120 días).
- Si no la da: 21 días para las urgentes, 14 para las medias y 7 para las leves.
- Una alerta se cierra antes si llega una noticia de resolución o si se pulsa **Solucionar**.

### Impacto económico (estimación para priorizar)

- **Sin respaldo:** el inventario se agota en *inicio del problema + días de cobertura*. Desde ese día hasta el fin previsto no llega material: días sin material × gasto anual / 365 × **valor de producción por € de material** (por defecto 8).
- **Cubrir con el alternativo:** días restantes × gasto diario del primario × **sobrecoste de compra urgente** (por defecto 15 %).
- **Plan B caído o Vigilar:** 0 €, porque el material sigue llegando.

Los dos supuestos se cambian en ⚙ Ajustes. Los servicios y la logística no tienen inventario, así que el impacto empieza el primer día.

### Avisos por correo

1. En ⚙ Ajustes, escribe la dirección de correo y marca **Enviar automáticamente**.
2. En **Servidor de correo (SMTP)**, elige Gmail u Outlook (o tu servidor) y pon usuario y contraseña. En Gmail hace falta una *contraseña de aplicación* (Cuenta de Google → Seguridad → Verificación en dos pasos → Contraseñas de aplicaciones).
3. Pulsa **Enviar correo de prueba**.

Cada vez que se suben noticias y aparece un material **sin respaldo** nuevo, se envía un correo con el material, los proveedores, la noticia, los productos, la fecha prevista de fin, el impacto y la acción recomendada. Cada alerta se avisa una sola vez. Si el envío falla, el correo se guarda en `datos_dashboard/correos_no_enviados/` y el error aparece en "Últimos envíos".

## Línea de comandos (sin web)

Requisitos: Python 3.9+ con `pandas`, `openpyxl` y `scikit-learn`.

    pip install pandas openpyxl scikit-learn

## Comandos

Validación (compara con las etiquetas del fichero):

    python radar_riesgo_proveedores.py --maestro Supplier_Risk_1_Training.xlsx --noticias Supplier_Risk_2_Validation.xlsx

Test (rellena las columnas Is_Risk, Urgent y Affected_Supplier_IDs):

    python radar_riesgo_proveedores.py --maestro Supplier_Risk_1_Training.xlsx --noticias Supplier_Risk_3_Test.xlsx --historial Supplier_Risk_1_Training.xlsx Supplier_Risk_2_Validation.xlsx

Live Day (sin tocar código; la memoria usa los días anteriores):

    python radar_riesgo_proveedores.py --maestro Supplier_Risk_1_Training.xlsx --noticias Supplier_Risk_LiveDay_Feed.xlsx --historial Supplier_Risk_1_Training.xlsx Supplier_Risk_2_Validation.xlsx Supplier_Risk_3_Test.xlsx

Opcional: `--clave Supplier_Risk_Answer_Key.xlsx` puntúa con la guía de los moderadores, `--salida carpeta` cambia la carpeta de salida y `--sin-ml` desactiva el clasificador de texto.

## Qué genera (carpeta `resultados/`)

Por cada fichero de noticias se generan dos salidas.

- **`<fichero>_RESULTADOS.xlsx`** es una copia del Excel original con las tres columnas de respuesta rellenas. Conserva los desplegables. Los urgentes van en rojo y los riesgos en ámbar. Añade estas hojas:
  - **Radar_Detalle:** tipo de evento, severidad, puntuación, cómo se enlazó el proveedor, acción del playbook y una explicación en lenguaje natural.
  - **Insights:** concentración geográfica, alternativos caídos y alertas cerradas.
  - **Evaluacion:** solo si hay etiquetas o se pasa `--clave`.
- **`<fichero>_respuestas.csv`** es el export de 4 columnas que se pide el día de la presentación.

## Cómo decide

1. **Proveedor afectado.** Primero busca el nombre: el nombre completo, las dos primeras palabras, o una palabra más el material del proveedor. Así se evitan trampas como "Celtic" (fútbol), "Lumina" (película) o "Stellar Shield". Si no hay nombre y el evento es de zona (desastre, huelga portuaria, corte eléctrico, logística o exportación), busca por ubicación: ciudad, región, país o corredor marítimo.
2. **¿Es un riesgo?** Hace falta un proveedor enlazado y un evento adverso detectado con el léxico de dominio. Las noticias positivas, las de "vuelta a la normalidad" y las que dicen "sin impacto en el suministro" no cuentan.
3. **¿Es urgente?** La puntuación es severidad + dependencia, y es urgente si llega a 4 o más.
   - **Severidad:** de 1 a 5 según el tipo de evento, con +1 si dura semanas o es indefinido y −1 si es corto, parcial, un aviso o una parada preventiva.
   - **Dependencia:** +1 si es fuente única de criticidad alta, −1 si tiene un alternativo sano y +1 si la parada supera la cobertura de inventario.
4. **Memoria.** Si el alternativo ya sufre una interrupción larga, deja de proteger. Las noticias de resolución cierran las alertas anteriores.
5. **Clasificador TF-IDF + regresión logística.** Se entrena solo con train y da una segunda opinión cuando se nombra un proveedor pero ningún evento del léxico encaja.
