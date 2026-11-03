# Automatización de compras y cumplimiento SAGRILAFT

[![Verificar código](https://github.com/Dsergio909/sagrilaft-compras-automation/actions/workflows/verificar.yml/badge.svg)](https://github.com/Dsergio909/sagrilaft-compras-automation/actions/workflows/verificar.yml)
![Google Apps Script](https://img.shields.io/badge/Google_Apps_Script-4285F4?logo=googleappsscript&logoColor=white)
![Herramientas de pago](https://img.shields.io/badge/herramientas_de_pago-0-2e7d32)

**Español** · [English](README_EN.md)

Sistema para pedir, rastrear y auditar los documentos que los proveedores deben mantener al día por la norma **SAGRILAFT** (Colombia), y para avisar de contratos por vencer. Funciona solo con Google Workspace: Sheets, Gmail, Drive y Calendar, conectados con Apps Script. Sin Zapier, Make ni APIs de pago.

Lo construí por iniciativa propia en mis prácticas profesionales en un área de compras. Este repositorio tiene el código con una configuración de ejemplo: no incluye datos de la empresa, de sus proveedores ni de su personal.

## Qué hace

- **Envía las solicitudes solo.** Martes y viernes revisa la base de proveedores y escribe a los que están a 60 días o menos de su fecha SAGRILAFT, con el formulario adjunto y la lista de documentos según sea persona jurídica o natural.
- **Lee Gmail para saber quién respondió**, con varios niveles de certeza (ver abajo). Lo que no es seguro nunca se marca como resuelto: queda un enlace para que una persona lo revise con un clic.
- **Detecta correos rebotados** y los marca para corregir la dirección.
- **Guarda una bitácora que no se borra.** Cada acción agrega una línea con fecha y responsable; nada se sobrescribe y la hoja está protegida.
- **Ordena el seguimiento manual.** Llamada, WhatsApp o recordatorio quedan registrados con responsable y fecha; el próximo seguimiento se calcula en días hábiles (con los festivos de Colombia) y los casos sin movimiento se detectan solos.
- **Arma tableros dentro de la misma hoja:** índice de cumplimiento, proveedores por actualizar y plan de acción priorizado.
- **Contratos:** un segundo módulo avisa al equipo, en un solo correo, de los contratos vencidos y de los que están por vencer.

## Cómo detecta las respuestas

| Capa | Cómo busca | Certeza | Qué hace |
|---|---|---|---|
| 0 | Rebote de `mailer-daemon` que menciona el correo del proveedor | Alta | Marca "correo rebotado" |
| 1 | ID del hilo guardado al enviar | Máxima | Actualiza el estado |
| 2 | Destinatario y asunto exactos | Alta | Actualiza el estado y repara el enlace si faltaba |
| 3a/3b/3c | Remitente, dominio de la empresa o NIT en el texto | Media o baja | **No toca el estado:** deja un enlace para revisión humana |

## Arquitectura

```
Google Sheet
 ├─ BASE DE PROVEEDORES COMPRAS   datos maestros
 ├─ Tabla Control                 estado de cada solicitud enviada
 ├─ Bitácora Auditoría            oculta y protegida, solo se le agregan filas
 ├─ DASHBOARD SAGRILAFT           se regenera desde el menú
 ├─ Contratos                     matriz de contratos
 └─ Control Alertas Contratos     registro para no repetir avisos

src/ (Apps Script vinculado a la hoja)
 ├─ Config.gs        correos, IDs, firma y documentos: lo único que se edita
 ├─ Proveedores.gs   envío, plantilla, detección, tablero y escalamiento
 ├─ Contratos.gs     alertas y tablero de contratos
 ├─ Menu.gs          menús de la barra
 └─ appsscript.json  zona horaria y motor V8
```

## Decisiones de diseño

- **Ante la duda, decide una persona.** Una confirmación equivocada marca "al día" a un proveedor que no lo está; una revisión de más cuesta unos segundos. Por eso las capas heurísticas solo sugieren.
- **El cumplimiento se calcula con fechas**, no con una columna de estado que se llena a mano y se desactualiza.
- **El historial se acumula.** Cada nota del sistema se agrega a las anteriores, con fecha y enlace clicable.
- **Columnas por nombre, no por número.** El tablero busca cada encabezado; si alguien inserta una columna, el sistema avisa en vez de leer datos equivocados en silencio.
- **Si falla una dependencia, el trabajo sigue.** Si no se puede leer el calendario de festivos, se asume día hábil en vez de frenar el seguimiento.
- **Toda la configuración en un archivo.** Para cambiar correos, documentos o firma no hay que tocar la lógica.

## Instalación

1. En una hoja de Google con las pestañas de arriba, abre **Extensiones → Apps Script** y copia los archivos de `src/`. Para pegar `appsscript.json`, activa "Mostrar el archivo de manifiesto" en la configuración del proyecto.
2. Completa `Config.gs` con tus correos, los IDs de Drive del formulario y de la firma, y la lista de documentos.
3. Recarga la hoja y usa el menú **🚀 Control Proveedores → ⚙️ Instalación**:
   1. Crear Tabla Control y Bitácora.
   2. Programar envío diario.
   3. Programar detección de casos sin seguimiento.
4. Acepta los permisos que pide Google (Sheets, Gmail, Drive y Calendar).

## Verificación

`node tools/verificar.js` revisa la sintaxis de cada archivo, que no haya dos funciones con el mismo nombre (en Apps Script todos los archivos comparten espacio y una pisaría a la otra) y que cada opción del menú y cada trigger apunten a una función que existe. Corre en GitHub Actions con cada cambio.

## Limitaciones

- Gmail devuelve como máximo 500 hilos por búsqueda.
- El sistema solo ve el buzón de la cuenta que lo ejecuta: si un proveedor responde a otra dirección sin copiarla, no hay forma de detectarlo desde el código.

## Autor

**Sergio García** · Negocios Internacionales, Bogotá · [Portafolio](https://dsergio909.github.io/) · [LinkedIn](https://linkedin.com/in/sergio-david-garcia-celis-836b8a250)
