// =========================================================================
// CONFIG.gs — CONFIGURACIÓN CENTRAL
// =========================================================================
// Todo lo que cambia de una organización a otra vive aquí: correos, IDs de
// Drive, firma, documentos que se piden y textos que marcan el fin de la
// lista. La lógica de los módulos no tiene ningún dato escrito a mano.
//
// Los valores de este archivo son de ejemplo. Complétalos en el proyecto
// de Apps Script vinculado a tu hoja; no subas a un repositorio público la
// versión con datos reales.
// =========================================================================

var CONFIG = {

  // --- Equipo de compras: va en copia de cada solicitud ---
  CORREOS_EQUIPO_COMPRAS: "compras@ejemplo.com",

  // --- Copias adicionales según la clasificación del proveedor ---
  // Si la clasificación contiene la etiqueta (por ejemplo "SEDE NORTE"),
  // se agrega ese correo en copia. Solo aplica a arrendamientos y
  // administraciones, que gestiona cada sede.
  CC_POR_SEDE: {
    "SEDE NORTE": "sede-norte@ejemplo.com",
    "SEDE SUR": "sede-sur@ejemplo.com"
  },

  // --- Dominios propios: sirven para distinguir respuestas internas de externas ---
  DOMINIOS_INTERNOS: ["ejemplo.com"],

  // --- Fin de la lista de proveedores activos ---
  // Si el nombre de una fila contiene alguno de estos textos, el envío se
  // detiene ahí: desde esa fila empieza la sección de proveedores
  // inactivos o inhabilitados, que ya no reciben solicitudes.
  MARCADORES_FIN_DE_LISTA: [],

  // --- Archivos en Drive ---
  ID_GIF_FIRMA_CONTRATOS: "ID_DEL_GIF_DE_FIRMA",
  ID_GIF_FIRMA_PROVEEDORES: "ID_DEL_GIF_DE_FIRMA",
  ID_PDF_FORMATO_PROVEEDOR: "ID_DEL_PDF_DEL_FORMULARIO",  // se adjunta a cada solicitud

  // --- Hoja externa de calificaciones (Módulo extra) ---
  URL_FORMULARIO_CALIFICACIONES: "URL_DE_LA_HOJA_DE_CALIFICACIONES",

  // --- Firma del correo ---
  FIRMA: {
    nombre: "Nombre Apellido",
    cargo: "Analista de Compras",
    correo: "compras@ejemplo.com",
    direccion: "Dirección de la empresa",
    telefono: "Teléfono de contacto",
    slogan: "Frase de la empresa"
  },

  NOMBRE_EMPRESA: "Empresa de Ejemplo S.A.S.",

  // --- Documentos que se piden a cada proveedor ---
  // {anioCerrado} se reemplaza por el último año fiscal cerrado.
  DOCUMENTOS: {
    juridica: [
      "<strong>Formulario de registro de proveedor</strong> (adjunto), diligenciado y firmado.",
      "<strong>RUT</strong> completo, generado este año.",
      "<strong>Certificado de existencia y representación legal</strong>, con menos de 90 días.",
      "<strong>Cédula del representante legal.</strong>",
      "<strong>Certificación bancaria</strong> a nombre de la empresa, con menos de 90 días.",
      "<strong>Estados financieros</strong> con corte a {anioCerrado}.",
      "<strong>Certificado de composición accionaria</strong> de este año."
    ],
    natural: [
      "<strong>Formulario de registro de proveedor</strong> (adjunto), diligenciado y firmado.",
      "<strong>Cédula de ciudadanía.</strong>",
      "<strong>RUT</strong> completo, generado este año.",
      "<strong>Certificación bancaria</strong> a nombre del titular, con menos de 90 días."
    ]
  },

  // --- Calendario público de festivos de Colombia ---
  ID_CALENDARIO_FESTIVOS_CO: "en.co#holiday@group.v.calendar.google.com"
};
