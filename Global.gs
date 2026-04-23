// ==========================================
// GLOBAL.GS — Configuración central, enrutamiento y utilidades
// ==========================================

/**
 * Servir HTML — Enrutador principal del Web App
 */
function doGet(e) {
  // CASO 1: Checkin de punto de control (QR de asamblea)
  if (e.parameter.control || e.parameter.action === 'checkin') {
    var template = HtmlService.createTemplateFromFile('QR_Asistencia');
    template.params = e.parameter;
    return template.evaluate()
        .setTitle('Registro Asistencia - Sindicato SLIM n°3')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
        .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no');
  }

  // CASO 2: Vinculación QR personal (action=register, rut=...)
  if (e.parameter.action || e.parameter.rut || e.parameter.asamblea) {
    var template = HtmlService.createTemplateFromFile('QR_Access');
    template.data = e.parameter;
    return template.evaluate()
        .setTitle('Control QR - Sindicato SLIM n°3')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
        .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no');
  }

  // CASO 3: Aplicación principal
  return HtmlService.createHtmlOutputFromFile('Index')
      .setTitle('Sindicato SLIM n°3 - App Socios')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no');
}

// ==========================================
// CONFIGURACIÓN GLOBAL — IDs DE SPREADSHEETS Y CARPETAS
// ==========================================
var CONFIG = {
  SPREADSHEETS: {
    USUARIOS:       "1m7KLd3b3BzKOAI10I5E32MVf_L34XWAGFonhTg37TVM",
    JUSTIFICACIONES:"1Hwbly__MXjl9uwJb-spXdah-R3v9SAMOCFHem92uOUg",
    APELACIONES:    "11nrvVsf84THWQ7j6NfAr_unyIcBV7aykxACS8R27PwE",
    PRESTAMOS:      "1h-_sJD4rOCuMjlfSouP7a6gfoodHyzI4MOBRUyOW5XU",
    PERMISOS_MEDICOS:"1VYfm7cOgL3mVfVoI8DubIm8WG2srzQw9a6DtIEs3UMM",
    CREDENCIALES:   "1HVyPxdYKuvIybeOCAPwAJaVHwlxEuOik4YW0XOXBE5o",
    ASISTENCIA:     "1SRQ8Mlc6bBdb0mitAfn4I-EUAS4BOrZRbqS9YAmg3Sk",
    GAMIFICACION:   "1SHDIhGv6XOc30Epm4vdusp3QGVD-pWzhIwzeD6iqbXQ"
  },
  HOJAS: {
    USUARIOS:               "BD_SLIMAPP",
    JUSTIFICACIONES:        "BD_JUSTIFICACIONES",
    CONFIG_JUSTIFICACIONES: "CONFIG_JUSTIFICACIONES",
    APELACIONES:            "BD_APELACIONES",
    PRESTAMOS:              "BD_PRESTAMOS",
    VALIDACION_PRESTAMOS:   "Validación-Prestamos",
    PERMISOS_MEDICOS:       "BD_Permisos medicos",
    CREDENCIALES:           "IMPRESION",
    HISTORIAL_CREDENCIALES: "HISTORIAL_CREDENCIALES",
    ASISTENCIA:             "BD_ASISTENCIA",
    PUNTOS_CONTROL:         "PUNTOS_CONTROL",
    GAMIFICACION:           "BD_GAMIFICACION",
    BANCO_PREGUNTAS:        "BANCO_PREGUNTAS"
  },
  CARPETAS: {
    JUSTIFICACIONES:             "1UD9hQz1FuacSb3QYrahRl7IfvlpKn8v6",
    APELACIONES_COMPROBANTES:    "15BmK5pf5Txrxdzdrny23S5q35NDxLy4P",
    APELACIONES_LIQUIDACIONES:   "1dR7fM6TW99tunNaMZliyvXc-L23nHKVY",
    APELACIONES_DEVOLUCIONES:    "1LGLKA3fiCJXf2ouIqlxq3jk_ZSxI3IyM",
    PERMISOS_MEDICOS:            "1nCYxD5sJLszBBA6s2DquGW8vlKGZp4ty",
    VESTUARIO_DOCS:              "1A4PVsIn8ndNMXdqnO9GZCovjtNdfr0BI"
  },
  CORREOS: {
    REPRESENTANTE_LEGAL: "juancarlos.pacheco@cl.issworld.com"
  },
  COLUMNAS: {
    USUARIOS: {
      RUT: 0, RUT_VALIDADO: 1, FECHA_INGRESO: 2, NOMBRE: 3, CARGO: 4,
      CORREO: 5, SITE: 6, REGION: 7, SEXO: 8, ESTADO: 9,
      DETALLE_DESVINCULACION: 10, ID_CREDENCIAL: 11, CORREO_REGISTRADO: 12,
      CONTACTO: 13, ROL: 14, LINK_REGISTRO: 15, QR_REGISTRO: 16,
      BANCO: 17, TIPO_CUENTA: 18, NUMERO_CUENTA: 19, ESTADO_NEG_COLECT: 20,
      TALLA_POLERA: 21, TALLA_POLAR: 22, TALLA_PANTALON: 23, TALLA_CALZADO: 24,
      CALZADO_ESPECIAL: 25, URL_CERT_PIE_DIABETICO: 26
    },
    JUSTIFICACIONES: {
      ID: 0, FECHA: 1, RUT: 2, NOMBRE: 3, REGION: 4, MOTIVO: 5,
      ARGUMENTO: 6, RESPALDO: 7, ESTADO: 8, OBSERVACION: 9, NOTIFICACION: 10,
      ASAMBLEA: 11, GESTION: 12, DIRIGENTE: 13, CORREO_DIRIGENTE: 14
    },
    APELACIONES: {
      ID: 0, FECHA_SOLICITUD: 1, RUT: 2, NOMBRE: 3, CORREO: 4,
      MES_APELACION: 5, TIPO_MOTIVO: 6, DETALLE_MOTIVO: 7, URL_COMPROBANTE: 8,
      URL_LIQUIDACION: 9, ESTADO: 10, OBSERVACION: 11, NOTIFICADO: 12,
      GESTION: 13, NOMBRE_DIRIGENTE: 14, CORREO_DIRIGENTE: 15,
      URL_COMPROBANTE_DEVOLUCION: 16, PERMISO_DEVOLUCION: 17, LOG_PERMISOS: 18
    },
    PRESTAMOS: {
      ID: 0, FECHA: 1, RUT: 2, NOMBRE: 3, CORREO: 4, TIPO: 5,
      MONTO: 6, CUOTAS: 7, MEDIO_PAGO: 8, ESTADO: 9, FECHA_TERMINO: 10,
      GESTION: 11, NOMBRE_DIRIGENTE: 12, CORREO_DIRIGENTE: 13, INFORME: 14,
      OBSERVACION: 15
    },
    PERMISOS_MEDICOS: {
      ID: 0, FECHA_SOLICITUD: 1, RUT: 2, NOMBRE: 3, CORREO: 4,
      TIPO_PERMISO: 5, FECHA_INICIO: 6, MOTIVO_DETALLE: 7, URL_DOCUMENTO: 8,
      ESTADO: 9, FECHA_SUBIDA: 10, NOTIFICADO_REP_LEGAL: 11, GESTION: 12,
      NOMBRE_DIRIGENTE: 13, CORREO_DIRIGENTE: 14, NOTIFICADO_SOCIO: 15
    },
    GAMIFICACION: {
      RUT: 0, NOMBRE: 1, XP_TOTAL: 2, GRADO: 3, LOGROS: 4,
      RACHA_ACTUAL: 5, RACHA_MAX: 6, ULTIMA_ACTIVIDAD: 7, QUIZ_ULTIMO_DIA: 8,
      QUIZZES_COMPLETADOS: 9, ESTADO: 10, QUIZZES_PERFECTOS: 11
    },
    BANCO_PREGUNTAS: {
      ID: 0, CATEGORIA: 1, NIVEL: 2, PREGUNTA: 3, OPCION_A: 4, OPCION_B: 5,
      OPCION_C: 6, OPCION_D: 7, RESPUESTA: 8, EXPLICACION: 9, XP: 10,
      ACTIVA: 11, FUENTE: 12
    }
  }
};

// ==========================================
// URL BASE DEL WEB APP
// ==========================================
var WEBAPP_BASE_URL = 'https://script.google.com/a/~/macros/s/AKfycbzrmy_GgdzMpOLfycvxxUPHU6iyuL9Jv6As_4kxG7mG8oQ4RbV-ALUZw0oeSJnqbvvc/exec';

// ==========================================
// HELPERS DE ACCESO A SPREADSHEETS Y HOJAS
// ==========================================

/**
 * Retorna el objeto Spreadsheet para una clave del CONFIG.
 * @param {string} spreadsheetKey
 * @returns {Spreadsheet}
 */
function getSpreadsheet(spreadsheetKey) {
  var spreadsheetId = CONFIG.SPREADSHEETS[spreadsheetKey];
  if (!spreadsheetId) {
    throw new Error('Spreadsheet key "' + spreadsheetKey + '" no encontrado en CONFIG');
  }
  return SpreadsheetApp.openById(spreadsheetId);
}

/**
 * Retorna una hoja específica con manejo de errores.
 * @param {string} spreadsheetKey
 * @param {string} sheetKey
 * @param {boolean} [createIfNotExists=false]
 * @returns {Sheet|null}
 */
function getSheet(spreadsheetKey, sheetKey, createIfNotExists) {
  createIfNotExists = createIfNotExists || false;
  try {
    var ss = getSpreadsheet(spreadsheetKey);
    var sheetName = CONFIG.HOJAS[sheetKey];

    if (!sheetName) {
      console.error('❌ Clave de hoja "' + sheetKey + '" no encontrada en CONFIG.HOJAS');
      return null;
    }

    var sheet = ss.getSheetByName(sheetName);

    if (!sheet && createIfNotExists) {
      console.warn('⚠️ Hoja "' + sheetName + '" no existe. Creándola...');
      sheet = ss.insertSheet(sheetName);
      console.log('✅ Hoja "' + sheetName + '" creada exitosamente');
    }

    if (!sheet) {
      console.error('❌ Hoja "' + sheetName + '" no encontrada en spreadsheet ' + spreadsheetKey);
      return null;
    }

    return sheet;

  } catch (e) {
    console.error('❌ Error obteniendo hoja ' + sheetKey + ' de ' + spreadsheetKey + ': ' + e.toString());
    return null;
  }
}

// ==========================================
// UTILIDADES DE FORMATO
// ==========================================

/**
 * Formatea una fecha a dd/mm/yyyy - hh:mm
 */
function formatearFechaConHora(fecha) {
  try {
    if (!fecha) return "";
    var fechaObj = (typeof fecha === 'string') ? new Date(fecha) : fecha;
    if (isNaN(fechaObj.getTime())) return fecha.toString();
    var dia = String(fechaObj.getDate()).padStart(2, '0');
    var mes = String(fechaObj.getMonth() + 1).padStart(2, '0');
    var anio = fechaObj.getFullYear();
    var hora = String(fechaObj.getHours()).padStart(2, '0');
    var min  = String(fechaObj.getMinutes()).padStart(2, '0');
    return dia + "/" + mes + "/" + anio + " - " + hora + ":" + min;
  } catch (e) {
    Logger.log('Error formateando fecha: ' + e.toString());
    return fecha ? fecha.toString() : "";
  }
}

/**
 * Formatea una fecha a dd/mm/yyyy (sin hora)
 */
function formatearFechaSinHora(fecha) {
  try {
    if (!fecha) return "";
    var fechaObj = (typeof fecha === 'string') ? new Date(fecha) : fecha;
    if (isNaN(fechaObj.getTime())) return fecha.toString();
    var dia = String(fechaObj.getDate()).padStart(2, '0');
    var mes = String(fechaObj.getMonth() + 1).padStart(2, '0');
    var anio = fechaObj.getFullYear();
    return dia + "/" + mes + "/" + anio;
  } catch (e) {
    Logger.log('Error formateando fecha: ' + e.toString());
    return fecha ? fecha.toString() : "";
  }
}

/**
 * Formatea un RUT con puntos y guión para visualización
 */
function formatRutDisplay(rut) {
  if (!rut) return '';
  var cleaned = cleanRut(rut);
  if (cleaned.length < 2) return cleaned;
  var dv = cleaned.slice(-1);
  var numero = cleaned.slice(0, -1);
  var formatted = numero.replace(/\B(?=(\d{3})+(?!\d))/g, '.');
  return formatted + '-' + dv;
}

/**
 * Formatea un RUT para correos (con puntos y guión)
 */
function formatRutServer(rut) {
  if (!rut) return "";
  var rutString = String(rut).trim();
  var value = rutString.replace(/[^0-9kK]/g, '').toUpperCase();
  if (value.length < 2) return value;
  var body = value.slice(0, -1);
  var dv = value.slice(-1);
  var formattedBody = "";
  for (var i = body.length - 1, j = 0; i >= 0; i--, j++) {
    formattedBody = body.charAt(i) + ((j > 0 && j % 3 === 0) ? "." : "") + formattedBody;
  }
  return formattedBody + "-" + dv;
}

/**
 * Limpia un RUT quitando puntos, guión y espacios
 */
function cleanRut(rut) {
  if (!rut) return "";
  return String(rut).replace(/\./g, '').replace(/-/g, '').toUpperCase().trim();
}

/**
 * Valida si un correo electrónico tiene formato válido
 */
function esCorreoValido(correo) {
  if (!correo || typeof correo !== 'string') return false;
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(correo.trim().toLowerCase());
}

/**
 * Ajusta un color hexadecimal más claro/oscuro (uso en plantillas de correo)
 */
function adjustColor(hexColor, percent) {
  var num = parseInt(hexColor.replace("#", ""), 16);
  var amt = Math.round(2.55 * percent);
  var R = (num >> 16) + amt;
  var G = (num >> 8 & 0x00FF) + amt;
  var B = (num & 0x0000FF) + amt;
  return "#" + (0x1000000 +
    (R < 255 ? R < 1 ? 0 : R : 255) * 0x10000 +
    (G < 255 ? G < 1 ? 0 : G : 255) * 0x100 +
    (B < 255 ? B < 1 ? 0 : B : 255))
    .toString(16).slice(1);
}

/**
 * Normaliza un string de hora al formato HH:mm
 */
function normalizarHoraHHmm(valor) {
  if (!valor) return '';
  var limpio = valor.replace(/\s/g, '').toUpperCase();
  var esPM = limpio.indexOf('PM') !== -1;
  var esAM = limpio.indexOf('AM') !== -1;
  limpio = limpio.replace('A.M.', '').replace('P.M.', '').replace('AM', '').replace('PM', '');
  var partes = limpio.split(':');
  if (partes.length < 2) return '';
  var horas   = parseInt(partes[0], 10);
  var minutos = parseInt(partes[1], 10);
  if (isNaN(horas) || isNaN(minutos)) return '';
  if (esAM && horas === 12) horas = 0;
  if (esPM && horas !== 12) horas += 12;
  return ('0' + horas).slice(-2) + ':' + ('0' + minutos).slice(-2);
}

/**
 * Genera código de asamblea en formato YYYY_MM
 */
function generarCodigoAsamblea(fecha) {
  if (!fecha || !(fecha instanceof Date)) fecha = new Date();
  var year  = fecha.getFullYear();
  var month = String(fecha.getMonth() + 1).padStart(2, '0');
  return year + "_" + month;
}

/**
 * Genera código de asamblea desde fecha de evento en formato YYYY_MM_DD
 */
function generarCodigoAsambleaEvento(fechaEvento) {
  try {
    var fecha;
    if (typeof fechaEvento === 'string') {
      var soloFecha = fechaEvento.split('T')[0];
      var partes = soloFecha.split('-');
      fecha = new Date(parseInt(partes[0]), parseInt(partes[1]) - 1, parseInt(partes[2]), 12, 0, 0);
    } else if (fechaEvento instanceof Date) {
      fecha = fechaEvento;
    } else {
      return generarCodigoAsamblea(new Date());
    }
    var year  = fecha.getFullYear();
    var month = String(fecha.getMonth() + 1).padStart(2, '0');
    var day   = String(fecha.getDate()).padStart(2, '0');
    return year + "_" + month + "_" + day;
  } catch (e) {
    Logger.log('Error en generarCodigoAsambleaEvento: ' + e.toString());
    return generarCodigoAsamblea(new Date());
  }
}

/**
 * Extrae la URL de una fórmula =IMAGE("URL")
 */
function extraerUrlDeImagen(formula) {
  if (!formula || typeof formula !== 'string') return '';
  var regex = /=IMAGE\s*\(\s*"([^"]+)"\s*\)/i;
  var match = formula.match(regex);
  if (match && match[1]) return match[1];
  if (formula.startsWith('http')) return formula;
  return '';
}

// ==========================================
// SISTEMA CENTRALIZADO DE PERMISOS DE ARCHIVOS
// ==========================================

/**
 * Valida los correos de los usuarios involucrados antes de procesar archivos
 */
function validarCorreosParaPermisos(beneficiario, gestor, esGestionDirigente) {
  var resultado = {
    valido: true,
    alertas: [],
    correosParaPermisos: [],
    alertaBeneficiario: false,
    alertaGestor: false
  };

  var correoBeneficiarioValido = esCorreoValido(beneficiario.correo);

  if (correoBeneficiarioValido) {
    resultado.correosParaPermisos.push({
      correo: beneficiario.correo.trim().toLowerCase(),
      tipo: 'beneficiario',
      nombre: beneficiario.nombre
    });
  } else {
    resultado.alertaBeneficiario = true;
    if (esGestionDirigente) {
      resultado.alertas.push({
        tipo: 'warning',
        mensaje: 'El socio ' + beneficiario.nombre + ' no tiene un correo electrónico válido registrado. No podrá acceder al archivo adjunto. Infórmele que debe actualizar sus datos en "Mis Datos".'
      });
    } else {
      resultado.alertas.push({
        tipo: 'warning',
        mensaje: 'No tienes un correo electrónico válido registrado. No podrás acceder al archivo adjunto desde tu correo. Por favor, actualiza tus datos en el módulo "Mis Datos".'
      });
    }
  }

  if (esGestionDirigente && gestor) {
    var correoGestorValido = esCorreoValido(gestor.correo);
    if (correoGestorValido) {
      var yaExiste = resultado.correosParaPermisos.some(function(c) {
        return c.correo === gestor.correo.trim().toLowerCase();
      });
      if (!yaExiste) {
        resultado.correosParaPermisos.push({
          correo: gestor.correo.trim().toLowerCase(),
          tipo: 'gestor',
          nombre: gestor.nombre
        });
      }
    } else {
      resultado.alertaGestor = true;
      resultado.alertas.push({
        tipo: 'info',
        mensaje: 'Tu correo electrónico no está registrado correctamente. El archivo se procesará, pero no recibirás acceso directo. Actualiza tus datos en "Mis Datos".'
      });
    }
  }

  return resultado;
}

/**
 * Sube un archivo a Google Drive y otorga permisos de lectura (silenciosos)
 */
function subirArchivoConPermisos(archivoData, carpetaId, nombreArchivo, correosParaPermisos, correosAdicionales) {
  correosAdicionales = correosAdicionales || [];

  var resultado = {
    success: false,
    url: '',
    permisosOtorgados: [],
    permisosError: [],
    mensajeError: ''
  };

  try {
    var sizeInBytes = (archivoData.base64.length * 3) / 4;
    if (sizeInBytes > 15 * 1024 * 1024) {
      resultado.mensajeError = "El archivo es demasiado grande (máximo 15MB).";
      return resultado;
    }

    var folder = DriveApp.getFolderById(carpetaId);
    var blob = Utilities.newBlob(
      Utilities.base64Decode(archivoData.base64),
      archivoData.mimeType,
      archivoData.fileName
    );

    var extension = "";
    var nameParts = archivoData.fileName.split('.');
    if (nameParts.length > 1) extension = "." + nameParts.pop();
    blob.setName(nombreArchivo + extension);

    var file = folder.createFile(blob);
    Utilities.sleep(1500);
    file.setSharing(DriveApp.Access.PRIVATE, DriveApp.Permission.NONE);
    Utilities.sleep(1000);

    var todosLosCorreos = correosParaPermisos.slice();
    correosAdicionales.forEach(function(correo) {
      if (esCorreoValido(correo)) {
        var yaExiste = todosLosCorreos.some(function(c) {
          return c.correo === correo.trim().toLowerCase();
        });
        if (!yaExiste) {
          todosLosCorreos.push({ correo: correo.trim().toLowerCase(), tipo: 'adicional', nombre: 'Usuario adicional' });
        }
      }
    });

    var fileId = file.getId();

    todosLosCorreos.forEach(function(item) {
      try {
        Drive.Permissions.insert(
          { 'role': 'reader', 'type': 'user', 'value': item.correo },
          fileId,
          { sendNotificationEmails: false }
        );
        resultado.permisosOtorgados.push({ correo: item.correo, tipo: item.tipo, nombre: item.nombre });
        Logger.log("✅ Permiso silencioso otorgado a " + item.tipo + ": " + item.correo);
      } catch (permError) {
        Logger.log("⚠️ Fallo API Avanzada para " + item.correo + " - Intentando addViewer...");
        Utilities.sleep(1000);
        try {
          file.addViewer(item.correo);
          resultado.permisosOtorgados.push({ correo: item.correo, tipo: item.tipo, nombre: item.nombre });
          Logger.log("✅ Permiso otorgado via addViewer a " + item.tipo + ": " + item.correo);
        } catch (fallbackError) {
          Utilities.sleep(2000);
          try {
            file.addViewer(item.correo);
            resultado.permisosOtorgados.push({ correo: item.correo, tipo: item.tipo, nombre: item.nombre });
            Logger.log("✅ Permiso otorgado en reintento final para: " + item.correo);
          } catch (finalError) {
            resultado.permisosError.push({ correo: item.correo, tipo: item.tipo, nombre: item.nombre, error: finalError.toString() });
            Logger.log("❌ Error fatal al otorgar permiso a " + item.tipo + " (" + item.correo + "): " + finalError);
          }
        }
      }
    });

    resultado.success = true;
    resultado.url = file.getUrl();
    return resultado;

  } catch (error) {
    Logger.log("❌ Error al subir archivo: " + error.toString());
    resultado.mensajeError = "Error al subir el archivo: " + error.toString();
    return resultado;
  }
}

/**
 * Genera el objeto de alerta de permisos para retornar al frontend
 */
function generarAlertaPermisos(validacionCorreos, resultadoSubida) {
  var alerta = { mostrarAlerta: false, tipoAlerta: 'info', mensajeAlerta: '', detalles: [] };

  if (validacionCorreos.alertas && validacionCorreos.alertas.length > 0) {
    alerta.mostrarAlerta = true;
    validacionCorreos.alertas.forEach(function(a) { alerta.detalles.push(a.mensaje); });
    if (validacionCorreos.alertaBeneficiario) alerta.tipoAlerta = 'warning';
  }

  if (resultadoSubida && resultadoSubida.permisosError && resultadoSubida.permisosError.length > 0) {
    alerta.mostrarAlerta = true;
    alerta.tipoAlerta = 'warning';
    resultadoSubida.permisosError.forEach(function(err) {
      alerta.detalles.push("No se pudo otorgar acceso a " + err.nombre + " (" + err.correo + ")");
    });
  }

  if (alerta.mostrarAlerta) alerta.mensajeAlerta = alerta.detalles.join('\n\n');
  return alerta;
}

// ==========================================
// CORREO ESTILIZADO CENTRALIZADO
// ==========================================

/**
 * Envía un correo HTML estilizado con tabla de detalles
 */
function enviarCorreoEstilizado(destinatario, asunto, titulo, mensaje, detalles, colorTema) {
  try {
    if (!destinatario || !destinatario.includes("@")) {
      console.log("Correo inválido: " + destinatario);
      return;
    }

    var detallesHtml = "";
    if (detalles && typeof detalles === "object") {
      detallesHtml = "<table style='width:100%;border-collapse:separate;border-spacing:0;margin-top:20px;border:1px solid #e2e8f0;border-radius:8px;overflow:hidden;'>";
      var isEven = false;
      for (var key in detalles) {
        var valor = detalles[key];
        if (valor === null || valor === undefined || valor === "") {
          valor = "<span style='color:#94a3b8;font-style:italic;'>S/D</span>";
        }
        var bgRow = isEven ? "#f8fafc" : "#ffffff";
        detallesHtml += "<tr style='background-color:" + bgRow + ";'>" +
          "<td style='padding:12px 15px;border-bottom:1px solid #e2e8f0;color:#64748b;font-weight:600;font-size:13px;width:35%;vertical-align:top;text-transform:uppercase;letter-spacing:0.05em;'>" + key + "</td>" +
          "<td style='padding:12px 15px;border-bottom:1px solid #e2e8f0;color:#1e293b;font-weight:500;font-size:14px;vertical-align:top;'>" + valor + "</td>" +
          "</tr>";
        isEven = !isEven;
      }
      detallesHtml += "</table>";
    }

    var uniqueId = Utilities.getUuid().slice(0, 8);
    var colorOscuro = adjustColor(colorTema, -40);

    var htmlBody = '<!DOCTYPE html><html><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1.0"></head>' +
      '<body style="margin:0;padding:0;font-family:\'Helvetica Neue\',Helvetica,Arial,sans-serif;background-color:#f1f5f9;">' +
      '<div style="max-width:600px;margin:20px auto;background:white;border-radius:16px;overflow:hidden;box-shadow:0 10px 15px -3px rgba(0,0,0,0.1);">' +
      '<div style="background:linear-gradient(135deg,' + colorTema + ' 0%,' + colorOscuro + ' 100%);padding:40px 30px;text-align:center;">' +
      '<h1 style="margin:0;color:white;font-size:24px;font-weight:800;letter-spacing:-0.5px;text-shadow:0 2px 4px rgba(0,0,0,0.1);">' + titulo + '</h1>' +
      '<p style="margin:10px 0 0 0;color:rgba(255,255,255,0.9);font-size:14px;">Sindicato SLIM N°3</p></div>' +
      '<div style="padding:40px 30px;background-color:#ffffff;">' +
      '<p style="color:#334155;font-size:16px;line-height:1.6;margin:0 0 25px 0;text-align:left;">' + mensaje + '</p>' +
      detallesHtml +
      '<div style="margin-top:30px;padding:15px;background-color:#eff6ff;border-left:4px solid ' + colorTema + ';border-radius:4px;">' +
      '<p style="color:#1e40af;font-size:12px;line-height:1.5;margin:0;"><strong>Nota Importante:</strong> Si el campo aparece como "S/D", significa que no hay datos registrados para ese ítem en el momento de la gestión.</p>' +
      '</div></div>' +
      '<div style="background:#f8fafc;padding:20px;text-align:center;border-top:1px solid #e2e8f0;">' +
      '<p style="color:#64748b;font-size:11px;margin:0;line-height:1.4;">Este es un mensaje automático. Por favor no respondas a este correo.<br>© ' + new Date().getFullYear() + ' Plataforma de Gestión Sindicato SLIM N°3</p>' +
      '<p style="color:#cbd5e1;font-size:9px;margin:10px 0 0 0;">Ref: ' + uniqueId + '</p>' +
      '</div></div></body></html>';

    MailApp.sendEmail({ to: destinatario, subject: asunto, htmlBody: htmlBody });

  } catch (e) {
    console.error("Error enviando correo a " + destinatario + ": " + e.toString());
  }
}

// ==========================================
// VERIFICACIÓN DE ROL DE USUARIO
// ==========================================

/**
 * Verifica si un usuario tiene un rol específico
 * @param {string} rut
 * @param {Array} rolesPermitidos
 * @returns {Object} {autorizado, mensaje, rol}
 */
function verificarRolUsuario(rut, rolesPermitidos) {
  try {
    var usuario = obtenerUsuarioPorRut(rut);
    if (!usuario.encontrado) {
      return { autorizado: false, mensaje: "Usuario no encontrado", rol: "" };
    }
    var rolUsuario = String(usuario.rol || "SOCIO").trim().toUpperCase();
    var tienePermiso = rolesPermitidos.some(function(rol) {
      return rol.toUpperCase() === rolUsuario;
    });
    if (!tienePermiso) {
      Logger.log('⚠️ INTENTO DE ACCESO NO AUTORIZADO: RUT=' + rut + ' Rol=' + rolUsuario + ' Requeridos=' + rolesPermitidos.join(', '));
      return { autorizado: false, mensaje: "No tienes permisos para realizar esta acción", rol: rolUsuario };
    }
    return { autorizado: true, mensaje: "Acceso autorizado", rol: rolUsuario };
  } catch (e) {
    Logger.log('❌ Error verificando rol: ' + e.toString());
    return { autorizado: false, mensaje: "Error de validación", rol: "" };
  }
}

// ==========================================
// ESTADOS SWITCHES PARA DASHBOARD (badges)
// ==========================================

/**
 * Retorna el estado habilitado/deshabilitado de todos los módulos en una sola llamada
 */
function obtenerEstadosSwitchDashboard() {
  try {
    var props = PropertiesService.getScriptProperties();
    var prestamos     = (props.getProperty('prestamos_habilitado')         !== 'false');
    var contrato      = (props.getProperty('contrato_colectivo_habilitado') !== 'false');
    var slimquest     = (props.getProperty('slimquest_habilitado')          !== 'false');
    var calculadora   = (props.getProperty('calculadora_habilitada')        !== 'false');
    var permisosMedicos = (props.getProperty('permisos_medicos_habilitado') !== 'false');
    var asistencia    = (props.getProperty('asistencia_habilitada')         !== 'false');
    var apelaciones   = (props.getProperty('apelaciones_habilitado')        !== 'false');

    var justificaciones = false;
    try {
      var resJ = obtenerEstadoSwitchJustificaciones();
      justificaciones = resJ.habilitado;
    } catch (eJ) { justificaciones = false; }

    return {
      success: true,
      prestamos:       prestamos,
      justificaciones: justificaciones,
      contrato:        contrato,
      slimquest:       slimquest,
      calculadora:     calculadora,
      permisosMedicos: permisosMedicos,
      asistencia:      asistencia,
      apelaciones:     apelaciones
    };
  } catch (e) {
    Logger.log('Error en obtenerEstadosSwitchDashboard: ' + e.toString());
    return { success: false };
  }
}
