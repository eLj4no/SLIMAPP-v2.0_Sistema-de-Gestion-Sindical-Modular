// ==========================================
// MODULO_ASISTENCIA.GS — Registro de asistencia QR y virtual
// ==========================================

// ==========================================
// REGISTRO PRESENCIAL (QR de punto de control)
// ==========================================

/**
 * Registra la asistencia de un socio vía QR de punto de control.
 * VERSIÓN OPTIMIZADA: Usuario con caché + lock reducido + correo delegado al trigger 20:00
 * BD_ASISTENCIA columnas: FECHA_HORA(A), RUT(B), NOMBRE(C), ASAMBLEA(D), TIPO_ASISTENCIA(E), GESTION(F), CODIGO_TEMP(G), NOTIF_CORREO(H)
 */
function registrarAsistencia(rutInput, nombreControl) {
  var rutLimpio = cleanRut(rutInput);
  if (!rutLimpio) return { success: false, message: "RUT inválido." };

  // Búsqueda de usuario CON CACHÉ, FUERA del lock
  var usuario = obtenerUsuarioPorRut(rutInput);

  // Validación ventana horaria
  try {
    var ssAsistVentana = getSpreadsheet('ASISTENCIA');
    var sheetPCtrl = ssAsistVentana.getSheetByName(CONFIG.HOJAS.PUNTOS_CONTROL);
    if (sheetPCtrl && sheetPCtrl.getLastRow() > 1) {
      var datosPC = sheetPCtrl.getDataRange().getDisplayValues();
      for (var pc = 1; pc < datosPC.length; pc++) {
        if (String(datosPC[pc][0]).trim() === nombreControl) {
          var horaApertura = normalizarHoraHHmm(String(datosPC[pc][4] || '').trim());
          var horaCierre   = normalizarHoraHHmm(String(datosPC[pc][5] || '').trim());
          if (horaApertura && horaCierre) {
            var horaActual   = Utilities.formatDate(new Date(), 'America/Santiago', 'HH:mm');
            var minActual    = parseInt(horaActual.split(':')[0], 10) * 60   + parseInt(horaActual.split(':')[1], 10);
            var minApertura  = parseInt(horaApertura.split(':')[0], 10) * 60 + parseInt(horaApertura.split(':')[1], 10);
            var minCierre    = parseInt(horaCierre.split(':')[0], 10) * 60   + parseInt(horaCierre.split(':')[1], 10);
            if (minActual < minApertura) {
              return { success: false, ventanaCerrada: true, tipoVentana: 'aun_no_abre', horaApertura: horaApertura, horaCierre: horaCierre, message: 'El registro de asistencia aun no ha comenzado. El modulo abre a las ' + horaApertura + ' hrs.' };
            }
            if (minActual > minCierre) {
              return { success: false, ventanaCerrada: true, tipoVentana: 'ya_cerro', horaApertura: horaApertura, horaCierre: horaCierre, message: 'El registro de asistencia ha cerrado. El periodo de registro fue de ' + horaApertura + ' a ' + horaCierre + ' hrs.' };
            }
          }
          break;
        }
      }
    }
  } catch (eVentana) {
    Logger.log('Advertencia: error verificando ventana horaria: ' + eVentana.toString());
  }

  var lock = LockService.getScriptLock();
  if (lock.tryLock(30000)) {
    try {
      var ssAsistencia = getSpreadsheet('ASISTENCIA');
      var sheetAsistencia = ssAsistencia.getSheetByName(CONFIG.HOJAS.ASISTENCIA);

      if (!sheetAsistencia) {
        sheetAsistencia = ssAsistencia.insertSheet(CONFIG.HOJAS.ASISTENCIA);
        sheetAsistencia.appendRow(["FECHA_HORA", "RUT", "NOMBRE", "ASAMBLEA", "TIPO_ASISTENCIA", "GESTION", "CODIGO_TEMP", "NOTIF_CORREO"]);
      }

      var dataAsistencia = sheetAsistencia.getDataRange().getDisplayValues();
      for (var i = 1; i < dataAsistencia.length; i++) {
        var row = dataAsistencia[i];
        if (cleanRut(row[1]) === rutLimpio && row[3] === nombreControl) {
          return { success: false, yaRegistrado: true, message: "Ya registraste tu asistencia en este punto de control." };
        }
      }

      var fechaStr = Utilities.formatDate(new Date(), 'America/Santiago', 'dd/MM/yyyy HH:mm:ss');
      sheetAsistencia.appendRow([
        fechaStr,
        usuario.rut,
        usuario.nombre,
        nombreControl,
        "Asistencia QR",
        "Sistema",
        "",
        ""
      ]);

      return {
        success: true,
        nombre: usuario.nombre,
        rut: usuario.rut,
        fecha: fechaStr,
        asamblea: nombreControl,
        correoEnviado: false,
        mensajeCorreo: (usuario.correo && usuario.correo.includes("@"))
          ? "Recibirás una confirmación en tu correo a más tardar esta noche."
          : "No tienes correo registrado. Puedes ver tu historial en el módulo 'Registro Asistencia'."
      };

    } catch (e) {
      Logger.log("❌ Error en registrarAsistencia: " + e.toString());
      return { success: false, message: "Error del servidor: " + e.toString() };
    } finally {
      lock.releaseLock();
    }
  } else {
    return { success: false, message: "Sistema ocupado, intenta nuevamente en unos segundos." };
  }
}

/**
 * Obtiene el historial de asistencias del usuario.
 */
function obtenerHistorialAsistencia(rutInput) {
  try {
    var rutLimpio = cleanRut(rutInput);
    if (!rutLimpio) return { success: false, message: "RUT inválido." };

    var ssAsistencia = getSpreadsheet('ASISTENCIA');
    var sheet = ssAsistencia.getSheetByName(CONFIG.HOJAS.ASISTENCIA);
    if (!sheet) return { success: true, registros: [] };

    var data = sheet.getDataRange().getDisplayValues();
    var registros = [];

    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (cleanRut(row[1]) === rutLimpio) {
        registros.push({
          fecha:    row[0] || "",
          asamblea: row[3] || "Asamblea",
          tipo:     row[4] || "Asistencia QR",
          gestion:  row[5] || "Sistema",
          dirigente: ""
        });
      }
    }

    registros.reverse();
    return { success: true, registros: registros };

  } catch (e) {
    Logger.log("❌ Error en obtenerHistorialAsistencia: " + e.toString());
    return { success: false, message: "Error: " + e.toString() };
  }
}

// ==========================================
// TRIGGER NOTIFICACIONES ASISTENCIA (diario 20:00)
// ==========================================

/**
 * Trigger diario a las 20:00 hrs.
 * Verifica registros en BD_ASISTENCIA sin notificación enviada (columna H vacía),
 * busca el correo del socio en BD_SLIMAPP y envía la notificación correspondiente.
 */
function verificarNotificacionesAsistencia() {
  try {
    Logger.log('🔔 Iniciando verificación de notificaciones pendientes de asistencia...');

    var ssAsistencia = getSpreadsheet('ASISTENCIA');
    var sheet = ssAsistencia.getSheetByName(CONFIG.HOJAS.ASISTENCIA);
    if (!sheet) { Logger.log('⚠️ Hoja BD_ASISTENCIA no encontrada.'); return; }

    var data = sheet.getDataRange().getDisplayValues();
    if (data.length < 2) { Logger.log('ℹ️ No hay registros en BD_ASISTENCIA.'); return; }

    var sheetUsuarios = getSheet('USUARIOS', 'USUARIOS');
    var dataUsuarios  = sheetUsuarios.getDataRange().getDisplayValues();
    var COL_U = CONFIG.COLUMNAS.USUARIOS;
    var mapaCorreos = {};
    for (var i = 1; i < dataUsuarios.length; i++) {
      var rutU = cleanRut(dataUsuarios[i][COL_U.RUT]);
      if (rutU) mapaCorreos[rutU] = dataUsuarios[i][COL_U.CORREO] || "";
    }

    var enviados = 0, sinCorreo = 0, omitidos = 0;

    for (var j = 1; j < data.length; j++) {
      var row = data[j];
      var notifEstado = String(row[7] || "").trim();

      if (notifEstado !== "") { omitidos++; continue; }

      var rutFila     = cleanRut(row[1]);
      var nombreFila  = row[2] || "Socio";
      var fechaFila   = row[0] || "";
      var asambleaFila= row[3] || "";
      var tipoFila    = row[4] || "Asistencia";
      var correoUsuario = mapaCorreos[rutFila] || "";

      if (correoUsuario && correoUsuario.includes("@")) {
        try {
          enviarCorreoEstilizado(
            correoUsuario,
            "Registro de Asistencia - Sindicato SLIM n°3",
            "Asistencia Registrada",
            "Tu asistencia ha sido registrada en el sistema del sindicato.",
            { "Nombre": nombreFila, "Asamblea": asambleaFila, "Tipo": tipoFila, "Fecha/Hora": fechaFila },
            "#10b981"
          );
          sheet.getRange(j + 1, 8).setValue("ENVIADO");
          enviados++;
          Logger.log("✅ Notificación enviada a " + correoUsuario + " (RUT: " + rutFila + ")");
        } catch (emailErr) {
          Logger.log("⚠️ Error enviando correo a " + correoUsuario + ": " + emailErr.toString());
        }
      } else {
        sheet.getRange(j + 1, 8).setValue("SIN CORREO");
        sinCorreo++;
        Logger.log("ℹ️ Sin correo registrado para RUT: " + rutFila);
      }
    }

    Logger.log("📊 Resultado: " + enviados + " enviados, " + sinCorreo + " sin correo, " + omitidos + " ya procesados.");

  } catch (e) {
    Logger.log("❌ Error en verificarNotificacionesAsistencia: " + e.toString());
  }
}

// ==========================================
// CHECK-IN QR (para QR_Asistencia.html)
// ==========================================

function checkinQR(rutInput, nombreAsamblea) {
  var lock = LockService.getScriptLock();
  if (lock.tryLock(30000)) {
    try {
      var rutLimpio = cleanRut(rutInput);
      var sheetUsers = getSheet('USUARIOS', 'USUARIOS');
      var dataUsers  = sheetUsers.getDataRange().getDisplayValues();
      var COL = CONFIG.COLUMNAS.USUARIOS;

      var usuario = null;
      for (var i = 1; i < dataUsers.length; i++) {
        if (cleanRut(dataUsers[i][COL.RUT]) === rutLimpio) {
          usuario = { rut: dataUsers[i][COL.RUT], nombre: dataUsers[i][COL.NOMBRE] };
          break;
        }
      }
      if (!usuario) throw new Error("Usuario no encontrado");

      var ssAsistencia = getSpreadsheet('ASISTENCIA');
      var sheetAsistencia = ssAsistencia.getSheetByName(CONFIG.HOJAS.ASISTENCIA);
      if (!sheetAsistencia) {
        sheetAsistencia = ssAsistencia.insertSheet(CONFIG.HOJAS.ASISTENCIA);
        sheetAsistencia.appendRow(["FECHA_HORA", "RUT", "NOMBRE", "ASAMBLEA", "TIPO_ASISTENCIA", "GESTION", "CODIGO_TEMP"]);
      }

      var dataAsistencia = sheetAsistencia.getDataRange().getDisplayValues();
      for (var j = 1; j < dataAsistencia.length; j++) {
        var row = dataAsistencia[j];
        if (cleanRut(row[1]) === rutLimpio && row[3] === nombreAsamblea) {
          throw new Error("Ya registraste tu asistencia en esta asamblea.");
        }
      }

      var fechaHora = new Date();
      var fechaStr  = Utilities.formatDate(fechaHora, 'America/Santiago', 'dd/MM/yyyy HH:mm:ss');
      sheetAsistencia.appendRow([fechaStr, usuario.rut, usuario.nombre, nombreAsamblea, "Asistencia QR", "Sistema", ""]);

      return { success: true, nombre: usuario.nombre, rut: usuario.rut, fecha: fechaStr };

    } catch (e) {
      throw new Error(e.message || e.toString());
    } finally {
      lock.releaseLock();
    }
  } else {
    throw new Error("Sistema ocupado, intenta nuevamente.");
  }
}

// ==========================================
// REGISTRO VIRTUAL (sin lock, appendRow atómico)
// ==========================================

/**
 * Registra asistencia virtual SIN lock.
 * Para asambleas virtuales con alta concurrencia.
 * appendRow es atómico en Sheets — seguro sin lock para este caso de uso.
 */
function registrarAsistenciaVirtual(rutInput, nombreControl) {
  try {
    var rutLimpio = cleanRut(rutInput);
    if (!rutLimpio) return { success: false, message: 'RUT inválido.' };

    var usuario = obtenerUsuarioPorRut(rutInput);
    if (!usuario.encontrado) return { success: false, message: 'RUT no encontrado en el sistema.' };

    // Validar ventana horaria
    try {
      var ssAsistVentana = getSpreadsheet('ASISTENCIA');
      var sheetPCtrl = ssAsistVentana.getSheetByName(CONFIG.HOJAS.PUNTOS_CONTROL);
      if (sheetPCtrl && sheetPCtrl.getLastRow() > 1) {
        var datosPC = sheetPCtrl.getDataRange().getDisplayValues();
        for (var pc = 1; pc < datosPC.length; pc++) {
          if (String(datosPC[pc][0]).trim() === nombreControl) {
            var horaApertura = normalizarHoraHHmm(String(datosPC[pc][4] || '').trim());
            var horaCierre   = normalizarHoraHHmm(String(datosPC[pc][5] || '').trim());
            if (horaApertura && horaCierre) {
              var horaActual = Utilities.formatDate(new Date(), 'America/Santiago', 'HH:mm');
              if (horaActual < horaApertura) return { success: false, ventanaCerrada: true, tipoVentana: 'aun_no_abre', horaApertura: horaApertura, horaCierre: horaCierre, message: 'El registro aun no ha comenzado. Abre a las ' + horaApertura + ' hrs.' };
              if (horaActual > horaCierre)   return { success: false, ventanaCerrada: true, tipoVentana: 'ya_cerro',     horaApertura: horaApertura, horaCierre: horaCierre, message: 'El registro ha cerrado. El periodo fue de ' + horaApertura + ' a ' + horaCierre + ' hrs.' };
            }
            break;
          }
        }
      }
    } catch (eVentana) {
      Logger.log('Advertencia ventana horaria virtual: ' + eVentana.toString());
    }

    var ssAsistencia    = getSpreadsheet('ASISTENCIA');
    var sheetAsistencia = ssAsistencia.getSheetByName(CONFIG.HOJAS.ASISTENCIA);
    if (!sheetAsistencia) {
      sheetAsistencia = ssAsistencia.insertSheet(CONFIG.HOJAS.ASISTENCIA);
      sheetAsistencia.appendRow(['FECHA_HORA', 'RUT', 'NOMBRE', 'ASAMBLEA', 'TIPO_ASISTENCIA', 'GESTION', 'CODIGO_TEMP', 'NOTIF_CORREO']);
    }

    // Extraer fecha del nombre del control (formato TIPO_DD-MM-YYYY_REGION)
    var partesControl = nombreControl.split('_');
    var fechaEvento = '';
    for (var p = 0; p < partesControl.length; p++) {
      if (/^\d{2}-\d{2}-\d{4}$/.test(partesControl[p])) {
        var fp = partesControl[p].split('-');
        fechaEvento = fp[0] + '/' + fp[1] + '/' + fp[2];
        break;
      }
    }

    var dataAsistencia = sheetAsistencia.getDataRange().getDisplayValues();
    for (var i = 1; i < dataAsistencia.length; i++) {
      var row = dataAsistencia[i];
      if (cleanRut(row[1]) !== rutLimpio) continue;

      if (row[3] === nombreControl) {
        return { success: false, yaRegistrado: true, message: 'Ya registraste tu asistencia en esta asamblea.' };
      }
      if (fechaEvento && row[4] === 'VIRTUAL') {
        var fechaRegistro = String(row[0]).split(' ')[0];
        if (fechaRegistro === fechaEvento) {
          return { success: false, yaRegistrado: true, message: 'Ya registraste tu asistencia en el evento de hoy desde otro dispositivo.' };
        }
      }
    }

    var fechaStr = Utilities.formatDate(new Date(), 'America/Santiago', 'dd/MM/yyyy HH:mm:ss');
    sheetAsistencia.appendRow([fechaStr, usuario.rut, usuario.nombre, nombreControl, 'VIRTUAL', 'Sistema', '', '']);

    return {
      success: true,
      nombre: usuario.nombre,
      rut: usuario.rut,
      fecha: fechaStr,
      mensajeCorreo: (usuario.correo && usuario.correo.includes('@'))
        ? 'Recibirás una confirmación en tu correo a más tardar esta noche.'
        : 'No tienes correo registrado. Puedes ver tu historial en el módulo Registro Asistencia.'
    };

  } catch (e) {
    Logger.log('Error en registrarAsistenciaVirtual: ' + e.toString());
    return { success: false, message: 'Error del servidor: ' + e.toString() };
  }
}

/**
 * Retorna TODAS las asambleas virtuales activas en este momento.
 */
function obtenerAsambleaVirtualActiva() {
  try {
    var ss    = getSpreadsheet('ASISTENCIA');
    var sheet = ss.getSheetByName(CONFIG.HOJAS.PUNTOS_CONTROL);
    if (!sheet || sheet.getLastRow() < 2) return { success: true, activa: false, asambleas: [] };

    var horaActual = Utilities.formatDate(new Date(), 'America/Santiago', 'HH:mm');
    var datos      = sheet.getDataRange().getDisplayValues();
    var asambleas  = [];

    for (var i = 1; i < datos.length; i++) {
      var nombre = String(datos[i][0] || '').trim();
      if (!nombre) continue;
      var tipo = String(datos[i][6] || '').trim().toUpperCase();
      if (tipo !== 'VIRTUAL') continue;

      var apertura = normalizarHoraHHmm(String(datos[i][4] || '').trim());
      var cierre   = normalizarHoraHHmm(String(datos[i][5] || '').trim());
      var activa   = (!apertura || !cierre) ? true : (horaActual >= apertura && horaActual <= cierre);
      if (activa) asambleas.push({ nombre: nombre, apertura: apertura, cierre: cierre });
    }

    return { success: true, activa: asambleas.length > 0, asambleas: asambleas };
  } catch (e) {
    Logger.log('Error en obtenerAsambleaVirtualActiva: ' + e.toString());
    return { success: false, activa: false, asambleas: [] };
  }
}

// ==========================================
// GESTIÓN PUNTOS DE CONTROL (ADMIN)
// ==========================================

/**
 * Retorna todos los puntos de control con su ventana horaria.
 * Columnas PUNTOS_CONTROL: A=NOMBRE, B=URL, C=QR_CODE, D=URL_BASE, E=HORA_APERTURA, F=HORA_CIERRE, G=TIPO
 */
function obtenerPuntosControl() {
  try {
    var ss    = getSpreadsheet('ASISTENCIA');
    var sheet = ss.getSheetByName(CONFIG.HOJAS.PUNTOS_CONTROL);
    if (!sheet || sheet.getLastRow() < 2) return { success: true, puntos: [] };

    var datos  = sheet.getDataRange().getDisplayValues();
    var puntos = [];

    for (var i = 1; i < datos.length; i++) {
      var nombre = String(datos[i][0] || '').trim();
      if (!nombre) continue;
      puntos.push({
        nombre:       nombre,
        horaApertura: normalizarHoraHHmm(String(datos[i][4] || '').trim()),
        horaCierre:   normalizarHoraHHmm(String(datos[i][5] || '').trim()),
        tipo:         String(datos[i][6] || 'PRESENCIAL').trim().toUpperCase() || 'PRESENCIAL'
      });
    }
    return { success: true, puntos: puntos };
  } catch (e) {
    Logger.log('Error en obtenerPuntosControl: ' + e.toString());
    return { success: false, message: e.toString() };
  }
}

/**
 * Crea un nuevo punto de control en la hoja PUNTOS_CONTROL.
 */
function crearPuntoControl(nombre, tipo) {
  try {
    nombre = String(nombre || '').trim();
    tipo   = (String(tipo || '').trim().toUpperCase() === 'VIRTUAL') ? 'VIRTUAL' : 'PRESENCIAL';
    if (!nombre) return { success: false, message: 'El nombre no puede estar vacío.' };

    var ss    = getSpreadsheet('ASISTENCIA');
    var sheet = ss.getSheetByName(CONFIG.HOJAS.PUNTOS_CONTROL);
    if (!sheet) return { success: false, message: 'Hoja PUNTOS_CONTROL no encontrada.' };

    if (sheet.getLastRow() > 1) {
      var existentes = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getDisplayValues();
      for (var i = 0; i < existentes.length; i++) {
        if (String(existentes[i][0]).trim() === nombre) {
          return { success: false, message: 'Ya existe un punto de control con ese nombre.' };
        }
      }
    }

    var nuevaFila = sheet.getLastRow() + 1;
    sheet.getRange(nuevaFila, 1).setValue(nombre);
    sheet.getRange(nuevaFila, 2).setFormula('=$D$1&"?action=checkin&control="&A' + nuevaFila);
    sheet.getRange(nuevaFila, 3).setFormula('=IMAGE("https://quickchart.io/qr?size=300&text="&ENCODEURL(B' + nuevaFila + '))');
    sheet.getRange(nuevaFila, 7).setValue(tipo);

    return { success: true, message: 'Punto de control creado correctamente.' };
  } catch (e) {
    Logger.log('Error en crearPuntoControl: ' + e.toString());
    return { success: false, message: e.toString() };
  }
}

/**
 * Guarda o actualiza la ventana horaria de un punto de control.
 */
function guardarVentanaPuntoControl(nombre, horaApertura, horaCierre) {
  try {
    nombre       = String(nombre       || '').trim();
    horaApertura = String(horaApertura || '').trim();
    horaCierre   = String(horaCierre   || '').trim();

    var reHora = /^([01]\d|2[0-3]):[0-5]\d$/;
    if (horaApertura && !reHora.test(horaApertura)) return { success: false, message: 'Hora de apertura inválida. Use formato HH:mm (ej: 08:30).' };
    if (horaCierre   && !reHora.test(horaCierre))   return { success: false, message: 'Hora de cierre inválida. Use formato HH:mm (ej: 10:30).' };

    var ss    = getSpreadsheet('ASISTENCIA');
    var sheet = ss.getSheetByName(CONFIG.HOJAS.PUNTOS_CONTROL);
    if (!sheet || sheet.getLastRow() < 2) return { success: false, message: 'Punto de control no encontrado.' };

    var datos = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getDisplayValues();
    for (var i = 0; i < datos.length; i++) {
      if (String(datos[i][0]).trim() === nombre) {
        sheet.getRange(i + 2, 5).setValue(horaApertura);
        sheet.getRange(i + 2, 6).setValue(horaCierre);
        return { success: true, message: 'Ventana horaria guardada correctamente.' };
      }
    }
    return { success: false, message: 'Punto de control no encontrado.' };
  } catch (e) {
    Logger.log('Error en guardarVentanaPuntoControl: ' + e.toString());
    return { success: false, message: e.toString() };
  }
}

/**
 * Elimina un punto de control completo de la hoja PUNTOS_CONTROL.
 */
function eliminarPuntoControl(nombre) {
  try {
    nombre = String(nombre || '').trim();
    var ss    = getSpreadsheet('ASISTENCIA');
    var sheet = ss.getSheetByName(CONFIG.HOJAS.PUNTOS_CONTROL);
    if (!sheet || sheet.getLastRow() < 2) return { success: false, message: 'Punto de control no encontrado.' };

    var datos = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getDisplayValues();
    for (var i = 0; i < datos.length; i++) {
      if (String(datos[i][0]).trim() === nombre) {
        sheet.deleteRow(i + 2);
        return { success: true, message: 'Punto de control eliminado.' };
      }
    }
    return { success: false, message: 'Punto de control no encontrado.' };
  } catch (e) {
    Logger.log('Error en eliminarPuntoControl: ' + e.toString());
    return { success: false, message: e.toString() };
  }
}

/**
 * Cierra el registro de asistencia de inmediato fijando HORA_CIERRE con la hora actual.
 */
function cerrarRegistroAhora(nombre) {
  try {
    nombre = String(nombre || '').trim();
    var horaActual = Utilities.formatDate(new Date(), 'America/Santiago', 'HH:mm');
    var ss    = getSpreadsheet('ASISTENCIA');
    var sheet = ss.getSheetByName(CONFIG.HOJAS.PUNTOS_CONTROL);
    if (!sheet || sheet.getLastRow() < 2) return { success: false, message: 'Punto de control no encontrado.' };

    var datos = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getDisplayValues();
    for (var i = 0; i < datos.length; i++) {
      if (String(datos[i][0]).trim() === nombre) {
        sheet.getRange(i + 2, 6).setValue(horaActual);
        return { success: true, message: 'Registro cerrado. Hora de cierre fijada en ' + horaActual + ' hrs.', horaCierre: horaActual };
      }
    }
    return { success: false, message: 'Punto de control no encontrado.' };
  } catch (e) {
    Logger.log('Error en cerrarRegistroAhora: ' + e.toString());
    return { success: false, message: e.toString() };
  }
}

// ==========================================
// SWITCH MÓDULO ASISTENCIA
// ==========================================

function obtenerEstadoSwitchAsistencia() {
  try {
    var estado = PropertiesService.getScriptProperties().getProperty('asistencia_habilitada');
    return { success: true, habilitado: (estado === null || estado === 'true') };
  } catch (e) {
    return { success: true, habilitado: true };
  }
}

function toggleSwitchAsistencia(estado) {
  try {
    PropertiesService.getScriptProperties().setProperty('asistencia_habilitada', estado ? 'true' : 'false');
    return { success: true };
  } catch (e) {
    return { success: false, message: 'Error: ' + e.toString() };
  }
}
