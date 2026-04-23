// ==========================================
// MODULO_PERMISOS_MEDICOS.GS — Permisos médicos laborales
// ==========================================

/**
 * Solicita un permiso médico
 */
function solicitarPermisoMedico(rutGestor, tipoPermiso, fechaInicio, motivo, rutBeneficiario, archivoData) {
  var CORREO_REPRESENTANTE_LEGAL = CONFIG.CORREOS.REPRESENTANTE_LEGAL;
  var CARPETA_ID = CONFIG.CARPETAS.PERMISOS_MEDICOS;

  var lock = LockService.getScriptLock();
  if (lock.tryLock(30000)) {
    try {
      var sheetUsers   = getSheet('USUARIOS', 'USUARIOS');
      var sheetPermisos= getSheet('PERMISOS_MEDICOS', 'PERMISOS_MEDICOS');
      var COL_USER = CONFIG.COLUMNAS.USUARIOS;
      var COL_PERM = CONFIG.COLUMNAS.PERMISOS_MEDICOS;
      var dataUsers = sheetUsers.getDataRange().getDisplayValues();

      // Validar gestor
      var gestor = null;
      var rutLimpioGestor = cleanRut(rutGestor);
      for (var i = 1; i < dataUsers.length; i++) {
        if (cleanRut(dataUsers[i][COL_USER.RUT]) === rutLimpioGestor) {
          gestor = { rut: dataUsers[i][COL_USER.RUT], nombre: dataUsers[i][COL_USER.NOMBRE], correo: dataUsers[i][COL_USER.CORREO] };
          break;
        }
      }
      if (!gestor) return { success: false, message: "Error de sesión." };

      // Determinar beneficiario
      var rutTarget = rutBeneficiario ? cleanRut(rutBeneficiario) : rutLimpioGestor;
      var beneficiario = null;
      if (rutTarget === rutLimpioGestor) {
        beneficiario = gestor;
      } else {
        for (var j = 1; j < dataUsers.length; j++) {
          if (cleanRut(dataUsers[j][COL_USER.RUT]) === rutTarget) {
            beneficiario = { rut: dataUsers[j][COL_USER.RUT], nombre: dataUsers[j][COL_USER.NOMBRE], correo: dataUsers[j][COL_USER.CORREO] };
            break;
          }
        }
        if (!beneficiario) return { success: false, message: "RUT del socio no encontrado." };
      }

      // Validar fecha de inicio (±7 días)
      var fechaInicioObj = new Date(fechaInicio + 'T12:00:00');
      var hoy = new Date();
      hoy.setHours(0, 0, 0, 0);
      fechaInicioObj.setHours(0, 0, 0, 0);
      var diffDias = Math.floor((fechaInicioObj - hoy) / (1000 * 60 * 60 * 24));
      if (diffDias < -7 || diffDias > 7) {
        return { success: false, message: "La fecha de inicio debe estar dentro del rango de 7 días antes o después de hoy." };
      }

      var fechaInicioNormalizada = fechaInicio.trim();
      Logger.log('🔍 Validando para RUT: ' + beneficiario.rut + ' | Fecha Inicio: ' + fechaInicioNormalizada);

      // Validar permiso activo con misma fecha de inicio
      var dataPermisos = sheetPermisos.getDataRange().getDisplayValues();
      var permisoConMismaFechaInicio = null, permisoAnuladoMismaFecha = null;

      for (var k = 1; k < dataPermisos.length; k++) {
        if (cleanRut(dataPermisos[k][COL_PERM.RUT]) === cleanRut(beneficiario.rut)) {
          var fechaInicioRegistro = dataPermisos[k][COL_PERM.FECHA_INICIO];
          var fechaInicioRegistroNorm = "";
          if (fechaInicioRegistro && fechaInicioRegistro.toString().trim() !== "") {
            if (fechaInicioRegistro.toString().match(/^\d{4}-\d{2}-\d{2}$/)) {
              fechaInicioRegistroNorm = fechaInicioRegistro.toString().trim();
            } else {
              try {
                var fo = new Date(fechaInicioRegistro);
                if (!isNaN(fo.getTime())) {
                  fechaInicioRegistroNorm = fo.getFullYear() + "-" + String(fo.getMonth()+1).padStart(2,'0') + "-" + String(fo.getDate()).padStart(2,'0');
                }
              } catch(e) { continue; }
            }
          }
          if (fechaInicioRegistroNorm === fechaInicioNormalizada) {
            if (dataPermisos[k][COL_PERM.ESTADO] === "Anulado") {
              permisoAnuladoMismaFecha = { id: dataPermisos[k][COL_PERM.ID] };
            } else {
              permisoConMismaFechaInicio = { id: dataPermisos[k][COL_PERM.ID], tipo: dataPermisos[k][COL_PERM.TIPO_PERMISO], estado: dataPermisos[k][COL_PERM.ESTADO] };
              break;
            }
          }
        }
      }

      if (permisoConMismaFechaInicio) {
        return { success: false, message: "❌ Ya existe un permiso médico ACTIVO con la misma fecha de inicio.\n\nID: " + permisoConMismaFechaInicio.id + "\nTipo: " + permisoConMismaFechaInicio.tipo + "\nEstado: " + permisoConMismaFechaInicio.estado + "\n\nSi cometió un error, puede anular el permiso existente desde el historial." };
      }
      if (permisoAnuladoMismaFecha) Logger.log('ℹ️ INFO: Hubo un permiso anulado para la fecha ' + fechaInicioNormalizada + ' (ID: ' + permisoAnuladoMismaFecha.id + '). Permitiendo crear uno nuevo.');

      var fechaHoyCompleta = new Date();
      var idUnico = Utilities.getUuid();
      var gestion = "Socio", nomDirigente = "", correoDirigente = "";
      if (rutTarget !== rutLimpioGestor) { gestion = "Dirigente"; nomDirigente = gestor.nombre; correoDirigente = gestor.correo; }

      var newRow = [];
      newRow[COL_PERM.ID]               = idUnico;
      newRow[COL_PERM.FECHA_SOLICITUD]  = fechaHoyCompleta;
      newRow[COL_PERM.RUT]              = beneficiario.rut;
      newRow[COL_PERM.NOMBRE]           = beneficiario.nombre;
      newRow[COL_PERM.CORREO]           = beneficiario.correo;
      newRow[COL_PERM.TIPO_PERMISO]     = tipoPermiso;
      newRow[COL_PERM.FECHA_INICIO]     = fechaInicioNormalizada;
      newRow[COL_PERM.MOTIVO_DETALLE]   = motivo;

      // Subir documento si fue adjuntado
      var urlDocFinal = "Sin documento";
      var estadoFinal = "Solicitado";
      var fechaSubidaFinal = "";

      if (archivoData && archivoData.base64) {
        var correosParaDoc = [];
        if (esCorreoValido(beneficiario.correo)) correosParaDoc.push({ correo: beneficiario.correo.trim().toLowerCase(), tipo: 'beneficiario', nombre: beneficiario.nombre });
        if (rutTarget !== rutLimpioGestor && esCorreoValido(gestor.correo) && gestor.correo.trim().toLowerCase() !== beneficiario.correo.trim().toLowerCase()) correosParaDoc.push({ correo: gestor.correo.trim().toLowerCase(), tipo: 'gestor', nombre: gestor.nombre });

        var resultadoSubida = subirArchivoConPermisos(archivoData, CARPETA_ID, "PERMISO-" + idUnico + "-" + cleanRut(beneficiario.rut), correosParaDoc, [CORREO_REPRESENTANTE_LEGAL]);
        if (!resultadoSubida.success) return { success: false, message: "Error al subir el documento adjunto: " + resultadoSubida.mensajeError };
        urlDocFinal = resultadoSubida.url;
        estadoFinal = "Solicitado con Documento";
        fechaSubidaFinal = fechaHoyCompleta;
      }

      newRow[COL_PERM.URL_DOCUMENTO]       = urlDocFinal;
      newRow[COL_PERM.ESTADO]              = estadoFinal;
      newRow[COL_PERM.FECHA_SUBIDA]        = fechaSubidaFinal;
      newRow[COL_PERM.NOTIFICADO_REP_LEGAL]= false;
      newRow[COL_PERM.NOTIFICADO_SOCIO]    = false;
      newRow[COL_PERM.GESTION]             = gestion;
      newRow[COL_PERM.NOMBRE_DIRIGENTE]    = nomDirigente;
      newRow[COL_PERM.CORREO_DIRIGENTE]    = correoDirigente;

      // Segunda validación anti-race-condition
      var dataPermisosPreEscritura = sheetPermisos.getDataRange().getDisplayValues();
      for (var n = 1; n < dataPermisosPreEscritura.length; n++) {
        if (cleanRut(dataPermisosPreEscritura[n][COL_PERM.RUT]) !== cleanRut(beneficiario.rut)) continue;
        var fechaInicioReg2 = dataPermisosPreEscritura[n][COL_PERM.FECHA_INICIO];
        var fechaNorm2 = "";
        if (fechaInicioReg2 && fechaInicioReg2.toString().trim() !== "") {
          if (fechaInicioReg2.toString().match(/^\d{4}-\d{2}-\d{2}$/)) { fechaNorm2 = fechaInicioReg2.toString().trim(); }
          else {
            try { var fo2 = new Date(fechaInicioReg2); fechaNorm2 = fo2.getFullYear() + "-" + String(fo2.getMonth()+1).padStart(2,'0') + "-" + String(fo2.getDate()).padStart(2,'0'); } catch(e) { continue; }
          }
        }
        if (fechaNorm2 === fechaInicioNormalizada && dataPermisosPreEscritura[n][COL_PERM.ESTADO] !== "Anulado") {
          return { success: false, message: "Se detectó otra solicitud en proceso con la misma fecha de inicio. Por favor, recarga la página y verifica tu historial." };
        }
      }

      sheetPermisos.appendRow(newRow);
      var filaRegistro = sheetPermisos.getLastRow();
      Logger.log('✅ Permiso creado exitosamente. ID: ' + idUnico);

      var fechaInicioEmailStr = new Date(fechaInicio + 'T12:00:00').toLocaleDateString('es-CL', { year: 'numeric', month: 'long', day: 'numeric' });
      var tieneDoc = urlDocFinal !== "Sin documento";

      // Correo al socio (si gestiona por sí mismo)
      if (gestion !== "Dirigente") {
        if (esCorreoValido(beneficiario.correo)) {
          var asuntoSocio = tieneDoc ? "Permiso Medico Registrado con Documento - Sindicato SLIM n3" : "Solicitud de Permiso Medico Registrada - Sindicato SLIM n3";
          var tituloSocio = tieneDoc ? "Solicitud Completa con Documento" : "Permiso Medico Solicitado";
          var mensajeSocio = tieneDoc
            ? "Hola <strong>" + beneficiario.nombre + "</strong>, tu permiso medico ha sido registrado correctamente y el documento medico de respaldo fue adjuntado en el mismo momento. No necesitas realizar ninguna accion adicional."
            : "Hola <strong>" + beneficiario.nombre + "</strong>, tu solicitud de permiso medico ha sido registrada exitosamente. Recuerda adjuntar el documento medico de respaldo desde el historial del modulo una vez realizada la atencion medica.";
          var datosSocio = { "ID": idUnico, "Tipo Permiso": tipoPermiso, "Fecha Inicio": fechaInicioEmailStr, "Motivo": motivo, "Estado": estadoFinal, "Documento": tieneDoc ? '<a href="' + urlDocFinal + '" style="color:#10b981;text-decoration:none;font-weight:600;">Ver Documento Adjunto</a>' : "Pendiente - Adjuntar desde historial una vez realizada la atencion medica" };
          try {
            enviarCorreoEstilizado(beneficiario.correo, asuntoSocio, tituloSocio, mensajeSocio, datosSocio, "#10b981");
            sheetPermisos.getRange(filaRegistro, COL_PERM.NOTIFICADO_SOCIO + 1).setValue(true);
          } catch (eSocio) { Logger.log("Advertencia: Fallo envio socio fila " + filaRegistro + " - " + eSocio.toString()); }
        } else {
          sheetPermisos.getRange(filaRegistro, COL_PERM.NOTIFICADO_SOCIO + 1).setValue("SIN_CORREO");
        }
      }

      // Correo al representante legal
      try {
        var asuntoRL = tieneDoc ? "Nueva Solicitud Permiso Medico con Documento - Sindicato SLIM n3" : "Nueva Solicitud Permiso Medico - Sindicato SLIM n3";
        var tituloRL = tieneDoc ? "Solicitud de Permiso Medico con Documento Adjunto" : "Solicitud de Permiso Medico Sin Documento";
        var mensajeRL = tieneDoc
          ? "El trabajador <strong>" + beneficiario.nombre + "</strong> ha registrado una solicitud de permiso medico con el documento de respaldo adjunto al momento del registro."
          : "El trabajador <strong>" + beneficiario.nombre + "</strong> ha registrado una solicitud de permiso medico. El documento de respaldo aun no ha sido adjuntado.";
        var datosRL = { "ID": idUnico, "Trabajador": beneficiario.nombre, "RUT": beneficiario.rut, "Tipo Permiso": tipoPermiso, "Fecha Inicio": fechaInicioNormalizada, "Motivo": motivo, "Estado": estadoFinal, "Documento": tieneDoc ? '<a href="' + urlDocFinal + '" style="color:#10b981;text-decoration:none;font-weight:600;">Ver Documento Adjunto</a>' : "Pendiente" };
        enviarCorreoEstilizado(CORREO_REPRESENTANTE_LEGAL, asuntoRL, tituloRL, mensajeRL, datosRL, "#10b981");
        sheetPermisos.getRange(filaRegistro, COL_PERM.NOTIFICADO_REP_LEGAL + 1).setValue(true);
      } catch (eRL) { Logger.log("Advertencia: Fallo envio Rep. Legal fila " + filaRegistro + " - " + eRL.toString()); }

      // Correo de respaldo al dirigente
      if (gestion === "Dirigente" && correoDirigente && correoDirigente.includes("@") && correoDirigente !== beneficiario.correo) {
        enviarCorreoEstilizado(correoDirigente, "Respaldo Gestión Permiso Médico - Sindicato SLIM n°3", "Permiso Médico Ingresado",
          "Has ingresado exitosamente un permiso médico para el socio <strong>" + beneficiario.nombre + "</strong>.",
          { "ID": idUnico, "Socio": beneficiario.nombre, "Tipo": tipoPermiso, "Estado": estadoFinal, "Documento": tieneDoc ? '<a href="' + urlDocFinal + '" style="color:#475569;text-decoration:none;font-weight:600;">📎 Ver Documento Adjunto</a>' : "Sin documento adjunto" },
          "#475569");
      }

      // Copia al socio cuando el dirigente gestiona en su nombre
      if (gestion === "Dirigente") {
        if (esCorreoValido(beneficiario.correo)) {
          var mensajeSocioDirigente = tieneDoc
            ? "Hola <strong>" + beneficiario.nombre + "</strong>, un dirigente ha solicitado un permiso medico a tu nombre y el documento de respaldo ha sido adjuntado exitosamente. No necesitas realizar ninguna accion adicional."
            : "Hola <strong>" + beneficiario.nombre + "</strong>, un dirigente ha solicitado un permiso medico a tu nombre.\n<strong>IMPORTANTE:</strong> Debes adjuntar el documento de respaldo en el historial del modulo una vez realizada la atencion medica.";
          try {
            enviarCorreoEstilizado(beneficiario.correo, "Permiso Medico Solicitado - Sindicato SLIM n3", "Permiso Medico Ingresado",
              mensajeSocioDirigente,
              { "ID": idUnico, "Trabajador": beneficiario.nombre, "RUT": beneficiario.rut, "Tipo": tipoPermiso, "Fecha Inicio": fechaInicioEmailStr, "Motivo": motivo, "Dirigente": nomDirigente, "Estado": estadoFinal, "Documento": tieneDoc ? '<a href="' + urlDocFinal + '" style="color:#10b981;text-decoration:none;font-weight:600;">Ver Documento Adjunto</a>' : "Pendiente de adjuntar desde el historial" },
              "#10b981");
            sheetPermisos.getRange(filaRegistro, COL_PERM.NOTIFICADO_SOCIO + 1).setValue(true);
          } catch (eSocioDirig) { Logger.log("Advertencia: Fallo envio socio (dirigente) fila " + filaRegistro + " - " + eSocioDirig.toString()); }
        } else {
          sheetPermisos.getRange(filaRegistro, COL_PERM.NOTIFICADO_SOCIO + 1).setValue("SIN_CORREO");
        }
      }

      return { success: true, message: tieneDoc ? "Permiso medico registrado con documento adjunto exitosamente." : "Permiso medico solicitado. No olvides adjuntar el documento de respaldo desde el historial." };

    } catch (e) {
      Logger.log("❌ Error en solicitarPermisoMedico: " + e.toString());
      return { success: false, message: "Error: " + e.toString() };
    } finally {
      lock.releaseLock();
    }
  } else {
    return { success: false, message: "Servidor ocupado. Intenta nuevamente." };
  }
}

/**
 * Adjunta un documento de respaldo a un permiso médico existente
 */
function adjuntarDocumentoPermiso(idPermiso, archivoData) {
  var CARPETA_ID = CONFIG.CARPETAS.PERMISOS_MEDICOS;
  var CORREO_REPRESENTANTE_LEGAL = CONFIG.CORREOS.REPRESENTANTE_LEGAL;

  var lock = LockService.getScriptLock();
  if (lock.tryLock(30000)) {
    try {
      var sheetPermisos = getSheet('PERMISOS_MEDICOS', 'PERMISOS_MEDICOS');
      var data = sheetPermisos.getDataRange().getValues();
      var COL = CONFIG.COLUMNAS.PERMISOS_MEDICOS;

      var rowIndex = -1, beneficiario = null, tipoPermiso = "", gestionTipo = "", correoGestor = "";
      for (var i = 1; i < data.length; i++) {
        if (String(data[i][COL.ID]) === String(idPermiso)) {
          rowIndex = i + 1;
          beneficiario = { nombre: data[i][COL.NOMBRE], correo: data[i][COL.CORREO], rut: data[i][COL.RUT] };
          tipoPermiso = data[i][COL.TIPO_PERMISO];
          gestionTipo = data[i][COL.GESTION];
          correoGestor = data[i][COL.CORREO_DIRIGENTE];
          break;
        }
      }
      if (rowIndex === -1) return { success: false, message: "Permiso no encontrado." };

      var esGestionDirigente = gestionTipo === "Dirigente" && esCorreoValido(correoGestor);
      var correosParaPermisos = [];
      var alertas = [];

      if (esCorreoValido(beneficiario.correo)) {
        correosParaPermisos.push({ correo: beneficiario.correo.trim().toLowerCase(), tipo: 'beneficiario', nombre: beneficiario.nombre });
      } else {
        alertas.push("El socio " + beneficiario.nombre + " no tiene correo válido. No podrá acceder al documento.");
      }
      if (esGestionDirigente && correoGestor !== beneficiario.correo) {
        correosParaPermisos.push({ correo: correoGestor.trim().toLowerCase(), tipo: 'gestor', nombre: 'Dirigente' });
      }

      var resultadoSubida = subirArchivoConPermisos(archivoData, CARPETA_ID, "PERMISO-" + idPermiso + "-" + cleanRut(beneficiario.rut), correosParaPermisos, [CORREO_REPRESENTANTE_LEGAL]);
      if (!resultadoSubida.success) return { success: false, message: resultadoSubida.mensajeError };

      var fechaSubida = new Date();
      var nuevoEstado = "Documento Adjuntado";
      sheetPermisos.getRange(rowIndex, COL.URL_DOCUMENTO + 1).setValue(resultadoSubida.url);
      sheetPermisos.getRange(rowIndex, COL.ESTADO + 1).setValue(nuevoEstado);
      sheetPermisos.getRange(rowIndex, COL.FECHA_SUBIDA + 1).setValue(fechaSubida);
      sheetPermisos.getRange(rowIndex, COL.NOTIFICADO_REP_LEGAL + 1).setValue(false);
      sheetPermisos.getRange(rowIndex, COL.NOTIFICADO_SOCIO + 1).setValue(false);

      if (esCorreoValido(beneficiario.correo)) {
        try {
          enviarCorreoEstilizado(beneficiario.correo, "Documento Adjuntado - Sindicato SLIM n3", "Documento de Permiso Medico Adjuntado", "Hola " + beneficiario.nombre + ", tu documento de respaldo ha sido adjuntado exitosamente.", { "ID": idPermiso, "Tipo Permiso": tipoPermiso, "Estado": nuevoEstado, "Documento": '<a href="' + resultadoSubida.url + '" style="color:#10b981;text-decoration:none;font-weight:600;">Ver Documento</a>' }, "#10b981");
          sheetPermisos.getRange(rowIndex, COL.NOTIFICADO_SOCIO + 1).setValue(true);
        } catch(e) {}
      } else {
        sheetPermisos.getRange(rowIndex, COL.NOTIFICADO_SOCIO + 1).setValue("SIN_CORREO");
      }

      try {
        enviarCorreoEstilizado(CORREO_REPRESENTANTE_LEGAL, "Documento Permiso Medico Adjuntado - Sindicato SLIM n3", "Documento de Permiso Medico Disponible", "El trabajador <strong>" + beneficiario.nombre + "</strong> ha adjuntado el documento de respaldo para su permiso medico.", { "ID": idPermiso, "Trabajador": beneficiario.nombre, "RUT": beneficiario.rut, "Tipo Permiso": tipoPermiso, "Documento": '<a href="' + resultadoSubida.url + '" style="color:#10b981;font-weight:bold;">Disponible para revision</a>', "Fecha Adjunto": fechaSubida.toLocaleDateString() }, "#475569");
        sheetPermisos.getRange(rowIndex, COL.NOTIFICADO_REP_LEGAL + 1).setValue(true);
      } catch(e) { Logger.log("Advertencia adjuntarDocumentoPermiso: Fallo envio Rep. Legal - " + e.toString()); }

      var respuesta = { success: true, message: "Documento adjuntado y notificaciones enviadas." };
      if (alertas.length > 0 || (resultadoSubida.permisosError && resultadoSubida.permisosError.length > 0)) {
        var todosDetalles = alertas.slice();
        if (resultadoSubida.permisosError) resultadoSubida.permisosError.forEach(function(err) { todosDetalles.push("No se pudo dar acceso a " + err.nombre); });
        respuesta.mostrarAlerta = true; respuesta.tipoAlerta = 'warning'; respuesta.mensajeAlerta = todosDetalles.join('\n\n');
      }
      return respuesta;

    } catch (e) {
      return { success: false, message: "Error: " + e.toString() };
    } finally {
      lock.releaseLock();
    }
  } else {
    return { success: false, message: "Servidor ocupado." };
  }
}

/**
 * Obtiene el historial de permisos médicos de un usuario
 */
function obtenerHistorialPermisosMedicos(rutInput) {
  try {
    var sheet = getSheet('PERMISOS_MEDICOS', 'PERMISOS_MEDICOS');
    var COL = CONFIG.COLUMNAS.PERMISOS_MEDICOS;
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return { success: true, registros: [] };
    var lastCol = sheet.getLastColumn();
    var data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
    var rutLimpio = cleanRut(rutInput);
    var registros = [];
    for (var i = 0; i < data.length; i++) {
      var row = data[i];
      if (cleanRut(row[COL.RUT]) === rutLimpio) {
        registros.push({ id: row[COL.ID], fecha: formatearFechaConHora(row[COL.FECHA_SOLICITUD]), tipoPermiso: row[COL.TIPO_PERMISO], fechaInicio: formatearFechaSinHora(row[COL.FECHA_INICIO]), motivo: row[COL.MOTIVO_DETALLE], urlDocumento: row[COL.URL_DOCUMENTO], estado: row[COL.ESTADO], gestion: row[COL.GESTION], nomDirigente: row[COL.NOMBRE_DIRIGENTE] });
      }
    }
    registros.reverse();
    return { success: true, registros: registros };
  } catch (e) {
    Logger.log("❌ Error en obtenerHistorialPermisosMedicos: " + e.toString());
    return { success: false, message: "Error: " + e.toString() };
  }
}

/**
 * Anula un permiso médico en estado "Solicitado"
 */
function eliminarPermisoMedico(idPermiso) {
  var CORREO_REPRESENTANTE_LEGAL = CONFIG.CORREOS.REPRESENTANTE_LEGAL;
  var lock = LockService.getScriptLock();
  if (lock.tryLock(30000)) {
    try {
      var sheet = getSheet('PERMISOS_MEDICOS', 'PERMISOS_MEDICOS');
      var data = sheet.getDataRange().getDisplayValues();
      var COL = CONFIG.COLUMNAS.PERMISOS_MEDICOS;

      for (var i = 1; i < data.length; i++) {
        if (String(data[i][COL.ID]) === String(idPermiso)) {
          if (String(data[i][COL.ESTADO]) !== "Solicitado") return { success: false, message: "Solo se pueden anular permisos en estado 'Solicitado'." };

          var beneficiario = { nombre: data[i][COL.NOMBRE], correo: data[i][COL.CORREO], rut: data[i][COL.RUT] };
          var tipoPermiso = data[i][COL.TIPO_PERMISO];
          var fechaInicio = data[i][COL.FECHA_INICIO];

          if (beneficiario.correo && beneficiario.correo.includes("@")) {
            enviarCorreoEstilizado(beneficiario.correo, "Permiso Médico Anulado - Sindicato SLIM n°3", "Solicitud de Permiso Anulada",
              "Hola " + beneficiario.nombre + ", tu solicitud de permiso médico ha sido anulada. No se hará uso de este permiso.",
              { "ID": idPermiso, "Tipo Permiso": tipoPermiso, "Fecha Inicio": fechaInicio, "Estado": "Anulado", "Acción": "Solicitud eliminada del sistema" },
              "#ef4444");
          }
          enviarCorreoEstilizado(CORREO_REPRESENTANTE_LEGAL, "Permiso Médico Anulado - Sindicato SLIM n°3", "Solicitud de Permiso Anulada",
            "La solicitud de permiso médico del trabajador <strong>" + beneficiario.nombre + "</strong> ha sido anulada.",
            { "ID": idPermiso, "Trabajador": beneficiario.nombre, "RUT": beneficiario.rut, "Tipo Permiso": tipoPermiso, "Fecha Inicio": fechaInicio, "Estado": "Anulado por el usuario" },
            "#475569");

          sheet.deleteRow(i + 1);
          return { success: true, message: "Permiso anulado y notificaciones enviadas." };
        }
      }
      return { success: false, message: "Permiso no encontrado." };
    } catch (e) {
      return { success: false, message: "Error: " + e.toString() };
    } finally {
      lock.releaseLock();
    }
  } else {
    return { success: false, message: "Servidor ocupado." };
  }
}

// ==========================================
// TRIGGERS DE NOTIFICACIÓN — REINTENTO
// ==========================================

/**
 * Trigger: cada 30 minutos. Reintento de notificación al socio.
 */
function reintentarNotificacionSocio() {
  var COL = CONFIG.COLUMNAS.PERMISOS_MEDICOS;
  try {
    var sheetPermisos = getSheet('PERMISOS_MEDICOS', 'PERMISOS_MEDICOS');
    var data = sheetPermisos.getDataRange().getValues();
    var pendientes = 0, exitosos = 0;

    for (var i = 1; i < data.length; i++) {
      var fila = data[i];
      var notificado = fila[COL.NOTIFICADO_SOCIO];
      var estado     = String(fila[COL.ESTADO]);
      var correo     = String(fila[COL.CORREO] || "");

      if (estado === '' || estado === 'Anulado') continue;
      if (notificado === true || String(notificado).toUpperCase() === 'TRUE') continue;
      if (String(notificado).toUpperCase() === 'SIN_CORREO') continue;
      if (!esCorreoValido(correo)) { sheetPermisos.getRange(i + 1, COL.NOTIFICADO_SOCIO + 1).setValue("SIN_CORREO"); continue; }

      pendientes++;
      var idPermiso   = String(fila[COL.ID]);
      var nombre      = String(fila[COL.NOMBRE]);
      var rut         = String(fila[COL.RUT]);
      var tipoPermiso = String(fila[COL.TIPO_PERMISO]);
      var urlDoc      = String(fila[COL.URL_DOCUMENTO]);
      var motivo      = String(fila[COL.MOTIVO_DETALLE]);
      var gestion     = String(fila[COL.GESTION]);
      var fechaVal    = fila[COL.FECHA_INICIO];
      var fechaInicioStr = (fechaVal instanceof Date) ? fechaVal.toLocaleDateString('es-CL', { year:'numeric', month:'long', day:'numeric' }) : String(fechaVal);
      var tieneDoc = urlDoc && urlDoc !== '' && urlDoc !== 'Sin documento';

      try {
        if (estado === 'Solicitado con Documento') {
          enviarCorreoEstilizado(correo, "Permiso Medico Registrado con Documento - Sindicato SLIM n3", "Solicitud Completa con Documento",
            "Hola <strong>" + nombre + "</strong>, tu permiso medico ha sido registrado correctamente y el documento medico de respaldo fue adjuntado en el mismo momento.",
            { "ID": idPermiso, "Tipo Permiso": tipoPermiso, "Fecha Inicio": fechaInicioStr, "Estado": estado, "Documento": tieneDoc ? '<a href="' + urlDoc + '" style="color:#10b981;font-weight:bold;">Ver Documento Adjunto</a>' : "Sin documento" }, "#10b981");
        } else if (estado === 'Documento Adjuntado') {
          enviarCorreoEstilizado(correo, "Documento de Respaldo Adjuntado - Sindicato SLIM n3", "Documento de Permiso Medico Adjuntado",
            "Hola " + nombre + ", has adjuntado tu documento medico de respaldo exitosamente a tu permiso existente.",
            { "ID": idPermiso, "Tipo Permiso": tipoPermiso, "Fecha Inicio": fechaInicioStr, "Estado": estado, "Documento": tieneDoc ? '<a href="' + urlDoc + '" style="color:#10b981;font-weight:bold;">Ver Documento</a>' : "Sin documento" }, "#10b981");
        } else {
          var esDirigente = gestion === "Dirigente";
          var mensajeSocio = esDirigente
            ? "Hola <strong>" + nombre + "</strong>, un dirigente ha solicitado un permiso medico a tu nombre. Recuerda adjuntar el documento de respaldo desde el historial del modulo una vez realizada la atencion medica."
            : "Hola " + nombre + ", se ha registrado tu solicitud de permiso medico. Debes adjuntar el documento de respaldo en el historial del modulo una vez realizada la atencion medica.";
          enviarCorreoEstilizado(correo, "Solicitud Permiso Medico - Sindicato SLIM n3", "Permiso Medico Solicitado", mensajeSocio,
            { "ID": idPermiso, "Trabajador": nombre, "RUT": rut, "Tipo": tipoPermiso, "Fecha Inicio": fechaInicioStr, "Motivo": motivo, "Estado": estado }, "#10b981");
        }
        sheetPermisos.getRange(i + 1, COL.NOTIFICADO_SOCIO + 1).setValue(true);
        exitosos++;
        Utilities.sleep(600);
      } catch (eEmail) { Logger.log("reintentarNotificacionSocio - Fila " + (i + 1) + ": " + eEmail.toString()); }
    }
    Logger.log("reintentarNotificacionSocio: " + pendientes + " pendientes, " + exitosos + " enviados.");
  } catch (e) { Logger.log("Error en reintentarNotificacionSocio: " + e.toString()); }
}

/**
 * Trigger: cada 30 minutos. Reintento de notificación al representante legal.
 */
function reintentarNotificacionRepLegal() {
  var CORREO_REPRESENTANTE_LEGAL = CONFIG.CORREOS.REPRESENTANTE_LEGAL;
  var COL = CONFIG.COLUMNAS.PERMISOS_MEDICOS;
  try {
    var sheetPermisos = getSheet('PERMISOS_MEDICOS', 'PERMISOS_MEDICOS');
    var data = sheetPermisos.getDataRange().getValues();
    var pendientes = 0, exitosos = 0;

    for (var i = 1; i < data.length; i++) {
      var fila = data[i];
      var notificado = fila[COL.NOTIFICADO_REP_LEGAL];
      var estado = String(fila[COL.ESTADO]);

      if (estado === '' || estado === 'Anulado') continue;
      if (notificado === true || String(notificado).toUpperCase() === 'TRUE') continue;

      pendientes++;
      var idPermiso   = String(fila[COL.ID]);
      var nombre      = String(fila[COL.NOMBRE]);
      var rut         = String(fila[COL.RUT]);
      var tipoPermiso = String(fila[COL.TIPO_PERMISO]);
      var urlDoc      = String(fila[COL.URL_DOCUMENTO]);
      var motivo      = String(fila[COL.MOTIVO_DETALLE]);
      var fechaVal    = fila[COL.FECHA_INICIO];
      var fechaInicioStr = (fechaVal instanceof Date) ? fechaVal.toLocaleDateString('es-CL', { year:'numeric', month:'long', day:'numeric' }) : String(fechaVal);
      var tieneDoc = urlDoc && urlDoc !== '' && urlDoc !== 'Sin documento';

      try {
        if (estado === 'Solicitado con Documento') {
          enviarCorreoEstilizado(CORREO_REPRESENTANTE_LEGAL, "Nueva Solicitud Permiso Medico con Documento - Sindicato SLIM n3", "Solicitud de Permiso Medico con Documento Adjunto",
            "El trabajador <strong>" + nombre + "</strong> ha registrado una solicitud de permiso medico con el documento de respaldo adjunto al momento del registro.",
            { "ID": idPermiso, "Trabajador": nombre, "RUT": rut, "Tipo Permiso": tipoPermiso, "Fecha Inicio": fechaInicioStr, "Estado": estado, "Documento": tieneDoc ? '<a href="' + urlDoc + '" style="color:#10b981;font-weight:bold;">Ver Documento Adjunto</a>' : "Sin documento" }, "#10b981");
        } else if (estado === 'Documento Adjuntado') {
          enviarCorreoEstilizado(CORREO_REPRESENTANTE_LEGAL, "Documento de Respaldo Adjuntado - Permiso Medico - Sindicato SLIM n3", "Documento de Permiso Medico Disponible",
            "El trabajador <strong>" + nombre + "</strong> ha adjuntado el documento medico de respaldo a su permiso existente.",
            { "ID": idPermiso, "Trabajador": nombre, "RUT": rut, "Tipo Permiso": tipoPermiso, "Fecha Inicio": fechaInicioStr, "Estado": estado, "Documento": tieneDoc ? '<a href="' + urlDoc + '" style="color:#10b981;font-weight:bold;">Ver Documento Adjunto</a>' : "Sin documento" }, "#475569");
        } else {
          enviarCorreoEstilizado(CORREO_REPRESENTANTE_LEGAL, "Nueva Solicitud Permiso Medico - Sindicato SLIM n3", "Solicitud de Permiso Medico Sin Documento",
            "El trabajador <strong>" + nombre + "</strong> ha registrado una solicitud de permiso medico. El documento de respaldo aun no ha sido adjuntado.",
            { "ID": idPermiso, "Trabajador": nombre, "RUT": rut, "Tipo Permiso": tipoPermiso, "Fecha Inicio": fechaInicioStr, "Motivo": motivo, "Estado": estado, "Documento": "Pendiente" }, "#10b981");
        }
        sheetPermisos.getRange(i + 1, COL.NOTIFICADO_REP_LEGAL + 1).setValue(true);
        exitosos++;
        Utilities.sleep(600);
      } catch (eEmail) { Logger.log("reintentarNotificacionRepLegal - Fila " + (i + 1) + ": " + eEmail.toString()); }
    }
    Logger.log("reintentarNotificacionRepLegal: " + pendientes + " pendientes, " + exitosos + " enviados.");
  } catch (e) { Logger.log("Error en reintentarNotificacionRepLegal: " + e.toString()); }
}

// ==========================================
// SWITCH MÓDULO PERMISOS MÉDICOS
// ==========================================

function obtenerEstadoSwitchPermisosMedicos() {
  try {
    var estado = PropertiesService.getScriptProperties().getProperty('permisos_medicos_habilitado');
    return { success: true, habilitado: (estado === null || estado === 'true') };
  } catch (e) {
    return { success: true, habilitado: true };
  }
}

function toggleSwitchPermisosMedicos(estado) {
  try {
    PropertiesService.getScriptProperties().setProperty('permisos_medicos_habilitado', estado ? 'true' : 'false');
    return { success: true };
  } catch (e) {
    return { success: false, message: 'Error: ' + e.toString() };
  }
}
