// ==========================================
// MODULO_PRESTAMOS.GS — Lógica completa de préstamos
// ==========================================

/**
 * Crea una solicitud de préstamo
 */
function crearSolicitudPrestamo(rutGestor, tipo, cuotas, medioPago, rutBeneficiario) {
  var lock = LockService.getScriptLock();
  if (lock.tryLock(30000)) {
    try {
      var sheetUsers    = getSheet('USUARIOS',  'USUARIOS');
      var sheetPrestamos= getSheet('PRESTAMOS', 'PRESTAMOS');
      var COL_USER = CONFIG.COLUMNAS.USUARIOS;
      var COL_PRES = CONFIG.COLUMNAS.PRESTAMOS;

      var dataUsers = sheetUsers.getDataRange().getDisplayValues();

      // 1. Identificar al Gestor
      var gestor = null;
      var rutLimpioGestor = cleanRut(rutGestor);
      for (var i = 1; i < dataUsers.length; i++) {
        if (cleanRut(dataUsers[i][COL_USER.RUT]) === rutLimpioGestor) {
          gestor = { rut: dataUsers[i][COL_USER.RUT], nombre: dataUsers[i][COL_USER.NOMBRE], correo: dataUsers[i][COL_USER.CORREO] };
          break;
        }
      }
      if (!gestor) return { success: false, message: "Error de sesión." };

      // 2. Identificar al Beneficiario
      var rutTarget = rutBeneficiario ? cleanRut(rutBeneficiario) : rutLimpioGestor;
      var beneficiario = null;
      var esGestionDirigente = (rutTarget !== rutLimpioGestor);

      if (!esGestionDirigente) {
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

      // 3. Validar préstamos activos del mismo tipo base
      var dataPrestamos = sheetPrestamos.getDataRange().getDisplayValues();
      var tipoBaseNuevo = tipo.split(' - ')[0].trim();

      for (var k = 1; k < dataPrestamos.length; k++) {
        var row = dataPrestamos[k];
        var rowRut   = cleanRut(row[COL_PRES.RUT]);
        var rowEstado= row[COL_PRES.ESTADO];
        var rowTipo  = String(row[COL_PRES.TIPO] || '');
        var tipoBaseExistente = rowTipo.split(' - ')[0].trim();

        if (rowRut === cleanRut(beneficiario.rut) &&
            ["Solicitado","Enviado","Vigente"].indexOf(rowEstado) !== -1 &&
            tipoBaseExistente === tipoBaseNuevo) {
          return {
            success: false,
            message: 'Tienes un préstamo de ' + tipoBaseNuevo + ' en estado "' + rowEstado + '". Solo puedes solicitar uno nuevo cuando el préstamo actual esté Pagado o Rechazado.'
          };
        }
      }

      // 4. Calcular monto
      var montoTexto = "$0";
      if (tipo.includes('Emergencia')) {
        montoTexto = (tipo.includes('Opcion B') || tipo.includes('Opción B')) ? "$400.000" : "$300.000";
      } else if (tipo.includes('Vacaciones')) {
        montoTexto = (tipo.includes('Opcion B') || tipo.includes('Opción B')) ? "$300.000" : "$200.000";
      }

      // 5. Calcular fechas
      var fechaSolicitud = new Date();
      var diaSolicitud = fechaSolicitud.getDate();
      var idUnico = Utilities.getUuid();
      var fechaInicioPago = new Date(fechaSolicitud);

      if (diaSolicitud > 24) fechaInicioPago.setMonth(fechaInicioPago.getMonth() + 1);

      var fechaTermino = new Date(fechaInicioPago);
      var numCuotas = parseInt(cuotas);
      if (!isNaN(numCuotas)) {
        fechaTermino.setMonth(fechaTermino.getMonth() + numCuotas);
        fechaTermino = new Date(fechaTermino.getFullYear(), fechaTermino.getMonth() + 1, 0);
      }

      var gestion       = esGestionDirigente ? "Dirigente" : "Socio";
      var nomDirigente  = esGestionDirigente ? gestor.nombre : "";
      var correoDirigente = esGestionDirigente ? gestor.correo : "";

      // 6. Guardar en BD
      var newRow = [];
      newRow[COL_PRES.ID]              = idUnico;
      newRow[COL_PRES.FECHA]           = fechaSolicitud;
      newRow[COL_PRES.RUT]             = beneficiario.rut;
      newRow[COL_PRES.NOMBRE]          = beneficiario.nombre;
      newRow[COL_PRES.CORREO]          = beneficiario.correo;
      newRow[COL_PRES.TIPO]            = tipo;
      newRow[COL_PRES.MONTO]           = "'" + montoTexto;
      newRow[COL_PRES.CUOTAS]          = cuotas;
      newRow[COL_PRES.MEDIO_PAGO]      = medioPago;
      newRow[COL_PRES.ESTADO]          = "Solicitado";
      newRow[COL_PRES.FECHA_TERMINO]   = fechaTermino;
      newRow[COL_PRES.GESTION]         = gestion;
      newRow[COL_PRES.NOMBRE_DIRIGENTE]= nomDirigente;
      newRow[COL_PRES.CORREO_DIRIGENTE]= correoDirigente;
      newRow[COL_PRES.INFORME]         = "";
      sheetPrestamos.appendRow(newRow);

      // 7. Enviar correos
      if (esCorreoValido(beneficiario.correo)) {
        enviarCorreoEstilizado(
          beneficiario.correo,
          "Solicitud de Préstamo - Sindicato SLIM n°3",
          "Solicitud de Préstamo Ingresada",
          "Hola <strong>" + beneficiario.nombre + "</strong>, se ha ingresado exitosamente una solicitud de préstamo a tu nombre.",
          {
            "FECHA SOLICITUD": Utilities.formatDate(fechaSolicitud, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm"),
            "RUT": formatRutServer(beneficiario.rut),
            "NOMBRE": beneficiario.nombre,
            "TIPO PRÉSTAMO": tipo,
            "MONTO": montoTexto,
            "CUOTAS": cuotas,
            "MEDIO PAGO": medioPago,
            "FECHA TÉRMINO": Utilities.formatDate(fechaTermino, Session.getScriptTimeZone(), "dd/MM/yyyy"),
            "GESTION": gestion,
            "NOMBRE DIRIGENTE": nomDirigente || ""
          },
          "#2563eb"
        );
      }

      if (esGestionDirigente && esCorreoValido(correoDirigente) && correoDirigente !== beneficiario.correo) {
        enviarCorreoEstilizado(
          gestor.correo,
          "Respaldo Gestión Préstamo - Sindicato SLIM n°3",
          "Solicitud de Préstamo Creada",
          "Has ingresado una solicitud de préstamo para el socio <strong>" + beneficiario.nombre + "</strong>.",
          {
            "FECHA SOLICITUD": Utilities.formatDate(fechaSolicitud, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm"),
            "RUT SOCIO": formatRutServer(beneficiario.rut),
            "NOMBRE SOCIO": beneficiario.nombre,
            "TIPO PRÉSTAMO": tipo,
            "MONTO": montoTexto,
            "CUOTAS": cuotas,
            "FECHA TÉRMINO": Utilities.formatDate(fechaTermino, Session.getScriptTimeZone(), "dd/MM/yyyy"),
            "GESTION": "Dirigente"
          },
          "#475569"
        );
      }

      return { success: true, message: "Solicitud creada exitosamente." };

    } catch (e) {
      return { success: false, message: "Error al solicitar: " + e.toString() };
    } finally {
      lock.releaseLock();
    }
  } else {
    return { success: false, message: "Servidor ocupado." };
  }
}

/**
 * Sincronización automática: Validación → BD → Notificación
 * Trigger: diario a las 8 AM
 */
function procesarValidacionPrestamos() {
  var lock = LockService.getScriptLock();
  if (lock.tryLock(60000)) {
    try {
      var ss = getSpreadsheet('PRESTAMOS');
      var sheetValidacion = ss.getSheetByName(CONFIG.HOJAS.VALIDACION_PRESTAMOS);
      var sheetBD = getSheet('PRESTAMOS', 'PRESTAMOS');

      if (!sheetValidacion) {
        console.warn("⚠️ La hoja 'Validación-Prestamos' no existe. Creándola...");
        var nuevaHoja = ss.insertSheet(CONFIG.HOJAS.VALIDACION_PRESTAMOS);
        nuevaHoja.appendRow(["ID","Fecha","RUT","Nombre","Validación","Observación","Nombre Informe"]);
        return;
      }
      if (!sheetBD) { console.error("❌ No se encontró la hoja BD_PRESTAMOS."); return; }

      var dataValidacion = sheetValidacion.getDataRange().getValues();
      var dataBD = sheetBD.getDataRange().getValues();
      var COL_BD = CONFIG.COLUMNAS.PRESTAMOS;
      var VAL_COL = { ID: 0, VALIDACION: 4, OBS: 5 };
      var COL_INFORME = 14;
      var procesadosCount = 0;

      for (var i = 1; i < dataValidacion.length; i++) {
        var idSolicitud = String(dataValidacion[i][VAL_COL.ID]).trim();
        var resultadoValidacion = String(dataValidacion[i][VAL_COL.VALIDACION]).toUpperCase().trim();
        var observacionAdmin = String(dataValidacion[i][VAL_COL.OBS]);

        if (!idSolicitud || (resultadoValidacion !== "ACEPTADO" && resultadoValidacion !== "RECHAZADO")) continue;

        for (var j = 1; j < dataBD.length; j++) {
          var idBD = String(dataBD[j][COL_BD.ID]).trim();
          var informeEnviado = String(dataBD[j][COL_INFORME]);

          if (idBD !== idSolicitud) continue;
          if (informeEnviado === "OK") { console.log('ℹ️ ID ' + idSolicitud + ': Ya procesado.'); continue; }

          var nuevoEstado = resultadoValidacion === "ACEPTADO" ? "Vigente" : "Rechazado";
          var tituloCorreo = resultadoValidacion === "ACEPTADO" ? "Solicitud Aprobada" : "Solicitud Rechazada";
          var colorCorreo  = resultadoValidacion === "ACEPTADO" ? "#15803d" : "#b91c1c";
          var mensajeIntro = resultadoValidacion === "ACEPTADO"
            ? "Nos complace informarte que tu solicitud de préstamo ha sido <strong>APROBADA</strong> por la empresa."
            : "Te informamos que tu solicitud de préstamo ha sido <strong>RECHAZADA</strong> por la empresa.";

          sheetBD.getRange(j + 1, COL_BD.ESTADO + 1).setValue(nuevoEstado);

          var correoUsuario = dataBD[j][COL_BD.CORREO];
          var nombreUsuario = dataBD[j][COL_BD.NOMBRE];

          if (esCorreoValido(correoUsuario)) {
            var fechaTerminoStr = "S/D";
            var fechaTerminoRaw = dataBD[j][COL_BD.FECHA_TERMINO];
            if (fechaTerminoRaw) {
              try {
                var ftObj = new Date(fechaTerminoRaw);
                if (!isNaN(ftObj.getTime())) fechaTerminoStr = Utilities.formatDate(ftObj, Session.getScriptTimeZone(), "dd/MM/yyyy");
              } catch (e) {}
            }

            enviarCorreoEstilizado(
              correoUsuario,
              "Resultado Solicitud Préstamo - Sindicato SLIM n°3",
              tituloCorreo,
              "Hola <strong>" + nombreUsuario + "</strong>, " + mensajeIntro,
              {
                "FECHA SOLICITUD": Utilities.formatDate(new Date(dataBD[j][COL_BD.FECHA]), Session.getScriptTimeZone(), "dd/MM/yyyy"),
                "RUT": formatRutServer(dataBD[j][COL_BD.RUT]),
                "NOMBRE": nombreUsuario,
                "TIPO PRÉSTAMO": dataBD[j][COL_BD.TIPO],
                "MONTO": dataBD[j][COL_BD.MONTO] || "$0",
                "ESTADO": nuevoEstado.toUpperCase(),
                "FECHA TÉRMINO": fechaTerminoStr,
                "OBSERVACIÓN": observacionAdmin || "Sin observaciones",
                "RESULTADO": resultadoValidacion
              },
              colorCorreo
            );
            sheetBD.getRange(j + 1, COL_INFORME + 1).setValue("OK");
            procesadosCount++;
          } else {
            sheetBD.getRange(j + 1, COL_INFORME + 1).setValue("ERROR_NO_MAIL");
          }
          break;
        }
      }

      console.log(procesadosCount > 0
        ? '✅ Resumen final: ' + procesadosCount + ' solicitudes nuevas procesadas.'
        : 'ℹ️ No hay solicitudes nuevas para procesar.');

    } catch (e) {
      console.error("❌ Error en procesarValidacionPrestamos: " + e.toString());
    } finally {
      lock.releaseLock();
    }
  }
}

/**
 * Obtiene el historial de préstamos de un usuario
 */
function obtenerHistorialPrestamos(rutInput) {
  try {
    var sheet = getSheet('PRESTAMOS', 'PRESTAMOS');
    var COL = CONFIG.COLUMNAS.PRESTAMOS;

    if (!sheet) return { success: false, message: "Hoja no encontrada" };
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return { success: true, registros: [] };

    var lastCol = sheet.getLastColumn();
    var data = sheet.getRange(2, 1, lastRow - 1, lastCol).getDisplayValues();
    var rutLimpio = cleanRut(rutInput);
    var registros = [];

    for (var i = 0; i < data.length; i++) {
      var row = data[i];
      if (!row[COL.RUT]) continue;
      if (cleanRut(row[COL.RUT]) !== rutLimpio) continue;

      var fechaTerminoStr = "S/D";
      var ftRaw = row[COL.FECHA_TERMINO];
      if (ftRaw) {
        try {
          var d = new Date(ftRaw);
          fechaTerminoStr = !isNaN(d.getTime())
            ? Utilities.formatDate(d, Session.getScriptTimeZone(), "dd/MM/yyyy")
            : String(ftRaw).split(' ')[0];
        } catch(e) { fechaTerminoStr = String(ftRaw).split(' ')[0]; }
      }

      registros.push({
        id:           row[COL.ID]           || "",
        fecha:        formatearFechaConHora(row[COL.FECHA]) || "",
        tipo:         row[COL.TIPO]         || "Préstamo",
        monto:        row[COL.MONTO]        || "$0",
        cuotas:       row[COL.CUOTAS]       || "S/D",
        medio:        row[COL.MEDIO_PAGO]   || "S/D",
        estado:       row[COL.ESTADO]       || "Solicitado",
        observacion:  row[COL.OBSERVACION]  || "",
        fechaTermino: formatearFechaSinHora(row[COL.FECHA_TERMINO]),
        gestion:      row[COL.GESTION]      || "Socio",
        nomDirigente: row[COL.NOMBRE_DIRIGENTE] || ""
      });
    }

    registros.reverse();
    return { success: true, registros: registros };

  } catch (e) {
    Logger.log('❌ ERROR en obtenerHistorialPrestamos: ' + e.toString());
    return { success: false, message: "Error: " + e.toString() };
  }
}

/**
 * Elimina una solicitud de préstamo con respaldo histórico
 */
function eliminarSolicitud(idSolicitud) {
  var lock = LockService.getScriptLock();
  if (lock.tryLock(30000)) {
    try {
      var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEETS.PRESTAMOS);
      var sheet = ss.getSheetByName("BD_PRESTAMOS");
      var data = sheet.getDataRange().getValues();
      var COL = CONFIG.COLUMNAS.PRESTAMOS;

      for (var i = 1; i < data.length; i++) {
        if (String(data[i][COL.ID]) === String(idSolicitud)) {
          var sheetEliminados = ss.getSheetByName("Registros-eliminados");
          if (sheetEliminados) {
            sheetEliminados.appendRow(data[i]);
          } else {
            return { success: false, message: "Error crítico: No existe la hoja de respaldo." };
          }
          sheet.deleteRow(i + 1);
          return { success: true, message: "Registro eliminado y respaldado correctamente." };
        }
      }
      return { success: false, message: "No encontrado." };
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
 * Modifica cuotas/medio de pago de una solicitud en estado "Solicitado"
 */
function modificarSolicitud(idSolicitud, nuevasCuotas, nuevoMedio) {
  var lock = LockService.getScriptLock();
  if (lock.tryLock(30000)) {
    try {
      var sheet = getSheet('PRESTAMOS', 'PRESTAMOS');
      var data = sheet.getDataRange().getValues();
      var COL = CONFIG.COLUMNAS.PRESTAMOS;

      for (var i = 1; i < data.length; i++) {
        if (String(data[i][COL.ID]) === String(idSolicitud)) {
          var estado = String(data[i][COL.ESTADO]);
          if (estado !== "Solicitado") return { success: false, message: "No se puede editar. Estado: " + estado };

          var fechaSolicitud = new Date(data[i][COL.FECHA]);
          var diaSolicitud = fechaSolicitud.getDate();
          var fechaInicioPago = new Date(fechaSolicitud);
          if (diaSolicitud > 24) fechaInicioPago.setMonth(fechaInicioPago.getMonth() + 1);

          var fechaTermino = new Date(fechaInicioPago);
          fechaTermino.setMonth(fechaTermino.getMonth() + parseInt(nuevasCuotas));
          fechaTermino = new Date(fechaTermino.getFullYear(), fechaTermino.getMonth() + 1, 0);

          sheet.getRange(i + 1, COL.FECHA_TERMINO + 1).setValue(fechaTermino);
          return { success: true, message: "Modificado correctamente." };
        }
      }
      return { success: false, message: "No encontrado." };
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
 * Trigger diario a las 8 AM: cambia préstamos "Vigente" → "Pagado" si venció fecha de término
 */
function verificarCambiosPrestamos() {
  try {
    var sheet = getSheet('PRESTAMOS', 'PRESTAMOS');
    if (!sheet) { console.error("❌ No se pudo acceder a la hoja BD_PRESTAMOS"); return { success: false }; }

    var data = sheet.getDataRange().getValues();
    var COL = CONFIG.COLUMNAS.PRESTAMOS;
    var hoy = new Date();
    hoy.setHours(0, 0, 0, 0);
    var prestamosActualizados = 0;

    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (String(row[COL.ESTADO]).trim() !== "Vigente") continue;
      if (!row[COL.FECHA_TERMINO]) continue;

      var fechaTermino;
      try {
        fechaTermino = new Date(row[COL.FECHA_TERMINO]);
        fechaTermino.setHours(0, 0, 0, 0);
      } catch (e) { continue; }
      if (isNaN(fechaTermino.getTime())) continue;
      if (hoy <= fechaTermino) continue;

      sheet.getRange(i + 1, COL.ESTADO + 1).setValue("Pagado");

      var correo = row[COL.CORREO];
      var nombre = row[COL.NOMBRE];
      var tipo   = row[COL.TIPO];
      var monto  = row[COL.MONTO];
      var cuotas = row[COL.CUOTAS];
      var idPrestamo = row[COL.ID];

      if (esCorreoValido(correo)) {
        try {
          enviarCorreoEstilizado(
            correo,
            "Préstamo Completado - Sindicato SLIM n°3",
            "Préstamo Finalizado",
            "Hola <strong>" + nombre + "</strong>, tu préstamo ha sido completado exitosamente.",
            {
              "ID": idPrestamo,
              "TIPO PRÉSTAMO": tipo,
              "MONTO": monto,
              "CUOTAS": cuotas,
              "ESTADO": "PAGADO",
              "FECHA TÉRMINO": Utilities.formatDate(fechaTermino, Session.getScriptTimeZone(), 'dd/MM/yyyy'),
              "FECHA FINALIZACIÓN": Utilities.formatDate(hoy, Session.getScriptTimeZone(), 'dd/MM/yyyy')
            },
            "#10b981"
          );
        } catch (mailError) { console.error('⚠️ Error enviando correo: ' + mailError); }
      }
      prestamosActualizados++;
    }

    console.log(prestamosActualizados > 0
      ? '📊 RESUMEN: ' + prestamosActualizados + ' préstamo(s) actualizado(s) a "Pagado"'
      : 'ℹ️ No hay préstamos que actualizar.');

    return { success: true, prestamosActualizados: prestamosActualizados };

  } catch (e) {
    console.error("❌ Error verificando préstamos: " + e.toString());
    return { success: false, error: e.toString() };
  }
}

// ==========================================
// SWITCH MÓDULO PRÉSTAMOS
// ==========================================

function obtenerEstadoSwitchPrestamos() {
  try {
    var estado = PropertiesService.getScriptProperties().getProperty('prestamos_habilitado');
    return { success: true, habilitado: (estado === null || estado === 'true') };
  } catch (e) {
    return { success: true, habilitado: true };
  }
}

function toggleSwitchPrestamos(estado) {
  try {
    PropertiesService.getScriptProperties().setProperty('prestamos_habilitado', estado ? 'true' : 'false');
    return { success: true };
  } catch (e) {
    return { success: false, message: 'Error: ' + e.toString() };
  }
}
