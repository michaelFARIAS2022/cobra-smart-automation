const CONFIG = {
  APP_NAME: "Generador Central de Borradores",

  LIMITE_POR_EJECUCION: 60,

  HEADERS: {
    NOMBRE: "Nombre",
    DNI: "DNI",
    CORREO: "Correo",
    ESTADO: "Estado",
    GESTOR: "Gestor",
    WHATSAPP: "WhatsApp",
    ENTIDAD: "Entidad",
  },

  META_LAST_RUN_CELL: "H1",
  META_CREATED_CELL: "I1",

  STATUS_PENDIENTE: "PENDIENTE",
  STATUS_SIN_CORREO: "SIN CORREO",
  STATUS_BORRADOR_OK: "BORRADOR CREADO",
  STATUS_ERROR: "ERROR",
  STATUS_VACIO: "DATOS INCOMPLETOS",

  PROP_MASTER_ENABLED: "MASTER_ENABLED",
  PROP_LICENSES_JSON: "LICENSES_JSON",
  PROP_TEMPLATE_SUBJECTS: "TEMPLATE_SUBJECTS_JSON",
  PROP_TEMPLATE_BODIES: "TEMPLATE_BODIES_JSON",
  PROP_LAST_RUN_PREFIX: "LAST_RUN_",

  DEFAULT_GESTOR_GENERIC: "Departamento de Legales",
  DEFAULT_ENTIDAD_GENERIC: "Entidad",
};

/***************************************
 * WEB APP
 ***************************************/

function doGet(e) {
  const sid = extractSpreadsheetId_(e);
  const gid = extractSheetGid_(e);

  const template = HtmlService.createTemplateFromFile("Run");
  template.appName = CONFIG.APP_NAME;
  template.spreadsheetId = sid;
  template.sheetGid = gid;

  return template
    .evaluate()
    .setTitle(CONFIG.APP_NAME)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function extractSpreadsheetId_(e) {
  let sid = "";

  if (e && e.parameter) {
    sid = e.parameter.sid || e.parameter.id || "";
  }

  if (!sid && e && e.queryString) {
    const match = String(e.queryString).match(/(?:^|&)(?:sid|id)=([^&]+)/i);
    if (match && match[1]) {
      sid = decodeURIComponent(match[1]);
    }
  }

  sid = String(sid || "")
    .trim()
    .replace(/^["'(]+/, "")
    .replace(/["')]+$/, "");

  return sid;
}

/***************************************
 * PROCESO PRINCIPAL
 ***************************************/

function runRemoteDraftJob(spreadsheetId, sheetGid) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    validateMasterSwitch_();
    validateSpreadsheetId_(spreadsheetId);

    const activeEmail = getActiveUserEmailSafe_();
    const tempUserKey = Session.getTemporaryActiveUserKey();
    const license = validateLicense_(spreadsheetId, activeEmail, tempUserKey);

    let enviadosHoy = getPlanDailyCount_(spreadsheetId);

    if (enviadosHoy >= CONFIG.LIMITE_POR_EJECUCION) {
      throw new Error(
        "Límite diario alcanzado (" +
          CONFIG.LIMITE_POR_EJECUCION +
          ") para esta planilla.",
      );
    }

    const ss = SpreadsheetApp.openById(spreadsheetId);

    let sheet = null;
    if (sheetGid) {
      sheet = ss.getSheets().find(function (s) {
        return String(s.getSheetId()) === String(sheetGid);
      });
    }

    if (!sheet) {
      sheet = ss.getActiveSheet();
    }
    ////////////////////////////////////
    const values = sheet.getDataRange().getValues();

    if (!values || values.length < 2) {
      throw new Error("La planilla no contiene datos para procesar.");
    }

    const headerMap = buildHeaderMap_(values[0]);
    validateRequiredHeaders_(headerMap);

    const asuntos = getSubjectTemplates_();
    const bodies = getBodyTemplates_();

    let creados = 0;
    let errores = 0;
    let sinCorreo = 0;

    sheet.getRange(CONFIG.META_LAST_RUN_CELL).setValue(new Date());
    sheet.getRange(CONFIG.META_CREATED_CELL).setValue(0);
    SpreadsheetApp.flush();

    for (let r = 1; r < values.length; r++) {
      if (creados >= CONFIG.LIMITE_POR_EJECUCION) break;
      if (enviadosHoy >= CONFIG.LIMITE_POR_EJECUCION) break;

      const row = values[r];

      const estadoIndex = headerMap[CONFIG.HEADERS.ESTADO];
      const estado =
        estadoIndex !== undefined ? normalize_(row[estadoIndex]) : "";

      // Procesa si Estado está vacío o PENDIENTE
      if (estado && estado !== CONFIG.STATUS_PENDIENTE) {
        continue;
      }

      const nombre = safeCell_(row[headerMap[CONFIG.HEADERS.NOMBRE]]);

      const dniIndex = headerMap[CONFIG.HEADERS.DNI];
      const dni = dniIndex !== undefined ? normalizeDni_(row[dniIndex]) : "";

      const correoIndex = headerMap[CONFIG.HEADERS.CORREO];
      const correoRaw =
        correoIndex !== undefined ? safeCell_(row[correoIndex]) : "";

      const gestorIndex = headerMap[CONFIG.HEADERS.GESTOR];
      const gestor =
        gestorIndex !== undefined
          ? safeCell_(row[gestorIndex]) ||
            license.defaultGestor ||
            CONFIG.DEFAULT_GESTOR_GENERIC
          : license.defaultGestor || CONFIG.DEFAULT_GESTOR_GENERIC;

      const whatsappIndex = headerMap[CONFIG.HEADERS.WHATSAPP];
      const numWhatsappFila = normalizeWhatsapp_(
        (whatsappIndex !== undefined ? safeCell_(row[whatsappIndex]) : "") ||
          license.defaultWhatsapp ||
          "5490000000000",
      );

      const entidadIndex = headerMap[CONFIG.HEADERS.ENTIDAD];
      const entidad =
        entidadIndex !== undefined
          ? safeCell_(row[entidadIndex]) ||
            license.defaultEntidad ||
            CONFIG.DEFAULT_ENTIDAD_GENERIC
          : license.defaultEntidad || CONFIG.DEFAULT_ENTIDAD_GENERIC;

      const estadoCell =
        estadoIndex !== undefined
          ? sheet.getRange(r + 1, estadoIndex + 1)
          : null;

      // Campos críticos
      if (!nombre || !numWhatsappFila) {
        if (estadoCell) {
          estadoCell.setValue(CONFIG.STATUS_VACIO + ": faltan Nombre/WhatsApp");
        }
        errores++;
        continue;
      }

      const emails = obtenerCorreos_(correoRaw);

      // Sin correo -> solo salta esa fila
      if (!correoRaw || emails.length === 0) {
        if (estadoCell) {
          estadoCell.setValue(CONFIG.STATUS_SIN_CORREO);
        }
        sinCorreo++;
        continue;
      }

      try {
        const idUnico = generateId_();
        const dniOculto = ocultarDni_(dni);

        const subjectBase = asuntos[Math.floor(Math.random() * asuntos.length)];
        const bodyBase = bodies[Math.floor(Math.random() * bodies.length)];

        const subject = renderTemplate_(subjectBase, {
          Nombre: nombre,
          DNI: dniOculto,
          GESTOR: gestor,
          NUM_WSP: numWhatsappFila,
          ENTIDAD: entidad,
          ID_UNICO: idUnico,
        });

        const bodyHtml = renderTemplate_(bodyBase, {
          Nombre: nombre,
          DNI: dni,
          GESTOR: gestor,
          NUM_WSP: numWhatsappFila,
          ENTIDAD: entidad,
          ID_UNICO: idUnico,
        });

        GmailApp.createDraft(
          emails.join(","),
          subject,
          "Este borrador requiere HTML.",
          { htmlBody: bodyHtml },
        );

        creados++;

        enviadosHoy = incrementPlanDailyCount_(spreadsheetId);

        if (estadoCell) {
          estadoCell.setValue(
            CONFIG.STATUS_BORRADOR_OK + " - " + formatDateTime_(new Date()),
          );
        }

        sheet.getRange(CONFIG.META_LAST_RUN_CELL).setValue(new Date());
        sheet.getRange(CONFIG.META_CREATED_CELL).setValue(creados);

        SpreadsheetApp.flush();
        Utilities.sleep(250);
      } catch (errRow) {
        if (estadoCell) {
          estadoCell.setValue(
            CONFIG.STATUS_ERROR + ": " + truncate_(errRow.message, 180),
          );
        }
        errores++;
        console.error(
          JSON.stringify({
            event: "row_error",
            spreadsheetId,
            row: r + 1,
            email: activeEmail || "",
            tempUserKey,
            error: String(errRow.message || errRow),
          }),
        );
      }
    }

    console.log(
      JSON.stringify({
        event: "run_finished",
        spreadsheetId,
        activeEmail: activeEmail || "",
        tempUserKey,
        created: creados,
        errors: errores,
        noEmail: sinCorreo,
      }),
    );

    return {
      ok: true,
      created: creados,
      errors: errores,
      noEmail: sinCorreo,
      executionLimit: CONFIG.LIMITE_POR_EJECUCION,
      message: "Proceso finalizado. Borradores creados: " + creados,
    };
  } finally {
    lock.releaseLock();
  }
}

/***************************************
 * LIMITE DIARIO POR PLANILLA
 ***************************************/

function getPlanDailyKeys_(spreadsheetId) {
  return {
    fecha: "fecha_" + spreadsheetId,
    contador: "contador_" + spreadsheetId,
  };
}

function getPlanDailyCount_(spreadsheetId) {
  const props = PropertiesService.getScriptProperties();
  const keys = getPlanDailyKeys_(spreadsheetId);
  const hoy = new Date().toDateString();

  const fechaGuardada = props.getProperty(keys.fecha);
  let contador = parseInt(props.getProperty(keys.contador) || "0", 10);

  if (fechaGuardada !== hoy) {
    props.setProperty(keys.fecha, hoy);
    props.setProperty(keys.contador, "0");
    return 0;
  }

  return contador;
}

function incrementPlanDailyCount_(spreadsheetId) {
  const props = PropertiesService.getScriptProperties();
  const keys = getPlanDailyKeys_(spreadsheetId);
  const hoy = new Date().toDateString();

  const fechaGuardada = props.getProperty(keys.fecha);
  let contador = parseInt(props.getProperty(keys.contador) || "0", 10);

  if (fechaGuardada !== hoy) {
    props.setProperty(keys.fecha, hoy);
    props.setProperty(keys.contador, "1");
    return 1;
  }

  contador++;
  props.setProperty(keys.contador, contador.toString());
  return contador;
}

/***************************************
 * AUTORIZACION / LICENCIAS
 ***************************************/

function validateMasterSwitch_() {
  const props = PropertiesService.getScriptProperties();
  const enabled = (
    props.getProperty(CONFIG.PROP_MASTER_ENABLED) || "true"
  ).toLowerCase();

  if (enabled !== "true") {
    throw new Error(
      "El sistema está desactivado temporalmente por administración.",
    );
  }
}

function validateLicense_(spreadsheetId, activeEmail, tempUserKey) {
  const props = PropertiesService.getScriptProperties();
  const raw = props.getProperty(CONFIG.PROP_LICENSES_JSON) || "{}";
  const licenses = JSON.parse(raw);
  const lic = licenses[spreadsheetId];

  if (!lic) {
    throw new Error("Esta planilla no está autorizada.");
  }

  if (!lic.enabled) {
    throw new Error("La licencia de esta planilla está desactivada.");
  }

  if (
    lic.allowedEmail &&
    activeEmail &&
    lic.allowedEmail.toLowerCase() !== activeEmail.toLowerCase()
  ) {
    throw new Error(
      "El usuario autenticado no coincide con el autorizado para esta planilla.",
    );
  }

  return lic;
}

function adminSetMasterEnabled(enabled) {
  PropertiesService.getScriptProperties().setProperty(
    CONFIG.PROP_MASTER_ENABLED,
    String(!!enabled),
  );
}

function adminUpsertLicense(spreadsheetId, options) {
  const props = PropertiesService.getScriptProperties();
  const raw = props.getProperty(CONFIG.PROP_LICENSES_JSON) || "{}";
  const licenses = JSON.parse(raw);

  licenses[spreadsheetId] = Object.assign(
    {},
    licenses[spreadsheetId] || {},
    options || {},
  );
  props.setProperty(
    CONFIG.PROP_LICENSES_JSON,
    JSON.stringify(licenses, null, 2),
  );
}

function adminDisableLicense(spreadsheetId) {
  adminUpsertLicense(spreadsheetId, { enabled: false });
}

function adminEnableLicense(spreadsheetId) {
  adminUpsertLicense(spreadsheetId, { enabled: true });
}

/***************************************
 * PLANTILLAS
 ***************************************/

function seedDefaultTemplates_() {
  const props = PropertiesService.getScriptProperties();

  const asuntos = [
    "Instancia final de gestión – Cuenta [ENTIDAD] - [Nombre] / [DNI]",
    "Último aviso antes de instancia judicial – Cuenta [ENTIDAD] - [Nombre] / [DNI]",
    "Regularización pendiente – Cuenta [ENTIDAD] - [Nombre] / [DNI]",
    "Notificación importante sobre su cuenta – Cuenta [ENTIDAD] - [Nombre] / [DNI]",
    "Aviso urgente – Cuenta [ENTIDAD] - [Nombre] / [DNI]",
  ];
  const bodyTemplates = [
    "<p>Estimado/a <strong>[Nombre]</strong>,</p>" +
      "<p>Le informamos que su cuenta con <strong>[ENTIDAD]</strong> se encuentra en etapa avanzada de recupero de mora.</p>" +
      "<p>A la fecha, registra saldo pendiente que, de no regularizarse, será necesario avanzar con la siguiente instancia de gestión correspondiente.</p>" +
      "<p>Antes de que la instancia actual avance, aún tiene la posibilidad de acceder a una cancelación con descuentos especiales.</p>" +
      "<hr>" +
      "<p><strong>IMPORTANTE:</strong></p>" +
      "<ul>" +
      "<li>La falta de respuesta a este aviso se interpretará como falta de intención en solucionar su situación de manera extrajudicial y se avanzará con la etapa correspondiente</li>" +
      "<li>Los costos derivados de un eventual proceso judicial estarán a su cargo, incrementando el monto total de la deuda.</li>" +
      "<li>Para una correcta gestión de su legajo, también puede responder este correo con la palabra: <strong>AHORA</strong>.</li>" +
      "</ul>" +
      "<p><strong>Para recibir el detalle actualizado y definir una solución, puede comunicarse a la brevedad por nuestros canales:</strong></p>" +
      "<ul>" +
      "<li>Tel: 358 4640734</li>" +
      '<li>WhatsApp: <a href="https://wa.me/[NUM_WSP]" target="_blank">Iniciar conversación</a></li>' +
      "</ul>" +
      '<div style="margin: 20px 0;">' +
      '<a href="https://wa.me/[NUM_WSP]?text=Gracias%20por%20la%20notificacion.%20Confirmo%20recepcion%20para%20detener%20avisos%20de%20deuda%20de%20-[Nombre]-" ' +
      'style="background-color: #2c3e50; color: white; padding: 12px 18px; text-decoration: none; border-radius: 4px; font-weight: bold; display: inline-block; font-family: Arial, sans-serif; font-size: 13px;">' +
      "✓ Confirmar recepción y detener avisos</a>" +
      "</div>" +
      "<p><strong>Aguardamos su pronta respuesta.</strong></p>" +
      "<p>Saludos cordiales,</p>" +
      "<p><strong>Estudio Demo</strong><br>" +
      "Departamento de Legales | [GESTOR]<br>" +
      "<strong>Empresa Demo</strong><br>" +
      "Dra. Lucrecia del Rosario Córdoba<br>" +
      "Direccion Demo<br>" +
      "Tel: 5490000000000<br>" +
      "WhatsApp: [NUM_WSP]<br>" +
      '<a href="https://www.ejemplo.com" target="_blank">www.ejemplo.com</a></p>' +
      '<p style="color:#ffffff; font-size:1px; margin:0;">Ref-ID: [ID_UNICO]</p>',
  ];

  props.setProperty(CONFIG.PROP_TEMPLATE_SUBJECTS, JSON.stringify(asuntos));
  props.setProperty(CONFIG.PROP_TEMPLATE_BODIES, JSON.stringify(bodyTemplates));
}

function getSubjectTemplates_() {
  const props = PropertiesService.getScriptProperties();
  const raw = props.getProperty(CONFIG.PROP_TEMPLATE_SUBJECTS);

  if (!raw) {
    seedDefaultTemplates_();
    return JSON.parse(
      PropertiesService.getScriptProperties().getProperty(
        CONFIG.PROP_TEMPLATE_SUBJECTS,
      ),
    );
  }

  return JSON.parse(raw);
}

function getBodyTemplates_() {
  const props = PropertiesService.getScriptProperties();
  const raw = props.getProperty(CONFIG.PROP_TEMPLATE_BODIES);

  if (!raw) {
    seedDefaultTemplates_();
    return JSON.parse(
      PropertiesService.getScriptProperties().getProperty(
        CONFIG.PROP_TEMPLATE_BODIES,
      ),
    );
  }

  return JSON.parse(raw);
}

/***************************************
 * UTILIDADES DE PLANILLA
 ***************************************/

function buildHeaderMap_(headersRow) {
  const map = {};
  headersRow.forEach(function (header, idx) {
    const key = String(header || "").trim();
    if (key) map[key] = idx;
  });
  return map;
}

function validateRequiredHeaders_(headerMap) {
  const required = [CONFIG.HEADERS.NOMBRE, CONFIG.HEADERS.WHATSAPP];

  const missing = required.filter(function (name) {
    return !(name in headerMap);
  });

  if (missing.length) {
    throw new Error("Faltan columnas obligatorias: " + missing.join(", "));
  }
}

function validateSpreadsheetId_(spreadsheetId) {
  if (!spreadsheetId) {
    throw new Error("No se recibió el ID de la planilla.");
  }
}

/***************************************
 * UTILIDADES DE DATOS
 ***************************************/

function safeCell_(v) {
  return v === null || v === undefined ? "" : String(v).trim();
}
function normalize_(v) {
  return safeCell_(v).toUpperCase();
}

function normalizeDni_(dni) {
  return safeCell_(dni).replace(/[^\d]/g, "");
}

function ocultarDni_(dni) {
  const limpio = normalizeDni_(dni);
  if (!limpio) return "";
  if (limpio.length <= 3) return limpio;
  return "X".repeat(limpio.length - 3) + limpio.slice(-3);
}

function normalizeWhatsapp_(num) {
  return safeCell_(num).replace(/[^\d]/g, "");
}

function isValidEmail_(email) {
  if (!email) return false;
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(email).trim());
}

function obtenerCorreos_(celda) {
  if (!celda) return [];
  const unique = {};

  return String(celda)
    .split(/[,;]+/)
    .map(function (s) {
      return s.trim();
    })
    .filter(isValidEmail_)
    .filter(function (mail) {
      const key = mail.toLowerCase();
      if (unique[key]) return false;
      unique[key] = true;
      return true;
    });
}

function generateId_() {
  return Utilities.getUuid().replace(/-/g, "").substring(0, 12).toUpperCase();
}

function renderTemplate_(template, values) {
  let out = String(template || "");
  Object.keys(values).forEach(function (key) {
    const re = new RegExp("\\[" + escapeRegExp_(key) + "\\]", "g");
    out = out.replace(re, values[key] == null ? "" : String(values[key]));
  });
  return out;
}

function escapeRegExp_(text) {
  return String(text).replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

function truncate_(text, len) {
  text = String(text || "");
  return text.length > len ? text.substring(0, len - 3) + "..." : text;
}

function formatDateTime_(date) {
  return Utilities.formatDate(
    date,
    Session.getScriptTimeZone(),
    "yyyy-MM-dd HH:mm:ss",
  );
}

function getActiveUserEmailSafe_() {
  try {
    return Session.getActiveUser().getEmail() || "";
  } catch (err) {
    return "";
  }
}

/***************************************
 * UTILIDADES MANUALES
 ***************************************/

function cargarPlantillas() {
  seedDefaultTemplates_();
}

function altaGestor() {
  adminUpsertLicense("SPREADSHEET_ID_DEMO", {
    enabled: true,
    allowedEmail: "",
    defaultGestor: "Gestor Demo",
    defaultWhatsapp: "5490000000000",
    defaultEntidad: "Entidad Demo",
  });
}

function desactivarPlanilla() {
  adminDisableLicense("PEGAR_ID");
}

function activarPlanilla() {
  adminEnableLicense("PEGAR_ID");
}

function resetearContadorPlanilla() {
  const spreadsheetId = "SPREADSHEET_ID_DEMO";
  const props = PropertiesService.getScriptProperties();
  props.deleteProperty("fecha_" + spreadsheetId);
  props.deleteProperty("contador_" + spreadsheetId);
}

function extractSheetGid_(e) {
  let gid = "";

  if (e && e.parameter) {
    gid = e.parameter.gid || "";
  }

  if (!gid && e && e.queryString) {
    const match = String(e.queryString).match(/(?:^|&)gid=([^&]+)/i);
    if (match && match[1]) {
      gid = decodeURIComponent(match[1]);
    }
  }

  return String(gid || "").trim();
}
