// =========================================================================
// BACKEND SEGURO: Google Apps Script para CeluControl
// =========================================================================

function doGet(e) {
  return handleRequest(e);
}

function doPost(e) {
  return handleRequest(e);
}

function handleRequest(e) {
  let params = {};
  if (e && e.postData && e.postData.contents) {
    try {
      params = JSON.parse(e.postData.contents);
    } catch (err) {
      params = {};
    }
  } else if (e && e.parameter) {
    params = e.parameter;
  }

  const action = params.action || 'getData';

  // Solo las acciones de escritura toman el lock; las lecturas corren en paralelo.
  const accionesLectura = ['getData', 'login', 'getHistorialCompleto', 'getHistorialIMEI', 'getRegalosDelMes'];
  const requiereLock = accionesLectura.indexOf(action) === -1;
  let lock = null;

  if (requiereLock) {
    lock = LockService.getScriptLock();
    try {
      lock.waitLock(5000);
    } catch (err) {
      return jsonResponse({ success: false, error: "Servidor ocupado, intente de nuevo." });
    }
  }

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    initSheets(ss);

    const result = { success: true };

    if (action !== 'getData' && action !== 'login') {
      validateUserAuth(ss, params.userId || params.vendedor_id, action);
    }

    switch (action) {
      case 'getData':
        result.data = getAllData(ss);
        break;

      case 'getHistorialCompleto':
        result.data = getHistorialCompleto(ss);
        break;

      case 'login':
        result.data = loginUser(ss, params.email, params.password);
        break;

      case 'registrarVenta':
        result.data = registrarVenta(ss, params);
        break;

      case 'editarVenta':
        result.data = editarVenta(ss, params);
        break;

      case 'anularVenta':
        result.data = anularVenta(ss, params.ventaId, params.userId);
        break;

      case 'getHistorialIMEI':
        result.data = getHistorialIMEI(ss, params.imei);
        break;

      case 'getRegalosDelMes':
        result.data = getRegalosDelMes(ss);
        break;

      case 'registrarPerdida':
        result.data = registrarPerdida(ss, params);
        break;

      case 'agregarCelular':
        result.data = agregarCelular(ss, params);
        break;

      case 'guardarAccesorio':
        result.data = guardarAccesorio(ss, params);
        break;

      case 'actualizarTasa':
        result.data = actualizarTasa(ss, params.tasa);
        break;

      case 'agregarModelo':
        result.data = agregarModelo(ss, params);
        break;

      case 'agregarReferenciaRetoma':
        result.data = agregarReferenciaRetoma(ss, params);
        break;

      case 'crearEmpleado':
        result.data = crearEmpleado(ss, params);
        break;

      default:
        throw new Error("Acción desconocida: " + action);
    }

    if (lock) lock.releaseLock();
    return jsonResponse(result);

  } catch (err) {
    if (lock) lock.releaseLock();
    return jsonResponse({ success: false, error: err.toString() });
  }
}

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function hashPassword(password) {
  const rawHash = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, password, Utilities.Charset.UTF_8);
  return rawHash.map(byte => ('0' + (byte & 0xFF).toString(16)).slice(-2)).join('');
}

function initSheets(ss) {
  if (!ss) ss = SpreadsheetApp.getActiveSpreadsheet();
  const structure = {
    'Usuarios': ['id', 'email', 'password_hash', 'nombre', 'rol', 'creado_en', 'username'],
    'Configuracion': ['clave', 'valor', 'actualizado_en'],
    'ModelosCelular': ['id', 'nombre', 'precio_referencia_usd', 'precio_referencia_cop', 'actualizado_en'],
    'ReferenciaRetoma': ['id', 'modelo_id', 'capacidad', 'grado', 'precio_compra_usd', 'precio_compra_cop', 'precio_venta_usd', 'precio_venta_cop', 'actualizado_en'],
    'UnidadesCelular': ['id', 'modelo_id', 'imei', 'capacidad', 'grado', 'costo_real', 'precio_venta_real', 'fecha_ingreso', 'estado', 'color'],
    'Accesorios': ['id', 'nombre', 'costo', 'precio_recomendado', 'stock', 'actualizado_en'],
    'Ventas': ['id', 'numero_correlativo', 'fecha_hora', 'vendedor_id', 'tipo_producto', 'unidad_id', 'producto_id', 'imei', 'nombre_cliente', 'nota', 'metodo_pago', 'cantidad', 'valor_venta', 'costo_total', 'ganancia', 'creado_en', 'regalo_accesorio_id', 'editado', 'editado_en'],
    'HistorialIMEI': ['fecha_hora', 'imei', 'accion', 'usuario_id', 'detalle'],
    'Perdidas': ['id', 'fecha_hora', 'tipo', 'accesorio_id', 'concepto', 'costo', 'usuario_id', 'detalle']
  };

  for (let sheetName in structure) {
    let sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      sheet.appendRow(structure[sheetName]);
      if (sheetName === 'Configuracion') {
        sheet.appendRow(['tasa_cambio_usd_cop', '4000', new Date().toISOString()]);
      }
      if (sheetName === 'Usuarios') {
        const defaultHash = hashPassword('admin123');
        sheet.appendRow(['user_admin_default', 'admin@local.com', defaultHash, 'Administrador', 'admin', new Date().toISOString(), '']);
      }
    }
  }

  // Hojas existentes: agrega al final las columnas nuevas que falten,
  // para que el esquema quede siempre sincronizado sin tocar las posiciones actuales.
  for (let sheetName in structure) {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet || sheet.getLastRow() < 1) continue;
    const existentes = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]
      .map(h => String(h).trim().toLowerCase());
    let colIndex = sheet.getLastColumn() + 1;
    structure[sheetName].forEach(h => {
      const ident = String(h).trim().toLowerCase();
      if (!existentes.includes(ident)) {
        sheet.getRange(1, colIndex).setValue(h);
        colIndex++;
      }
    });
  }

  actualizarAdminRyu(ss);
}

// Actualiza las credenciales del admin RYU (ryutechnology.oficial@gmail.com):
// username de login "RYU" y contraseña "1234" (guardada como hash SHA-256).
// Es idempotente: solo escribe si hace falta, y no toca los demás usuarios.
function actualizarAdminRyu(ss) {
  const usuarios = getSheetData(ss, 'Usuarios');
  const admin = usuarios.find(u => String(u.email).trim().toLowerCase() === 'ryutechnology.oficial@gmail.com');
  if (!admin) return;

  const nuevoUsername = 'RYU';
  const nuevoHash = hashPassword('1234');
  if (String(admin.username || '').trim() === nuevoUsername && admin.password_hash === nuevoHash) return;

  const sheet = ss.getSheetByName('Usuarios');
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]
    .map(h => String(h).trim().toLowerCase());
  const col = (name) => headers.indexOf(name) + 1;
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(admin.id)) {
      if (col('username')) sheet.getRange(i + 1, col('username')).setValue(nuevoUsername);
      if (col('password_hash')) sheet.getRange(i + 1, col('password_hash')).setValue(nuevoHash);
      break;
    }
  }
}

function getSheetData(ss, sheetName) {
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return [];
  const rows = sheet.getDataRange().getValues();
  if (rows.length <= 1) return [];
  const headers = rows[0];
  const data = [];
  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    const obj = {};
    for (let j = 0; j < headers.length; j++) {
      obj[headers[j]] = row[j];
    }
    data.push(obj);
  }
  return data;
}

function getAllData(ss) {
  const usuarios = getSheetData(ss, 'Usuarios').map(u => {
    delete u.password_hash;
    return u;
  });
  const ventasOrdenadas = getSheetData(ss, 'Ventas')
    .sort((a, b) => String(b.fecha_hora || '').localeCompare(String(a.fecha_hora || '')));
  return {
    ss_url: String(ss.getUrl() || ''),
    usuarios: usuarios,
    configuracion: getSheetData(ss, 'Configuracion'),
    modelos_celular: getSheetData(ss, 'ModelosCelular'),
    referencia_retoma: getSheetData(ss, 'ReferenciaRetoma'),
    unidades_celular: getSheetData(ss, 'UnidadesCelular'),
    accesorios: getSheetData(ss, 'Accesorios'),
    ventas: ventasOrdenadas.slice(0, 200)
  };
}

// Historial completo de ventas (todas, ordenadas por fecha desc) para cuando
// el usuario quiera ver más atrás de las 200 recientes del dashboard.
function getHistorialCompleto(ss) {
  const ventas = getSheetData(ss, 'Ventas')
    .sort((a, b) => String(b.fecha_hora || '').localeCompare(String(a.fecha_hora || '')));
  return {
    ventas: ventas,
    accesorios: getSheetData(ss, 'Accesorios')
  };
}

// =====================================================
// HISTORIAL DE MOVIMIENTOS POR IMEI
// =====================================================
function registrarHistorialIMEI(ss, imei, accion, usuarioId, detalle) {
  const hoja = ss.getSheetByName('HistorialIMEI');
  if (!hoja || !imei) return;
  hoja.appendRow([new Date().toISOString(), String(imei).trim(), accion, usuarioId || '', detalle || '']);
}

function getHistorialIMEI(ss, imei) {
  if (!imei) throw new Error('Falta el IMEI para consultar el historial.');
  const movimientos = getSheetData(ss, 'HistorialIMEI')
    .filter(m => String(m.imei).trim() === String(imei).trim())
    .sort((a, b) => String(b.fecha_hora || '').localeCompare(String(a.fecha_hora || '')));
  return { imei: String(imei).trim(), movimientos: movimientos };
}

// =====================================================
// REPORTE DE PÉRDIDAS POR REGALOS (mes actual)
// =====================================================
function getRegalosDelMes(ss) {
  const ventas = getSheetData(ss, 'Ventas');
  const accesorios = getSheetData(ss, 'Accesorios');
  const perdidas = getSheetData(ss, 'Perdidas');
  const mesActual = new Date().toISOString().substring(0, 7);

  const conRegalo = ventas.filter(v =>
    v.regalo_accesorio_id && String(v.regalo_accesorio_id).trim() !== '' &&
    v.fecha_hora && String(v.fecha_hora).startsWith(mesActual)
  );

  const detalle = conRegalo.map(v => {
    const acc = accesorios.find(a => String(a.id) === String(v.regalo_accesorio_id));
    return {
      fecha_hora: v.fecha_hora,
      origen: 'venta',
      imei: v.imei || '',
      numero_correlativo: v.numero_correlativo || '',
      nombre_cliente: v.nombre_cliente || '',
      regalo_accesorio_id: v.regalo_accesorio_id,
      regalo_nombre: acc ? acc.nombre : 'Desconocido',
      regalo_costo: acc ? Number(acc.costo || 0) : 0
    };
  });

  const perdidasMes = perdidas.filter(p =>
    p.fecha_hora && String(p.fecha_hora).startsWith(mesActual)
  );

  perdidasMes.forEach(p => {
    detalle.push({
      fecha_hora: p.fecha_hora,
      origen: 'manual',
      imei: '',
      numero_correlativo: '',
      nombre_cliente: p.concepto || 'Pérdida manual',
      regalo_nombre: p.concepto || 'Pérdida manual',
      regalo_costo: Number(p.costo || 0)
    });
  });

  detalle.sort((a, b) => String(b.fecha_hora || '').localeCompare(String(a.fecha_hora || '')));

  const total = detalle.reduce((s, d) => s + (Number(d.regalo_costo) || 0), 0);
  return { total: total, detalle: detalle };
}

function registrarPerdida(ss, params) {
  const perdidas = getSheetData(ss, 'Perdidas');
  const sheet = ss.getSheetByName('Perdidas');
  const fecha = new Date().toISOString();
  let acc;
  let costo = Number(params.costo || 0);
  let concepto = String(params.concepto || '').trim();

  if (params.accesorio_id) {
    const accesorios = getSheetData(ss, 'Accesorios');
    acc = accesorios.find(a => String(a.id) === String(params.accesorio_id));
    if (!acc) throw new Error('Accesorio no encontrado.');
    if (Number(acc.stock || 0) < 1) throw new Error('El accesorio ya no tiene stock.');
    concepto = acc.nombre || concepto;
    costo = Number(acc.costo || 0);
    const rows = ss.getSheetByName('Accesorios').getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) === String(params.accesorio_id)) {
        ss.getSheetByName('Accesorios').getRange(i + 1, 5).setValue(Number(acc.stock || 0) - 1);
        break;
      }
    }
  }

  if (!(costo > 0)) throw new Error('La pérdida debe tener un costo válido mayor a 0.');
  if (!concepto) throw new Error('Indica un concepto para la pérdida.');

  const nuevoId = perdidas.length > 0 ? Math.max(...perdidas.map(p => Number(p.id) || 0)) + 1 : 1;
  sheet.appendRow([nuevoId, fecha, 'manual', params.accesorio_id || '', concepto, costo, params.userId || '', String(params.detalle || '').trim()]);
  return { success: true, id: nuevoId };
}

function validateUserAuth(ss, userId, action) {
  if (!userId) throw new Error('Acceso no autorizado: falta identificador de usuario.');

  const usuarios = getSheetData(ss, 'Usuarios');
  const user = usuarios.find(u => String(u.id) === String(userId));
  if (!user) throw new Error('Usuario no encontrado o sesión inválida.');

  const adminActions = ['anularVenta', 'editarVenta', 'actualizarTasa', 'agregarModelo', 'agregarReferenciaRetoma', 'crearEmpleado', 'agregarCelular', 'guardarAccesorio'];

  if (adminActions.includes(action) && user.rol !== 'admin') {
    throw new Error('Acceso denegado: el rol de vendedor no tiene permisos para ejecutar esta acción.');
  }

  return user;
}

function loginUser(ss, email, password) {
  const usuarios = getSheetData(ss, 'Usuarios');
  const hashedInput = hashPassword(password);
  const identificador = String(email || '').trim().toLowerCase();

  const user = usuarios.find(u => {
    if (u.password_hash !== hashedInput) return false;
    const correo = String(u.email || '').trim().toLowerCase();
    const usuario = String(u.username || '').trim().toLowerCase();
    return correo === identificador || usuario === identificador;
  });

  if (!user) {
    throw new Error('Correo, usuario o contraseña incorrectos');
  }
  return { id: user.id, email: user.email, nombre: user.nombre, rol: user.rol };
}

function registrarVenta(ss, params) {
  const sheet = ss.getSheetByName('Ventas');
  const unidadesSheet = ss.getSheetByName('UnidadesCelular');
  const accSheet = ss.getSheetByName('Accesorios');

  const ventas = getSheetData(ss, 'Ventas');
  const nuevoId = ventas.length > 0 ? Math.max(...ventas.map(v => Number(v.id) || 0)) + 1 : 1;
  const correlativo = nuevoId;
  const fechaHora = new Date().toISOString();

  if (params.tipo_producto === 'celular') {
    const rows = unidadesSheet.getDataRange().getValues();
    let rowIndex = -1;
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) === String(params.unidad_id)) {
        rowIndex = i + 1;
        if (rows[i][8] === 'vendido') {
          throw new Error('Este celular ya fue vendido por otro usuario en este mismo instante.');
        }
        break;
      }
    }
    if (rowIndex === -1) throw new Error('Celular no encontrado en inventario');
    unidadesSheet.getRange(rowIndex, 9).setValue('vendido');
  } else if (params.tipo_producto === 'accesorio') {
    const rows = accSheet.getDataRange().getValues();
    let rowIndex = -1;
    let currentStock = 0;
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) === String(params.producto_id)) {
        rowIndex = i + 1;
        currentStock = Number(rows[i][4]);
        break;
      }
    }
    if (rowIndex === -1) throw new Error('Accesorio no encontrado');
    const cantidad = Number(params.cantidad) || 1;
    if (currentStock < cantidad) {
      throw new Error('Stock insuficiente. Quedan ' + currentStock + ' unidades.');
    }
    accSheet.getRange(rowIndex, 5).setValue(currentStock - cantidad);
    accSheet.getRange(rowIndex, 6).setValue(fechaHora);
  }

  let costoTotal = Number(params.costo_total) || 0;
  const regaloId = params.accesorio_regalo_id || params.regalo_accesorio_id || '';

  // REGALO DE ACCESORIO en venta de celular:
  // descuenta 1 unidad del stock, NO suma al valor_venta, y añade su costo al costo_total.
  if (params.tipo_producto === 'celular' && regaloId) {
    const accRows = accSheet.getDataRange().getValues();
    let accRowIndex = -1;
    let accStock = 0;
    let accCosto = 0;
    let accNombre = '';
    for (let i = 1; i < accRows.length; i++) {
      if (String(accRows[i][0]) === String(regaloId)) {
        accRowIndex = i + 1;
        accStock = Number(accRows[i][4]);
        accCosto = Number(accRows[i][2]);
        accNombre = String(accRows[i][1]);
        break;
      }
    }
    if (accRowIndex === -1) throw new Error('El accesorio de regalo no fue encontrado.');
    if (accStock < 1) throw new Error('No hay stock disponible del accesorio seleccionado para regalar.');
    accSheet.getRange(accRowIndex, 5).setValue(accStock - 1);
    accSheet.getRange(accRowIndex, 6).setValue(fechaHora);
    costoTotal = costoTotal + accCosto;
    const detalleRegalo = 'Regalo: ' + accNombre;
    params.nota = params.nota ? params.nota + ' | ' + detalleRegalo : detalleRegalo;
  }

  // La ganancia se recalcula SIEMPRE en el servidor (fuente de verdad).
  const valorVenta = Number(params.valor_venta) || 0;
  const ganancia = Number((valorVenta - costoTotal).toFixed(2));

  sheet.appendRow([
    nuevoId,
    correlativo,
    fechaHora,
    params.vendedor_id,
    params.tipo_producto,
    params.unidad_id || '',
    params.producto_id || '',
    params.imei || '',
    params.nombre_cliente,
    params.nota || '',
    params.metodo_pago,
    params.cantidad || 1,
    valorVenta,
    costoTotal,
    ganancia,
    fechaHora,
    regaloId || '',
    '',
    ''
  ]);

  if (params.tipo_producto === 'celular') {
    registrarHistorialIMEI(ss, params.imei, 'venta', params.userId || params.vendedor_id,
      'Venta N° ' + nuevoId + (params.nombre_cliente ? ' | Cliente: ' + params.nombre_cliente : '') + (params.nota ? ' | ' + params.nota : ''));
  }

  return { success: true, ventaId: nuevoId };
}

function editarVenta(ss, params) {
  const sheet = ss.getSheetByName('Ventas');
  const rows = sheet.getDataRange().getValues();
  const headers = rows[0];
  let rowIndex = -1;
  let venta = null;

  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(params.ventaId)) {
      rowIndex = i + 1;
      venta = {};
      for (let j = 0; j < headers.length; j++) {
        venta[headers[j]] = rows[i][j];
      }
      break;
    }
  }

  if (!venta) throw new Error('Venta no encontrada');

  const colIdx = (name) => headers.indexOf(name) + 1;

  const nuevoValor = params.valor_venta !== undefined && params.valor_venta !== null && params.valor_venta !== ''
    ? Number(params.valor_venta)
    : Number(venta.valor_venta || 0);

  const nuevoCosto = params.costo_total !== undefined && params.costo_total !== null && params.costo_total !== ''
    ? Number(params.costo_total)
    : Number(venta.costo_total || 0);

  if (colIdx('valor_venta')) sheet.getRange(rowIndex, colIdx('valor_venta')).setValue(nuevoValor);
  if (colIdx('costo_total')) sheet.getRange(rowIndex, colIdx('costo_total')).setValue(nuevoCosto);

  // Recalcula la ganancia y deja constancia de la corrección.
  if (colIdx('ganancia')) sheet.getRange(rowIndex, colIdx('ganancia')).setValue(Number((nuevoValor - nuevoCosto).toFixed(2)));
  if (colIdx('editado')) sheet.getRange(rowIndex, colIdx('editado')).setValue(true);
  if (colIdx('editado_en')) sheet.getRange(rowIndex, colIdx('editado_en')).setValue(new Date().toISOString());

  const hdr = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const updated = sheet.getRange(rowIndex, 1, 1, sheet.getLastColumn()).getValues()[0];
  const result = {};
  for (let j = 0; j < hdr.length; j++) {
    result[hdr[j]] = updated[j];
  }

  if (venta.imei) {
    registrarHistorialIMEI(ss, venta.imei, 'edicion', params.userId || params.vendedor_id,
      'Corrección venta N° ' + (venta.numero_correlativo || venta.id) + ' → valor de venta ' + nuevoValor + ' / costo ' + nuevoCosto);
  }

  return result;
}

function anularVenta(ss, ventaId, userId) {
  const ventasSheet = ss.getSheetByName('Ventas');
  const rows = ventasSheet.getDataRange().getValues();
  let rowIndex = -1;
  let venta = null;
  const headers = rows[0];

  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(ventaId)) {
      rowIndex = i + 1;
      venta = {};
      for (let j = 0; j < headers.length; j++) {
        venta[headers[j]] = rows[i][j];
      }
      break;
    }
  }

  if (!venta) throw new Error('Venta no encontrada');

  if (venta.tipo_producto === 'celular' && venta.unidad_id) {
    const unidadesSheet = ss.getSheetByName('UnidadesCelular');
    const uRows = unidadesSheet.getDataRange().getValues();
    for (let i = 1; i < uRows.length; i++) {
      if (String(uRows[i][0]) === String(venta.unidad_id)) {
        unidadesSheet.getRange(i + 1, 9).setValue('disponible');
        break;
      }
    }
  } else if (venta.tipo_producto === 'accesorio' && venta.producto_id) {
    const accSheet = ss.getSheetByName('Accesorios');
    const aRows = accSheet.getDataRange().getValues();
    for (let i = 1; i < aRows.length; i++) {
      if (String(aRows[i][0]) === String(venta.producto_id)) {
        const currentStock = Number(aRows[i][4]);
        accSheet.getRange(i + 1, 5).setValue(currentStock + Number(venta.cantidad));
        break;
      }
    }
  }

  if (venta.regalo_accesorio_id) {
    const accSheet = ss.getSheetByName('Accesorios');
    const aRows = accSheet.getDataRange().getValues();
    for (let i = 1; i < aRows.length; i++) {
      if (String(aRows[i][0]) === String(venta.regalo_accesorio_id)) {
        accSheet.getRange(i + 1, 5).setValue(Number(aRows[i][4]) + 1);
        break;
      }
    }
  }

  if (venta.imei) {
    registrarHistorialIMEI(ss, venta.imei, 'anulacion', userId || venta.vendedor_id,
      'Venta N° ' + ventaId + ' anulada' + (venta.nombre_cliente ? ' | Cliente: ' + venta.nombre_cliente : ''));
  }

  ventasSheet.deleteRow(rowIndex);
  return { success: true };
}

function agregarCelular(ss, params) {
  const sheet = ss.getSheetByName('UnidadesCelular');
  const celulares = getSheetData(ss, 'UnidadesCelular');

  if (celulares.some(c => String(c.imei).trim() === String(params.imei).trim())) {
    throw new Error('Ya existe un celular registrado con este IMEI.');
  }

  const nuevoId = celulares.length > 0 ? Math.max(...celulares.map(c => Number(c.id) || 0)) + 1 : 1;
  const fechaIngreso = new Date().toISOString();

  sheet.appendRow([
    nuevoId,
    params.modelo_id,
    params.imei.trim(),
    params.capacidad,
    params.grado,
    params.costo_real,
    params.precio_venta_real,
    fechaIngreso,
    'disponible',
    params.color || ''
  ]);

  registrarHistorialIMEI(ss, params.imei, 'ingreso', params.userId,
    'Ingreso: ' + [params.modelo_id, params.capacidad, params.grado, params.color].filter(Boolean).join(' / '));

  return { success: true };
}

function guardarAccesorio(ss, params) {
  const sheet = ss.getSheetByName('Accesorios');
  const accesorios = getSheetData(ss, 'Accesorios');
  const fechaActual = new Date().toISOString();

  if (params.id) {
    const rows = sheet.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) === String(params.id)) {
        sheet.getRange(i + 1, 2).setValue(params.nombre);
        sheet.getRange(i + 1, 3).setValue(params.costo);
        sheet.getRange(i + 1, 4).setValue(params.precio_recomendado);
        sheet.getRange(i + 1, 5).setValue(params.stock);
        sheet.getRange(i + 1, 6).setValue(fechaActual);
        return { success: true };
      }
    }
    throw new Error('Accesorio no encontrado');
  } else {
    const nuevoId = accesorios.length > 0 ? Math.max(...accesorios.map(a => Number(a.id) || 0)) + 1 : 1;
    sheet.appendRow([nuevoId, params.nombre, params.costo, params.precio_recomendado, params.stock, fechaActual]);
    return { success: true };
  }
}

function actualizarTasa(ss, tasa) {
  const sheet = ss.getSheetByName('Configuracion');
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === 'tasa_cambio_usd_cop') {
      sheet.getRange(i + 1, 2).setValue(tasa);
      sheet.getRange(i + 1, 3).setValue(new Date().toISOString());
      return { success: true };
    }
  }
  sheet.appendRow(['tasa_cambio_usd_cop', tasa, new Date().toISOString()]);
  return { success: true };
}

function agregarModelo(ss, params) {
  const sheet = ss.getSheetByName('ModelosCelular');
  const modelos = getSheetData(ss, 'ModelosCelular');
  const nuevoId = modelos.length > 0 ? Math.max(...modelos.map(m => Number(m.id) || 0)) + 1 : 1;
  const fecha = new Date().toISOString();

  sheet.appendRow([nuevoId, params.nombre, params.precio_referencia_usd, params.precio_referencia_cop, fecha]);
  return { success: true };
}

function agregarReferenciaRetoma(ss, params) {
  const sheet = ss.getSheetByName('ReferenciaRetoma');
  const refs = getSheetData(ss, 'ReferenciaRetoma');
  const nuevoId = refs.length > 0 ? Math.max(...refs.map(r => Number(r.id) || 0)) + 1 : 1;
  const fecha = new Date().toISOString();

  sheet.appendRow([
    nuevoId,
    params.modelo_id,
    params.capacidad,
    params.grado,
    params.precio_compra_usd,
    params.precio_compra_cop,
    params.precio_venta_usd,
    params.precio_venta_cop,
    fecha
  ]);

  return { success: true };
}

// CORRECCIÓN: la variable "usuarios" no existía (bug del código original).
// Ahora se obtiene correctamente con getSheetData antes de usarla.
function crearEmpleado(ss, params) {
  const sheet = ss.getSheetByName('Usuarios');
  const usuarios = getSheetData(ss, 'Usuarios');

  if (usuarios.some(u => String(u.email).toLowerCase() === String(params.email).toLowerCase())) {
    throw new Error('Ya existe un usuario con este correo electrónico.');
  }

  const nuevoId = 'user_' + new Date().getTime();
  const fecha = new Date().toISOString();
  const hashedPass = hashPassword(params.password);

  sheet.appendRow([nuevoId, params.email.trim(), hashedPass, params.nombre, params.rol || 'vendedor', fecha]);
  return { success: true, id: nuevoId };
}

// =========================================================================
// MENÚ PERSONALIZADO EN LA HOJA — para inicializar y probar sin depender
// del desplegable de funciones del editor de Apps Script.
// =========================================================================
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('CeluControl')
    .addItem('1. Inicializar / Reparar Hojas', 'menuInitSheets')
    .addItem('2. Aplicar Credenciales Admin RYU (usuario: RYU / pass: 1234)', 'menuActualizarAdminRyu')
    .addItem('3. Probar Login (admin@local.com / admin123)', 'menuProbarLogin')
    .addItem('4. Ver Inventario', 'menuVerInventario')
    .addToUi();
}

function menuActualizarAdminRyu() {
  try {
    actualizarAdminRyu(SpreadsheetApp.getActiveSpreadsheet());
    SpreadsheetApp.getUi().alert('✅ Credenciales del admin RYU actualizadas.\nLogin: RYU — Contraseña: 1234');
  } catch (err) {
    SpreadsheetApp.getUi().alert('❌ Error: ' + err.message);
  }
}

function menuInitSheets() {
  try {
    initSheets(SpreadsheetApp.getActiveSpreadsheet());
    SpreadsheetApp.getUi().alert('✅ Hojas inicializadas/reparadas correctamente.');
  } catch (err) {
    SpreadsheetApp.getUi().alert('❌ Error: ' + err.message);
  }
}

function menuProbarLogin() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const r = loginUser(ss, 'admin@local.com', 'admin123');
    SpreadsheetApp.getUi().alert('✅ Login exitoso.\nNombre: ' + r.nombre + '\nRol: ' + r.rol);
  } catch (err) {
    SpreadsheetApp.getUi().alert('❌ Error: ' + err.message + '\n\nSi dice que no existe, corre primero "1. Inicializar / Reparar Hojas".');
  }
}

function menuVerInventario() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const r = getSheetData(ss, 'UnidadesCelular');
    const resumen = r.length === 0
      ? 'El inventario está vacío.'
      : r.length + ' celular(es) registrados. Ejemplo: ' + JSON.stringify(r[0]);
    SpreadsheetApp.getUi().alert(resumen);
  } catch (err) {
    SpreadsheetApp.getUi().alert('❌ Error: ' + err.message);
  }
}
