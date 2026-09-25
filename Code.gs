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

      case 'editarReferenciaRetoma':
        result.data = editarReferenciaRetoma(ss, params);
        break;

      case 'sembrarCatalogo':
        result.data = sembrarCatalogoRetoma(ss);
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

  const adminActions = ['anularVenta', 'editarVenta', 'editarReferenciaRetoma', 'actualizarTasa', 'agregarModelo', 'agregarReferenciaRetoma', 'crearEmpleado', 'agregarCelular', 'guardarAccesorio', 'sembrarCatalogo'];

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

// Actualiza una referencia de retoma existente (busca la fila por su id).
function editarReferenciaRetoma(ss, params) {
  const sheet = ss.getSheetByName('ReferenciaRetoma');
  const refs = getSheetData(ss, 'ReferenciaRetoma');
  const existe = refs.some(r => String(r.id) === String(params.id));
  if (!existe) throw new Error('Referencia de retoma no encontrada.');

  const rows = sheet.getDataRange().getValues();
  const headers = rows[0].map(h => String(h).trim().toLowerCase());
  const col = (name) => headers.indexOf(name) + 1;
  const fecha = new Date().toISOString();

  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(params.id)) {
      if (params.modelo_id !== undefined && params.modelo_id !== null && String(params.modelo_id) !== '') {
        if (col('modelo_id')) sheet.getRange(i + 1, col('modelo_id')).setValue(params.modelo_id);
      }
      if (col('capacidad')) sheet.getRange(i + 1, col('capacidad')).setValue(params.capacidad);
      if (col('grado')) sheet.getRange(i + 1, col('grado')).setValue(params.grado);
      if (col('precio_compra_usd')) sheet.getRange(i + 1, col('precio_compra_usd')).setValue(params.precio_compra_usd);
      if (col('precio_compra_cop')) sheet.getRange(i + 1, col('precio_compra_cop')).setValue(params.precio_compra_cop);
      if (col('precio_venta_usd')) sheet.getRange(i + 1, col('precio_venta_usd')).setValue(params.precio_venta_usd);
      if (col('precio_venta_cop')) sheet.getRange(i + 1, col('precio_venta_cop')).setValue(params.precio_venta_cop);
      if (col('actualizado_en')) sheet.getRange(i + 1, col('actualizado_en')).setValue(fecha);
      return { success: true };
    }
  }
  throw new Error('Referencia de retoma no encontrada.');
}

// Carga el catálogo inicial de Referencias de Retoma (IPHONE / IPAD / WATCH /
// MACBOOK) EXACTAMENTE como fue proporcionado, sin corregir ni redondear nada.
// Es idempotente:
//   - Crea en ModelosCelular los modelos que no existan (solo agrega, no modifica).
//   - Solo agrega una referencia si NO existe ya una fila con el mismo
//     (modelo_id, capacidad, grado y precio_compra_cop). Así respeta registros
//     con el mismo Modelo+Capacidad pero precios distintos, y nunca duplica.
// Grado B = precio de compra grado B; si la línea del catálogo no trae grado B
// (valor null), solo se crea la fila del grado A.
// Convierte un valor de precio tal como se escribió en el catálogo, p. ej.
// "$400.000", "$360.00", "- $5.000", "$1.080.000". Devuelve número o null.
// Se guarda EXACTAMENTE lo escrito (sin recalcular, sin redondear, sin escalar):
// "$360.00" → 360, "$2.400.000" → 2400000, "-$4.50" → -4.5. El frontend decide
// cómo presentarlo visualmente (formato de 2 decimales o de miles) sin tocar
// el dato almacenado.
function parseCOP(s) {
  if (s === null || s === undefined) return null;
  var txt = String(s).replace(/[$,]/g, '').replace(/\s/g, '');
  if (txt === '') return null;
  var parts = txt.split('.');
  if (parts.length > 1 && parts[parts.length - 1].length === 2) {
    var entera = parts.slice(0, -1).join('').replace(/\./g, '');
    return parseFloat(entera + '.' + parts[parts.length - 1]);
  }
  return parseFloat(txt.replace(/\./g, ''));
}

// Carga en la hoja ReferenciaRetoma EXACTAMENTE las referencias entregadas por
// RYU (volcado real de la hoja): 803 filas en el MISMO orden, con los mismos
// valores, incluidos duplicados, negativos, capacidades '0' y valores raros.
//  - Cada entrada es UNA fila: [modelo, capacidad, grado, valor_cop, fecha].
//  - Reconstruye la hoja (no duplica): se ejecuta SOLO de forma manual.
//  - No toca ModelosCelular, UnidadesCelular, Ventas ni cualquier otra hoja.
function sembrarCatalogoRetoma(ss) {
    const FILAS = [
    ['IPHONE 16 PRO MAX', '256', 'A', '2.550.000', '2026-09-25'],
    ['IPHONE 16 PRO MAX', '256', 'B', '2.370.000', '2026-09-25'],
    ['IPHONE 16 PRO', '256', 'A', '2.150.000', '2026-09-25'],
    ['IPHONE 16 PRO', '256', 'B', '1.920.000', '2026-09-25'],
    ['IPHONE 16 PRO', '256', 'A', '2.250.000', '2026-09-25'],
    ['IPHONE 16 PRO', '256', 'B', '2.010.000', '2026-09-25'],
    ['IPHONE 16', '256', 'A', '1.550.000', '2026-09-25'],
    ['IPHONE 16', '256', 'B', '1.380.000', '2026-09-25'],
    ['IPHONE 16', '128', 'A', '1.450.000', '2026-09-25'],
    ['IPHONE 16', '128', 'B', '1.290.000', '2026-09-25'],
    ['IPHONE 15 PRO MAX', '128', 'A', '110', '2026-09-25'],
    ['IPHONE 15 PRO MAX', '128', 'B', '84', '2026-09-25'],
    ['IPHONE 15 PRO MAX', '512', 'A', '2.350.000', '2026-09-25'],
    ['IPHONE 15 PRO MAX', '512', 'B', '2.100.000', '2026-09-25'],
    ['IPHONE 15 PRO MAX', '256', 'A', '2.250.000', '2026-09-25'],
    ['IPHONE 15 PRO MAX', '256', 'B', '2.010.000', '2026-09-25'],
    ['IPHONE 15 PRO MAX', '128', 'A', '2.100.000', '2026-09-25'],
    ['IPHONE 15 PRO MAX', '128', 'B', '1.920.000', '2026-09-25'],
    ['IPHONE 15 PRO', '128', 'A', '2.050.000', '2026-09-25'],
    ['IPHONE 15 PRO', '128', 'B', '1.830.000', '2026-09-25'],
    ['IPHONE 15 PRO', '512', 'A', '1.950.000', '2026-09-25'],
    ['IPHONE 15 PRO', '512', 'B', '1.740.000', '2026-09-25'],
    ['IPHONE 15 PRO', '256', 'A', '1.850.000', '2026-09-25'],
    ['IPHONE 15 PRO', '256', 'B', '1.650.000', '2026-09-25'],
    ['IPHONE 15 PRO', '128', 'A', '1.750.000', '2026-09-25'],
    ['IPHONE 15 PRO', '128', 'B', '1.560.000', '2026-09-25'],
    ['IPHONE 15 Plus', '256', 'A', '1.650.000', '2026-09-25'],
    ['IPHONE 15 Plus', '256', 'B', '1.470.000', '2026-09-25'],
    ['IPHONE 15 Plus', '128', 'A', '1.550.000', '2026-09-25'],
    ['IPHONE 15 Plus', '128', 'B', '1.380.000', '2026-09-25'],
    ['IPHONE 15', '512', 'A', '1.550.000', '2026-09-25'],
    ['IPHONE 15', '512', 'B', '1.380.000', '2026-09-25'],
    ['IPHONE 15', '256', 'A', '1.450.000', '2026-09-25'],
    ['IPHONE 15', '256', 'B', '1.290.000', '2026-09-25'],
    ['IPHONE 15', '128', 'A', '1.350.000', '2026-09-25'],
    ['IPHONE 15', '128', 'B', '1.200.000', '2026-09-25'],
    ['IPHONE 14 PRO MAX', '128', 'A', '1.550.000', '2026-09-25'],
    ['IPHONE 14 PRO MAX', '128', 'B', '1.450.000', '2026-09-25'],
    ['IPHONE 14 PRO MAX', '512', 'A', '1.850.000', '2026-09-25'],
    ['IPHONE 14 PRO MAX', '512', 'B', '1.650.000', '2026-09-25'],
    ['IPHONE 14 PRO MAX', '256', 'A', '1.750.000', '2026-09-25'],
    ['IPHONE 14 PRO MAX', '256', 'B', '1.560.000', '2026-09-25'],
    ['IPHONE 14 PRO MAX', '128', 'A', '1.650.000', '2026-09-25'],
    ['IPHONE 14 PRO MAX', '128', 'B', '1.470.000', '2026-09-25'],
    ['IPHONE 14 PRO', '128', 'A', '1.650.000', '2026-09-25'],
    ['IPHONE 14 PRO', '128', 'B', '1.470.000', '2026-09-25'],
    ['IPHONE 14 PRO', '512', 'A', '1.550.000', '2026-09-25'],
    ['IPHONE 14 PRO', '512', 'B', '1.380.000', '2026-09-25'],
    ['IPHONE 14 PRO', '256', 'A', '1.450.000', '2026-09-25'],
    ['IPHONE 14 PRO', '256', 'B', '1.290.000', '2026-09-25'],
    ['IPHONE 14 PRO', '128', 'A', '1.350.000', '2026-09-25'],
    ['IPHONE 14 PRO', '128', 'B', '1.200.000', '2026-09-25'],
    ['IPHONE 14 PLUS', '512', 'A', '1.250.000', '2026-09-25'],
    ['IPHONE 14 PLUS', '512', 'B', '1.110.000', '2026-09-25'],
    ['IPHONE 14 PLUS', '256', 'A', '1.150.000', '2026-09-25'],
    ['IPHONE 14 PLUS', '256', 'B', '1.020.000', '2026-09-25'],
    ['IPHONE 14 PLUS', '128', 'A', '1.050.000', '2026-09-25'],
    ['IPHONE 14 PLUS', '128', 'B', '930', '2026-09-25'],
    ['IPHONE 14', '128', 'A', '1.150.000', '2026-09-25'],
    ['IPHONE 14', '128', 'B', '1.020.000', '2026-09-25'],
    ['IPHONE 14', '512', 'A', '1.050.000', '2026-09-25'],
    ['IPHONE 14', '512', 'B', '930', '2026-09-25'],
    ['IPHONE 14', '256', 'A', '950', '2026-09-25'],
    ['IPHONE 14', '256', 'B', '840', '2026-09-25'],
    ['IPHONE 14', '128', 'A', '900', '2026-09-25'],
    ['IPHONE 14', '128', 'B', '795', '2026-09-25'],
    ['IPHONE 13 PRO MAX', '128', 'A', '1.350.000', '2026-09-25'],
    ['IPHONE 13 PRO MAX', '128', 'B', '1.200.000', '2026-09-25'],
    ['IPHONE 13 PRO MAX', '512', 'A', '1.250.000', '2026-09-25'],
    ['IPHONE 13 PRO MAX', '512', 'B', '1.110.000', '2026-09-25'],
    ['IPHONE 13 PRO MAX', '256', 'A', '1.150.000', '2026-09-25'],
    ['IPHONE 13 PRO MAX', '256', 'B', '1.020.000', '2026-09-25'],
    ['IPHONE 13 PRO MAX', '128', 'A', '1.050.000', '2026-09-25'],
    ['IPHONE 13 PRO MAX', '128', 'B', '930', '2026-09-25'],
    ['IPHONE 13 PRO', '128', 'A', '1.250.000', '2026-09-25'],
    ['IPHONE 13 PRO', '128', 'B', '1.110.000', '2026-09-25'],
    ['IPHONE 13 PRO', '512', 'A', '1.150.000', '2026-09-25'],
    ['IPHONE 13 PRO', '512', 'B', '1.020.000', '2026-09-25'],
    ['IPHONE 13 PRO', '256', 'A', '1.050.000', '2026-09-25'],
    ['IPHONE 13 PRO', '256', 'B', '930', '2026-09-25'],
    ['IPHONE 13 PRO', '128', 'A', '950', '2026-09-25'],
    ['IPHONE 13 PRO', '128', 'B', '840', '2026-09-25'],
    ['IPHONE 13', '512', 'A', '950', '2026-09-25'],
    ['IPHONE 13', '512', 'B', '840', '2026-09-25'],
    ['IPHONE 13', '256', 'A', '850', '2026-09-25'],
    ['IPHONE 13', '256', 'B', '750', '2026-09-25'],
    ['IPHONE 13', '128', 'A', '750', '2026-09-25'],
    ['IPHONE 13', '128', 'B', '660', '2026-09-25'],
    ['IPHONE 13 MINI', '512', 'A', '700', '2026-09-25'],
    ['IPHONE 13 MINI', '512', 'B', '615', '2026-09-25'],
    ['IPHONE 13 MINI', '256', 'A', '650', '2026-09-25'],
    ['IPHONE 13 MINI', '256', 'B', '570', '2026-09-25'],
    ['IPHONE 13 MINI', '128', 'A', '550', '2026-09-25'],
    ['IPHONE 13 MINI', '128', 'B', '480', '2026-09-25'],
    ['IPHONE 12 PRO MAX', '512', 'A', '1.050.000', '2026-09-25'],
    ['IPHONE 12 PRO MAX', '512', 'B', '930', '2026-09-25'],
    ['IPHONE 12 PRO MAX', '256', 'A', '950', '2026-09-25'],
    ['IPHONE 12 PRO MAX', '256', 'B', '840', '2026-09-25'],
    ['IPHONE 12 PRO MAX', '128', 'A', '850', '2026-09-25'],
    ['IPHONE 12 PRO MAX', '128', 'B', '750', '2026-09-25'],
    ['IPHONE 12 PRO', '512', 'A', '850', '2026-09-25'],
    ['IPHONE 12 PRO', '256', 'A', '800', '2026-09-25'],
    ['IPHONE 12 PRO', '256', 'B', '750', '2026-09-25'],
    ['IPHONE 12 PRO', '128', 'A', '750', '2026-09-25'],
    ['IPHONE 12 PRO', '128', 'B', '660', '2026-09-25'],
    ['IPHONE 12', '256', 'A', '500', '2026-09-25'],
    ['IPHONE 12', '256', 'B', '435', '2026-09-25'],
    ['IPHONE 12', '128', 'A', '450', '2026-09-25'],
    ['IPHONE 12', '128', 'B', '390', '2026-09-25'],
    ['IPHONE 12', '64', 'A', '350', '2026-09-25'],
    ['IPHONE 12', '64', 'B', '300', '2026-09-25'],
    ['IPHONE 12 MINI', '256', 'A', '350', '2026-09-25'],
    ['IPHONE 12 MINI', '256', 'B', '300', '2026-09-25'],
    ['IPHONE 12 MINI', '128', 'A', '300', '2026-09-25'],
    ['IPHONE 12 MINI', '128', 'B', '255', '2026-09-25'],
    ['IPHONE 12 MINI', '64', 'A', '250', '2026-09-25'],
    ['IPHONE 12 MINI', '64', 'B', '210', '2026-09-25'],
    ['IPHONE 11 PRO MAX', '512', 'A', '600', '2026-09-25'],
    ['IPHONE 11 PRO MAX', '512', 'B', '525', '2026-09-25'],
    ['IPHONE 11 PRO MAX', '256', 'A', '550', '2026-09-25'],
    ['IPHONE 11 PRO MAX', '256', 'B', '480', '2026-09-25'],
    ['IPHONE 11 PRO MAX', '64', 'A', '500', '2026-09-25'],
    ['IPHONE 11 PRO MAX', '64', 'B', '435', '2026-09-25'],
    ['IPHONE 11 PRO', '512', 'A', '450', '2026-09-25'],
    ['IPHONE 11 PRO', '512', 'B', '390', '2026-09-25'],
    ['IPHONE 11 PRO', '64', 'A', '350', '2026-09-25'],
    ['IPHONE 11 PRO', '64', 'B', '300', '2026-09-25'],
    ['IPHONE 11', '256', 'A', '450', '2026-09-25'],
    ['IPHONE 11', '256', 'B', '390', '2026-09-25'],
    ['IPHONE 11', '128', 'A', '350', '2026-09-25'],
    ['IPHONE 11', '128', 'B', '300', '2026-09-25'],
    ['IPHONE 11', '64', 'A', '250', '2026-09-25'],
    ['IPHONE 11', '64', 'B', '210', '2026-09-25'],
    ['IPAD 9.7 5TH GEN GSM', '32', 'A', '100', '2026-09-25'],
    ['IPAD 9.7 5TH GEN GSM', '32', 'B', '75', '2026-09-25'],
    ['IPAD 9.7 5TH GEN GSM', '128', 'A', '100', '2026-09-25'],
    ['IPAD 9.7 5TH GEN GSM', '128', 'B', '75', '2026-09-25'],
    ['IPAD 9.7 5TH GEN WIFI', '32', 'A', '100', '2026-09-25'],
    ['IPAD 9.7 5TH GEN WIFI', '32', 'B', '75', '2026-09-25'],
    ['IPAD 9.7 5TH GEN WIFI', '128', 'A', '100', '2026-09-25'],
    ['IPAD 9.7 5TH GEN WIFI', '128', 'B', '75', '2026-09-25'],
    ['IPAD 9.7 6TH GEN GSM', '32', 'A', '100', '2026-09-25'],
    ['IPAD 9.7 6TH GEN GSM', '32', 'B', '75', '2026-09-25'],
    ['IPAD 9.7 6TH GEN GSM', '128', 'A', '100', '2026-09-25'],
    ['IPAD 9.7 6TH GEN GSM', '128', 'B', '75', '2026-09-25'],
    ['IPAD 9.7 6TH GEN WIFI', '32', 'A', '100', '2026-09-25'],
    ['IPAD 9.7 6TH GEN WIFI', '32', 'B', '75', '2026-09-25'],
    ['IPAD 9.7 6TH GEN WIFI', '128', 'A', '100', '2026-09-25'],
    ['IPAD 9.7 6TH GEN WIFI', '128', 'B', '75', '2026-09-25'],
    ['IPAD AIR 2 GSM', '16', 'A', '-100', '2026-09-25'],
    ['IPAD AIR 2 GSM', '16', 'B', '-105', '2026-09-25'],
    ['IPAD AIR 2 GSM', '32', 'A', '-75', '2026-09-25'],
    ['IPAD AIR 2 GSM', '32', 'B', '-82.5', '2026-09-25'],
    ['IPAD AIR 2 GSM', '64', 'A', '-50', '2026-09-25'],
    ['IPAD AIR 2 GSM', '64', 'B', '-60', '2026-09-25'],
    ['IPAD AIR 2 GSM', '128', 'A', '-50', '2026-09-25'],
    ['IPAD AIR 2 GSM', '128', 'B', '-60', '2026-09-25'],
    ['IPAD AIR 2 WIFI', '16', 'A', '-25', '2026-09-25'],
    ['IPAD AIR 2 WIFI', '16', 'B', '-37.5', '2026-09-25'],
    ['IPAD AIR 2 WIFI', '64', 'A', '0', '2026-09-25'],
    ['IPAD AIR 2 WIFI', '64', 'B', '-15', '2026-09-25'],
    ['IPAD AIR 2 WIFI', '128', 'A', '25', '2026-09-25'],
    ['IPAD AIR 2 WIFI', '128', 'B', '7.5', '2026-09-25'],
    ['IPAD AIR 3RD GEN GSM', '64', 'A', '50', '2026-09-25'],
    ['IPAD AIR 3RD GEN GSM', '64', 'B', '30', '2026-09-25'],
    ['IPAD AIR 3RD GEN GSM', '256', 'A', '100', '2026-09-25'],
    ['IPAD AIR 3RD GEN GSM', '256', 'B', '75', '2026-09-25'],
    ['IPAD AIR 3RD GEN WIFI', '64', 'A', '150', '2026-09-25'],
    ['IPAD AIR 3RD GEN WIFI', '64', 'B', '120', '2026-09-25'],
    ['IPAD AIR 3RD GEN WIFI', '256', 'A', '200', '2026-09-25'],
    ['IPAD AIR 3RD GEN WIFI', '256', 'B', '165', '2026-09-25'],
    ['IPAD AIR GSM', '16', 'A', '-100', '2026-09-25'],
    ['IPAD AIR GSM', '16', 'B', '-105', '2026-09-25'],
    ['IPAD AIR GSM', '32', 'A', '-75', '2026-09-25'],
    ['IPAD AIR GSM', '32', 'B', '-82.5', '2026-09-25'],
    ['IPAD AIR GSM', '64', 'A', '-50', '2026-09-25'],
    ['IPAD AIR GSM', '64', 'B', '-60', '2026-09-25'],
    ['IPAD AIR GSM', '128', 'A', '-25', '2026-09-25'],
    ['IPAD AIR GSM', '128', 'B', '-37.5', '2026-09-25'],
    ['IPAD AIR WIFI', '16', 'A', '-100', '2026-09-25'],
    ['IPAD AIR WIFI', '16', 'B', '-105', '2026-09-25'],
    ['IPAD AIR WIFI', '32', 'A', '-75', '2026-09-25'],
    ['IPAD AIR WIFI', '32', 'B', '-82.5', '2026-09-25'],
    ['IPAD AIR WIFI', '64', 'A', '-50', '2026-09-25'],
    ['IPAD AIR WIFI', '64', 'B', '-60', '2026-09-25'],
    ['IPAD AIR WIFI', '128', 'A', '-25', '2026-09-25'],
    ['IPAD AIR WIFI', '128', 'B', '-37.5', '2026-09-25'],
    ['IPAD MINI 2 WIFI', '32', 'A', '-100', '2026-09-25'],
    ['IPAD MINI 2 WIFI', '32', 'B', '-105', '2026-09-25'],
    ['IPAD MINI 3 WIFI', '32', 'A', '-80', '2026-09-25'],
    ['IPAD MINI 3 WIFI', '32', 'B', '-87', '2026-09-25'],
    ['IPAD MINI 4 GSM', '16', 'A', '-70', '2026-09-25'],
    ['IPAD MINI 4 GSM', '16', 'B', '-78', '2026-09-25'],
    ['IPAD MINI 4 GSM', '32', 'A', '-70', '2026-09-25'],
    ['IPAD MINI 4 GSM', '32', 'B', '-78', '2026-09-25'],
    ['IPAD MINI 4 GSM', '64', 'A', '-50', '2026-09-25'],
    ['IPAD MINI 4 GSM', '64', 'B', '-60', '2026-09-25'],
    ['IPAD MINI 4 GSM', '128', 'A', '-25', '2026-09-25'],
    ['IPAD MINI 4 GSM', '128', 'B', '-37.5', '2026-09-25'],
    ['IPAD MINI 4 WIFI', '16', 'A', '-100', '2026-09-25'],
    ['IPAD MINI 4 WIFI', '16', 'B', '-105', '2026-09-25'],
    ['IPAD MINI 4 WIFI', '32', 'A', '-60', '2026-09-25'],
    ['IPAD MINI 4 WIFI', '32', 'B', '-69', '2026-09-25'],
    ['IPAD MINI 4 WIFI', '64', 'A', '-40', '2026-09-25'],
    ['IPAD MINI 4 WIFI', '64', 'B', '-51', '2026-09-25'],
    ['IPAD MINI 4 WIFI', '128', 'A', '-30', '2026-09-25'],
    ['IPAD MINI 4 WIFI', '128', 'B', '-42', '2026-09-25'],
    ['IPAD MINI 5TH GEN GSM', '64', 'A', '0', '2026-09-25'],
    ['IPAD MINI 5TH GEN GSM', '64', 'B', '-15', '2026-09-25'],
    ['IPAD MINI 5TH GEN GSM', '256', 'A', '25', '2026-09-25'],
    ['IPAD MINI 5TH GEN GSM', '256', 'B', '7.5', '2026-09-25'],
    ['IPAD MINI 5TH GEN WIFI', '64', 'A', '50', '2026-09-25'],
    ['IPAD MINI 5TH GEN WIFI', '64', 'B', '30', '2026-09-25'],
    ['IPAD MINI 5TH GEN WIFI', '256', 'A', '75', '2026-09-25'],
    ['IPAD MINI 5TH GEN WIFI', '256', 'B', '52.5', '2026-09-25'],
    ['IPAD MINI WIFI', '16', 'A', '-125', '2026-09-25'],
    ['IPAD MINI WIFI', '16', 'B', '-127.5', '2026-09-25'],
    ['IPAD 8A GEN', '32', 'A', '100', '2026-09-25'],
    ['IPAD 8A GEN', '32', 'B', '75', '2026-09-25'],
    ['IPAD 10A GEN', '64', 'A', '200', '2026-09-25'],
    ['IPAD 10A GEN', '64', 'B', '165', '2026-09-25'],
    ['IPAD PRO 10.5 GSM', '64', 'A', '300', '2026-09-25'],
    ['IPAD PRO 10.5 GSM', '64', 'B', '255', '2026-09-25'],
    ['IPAD PRO 10.5 GSM', '256', 'A', '400', '2026-09-25'],
    ['IPAD PRO 10.5 GSM', '256', 'B', '345', '2026-09-25'],
    ['IPAD PRO 10.5 GSM', '512', 'A', '500', '2026-09-25'],
    ['IPAD PRO 10.5 GSM', '512', 'B', '435', '2026-09-25'],
    ['IPAD PRO 10.5 WIFI', '64', 'A', '350', '2026-09-25'],
    ['IPAD PRO 10.5 WIFI', '64', 'B', '300', '2026-09-25'],
    ['IPAD PRO 10.5 WIFI', '256', 'A', '400', '2026-09-25'],
    ['IPAD PRO 10.5 WIFI', '256', 'B', '345', '2026-09-25'],
    ['IPAD PRO 10.5 WIFI', '512', 'A', '400', '2026-09-25'],
    ['IPAD PRO 10.5 WIFI', '512', 'B', '345', '2026-09-25'],
    ['IPAD PRO 11 GSM', '64', 'A', '450', '2026-09-25'],
    ['IPAD PRO 11 GSM', '64', 'B', '390', '2026-09-25'],
    ['IPAD PRO 11 GSM', '256', 'A', '500', '2026-09-25'],
    ['IPAD PRO 11 GSM', '256', 'B', '435', '2026-09-25'],
    ['IPAD PRO 11 GSM', '512', 'A', '600', '2026-09-25'],
    ['IPAD PRO 11 GSM', '512', 'B', '525', '2026-09-25'],
    ['IPAD PRO 11 GSM', '128', 'A', '700', '2026-09-25'],
    ['IPAD PRO 11 GSM', '128', 'B', '615', '2026-09-25'],
    ['IPAD PRO 11 WIFI', '64', 'A', '500', '2026-09-25'],
    ['IPAD PRO 11 WIFI', '64', 'B', '435', '2026-09-25'],
    ['IPAD PRO 11 WIFI', '256', 'A', '600', '2026-09-25'],
    ['IPAD PRO 11 WIFI', '256', 'B', '525', '2026-09-25'],
    ['IPAD PRO 11 WIFI', '512', 'A', '600', '2026-09-25'],
    ['IPAD PRO 11 WIFI', '512', 'B', '525', '2026-09-25'],
    ['IPAD PRO 11 WIFI', '128', 'A', '700', '2026-09-25'],
    ['IPAD PRO 11 WIFI', '128', 'B', '615', '2026-09-25'],
    ['IPAD PRO 12.9 1ST GEN GSM', '32', 'A', '300', '2026-09-25'],
    ['IPAD PRO 12.9 1ST GEN GSM', '32', 'B', '255', '2026-09-25'],
    ['IPAD PRO 12.9 1ST GEN GSM', '128', 'A', '400', '2026-09-25'],
    ['IPAD PRO 12.9 1ST GEN GSM', '128', 'B', '345', '2026-09-25'],
    ['IPAD PRO 12.9 1ST GEN GSM', '256', 'A', '500', '2026-09-25'],
    ['IPAD PRO 12.9 1ST GEN GSM', '256', 'B', '435', '2026-09-25'],
    ['IPAD PRO 12.9 1ST GEN WIFI', '32', 'A', '439', '2026-09-25'],
    ['IPAD PRO 12.9 1ST GEN WIFI', '32', 'B', '380.1', '2026-09-25'],
    ['IPAD PRO 12.9 1ST GEN WIFI', '128', 'A', '475', '2026-09-25'],
    ['IPAD PRO 12.9 1ST GEN WIFI', '128', 'B', '412.5', '2026-09-25'],
    ['IPAD PRO 12.9 1ST GEN WIFI', '256', 'A', '556', '2026-09-25'],
    ['IPAD PRO 12.9 1ST GEN WIFI', '256', 'B', '485.4', '2026-09-25'],
    ['IPAD PRO 12.9 2ND GEN GSM', '64', 'A', '646', '2026-09-25'],
    ['IPAD PRO 12.9 2ND GEN GSM', '64', 'B', '566.4', '2026-09-25'],
    ['IPAD PRO 12.9 2ND GEN GSM', '256', 'A', '682', '2026-09-25'],
    ['IPAD PRO 12.9 2ND GEN GSM', '256', 'B', '598.8', '2026-09-25'],
    ['IPAD PRO 12.9 2ND GEN GSM', '512', 'A', '691', '2026-09-25'],
    ['IPAD PRO 12.9 2ND GEN GSM', '512', 'B', '606.9', '2026-09-25'],
    ['IPAD PRO 12.9 2ND GEN WIFI', '64', 'A', '628', '2026-09-25'],
    ['IPAD PRO 12.9 2ND GEN WIFI', '64', 'B', '550.2', '2026-09-25'],
    ['IPAD PRO 12.9 2ND GEN WIFI', '256', 'A', '700', '2026-09-25'],
    ['IPAD PRO 12.9 2ND GEN WIFI', '256', 'B', '615', '2026-09-25'],
    ['IPAD PRO 12.9 2ND GEN WIFI', '512', 'A', '700', '2026-09-25'],
    ['IPAD PRO 12.9 2ND GEN WIFI', '512', 'B', '615', '2026-09-25'],
    ['IPAD PRO 12.9 3RD GEN GSM', '64', 'A', '790', '2026-09-25'],
    ['IPAD PRO 12.9 3RD GEN GSM', '64', 'B', '696', '2026-09-25'],
    ['IPAD PRO 12.9 3RD GEN GSM', '256', 'A', '880', '2026-09-25'],
    ['IPAD PRO 12.9 3RD GEN GSM', '256', 'B', '777', '2026-09-25'],
    ['IPAD PRO 12.9 3RD GEN GSM', '512', 'A', '970', '2026-09-25'],
    ['IPAD PRO 12.9 3RD GEN GSM', '512', 'B', '858', '2026-09-25'],
    ['IPAD PRO 12.9 3RD GEN GSM', '128', 'A', '1.060.000', '2026-09-25'],
    ['IPAD PRO 12.9 3RD GEN GSM', '128', 'B', '939', '2026-09-25'],
    ['IPAD PRO 12.9 3RD GEN WIFI', '64', 'A', '745', '2026-09-25'],
    ['IPAD PRO 12.9 3RD GEN WIFI', '64', 'B', '655.5', '2026-09-25'],
    ['IPAD PRO 12.9 3RD GEN WIFI', '256', 'A', '880', '2026-09-25'],
    ['IPAD PRO 12.9 3RD GEN WIFI', '256', 'B', '777', '2026-09-25'],
    ['IPAD PRO 12.9 3RD GEN WIFI', '512', 'A', '970', '2026-09-25'],
    ['IPAD PRO 12.9 3RD GEN WIFI', '512', 'B', '858', '2026-09-25'],
    ['IPAD PRO 12.9 3RD GEN WIFI', '128', 'A', '1.060.000', '2026-09-25'],
    ['IPAD PRO 12.9 3RD GEN WIFI', '128', 'B', '939', '2026-09-25'],
    ['IPAD PRO 9.7 GSM', '32', 'A', '205', '2026-09-25'],
    ['IPAD PRO 9.7 GSM', '32', 'B', '169.5', '2026-09-25'],
    ['IPAD PRO 9.7 GSM', '128', 'A', '268', '2026-09-25'],
    ['IPAD PRO 9.7 GSM', '128', 'B', '226.2', '2026-09-25'],
    ['IPAD PRO 9.7 GSM', '256', 'A', '295', '2026-09-25'],
    ['IPAD PRO 9.7 GSM', '256', 'B', '250.5', '2026-09-25'],
    ['IPAD PRO 9.7 WIFI', '32', 'A', '205', '2026-09-25'],
    ['IPAD PRO 9.7 WIFI', '32', 'B', '169.5', '2026-09-25'],
    ['IPAD PRO 9.7 WIFI', '128', 'A', '250', '2026-09-25'],
    ['IPAD PRO 9.7 WIFI', '128', 'B', '210', '2026-09-25'],
    ['IPAD PRO 9.7 WIFI', '256', 'A', '295', '2026-09-25'],
    ['IPAD PRO 9.7 WIFI', '256', 'B', '250.5', '2026-09-25'],
    ['Watch SE', '0', 'A', '50', '2026-09-25'],
    ['Watch SE', '0', 'B', '30', '2026-09-25'],
    ['Watch SE2 40mm', '0', 'A', '100', '2026-09-25'],
    ['Watch SE2 40mm', '0', 'B', '75', '2026-09-25'],
    ['Watch SE2 44mm', '0', 'A', '150', '2026-09-25'],
    ['Watch SE2 44mm', '0', 'B', '120', '2026-09-25'],
    ['Watch Serie 1', '0', 'A', '-130', '2026-09-25'],
    ['Watch Serie 1', '0', 'B', '-132', '2026-09-25'],
    ['Watch Serie 2', '0', 'A', '-100', '2026-09-25'],
    ['Watch Serie 2', '0', 'B', '-105', '2026-09-25'],
    ['Watch Serie 3', '0', 'A', '-25', '2026-09-25'],
    ['Watch Serie 3', '0', 'B', '-37.5', '2026-09-25'],
    ['Watch Serie 4', '0', 'A', '50', '2026-09-25'],
    ['Watch Serie 4', '0', 'B', '30', '2026-09-25'],
    ['Watch Serie 5', '0', 'A', '100', '2026-09-25'],
    ['Watch Serie 5', '0', 'B', '75', '2026-09-25'],
    ['Watch Serie 6', '0', 'A', '150', '2026-09-25'],
    ['Watch Serie 6', '0', 'B', '120', '2026-09-25'],
    ['Watch Serie 7', '0', 'A', '200', '2026-09-25'],
    ['Watch Serie 7', '0', 'B', '165', '2026-09-25'],
    ['Watch Serie 8', '0', 'A', '300', '2026-09-25'],
    ['Watch Serie 8', '0', 'B', '255', '2026-09-25'],
    ['Watch Serie 9', '0', 'A', '400', '2026-09-25'],
    ['Watch Serie 9', '0', 'B', '345', '2026-09-25'],
    ['Macbook Air 2020 M1 512', '0', 'A', '1.200.000', '2026-09-25'],
    ['Macbook Air 2020 M1 512', '0', 'B', '1.065.000', '2026-09-25'],
    ['Magic 5 Lite', '0', 'A', '50', '2026-09-25'],
    ['Magic 5 Lite', '0', 'B', '30', '2026-09-25'],
    ['Magic 6 Lite', '0', 'A', '100', '2026-09-25'],
    ['Magic 6 Lite', '0', 'B', '75', '2026-09-25'],
    ['X7a', '0', 'A', '0', '2026-09-25'],
    ['X7a', '0', 'B', '-15', '2026-09-25'],
    ['X7b', '0', 'A', '50', '2026-09-25'],
    ['X7b', '0', 'B', '30', '2026-09-25'],
    ['X8a', '0', 'A', '20', '2026-09-25'],
    ['X8a', '0', 'B', '3', '2026-09-25'],
    ['X8b 256', '0', 'A', '100', '2026-09-25'],
    ['X8b 256', '0', 'B', '75', '2026-09-25'],
    ['MATE 10', '0', 'A', '-65', '2026-09-25'],
    ['MATE 10', '0', 'B', '-73.5', '2026-09-25'],
    ['MATE 10 LITE', '0', 'A', '-83', '2026-09-25'],
    ['MATE 10 LITE', '0', 'B', '-89.7', '2026-09-25'],
    ['MATE 10 PRO', '0', 'A', '-20', '2026-09-25'],
    ['MATE 10 PRO', '0', 'B', '-33', '2026-09-25'],
    ['MATE 20', '0', 'A', '0', '2026-09-25'],
    ['MATE 20', '0', 'B', '-15', '2026-09-25'],
    ['MATE 20 LITE', '0', 'A', '-65', '2026-09-25'],
    ['MATE 20 LITE', '0', 'B', '-73.5', '2026-09-25'],
    ['MATE 20 Pro', '0', 'A', '0', '2026-09-25'],
    ['MATE 20 Pro', '0', 'B', '-15', '2026-09-25'],
    ['MATE 30', '0', 'A', '20', '2026-09-25'],
    ['MATE 30', '0', 'B', '3', '2026-09-25'],
    ['MATE 30 Pro', '0', 'A', '100', '2026-09-25'],
    ['MATE 30 Pro', '0', 'B', '75', '2026-09-25'],
    ['MATE 9', '0', 'A', '-92', '2026-09-25'],
    ['MATE 9', '0', 'B', '-97.8', '2026-09-25'],
    ['MATE 9 LITE', '0', 'A', '-74', '2026-09-25'],
    ['MATE 9 LITE', '0', 'B', '-81.6', '2026-09-25'],
    ['NOVA 5T', '0', 'A', '-65', '2026-09-25'],
    ['NOVA 5T', '0', 'B', '-73.5', '2026-09-25'],
    ['P SMART', '0', 'A', '-110', '2026-09-25'],
    ['P SMART', '0', 'B', '-114', '2026-09-25'],
    ['P SMART (2019)', '0', 'A', '-92', '2026-09-25'],
    ['P SMART (2019)', '0', 'B', '-97.8', '2026-09-25'],
    ['P10', '0', 'A', '-92', '2026-09-25'],
    ['P10', '0', 'B', '-97.8', '2026-09-25'],
    ['P10 LITE', '0', 'A', '-110', '2026-09-25'],
    ['P10 LITE', '0', 'B', '-114', '2026-09-25'],
    ['P10 PLUS', '0', 'A', '-74', '2026-09-25'],
    ['P10 PLUS', '0', 'B', '-81.6', '2026-09-25'],
    ['P20', '0', 'A', '-38', '2026-09-25'],
    ['P20', '0', 'B', '-49.2', '2026-09-25'],
    ['P20 LITE', '0', 'A', '-74', '2026-09-25'],
    ['P20 LITE', '0', 'B', '-81.6', '2026-09-25'],
    ['P20 PRO', '0', 'A', '-20', '2026-09-25'],
    ['P20 PRO', '0', 'B', '-33', '2026-09-25'],
    ['P30', '0', 'A', '25', '2026-09-25'],
    ['P30', '0', 'B', '7.5', '2026-09-25'],
    ['P30 LITE', '0', 'A', '-65', '2026-09-25'],
    ['P30 LITE', '0', 'B', '-73.5', '2026-09-25'],
    ['P30 PRO', '0', 'A', '70', '2026-09-25'],
    ['P30 PRO', '0', 'B', '48', '2026-09-25'],
    ['P40', '0', 'A', '70', '2026-09-25'],
    ['P40', '0', 'B', '48', '2026-09-25'],
    ['P40 Lite', '0', 'A', '-20', '2026-09-25'],
    ['P40 Lite', '0', 'B', '-33', '2026-09-25'],
    ['P40 Pro', '0', 'A', '115', '2026-09-25'],
    ['P40 Pro', '0', 'B', '88.5', '2026-09-25'],
    ['P60 Pro', '0', 'A', '850', '2026-09-25'],
    ['P60 Pro', '0', 'B', '750', '2026-09-25'],
    ['P9', '0', 'A', '-128', '2026-09-25'],
    ['P9', '0', 'B', '-130.2', '2026-09-25'],
    ['P9 LITE', '0', 'A', '-137', '2026-09-25'],
    ['P9 LITE', '0', 'B', '-138.3', '2026-09-25'],
    ['P9 PLUS', '0', 'A', '-110', '2026-09-25'],
    ['P9 PLUS', '0', 'B', '-114', '2026-09-25'],
    ['Y5 (2017)', '0', 'A', '-155', '2026-09-25'],
    ['Y5 (2017)', '0', 'B', '-154.5', '2026-09-25'],
    ['Y5 (2018)', '0', 'A', '-128', '2026-09-25'],
    ['Y5 (2018)', '0', 'B', '-130.2', '2026-09-25'],
    ['Y5 (2019)', '0', 'A', '-119', '2026-09-25'],
    ['Y5 (2019)', '0', 'B', '-122.1', '2026-09-25'],
    ['Y6 (2017)', '0', 'A', '-110', '2026-09-25'],
    ['Y6 (2017)', '0', 'B', '-114', '2026-09-25'],
    ['Y6 (2018)', '0', 'A', '-92', '2026-09-25'],
    ['Y6 (2018)', '0', 'B', '-97.8', '2026-09-25'],
    ['Y6 (2019)', '0', 'A', '-65', '2026-09-25'],
    ['Y6 (2019)', '0', 'B', '-73.5', '2026-09-25'],
    ['Y6P', '0', 'A', '-65', '2026-09-25'],
    ['Y6P', '0', 'B', '-73.5', '2026-09-25'],
    ['Y7 (2017)', '0', 'A', '-110', '2026-09-25'],
    ['Y7 (2017)', '0', 'B', '-114', '2026-09-25'],
    ['Y7 (2018)', '0', 'A', '-92', '2026-09-25'],
    ['Y7 (2018)', '0', 'B', '-97.8', '2026-09-25'],
    ['Y7 (2019)', '0', 'A', '-65', '2026-09-25'],
    ['Y7 (2019)', '0', 'B', '-73.5', '2026-09-25'],
    ['Y7 PRIME', '0', 'A', '-74', '2026-09-25'],
    ['Y7 PRIME', '0', 'B', '-81.6', '2026-09-25'],
    ['Y7A', '0', 'A', '-65', '2026-09-25'],
    ['Y7A', '0', 'B', '-73.5', '2026-09-25'],
    ['Y7P', '0', 'A', '-65', '2026-09-25'],
    ['Y7P', '0', 'B', '-73.5', '2026-09-25'],
    ['Y8P', '0', 'A', '-38', '2026-09-25'],
    ['Y8P', '0', 'B', '-49.2', '2026-09-25'],
    ['Y9 (2018)', '0', 'A', '-65', '2026-09-25'],
    ['Y9 (2018)', '0', 'B', '-73.5', '2026-09-25'],
    ['Y9 (2019)', '0', 'A', '-65', '2026-09-25'],
    ['Y9 (2019)', '0', 'B', '-73.5', '2026-09-25'],
    ['Y9 PRIME', '0', 'A', '-56', '2026-09-25'],
    ['Y9 PRIME', '0', 'B', '-65.4', '2026-09-25'],
    ['Y9A', '0', 'A', '-50', '2026-09-25'],
    ['Y9A', '0', 'B', '-60', '2026-09-25'],
    ['Y9S (2019)', '0', 'A', '-20', '2026-09-25'],
    ['Y9S (2019)', '0', 'B', '-33', '2026-09-25'],
    ['Nova 8', '0', 'A', '-20', '2026-09-25'],
    ['Nova 8', '0', 'B', '-33', '2026-09-25'],
    ['Nova 8i', '0', 'A', '0', '2026-09-25'],
    ['Nova 8i', '0', 'B', '-15', '2026-09-25'],
    ['Nova 9', '0', 'A', '20', '2026-09-25'],
    ['Nova 9', '0', 'B', '3', '2026-09-25'],
    ['Nova 10 Pro', '0', 'A', '500', '2026-09-25'],
    ['Nova 10 Pro', '0', 'B', '435', '2026-09-25'],
    ['Nova 11i', '0', 'A', '25', '2026-09-25'],
    ['Nova 11i', '0', 'B', '7.5', '2026-09-25'],
    ['Nova Y91', '0', 'A', '80', '2026-09-25'],
    ['Nova Y91', '0', 'B', '57', '2026-09-25'],
    ['MOTO E13', '0', 'A', '-80', '2026-09-25'],
    ['MOTO E13', '0', 'B', '-87', '2026-09-25'],
    ['MOTO E14', '0', 'A', '-60', '2026-09-25'],
    ['MOTO E14', '0', 'B', '-69', '2026-09-25'],
    ['MOTO E20', '0', 'A', '-92', '2026-09-25'],
    ['MOTO E20', '0', 'B', '-97.8', '2026-09-25'],
    ['MOTO E32', '0', 'A', '-75', '2026-09-25'],
    ['MOTO E32', '0', 'B', '-82.5', '2026-09-25'],
    ['MOTO E40', '0', 'A', '-65', '2026-09-25'],
    ['MOTO E40', '0', 'B', '-73.5', '2026-09-25'],
    ['MOTO E4', '0', 'A', '-128', '2026-09-25'],
    ['MOTO E4', '0', 'B', '-130.2', '2026-09-25'],
    ['MOTO E4 PLUS', '0', 'A', '-119', '2026-09-25'],
    ['MOTO E4 PLUS', '0', 'B', '-122.1', '2026-09-25'],
    ['MOTO E5', '0', 'A', '-110', '2026-09-25'],
    ['MOTO E5', '0', 'B', '-114', '2026-09-25'],
    ['MOTO E5 PLAY', '0', 'A', '-110', '2026-09-25'],
    ['MOTO E5 PLAY', '0', 'B', '-114', '2026-09-25'],
    ['MOTO E5 PLUS', '0', 'A', '-110', '2026-09-25'],
    ['MOTO E5 PLUS', '0', 'B', '-114', '2026-09-25'],
    ['MOTO E6', '0', 'A', '-92', '2026-09-25'],
    ['MOTO E6', '0', 'B', '-97.8', '2026-09-25'],
    ['MOTO E6 PLAY', '0', 'A', '-92', '2026-09-25'],
    ['MOTO E6 PLAY', '0', 'B', '-97.8', '2026-09-25'],
    ['MOTO E6 PLUS', '0', 'A', '-92', '2026-09-25'],
    ['MOTO E6 PLUS', '0', 'B', '-97.8', '2026-09-25'],
    ['MOTO G100', '0', 'A', '160', '2026-09-25'],
    ['MOTO G100', '0', 'B', '129', '2026-09-25'],
    ['MOTO G13', '0', 'A', '0', '2026-09-25'],
    ['MOTO G13', '0', 'B', '-15', '2026-09-25'],
    ['MOTO G20', '0', 'A', '-65', '2026-09-25'],
    ['MOTO G20', '0', 'B', '-73.5', '2026-09-25'],
    ['MOTO G22', '0', 'A', '-20', '2026-09-25'],
    ['MOTO G22', '0', 'B', '-33', '2026-09-25'],
    ['MOTO G24', '0', 'A', '50', '2026-09-25'],
    ['MOTO G24', '0', 'B', '30', '2026-09-25'],
    ['MOTO G30', '0', 'A', '-65', '2026-09-25'],
    ['MOTO G30', '0', 'B', '-73.5', '2026-09-25'],
    ['MOTO G32', '0', 'A', '20', '2026-09-25'],
    ['MOTO G32', '0', 'B', '3', '2026-09-25'],
    ['MOTO G34', '0', 'A', '40', '2026-09-25'],
    ['MOTO G34', '0', 'B', '21', '2026-09-25'],
    ['MOTO G41', '0', 'A', '-20', '2026-09-25'],
    ['MOTO G41', '0', 'B', '-33', '2026-09-25'],
    ['MOTO G50', '0', 'A', '-20', '2026-09-25'],
    ['MOTO G50', '0', 'B', '-33', '2026-09-25'],
    ['MOTO G54', '0', 'A', '20', '2026-09-25'],
    ['MOTO G54', '0', 'B', '3', '2026-09-25'],
    ['MOTO G60', '0', 'A', '2.5', '2026-09-25'],
    ['MOTO G60', '0', 'B', '-12.75', '2026-09-25'],
    ['MOTO G60S', '0', 'A', '25', '2026-09-25'],
    ['MOTO G60S', '0', 'B', '7.5', '2026-09-25'],
    ['MOTO G71', '0', 'A', '50', '2026-09-25'],
    ['MOTO G71', '0', 'B', '30', '2026-09-25'],
    ['MOTO G72', '0', 'A', '50', '2026-09-25'],
    ['MOTO G72', '0', 'B', '30', '2026-09-25'],
    ['MOTO G84', '0', 'A', '60', '2026-09-25'],
    ['MOTO G84', '0', 'B', '39', '2026-09-25'],
    ['MOTO G85', '0', 'A', '100', '2026-09-25'],
    ['MOTO G85', '0', 'B', '75', '2026-09-25'],
    ['MOTO G04', '0', 'A', '-50', '2026-09-25'],
    ['MOTO G04', '0', 'B', '-60', '2026-09-25'],
    ['MOTO G5', '0', 'A', '-110', '2026-09-25'],
    ['MOTO G5', '0', 'B', '-114', '2026-09-25'],
    ['MOTO G5 PLUS', '0', 'A', '-110', '2026-09-25'],
    ['MOTO G5 PLUS', '0', 'B', '-114', '2026-09-25'],
    ['MOTO G5S', '0', 'A', '-110', '2026-09-25'],
    ['MOTO G5S', '0', 'B', '-114', '2026-09-25'],
    ['MOTO G5S PLUS', '0', 'A', '-110', '2026-09-25'],
    ['MOTO G5S PLUS', '0', 'B', '-114', '2026-09-25'],
    ['MOTO G6', '0', 'A', '-92', '2026-09-25'],
    ['MOTO G6', '0', 'B', '-97.8', '2026-09-25'],
    ['MOTO G6 PLAY', '0', 'A', '-110', '2026-09-25'],
    ['MOTO G6 PLAY', '0', 'B', '-114', '2026-09-25'],
    ['MOTO G6 PLUS', '0', 'A', '-74', '2026-09-25'],
    ['MOTO G6 PLUS', '0', 'B', '-81.6', '2026-09-25'],
    ['MOTO G7', '0', 'A', '-65', '2026-09-25'],
    ['MOTO G7', '0', 'B', '-73.5', '2026-09-25'],
    ['MOTO G7 PLAY', '0', 'A', '-74', '2026-09-25'],
    ['MOTO G7 PLAY', '0', 'B', '-81.6', '2026-09-25'],
    ['MOTO G7 PLUS', '0', 'A', '-56', '2026-09-25'],
    ['MOTO G7 PLUS', '0', 'B', '-65.4', '2026-09-25'],
    ['MOTO G7 POWER', '0', 'A', '-65', '2026-09-25'],
    ['MOTO G7 POWER', '0', 'B', '-73.5', '2026-09-25'],
    ['MOTO G8', '0', 'A', '-47', '2026-09-25'],
    ['MOTO G8', '0', 'B', '-57.3', '2026-09-25'],
    ['MOTO G8 PLAY', '0', 'A', '-65', '2026-09-25'],
    ['MOTO G8 PLAY', '0', 'B', '-73.5', '2026-09-25'],
    ['MOTO G8 PLUS', '0', 'A', '-38', '2026-09-25'],
    ['MOTO G8 PLUS', '0', 'B', '-49.2', '2026-09-25'],
    ['MOTO G8 Power', '0', 'A', '-56', '2026-09-25'],
    ['MOTO G8 Power', '0', 'B', '-65.4', '2026-09-25'],
    ['MOTO G9 PLAY', '0', 'A', '-65', '2026-09-25'],
    ['MOTO G9 PLAY', '0', 'B', '-73.5', '2026-09-25'],
    ['MOTO G9 POWER', '0', 'A', '-65', '2026-09-25'],
    ['MOTO G9 POWER', '0', 'B', '-73.5', '2026-09-25'],
    ['MOTO G9 PLUS', '0', 'A', '-38', '2026-09-25'],
    ['MOTO G9 PLUS', '0', 'B', '-49.2', '2026-09-25'],
    ['MOTO Edge 20 Lite', '0', 'A', '0', '2026-09-25'],
    ['MOTO Edge 20 Lite', '0', 'B', '-15', '2026-09-25'],
    ['MOTO Edge 20 Pro', '0', 'A', '150', '2026-09-25'],
    ['MOTO Edge 20 Pro', '0', 'B', '120', '2026-09-25'],
    ['MOTO Edge 30 Neo', '0', 'A', '100', '2026-09-25'],
    ['MOTO Edge 30 Neo', '0', 'B', '75', '2026-09-25'],
    ['MOTO Edge 30 Fusion', '0', 'A', '250', '2026-09-25'],
    ['MOTO Edge 30 Fusion', '0', 'B', '210', '2026-09-25'],
    ['MOTO Edge 30', '0', 'A', '150', '2026-09-25'],
    ['MOTO Edge 30', '0', 'B', '120', '2026-09-25'],
    ['MOTO Edge 40', '0', 'A', '150', '2026-09-25'],
    ['MOTO Edge 40', '0', 'B', '120', '2026-09-25'],
    ['MOTO Edge 40 Neo', '0', 'A', '150', '2026-09-25'],
    ['MOTO Edge 40 Neo', '0', 'B', '120', '2026-09-25'],
    ['MOTO Edge 50', '0', 'A', '250', '2026-09-25'],
    ['MOTO Edge 50', '0', 'B', '210', '2026-09-25'],
    ['MOTO Edge 50 Fusion', '0', 'A', '150', '2026-09-25'],
    ['MOTO Edge 50 Fusion', '0', 'B', '120', '2026-09-25'],
    ['MOTO Edge 50 Pro', '0', 'A', '300', '2026-09-25'],
    ['MOTO Edge 50 Pro', '0', 'B', '255', '2026-09-25'],
    ['MOTO ONE', '0', 'A', '-70', '2026-09-25'],
    ['MOTO ONE', '0', 'B', '-78', '2026-09-25'],
    ['MOTO ONE ACTION', '0', 'A', '-65', '2026-09-25'],
    ['MOTO ONE ACTION', '0', 'B', '-73.5', '2026-09-25'],
    ['MOTO ONE FUSION', '0', 'A', '-65', '2026-09-25'],
    ['MOTO ONE FUSION', '0', 'B', '-73.5', '2026-09-25'],
    ['MOTO ONE HYPER', '0', 'A', '-50', '2026-09-25'],
    ['MOTO ONE HYPER', '0', 'B', '-60', '2026-09-25'],
    ['MOTO ONE MACRO', '0', 'A', '-50', '2026-09-25'],
    ['MOTO ONE MACRO', '0', 'B', '-60', '2026-09-25'],
    ['MOTO ONE VISION', '0', 'A', '-40', '2026-09-25'],
    ['MOTO ONE VISION', '0', 'B', '-51', '2026-09-25'],
    ['MOTO ONE ZOOM', '0', 'A', '-40', '2026-09-25'],
    ['MOTO ONE ZOOM', '0', 'B', '-51', '2026-09-25'],
    ['MOTO RAZR PLUS 2023', '0', 'A', '400', '2026-09-25'],
    ['MOTO RAZR PLUS 2023', '0', 'B', '345', '2026-09-25'],
    ['X5 Pro', '0', 'A', '200', '2026-09-25'],
    ['X5 Pro', '0', 'B', '165', '2026-09-25'],
    ['X6', '0', 'A', '50', '2026-09-25'],
    ['X6', '0', 'B', '30', '2026-09-25'],
    ['M3', '0', 'A', '50', '2026-09-25'],
    ['M3', '0', 'B', '30', '2026-09-25'],
    ['F3', '0', 'A', '100', '2026-09-25'],
    ['F3', '0', 'B', '75', '2026-09-25'],
    ['GALAXY A01', '0', 'A', '-110', '2026-09-25'],
    ['GALAXY A01', '0', 'B', '-114', '2026-09-25'],
    ['GALAXY A02', '0', 'A', '-110', '2026-09-25'],
    ['GALAXY A02', '0', 'B', '-114', '2026-09-25'],
    ['GALAXY A3 (2017)', '0', 'A', '-110', '2026-09-25'],
    ['GALAXY A3 (2017)', '0', 'B', '-114', '2026-09-25'],
    ['GALAXY A04s', '0', 'A', '-20', '2026-09-25'],
    ['GALAXY A04s', '0', 'B', '-33', '2026-09-25'],
    ['GALAXY A10', '0', 'A', '-92', '2026-09-25'],
    ['GALAXY A10', '0', 'B', '-97.8', '2026-09-25'],
    ['GALAXY A10s', '0', 'A', '-92', '2026-09-25'],
    ['GALAXY A10s', '0', 'B', '-97.8', '2026-09-25'],
    ['GALAXY A11', '0', 'A', '-65', '2026-09-25'],
    ['GALAXY A11', '0', 'B', '-73.5', '2026-09-25'],
    ['GALAXY A12', '0', 'A', '-65', '2026-09-25'],
    ['GALAXY A12', '0', 'B', '-73.5', '2026-09-25'],
    ['GALAXY A13', '0', 'A', '0', '2026-09-25'],
    ['GALAXY A13', '0', 'B', '-15', '2026-09-25'],
    ['GALAXY A14', '0', 'A', '50', '2026-09-25'],
    ['GALAXY A14', '0', 'B', '30', '2026-09-25'],
    ['GALAXY A15', '0', 'A', '40', '2026-09-25'],
    ['GALAXY A15', '0', 'B', '21', '2026-09-25'],
    ['GALAXY A20', '0', 'A', '-65', '2026-09-25'],
    ['GALAXY A20', '0', 'B', '-73.5', '2026-09-25'],
    ['GALAXY A20s', '0', 'A', '-65', '2026-09-25'],
    ['GALAXY A20s', '0', 'B', '-73.5', '2026-09-25'],
    ['GALAXY A21s', '0', 'A', '-38', '2026-09-25'],
    ['GALAXY A21s', '0', 'B', '-49.2', '2026-09-25'],
    ['GALAXY A22', '0', 'A', '-20', '2026-09-25'],
    ['GALAXY A22', '0', 'B', '-33', '2026-09-25'],
    ['GALAXY A23', '0', 'A', '0', '2026-09-25'],
    ['GALAXY A23', '0', 'B', '-15', '2026-09-25'],
    ['GALAXY A24', '0', 'A', '30', '2026-09-25'],
    ['GALAXY A24', '0', 'B', '12', '2026-09-25'],
    ['GALAXY A25', '0', 'A', '70', '2026-09-25'],
    ['GALAXY A25', '0', 'B', '48', '2026-09-25'],
    ['GALAXY A30', '0', 'A', '-87.5', '2026-09-25'],
    ['GALAXY A30', '0', 'B', '-93.75', '2026-09-25'],
    ['GALAXY A30s', '0', 'A', '-65', '2026-09-25'],
    ['GALAXY A30s', '0', 'B', '-73.5', '2026-09-25'],
    ['GALAXY A31', '0', 'A', '-20', '2026-09-25'],
    ['GALAXY A31', '0', 'B', '-33', '2026-09-25'],
    ['GALAXY A32', '0', 'A', '0', '2026-09-25'],
    ['GALAXY A32', '0', 'B', '-15', '2026-09-25'],
    ['GALAXY A33', '0', 'A', '100', '2026-09-25'],
    ['GALAXY A33', '0', 'B', '75', '2026-09-25'],
    ['GALAXY A34', '0', 'A', '150', '2026-09-25'],
    ['GALAXY A34', '0', 'B', '120', '2026-09-25'],
    ['GALAXY A35', '0', 'A', '200', '2026-09-25'],
    ['GALAXY A35', '0', 'B', '165', '2026-09-25'],
    ['GALAXY A5 (2017)', '0', 'A', '-110', '2026-09-25'],
    ['GALAXY A5 (2017)', '0', 'B', '-114', '2026-09-25'],
    ['GALAXY A50', '0', 'A', '-38', '2026-09-25'],
    ['GALAXY A50', '0', 'B', '-49.2', '2026-09-25'],
    ['GALAXY A51', '0', 'A', '25', '2026-09-25'],
    ['GALAXY A51', '0', 'B', '7.5', '2026-09-25'],
    ['GALAXY A52', '0', 'A', '52', '2026-09-25'],
    ['GALAXY A52', '0', 'B', '31.8', '2026-09-25'],
    ['GALAXY A53', '0', 'A', '100', '2026-09-25'],
    ['GALAXY A53', '0', 'B', '75', '2026-09-25'],
    ['GALAXY A54', '0', 'A', '120', '2026-09-25'],
    ['GALAXY A54', '0', 'B', '93', '2026-09-25'],
    ['GALAXY A55', '0', 'A', '150', '2026-09-25'],
    ['GALAXY A55', '0', 'B', '120', '2026-09-25'],
    ['GALAXY A6 (2018)', '0', 'A', '-92', '2026-09-25'],
    ['GALAXY A6 (2018)', '0', 'B', '-97.8', '2026-09-25'],
    ['GALAXY A6+ (2018)', '0', 'A', '-92', '2026-09-25'],
    ['GALAXY A6+ (2018)', '0', 'B', '-97.8', '2026-09-25'],
    ['GALAXY A7 (2017)', '0', 'A', '-65', '2026-09-25'],
    ['GALAXY A7 (2017)', '0', 'B', '-73.5', '2026-09-25'],
    ['GALAXY A7 (2018)', '0', 'A', '-65', '2026-09-25'],
    ['GALAXY A7 (2018)', '0', 'B', '-73.5', '2026-09-25'],
    ['GALAXY A70', '0', 'A', '25', '2026-09-25'],
    ['GALAXY A70', '0', 'B', '7.5', '2026-09-25'],
    ['GALAXY A71', '0', 'A', '70', '2026-09-25'],
    ['GALAXY A71', '0', 'B', '48', '2026-09-25'],
    ['GALAXY A72', '0', 'A', '115', '2026-09-25'],
    ['GALAXY A72', '0', 'B', '88.5', '2026-09-25'],
    ['GALAXY A8 (2018)', '0', 'A', '-65', '2026-09-25'],
    ['GALAXY A8 (2018)', '0', 'B', '-73.5', '2026-09-25'],
    ['GALAXY A8+ (2018)', '0', 'A', '-47', '2026-09-25'],
    ['GALAXY A8+ (2018)', '0', 'B', '-57.3', '2026-09-25'],
    ['GALAXY A80', '0', 'A', '25', '2026-09-25'],
    ['GALAXY A80', '0', 'B', '7.5', '2026-09-25'],
    ['GALAXY A9 (2018)', '0', 'A', '-65', '2026-09-25'],
    ['GALAXY A9 (2018)', '0', 'B', '-73.5', '2026-09-25'],
    ['GALAXY J4', '0', 'A', '-110', '2026-09-25'],
    ['GALAXY J4', '0', 'B', '-114', '2026-09-25'],
    ['GALAXY J4 CORE', '0', 'A', '-110', '2026-09-25'],
    ['GALAXY J4 CORE', '0', 'B', '-114', '2026-09-25'],
    ['GALAXY J4 PLUS', '0', 'A', '-92', '2026-09-25'],
    ['GALAXY J4 PLUS', '0', 'B', '-97.8', '2026-09-25'],
    ['GALAXY J5 METAL', '0', 'A', '-128', '2026-09-25'],
    ['GALAXY J5 METAL', '0', 'B', '-130.2', '2026-09-25'],
    ['GALAXY J5 PRIME', '0', 'A', '-128', '2026-09-25'],
    ['GALAXY J5 PRIME', '0', 'B', '-130.2', '2026-09-25'],
    ['GALAXY J5 PRO', '0', 'A', '-128', '2026-09-25'],
    ['GALAXY J5 PRO', '0', 'B', '-130.2', '2026-09-25'],
    ['GALAXY J6', '0', 'A', '-110', '2026-09-25'],
    ['GALAXY J6', '0', 'B', '-114', '2026-09-25'],
    ['GALAXY J6+', '0', 'A', '-110', '2026-09-25'],
    ['GALAXY J6+', '0', 'B', '-114', '2026-09-25'],
    ['GALAXY J7 (2016)', '0', 'A', '-128', '2026-09-25'],
    ['GALAXY J7 (2016)', '0', 'B', '-130.2', '2026-09-25'],
    ['GALAXY J7 (2017)', '0', 'A', '-128', '2026-09-25'],
    ['GALAXY J7 (2017)', '0', 'B', '-130.2', '2026-09-25'],
    ['GALAXY J7 (2018)', '0', 'A', '-128', '2026-09-25'],
    ['GALAXY J7 (2018)', '0', 'B', '-130.2', '2026-09-25'],
    ['GALAXY J7 NEO', '0', 'A', '-128', '2026-09-25'],
    ['GALAXY J7 NEO', '0', 'B', '-130.2', '2026-09-25'],
    ['GALAXY J7 PRIME', '0', 'A', '-128', '2026-09-25'],
    ['GALAXY J7 PRIME', '0', 'B', '-130.2', '2026-09-25'],
    ['GALAXY J7 PRO', '0', 'A', '-128', '2026-09-25'],
    ['GALAXY J7 PRO', '0', 'B', '-130.2', '2026-09-25'],
    ['GALAXY J8', '0', 'A', '-110', '2026-09-25'],
    ['GALAXY J8', '0', 'B', '-114', '2026-09-25'],
    ['GALAXY M12', '0', 'A', '-65', '2026-09-25'],
    ['GALAXY M12', '0', 'B', '-73.5', '2026-09-25'],
    ['GALAXY M23', '0', 'A', '50', '2026-09-25'],
    ['GALAXY M23', '0', 'B', '30', '2026-09-25'],
    ['GALAXY M31', '0', 'A', '-38', '2026-09-25'],
    ['GALAXY M31', '0', 'B', '-49.2', '2026-09-25'],
    ['GALAXY M32', '0', 'A', '-20', '2026-09-25'],
    ['GALAXY M32', '0', 'B', '-33', '2026-09-25'],
    ['GALAXY NOTE 9', '0', 'A', '-110', '2026-09-25'],
    ['GALAXY NOTE 9', '0', 'B', '-114', '2026-09-25'],
    ['GALAXY NOTE 10', '0', 'A', '100', '2026-09-25'],
    ['GALAXY NOTE 10', '0', 'B', '75', '2026-09-25'],
    ['GALAXY NOTE 10 LITE', '0', 'A', '50', '2026-09-25'],
    ['GALAXY NOTE 10 LITE', '0', 'B', '30', '2026-09-25'],
    ['GALAXY NOTE 10+', '0', 'A', '150', '2026-09-25'],
    ['GALAXY NOTE 10+', '0', 'B', '120', '2026-09-25'],
    ['GALAXY NOTE 20', '0', 'A', '500', '2026-09-25'],
    ['GALAXY NOTE 20', '0', 'B', '435', '2026-09-25'],
    ['GALAXY NOTE 20 ULTRA', '0', 'A', '700', '2026-09-25'],
    ['GALAXY NOTE 20 ULTRA', '0', 'B', '615', '2026-09-25'],
    ['GALAXY S7 EDGE', '0', 'A', '-110', '2026-09-25'],
    ['GALAXY S7 EDGE', '0', 'B', '-114', '2026-09-25'],
    ['GALAXY S8', '0', 'A', '-65', '2026-09-25'],
    ['GALAXY S8', '0', 'B', '-73.5', '2026-09-25'],
    ['GALAXY S8+', '0', 'A', '-65', '2026-09-25'],
    ['GALAXY S8+', '0', 'B', '-73.5', '2026-09-25'],
    ['GALAXY S9', '0', 'A', '-50', '2026-09-25'],
    ['GALAXY S9', '0', 'B', '-60', '2026-09-25'],
    ['GALAXY S9+', '64', 'A', '-50', '2026-09-25'],
    ['GALAXY S9+', '64', 'B', '-60', '2026-09-25'],
    ['GALAXY S9+', '128', 'A', '-50', '2026-09-25'],
    ['GALAXY S9+', '128', 'B', '-60', '2026-09-25'],
    ['GALAXY S10', '0', 'A', '-30', '2026-09-25'],
    ['GALAXY S10', '0', 'B', '-42', '2026-09-25'],
    ['GALAXY S10 E', '0', 'A', '-30', '2026-09-25'],
    ['GALAXY S10 E', '0', 'B', '-42', '2026-09-25'],
    ['GALAXY S10+', '0', 'A', '0', '2026-09-25'],
    ['GALAXY S10+', '0', 'B', '-15', '2026-09-25'],
    ['GALAXY S20', '0', 'A', '50', '2026-09-25'],
    ['GALAXY S20', '0', 'B', '30', '2026-09-25'],
    ['GALAXY S20 FE', '0', 'A', '50', '2026-09-25'],
    ['GALAXY S20 FE', '0', 'B', '30', '2026-09-25'],
    ['GALAXY S20 Plus', '0', 'A', '100', '2026-09-25'],
    ['GALAXY S20 Plus', '0', 'B', '75', '2026-09-25'],
    ['GALAXY S20 Ultra', '0', 'A', '250', '2026-09-25'],
    ['GALAXY S20 Ultra', '0', 'B', '210', '2026-09-25'],
    ['GALAXY S21 FE', '0', 'A', '250', '2026-09-25'],
    ['GALAXY S21 FE', '0', 'B', '210', '2026-09-25'],
    ['GALAXY S21', '0', 'A', '250', '2026-09-25'],
    ['GALAXY S21', '0', 'B', '210', '2026-09-25'],
    ['GALAXY S21 Plus', '0', 'A', '350', '2026-09-25'],
    ['GALAXY S21 Plus', '0', 'B', '300', '2026-09-25'],
    ['GALAXY S21 Ultra', '0', 'A', '450', '2026-09-25'],
    ['GALAXY S21 Ultra', '0', 'B', '390', '2026-09-25'],
    ['GALAXY S22', '0', 'A', '350', '2026-09-25'],
    ['GALAXY S22', '0', 'B', '300', '2026-09-25'],
    ['GALAXY S22 Plus', '0', 'A', '450', '2026-09-25'],
    ['GALAXY S22 Plus', '0', 'B', '390', '2026-09-25'],
    ['GALAXY S22 Ultra', '0', 'A', '850', '2026-09-25'],
    ['GALAXY S22 Ultra', '0', 'B', '750', '2026-09-25'],
    ['GALAXY S23 FE', '0', 'A', '550', '2026-09-25'],
    ['GALAXY S23 FE', '0', 'B', '480', '2026-09-25'],
    ['GALAXY S23', '0', 'A', '550', '2026-09-25'],
    ['GALAXY S23', '0', 'B', '480', '2026-09-25'],
    ['GALAXY S23 Plus', '0', 'A', '850', '2026-09-25'],
    ['GALAXY S23 Plus', '0', 'B', '750', '2026-09-25'],
    ['GALAXY S23 Ultra', '0', 'A', '1.050.000', '2026-09-25'],
    ['GALAXY S23 Ultra', '0', 'B', '930', '2026-09-25'],
    ['GALAXY S24 FE', '0', 'A', '750', '2026-09-25'],
    ['GALAXY S24 FE', '0', 'B', '660', '2026-09-25'],
    ['GALAXY S24', '0', 'A', '950', '2026-09-25'],
    ['GALAXY S24', '0', 'B', '840', '2026-09-25'],
    ['GALAXY S24 Plus', '0', 'A', '1.050.000', '2026-09-25'],
    ['GALAXY S24 Plus', '0', 'B', '930', '2026-09-25'],
    ['GALAXY S24 Ultra', '0', 'A', '1.550.000', '2026-09-25'],
    ['GALAXY S24 Ultra', '0', 'B', '1.380.000', '2026-09-25'],
    ['GALAXY Z Flip', '0', 'A', '250', '2026-09-25'],
    ['GALAXY Z Flip', '0', 'B', '210', '2026-09-25'],
    ['GALAXY Z Flip 3', '0', 'A', '450', '2026-09-25'],
    ['GALAXY Z Flip 3', '0', 'B', '390', '2026-09-25'],
    ['GALAXY Z Flip 4', '0', 'A', '650', '2026-09-25'],
    ['GALAXY Z Flip 4', '0', 'B', '570', '2026-09-25'],
    ['GALAXY Z Flip 5', '0', 'A', '850', '2026-09-25'],
    ['GALAXY Z Flip 5', '0', 'B', '750', '2026-09-25'],
    ['GALAXY Z Flip 6', '0', 'A', '1.050.000', '2026-09-25'],
    ['GALAXY Z Flip 6', '0', 'B', '930', '2026-09-25'],
    ['GALAXY Z Fold', '0', 'A', '450', '2026-09-25'],
    ['GALAXY Z Fold', '0', 'B', '390', '2026-09-25'],
    ['GALAXY Z Fold 2', '0', 'A', '550', '2026-09-25'],
    ['GALAXY Z Fold 2', '0', 'B', '480', '2026-09-25'],
    ['GALAXY Z Fold 3', '0', 'A', '650', '2026-09-25'],
    ['GALAXY Z Fold 3', '0', 'B', '570', '2026-09-25'],
    ['GALAXY Z Fold 4', '0', 'A', '750', '2026-09-25'],
    ['GALAXY Z Fold 4', '0', 'B', '660', '2026-09-25'],
    ['GALAXY Z Fold 5', '0', 'A', '850', '2026-09-25'],
    ['GALAXY Z Fold 5', '0', 'B', '750', '2026-09-25'],
    ['4', '128', 'A', '5000000', '2026-09-25T07:15:38.273Z'],
    ['4', '128', 'B', '4700000', '2026-09-25T07:15:44.471Z'],
    ];

    const refsHoja = ss.getSheetByName('ReferenciaRetoma');
    if (!refsHoja) throw new Error('No se encontró la hoja ReferenciaRetoma.');

    refsHoja.clearContents();
    refsHoja.appendRow(['id', 'modelo_id', 'capacidad', 'grado', 'precio_compra_usd', 'precio_compra_cop', 'precio_venta_usd', 'precio_venta_cop', 'actualizado_en']);

    let refsCreadas = 0;
    FILAS.forEach(function (fila, i) {
      const modelo = String(fila[0]).trim();
      const cap = String(fila[1]).trim();
      const grado = String(fila[2]).trim().toUpperCase();
      const valor = parseCOP(fila[3]);
      const fecha = fila[4] || new Date().toISOString();
      if (modelo === '' || (grado !== 'A' && grado !== 'B')) return;
      refsHoja.appendRow([i + 1, modelo, cap, grado, '', valor === null ? '' : valor, '', '', fecha]);
      refsCreadas++;
    });

    return { referencia_retoma: refsCreadas };
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
    .addItem('5. Sembrar Referencias de Retoma (catálogo inicial)', 'menuSembrarCatalogoRetoma')
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

function menuSembrarCatalogoRetoma() {
  try {
    sembrarCatalogoRetoma(SpreadsheetApp.getActiveSpreadsheet());
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const refs = getSheetData(ss, 'ReferenciaRetoma');
    SpreadsheetApp.getUi().alert('✅ Catálogo sembrado / actualizado.\nReferencias en la hoja: ' + refs.length);
  } catch (err) {
    SpreadsheetApp.getUi().alert('❌ Error: ' + err.message);
  }
}
