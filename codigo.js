var SPREADSHEET_ID = "1ZfLXV_SRGur_9YcPEDJJT93xv46AREvIKC7NJm2DKUw";
var HOJA_PRODUCTOS   = "Productos";
var HOJA_COTIZ       = "Cotizaciones";
var HOJA_CLIENTES    = "Clientes";
var HOJA_PROVEEDORES = "Proveedores";
var HOJA_OC          = "OrdenesCompra";
function doGet(e) {
  return HtmlService.createTemplateFromFile("index")
    .evaluate()
    .setTitle("OPTIQRO · Cotizador")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag("viewport", "width=device-width, initial-scale=1");
}
function getSS() { return SpreadsheetApp.openById(SPREADSHEET_ID); }
// ══════════════════ PRODUCTOS ══════════════════
function getProductos() {
  try {
    var sh = getSS().getSheetByName(HOJA_PRODUCTOS);
    if (!sh) return JSON.stringify({ok:false, msg:"Hoja no encontrada"});
    var rows = sh.getDataRange().getValues(), out = [];
    for (var i = 1; i < rows.length; i++) {
      var r = rows[i];
      if (r[0] && r[1]) out.push({codigo:String(r[0]), nombre:String(r[1]), categoria:String(r[2]||"OTROS"), precio:Number(r[3])||0, fila:i+1});
    }
    return JSON.stringify({ok:true, data:out});
  } catch(e) { return JSON.stringify({ok:false, msg:e.message}); }
}
function agregarProducto(d) {
  try { var p=JSON.parse(d); getSS().getSheetByName(HOJA_PRODUCTOS).appendRow([p.codigo,p.nombre,p.categoria,p.precio]); return JSON.stringify({ok:true}); }
  catch(e) { return JSON.stringify({ok:false, msg:e.message}); }
}
function editarProducto(d) {
  try { var p=JSON.parse(d); getSS().getSheetByName(HOJA_PRODUCTOS).getRange(p.fila,1,1,4).setValues([[p.codigo,p.nombre,p.categoria,p.precio]]); return JSON.stringify({ok:true}); }
  catch(e) { return JSON.stringify({ok:false, msg:e.message}); }
}
function eliminarProducto(fila) {
  try { getSS().getSheetByName(HOJA_PRODUCTOS).deleteRow(fila); return JSON.stringify({ok:true}); }
  catch(e) { return JSON.stringify({ok:false, msg:e.message}); }
}

// ══════════════════ CLIENTES ══════════════════
function getClientes() {
  try {
    var sh = getSS().getSheetByName(HOJA_CLIENTES);
    if (!sh) return JSON.stringify({ok:false, msg:"Hoja no encontrada"});
    var rows = sh.getDataRange().getValues(), out = [];
    for (var i = 1; i < rows.length; i++) {
      var r = rows[i];
      if (r[0]) out.push({nombre:String(r[0]), rfc:String(r[1]||""), contacto:String(r[2]||""), tel:String(r[3]||""), dir:String(r[4]||""), fila:i+1});
    }
    return JSON.stringify({ok:true, data:out});
  } catch(e) { return JSON.stringify({ok:false, msg:e.message}); }
}
function agregarCliente(d) {
  try { var c=JSON.parse(d); getSS().getSheetByName(HOJA_CLIENTES).appendRow([c.nombre,c.rfc,c.contacto,c.tel,c.dir]); return JSON.stringify({ok:true}); }
  catch(e) { return JSON.stringify({ok:false, msg:e.message}); }
}
function editarCliente(d) {
  try { var c=JSON.parse(d); getSS().getSheetByName(HOJA_CLIENTES).getRange(c.fila,1,1,5).setValues([[c.nombre,c.rfc,c.contacto,c.tel,c.dir]]); return JSON.stringify({ok:true}); }
  catch(e) { return JSON.stringify({ok:false, msg:e.message}); }
}
function eliminarCliente(fila) {
  try { getSS().getSheetByName(HOJA_CLIENTES).deleteRow(fila); return JSON.stringify({ok:true}); }
  catch(e) { return JSON.stringify({ok:false, msg:e.message}); }
}

// ══════════════════ PROVEEDORES ══════════════════
function getProveedores() {
  try {
    var ss = getSS();
    var sh = ss.getSheetByName(HOJA_PROVEEDORES);
    if (!sh) {
      sh = ss.insertSheet(HOJA_PROVEEDORES);
      sh.appendRow(["NOMBRE","RFC","CONTACTO","TELÉFONO/CORREO","DIRECCIÓN"]);
    }
    var rows = sh.getDataRange().getValues(), out = [];
    for (var i = 1; i < rows.length; i++) {
      var r = rows[i];
      if (r[0]) out.push({nombre:String(r[0]), rfc:String(r[1]||""), contacto:String(r[2]||""), tel:String(r[3]||""), dir:String(r[4]||""), fila:i+1});
    }
    return JSON.stringify({ok:true, data:out});
  } catch(e) { return JSON.stringify({ok:false, msg:e.message}); }
}
function agregarProveedor(d) {
  try { var p=JSON.parse(d); getSS().getSheetByName(HOJA_PROVEEDORES).appendRow([p.nombre,p.rfc,p.contacto,p.tel,p.dir]); return JSON.stringify({ok:true}); }
  catch(e) { return JSON.stringify({ok:false, msg:e.message}); }
}
function editarProveedor(d) {
  try { var p=JSON.parse(d); getSS().getSheetByName(HOJA_PROVEEDORES).getRange(p.fila,1,1,5).setValues([[p.nombre,p.rfc,p.contacto,p.tel,p.dir]]); return JSON.stringify({ok:true}); }
  catch(e) { return JSON.stringify({ok:false, msg:e.message}); }
}
function eliminarProveedor(fila) {
  try { getSS().getSheetByName(HOJA_PROVEEDORES).deleteRow(fila); return JSON.stringify({ok:true}); }
  catch(e) { return JSON.stringify({ok:false, msg:e.message}); }
}

// ══════════════════ COTIZACIONES ══════════════════
function getCotizaciones() {
  try {
    var sh = getSS().getSheetByName(HOJA_COTIZ);
    var rows = sh.getDataRange().getValues(), out = [];
    for (var i = 1; i < rows.length; i++) {
      out.push({folio:rows[i][0], fecha:rows[i][1], cliente:rows[i][2], total:rows[i][11], estatus:rows[i][12]});
    }
    return JSON.stringify({ok:true, data:out.reverse()});
  } catch(e) { return JSON.stringify({ok:false, msg:e.message}); }
}
function guardarCotizacion(d) {
  try {
    var data = JSON.parse(d);
    var sh = getSS().getSheetByName(HOJA_COTIZ);
    var folio = "OPT-" + String(Math.max(sh.getLastRow(),1)).padStart(4,"0");
    sh.appendRow([folio, new Date(), data.cliente_nombre, data.cliente_rfc,
      data.cliente_contacto, data.cliente_tel, data.cliente_dir,
      data.items.map(function(i){return i.nombre+" x"+i.qty;}).join(" | "),
      data.subtotal, data.iva_amt||0, data.cargo_amt||0, data.total, "EMITIDA"]);
    var pdf = generarPDFCotizacion(data, folio);
    var archivo = getPDFFolder().createFile(pdf);
    archivo.setName("Cotizacion_OPTIQRO_"+folio+".pdf");
    archivo.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return JSON.stringify({ok:true, folio:folio, url:archivo.getUrl()});
  } catch(e) { return JSON.stringify({ok:false, msg:e.message}); }
}

// ══════════════════ ÓRDENES DE COMPRA ══════════════════
function getOrdenesCompra() {
  try {
    var ss = getSS();
    var sh = ss.getSheetByName(HOJA_OC);
    if (!sh) {
      sh = ss.insertSheet(HOJA_OC);
      sh.appendRow(["FOLIO","FECHA","PROVEEDOR","RFC","CONTACTO","TELÉFONO","DIRECCIÓN","ARTÍCULOS","SUBTOTAL","IVA","CARGO","TOTAL","ESTATUS"]);
    }
    var rows = sh.getDataRange().getValues(), out = [];
    for (var i = 1; i < rows.length; i++) {
      out.push({folio:rows[i][0], fecha:rows[i][1], proveedor:rows[i][2], total:rows[i][11], estatus:rows[i][12]});
    }
    return JSON.stringify({ok:true, data:out.reverse()});
  } catch(e) { return JSON.stringify({ok:false, msg:e.message}); }
}
function guardarOrdenCompra(d) {
  try {
    var data = JSON.parse(d);
    var ss = getSS();
    var sh = ss.getSheetByName(HOJA_OC);
    if (!sh) {
      sh = ss.insertSheet(HOJA_OC);
      sh.appendRow(["FOLIO","FECHA","PROVEEDOR","RFC","CONTACTO","TELÉFONO","DIRECCIÓN","ARTÍCULOS","SUBTOTAL","IVA","CARGO","TOTAL","ESTATUS"]);
    }
    var folio = "OC-" + String(Math.max(sh.getLastRow(),1)).padStart(4,"0");
    sh.appendRow([folio, new Date(), data.cliente_nombre, data.cliente_rfc,
      data.cliente_contacto, data.cliente_tel, data.cliente_dir,
      data.items.map(function(i){return i.nombre+" x"+i.qty;}).join(" | "),
      data.subtotal, data.iva_amt||0, data.cargo_amt||0, data.total, "EMITIDA"]);
    var pdf = generarPDF(data, folio, "ORDEN DE COMPRA", "PROVEEDOR");
    var archivo = getPDFFolder().createFile(pdf);
    archivo.setName("OrdenCompra_OPTIQRO_"+folio+".pdf");
    archivo.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return JSON.stringify({ok:true, folio:folio, url:archivo.getUrl()});
  } catch(e) { return JSON.stringify({ok:false, msg:e.message}); }
}

// ══════════════════ GENERADOR DE PDF (compartido) ══════════════════
function generarPDF(data, folio, tituloDoc, etiquetaContacto) {
  var hoy = new Date(), vig = new Date(hoy);
  vig.setDate(vig.getDate()+15);
  var fd = function(d){ return Utilities.formatDate(d, Session.getScriptTimeZone(), "dd/MM/yyyy"); };
  var filas = "", sub = 0;
  data.items.forEach(function(item){
    var s = item.precio * item.qty; sub += s;
    filas += "<tr><td>"+item.nombre+"</td>"
      +"<td style='text-align:center'>"+item.qty+"</td>"
      +"<td style='text-align:right'>$"+Number(item.precio).toFixed(2)+"</td>"
      +"<td style='text-align:right'>$"+s.toFixed(2)+"</td></tr>";
  });
  var iva = data.iva ? sub*0.16 : 0;
  var cargo = data.urgente ? sub*0.10 : 0;
  var total = sub + iva + cargo;
  var html = "<!DOCTYPE html><html><head><meta charset='utf-8'>"
    +"<style>"
    +"body{font-family:Arial,sans-serif;font-size:12px;margin:0;color:#111}"
    +".hdr{background:#111;color:#fff;padding:16px 24px;display:flex;justify-content:space-between}"
    +".hdr h1{margin:0;font-size:20px}"
    +".yline{height:3px;background:#F5C518}"
    +".sec{margin:14px 24px 0}"
    +"table{width:100%;border-collapse:collapse;font-size:11px;margin-top:6px}"
    +"th{background:#111;color:#fff;padding:6px 8px;text-align:left;font-size:9px}"
    +"td{padding:6px 8px;border-bottom:.3px solid #eee}"
    +"tr:nth-child(even) td{background:#F9F9F9}"
    +".tot{margin:10px 24px 0;display:flex;justify-content:flex-end}"
    +".tot table{width:auto}"
    +".tot td{padding:3px 8px}"
    +".grand td{font-weight:700;font-size:14px;background:#F5C518;padding:8px}"
    +".footer{background:#111;color:#888;text-align:center;font-size:8px;padding:8px;margin-top:16px}"
    +".firmas{display:flex;justify-content:space-around;margin:24px 24px 0;font-size:9px;color:#888;text-align:center}"
    +".fline{border-top:1px solid #888;width:170px;margin:0 auto 5px}"
    +"</style></head><body>"
    +"<div class='hdr'>"
    +"<div><h1 style='margin:0'>OPTIQRO</h1><div style='color:#666;font-size:10px'>Soluciones Integrales para la Industria</div></div>"
    +"<div><div style='color:#F5C518;font-weight:700'>"+tituloDoc+" "+folio+"</div>"
    +"<div style='color:#ccc;font-size:10px'>"+fd(hoy)+"&nbsp; Vigencia: "+fd(vig)+"</div></div>"
    +"</div>"
    +"<div class='yline'></div>"
    +"<div class='sec'><table>"
    +"<tr><th>"+etiquetaContacto+"</th><th>RFC</th><th>CONTACTO</th><th>TEL / CORREO</th></tr>"
    +"<tr><td><strong>"+(data.cliente_nombre||"—")+"</strong></td>"
    +"<td>"+(data.cliente_rfc||"—")+"</td>"
    +"<td>"+(data.cliente_contacto||"—")+"</td>"
    +"<td>"+(data.cliente_tel||"—")+"</td></tr>"
    +"</table></div>"
    +"<div class='sec'><table>"
    +"<thead><tr>"
    +"<th style='width:55%'>DESCRIPCION</th>"
    +"<th>CANT.</th>"
    +"<th style='text-align:right'>PRECIO UNIT.</th>"
    +"<th style='text-align:right'>SUBTOTAL</th>"
    +"</tr></thead><tbody>"+filas+"</tbody></table></div>"
    +"<div class='tot'><table>"
    +"<tr><td style='text-align:right;color:#555'>Subtotal:</td>"
    +"<td style='text-align:right'>$"+sub.toFixed(2)+"</td></tr>"
    +(data.iva?"<tr><td style='text-align:right;color:#555'>IVA 16%:</td><td style='text-align:right'>$"+iva.toFixed(2)+"</td></tr>":"")
    +(data.urgente?"<tr><td style='text-align:right;color:#555'>Cargo urgente:</td><td style='text-align:right'>$"+cargo.toFixed(2)+"</td></tr>":"")
    +"<tr class='grand'><td>TOTAL:</td><td style='text-align:right'>$"+total.toFixed(2)+"</td></tr>"
    +"</table></div>"
    +(data.notas?"<div class='sec'><p style='font-size:10px;color:#555;font-style:italic'><strong>Notas:</strong> "+data.notas+"</p></div>":"")
    +"<div class='firmas'>"
    +"<div><div class='fline'></div>Autorizado por OPTIQRO</div>"
    +"<div><div class='fline'></div>Recibido / Aceptado</div>"
    +"</div>"
    +"<div class='footer'>OPTIQRO &middot; Papelería &middot; Herramientas &middot; EPP &middot; Limpieza &middot; Refacciones Industriales</div>"
    +"</body></html>";
  return Utilities.newBlob(html, "text/html", "t.html")
    .getAs("application/pdf")
    .setName(tituloDoc+"_"+folio+".pdf");
}
