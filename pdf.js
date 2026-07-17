const TEMPLATE_COTIZACION_ID = "1QKyfB1bIu_rPIpZFVr3R4CsFDmAXnlqfVxuhywZAYIY";
const PDF_FOLDER_NAME = "PDF";

function getPDFFolder() {

  var root = DriveApp.getFoldersByName("OPTIQRO Business Suite");

  if (!root.hasNext()) {
    throw new Error("No existe la carpeta OPTIQRO Business Suite");
  }

  var carpetaPrincipal = root.next();

  var carpetas = carpetaPrincipal.getFoldersByName(PDF_FOLDER_NAME);

  if (carpetas.hasNext()) {
    return carpetas.next();
  }

  return carpetaPrincipal.createFolder(PDF_FOLDER_NAME);
}

function generarPDFPrueba() {

  var copia = DriveApp
      .getFileById(TEMPLATE_COTIZACION_ID)
      .makeCopy("COTIZACION_PRUEBA", getPDFFolder());

  var doc = DocumentApp.openById(copia.getId());
  var body = doc.getBody();

  body.replaceText("{{FOLIO}}","COT-000001");
  body.replaceText("{{FECHA}}",Utilities.formatDate(new Date(),Session.getScriptTimeZone(),"dd/MM/yyyy"));
  body.replaceText("{{CLIENTE}}","CLIENTE DEMO SA DE CV");
  body.replaceText("{{RFC}}","XAXX010101000");
  body.replaceText("{{CONTACTO}}","Juan Pérez");
  body.replaceText("{{TELEFONO}}","4421234567");
  body.replaceText("{{DIRECCION}}","Querétaro, Qro.");

  body.replaceText("{{TABLA_PRODUCTOS}}",
    "2 Banda Transportadora\n" +
    "1 Rodamiento SKF\n" +
    "5 Grapas Flexco");

  body.replaceText("{{SUBTOTAL}}","$12,500.00");
  body.replaceText("{{IVA}}","$2,000.00");
  body.replaceText("{{TOTAL}}","$14,500.00");

  body.replaceText("{{OBSERVACIONES}}","Cotización de prueba.");
  body.replaceText("{{ENTREGA}}","5 días hábiles");
  body.replaceText("{{PAGO}}","Contado");
  body.replaceText("{{VIGENCIA}}","15 días");

  doc.saveAndClose();

  var pdf = copia.getAs(MimeType.PDF).setName("COT-000001.pdf");
  getPDFFolder().createFile(pdf);

  return "PDF generado correctamente.";
}

function generarPDFCotizacion(data, folio) {

  var copia = DriveApp
      .getFileById(TEMPLATE_COTIZACION_ID)
      .makeCopy("COT_" + folio, getPDFFolder());

  var doc = DocumentApp.openById(copia.getId());
  var body = doc.getBody();

  body.replaceText("{{FOLIO}}", folio);

  body.replaceText(
    "{{FECHA}}",
    Utilities.formatDate(
      new Date(),
      Session.getScriptTimeZone(),
      "dd/MM/yyyy"
    )
  );

  body.replaceText("{{CLIENTE}}", data.cliente_nombre || "");
  body.replaceText("{{RFC}}", data.cliente_rfc || "");
  body.replaceText("{{CONTACTO}}", data.cliente_contacto || "");
  body.replaceText("{{TELEFONO}}", data.cliente_tel || "");
  body.replaceText("{{DIRECCION}}", data.cliente_dir || "");

  var tabla = "";

  (data.items || []).forEach(function(i) {

    var subtotal = Number(i.precio) * Number(i.qty);

    tabla +=
      i.qty + "    " +
      i.nombre +
      "    $" + Number(i.precio).toFixed(2) +
      "    $" + subtotal.toFixed(2) +
      "\n";

  });

  body.replaceText("{{TABLA_PRODUCTOS}}", tabla);

  body.replaceText("{{SUBTOTAL}}", "$" + Number(data.subtotal || 0).toFixed(2));
  body.replaceText("{{IVA}}", "$" + Number(data.iva_amt || 0).toFixed(2));
  body.replaceText("{{TOTAL}}", "$" + Number(data.total || 0).toFixed(2));

  body.replaceText("{{OBSERVACIONES}}", data.notas || "");
  body.replaceText("{{ENTREGA}}", "5 días hábiles");
  body.replaceText("{{PAGO}}", "Contado");
  body.replaceText("{{VIGENCIA}}", "15 días");

  doc.saveAndClose();

  var pdf = copia.getAs(MimeType.PDF);
  pdf.setName("Cotizacion_OPTIQRO_" + folio + ".pdf");
DriveApp.getFileById(copia.getId()).setTrashed(true);
  return pdf;
}
