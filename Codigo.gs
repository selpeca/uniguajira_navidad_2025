function enviarCorreosNavidad() {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName("Base de datos");
  const data = sheet.getDataRange().getValues();

  // Quitar encabezados
  data.shift();

  data.forEach((row, index) => {
    const comunidad = row[0];
    const nombreNino = row[1];
    const edad = row[2];
    const sexo = row[3];
    const nombreFuncionario = row[4];
    const correoFuncionario = row[6];

    // Validación mínima
    if (!correoFuncionario) return;

    const template = HtmlService.createTemplateFromFile("template.html");

    template.nombreFuncionario = nombreFuncionario;
    template.nombreNino = nombreNino;
    template.edad = edad;
    template.sexo = sexo;
    template.comunidad = comunidad;

    const mensaje = template.evaluate().getContent();
    try {
      GmailApp.sendEmail(
        correoFuncionario,
        "Únete a la campaña: Desde Uniguajira, una Navidad para Compartir y Sonreír",
        "",
        {
          htmlBody: mensaje,
          name: "Dirección de Extensión y Proyección Social",
        }
      );

      // Solo si no falla:
      sheet.getRange(index + 2, 10).setValue(new Date());
    } catch (error) {
      sheet.getRange(index + 2, 11).setValue("ERROR: " + error);
    }
  });
}
