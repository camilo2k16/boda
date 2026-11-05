/**
 * Confirmaciones · Boda de Camilo y Carolina
 *
 * 1. Ejecuta la función "configurar" una sola vez (crea las hojas Resumen, Invitados y Confirmaciones).
 * 2. Implementar > Nueva implementación > Aplicación web
 *      Ejecutar como: Yo   ·   Quién tiene acceso: Cualquier usuario
 * 3. Copia la URL que termina en /exec y pégala en config.js (en GitHub).
 *
 * Cada invitado queda en UNA sola fila: si vuelve a confirmar desde su link, se actualiza su respuesta.
 */

// Código del link · Nombre en la invitación · Lado · Cupo máximo · (padrinos)
var INVITADOS = [
  ['byron-oriana', 'Byron y Oriana', 'Novio', 2, true],
  ['carlos', 'Carlos', 'Novio', 1],
  ['laura', 'Laura', 'Novio', 1],
  ['daniel-valencia', 'Daniel Valencia', 'Novio', 1],
  ['nicolas-ospina', 'Nicolás Ospina', 'Novio', 1],
  ['karla-puerto', 'Karla Puerto', 'Novio', 2],
  ['angel', 'Ángel', 'Novio', 1],
  ['fredy', 'Fredy', 'Novio', 3],
  ['dayana', 'Dayana', 'Novio', 1],
  ['angie', 'Angie', 'Novio', 1],
  ['ingrid', 'Ingrid', 'Novio', 1],
  ['esly', 'Esly', 'Novio', 1],
  ['estefany', 'Estefany', 'Novio', 1],
  ['heidy', 'Heidy', 'Novio', 1],
  ['jeldin', 'Jeldin', 'Novio', 1],
  ['karina', 'Karina', 'Novio', 2],
  ['vanessa', 'Vanessa', 'Novio', 2],
  ['marisol', 'Marisol', 'Novia', 1],
  ['alexandra', 'Alexandra', 'Novia', 2],
  ['monica', 'Mónica', 'Novia', 4],
  ['leidy', 'Leidy', 'Novia', 1],
  ['maria-monica', 'María Mónica', 'Novia', 1],
  ['mariana-julian', 'Mariana y Julián', 'Novia', 2],
  ['camilo', 'Camilo', 'Novia', 3],
  ['heber', 'Heber', 'Novia', 3],
  ['juliana', 'Juliana', 'Novia', 1],
  ['johana', 'Johana', 'Novia', 1],
  ['kathe', 'Kathe', 'Novia', 2],
  ['francy', 'Francy', 'Novia', 4],
  ['papa-william', 'Papá William', 'Novia', 1],
  ['lorena-urena', 'Lorena Ureña', 'Novia', 4],
  ['primo-juan-carlos', 'Primo Juan Carlos', 'Novia', 2]
];

function buscarInvitado_(codigo) {
  for (var i = 0; i < INVITADOS.length; i++) if (INVITADOS[i][0] === codigo) return INVITADOS[i];
  return null;
}

function configurar() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var conf = ss.getSheetByName('Confirmaciones') || ss.insertSheet('Confirmaciones');
  if (conf.getLastRow() === 0) {
    conf.appendRow(['Fecha', 'Código', 'Invitación', 'Lado', 'Cupo', 'Nombre escrito', '¿Asiste?', 'Personas', '¿Va a la cena?']);
  }
  conf.setFrozenRows(1);
  conf.getRange(1, 1, 1, 9).setFontWeight('bold').setBackground('#F4D0DC');
  conf.getRange('A:A').setNumberFormat('dd/mm/yyyy hh:mm');

  var res = ss.getSheetByName('Resumen') || ss.insertSheet('Resumen', 0);
  res.clear();
  res.getRange('A1').setValue('Boda de Camilo y Carolina · Resumen').setFontWeight('bold').setFontSize(14);
  res.getRange('A2:B2').setValues([['Link de la invitación (GitHub Pages)', 'https://TU-USUARIO.github.io/boda/']]);
  res.getRange('B2').setBackground('#FFF3D6');
  res.getRange('A4:B10').setValues([
    ['Invitaciones', '=COUNTA(Invitados!A2:A)'],
    ['Cupos totales', '=SUM(Invitados!D2:D)'],
    ['Invitaciones que respondieron', '=COUNTIF(Invitados!F2:F,"<>Pendiente")'],
    ['Invitaciones sin responder', '=COUNTIF(Invitados!F2:F,"Pendiente")'],
    ['Personas que asistirán', '=SUM(Confirmaciones!H2:H)'],
    ['Personas que van a la cena', '=SUMIFS(Confirmaciones!H2:H,Confirmaciones!I2:I,"Sí")'],
    ['Invitaciones que no asistirán', '=COUNTIF(Invitados!F2:F,"No asiste")']
  ]);
  res.getRange('A4:A10').setFontWeight('bold');
  res.setColumnWidth(1, 260); res.setColumnWidth(2, 320);

  var inv = ss.getSheetByName('Invitados') || ss.insertSheet('Invitados', 1);
  inv.clear();
  var filas = [['Código', 'Invitación', 'Lado', 'Cupo', 'Link personal', 'Estado', 'Personas', 'Cena', 'Mensaje para WhatsApp']];
  INVITADOS.forEach(function (x, i) {
    var f = i + 2, plural = x[3] > 1;
    var saludo = x[4]
      ? 'Queridos padrinos " & B' + f + ' & ", tenemos algo muy especial para ustedes. Abran su invitación aquí: '
      : 'Hola " & B' + f + ' & ", con mucho cariño ' + (plural ? 'los' : 'te') + ' invitamos a nuestra boda. Aquí está ' + (plural ? 'su' : 'tu') + ' invitación: ';
    filas.push([x[0], x[1], x[2], x[3],
      '=Resumen!$B$2 & "?i=" & A' + f,
      '=IFERROR(IF(VLOOKUP(A' + f + ',Confirmaciones!B:G,6,FALSE)="Sí","Confirmó","No asiste"),"Pendiente")',
      '=IFERROR(VLOOKUP(A' + f + ',Confirmaciones!B:H,7,FALSE),"")',
      '=IFERROR(VLOOKUP(A' + f + ',Confirmaciones!B:I,8,FALSE),"")',
      '="' + saludo + '" & E' + f]);
  });
  inv.getRange(1, 1, filas.length, filas[0].length).setValues(filas);
  inv.setFrozenRows(1);
  inv.getRange(1, 1, 1, 9).setFontWeight('bold').setBackground('#F4D0DC');
  inv.setColumnWidth(5, 330); inv.setColumnWidth(9, 520);
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var p = (e && e.parameter) || {};
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var hoja = ss.getSheetByName('Confirmaciones');
    if (!hoja) { configurar(); hoja = ss.getSheetByName('Confirmaciones'); }

    var codigo = String(p.codigo || '').slice(0, 40);
    var x = buscarInvitado_(codigo);
    var cupo = x ? x[3] : Math.max(1, Math.min(Number(p.cupo) || 1, 10));
    var asiste = p.asiste === 'si';
    var personas = asiste ? Math.max(1, Math.min(Number(p.personas) || 1, cupo)) : 0;
    var nombre = String(p.nombre || '').slice(0, 80);
    if (/^[=+\-@]/.test(nombre)) nombre = "'" + nombre;   // evita fórmulas escritas en el formulario

    var fila = [new Date(), codigo || '(sin código)', x ? x[1] : '', x ? x[2] : '', cupo, nombre,
                asiste ? 'Sí' : 'No', personas, asiste && p.cena === 'si' ? 'Sí' : 'No'];

    var n = hoja.getLastRow(), destino = -1;
    if (codigo && n > 1) {
      var cods = hoja.getRange(2, 2, n - 1, 1).getValues();
      for (var k = 0; k < cods.length; k++) if (cods[k][0] === codigo) { destino = k + 2; break; }
    }
    if (destino > 0) hoja.getRange(destino, 1, 1, fila.length).setValues([fila]);
    else hoja.appendRow(fila);

    return ContentService.createTextOutput('{"ok":true}').setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

function doGet() {
  return ContentService.createTextOutput('Confirmaciones de la boda de Camilo y Carolina: activo.');
}
