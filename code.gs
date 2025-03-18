function sendMonthlyFinanceReport() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Hárok2"); // Uprav podľa názvu tabuľky
  var data = sheet.getDataRange().getValues();

  var today = new Date();
  var currentMonth = today.getMonth() + 1;
  var currentYear = today.getFullYear();
  
  var reportMap = {};
  var totalSum = 0;
  var categoryColors = {
    "Zábava": "#0a53a8",
    "Škola": "#ffe5a0",
    "Jedlo": "#ffc8aa",
    "Iné potreby": "#473822",
    "Skaut": "#11734b"
  };

  var headers = data[0];
  var dateIndex = headers.indexOf("Dátum");
  var itemIndex = headers.indexOf("Položka");
  var countIndex = headers.indexOf("Počet");
  var priceIndex = headers.indexOf("Cena");
  var categoryIndex = headers.indexOf("Typ transakcie");

  if (dateIndex === -1 || itemIndex === -1 || countIndex === -1 || priceIndex === -1 || categoryIndex === -1) {
    Logger.log("Chyba: Niektoré stĺpce neboli nájdené.");
    return;
  }

  for (var i = 1; i < data.length; i++) {
    if (!data[i][dateIndex]) continue;

    var date;
    try {
      date = new Date(data[i][dateIndex]);
    } catch (e) {
      Logger.log("Chyba konverzie dátumu v riadku " + (i+1));
      continue;
    }

    var month = date.getMonth() + 1;
    var year = date.getFullYear();

    if (month === currentMonth && year === currentYear) {
      var category = data[i][categoryIndex] ? data[i][categoryIndex].toString().trim() : "Nezaradené";
      var item = data[i][itemIndex] ? data[i][itemIndex].toString().trim() : "Neznáma položka";
      var count = parseInt(data[i][countIndex], 10) || 1;
      var priceRaw = data[i][priceIndex] ? data[i][priceIndex].toString() : "0";
      var price = parseFloat(priceRaw.replace(",", ".").replace("€", "").trim()) || 0;

      if (!reportMap[category]) {
        reportMap[category] = {};
      }

      if (!reportMap[category][item]) {
        reportMap[category][item] = { count: 0, totalPrice: 0 };
      }

      reportMap[category][item].count += count;
      reportMap[category][item].totalPrice += price;
      totalSum += price;
    }
  }

  // Pridanie všetkých kategórií, aj keď nemajú položky
  var allCategories = ["Zábava", "Škola", "Jedlo", "Iné potreby", "Skaut"];
  allCategories.forEach(function(category) {
    if (!reportMap[category]) {
      reportMap[category] = {}; // Ak kategória nemá položky, pridáme ju s prázdnym objektom
    }
  });

  if (Object.keys(reportMap).length === 0) {
    Logger.log("Žiadne nákupy tento mesiac.");
    return;
  }

  // Vytvorenie Google Docs dokumentu
  var doc = DocumentApp.create(`Finančný report ${currentMonth}-${currentYear}`);
  var body = doc.getBody();
  
  // Hlavička
  body.appendParagraph("Finančný report")
      .setHeading(DocumentApp.ParagraphHeading.HEADING1)
      .setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  
  body.appendParagraph(`Mesiac: ${currentMonth}/${currentYear}\n`).setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  
  var categoriesForChart = [];
  var chartData = [];

  // Pridanie tabuľky pre kategórie s položkami a grafu
  for (var category in reportMap) {
    if (Object.keys(reportMap[category]).length === 0) {
      body.appendParagraph(`V mesiaci ${currentMonth}/${currentYear} neboli žiadne nákupy v kategórii: ${category}.`)
          .setAlignment(DocumentApp.HorizontalAlignment.LEFT);
    } else {
      body.appendParagraph(category)
          .setHeading(DocumentApp.ParagraphHeading.HEADING2)
          .setBold(true);

      var tableData = [["Položka", "Počet", "Cena (€)"]];
      for (var item in reportMap[category]) {
        tableData.push([
          item,
          reportMap[category][item].count.toString(),
          reportMap[category][item].totalPrice.toFixed(2)
        ]);
      }

      var table = body.appendTable(tableData);
      table.getRow(0).editAsText().setBold(true); // Hlavička tabuľky

      categoriesForChart.push(category);
      chartData.push([category, totalSum]);  // Skutočný výpočet celkových výdavkov podľa kategórie
    }
  }

  // Celková suma
  body.appendParagraph(`Celková suma: ${totalSum.toFixed(2)} €`)
      .setBold(true)
      .setAlignment(DocumentApp.HorizontalAlignment.RIGHT);

  doc.saveAndClose();

  // Príprava dát pre koláčový graf
  var pieChartData = [];
  for (var category in reportMap) {
    var categoryTotal = 0;
    for (var item in reportMap[category]) {
      categoryTotal += reportMap[category][item].totalPrice;
    }
    pieChartData.push([category, categoryTotal]);
  }

  // Vytvorenie koláčového grafu v Google Sheets
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var chart = sheet.newChart()
    .setChartType(Charts.ChartType.PIE)
    .addRange(sheet.getRange('A1:B' + pieChartData.length)) // Za predpokladu, že máme dáta v stĺpcoch A a B
    .setPosition(2, 2, 0, 0)
    .setOption('title', 'Distribúcia výdavkov podľa kategórií')
    .setOption('colors', [categoryColors["Zábava"], categoryColors["Škola"], categoryColors["Jedlo"], categoryColors["Iné potreby"], categoryColors["Skaut"]])
    .build();

  sheet.insertChart(chart);

  // Konvertovanie dokumentu do PDF
  var pdfFile = DriveApp.getFileById(doc.getId()).getAs(MimeType.PDF);
  var folder = DriveApp.getRootFolder();
  var pdf = folder.createFile(pdfFile.setName(`Finančný report ${currentMonth}-${currentYear}.pdf`));
  // Datovanie reportu
  var startedDate = new Date;
  var formattedStartedDate = Utilities.formatDate(startedDate, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");

  // Odoslanie e-mailu s prílohou
  var emailBody = `Finančný report za ${currentMonth}/${currentYear}.\n Report bol vyhotovený ${formattedStartedDate}`;
  MailApp.sendEmail({
    to: "email@email.com",
    subject: `Mesačný finančný report - ${currentMonth}/${currentYear}`,
    body: emailBody,
    attachments: [pdf.getAs(MimeType.PDF)]  // Posielame PDF ako prílohu
  });

  Logger.log("Mesačný report bol odoslaný s PDF prílohou.");
  Logger.log(`Report bol vyhotovený ${formattedStartedDate}`)
}
