function onOpen() {
  var ui = SpreadsheetApp.getUi();
  // Or DocumentApp or FormApp.
  ui.createMenu('Ticketing')
      .addItem('Send Ticket', 'sendTicket')
      .addToUi();
 }

var TIKET_TERKIRIM = "TIKET_TERKIRIM";
var EMAIL_SUBJECT = "Tiket GFD 2018";

function sendTicket() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var dataRange = sheet.getRange(2, 1, sheet.getLastRow()-1, 11);
  // Fetch values for each row in the Range.
  var data = dataRange.getValues();

  for (var i = 0; i < data.length; ++i) {
    var tiketSaintek = Number(sheet.getRange(1, 13).getValues());
    var tiketSoshum = Number(sheet.getRange(1, 14).getValues());
    if (tiketSaintek + tiketSoshum > 1175) {
      break;
    };

    var TicketSend = data[i][9];
    var statusPembayaran = data[i][8];

    if ((TicketSend != TIKET_TERKIRIM)&&(statusPembayaran=='LUNAS')) {  // Prevents sending duplicates
      var row = data[i];
      if (row[4] == "SAINTEK") {
        if (tiketSaintek > 800) { continue; }
      } else {
        if (tiketSoshum > 400) { continue; }
      }

      var recipient = row[3];
      var nomorPeserta;
      if (row[10] == "") {
        nomorPeserta = createUniqNumber(sheet,row[4]);
      } else {
        nomorPeserta = row[10];
      }
      Logger.log(nomorPeserta);
      row[10] = nomorPeserta;

      MailApp.sendEmail(recipient,
                        EMAIL_SUBJECT,
                        createEmailBody(nomorPeserta),
                       {attachments: [createTicket(row)]});
      sheet.getRange(2 + i, 10).setValue(TIKET_TERKIRIM);
      sheet.getRange(2 + i, 11).setValue("'"+nomorPeserta.toString());
      SpreadsheetApp.flush();
    };
  }
}

// buat nomor unik
function createUniqNumber(sheet,paket) {
  var pkt = {"SAINTEK":13, "SOSHUM":14};
  var nomor = Number(pkt[paket]);

  var noUrut = Number(sheet.getRange(1, nomor).getValues())+1;
  sheet.getRange(1, nomor).setValue(noUrut.toString());
  var noUrutString = noUrut.toString();
  for (var i=1;i<=(5-((noUrut).toString()).length);i++) {
    noUrutString = "0" + noUrutString;
  };

  var noPes = ("0040" + (nomor-12).toString() + noUrutString).toString();

  return (noPes.toString());
}


var EMAIL_BODY_ID = "1wY_S1FZRc6n8DazoFpNrXN8CL1Vcbtxv0W0c0qhgymA";
// buat badan email
function createEmailBody(nomorPeserta) {

  var copyFile = DriveApp.getFileById(EMAIL_BODY_ID).makeCopy(),
      copyId = copyFile.getId(),
      copyDoc = DocumentApp.openById(copyId),
      copyBody = copyDoc.getActiveSection();

  copyBody.replaceText("%noPes%", nomorPeserta);
  var txt = copyBody.getText();
  DriveApp.getRootFolder().removeFile(copyFile);
  copyFile.setTrashed(true);

  return (txt);
}

function createTicket(row) {

  // Set up the slide and the spreadsheet access
  var TEMPLATE_ID = "1cFCO-3NcqgQNoX7-79723WVGaPxJeZjCpW4XnRb-SnA";

  var copyFile = DriveApp.getFileById(TEMPLATE_ID).makeCopy(),
      copyId = copyFile.getId();
  var slide = Slides.Presentations.get(copyId);
  var ElementObjectIds = {1:slide.slides[0].pageElements[1].objectId,
                          2:slide.slides[0].pageElements[2].objectId,
                          3:slide.slides[0].pageElements[3].objectId};

  var replacementText = {1:row[1],2:row[10],3:row[2]};
  for (var i = 1;i<=3; i++) {
    var requests = [{
      "deleteText": {
        "objectId": ElementObjectIds[i],
        "textRange": {
          "type": 'ALL'
        }
      }
    }, {
      "insertText": {
        "objectId": ElementObjectIds[i],
        "insertionIndex": 0,
        "text": replacementText[i]
      }
    }];

  // Execute the requests.
    Slides.Presentations.batchUpdate({'requests': requests}, copyId);
  };

  var newFile = DriveApp.createFile(copyFile.getAs('application/pdf'));
  newFile.setName(("Tiket "+ row[1] + " " + row[10]));

  DriveApp.getFolderById("1ovgR3F4fQ8SDHzlFr6u8sLocQFAl3Q40").removeFile(copyFile);
  DriveApp.getRootFolder().removeFile(newFile);

  copyFile.setTrashed(true)

  Logger.log('Ticket for '+row[1]+ ' has been created');
  return(newFile);

} // createTicket()
