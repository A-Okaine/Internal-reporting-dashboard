const config = {

  id: "1i4P9hMgLTapqfuB3JkXNPhfkAzwYPgMWHO5gnvs_sGc",

}

//Gmail template for RP dashboard

var emailRange = SpreadsheetApp.openById("1i4P9hMgLTapqfuB3JkXNPhfkAzwYPgMWHO5gnvs_sGc")

const email_scheme = {

  recipient: "product.department@lotto.com, bi.department@lotto.com",
  // from: "product.department@lotto.com",
  bcc: null,
  cc: null,

  subject: "Updated RP Dashboard " + Utilities.formatDate(new Date(), "GMT-4", "MM/dd/YYYY"),
  name: "RP Dashboard",

  body: "<br><br><p><a class='button' style='width: 312px; margin: 0;' href='link to File'>Label of Button</a></p>",
  signature: "<br><br>Best Regards,<br>The Lotto.com Product Department",
  attachments: [],

  htmlBody: 'Template'

}

// email_scheme.body = "<br><br><p><a class='button' style='width: 312px; margin: 0;' href='link to File'>Label of Button</a></p>"
// sendEmailByConfig(email_scheme);

function sendEmailByConfig(config) {

  GmailApp.sendEmail(
    config.recipient,
    config.subject,
    null,
    {
      // attachments:('RP reporting Dashboard'),
      bcc: config.bcc,
      cc: config.cc,
      from: config.from,
      replyto: config.from,
      name: config.name,
      htmlBody: HtmlService.createHtmlOutputFromFile("Template").getContent().replace("{{Content}}", config.body + config.signature),
    }


  );
  Logger.log("Email was Sent")
}


//Make sure you have the 'Drive" and 'Gmail' API services enabled
function xlsxToSheet() {
  Logger.log(config.id)
  let sheets = SpreadsheetApp.openById(config.id).getSheets()
  let date = new Date();
  let threads = GmailApp.search('label:"data-reporters-responsible-play-dashboard" is:unread');

  for (let i = 0; i < threads.length; i++) {
    var messages = threads[i].getMessages();
    var attachment = messages[0].getAttachments()[0];
    var msgDate = messages[0].getDate();
    var calendarmsgDate = Utilities.formatDate(msgDate, 'GMT-4', 'MM/dd/YYYY')
    var todaysDate = Utilities.formatDate(date, 'GMT-4', 'MM/dd/YYYY')

    if (calendarmsgDate === todaysDate) {

      Logger.log('import')

      let blob = attachment.copyBlob();

      Logger.log(attachment.getName());
      Logger.log(attachment.getContentType());
      let data;

      if (attachment.getName() == '00 Abnormal LT TTV Avg Incr.xlsx') {

        data = parseMSExcelBlob(blob)['00 Abnormal LT TTV Avg Incr']

        data.shift()

        for (let i = 0; i < data.length; i++) {
          data[i].shift()
        }
      }
      else {

        data = parseMSExcelBlob(blob)['Layout 1']
        Logger.log(data)
        data.shift()

        for (let i = 0; i < data.length; i++) {
          data[i].shift()
        }
        if (attachment.getName() == 'Abnormal Limit Change.xlsx') {
          data.splice(0, 1)
        }
      }
      Logger.log(data)

      let attachmentName = attachment.getName().slice(0, -5)
      Logger.log(attachmentName)
      let targetSheet = SpreadsheetApp.openById('1i4P9hMgLTapqfuB3JkXNPhfkAzwYPgMWHO5gnvs_sGc').getSheetByName(attachmentName)

      insertData(targetSheet, data)

    }

    else {
      Logger.log('do not import')
    }
    GmailApp.markThreadRead(threads[i])
  }

  let dateYear = Utilities.formatDate(date, "GMT-4", "YYYY");
  let dateMonth = Utilities.formatDate(date, "GMT-4", "MMMMMM");
  let rootID = "18Fr_Xv3FyQt-5oIsYDnhPg9F2zNQHJHE";
  let rootFolder = DriveApp.getFolderById(rootID);

  //   // Folder Directory creation
  let folderYear = mkdir(rootFolder, dateYear, "-p");
  let folderMonth = mkdir(folderYear, dateMonth, "-p");
  var name = 'RP dashboard ' + todaysDate
  let newspreadsheet = SpreadsheetApp.create(name)
  let newfile = DriveApp.getFileById(newspreadsheet.getId())
  newfile.moveTo(folderMonth)


  let ogsheets = SpreadsheetApp.openById("1i4P9hMgLTapqfuB3JkXNPhfkAzwYPgMWHO5gnvs_sGc").getSheets()
  ogsheets.forEach((sheet, index) => {
    sheet.copyTo(newspreadsheet)


  })
  let sheet1 = newspreadsheet.getSheetByName('Sheet1')
  newspreadsheet.deleteSheet(sheet1)


  let arraysheets = newspreadsheet.getSheets()
  for (let i = 0; i < arraysheets.length; i++) {
    let newsheetName = arraysheets[i].getName()
    let slicednewsheetname = newsheetName.slice(8)
    Logger.log(slicednewsheetname)
    arraysheets[i].setName(slicednewsheetname)
  }

  let copyUrl = newspreadsheet.getUrl()

  email_scheme.body = "<p> This email is a reference to the copy of today's responsible gaming dashboard </p><br><p><a class='button' style='width: 312px; margin: 0;' href=" + copyUrl + ">Responsible Gaming Dashboard</a></p>"
  sendEmailByConfig(email_scheme);

  let reference = SpreadsheetApp.openById('1i4P9hMgLTapqfuB3JkXNPhfkAzwYPgMWHO5gnvs_sGc').getSheets()
  for (let i = 1; i < reference.length; i++) {
    if (reference[i].getLastRow() !== 0 && reference[i].getLastColumn() !== 0) {
      Logger.log(reference[i].getLastRow())
      Logger.log(reference[i].getLastColumn())

      if (reference[i].getName() === '00 Abnormal LT TTV Avg Incr') {
        reference[i].getRange(3, 1, reference[i].getLastRow(), reference[i].getLastColumn()).clearContent()
      }

      else {
        reference[i].getRange(2, 1, reference[i].getLastRow(), reference[i].getLastColumn()).clearContent()
      }
      Logger.log('works')
    }
  }

  function insertData(sheet, data) {
    let ss = SpreadsheetApp.openById('1i4P9hMgLTapqfuB3JkXNPhfkAzwYPgMWHO5gnvs_sGc')

    if (data.length > 0) {
      Logger.log("insertData first cell value 1 is: " + sheet.getRange(1, 1).getValue())
      sheet.insertRowsAfter(1, data.length);

      sheet.getRange(1, 1, data.length, data[0].length).setValues(data)
      Logger.log("insertData first cell value 2 is: " + sheet.getRange(1, 1).getValue())
      Logger.log('right')

    } else {

      Logger.log('wrong')
    }
  }

  //function for new directory
  function mkdir(parentfolder, folderName, option) {
    Logger.log('Launching mkdir');

    var parentFolderContents = parentfolder.getFoldersByName(folderName);

    if (parentFolderContents.hasNext() == true) {
      Logger.log('Folder contains an folder(s) named "' + folderName + '"');

      return parentFolderContents.next(); // return the parent folder contents and move to the next
    } else { // if the first stsement is not the caase then run this
      Logger.log('Folder does not contain any folder(s) named "' + folderName + '"'); // log activity that folder does not contain anything

      if (option == "-p") { // if statement is true for failed new directory location
        Logger.log('[-p] Option enabled, Creating missing folder'); // log activity and then create missing folder

        parentfolder.createFolder(folderName); //Creates folder as a parent folder

        Logger.log('Folder "' + folderName + '" has been created'); // log activity that the folder, name has been created 

        return parentfolder.getFoldersByName(folderName).next(); // returns parent folder with the name and moves on to next one  
      }

      Logger.log('Exiting mkdir with false value with no actions'); // logs acitity that all directories have been completed
      return false; // returns what directories not completed
    }
  }

  /**
  * Parsing MS Excel files and returns values in JSON format.
  *
  * @param {BlobSource} blob the blob from MS Excel file
  * @param {String[]} requiredSheets the array of required sheet names (if omitted returns all)
  * @return {Object} Object of sheet names and values (2D arrays)
  */
  function parseMSExcelBlob(blob, requiredSheets) {
    var col_cache = {};
    var forbidden_chars = {
      "&lt;": "<",
      "&gt;": ">",
      "&amp;": "&",
      "&apos;": "'",
      "&quot;": '"'
    };

    blob.setContentType("application/zip");
    var parts = Utilities.unzip(blob);

    var relationships = {};
    for (var part of parts) {
      var part_name = part.getName();
      if (part_name === "xl/_rels/workbook.xml.rels") {
        var txt = part.getDataAsString();
        var rels = breakUpString(txt, '<Relationship ', '/>');
        for (var i = 0; i < rels.length; i++) {
          var rId = breakUpString(rels[i], 'Id="', '"')[0];
          var path = breakUpString(rels[i], 'Target="', '"')[0];
          relationships[rId] = "xl/" + path;
        }
      }
    }

    var worksheets = {};
    for (var part of parts) {
      var part_name = part.getName();
      if (part_name === "xl/workbook.xml") {
        var txt = part.getDataAsString();
        var sheets = breakUpString(txt, '<sheet ', '/>');
        for (var i = 0; i < sheets.length; i++) {
          var sh_name = breakUpString(sheets[i], 'name="', '"')[0];
          sh_name = decodeForbiddenChars(sh_name);
          var rId = breakUpString(sheets[i], 'r:id="', '"')[0];
          var path = relationships[rId];
          if (path.includes("worksheets")) {
            worksheets[path] = sh_name;
          }
        }
      }
    }

    requiredSheets = Array.isArray(requiredSheets) && requiredSheets.length && requiredSheets || [];
    var worksheets_needed = [];
    for (var path in worksheets) {
      if (!requiredSheets.length || requiredSheets.includes(worksheets[path])) {
        worksheets_needed.push(path);
      }
    }
    if (!worksheets_needed.length) return { "Error": "Requested worksheets not found" };

    var sharedStrings = [];
    for (var part of parts) {
      var part_name = part.getName();
      if (part_name === "xl/sharedStrings.xml") {
        var txt = part.getDataAsString();
        txt = txt.replace(/ xml:space="preserve"/g, "");
        sharedStrings = breakUpString(txt, '<t>', '</t>');
        for (var i = 0; i < sharedStrings.length; i++) {
          sharedStrings[i] = decodeForbiddenChars(sharedStrings[i]);
        }
      }
    }

    var result = {};
    for (var part of parts) {
      var part_name = part.getName();
      if (worksheets_needed.includes(part_name)) {
        var txt = part.getDataAsString();
        var cells = breakUpString(txt, '<c ', '</c>');
        var tbl = [[]];
        for (var i = 0; i < cells.length; i++) {
          var r = breakUpString(cells[i], 'r="', '"')[0];
          var t = breakUpString(cells[i], 't="', '"')[0];
          if (t === "inlineStr") {
            var data = breakUpString(cells[i].replace(/ xml:space="preserve"/g, ""), '<t>', '</t>')[0];
            data = decodeForbiddenChars(data);
          } else if (t === "s") {
            var v = breakUpString(cells[i], '<v>', '</v>')[0];
            var data = sharedStrings[v];
          } else {
            var v = breakUpString(cells[i], '<v>', '</v>')[0];
            var data = Number(v);
          }
          var row = r.replace(/[A-Z]/g, "") - 1;
          var col = colNum(r.replace(/[0-9]/g, "")) - 1;
          if (tbl[row]) {
            tbl[row][col] = data;
          } else {
            tbl[row] = [];
            tbl[row][col] = data;
          }
        }
        var sh_name = worksheets[part_name];
        result[sh_name] = squareTbl(tbl);
      }
    }


    function decodeForbiddenChars(txt) {
      for (var char in forbidden_chars) {
        var regex = new RegExp(char, "g");
        txt = txt.replace(regex, forbidden_chars[char]);
      }
      return txt;
    }

    function breakUpString(str, start_patern, end_patern) {
      var arr = [], raw = str.split(start_patern), i = 1, len = raw.length;
      while (i < len) { arr[i - 1] = raw[i].split(end_patern, 1)[0]; i++ };
      return arr;
    }

    function colNum(char) {
      if (col_cache[char]) return col_cache[char];
      var alph = "ABCDEFGHIJKLMNOPQRSTUVWXYZ", i, j, result = 0;
      for (i = 0, j = char.length - 1; i < char.length; i++, j--) {
        result += Math.pow(alph.length, j) * (alph.indexOf(char[i]) + 1);
      }
      col_cache[char] = result;
      return result;
    }

    function squareTbl(arr) {
      var tbl = [];
      var x_max = 0;
      var y_max = arr.length;
      for (var y = 0; y < y_max; y++) {
        arr[y] = arr[y] || [];
        if (arr[y].length > x_max) { x_max = arr[y].length };
      }
      for (var y = 0; y < y_max; y++) {
        var row = [];
        for (var x = 0; x < x_max; x++) {
          row.push(arr[y][x] || arr[y][x] === 0 ? arr[y][x] : "");
        }
        tbl.push(row);
      }
      return tbl.length ? tbl : [[]];
    }


    return result;
  }
}