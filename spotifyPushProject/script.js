function onOpen() {
    var ui = SpreadsheetApp.getUi();
    ui.createMenu('Invoice')
        .addItem('Who to charge', 'executeCheck')
        .addToUi();

}

function doOverdueCheck(columnDate, columnPeriod, columnPayment) //for each table
{
    var sheet = SpreadsheetApp.getActiveSheet();
    var dataRange = sheet.getDataRange();
    var lastRow = dataRange.getLastRow()
    var lastcolumn = dataRange.getLastColumn();
    var today = new Date();
    //var defaultDate = dataRange.getCell(6, 7).getValue();

    today.setHours(0, 0, 0, 0);

    sheet.getRange('H:H').clearContent();
    var overdue = sheet.getRange(1, 1);
    for (var r = 4; r <= lastRow; r++) {
        var payment; //another deadline
        var inv_date = dataRange.getCell(r, columnDate).getValue(); //column with dates
        var period = dataRange.getCell(r, columnPeriod).getValue();   //column with periods
        inv_date.setHours(0, 0, 0, 0);

        if (r === lastRow) {
            payment = 0
        } else {
            payment = sheet.getRange(r + 1, columnPayment).getValue();
        }


        if (payment != 0 && period != 0 && today >= inv_date) {
            sheet.getRange(r, columnDate).setBackgroundRGB(247, 245, 173); //yellow, paid before

        } else if (payment === 0 && period != 0 && today <= inv_date) {
            sheet.getRange(r, columnDate).setBackgroundRGB(218, 247, 173); //green, not expired

        } else if (payment === 0 && period != 0 && today >= inv_date) {
            sheet.getRange(r, columnDate).setBackgroundRGB(247, 176, 173); //red, expired
            overdue = sheet.getRange(r, columnDate);

        } else if (payment === 0 && period === 0 && today >= inv_date) {
            sheet.getRange(r, columnDate).setBackgroundRGB(255, 227, 248); //pink, awaits
        }
    }

    return overdue;
}

function getOverdueInfo(column) //column where the name is
{
    var sheet = SpreadsheetApp.getActiveSheet();
    var values = sheet.getRange(3, column, 2, 1).getValues();
    var date = doOverdueCheck(column + 5, column + 4, column + 1); //coordinates to cell - getRange

    var overdue =
    {
        name: values[0],
        email: values[1],
        dueDate: date.getValue(),
        dueDate_str: date.getDisplayValue()
    };

    var due_date = new Date(overdue.dueDate);
    due_date.setHours(0, 0, 0, 0);
    var today = new Date();
    today.setHours(0, 0, 0, 0);

    var difference = Math.abs(today.getTime() - due_date.getTime());
    overdue.numDays = Math.round(difference / (24 * 60 * 60 * 1000));

    return overdue;
}

function sendEmail(column) {
    var sheet = SpreadsheetApp.getActiveSheet();
    var overdue = getOverdueInfo(column);

    var templ = HtmlService
        .createTemplateFromFile('mail');

    templ.overdue = overdue;

    var message = templ.evaluate().getContent();

    var templ2 = HtmlService
        .createTemplateFromFile('mail_for_me');
    templ2.overdue = overdue;
    var myMessage = templ2.evaluate().getContent();

    MailApp.sendEmail({
        to: overdue.email.toString(),
        subject: 'Cześć, tu Weronka i jej fantastyczny skrypt kontrolujący Twoje płatności na Spotify!',
        htmlBody: message
    });

    MailApp.sendEmail({
        to: 'weronkagolonka@gmail.com',
        subject: 'Przypominajka o zaległościach',
        htmlBody: myMessage
    });

}

function checkAndSend(column) {
    var isOverdue = doOverdueCheck(column + 5, column + 4, column + 1).getDisplayValue()
    if (isOverdue != '') {
        sendEmail(column);
    }

}

function executeCheck() {
    checkAndSend(2);
    checkAndSend(9);
}