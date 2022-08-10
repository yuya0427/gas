function myFunction() {
    //シート名指定
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var sheet_event = spreadsheet.getSheetByName("テストシート");

    //リストから日付と行事を取得
    const date = sheet_event.getRange(1, 1, sheet_event.getLastRow()).getDisplayValues();
    const event = sheet_event.getRange(1, 3, sheet_event.getLastRow()).getValues();

    //カレンダーシートに行事を追記
    for (let i = 0; i < date.length; i++) {
        var spreaddate = String(date[0, i]).split('/');
        //行事を書き込む対象シートを検索
        var sheet_calendar = spreadsheet.getSheetByName(spreaddate[1] + "月");
        //行事を書き込む対象セルを検索
        var textFinder = sheet_calendar.createTextFinder(spreaddate[2]).matchEntireCell(true);
        var cells = textFinder.findAll();
        //行事を追記
        for (var j = 0; j < cells.length; j++) {
            var writtenevent = sheet_calendar.getRange(cells[j].offset(0, 1).getA1Notation()).getValue();
            if (writtenevent) {
                writtenevent = writtenevent + ","
            }
            sheet_calendar.getRange(cells[j].offset(0, 1).getA1Notation()).setValue(writtenevent + event[i]);
        }

    }

}