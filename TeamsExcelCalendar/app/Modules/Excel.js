async function CreateSheetTitles(Token,ItemId,workBookSession) {
    try {
        var BatchBuilder = [];
        var ContentHeader = {};
        ContentHeader["Content-Type"] = "application/json"
        ContentHeader["workbook-session-id"] = workBookSession.id;
        var TitleArray = [];
        TitleArray.push("Month");
        TitleArray.push("Day");
        TitleArray.push("Calendar");
        TitleArray.push("Event");
        TitleArray.push("Start");
        TitleArray.push("End");
        var TitleData = {};
        TitleData["values"] = [];
        TitleData["values"].push(TitleArray);
        var RangeUpdate = {};
        RangeUpdate.id = 1;
        RangeUpdate.method = "PATCH";
        RangeUpdate.url =  "/me/drive/items/" + ItemId + "/workbook/worksheets/Sheet1/range(address='A1:F1')";
        RangeUpdate.body = TitleData;                    
        RangeUpdate.headers = ContentHeader;
        BatchBuilder.push(RangeUpdate);      

        try {
            var RangeFormat = {};
            RangeFormat["bold"] = true;
           RangeUpdate = {};
           RangeUpdate.id = 2;
           RangeUpdate.method = "PATCH";
           RangeUpdate.url =  "/me/drive/items/" + ItemId + "/workbook/worksheets/Sheet1/range(address='A1:F1')/format/font";
           RangeUpdate.body = RangeFormat;                    
           RangeUpdate.headers = ContentHeader;
           BatchBuilder.push(RangeUpdate);
           RangeFormat = {};
           RangeFormat["horizontalAlignment"] = "Center";
           RangeFormat["verticalAlignment"] = "Center";
           RangeUpdate = {};
           RangeUpdate.id =3;
           RangeUpdate.method = "PATCH";
           RangeUpdate.url =  "/me/drive/items/" + ItemId + "/workbook/worksheets/Sheet1/range(address='A1:F1')/format";
           RangeUpdate.body = RangeFormat;                    
           RangeUpdate.headers = ContentHeader;
           BatchBuilder.push(RangeUpdate);
           var Batch = {};
           Batch["requests"] = BatchBuilder;
           await WorkBookPOST(Token,"https://graph.microsoft.com/v1.0/$batch",workBookSession.id,JSON.stringify(Batch) );
        } catch (error) {
            console.log(error);
        }
    } catch (error) {
        console.log(error);
    }
}

async function UpdateRangeData(Token,CalendarJson,ItemId,workBookSession){
    const monthNames = ["January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December"
    ];
    var i;
    var RangeStart = 2;
    var MonthRangeStart = 1;
    var MonthRangeEnd = 1;
    var BatchBuilder = [];
    var BatchBuilderMerge = [];
    var BatchBuilderFormat = [];
    var RowCount =0;
    for (i = 0; i <= 11; i++) {
        var RangeEnd = RangeStart;
        MonthRangeStart = MonthRangeEnd+1;
        var MonthTest = monthNames[i];
        var PostData = {};
        PostData["values"] = [];
        if (CalendarJson.has(MonthTest)) {
            var MonthMap = CalendarJson.get(MonthTest);
            var dcount;
            for (dcount = 1; dcount <= 32; dcount++) {
                if (MonthMap.has(dcount)) {                    
                    var DayMap = MonthMap.get(dcount);
                    DayMap.forEach(function (arrayValue) {
                        RowCount++;
                        PostData["values"].push(arrayValue);             
                    }                                     
                    );
                    RangeEnd += (DayMap.length - 1);           
                    RangeEnd++;
                    MonthRangeEnd = RangeEnd;
                    RangeStart = RangeEnd;                               
                }
            }
            if (MonthRangeStart != MonthRangeEnd) {
               
                var ContentHeader = {};
                ContentHeader["Content-Type"] = "application/json";
                ContentHeader["workbook-session-id"] = workBookSession.id;
                MonthRangeEnd = RangeEnd - 1;
                try {
                    var RangeValue = {};
                    RangeValue.id = BatchBuilder.length;
                    RangeValue.method = "PATCH";
                    RangeValue.url =  "/me/drive/items/" + ItemId + "/workbook/worksheets/Sheet1/range(address='" + ("A" + MonthRangeStart + ":F" + MonthRangeEnd) + "')";
                    RangeValue.body = PostData;                    
                    RangeValue.headers = ContentHeader;
                    BatchBuilder.push(RangeValue);
                    if(RowCount > 20){
                        Batch = {};
                        Batch["requests"] = BatchBuilder;
                        await WorkBookPOST(Token,"https://graph.microsoft.com/v1.0/$batch",workBookSession.id,JSON.stringify(Batch));
                        BatchBuilder = [];
                        RowCount = 0;
                    }
                } catch (error) {
                    console.log(error);
                }
                try {
                    var RangeSetting = {};
                    RangeSetting["across"] = false;
                    RangeValue = {};
                    RangeValue.id = BatchBuilderMerge.length;
                    RangeValue.method = "POST";
                    RangeValue.url =  "/me/drive/items/" + ItemId + "/workbook/worksheets/Sheet1/range(address='" + ("A" + MonthRangeStart + ":A" + MonthRangeEnd) + "')/merge";
                    RangeValue.body = RangeSetting;
                    RangeValue.headers = ContentHeader;
                    BatchBuilderMerge.push(RangeValue);
                   
                } catch (error) {
                    console.log(error);
                }
                try {
                    var RangeFormat = {};
                    RangeFormat["horizontalAlignment"] = "Center";
                    RangeFormat["verticalAlignment"] = "Center";
                    RangeValue = {};
                    RangeValue.id = BatchBuilderFormat.length;
                    RangeValue.method = "PATCH";
                    RangeValue.url =  "/me/drive/items/" + ItemId + "/workbook/worksheets/Sheet1/range(address='" + ("A" + MonthRangeStart + ":A" + MonthRangeEnd) + "')/format";
                    RangeValue.body = RangeFormat;
                    RangeValue.headers = ContentHeader;
                    BatchBuilderFormat.push(RangeValue);
               

                } catch (error) {
                    console.log(error);
                }
            }
           

        }
    }
    if(BatchBuilder.length > 0){
        Batch = {};
        Batch["requests"] = BatchBuilder;
        await WorkBookPOST(Token,"https://graph.microsoft.com/v1.0/$batch",workBookSession.id,JSON.stringify(Batch));
    }
    Batch = {};
    Batch["requests"] = BatchBuilderMerge;
    await WorkBookPOST(Token,"https://graph.microsoft.com/v1.0/$batch",workBookSession.id,JSON.stringify(Batch));
    Batch = {};
    Batch["requests"] = BatchBuilderFormat;
    await WorkBookPOST(Token,"https://graph.microsoft.com/v1.0/$batch",workBookSession.id,JSON.stringify(Batch));
    var AutoFitUpdateURL = "https://graph.microsoft.com/v1.0/me/drive/items/" + ItemId + "/workbook/worksheets/Sheet1/range(address='" + ("A1:F" + MonthRangeEnd) + "')/format/autofitColumns";
    await WorkBookPOST(Token, AutoFitUpdateURL, workBookSession.id,"");

}