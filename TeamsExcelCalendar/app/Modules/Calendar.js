async function GetCalendarDataYearlyAg(CalendarMap, Token, CalendarEmailAddress, CalendarId,CalendarName) {
    var Start = new Date()
    var StartString = Start.getFullYear() + "-01-01T00:00:01";
    var EndString = Start.getFullYear() + "-12-31T23:59:00";        
    
    const monthNames = ["January", "February", "March", "April", "May", "June",
        "July", "August", "September", "October", "November", "December"
    ];
    var calendarURL = "https://graph.microsoft.com/v1.0/users/" + CalendarEmailAddress + "/calendars('" + CalendarId + "')/calendarview?startdatetime=" + StartString + "&enddatetime=" + EndString + "&Top=500";
    var CalendarData = null;
    do {
        CalendarData = await GenericGraphGet(Token, calendarURL);
        CalendarData.value.forEach(function (element) {
            var StartTime = new Date(element.start.dateTime);
            var EndTime = new Date(element.end.dateTime);
            var MonthString = monthNames[StartTime.getMonth()];
            var DayString = StartTime.getDate();
            var CalendarEntry = [];
            CalendarEntry.push(monthNames[StartTime.getMonth()]);
            CalendarEntry.push(DayString);
            CalendarEntry.push(CalendarName);
            CalendarEntry.push(element.subject);
            if (element.isAllDay) {
                CalendarEntry.push("");
                CalendarEntry.push("");
            } else {
                CalendarEntry.push(StartTime.toLocaleTimeString());
                CalendarEntry.push(EndTime.toLocaleTimeString());
            }
            if (!CalendarMap.has(MonthString)) {
                var DayMap = new Map();
                CalendarMap.set(MonthString, DayMap); 
            }
            var MonthMap = CalendarMap.get(MonthString);
            var Checkdup = false;
            var PushCalendarEntry = true;
            if (!MonthMap.has(DayString)) {
                var DayArray = [];
                MonthMap.set(DayString, DayArray);               
            }else{
                Checkdup = true;
            }
            var DayMap = MonthMap.get(DayString);
            if(Checkdup){
                var dCount;
                for(dCount=0;dCount < DayMap.length;dCount++){
                    if(DayMap[dCount][3] == element.subject){
                        DayMap[dCount][2] += ("," + CalendarName);
                        PushCalendarEntry= false;
                    }
                }
            }
            if(PushCalendarEntry){
                DayMap.push(CalendarEntry);
            }
            MonthMap.set(DayString, DayMap);
            CalendarMap.set(MonthString, MonthMap);
        });
        if (CalendarData.hasOwnProperty("@odata.nextLink")) {
            calendarURL = CalendarData["@odata.nextLink"];
        } 
    }
    while (CalendarData.hasOwnProperty("@odata.nextLink"));
    return CalendarMap;
};









