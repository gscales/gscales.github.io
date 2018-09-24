const buildScheduleTable = (Schedules,displayNameMap) => {
    var fbBoard = "<table><tr bgcolor=\"#95aedc\">";
    fbBoard += "<td align=\"center\" style=\"width=200;\" ><font size=\"2\"><b>User</b></td><td></td>";
    var dminute = "30";
    var stime = 8;
    var displayEndTime = 18;
    for (stime = 8; stime < displayEndTime; stime+=1) {
        if(dminute == "30"){
            dminute = "00";
        }else{
            dminute = "30"
        }
        var dstime = "0" + stime + ":" + dminute;
        if(stime > 10){
            dstime = stime + ":" + dminute;            
        }
        fbBoard += "<td align=\"center\" style=\"width=50;\" ><b><font size=\"2\">" + dstime + "</b></td>";
        if(dminute == "30"){
            dminute = "00";
        }else{
            dminute = "30"
        }
        var dstime = "0" + stime + ":" + dminute;
        if(stime > 10){
            dstime = stime + ":" + dminute;            
        }  
        fbBoard += "<td align=\"center\" style=\"width=50;\" ><b><font size=\"2\">" + dstime + "</b></td>";  
    }
    fbBoard += "</tr>";
   
    for (index = 0; index < Schedules.value.length; ++index) {

        var entry = Schedules.value[index];
        fbBoard += "<td bgcolor=\"#CFECEC\"><b><font size=\"2\">" + displayNameMap[entry.scheduleId] + "</b></td><td><img id=\"img" + entry.scheduleId + "\" style=\"border: 2px solid green;\" src=\"\" /></td>";   
        var availbilityArray = entry.availabilityView.split('');
        var ai =0;
        for(ai=0;ai<availbilityArray.length;++ai){
           var bgColour = "";
           switch (availbilityArray[ai]) {
               case "0" : bgColour = "bgcolor=\"#41A317\"";
               break;
               case "1" : bgColour = "bgcolor=\"#52F3FF\"";
               break;
               case "2" : bgColour = "bgcolor=\"#153E7E\"";
               break;
               case "3" : bgColour = "bgcolor=\"#4E387E\"";
               break;
               case "4" : bgColour = "bgcolor=\"#98AFC7\"";
               break;
            }
            fbBoard += "<td " + bgColour + "></td>";            
           
          
        }
        fbBoard += "</tr>";
  

    }
    fbBoard += "</table>"; 
    $('#data').empty()
    $('#data').append(fbBoard);
}