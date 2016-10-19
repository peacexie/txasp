<% 
response.expires=-1000000

pxDatS=Request("DatS")
pxPreN=Request("PreN")
pxNxtN=Request("NxtN")
if ( not IsNumeric(pxPreN) ) or ( not IsNumeric(pxNxtN) ) then
    pxPreN = 50
    pxNxtN = 30
end if
if not ( (pxPreN - pxNxtN)<=305 ) then
    pxPreN = 80
    pxNxtN = 50
end if

pxDatS = Replace(pxDatS,"/","-")
pxDatS = Replace(pxDatS,".","-")
arDatS = Split(pxDatS,"-")
if isDate(pxDatS) then
    pxYY=arDatS(0)
    pxMM=arDatS(1)
    pxDD=arDatS(2)
else
    pxYY=year(now())
    pxMM=month(now())
    pxDD=day(now())
end if

'Response.Write "<br>"&pxDatS&"||"&pxYY&"||"&pxDD&"||"&pxMM&""  '//debug
%>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<SCRIPT LANGUAGE="JavaScript">
	this.now = new Date();
    this.year = this.now.getFullYear();
	var y1 = this.year - <%=pxPreN%>;  
	var y2 = this.year + <%=pxNxtN%>;  

	var mm1,mm2;
	var m1 = 0;
	var m2 = 11;

	var dd1;
	var d1 = 1;

</Script>
<HTML>
<HEAD>
<TITLE>Date Select</TITLE>
<style type="text/css">
<!--
body {
	margin: 0px;
}
td,body { font-family:Arial;font-size:12pt }
A:link { color:blue; text-decoration:none; }
A:visited { color:blue; text-decoration:none; }
A:hover { text-decoration:underline overline; }
.bodytext {color:blue;}
.today {color:red;}
.sunday {color:black;}
--> 
</style>
<SCRIPT LANGUAGE="JavaScript">
         var MonthsList = new Array("1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12");
         var DaysInMonth = new Array(31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31);
         var WeekDaysList = new Array(" SUN ", " MON ", " TUE ", " WED ", " THR ", " FRI ", " SET ");

         function GetDays(month, year) 
		 {
            if (1 == month)
               return ((0 == year % 4) && (0 != (year % 100))) || (0 == year % 400) ? 29 : 28;
            else
               return DaysInMonth[month];
         }
         
         function GetToday() 
		 {
			this.year = <%=pxYY%>; //PNDate.getFullYear();
            this.month = <%=pxMM-1%>; //this.now.getMonth();
            this.day = <%=pxDD%>; //this.now.GetDate();
         }

         today = new GetToday();

         function newCalendar() 
		 {
            today = new GetToday();
            var parseYear = parseInt(document.all.year[document.all.year.selectedIndex].text);
			var parseMonth = parseInt(document.all.month[document.all.month.selectedIndex].text-1);
            var newCal = new Date(parseYear,parseMonth,dd1);
            var dayFlg = -1;
            var startDay = newCal.getDay();
            var daily = 0;
            if ((today.year == newCal.getFullYear()) && (today.month == newCal.getMonth())){ dayFlg = today.day;	}		   
            var tableCal = document.all.Calendar.tBodies.dayList;
            var intDaysInMonth = GetDays(newCal.getMonth(), newCal.getFullYear());
			for (var intWeek = 0; intWeek < tableCal.rows.length; intWeek++)
               for (var intDay = 0; intDay < tableCal.rows[intWeek].cells.length; intDay++) 
			   {
                  var cell = tableCal.rows[intWeek].cells[intDay];
                  if ((intDay == startDay) && (0 == daily)){ daily = dd1;}               
                  if ((daily > 0) && (daily <= intDaysInMonth)){ cell.innerText = daily++;}
                  else                                         { cell.innerText = "";}
				  cell.className = "bodytext";
				  if((intDay==0)||(intDay==6))  { cell.className = "sunday"; }
				  if(dayFlg==(daily-1))         { cell.className = "today"; }
			   }
            }

         function GetDate() {
		 	var sDate;
            var strArr;
            if ("TD" == event.srcElement.tagName)
               if ("" != event.srcElement.innerText) 
			   {
				  sDate = document.all.year.value + "-" + document.all.month.value + "-" + event.srcElement.innerText;				 				  
				  document.all.ret.value = sDate;
 		  		  window.close();
				  //strArr=sDate.split("-");
                  //sDate=strArr[0];
				}
         }

         function CheckYear() 
		 {
		 	if (document.all.year.value ==y1) 
			{
				mm1 = m1;
				mm2 = m2;
			} else {
				mm1=0;
				mm2=11;
			}
			ListMonth();
			CheckDay();
			newCalendar();
		 }
		 
         function CheckDay() 
		 {
		 	if ((document.all.year.value ==y1) && (document.all.month.value ==m1+1)) {
				dd1 = d1;
			} else {
				dd1 = 1;
			}
			newCalendar();
		 }

		 function ListYear() 
		 {
		 	document.all.year.length =y2-y1+1;
         	for (var intLoop = y1; intLoop <= y2;intLoop++) 
			{
				document.all.year.options[intLoop-y1].value = intLoop;
				document.all.year.options[intLoop-y1].text= intLoop;
				if (today.year == intLoop) 
					document.all.year.options[intLoop-y1].selected = true;
			}
		 }

		 function ListMonth() 
		 {
		 	document.all.month.length =mm2-mm1+1;
         	for (var intLoop = mm1; intLoop <= mm2;intLoop++) {
				document.all.month.options[intLoop-mm1].value = intLoop+1;
				document.all.month.options[intLoop-mm1].text= MonthsList[intLoop];
				if (today.month == intLoop) 
					document.all.month.options[intLoop-mm1].selected = true;
			}
		 }
      </SCRIPT>
</head>
<BODY ONLOAD="ListYear();CheckYear();newCalendar();" OnUnload="window.returnValue = document.all.ret.value;">
<TABLE ID="Calendar" cellpadding=1 cellspacing=1 border=0 align=center>
  <input type=hidden name=ret>
  <THEAD>
    <tr>
      <td bgcolor=#CCCCCC colspan=7 height=3></td>
    </tr>
    <TR>
      <TD COLSPAN=7 ALIGN=CENTER> Year
          <SELECT ID=year ONCHANGE="CheckYear();">
          </SELECT>
        - Month
        <SELECT ID=month ONCHANGE="CheckDay()">
        </SELECT>
      </TD>
    </TR>
    <tr>
      <td bgcolor=#999999 colspan=7 height=3></td>
    </tr>
    <TR>
      <SCRIPT LANGUAGE='JavaScript'>
         for (var intLoop = 0; intLoop < WeekDaysList.length; intLoop++)
         document.write("<TD bgcolor=#CCCCCC align=CENTER width=36>" + WeekDaysList[intLoop] + "</TD>");                                        
      </SCRIPT>
    </TR>
    <tr>
      <td bgcolor=#999999 colspan=7 height=1></td>
    </tr>
  </THEAD>
  <TBODY ID=dayList ALIGN=right ONCLICK="GetDate()" bgcolor=#F0F0F0>
    <SCRIPT LANGUAGE="JavaScript">
      for (var intWeeks = 0; intWeeks < 6; intWeeks++) 
      {
         document.write("<TR>");
         for (var intDays = 0; intDays < WeekDaysList.length;intDays++)
          { 
		    if((intDays==0)||(intDays==6)) { document.write("<TD title='Click' bgcolor=#E0E0E0 width=36>&nbsp;</TD>");}
		    else                           { document.write("<TD title='Click' width=36>&nbsp;</TD>");} 
		  }
         document.write("</TR>");
      }
    </SCRIPT>
    <tr>
      <td bgcolor=#999999 colspan=7 height=3></td>
    </tr>
  </TBODY>
</TABLE>
<Script Language="JavaScript1.2">
function Cancel() 
{
	document.all.ret.value = "";
	window.close();
}
</script>
</BODY>
</HTML>
