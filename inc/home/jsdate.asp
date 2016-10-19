
<%

verNow = Request("verNow") : If NOT isNumeric(verNow) Then verNow=0
divObj = Request("divObj") : idFunc = ""&verNow&divObj&""

%>

<%If aNull="bNull" Then%>
<script type="text/javascript">
<%End If%>

function jsDateMsg_<%=idFunc%>()
{
  
  var verNow = <%=verNow%>;
  var divObj = "<%=divObj%>";
  
  var sysDate = new Date();
  var nWeek   = sysDate.getDay();
  
  var nYear  = sysDate.getYear(); 
  var nMonth = sysDate.getMonth()+1;                                      
  var nDay   = sysDate.getDate();                                      
                                       
  var nHour   = sysDate.getHours();  
  var nMinute = sysDate.getMinutes(); 
  var nSecond = sysDate.getSeconds();
  
  var aWeek,sWeek,sDate,sTime;
  if(verNow==2) {  //英文
    aWeek = new Array("SUN.","MON.","TRU.","WED.","THR.","FRI.","SAT."); 
	sWeek = " "+aWeek[nWeek];
	sDate = nYear+"-"+nMonth+"-"+nDay+" ";
  } else if(verNow==3) {  //日文
    aWeek = new Array("日曜","月曜","火曜","水曜","木曜","金曜","土曜");
	sWeek = " "+aWeek[nWeek]+"日";
	sDate = nYear+"-"+nMonth+"-"+nDay+" ";
  } else if(verNow==4) {  //德文
    aWeek = new Array("Sonntag","Montag","Dienstag","Mittwoch","Donnerstag","Freitag","Samstag");
	sWeek = " "+aWeek[nWeek]+" ";
	sDate = nYear+"-"+nMonth+"-"+nDay+" ";
  } else { //默认中文
    aWeek = new Array("日","一","二","三","四","五","六");
	sWeek = " 星期"+aWeek[nWeek];
	sDate = nYear+"年"+nMonth+"月"+nDay+"日 ";
  }
  
  sTime = nHour+":"+nMinute+":"+nSecond;
  document.getElementById(divObj).innerHTML = sDate+sTime+sWeek;

}

setInterval("jsDateMsg_<%=idFunc%>()",1000);
// jsdate.asp?verNow=1&divObj=DateTime
// new Date().toLocaleString();
// "日一二三四五六".charAt(new Date().getDay();


<%If aNull="bNull" Then%>
</script>
<%End If%>


