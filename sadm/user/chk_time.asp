<!-- check time 1 -->
<style type="text/css">
.chkTim1 {
	color:#999;
}
.chkTim2 {
	color:#309;
	cursor:pointer;
	text-decoration:underline;
}
</style>
<script type="text/javascript" language="javascript" runat="server">
  var sTime,sDiff;
  var sysDate = new Date();
  var nYear   = sysDate.getYear(); 
  var nMonth  = sysDate.getMonth()+1;                                      
  var nDay    = sysDate.getDate();                                   
  sTime = nYear+"-"+nMonth+"-"+nDay+" "+sysDate.toTimeString();
  sDiff = sysDate.getTime(); 
  //sDiff = sDiff+1000000; // for Test
  //Response.Write(sTime); OK
</script>
<script type="text/javascript">
  var sTime = "<%=sTime%>";
  var sDiff = "<%=sDiff%>";
  var cTime,cDiff;
  var sysDate = new Date();
  var nYear   = sysDate.getFullYear(); 
  var nMonth  = sysDate.getMonth()+1;                                      
  var nDay    = sysDate.getDate();                                   
  var cTime = nYear+"-"+nMonth+"-"+nDay+" "+sysDate.toTimeString().replace(" Standard Time","");
  var cDiff = sysDate.getTime(); //alert(cDiff);
function checkTime()
{ 
  var nDiff = (cDiff-sDiff)/1000;
  if(nDiff<0){ nDiff = 0-nDiff; }
  var sMsg = "", chkMsg = "";
  if(nDiff>(60*60*24)) {
	chkMsg = "你的本地日期设置不正确!";
  } else if(nDiff>60*15) {
	chkMsg = "你的本地时间设置不正确!";
  } else { 
	chkMsg = "时差在误差范围内! ";
  } 
     sMsg += "服务器时间："+sTime+"\n";
     sMsg += "  本地时间："+cTime+"\n";
  if((sTime.indexOf("CST")<0)&&(cTime.indexOf("CST")<0)){
	 sMsg += " GMT时间差："+getDiff(nDiff)+" (sec)\n"; 
	 sMsg += "       -=>  "+chkMsg;
  }else{
	 sMsg += "  无法比较：CST时间 可表示为以下4个时间: \n"; 
	 sMsg += "1. UTC-600：Central Standard Time (USA) \n"; 
	 sMsg += "2. UTC-400：Cuba Standard Time \n"; 
	 sMsg += "3. UTC+800：China Standard Time \n"; 
	 sMsg += "4. UTC+900：Central Standard Time (Australia)。 "; 
  }
     alert(sMsg);
}
 
function getDiff(n){  
  //return n;  
  var g,k,s = "";  
  var g4 = new Array(4); //gap
  var s4 = new Array(4); //str
  var u4 = new Array(4); //unit
  u4[0] = "天 "; s4[0] = 0; g4[0] = 60*60*24;
  u4[1] = ":";   s4[1] = 0; g4[1] = 60*60;
  u4[2] = ":";   s4[2] = 0; g4[2] = 60;
  u4[3] = ".";   s4[3] = 0; g4[3] = 1; 
  for(i=0;i<g4.length;i++){
    g = g4[i];  
    if(n>g) { 
	  k = parseInt(n/g,10);
	  s4[i] = k; 
	  n = n-g*k;  
    } 
	if(!((i==0)&&(s4[0]==0))) { s += s4[i]+u4[i]; }
  }
  n = n.toFixed(3).replace("0.",""); //alert(n); // Math.round/ceil/floor S.toFixed
  return s+n;
}

</script>
<%=sTime%>
 &nbsp; <span class='chkTim1'>本地:</span><span class='chkTim2' title='[点击对时]' onClick="checkTime()">
<script type="text/javascript">document.write(cTime);</script>
</span> 
<!-- check time 2 -->