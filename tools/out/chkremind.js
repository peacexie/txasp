

function chkRemind(){
				   
  var yDate   = new Date();
  var nYear   = yDate.getYear(); 
  var nMonth  = yDate.getMonth()+1;                                      
  var nDay    = yDate.getDate();
  var nHour   = yDate.getHours();  
  var nMinute = yDate.getMinutes(); 
  var nSecond = yDate.getSeconds();
  var nWeek   = yDate.getDay();
  var sDate   = nYear+"-"+nMonth+"-"+nDay;
  m0Sec = nHour*3600+nMinute*60+nSecond;
	
  srArr  = Rem_Flag.split("^");
  srFlag = srArr[0];                   // X-不启用;A-某1天;D-每天;W-按星期[0123456]；
  srVal  = srArr[1].replace("-0","-"); // Null; 2010-08-17; Null; 123456 
  if(srFlag=="D") {
	srFlag = "(Rem)D";	
  } else if((srFlag=="A")&&(sDate==srVal)) {
	srFlag = "(Rem)A";	
  } else if((srFlag=="W")&&(srVal.indexOf(nWeek)>=0)) {
	srFlag = "(Rem)W";	
  } else {
	srFlag = "(Silent)";	
  } //alert(srFlag);
  
  if(srFlag.substring(0,5)=="(Rem)") {
	for(i=0;i<Rem_Arr.length;i++){
	  iArr = Rem_Arr[i].split("^");
	  if(iArr[0]=="V"){
		jArr = iArr[1].split(":");
		secMin = jArr[0]*3600+jArr[1]*60;
		secMax = secMin + 60*5; //5分钟
		if((m0Sec>=secMin)&&(m0Sec<=secMax))
		  if(chkCookie(sDate,iArr[1])) {
			alert("系统提醒：\n"+iArr[1]+" "+iArr[2]+" ");
		  }
	  }
	}
  }

}

function chkCookie(xDate,xVal){ 
  var temp = document.cookie+';'; 
  var val = xDate+'@'+xVal;
  var rid = 'PeaceRemind'+xDate.replace("-","_")+'__'+xVal.replace(":","_");
  var str = ""; //alert(val +'\n'+ rid);
  var Pos = temp.indexOf('=',temp.indexOf(rid+'=')); 
  if(temp.indexOf(rid+'=')>=0)  
   { str = temp.substring(Pos+1,temp.indexOf(';',Pos)); }
  if(str.indexOf(val)>=0) {
	return false;
  } else {
    document.cookie=rid+'='+val+';';
    document.cookie='expires=Wednesday, 07-Sep-50 23:12:40 GMT';
	return true;  
  }
} 

setInterval("chkRemind()",15000);

