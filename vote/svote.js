function chkVData(e,xUrl,xType)
{
  v = grVal("iVRad");
  if(xType=="Vote"){
    if (v!=""){ // fmObj.submit();
	  e.disabled = true;
	  showModalDialog(xUrl+"&iRad="+v,"center=yes;dialogWidth=640px;dialogHeight=480px"); 
	} else { 
	  alert("请选一个项目！"); }
  }else{
	  showModalDialog(xUrl,'x',"center=yes;dialogWidth=640px;dialogHeight=480px");  
  }
} 

function grVal(xID)
{
  var R = document.getElementsByName(xID);
  for (i=0;i<R.length;i++){
    if(R[i].checked){ return R[i].value; }
  }
  return "";
}