<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Untitled Document</title>
</head>
<body>
<table>
  <form name="fm01" method="post" action="para_date.asp">
    <tr bgcolor="f0f0f0">
      <td height="23" align="center">NO</td>
      <td ><input type='text' name='PRCode' size='12' maxlength='12'></td>
      <td><input type='text' name='PRName' size='18' maxlength='24'></td>
      <td><input name='PRDate' type='text' size='12' maxlength='8' 
			    >
        <img 
			  src="/img/tool/opendate.JPG" width="16" height="16" align="absmiddle"
			  onClick="OpenDate(document.fm01.PRDate,'','20','30','10','')"></td>
      <td colspan="3"><input type="button" name="Button" value="增加" onClick="chkData()">
        <input name="send" type="hidden" id="send" value="ins"></td>
    </tr>
  </form>
</table>
<script language="javascript">


//document.fm01.submit();
//alert("xx");

function Del_YN(YNaddr,msg)
{
    if(confirm(msg))
     {
         location.href = YNaddr;
         return true;
     }
         return false;
}

    function OpenDate(ObjDate,DatS,PreN,NxtN,RetB,xPath) 
    {   
        if(window.showModalDialog) 
        {   var sRtn; var i,resLen;
            resRtn = showModalDialog(xPath+"WinDate.asp?DatS="+DatS+"&PreN="+PreN+"&NxtN="+NxtN+"",
                                     "","center=yes;dialogWidth=240pt;dialogHeight=150pt");
            if (resRtn!="")
            {   switch (RetB) 
                {   case "6" : { // yyyymm
                       ObjDate.value = resRtn.substr(0,4) + resRtn.substr(5,2); break;}
                    case "7" : { // yyyy/mm  
                       ObjDate.value = resRtn.substr(0,7); break;}
                    case "8" : { // yyyymmdd  
                       ObjDate.value = resRtn.substr(0,4) + resRtn.substr(5,2) + resRtn.substr(8,2); break;}
                    default : { 
                       ObjDate.value = resRtn; break;}
            }   }
        }else{ 
            alert("Internet Explorer 4.0 or Later Is Required."); }
    }    

</script>
</body>
</html>
