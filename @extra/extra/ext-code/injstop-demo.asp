<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Demo --- Stop Injection</title>
</head>

<body>
<form id="form1" name="form1" method="post" action="">
  <label>
    <input type="text" name="textfield" id="textfield" onblur="injStopSet()" />
  </label> 
  <a href="injstop.asp?yAct=Show" target="_blank">View </a> &nbsp; 
  <a href="injstop.asp?yAct=Clear" target="_blank">Clear </a>
</form>
<div style="display:none" id="injStopField"></div>
<script language="javascript">
function injStopSet() {
  var sFlag = "scr"+"ipt";
  var sUrl = "injstop.asp?yAct=Set";
  var sJava = "<"+sFlag+" language='javascript' src='"+sUrl+"'></"+sFlag+">";
  injStopField.innerHTML = sJava;
} 
</script>

</body>
</html>
