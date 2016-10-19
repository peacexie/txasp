<!--#include file="config.asp"-->
<html>
<head>
<meta http-equiv="Pragma" content="no-cache">
<meta http-equiv="Expires" content="0">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title><%=Config_Name%>后台管理中心</title>
<link href="../../inc/adm_inc/adm_style.css" rel="stylesheet" type="text/css">
<script src="../func1/Func_JS.js" type="text/javascript"></script>
<style type="text/css">
fieldset{ width:640px; overflow:hidden; padding:8px; margin:12px; }
legend { color:#00F; font-weight:bold; padding:3px; }
</style>

</head>
<body>
<%

CD = RequestS("ParCode",3,48)
dt = rs_val("","SELECT ParRem FROM AdmPara WHERE ParCode='"&CD&"'")
dt2=dt
dt3=dt
dt4=dt

If Session("MemID")&Session("UsrID")&Session("InnID")&""="" AND Right(lCase(CD),4)=".asp" Then
  Response.End()
End If

%>

<fieldset>
  <legend>[<%=CD%>]*** OrgText: </legend>
<%=dt%>
</fieldset>

<fieldset>
  <legend>[<%=CD%>]*** Org &lt;Pre&gt;: </legend>
<pre>
<%=dt4%>
</pre>
</fieldset>

<fieldset>
  <legend>[<%=CD%>]*** Server.HTMLEncode: </legend>
<%=Server.HTMLEncode(dt2)%> 
</fieldset>

<fieldset>
  <legend>[<%=CD%>]*** Text: </legend>
<%=Show_Text(dt3)%>  
</fieldset>


</body>
</html>