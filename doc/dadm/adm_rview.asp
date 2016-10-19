<!--#include file="_config.asp"-->
<%

ID = RequestS("ID",3,48) 
ModTab = "DocsRemark"
SET rs=Server.CreateObject("Adodb.Recordset") 


InfShow = ""
rs.Open "SELECT * FROM "&ModTab&" WHERE KeyID='"&ID&"' "&sqlK& " ",conn,1,1 
  if NOT rs.eof then 
  Do While Not rs.EOF
KeyID = rs("KeyID")
KeyMod = rs("KeyMod")
KeyCode = rs("KeyCode")
InfType = rs("InfType")
InfSubj = Show_Text(rs("InfSubj"))
LogAddIP = rs("LogAddIP") 
LogATime = rs("LogATime")
LogAUser = rs("LogAUser")
InfCont = rs("InfCont") 
  imgStr = ""
InfShow = vbcrlf&vbcrlf&"<p>"& InfCont&"</p>"
  rs.MoveNext
  Loop
  end if 
rs.Close()
SET rs=Nothing 
    If KeyID = "" Then 
        Response.End() '.Redirect("adm_list.asp")
    End If

%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title><%=InfSubj%></title>
<link href="../../inc/adm_inc/adm_style.css" rel="stylesheet" type="text/css">
</head>
<body>
<br>
<table width="540" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="F0F0F0">
  <tr bgcolor="#FFFFFF">
    <td colspan="3" align="center" nowrap bgcolor="#F0F0F0" class="col00F"><%=InfSubj%></td>
  </tr>
  <tr bgcolor="#FFFFFF">
    <td width="15%" align="center" bgcolor="#FFFFFF">姓名</td>
    <td bgcolor="#FFFFFF" class="col00F"><%=LogAUser%> &nbsp;&nbsp;</td>
    <td width="20%" align="center" bgcolor="#FFFFFF">态度: <span class="col00F"><%=InfType%></span></td>
  </tr>
  <tr bgcolor="#FFFFFF">
    <td colspan="3" bgcolor="#FFFFFF" style="font-size:14px;line-height:150%; padding:8px; color:#666666;"><%=InfCont%> </td>
  </tr>
  <tr bgcolor="#FFFFFF">
    <td colspan="3" align="center" bgcolor="#F0F0F0"><%=LogATime%> - <%=KeyCode%> - <%=KeyMod%> - <%=LogAddIP%> </td>
  </tr>
</table>

</body>
</html>
