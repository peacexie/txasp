<!--#include file="config.asp"-->
<%

ID = RequestS("ID",3,48) 
KeyID = ""

SET rs=Server.CreateObject("Adodb.Recordset") 
rs.Open "SELECT * FROM "&ModTab&" WHERE KeyID='"&ID&"' ",conn,1,1 
  if NOT rs.eof then 
  Do While Not rs.EOF
KeyID = rs("KeyID")
KeyMod = rs("KeyMod")
KeyCode = rs("KeyCode")
InfType = rs("InfType")
InfSubj = Show_Text(rs("InfSubj"))
xxxCont = rs("InfCont")
InfPara = rs("InfPara")
SetRead = rs("SetRead")
SetHot = rs("SetHot")
SetTop = rs("SetTop")
SetShow = rs("SetShow")
ImgName = rs("ImgName")&""
LogATime = rs("LogATime")
LogAddIP = rs("LogAddIP")
LogAUser = rs("LogAUser")
LogETime = rs("LogETime")
LogEditIP = rs("LogEditIP")
LogEUser = rs("LogEUser")
  rs.MoveNext
  Loop
  end if 
rs.Close()
SET rs=Nothing 

Response.Write PathRoot
If KeyID = "" Then 
  Response.End() '.Redirect("adm_list.asp")
End If

%>
<html>
<head>
<meta http-equiv="Pragma" content="no-cache">
<meta http-equiv="Expires" content="0">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title><%=InfSubj%></title>
</head>
<body>
<table width="100%"  border="0" align="center" cellpadding="3" cellspacing="1">
  <tr>
    <th height="24" colspan="2"><strong><%=InfSubj%></strong></th>
  </tr>
  <tr>
    <th height="24" colspan="2"><hr></th>
  </tr>
  <tr>
    <td colspan="2"><%Call Show_sfData(ID,"fcont.htm")%></td>
  </tr>
  <tr>
    <th height="24" colspan="2"><hr></th>
  </tr>
  <tr>
    <td nowrap>阅读: <%=SetRead%></td>
    <td align="left" nowrap>发布: <%=LogATime%> &nbsp;&nbsp;<br>
      修改: <%=LogETime%> &nbsp;&nbsp;</td>
    <td align="left" nowrap>IP:<%=LogAddIP%><br>
    IP:<%=LogEditIP%></td>
    <td align="left" nowrap>ID:<%=LogAUser%><br>
    ID:<%=LogEUser%></td>
  </tr>
  <tr>
    <td>InfPara: <%=InfPara%><br>
      ImgName: <%=ImgName%></td>
    <td width="30%" nowrap>(<%=LogAddIP%>)</td>
  </tr>
</table>
</body>
</html>
