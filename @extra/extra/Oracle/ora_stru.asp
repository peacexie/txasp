<!--#include file="inc_perm.asp"-->
<%

Server.ScriptTimeout = 3600  'Second 999999
TBName = RequestS("xTab","C",96)

%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title>Tab Structure</title>
<meta http-equiv="Pragma" content="no-cache">
</head>
<body>
<table width='540' border=1 align='center'>
  <tr>
    <td colspan='4'><%=TBName%></td>
  </tr>
  <tr>
    <td nowrap>名称</td>
    <td nowrap>类别</td>
    <td nowrap>长度</td>
    <td width='50%' nowrap>说明</td>
  </tr>
  <%

SET rs=Server.CreateObject("Adodb.Recordset")  
rs.Open " SELECT TOP 1 * FROM "&TBName&" ",conn,1,1
vdt1 = ""
vdt2 = ""
vdt3 = ""
  FOR i = 0 TO rs.Fields.Count-1
	fName = rs.Fields(i).Name
	fType = rs.Fields(i).Type
	fSize = rs.Fields(i).DefinedSize
 If CStr(fType)="3" Then
   fTyp2 = "INT"
 ElseIf CStr(fType)="5" Then
   fTyp2 = "Float"
 ElseIf CStr(fType)="135" Then
   fTyp2 = "DateTime"
 ElseIf CStr(fType)="201" Then
   fTyp2 = "Memo"
 ElseIf CStr(fType)="202" Then
   fTyp2 = "varChar"
 Else
   fTyp2 = "("&fType&")"
 End If
%>
  <tr>
    <td><%=fName%></td>
    <td><%=fTyp2%></td>
    <td><%=fSize%></td>
    <td>&nbsp;</td>
  </tr>
  <%
  NEXT
rs.Close
SET rs=Nothing  


%>
</table>
</body>
</html>
