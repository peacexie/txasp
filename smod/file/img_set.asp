<!--#include file="iconfig.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Pragma" content="no-cache">
<meta http-equiv="Expires" content="0">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title><%=Config_Name%>后台管理中心</title>
<link href="../../inc/adm_inc/adm_style.css" rel="stylesheet" type="text/css">
<script src="../../sadm/func1/Func_JS.js" type="text/javascript"></script>
<script src="../../inc/home/jsInfo.js" type="text/javascript"></script>
</head>
<body>
<%

' TabID,upPath, ID,KeyID,ColID,ColNO,

send = Request("send")

TabID = RequestS("TabID",3,48) 
If TabID="" Then 
  Response.Write "错误！" 
  Response.End()
End If
upPath = ImgfPath()
'Response.Write ModID&" --- "&TabID&" --- "&upPath

ID = RequestS("ID",3,48)
KeyID = RequestS("KeyID",3,48) : If KeyID="" Then KeyID="KeyID"
ColID = RequestS("ColID",3,48) : If ColID="" Then ColID="ImgName"
NO = RequestS("NO","N",0)
WW = RequestS("WW","N",360)


if send = "del" then
  dName = RequestS("Img",3,255)  
  if dName <> "" then
  If NO<>"0" Then
    cName = rs_Val(""," SELECT "&ColID&" FROM "&TabID&" WHERE "&KeyID&"='"&ID&"'")
	aImg = Split(cName&"^^^^^^^^^^","^")
	sName = ""
	For i=1 To 9
	  If Int(NO)=i Then
	    sName = sName&"^"&""
	  Else
	    sName = sName&"^"&aImg(i)
	  End If
	Next
	sql = "UPDATE "&TabID&" SET "&ColID&"='"&sName&"' WHERE "&KeyID&"='"&ID&"'"
    Call rs_DoSql(conn,sql)
    Call fil_del(upRoot&upPath&"/"&dName)
  Else
	sql = "UPDATE "&TabID&" SET "&ColID&"='' WHERE "&KeyID&"='"&ID&"'"
    Call rs_DoSql(conn,sql)
    Call fil_del(upRoot&upPath&"/"&dName)
  End if
  End if
End if


SET rs=Server.CreateObject("Adodb.Recordset") 
rs.Open "SELECT "&ColID&" FROM "&TabID&" WHERE "&KeyID&"='"&ID&"'",conn,1,1 
  if NOT rs.eof then 
    fName = Trim(""&rs(ColID))
  else
    fName = ""
  end if
rs.Close()
SET rs=Nothing


'Response.Write " ** "&fName&NO '&ColID&"SELECT "&ColID&" FROM "&TabID&" WHERE "&KeyID&"='"&ID&"'"
If fName<>"" Then
  If NO="0" Then
	p = InStrRev(fName,"/")
	upPath = Left(fName,p-1)
	fName = Mid(fName,p+1)
  Else
    aImg = Split(fName&"^^^^^^^^^^","^")
    If aImg(NO)<>"" Then
      fName = aImg(NO)
    Else
	  fName = ""  
	End If
  End If
End If

'Response.Write upPath&" --- "&fName&" --- "&NO
%>
<table height="100%" border="0" align="left" cellpadding="0" cellspacing="0" bgcolor="F0F0F0">
  <form name="ffimg1" id="ffimg1" action="img_up.asp" enctype="multipart/form-data" method="post">
    <%
	if fName <> "" then
	fExt = lCase(Mid(fName,InStrRev(fName,"."),8))
	%>
    <tr bgcolor="#FFFFFF">
      <td bgcolor="#FFFFFF">&nbsp;<a href="<%=upRoot&upPath&"/"&fName%>" target="_blank">查看</a> &nbsp; <a href="?send=del&TabID=<%=TabID%>&upPath=<%=upPath%>&ID=<%=ID%>&KeyID=<%=KeyID%>&ColID=<%=ColID%>&NO=<%=NO%>&Img=<%=fName%>&WW=<%=WW%>">删除</a>
        <%
		If inStr(".jpg.jpeg.gif.",lCase(fExt))>0 Then
		fShow = "<img src='"&upRoot&upPath&"/"&fName&"' width='24' height='18' border=0 onload='javascript:setImgSize(this);' align='absmiddle' />"

		%>
        &nbsp; <a href="<%=upRoot&upPath%>/<%=fName%>?WW=<%=WW%>" target="_blank"><%=fShow%></a>
        <%End If%></td>
    </tr>
    <%else%>
    <tr bgcolor="#FFFFFF">
      <td nowrap bgcolor="#FFFFFF"><input type='file' name='<%=ColID%>' style="width:<%=WW%>px; ">
        <input name=view type=submit id="Button1" value="上传">
        <input name="TabID" type="hidden" id="TabID" value="<%=TabID%>">
        <input name="upPath" type="hidden" id="upPath" value="<%=upPath%>">
        <input name="ID" type="hidden" id="ID" value="<%=ID%>">
        <input name="KeyID" type="hidden" id="KeyID" value="<%=KeyID%>">
        <input name="ColID" type="hidden" id="ColID" value="<%=ColID%>">
        <input name="NO" type="hidden" id="NO" value="<%=NO%>">
      <input name="WW" type="hidden" id="WW" value="<%=WW%>"></td>
    </tr>
    <%end if%>
  </form>
</table>
</body>
</html>
