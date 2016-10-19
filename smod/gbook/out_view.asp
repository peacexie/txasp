
<!--#include file="config.asp"-->
<%

ID = RequestS("ID",3,48) 
ModTab = "GboSend"
yAct = RequestS("yAct",3,48)

If yAct<>"" Then
 If yAct="Del" Then
  Call rs_DoSql(conn,"DELETE FROM GboSend WHERE KeyID='"&ID&"'")
  Call del_sfCont("GboSend",ID)
 ElseIf yAct="Check" Then
  Call rs_DoSql(conn,"UPDATE GboSend SET SetShow='Y' WHERE KeyID='"&ID&"'")
 End If
 Response.Write "<script type='text/javascript'>alert('操作成功！点击关闭!');window.close();</script>"
End If

SET rs=Server.CreateObject("Adodb.Recordset") 
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
xxxCont = rs("InfCont") 
  imgStr = ""
  rs.MoveNext
  Loop
  end if 
rs.Close()
SET rs=Nothing 
    If KeyID = "" Then 
        Response.End() '.Redirect("adm_list.asp")
    End If

'xxxCont = rs("InfCont") 
InfCont = Show_sfRead(ID,"/fcont.htm")

If Left(lCase(InfCont),7)="<table " Then
  InfCont = Show_Html(InfCont)
  InfCont = Replace(InfCont,"&#60;!--KillA--&#62;","<!--KillA")
  InfCont = Replace(InfCont,"&#60;!--KillB--&#62;","KillB-->")
  InfCont = Replace(InfCont,"<INPUT id=button onclick=ChkFData(document.fmMail) type=button value=按钮 name=button>","")
  InfCont = Replace(InfCont,"<INPUT id=button2 type=reset value=重置 name=button2>","")
  InfCont = Replace(InfCont,"请认真输入，提交资料后，我们尽快处理","")
Else
  InfCont = InfCont 'Show_Form()
  InfCFlg = "Text"
End If


%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title><%=InfSubj%></title>
<link href="../../inc/adm_inc/adm_style.css" rel="stylesheet" type="text/css">
</head>
<body>
<%If InfCFlg<>"Text" Then%>
<%=InfCont%>
<%Else%>
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
    <td colspan="3" bgcolor="#FFFFFF" style="font-size:14px;line-height:150%; padding:8px; color:#666666;">
	<%=InfCont%>
    </td>
  </tr>
  <tr bgcolor="#FFFFFF">
    <td colspan="3" align="center" bgcolor="#F0F0F0"><%=LogATime%> - <%=KeyCode%> - <%=KeyMod%> - <%=LogAddIP%> </td>
  </tr>
  <tr bgcolor="#FFFFFF">
    <td colspan="3" align="center" bgcolor="#FFFFFF"><a href="?yAct=Check&ID=<%=ID%>">审核</a>　 |　<a href="?yAct=Del&ID=<%=ID%>">删除</a></td>
  </tr>
</table>
<%End If%>
</body>
</html>
