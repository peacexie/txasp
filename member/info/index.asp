<!--#include file="../../sadm/func1/func1.asp"-->
<!--#include file="../../sadm/func2/func2.asp"-->
<!--#include file="../admin/mconfig.asp"-->
<!--#include file="../../pfile/lang/vmemb.asp"-->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title>用户中心 (Member Center)</title>
<meta http-equiv="Pragma" content="no-cache">
<link href="../../inc/mem_inc/mem_style.css" rel="stylesheet" type="text/css">
</head>
<body>

<!--#include file="../../inc/mem_inc/mem_top.asp"-->
      
      <%

Call Chk_Perm2("","")

If Request("send")="OutAdmin" Then
  Session(UsrPStr)=""
  Session("UsrID")=""
ElseIf Request("send")="OutInner" Then 
  Session("InnPerm")=""
  Session("InnID")=""
End If

      msg2 = vMINX_A1&" [ "&vPMsg_WName&" ] "&vMINX_A2

	  MemID    = Session("MemID") 
	  MemName  = rs_Val("","SELECT MemName  FROM [Member"&Mem_aMemb&"] WHERE MemID='"&MemID&"'")   
	  MemStr   = "["&Session("MemID")&"] "&MemName
	  MemIP    = Get_CIP()
	  STime    = now()

%>
      <br>
      <table width="540" border="0" align="center" cellpadding="8" cellspacing="5" bgcolor="#F0F0F0">
        <tr>
          <td><table width="100%" border="0" cellpadding="2" cellspacing="1" bgcolor="#DDDDDD">
              <tr align="center" bgcolor="#FFFFFF">
                <td height="27" colspan="2">
                  <%=msg2%></td>
              </tr>
              <tr bgcolor="F8F8F8">
                <td colspan="2" align="right"></td>
              </tr>
              <%If Len(Session("UsrID")&"")>0 Then%>
              <tr bgcolor="#FFFFCC">
                <td align="right">超管用户:</td>
                <td>&nbsp;<a href="?send=OutAdmin"><span class="fnt00F">退出 超管用户[<%=Session("UsrID")%>]</span></a></td>
              </tr>
              <%End If%>
              <%If Len(Session("InnID")&"")>0 Then%>
              <tr bgcolor="#FFFFCC">
                <td align="right">内部用户:</td>
                <td>&nbsp;<a href="?send=OutInner"><span class="fnt00F">退出 内部用户[<%=Session("InnID")%>]</span></a></td>
              </tr>
              <%End If%>
              <tr bgcolor="#FFFFFF">
                <td width="30%" align="right"><%=vMINX_A3%>:</td>
                <td>&nbsp;<font color=blue><%=MemStr%></font> &nbsp; </td>
              </tr>
              <tr bgcolor="#FFFFFF">
                <td align="right">IP:</td>
                <td>&nbsp;<%=MemIP%></td>
              </tr>
              <tr bgcolor="#FFFFFF">
                <td align="right"><%=vMINX_A4%>:</td>
                <td><%=STime%> &nbsp;</td>
              </tr>
              <tr bgcolor="#FFFFFF">
                <td colspan="2" align="right"></td>
              </tr>
              <tr bgcolor="#FFFFFF">
                <td align="right"><font color="#0000FF"><%=vMINX_B1%></font></a></td>
                <td align="center">
                <a href="../../smod/gbook/info_list.asp?ModID=MemB324&PrmFlag=(Mem)"><%=vMINX_B2%></a> | 
                <A href="../../smod/gbook/info_list.asp?ModID=MemB224&PrmFlag=(Mem)"><%=vMINX_B3%></A> | 
                <A href="../../smod/gbook/info_list.asp?ModID=MemB124&PrmFlag=(Mem)"><%=vMINX_B4%></A> </td>
              </tr>
              <tr bgcolor="#FFFFFF">
                <td align="right"><font color="#FF0000"><%=vMINX_B5%></font></a></td>
                <td><a href="#" onClick="javascript:window.open('error.asp?send=SysKeyWD','SysKeyWD', 'left=120,top=120,width=480,height=360');">系统过滤关键字 (System keywords)</a></td>
              </tr>

              <tr bgcolor="F8F8F8">
                <td colspan="2" align="right"></td>
              </tr>
            </table>
            <table width="100%"  border="0" cellpadding="2" cellspacing="1">
              <tr bgcolor="#FFFFFF">
                <td><strong>系统通知 (System Notes)</strong></td>
                <td width="30%" align="center">&nbsp;</td>
              </tr>
              <%
    '" KeyMod='MemB224' AND (LnkName='"&Session("MemID")&"' OR LnkName='(Public)' )"
	sql = " SELECT TOP 5 * FROM [GboInfo] "
	sql =sql& " WHERE KeyMod='MemB224' AND (LnkName='"&Session("MemID")&"' OR LnkName='(Public)' )"
	sql =sql& " ORDER BY KeyID DESC" 
   Set rs=Server.Createobject("Adodb.Recordset")
   rs.Open Sql,conn,1,1
If NOT rs.EOF Then
Do While NOT rs.EOF
KeyID = rs("KeyID")
KeyMod = rs("KeyMod")
KeyFlag = rs("KeyFlag")
InfType = rs("InfType")
TypName = rs_Val("","SELECT TypName FROM WebTyps WHERE TypID='"&InfType&"'")
InfSubj = Show_Text(rs("InfSubj"))
InfCont = rs("InfCont")
InfReply = rs("InfReply")&""
LnkName = rs("LnkName")
LnkEmail = rs("LnkEmail")
SetRead = rs("SetRead")
SetShow = rs("SetShow")
ImgName = rs("ImgName")
LogAddIP = rs("LogAddIP")
LogAUser = rs("LogAUser")
LogATime = rs("LogATime")
If LnkName=Session("MemID") Then
  LnkName = "<font color=#ff0000>管理员留言</font>"
Else
  LnkName = "<font color=#ff00ff>系统公告</font>"
End If
   
%>
              <tr bgcolor="#FFFFFF">
                <td><font color=#FF00FF>[<%=TypName%>]</font> <a href="../gbook/info_view.asp?ID=<%=KeyID%>"><%=InfSubj%></a> (<%=LnkName%>)</td>
                <td><%=LogATime%></td>
              </tr>
              <%
rs.MoveNext
Loop
Else

End If  
 rs.Close()
 Set rs = Nothing
%>
              <tr bgcolor="#FFFFFF">
                <td>&nbsp; 
                  <!--<%=vbcrlf&vbcrlf&vbcrlf&vbcrlf&Session("MemPerm")%>--></td>
                <td align="center"><a href="../../smod/gbook/info_list.asp?ModID=MemB224&PrmFlag=(Mem)"><%=vMINX_C1%>&gt;&gt;&gt;</a></td>
              </tr>
            </table></td>
        </tr>
      </table>
      <!--#include file="../../inc/mem_inc/mem_bot.asp"-->
<%
scr = "script"
If verMemb="0" Then
 Response.Write "<"&scr&" type='text/javascript'>SetLang('Big5');</"&scr&">"
Else 'If verMemb="1" Then
 Response.Write "<"&scr&" type='text/javascript'>SetLang('xOrg');</"&scr&">"
End If
%>

</body>
</html>
