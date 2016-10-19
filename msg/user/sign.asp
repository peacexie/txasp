<!--#include file="_config.asp"-->
<html>
<head>
<meta http-equiv="Pragma" content="no-cache">
<meta http-equiv="Expires" content="0">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title><%=Config_Name%>后台管理中心</title>
<script src="../../sadm/func1/Func_JS.js" type="text/javascript"></script>
<link href="../inc/spub.css" rel="stylesheet" type="text/css">
</head>
<body>
<%

'PrmFlag,ModID,

ModID = RequestS("ModID","C",48) 'sysSign,userSign,
TypID = Trim(RequestS("TypID",3,48))
HidID = RequestS("HidID","C",48)
ShwID = RequestS("ShwID","C",12)
If ModID="" Then ModID="sysSign"

yAct = Request("yAct")
send = Request("send") 
Page = RequestS("Page","N",1)

KW = RequestS("KW","C",96)
KT = RequestS("KT","C",96)

If PrmFlag="(Mem)" Then
  ModID = "userSign"
  ModName = "签名设置"
  sqlK = " And LogAUser='"&PrmUser&"' "
Else
 If ModID="userSign" Then
  ModName = "签名(用户)"
 Else 'If ModID="" Then
  ModName = "签名(管理员)"
 End If 
  sqlK = "  "
End If
sqlK = " TypMod='"&ModID&"' "&sqlK

if KW&"" <> "" then
  If KT="ID" Then
    sqlK = sqlK& " AND ( TypID LIKE '%"&KW&"%' ) " 
  ElseIf KT="Name" Then
    sqlK = sqlK& " AND ( TypName LIKE '%"&KW&"%' ) "
  ElseIf KT="User" Then
    sqlK = sqlK& " AND ( LogAUser LIKE '%"&KW&"%' ) "
  Else
    sqlK = sqlK& " AND ( TypID LIKE '%"&KW&"%' ) " 
  End If
end if

If send = "ins" Then
  Msg = ""
  id = HidID&ShwID
  sql = "SELECT * FROM [SmsType] WHERE TypID='"&id&"' "
  exF2 = rs_Exist(conn,sql)
  If exF2 = "YES" Then
    Msg = "新增失败！["&TypID&"]已经存在！"
  Else
	sql = "INSERT INTO [SmsType] (TypID,TypMod,TypName,LogAddIP,LogAUser,LogATime)VALUES"
	sql =sql& "('"&id&"','"&ModID&"','"&RequestS("TypName",3,48)&"','"&Get_CIP()&"','"&PrmUser&"','"&Now()&"')"
	Call rs_DoSql(conn,sql)
	Msg = "["&ShwID&"]新增成功！"
  End If
ElseIf send = "del" Then
    sql = " DELETE FROM [SmsType] WHERE TypID='"&TypID&"' "
	Call rs_DoSql(conn,sql)
	Msg = "删除成功!"
ElseIf send = "edt" Then
    sql = " UPDATE [SmsType] SET TypName='"&RequestS("TypName",3,48)&"' "
    sql = sql & " ,LogEditIP='"&Get_CIP()&"',LogEUser='"&PrmUser&"',LogETime='"&Now()&"' "
	sql = sql & " WHERE TypID='"&TypID&"' "
	Call rs_DoSql(conn,sql)
	Msg = "修改成功!"

Else
    '
End If

If Left(ModID,4)="user" And PrmFlag<>"(Mem)" Then
  sqlO = " ORDER BY TypID DESC "
Else
  sqlO = " ORDER BY TypID "
End If
TypID = ""

    sql = " SELECT * FROM [SmsType] WHERE "&sqlK&sqlO
   Set rs=Server.Createobject("Adodb.Recordset") ':Response.Write sql
   rs.Open Sql,conn,1,1
   rs.PageSize = 15
if int(Page)>rs.PageCount or int(Page)<1 Then
  Page = 1
End If
%>
<br>
<table width="620" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="#CCCCCC">
  <tr align="center" bgcolor="E0E0E0">
    <td height="21" colspan="7" align="center" bgcolor="f8f8f8"><table width="100%"  border="0" align="center" cellpadding="0" cellspacing="0">
        <form action="?" method="post" name="fsch9">
          
          <tr bgcolor="#FFFFFF">
            <td width="40%" rowspan="2" align="center" nowrap bgcolor="#FFFFFF"><strong><%=ModName%></strong></td>
            <td colspan="3" align="left" nowrap bgcolor="#FFFFFF">
            <%If PrmFlag="(Mem)" Then%>
            <a 
            href="?PrmFlag=(Mem)&ModID=userSign">签名设置</a>
            <%Else%>
            <a 
            href="?ModID=sysSign">签名设置</a> | <a 
            href="?ModID=userSign">签名设置(用户)</a>
            <%End If%>
             | <a 
            href="<%If PrmFlag<>"(Mem)" Then Response.Write "../admin/"%>send.asp?PrmFlag=<%=PrmFlag%>">返回&lt;&lt;&lt;</a>
            &nbsp; </td>
          </tr>
          
          <tr bgcolor="#FFFFFF">
            <td align="center" nowrap bgcolor="#FFFFFF"><font color="#FF0000"><%=MSG%></font></td>
            <td width="10%" align="center" nowrap bgcolor="#FFFFFF"><input name="ModID" type="hidden" id="ModID" value="<%=ModID%>">
            <input name="PrmFlag" type="hidden" id="PrmFlag" value="<%=PrmFlag%>"></td>
            <td width="30%" align="right" nowrap bgcolor="#FFFFFF"><select name="KT" id="KT">
                <option value="">(所有)</option>
                <%=Get_SOpt("ID;Name;User","代码;名称;User",KT,"")%>
              </select>
              <input name="KW" type="text" id="KW" value="<%=KW%>" size="8" maxlength="12">
              <input type="submit" name="Submit3" value="查询"></td>
          </tr>
        </form>
      </table></td>
  </tr>
  <%
  if not rs.eof then
  %>
  <tr bgcolor="#FFFFFF">
    <td colspan="7" align="right"><%=RS_Page(rs,Page,"?send=pag&KW="&KW&"&KT="&KT&"&ModID="&ModID&"",1)%></td>
  </tr>
  <tr align="center" bgcolor="#FFFFFF">
    <td height="27" align="center" bgcolor="#e8e8e8">NO</td>
    <td height="27" align="center" bgcolor="#e8e8e8">签名代码</td>
    <td height="27" align="center" bgcolor="#e8e8e8">签名内容</td>
    <td align="center" bgcolor="#e8e8e8">User</td>
    <td align="center" bgcolor="#e8e8e8">Mod</td>
    <td height="27" align="center" bgcolor="#e8e8e8">修改</td>
    <td height="27" align="center" bgcolor="#e8e8e8">删除</td>
  </tr>
  <%
  rs.AbsolutePage = Page
  for i = 1 to rs.PageSize
		if i mod 2 = 1 then
		  col = "#ffffff"
		else
		  col = "#f8f8f8"
		end if
		TypID = rs("TypID")
		TypName = rs("TypName")
		TypMod = rs("TypMod")
		LogAUser = rs("LogAUser")
	  %>
  <form name="ff<%=i%>" method="post" action="?">
    <tr bgcolor="<%=col%>">
      <td align="right"><input name="send" type="hidden" id="XXsnd" value="edt">
        <%=i%></td>
      <td><%=Mid(TypID,13)%>
        <input name="TypID" type="hidden" id="TypID" value="<%=TypID%>"></td>
      <td><input name="TypName" type="text" id="TypName" value="<%=TypName%>" size="18" maxlength="8"></td>
      <td align="center"><%=LogAUser%>
      <input name="PrmFlag" type="hidden" id="PrmFlag" value="<%=PrmFlag%>"></td>
      <td align="center"><%=TypMod%>
        <input name="ModID" type="hidden" id="ModID" value="<%=ModID%>"></td>
      <td align="center"><input type="submit" name="Submit" value="修改" ></td>
      <td align="center"><input type="button" name="Button" value="删除" 
			    onClick="Del_YN('?send=del&TypID=<%=TypID%>&ModID=<%=ModID%>&PrmFlag=<%=PrmFlag%>','<%=Show_jsStr(TypName)%>')"></td>
    </tr>
  </form>
  <%
    rs.movenext
    if rs.eof then exit for
    next
  Else
%>
  <tr align="center" bgcolor="#FFFFFF">
    <td height="27" align="center" bgcolor="#FFFFFF">NO</td>
    <td height="27" align="center" bgcolor="#FFFFFF">类别代码</td>
    <td height="27" align="center" bgcolor="#FFFFFF">类别名称</td>
    <td align="center" bgcolor="#FFFFFF">User</td>
    <td align="center" bgcolor="#FFFFFF">Mod</td>
    <td height="27" align="center" bgcolor="#FFFFFF">修改</td>
    <td height="27" align="center" bgcolor="#FFFFFF">删除</td>
  </tr>
  <%
  end if
  rs.Close
  set rs = nothing


If TypID<>"" Then
  HidID = Left(TypID,12)
  DefID = Next_ID(Mid(TypID,13),"A012",4)
Else
  HidID = Get_AutoID(12)
  DefID = "A012"
End If
'Response.Write "<br>"&TypID&MinID
	  
	  %>
  <form name="ff" method="post" action="?">
    <tr bgcolor="#e8e8e8">
      <td align="right" bgcolor="#e8e8e8"><input name="send" type="hidden" id="send" value="ins">
        新增 </td>
      <td bgcolor="#e8e8e8"><input name="ShwID" type="text" id="ShwID" value="<%=DefID%>" size="12" maxlength="12">
        <input name="HidID" type="hidden" id="HidID" value="<%=HidID%>"></td>
      <td><input name="TypName" type="text" id="TypName" size="18" maxlength="8"></td>
      <td align="center">(User)
      <input name="PrmFlag" type="hidden" id="PrmFlag" value="<%=PrmFlag%>"></td>
      <td align="center"><input name="ModID" type="hidden" id="ModID" value="<%=ModID%>"></td>
      <td colspan="2" align="center">
      <%If i<=6 Then %>
      <input type="button" name="Button" value="新增类别" onClick="chkData()">
      <%Else%>
      <input type="button" name="Button" value="多于6个，不能新增" disabled>
      <%End If%>
      </td>
    </tr>
  </form>
  <tr align="left" bgcolor="#FFFFFF">
    <td colspan="9">说明: <%=Get_AutoID(12)%></td>
  </tr>
</table>
<script type="text/javascript">
function chkData()
{   errflag=1
        for(i=0;i<1;i++)
        {  
	 if (document.ff.ShwID.value.length==0)
           { alert('[错误]\n 类别代码不能为空!');
             document.ff.ShwID.focus();
             errflag=0;
             break;
     }
	 if (document.ff.TypName.value.length==0)
           { alert('[错误]\n 请输入名称!');
             document.ff.TypName.focus();
             errflag=0;
             break;
     }

        }
          if (errflag==1)
          {    
		  document.ff.submit()
          }
}
  
                </script>
</body>
</html>
