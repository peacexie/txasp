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

ModID = RequestS("ModID","C",48) 
TelID = Trim(RequestS("TelID",3,48))
If ModID="" Then ModID="sysTels"

yAct = Request("yAct")
send = Request("send") 
Page = RequestS("Page","N",1)

KW = RequestS("KW","C",96)
KT = RequestS("KT","C",96)
TP = RequestS("TP","C",96)

If ModID="sysTels" Then
  ModName = "(系统)电话薄"
ElseIf ModID="userTels" Then
  ModName = "(用户)电话薄"
End If 
sqlK = " TelMod='"&ModID&"' "
if KW&"" <> "" then
  If KT="Name" Then
    sqlK = sqlK& " AND ( TelName LIKE '%"&KW&"%' ) "
  ElseIf KT="Num" Then
    sqlK = sqlK& " AND ( TelNum LIKE '%"&KW&"%' ) "
  ElseIf KT="User" Then
    sqlK = sqlK& " AND ( LogAUser LIKE '%"&KW&"%' ) "
  Else
    sqlK = sqlK& " AND ( TelID LIKE '%"&KW&"%' ) " 
  End If
end if
if TP&"" <> "" then
    sqlK = sqlK& " AND ( TelType='"&TP&"' ) " 
end if

If send = "ins" Then
  Msg = ""
  sql = "INSERT INTO [SmsTels] (TelID,TelMod,TelType,TelName,TelNum,LogAddIP,LogAUser,LogATime)VALUES"
  sql =sql& "('"&Get_AutoID(24)&"','"&ModID&"','"&RequestS("TelType",3,48)&"','"&RequestS("TelName",3,48)&"','"&RequestS("TelNum",3,36)&"','"&Get_CIP()&"','"&PrmUser&"','"&Now()&"')"
  'Response.Write sql
  Call rs_DoSql(conn,sql)
  Msg = "新增成功！"
ElseIf send = "del" Then
    sql = " DELETE FROM [SmsTels] WHERE TelID='"&TelID&"' "
	Call rs_DoSql(conn,sql)
	Msg = "删除成功!"
ElseIf send = "edt" Then
    sql = " UPDATE [SmsTels] SET TelType='"&RequestS("TelType",3,48)&"',TelName='"&RequestS("TelName",3,48)&"',TelNum='"&RequestS("TelNum",3,36)&"' "
    sql = sql & " ,LogEditIP='"&Get_CIP()&"',LogEUser='"&PrmUser&"',LogETime='"&Now()&"' "
	sql = sql & " WHERE TelID='"&TelID&"' "
	Call rs_DoSql(conn,sql)
	Msg = "修改成功!"

Else
    '
End If

sqlUBak = " And LogAUser='(user)' "
If PrmFlag="(Mem)" Then
  sqlK = sqlK&Replace(sqlUBak,"(user)",PrmUser)
End If
If Left(ModID,4)="user" And PrmFlag<>"(Mem)" Then
  sqlO = " ORDER BY TelID DESC "
Else
  sqlO = " ORDER BY TelType,TelID ASC "
End If
TelID = ""

    sql = " SELECT * FROM [SmsTels] WHERE "&sqlK&sqlO
   Set rs=Server.Createobject("Adodb.Recordset") ':Response.Write sql
   rs.Open Sql,conn,1,1
   rs.PageSize = 15
if int(Page)>rs.PageCount or int(Page)<1 Then
  Page = 1
End If
%>

<table width="100%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="#CCCCCC">
  <tr align="center" bgcolor="E0E0E0">
    <td height="21" colspan="8" align="center" bgcolor="f8f8f8"><table width="100%"  border="0" align="center" cellpadding="0" cellspacing="0">
        <form action="?" method="post" name="fsch9">
          <tr bgcolor="#FFFFFF">
            <td width="40%" rowspan="2" align="center" nowrap bgcolor="#FFFFFF"><strong><%=ModName%></strong> | <a href="tset.asp?send=tset&ModID=<%=ModID%>&PrmFlag=<%=PrmFlag%>">批量增加&gt;&gt;&gt;</a></td>
            <td colspan="2" align="right" nowrap bgcolor="#FFFFFF"><%If PrmFlag="(Mem)" Then%>
              (用户)电话薄
              <%Else%>
              <a 
            href="?ModID=sysTels">(系统)电话薄</a> | <a 
            href="?ModID=userTels">(用户)电话薄</a></a>
              <%End If%> 
              | <a href="type.asp?ModID=<%=ModID%>&PrmFlag=<%=PrmFlag%>">类别设置</a>&nbsp;</td>
          </tr>
          <tr bgcolor="#FFFFFF">
            <td align="center" nowrap bgcolor="#FFFFFF"><font color="#FF0000"><%=MSG%></font>
              <input name="ModID" type="hidden" id="ModID" value="<%=ModID%>">
            <input name="PrmFlag" type="hidden" id="PrmFlag" value="<%=PrmFlag%>"></td>
            <td width="30%" align="right" nowrap bgcolor="#FFFFFF"><select name="TP" id="TP">
              <option value="">(所有)</option>
              <%
			  If PrmFlag="(Mem)" Or Left(ModID,3)="sys" Then
			    If PrmFlag="(Mem)" Then 
				  sqlUsr2 = Replace(sqlUBak,"(user)",PrmUser)
				Else
				  sqlUsr2 = ""
				End If
			  %>
              <%=Get_rsOpt(conn,"SELECT TypID,TypName FROM SmsType WHERE TypMod='"&ModID&"' "&sqlUsr2&" ORDER BY TypID",TP)%>
              <%End If%>
              </select>
              <input name="KW" type="text" id="KW" value="<%=KW%>" size="8" maxlength="12">
              <select name="KT" id="KT">
                <option value="">(所有)</option>
                <%=Get_SOpt("Name;Num;User","名称;号码;User",KT,"")%>
              </select>
  <input type="submit" name="Submit3" value="查询"></td>
          </tr>
        </form>
      </table></td>
  </tr>
  <%
  if not rs.eof then
  %>
  <tr bgcolor="#FFFFFF">
    <td colspan="8" align="right"><%=RS_Page(rs,Page,"?send=pag&KW="&KW&"&KT="&KT&"&ModID="&ModID&"",1)%></td>
  </tr>
  <tr align="center" bgcolor="#FFFFFF">
    <td height="27" align="center" bgcolor="#e8e8e8">NO</td>
    <td height="27" align="center" bgcolor="#e8e8e8">名称</td>
    <td height="27" align="center" bgcolor="#e8e8e8">号码</td>
    <td align="center" bgcolor="#e8e8e8">类别</td>
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
		TelID = rs("TelID")
		TelMod = rs("TelMod")
		TelType = rs("TelType")
		TelName = rs("TelName")
		TelNum = rs("TelNum")
		LogAUser = rs("LogAUser")
		sqlUsr2 = Replace(sqlUBak,"(user)",LogAUser)
		
			    If PrmFlag="(Mem)" Then 
				  sqlUsr2 = Replace(sqlUBak,"(user)",PrmUser)
				Else
				  sqlUsr2 = ""
				End If
		
	  %>
  <form name="ff<%=i%>" method="post" action="?">
    <tr bgcolor="<%=col%>">
      <td align="right"><input name="send" type="hidden" id="XXsnd" value="edt">
        <input name="TP" type="hidden" id="XXsnd2" value="<%=TP%>">
        <%=i%></td>
      <td nowrap><input name="TelName" type="text" id="TelName" value="<%=TelName%>" size="18" maxlength="24">
      <input name="TelID" type="hidden" id="TelID" value="<%=TelID%>"></td>
      <td><input name="TelNum" type="text" id="TelNum" value="<%=TelNum%>" size="18" maxlength="24"></td>
      <td align="center"><select name="TelType" id="TelType">
        <%=Get_rsOpt(conn,"SELECT TypID,TypName FROM SmsType WHERE TypMod='"&ModID&"' "&sqlUsr2&" ORDER BY TypID",TelType)%>
      </select></td>
      <td align="center"><%=LogAUser%>
        <input name="PrmFlag" type="hidden" id="PrmFlag" value="<%=PrmFlag%>"></td>
      <td align="center"><%=TelMod%>
        <input name="ModID" type="hidden" id="ModID" value="<%=ModID%>"></td>
      <td align="center"><input type="submit" name="Submit" value="修改" ></td>
      <td align="center"><input type="button" name="Button" value="删除" 
			    onClick="Del_YN('?send=del&TelID=<%=TelID%>&ModID=<%=ModID%>&KT=<%=KT%>&KW=<%=KW%>&TP=<%=TP%>','<%=Show_jsStr(TelName)%>')"></td>
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
    <td height="27" align="center" bgcolor="#FFFFFF">名称</td>
    <td height="27" align="center" bgcolor="#FFFFFF">号码</td>
    <td align="center" bgcolor="#FFFFFF">&nbsp;</td>
    <td align="center" bgcolor="#FFFFFF">User</td>
    <td align="center" bgcolor="#FFFFFF">Mod</td>
    <td height="27" align="center" bgcolor="#FFFFFF">修改</td>
    <td height="27" align="center" bgcolor="#FFFFFF">删除</td>
  </tr>
  <%
  end if
  rs.Close
  set rs = nothing


'DefID = Get_AutoID(24)
'Response.Write "<br>"&TelID&MinID
	  
	  %>
  <form name="ff" method="post" action="?">
    <tr bgcolor="#e8e8e8">
      <td align="right" bgcolor="#e8e8e8"><input name="send" type="hidden" id="send" value="ins">
        <input name="TP" type="hidden" id="XXsnd3" value="<%=TP%>">
        新增 </td>
      <td bgcolor="#e8e8e8"><input name="TelName" type="text" id="TelName" size="18" maxlength="24"></td>
      <td><input name="TelNum" type="text" id="TelNum" value="13" size="18" maxlength="24"></td>
      <td align="center"><select name="TelType" id="TelType">
      
              <%
			  If PrmFlag="(Mem)" Or Left(ModID,3)="sys" Then
			    If PrmFlag="(Mem)" Then 
				  sqlUsr2 = Replace(sqlUBak,"(user)",PrmUser)
				Else
				  sqlUsr2 = ""
				End If
			  %>
              <%=Get_rsOpt(conn,"SELECT TypID,TypName FROM SmsType WHERE TypMod='"&ModID&"' "&sqlUsr2&" ORDER BY TypID",TelType)%>
              <%End If%>
      
      </select></td>
      <td align="center">(User)
        <input name="PrmFlag" type="hidden" id="PrmFlag" value="<%=PrmFlag%>"></td>
      <td align="center"><input name="ModID" type="hidden" id="ModID" value="<%=ModID%>"></td>
      <td colspan="2" align="center"><%If i<=12 Then %>
        <input type="button" name="Button" value="新增类别" onClick="chkData()">
        <%Else%>
        <input type="button" name="Button" value="多于12个，不能新增" disabled>
      <%End If%></td>
    </tr>
  </form>
  <tr align="left" bgcolor="#FFFFFF">
    <td colspan="10">说明: </td>
  </tr>
</table>
<script type="text/javascript">
function chkData()
{   errflag=1
        for(i=0;i<1;i++)
        {  

	 if (document.ff.TelName.value.length==0)
           { alert('[错误]\n 请输入名称!');
             document.ff.TelName.focus();
             errflag=0;
             break;
     }
	 if (document.ff.TelNum.value.length<=5)
           { alert('[错误]\n 号码 太短!');
             document.ff.TelNum.focus();
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
