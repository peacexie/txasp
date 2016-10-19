<!--#include file="_config.asp"-->
<html>
<head>
<meta http-equiv="Pragma" content="no-cache">
<meta http-equiv="Expires" content="0">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title><%=Config_Name%>后台管理中心</title>
<script src="../inc/send_check.js" type="text/javascript"></script>
<script src="../../sadm/func1/Func_JS.js" type="text/javascript"></script>
<link href="../inc/spub.css" rel="stylesheet" type="text/css">
</head>
<body>
<%

'PrmFlag,ModID,

ModID = RequestS("ModID","C",48) 'sysTels,sysTemp,userTels,userTels
TmpID = Trim(RequestS("TmpID",3,48))
If ModID="" Then ModID="sysTemp"

yAct = Request("yAct")
send = Request("send") 
Page = RequestS("Page","N",1)

KW = RequestS("KW","C",96)
KT = RequestS("KT","C",96)
TP = RequestS("TP","C",96)

If ModID="sysTemp" Then
  ModName = "(系统)短信范本"
ElseIf ModID="userTemp" Then
  ModName = "(用户)短信范本"
End If 
sqlK = " TmpMod='"&ModID&"' "
if KW&"" <> "" then
  If KT="Cont" Then
    sqlK = sqlK& " AND ( TmpCont LIKE '%"&KW&"%' ) " 
  ElseIf KT="User" Then
    sqlK = sqlK& " AND ( LogAUser LIKE '%"&KW&"%' ) "
  Else
    sqlK = sqlK& " AND ( TmpID LIKE '%"&KW&"%' ) " 
  End If
end if
if TP&"" <> "" then
    sqlK = sqlK& " AND ( TmpType='"&TP&"' ) " 
end if

If send = "ins" Then
  Msg = ""
  sql = "INSERT INTO [SmsTemp] (TmpID,TmpMod,TmpType,TmpCont,LogAddIP,LogAUser,LogATime)VALUES"
  sql =sql& "('"&TmpID&"','"&ModID&"','"&RequestS("TmpType",3,48)&"','"&RequestS("TmpCont",3,140)&"','"&Get_CIP()&"','"&PrmUser&"','"&Now()&"')"
  Call rs_DoSql(conn,sql)
  Msg = "新增成功！"
ElseIf send = "del" Then
    sql = " DELETE FROM [SmsTemp] WHERE TmpID='"&TmpID&"' "
	Call rs_DoSql(conn,sql)
	Msg = "删除成功!"
ElseIf send = "edt" Then
    sql = " UPDATE [SmsTemp] SET TmpType='"&RequestS("TmpType",3,48)&"',TmpCont='"&RequestS("TmpCont",3,140)&"' "
    sql = sql & " ,LogEditIP='"&Get_CIP()&"',LogEUser='"&PrmUser&"',LogETime='"&Now()&"' "
	sql = sql & " WHERE TmpID='"&TmpID&"' "
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
  sqlO = " ORDER BY TmpID DESC "
Else
  sqlO = " ORDER BY TmpID DESC "
End If
TmpID = ""

    sql = " SELECT * FROM [SmsTemp] WHERE "&sqlK&sqlO
   Set rs=Server.Createobject("Adodb.Recordset") ':Response.Write sql
   rs.Open Sql,conn,1,1
   rs.PageSize = 5
if int(Page)>rs.PageCount or int(Page)<1 Then
  Page = 1
End If
%>
<table width="100%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="#CCCCCC">
  <tr align="center" bgcolor="E0E0E0">
    <td height="21" colspan="5" align="center" bgcolor="f8f8f8"><table width="100%"  border="0" align="center" cellpadding="0" cellspacing="0">
        <form action="?" method="post" name="fsch9">
          <tr bgcolor="#FFFFFF">
            <td width="40%" rowspan="2" align="center" nowrap bgcolor="#FFFFFF"><strong><%=ModName%></strong></td>
            <td colspan="3" align="right" nowrap bgcolor="#FFFFFF"><%If PrmFlag="(Mem)" Then%>
              (用户)短信范本
              <%Else%>
              <a 
            href="?ModID=sysTemp">(系统)短信范本</a> | <a 
            href="?ModID=userTemp">(用户)短信范本</a></a>
              <%End If%> 
              | <a href="type.asp?ModID=<%=ModID%>&PrmFlag=<%=PrmFlag%>">类别设置</a></td>
          </tr>
          <tr bgcolor="#FFFFFF">
            <td align="center" nowrap bgcolor="#FFFFFF"><font color="#FF0000"><%=MSG%></font></td>
            <td width="10%" align="center" nowrap bgcolor="#FFFFFF"><input name="ModID" type="hidden" id="ModID" value="<%=ModID%>">
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
                <%=Get_SOpt("Cont;User","内容;User",KT,"")%>
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
    <td colspan="5" align="right"><%=RS_Page(rs,Page,"?send=pag&KW="&KW&"&KT="&KT&"&TP="&TP&"&ModID="&ModID&"",1)%></td>
  </tr>
  <tr align="center" bgcolor="#FFFFFF">
    <td height="27" align="center" bgcolor="#e8e8e8">NO</td>
    <td height="27" align="center" bgcolor="#e8e8e8">内容</td>
    <td height="27" align="center" bgcolor="#e8e8e8">类别/User/Mod</td>
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
		TmpID = rs("TmpID")
		TmpMod = rs("TmpMod")
		TmpType = rs("TmpType")
		TmpCont = rs("TmpCont")
		LogAUser = rs("LogAUser")
		sqlUsr2 = Replace(sqlUBak,"(user)",LogAUser)
	  %>
  <form name="ff<%=i%>" method="post" action="?">
    <tr bgcolor="<%=col%>">
      <td align="right"><input name="send" type="hidden" id="XXsnd" value="edt">
        <input name="TmpID" type="hidden" id="TmpID" value="<%=TmpID%>">
        <input name="TP" type="hidden" id="XXsnd2" value="<%=TP%>">
        <%=i%></td>
      <td align="center"><textarea name="TmpCont" cols="36" rows="4" id="TmpCont" onBlur="chkChars(this,'nowChars',70)"><%=TmpCont%></textarea></td>
      <td align="center"><select name="TmpType" id="TmpType" style="width:120px">
          <%=Get_rsOpt(conn,"SELECT TypID,TypName FROM SmsType WHERE TypMod='"&ModID&"' "&sqlUsr2&" ORDER BY TypID",TmpType)%>
        </select>
        <br>
        <%=LogAUser%>/<%=TmpMod%>
        <input name="PrmFlag" type="hidden" id="PrmFlag" value="<%=PrmFlag%>">
      <input name="ModID" type="hidden" id="ModID" value="<%=ModID%>">
      <input name="Page" type="hidden" id="Page" value="<%=Page%>"></td>
      <td align="center"><input type="submit" name="Submit" value="修改" ></td>
      <td align="center"><input type="button" name="Button" value="删除" 
			    onClick="Del_YN('?send=del&TmpID=<%=TmpID%>&ModID=<%=ModID%>&KT=<%=KT%>&KW=<%=KW%>&TP=<%=TP%>','<%=Show_jsStr(TmpCont)%>')"></td>
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
    <td height="27" align="center" bgcolor="#FFFFFF">内容</td>
    <td height="27" align="center" bgcolor="#FFFFFF">类别/User/Mod</td>
    <td height="27" align="center" bgcolor="#FFFFFF">修改</td>
    <td height="27" align="center" bgcolor="#FFFFFF">删除</td>
  </tr>
  <%
  end if
  rs.Close
  set rs = nothing


DefID = Get_AutoID(24)
'Response.Write "<br>"&TmpID&MinID
	  
	  %>
  <form name="ff" method="post" action="?">
    <tr bgcolor="#e8e8e8">
      <td align="right" bgcolor="#e8e8e8"><input name="send" type="hidden" id="send" value="ins">
        <input name="TmpID" type="hidden" id="TmpID" value="<%=DefID%>">
        <input name="TP" type="hidden" id="XXsnd3" value="<%=TP%>">
        新增 </td>
      <td align="center" bgcolor="#e8e8e8"><textarea name="TmpCont" cols="36" rows="4" id="TmpCont" onBlur="chkChars(this,'nowChars',70)"></textarea></td>
      <td align="center" nowrap><select name="TmpType" id="TmpType" style="width:120px">
              <%
			  If PrmFlag="(Mem)" Or Left(ModID,3)="sys" Then
			    If PrmFlag="(Mem)" Then 
				  sqlUsr2 = Replace(sqlUBak,"(user)",PrmUser)
				Else
				  sqlUsr2 = ""
				End If
			  %>
              <%=Get_rsOpt(conn,"SELECT TypID,TypName FROM SmsType WHERE TypMod='"&ModID&"' "&sqlUsr2&" ORDER BY TypID",TelType)%>
              <%End If%>        </select>
        <br>
        (User)
        <input name="PrmFlag" type="hidden" id="PrmFlag" value="<%=PrmFlag%>">
        <input name="ModID" type="hidden" id="ModID" value="<%=ModID%>"></td>
      <td colspan="2" align="center"><input type="button" name="Button" value="新增范本" onClick="chkData()"></td>
    </tr>
  </form>
  <tr align="left" bgcolor="#FFFFFF">
    <td colspan="7">说明: <span id="nowChars" class="fntF00">&nbsp;<%=msg%></span> &nbsp;</td>
  </tr>
</table>
<script type="text/javascript">
function chkData()
{   errflag=1
        for(i=0;i<1;i++)
        {  

	 if (document.ff.TmpCont.value.length==0)
           { alert('[错误]\n 请输入内容!');
             document.ff.TmpCont.focus();
             errflag=0;
             break;
     }
	 if (document.ff.TmpCont.value.length>70)
           { alert('[错误]\n 内容请少于70字!');
             document.ff.TmpCont.focus();
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
