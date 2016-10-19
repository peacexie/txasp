<!--#include file="config.asp"-->
<!--#include file="../func1/md5_func.asp"-->
<html>
<head>
<meta http-equiv="Pragma" content="no-cache">
<meta http-equiv="Expires" content="0">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title><%=Config_Name%>后台管理中心</title>
<link href="../../inc/adm_inc/adm_style.css" rel="stylesheet" type="text/css">
<script src="../../func1/FuncJS.js" type="text/javascript"></script>
</head>
<body>

<%

UsrID   = RequestS("UsrID",3,48)
UsrPW   = RequestS("UsrPW",3,48)
send   = RequestS("send",3,48)
UsrType = RequestS("UsrType",3,48) 
UsrName   = RequestS("UsrName",3,48)
UsrExp   = RequestS("UsrExp","D","1900-12-31")
Page = RequestS("Page","N",1)
ModID = RequestS("ModID",3,24) 
If ModID="Inner" Then
  ModName = "内部会员"
Else
  ModID = "Admin"
  ModName = "管理员"
End If

KW = RequestS("KW",3,48)
TP = RequestS("TP",3,48)
 sqlK = ""
If KW<>"" Then
 sqlK = sqlK& " AND (UsrID LIKE '%"&KW&"%' OR UsrName LIKE '%"&KW&"%') "
End If
If TP<>"" Then
 sqlK = sqlK& " AND (UsrType='"&TP&"') "
End If

If send = "ins" Then
  Msg = ""
  sql = "SELECT * FROM [AdmUser"&Adm_aUser&"] WHERE UsrID='"&UsrID&"' "
  exF1 = rs_Exist(conn,sql)
  sql = "SELECT * FROM [Member"&Mem_aMemb&"] WHERE MemID='"&UsrID&"' "
  exF2 = rs_Exist(conn,sql)
  If exF1 = "YES" OR exF2 = "YES" Then
    Msg = "新增失败！["&UsrID&"]已经存在 或 有其它用途!"
	'Response.Redirect(RedAddr)
  Else
    eStr = MD5_Adm(UsrPW&UsrID)
	sql = "INSERT INTO [AdmUser"&Adm_aUser&"] (UsrID,UsrName,UsrPW,UsrType,UsrPerm,UsrExp,LogAddIP,LogAUser,LogATime)VALUES"
	sql =sql& "('"&UsrID&"','"&UsrName&"','"&eStr&"','"&UsrType&"','{"&UsrType&"}','"&UsrExp&"','"&Get_CIP()&"','"&Session("UsrID")&"','"&Now()&"')"
	Call rs_DoSql(conn,sql)
	Msg = "["&UsrID&"]新增成功！"
  End If
ElseIf send = "del" Then
    sql = " DELETE FROM [AdmUser"&Adm_aUser&"] WHERE UsrID='"&UsrID&"'"
	Call rs_DoSql(conn,sql)
	Msg = "删除成功!"
ElseIf send = "edt" Then
    sql3 = ""
	if UsrPW <> "" then
	  eStr = MD5_Adm(UsrPW&UsrID)
	  sql3 = sql3 & " ,UsrPW='"&eStr&"' "
	  Msg  = " (密码已经修改)"
	else
	  sql3 = sql3 & " "
	  Msg  = " (密码未修改)"
	end if	
    if inStr(UsrType,"Admin")>0 then
	  sql3 = sql3 & " ,UsrPerm='{"&UsrType&"}:' "
	end if	
	sql = " UPDATE [AdmUser"&Adm_aUser&"] SET UsrType='"&UsrType&"',UsrName='"&UsrName&"',UsrExp='"&UsrExp&"' "&sql3&" "
	sql = sql & " ,LogEditIP='"&Get_CIP()&"',LogEUser='"&Session("UsrID")&"',LogETime='"&Now()&"' "
	sql = sql & " WHERE UsrID='"&UsrID&"'"
	Call rs_DoSql(conn,sql)
	Msg = "修改成功!"&Msg
	'Response.Write sql
Else

End If

    sql = " SELECT * FROM [AdmUser"&Adm_aUser&"] "
	sql = sql& " WHERE UsrType LIKE '"&Left(ModID,3)&"%' "&sqlK
	sql = sql& " ORDER BY UsrType, UsrID " 
   Set rs=Server.Createobject("Adodb.Recordset")
   rs.Open Sql,conn,1,1
   rs.PageSize = 15
if int(Page)>rs.PageCount or int(Page)<0 Then
  Page = 1
End If
%>
        <br>
        <table width="640" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="E0E0E0">

          <tr align="center" bgcolor="e8e8e8">
            <td height="24" colspan="9" align="center"  bgcolor="f8f8f8"><table width="100%" cellpadding="2" cellspacing="1" bgcolor="#FFFFFF">
              <tr>
                <td width="40%" rowspan="2" align="center" nowrap bgcolor="#FFFFFF"><strong><%=ModName%>设定</strong><br>
<font color="red"><%=msg%></font></td>
                <td width="20%" align="left" bgcolor="#FFFFFF">  &nbsp;<a 
				href="?ModID=Admin">管理员</a> | <a 
				href="?ModID=Inner">内部会员</a> | </td>
    
                  <form name="form1" method="post" action="?">
                    <td width="1%" align="right" nowrap><input name="ModID" type="hidden" id="ModID" value="<%=ModID%>">
                      <select id="TP" style="WIDTH: 90px" name="TP">
                      <option value="" >[不限]</option>
<%=Get_rsOpt(conn,"SELECT SysID,SysName FROM AdmSyst WHERE SysType='"&ModID&"'",UsrType)%>
                    </select>
                      <input name="KW" type="text" id="KW" value="" size="10" maxlength="10">
                        <input type="submit" name="Submit" value="搜索">
                    </td>
                </form>
                
              </tr>
              <tr>
                <td colspan="2" align="left"> <%=RS_Page(rs,Page,"?send=pag&ModID="&ModID&"&TP="&TP&"&KW="&KW&"",1)%></td>
              </tr>
            </table>
            </td>
          </tr>
          <tr align="center" bgcolor="e8e8e8">
            <td height="24" align="center"  bgcolor="e8e8e8">NO</td>
            <td align="center" >管理员帐号</td>
            <td align="center" >姓名</td>
            <td align="center" >密码</td>
            <td align="center" >所属分组</td>
            <td align="center" >&nbsp;</td>
            <td align="center" nowrap >权限</td>
            <td align="center" >修改</td>
            <td align="center" >删除</td>
          </tr>
          <tr bgcolor="#333333">
            <td colspan="9" align="right"></td>
          </tr>
          <%
  if not rs.eof then
  rs.AbsolutePage = Page
  for i = 1 to rs.PageSize
	    'i = i + 1
		if i mod 2 = 1 then
		  col = "#ffffff"
		else
		  col = "#f4f4f4"
		end if
		UsrID = rs("UsrID")
		UsrPW = rs("UsrPW")
		UsrType = rs("UsrType")
		UsrName = rs("UsrName") 
		UsrExp = rs("UsrExp") 
		
		If Left(UsrType,3)="Adm" Then
		 If UsrType="Admin" Then
		   setStr = "<font color=999999>设定</font>"
		 Else
		   setStr = "<a href='user_perm2.asp?PrmUser="&UsrID&"&ModID="&ModID&"'>设定</a>"
		 End If
		 goStr = "<font color=999999>登陆</font>"
		ElseIf Left(UsrType,3)="Inn" Then
		 If inStr(UsrType,"Admin")>0 Then
		   setStr = "<font color=999999>设定</font>"
		 Else
		   setStr = "<a href='user_pinn1.asp?PrmUser="&UsrID&"&ModID="&ModID&"'>设定</a>"
		 End If
		 goStr = "<a href='../../doc/index.asp?InnID="&UsrID&"&Act=AdmLogin' target='_blank'>登陆</a>"
		Else
		  setStr = "<font color=999999>设定</font>"
		  goStr = "<font color=999999>登陆</font>"
		End If

		If UsrID = Session("UsrID") Then
		  dis02 = "disabled"
		Else
		  dis02 = ""
		End If
		
	  %>
          <form name="ff<%=i%>" method="post" action="?">
            <tr bgcolor="<%=col%>">
              <td align="right"><input name="send" type="hidden" id="send" value="edt">
              <%=i%></td>
              <td><%=UsrID%>
                <input name="UsrID" type="hidden" id="UsrID" value="<%=UsrID%>">
                <input name="ModID" type="hidden" id="ModID" value="<%=ModID%>">
                <input name="TP" type="hidden" id="TP" value="<%=TP%>">
                <input name="KW" type="hidden" id="KW" value="<%=KW%>"></td>
              <td><input name="UsrName" type="text" id="UsrName" value="<%=UsrName%>" size="12" maxlength="12"></td>
              <td><input name="UsrPW" type="password" id="COName" value="" size="12" maxlength="24">
              </td>
              <td align="center"><select name="UsrType" id="UsrType">
                  <%=Get_rsOpt(conn,"SELECT SysID,SysName FROM AdmSyst WHERE SysType='"&ModID&"' ORDER BY SysTop,SysID",UsrType)%>
              </select></td>
              <td align="center" nowrap><%=goStr%></td>
              <td align="center" nowrap><%=setStr%></td>
              <td align="center"><input type="submit" name="Submit" value="修改" ></td>
              <td align="center"><input type="button" name="Button" value="删除" <%=dis02%> onClick="DelID('<%=UsrID%>','<%=UsrID%>')">
              </td>
            </tr>
          </form>
          <%
  rs.movenext
  if rs.eof then exit for
  next
  end if
	  rs.close()
	  set rs = nothing
	  
	  %>
          <form name="ffx" method="post" action="?">
            <tr bgcolor="cccccc">
              <td colspan="9" align="right"></td>
            </tr>
            <tr bgcolor="eeeeee">
              <td height="27" align="right" nowrap bgcolor="eeeeee"><input name="send" type="hidden" id="send" value="ins">
                <input name="ModID" type="hidden" id="ModID" value="<%=ModID%>">
                新增 </td>
              <td><input name="UsrID" type="text" id="UsrID" value="<%=lCase(Left(ModID,3)&Mid(Get_YYYYMMDD(""),3,4)&Rnd_ID("KEY",3))%>" size="12" maxlength="24"></td>
              <td><input name="UsrName" type="text" id="UsrName" size="12" maxlength="12"></td>
              <td><input name="UsrPW" type="password" id="UsrPW" size="12" maxlength="24"></td>
              <td align="center"><select name="UsrType" id="UsrType">
                <%=Get_rsOpt(conn,"SELECT SysID,SysName FROM AdmSyst WHERE SysType='"&ModID&"' ORDER BY SysTop,SysID",UsrType)%>
              </select></td>
              <td>&nbsp;</td>
              <td colspan="3"><input type="button" name="Button" value="新增管理员" onClick="chkData()">
              <input name="TP" type="hidden" id="TP" value="<%=TP%>">
              <input name="KW" type="hidden" id="KW" value="<%=KW%>"></td>
            </tr>
          </form>
          <tr bgcolor="#FFFFFF">
            <td colspan="9" align="left">注意：帐号建议用"<%=Left(ModID,3)%>"开头! <font color="red">注意:</font>&nbsp;如果密码为空,表示不修改密码!&nbsp;</td>
          </tr>
</table>
<script type="text/javascript">
function DelID(UsrID,msg)
{
     if (window.confirm("确认删除 ["+msg+"] ?"))
     {
        location="?send=del&ModID=<%=ModID%>&UsrID="+UsrID+"&TP=<%=TP%>&KW=<%=KW%>";
      }else
      {
       return;
      }
} 

function chkData()
{   errflag=1
        for(i=0;i<1;i++)
        {  
	 if (document.ffx.UsrID.value.length==0)
           { alert('[错误]\n 管理员帐号!');
             document.ffx.UsrID.focus();
             errflag=0;
             break;
     }
	 if (document.ffx.UsrPW.value.length==0)
           { alert('[错误]\n 请输入 管理员密码!');
             document.ffx.UsrPW.focus();
             errflag=0;
             break;
     }
        }
          if (errflag==1)
          {    document.ffx.submit()
          }
}

</script>
        
</body>
</html>