<!--#include file="config.asp"-->
<html>
<head>
<meta http-equiv="Pragma" content="no-cache">
<meta http-equiv="Expires" content="0">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title><%=Config_Name%>后台管理中心</title>
<script src="../../sadm/func1/Func_JS.js" type="text/javascript"></script>
<link href="../../inc/adm_inc/adm_style.css" rel="stylesheet" type="text/css">
</head>
<body>
<%

send   = Request("send") 
SysID   = RequestS("SysID",3,24)
SysName = RequestS("SysName",3,24) 
SysType = RequestS("SysType",3,24) 
SysFlag = Request("SysFlag")
SysDef  = RequestS("SysDef","c",24)
TPU = "System"
TPN = "模块"

If send = "ins" Then
  Msg = ""
  sql = "SELECT * FROM [MemSyst] WHERE SysID='"&SysID&"' "
  exF2 = rs_Exist(conn,sql)
  If exF2 = "YES" Then
    Msg = "新增失败！["&SysID&"]已经存在！"
  Else
    sql = "INSERT INTO [MemSyst] (SysID,[SysName],[SysType],SysDef,LogAddIP,LogAUser,LogATime)VALUES"
	sql =sql& "('"&SysID&"','"&SysName&"','"&SysType&"','"&SysDef&"','"&Get_CIP()&"','"&Session("UsrID")&"','"&Now()&"')"
	Call rs_DoSql(conn,sql)
	Msg = "["&SysID&"]新增成功！"
  End If
ElseIf send = "del" Then
    sql = " DELETE FROM [MemSyst] WHERE SysID='"&SysID&"'"
	Call rs_DoSql(conn,sql)
	Msg = "删除成功!"
	
ElseIf send = "edt" Then
    sql = " UPDATE [MemSyst] SET [SysName]='"&SysName&"',[SysDef]='"&SysDef&"'"
    sql = sql & ",LogEditIP='"&Get_CIP()&"',LogEUser='"&Session("UsrID")&"',LogETime='"&Now()&"' "
	sql = sql & " WHERE SysID='"&SysID&"'" ': Response.Write sql : Response.End()
	Call rs_DoSql(conn,sql)
	Msg = "修改成功!"
End If

    Page = RequestS("Page","N",1)
    sql = " SELECT * FROM [MemSyst] "
	sql =sql& " WHERE SysType='"&TPU&"' ORDER BY SysID " 
   Set rs=Server.Createobject("Adodb.Recordset")
   rs.Open Sql,conn,1,1
   rs.PageSize = 24 '2 '
if int(Page)>rs.PageCount or int(Page)<0 Then
  Page = 1
End If
%>
        <br>
        <table width="640" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="E0E0E0">
          <tr align="center" bgcolor="E0E0E0">
            <td height="27" colspan="6" align="center" bgcolor="f8f8f8"><table width="100%" border="0" cellpadding="3" cellspacing="0">
                <tr align="center" bgcolor="#FFFFFF">
                  <td width="50%" align="center" nowrap bgcolor="#FFFFFF"><strong><%=TPN%>设定:</strong> <a href="type.asp">社区</a> | <a href="system.asp">模块</a> | <a href="grade.asp">等级</a> </td>
                  <td align="center" bgcolor="#FFFFFF">&nbsp;&nbsp;&nbsp;</td>
                 
                    <td align="right" nowrap><font color="#FF0000"><%=msg%></font> </td>
                  
                </tr>
              </table></td>
          </tr>
          <tr align="center" bgcolor="E0E0E0">
            <td height="27" colspan="6" align="center" bgcolor="f8f8f8"><%if not rs.eof then
              Response.Write RS_Page(rs,Page,"?send=pag",1)
              end if%></td>
          </tr>
          <tr align="center" bgcolor="E0E0E0">
            <td height="27" align="center">NO</td>
            <td height="27" align="center"><%=TPN%>代码</td>
            <td height="27" align="center"><%=TPN%>名称</td>
            <td width="150" align="center">所属社区</td>
            <td height="27" align="center">修改</td>
            <td height="27" align="center">删除</td>
          </tr>
          <tr bgcolor="#333333">
            <td colspan="8" align="right"></td>
          </tr>
          <%
  if not rs.eof then
  rs.AbsolutePage = Page
  for i = 1 to rs.PageSize
		if i mod 2 = 1 then
		  col = "#ffffff"
		else
		  col = "#f4f4f4"
		end if
		SysID = rs("SysID")
		SysName = Trim(rs("SysName")&"")
		SysType = rs("SysType")
        SysDef = rs("SysDef")
	  %>
          <form name="ff<%=i%>" method="post" action="?">
            <tr bgcolor="<%=col%>">
              <td align="right"><input name="send" type="hidden" id="XXsnd" value="edt">
                <%=i%></td>
              <td><%=SysID%>
                <input name="SysID" type="hidden" id="SysID" value="<%=SysID%>">
              </td>
              <td><input name="SysName" type="text" id="SysName" value="<%=SysName%>" size="24" maxlength="12">
              </td>
              <td align="center"><select name="SysDef" id="SysDef" style="width:120; ">
				<%=Get_rsOpt(conn,"SELECT SysID,SysName FROM MemSyst WHERE SysType='Type' ORDER BY SysTop,SysID",SysDef)%>
              </select></td>
              <td><input type="submit" name="Submit" value="修改" >              </td>
              <td align="center"><input type="button" name="Button" value="删除" 
			    onClick="Del_YN('?send=del&SysID=<%=SysID%>&Img=<%=Img%>&TPU=<%=TPU%>','<%=SysName%>')">
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
          <tr bgcolor="#cccccc">
            <td colspan="8" align="right"></td>
          </tr>
          <form name="ff" method="post" action="?">
            <tr bgcolor="#e8e8e8">
              <td align="right"><input name="send" type="hidden" id="send" value="ins">
                新增 </td>
              <td bgcolor="#e8e8e8"><input name="SysID" type="text" id="SysID" value="x" size="12" maxlength="12"></td>
              <td><input name="SysName" type="text" id="SysName" size="24" maxlength="12"></td>
              <td align="center"> <select name="SysDef" id="SysDef" style="width:120; ">
                <%=Get_rsOpt(conn,"SELECT SysID,SysName FROM MemSyst WHERE SysType='Type' ORDER BY SysTop,SysID",SysDef)%>
              </select></td>
              <td colspan="2"><input type="button" name="Button" value="新增模块" onClick="chkData()">
                <input name="SysType" type="hidden" id="SysType" value="<%=TPU%>" size="12" maxlength="12" readonly></td>
            </tr>
          </form>
          <tr bgcolor="#000033">
            <td colspan="8" align="right"></td>
          </tr>
        </table>
        <script type="text/javascript">
function chkData()
{   errflag=1
        for(i=0;i<1;i++)
        {  
	 if (document.ff.SysID.value.length==0)
           { alert('[错误]\n 系统模块代码!');
             document.ff.SysID.focus();
             errflag=0;
             break;
     }
	 if (document.ff.SysName.value.length==0)
           { alert('[错误]\n 请输入 系统模块名称!');
             document.ff.SysName.focus();
             errflag=0;
             break;
     }
        }
          if (errflag==1)
          {    document.ff.submit()
          }
}
  
                </script>
        
</body>
</html>
