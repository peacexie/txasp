<!--#include file="config.asp"-->
<html>
<head>
<meta http-equiv="Pragma" content="no-cache">
<meta http-equiv="Expires" content="0">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title><%=Config_Name%>后台管理中心</title>
<link href="../../inc/adm_inc/adm_style.css" rel="stylesheet" type="text/css">
<script src="../func1/Func_JS.js" type="text/javascript"></script>
</head>
<body>

<%

send   = Request("send") 
SysID   = RequestS("SysID",3,24)
SysName = RequestS("SysName",3,120) 
SysNEng = RequestS("SysNEng",3,120) 
SysType = RequestS("SysType",3,24) 
SysFlag = Request("SysFlag")
SysTop = RequestS("SysTop",3,2) 
ModID = RequestS("ModID",3,24) 
If ModID="" Then
  ModID = "Module" 'Groups
End If
ModName = rs_Val("","SELECT [SysName] FROM AdmSyst WHERE [SysID]='x"&ModID&"' Or [SysID]='Mod"&ModID&"' ")

If send = "ins" Then
  Msg = ""
  sql = "SELECT * FROM [AdmSyst] WHERE SysID='"&SysID&"' "
  exF2 = rs_Exist(conn,sql)
  If exF2 = "YES" Then
    Msg = "新增失败！["&SysID&"]已经存在！"
  Else
    sql = "INSERT INTO [AdmSyst] (SysID,[SysName],SysNEng,SysType,SysTop,LogAddIP,LogAUser,LogATime)VALUES"
	sql =sql& "('"&SysID&"','"&SysName&"','"&SysNEng&"','"&SysType&"','"&SysTop&"','"&Get_CIP()&"','"&Session("UsrID")&"','"&Now()&"')"
	Call rs_DoSql(conn,sql)
	Msg = "["&SysID&"]新增成功！"
  End If
ElseIf send = "del" Then
    sql = " DELETE FROM [AdmSyst] WHERE SysID='"&SysID&"'"
	Call rs_DoSql(conn,sql)
	Call Add_Log(conn,Session("UsrID"),"删除:"&ModName&"!!!"&SysID&"","[sadm_system]",Msg)
	Msg = "删除成功!"
ElseIf send = "edt" Then
    sql = " UPDATE [AdmSyst] SET [SysName]='"&SysName&"',SysNEng='"&SysNEng&"',SysType='"&SysType&"',SysTop='"&SysTop&"'"
    sql = sql & ",LogEditIP='"&Get_CIP()&"',LogEUser='"&Session("UsrID")&"',LogETime='"&Now()&"' "
	sql = sql & " WHERE SysID='"&SysID&"'" ': Response.Write sql : Response.End()
	Call rs_DoSql(conn,sql)
	Msg = "修改成功!"
End If

    Page = RequestS("Page","N",1)
    sql = " SELECT * FROM [AdmSyst] "
	sql =sql& " WHERE SysType='"&ModID&"' ORDER BY SysTop,SysID " 
	
   Set rs=Server.Createobject("Adodb.Recordset")
   rs.Open Sql,conn,1,1

%>
<br>
<table width="720" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="E0E0E0">
  <tr align="center" bgcolor="E0E0E0">
    <td height="27" colspan="7" align="center" bgcolor="f8f8f8"><table width="100%" border="0" cellpadding="3" cellspacing="0">
        <tr align="center" bgcolor="#FFFFFF">
          <td width="30%" align="center" nowrap bgcolor="#FFFFFF"><strong>[<%=ModName%>] 设定:</strong></td>
          <td align="left" nowrap><font color="#FF0000">提示: <%=msg%></font></td>
        </tr>
        <tr align="center" bgcolor="#FFFFFF">
          <td colspan="2" align="right" nowrap bgcolor="#FFFFFF">
            <a href='upd_para.asp'><font color="#0000FF">刷新</font></a> |
            <%
			fExpert = rs_Val("","SELECT ParText FROM AdmPara WHERE ParCode='Config_Mode'")
			If fExpert="isExpert" Or Chk_PermSP() Then
			%>
			<!--#include file="../../upfile/sys/config/sf_Groups.htm"--><br>
		    <!--#include file="../../upfile/sys/config/sf_Module.htm"-->
            <%Else%>
            <a href=?ModID=Module>权限模块</a> | 
            <!--<a href=?ModID=Depart>科室部门</a> | -->
            <%End If%>
            

          </td>
        </tr>
      </table></td>
  </tr>
  <tr align="center" bgcolor="E0E0E0">
    <td height="27" align="center">NO</td>
    <td height="27" align="center">代码</td>
    <td height="27" align="center">名称
      <%If ModID="Depart" Then%>
    (显示标记) 
    <%End If%></td>
    <td align="center">TOP</td>
    <td align="center">标记</td>
    <td height="27" align="center">修改</td>
    <td height="27" align="center">删除</td>
  </tr>
  <tr bgcolor="#333333">
    <td colspan="9" align="right"></td>
  </tr>
  <%
  i=0
  Do While NOT rs.EOF
  i=i+1
		if i mod 2 = 1 then
		  col = "#ffffff"
		else
		  col = "#f4f4f4"
		end if
		SysID = rs("SysID")
		SysName = Trim(rs("SysName")&"")
		SysNEng = Trim(rs("SysNEng")&"")
		SysType = rs("SysType")
		SysTop = rs("SysTop")

	  %>
  <form name="ff<%=i%>" method="post" action="?">
    <tr bgcolor="<%=col%>">
      <td align="right"><input name="send" type="hidden" id="XXsnd" value="edt">
        <%=i%></td>
      <td><%=SysID%>
        <input name="SysID" type="hidden" id="SysID" value="<%=SysID%>">
        <input name="ModID" type="hidden" id="ModID" value="<%=ModID%>"></td>
      <td><input name="SysName" type="text" id="SysName" value="<%=SysName%>" size="15" maxlength="60">
      <%If ModID="Depart" Then%>
      <input name="SysNEng" type="text" id="SysNEng" value="<%=SysNEng%>" size="8" maxlength="12">
      <%End If%>
      </td>
      <td><input name="SysTop" type="text" id="SysTop" value="<%=SysTop%>" size="3" maxlength="1"></td>
      <td><input name="SysType" type="text" id="SysType" value="<%=SysType%>" size="6" read.only></td>
      <td><input type="submit" name="Submit" value="修改" >
      </td>
      <td align="center"><input type="button" name="Button" value="删除" 
			    onClick="Del_YN('?send=del&SysID=<%=SysID%>&Img=<%=Img%>&ModID=<%=ModID%>','确认删除?小心操作哦！\n也许使你系统崩溃？！！')">
      </td>
    </tr>
  </form>
  <%
  rs.MoveNext()
  Loop
	  
	  rs.close()
	  set rs = nothing
	
If SysID<>"" Then
  MinID = Left(SysID,4)&"124"
  DefID = Next_ID(SysID,MinID,4)
Else
  DefID = Left(ModID,4)&"124"
End If
	  
	  %>
  <tr bgcolor="#cccccc">
    <td colspan="9" align="right"></td>
  </tr>
  <form name="ff" method="post" action="system.asp">
    <tr bgcolor="#e8e8e8">
      <td align="right"><input name="send" type="hidden" id="send" value="ins">
        <input name="ModID" type="hidden" id="ModID" value="<%=ModID%>">
        新增 </td>
      <td bgcolor="#e8e8e8"><input name="SysID" type="text" id="SysID" value="<%=DefID%>" size="12" maxlength="12"></td>
      <td><input name="SysName" type="text" id="SysName" size="15" maxlength="60">
      </td>
      <td><input name="SysTop" type="text" id="SysTop" value="6" size="3" maxlength="1"></td>
      <td><input name="SysType" type="text" id="SysType" value="<%=ModID%>" size="6" readonly></td>
      <td colspan="2"><input type="button" name="Button" value="新增模块" onClick="chkData()">
      <input name="Page" type="hidden" id="Page" value="<%=Page%>">      </td>
    </tr>
  </form>
  <tr bgcolor="#FFFFFF">
    <td colspan="9" align="left"><%If ModID="Groups" Or ModID="Else" Or ModID="Home" Then%>
      此部分 [群组设置] 开发时已经设置好，请不要轻易修改！代码建议用"x" 或 <%=Left(ModID,3)%> 开头！
      <%Else%>
      注意：代码建议用"<%=Left(ModID,3)%>"开头!
      <%End If%>
    <br>
    *. 本参数表一般不需要修改，开发人员请郑重修改！<span class="colF0F">正常运行的系统，请不要随便执行该操作！否则后果自负！！！</span></td>
  </tr>
</table>
<script type="text/javascript">
function chkData()
{   errflag=1
        for(i=0;i<1;i++)
        {  
	 if (document.ff.SysID.value.length==0)
           { alert('[错误]\n 代码!');
             document.ff.SysID.focus();
             errflag=0;
             break;
     }
	 if (document.ff.SysName.value.length==0)
           { alert('[错误]\n 请输入 名称!');
             document.ff.SysName.focus();
             errflag=0;
             break;
     }
	 <%if ModID="Groups" Or ModID="Else" Or ModID="Home" Then%>
	 <%Else%>
	 var tID = document.ff.SysID.value; 
	 if ( (tID.substring(0,3)!="<%=Left(ModID,3)%>") )
           { 
		     //alert('[错误]\n 代码请用"<%=Left(ModID,3)%>"开头!');
             //document.ff.SysID.focus();
             //errflag=0;
             //break;
     }
	 <%end if%>
        }
          if (errflag==1)
          {    document.ff.submit()
          }
}
  
                </script>
</body>
</html>
