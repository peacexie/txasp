<!--#include file="config.asp"-->
<html>
<head>
<meta http-equiv="Pragma" content="no-cache">
<meta http-equiv="Expires" content="0">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title><%=Config_Name%>后台管理中心</title>
<script src="../func1/Func_JS.js" type="text/javascript"></script>
<link href="../../inc/adm_inc/adm_style.css" rel="stylesheet" type="text/css">
</head>
<body>
<%

ModID = RequestS("ModID",3,48) : If ModID="" Then ModID="BisT124"
ModNM = RequestS("ModNM",3,48) : If ModNM="" Then ModNM="商务种类"
send = Request("send") 
TypID = Trim(RequestS("TypID",3,48))

If send = "ins" Then
  Msg = ""
  sql = "SELECT * FROM [WebType] WHERE TypID='"&TypID&"' "
  exF2 = rs_Exist(conn,sql)
  If exF2 = "YES" Then
    Msg = "新增失败！["&TypID&"]已经存在！"
  Else
	sql = "INSERT INTO [WebType] (TypID,TypMod,TypName,TypRem,TypTop,LogAddIP,LogAUser,LogATime)VALUES"
	sql =sql& "('"&TypID&"','"&ModID&"','"&RequestS("TypName",3,48)&"','"&RequestS("TypRem",3,1020)&"','"&RequestS("TypTop","N",888)&"','"&Get_CIP()&"','"&Session("UsrID")&"','"&Now()&"')"
	Call rs_DoSql(conn,sql)
	Msg = "["&TypID&"]新增成功！"
  End If
ElseIf send = "del" Then
    sql = " DELETE FROM [WebType] WHERE TypID='"&TypID&"' "
	Call rs_DoSql(conn,sql)
	Msg = "删除成功!"
ElseIf send = "edt" Then
    sql = " UPDATE [WebType] SET TypName='"&RequestS("TypName",3,48)&"',TypRem='"&RequestS("TypRem",3,1020)&"',TypTop='"&RequestS("TypTop","N",888)&"' "
    sql = sql & " ,LogEditIP='"&Get_CIP()&"',LogEUser='"&Session("UsrID")&"',LogETime='"&Now()&"' "
	sql = sql & " WHERE TypID='"&TypID&"' "
	Call rs_DoSql(conn,sql)
	Msg = "修改成功!"

ElseIf send="upd" Then
Set rs=Server.Createobject("Adodb.Recordset")

  OldMod = ""
  s1="" : s2="" : s3="" : s4="" 
  s=""
 
 Response.Write " <hr>更新成功！"
 Response.End()
Set rs = Nothing
Else
    '
End If

TypID = ""

    sql = " SELECT * FROM [WebType] WHERE TypMod='"&ModID&"' ORDER BY TypTop,TypID " 
   Set rs=Server.Createobject("Adodb.Recordset") ':Response.Write sql
   rs.Open Sql,conn,1,1
%>
<br>

<table width="620" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="#CCCCCC">
  <tr align="center" bgcolor="E0E0E0">
    <td height="21" colspan="7" align="center" bgcolor="f8f8f8"><table width="100%" border="0" cellpadding="1" cellspacing="0">
        <tr align="center">
          <td width="30%" rowspan="2" align="center"><strong>[<%=ModNM%>]</strong> 类别设置</td>
          <td align="left">&nbsp;</td>
          <td align="right" nowrap> <a 
            href="?ModID=Fields&ModNM=行业类别" >行业类别</a> | <a 
            href="?ModID=ComTyp&ModNM=企业性质" >企业性质</a> | <a 
            href="?ModID=GrdTyp&ModNM=学历等级" >学历等级</a> | <a 
            href="?ModID=Minzu&ModNM=民族种类" >民族种类</a> | <a 
            href="?ModID=FilTyp&ModNM=文件类型" >文件类型</a></td>
        </tr>
        <tr align="center">
          <td align="left" nowrap>&nbsp;</td>
          <form name="form1" method="post" action="?">
            <td align="right" nowrap><font color="#FF0000"><%=msg%></font>&nbsp; &nbsp;<a 
            href="?send=upd" target="_blank">刷新</a> | <a 
            href="?ModID=(Test)&ModNM=测试" >测试</a> | <a 
            href="?ModID=InfHead&ModNM=新闻头条">新闻头条</a> | <a 
            href="?ModID=BisU124&ModNM=会员留言" >会员留言</a> | <a 
            href="?ModID=BisT124&ModNM=商务种类" >商务种类</a></td>
          </form>
        </tr>
      </table></td>
  </tr>
  <tr align="center" bgcolor="#FFFFFF">
    <td height="27" align="center" bgcolor="#e8e8e8">NO</td>
    <td height="27" align="center" bgcolor="#e8e8e8">类别代码</td>
    <td height="27" align="center" bgcolor="#e8e8e8">类别名称</td>
    <td align="center" bgcolor="#e8e8e8">Rem</td>
    <td align="center" bgcolor="#e8e8e8">Top</td>
    <td height="27" align="center" bgcolor="#e8e8e8">修改</td>
    <td height="27" align="center" bgcolor="#e8e8e8">删除</td>
  </tr>

  <%
  if not rs.eof then
  i = 0
  Do While NOT rs.EOF
  i = i + 1
		if i mod 2 = 1 then
		  col = "#ffffff"
		else
		  col = "#f4f4f4"
		end if
		TypID = rs("TypID")
		TypName = rs("TypName")
		TypRem = rs("TypRem")
		TypTop = rs("TypTop")
	  %>
  <form name="ff<%=i%>" method="post" action="?">
    <tr bgcolor="<%=col%>">
      <td align="right"><input name="ModNM" type="hidden" id="ModNM" value="<%=ModNM%>">
        <input name="send" type="hidden" id="XXsnd" value="edt">
      <%=i%></td>
      <td><%=TypID%>
        <input name="TypID" type="hidden" id="TypID" value="<%=TypID%>">
        <input name="ModID" type="hidden" id="ModID" value="<%=ModID%>"></td>
      <td><input name="TypName" type="text" id="TypName" value="<%=TypName%>" size="15" maxlength="24">      </td>
      <td align="center"><input name="TypRem" type="text" id="TypRem" value="<%=TypRem%>" size="18" maxlength="255"></td>
      <td align="center"><input name="TypTop" type="text" id="TypTop" value="<%=TypTop%>" size="6" maxlength="3"></td>
      <td align="center"><input type="submit" name="Submit" value="修改" >      </td>
      <td align="center"><input type="button" name="Button" value="删除" 
			    onClick="Del_YN('?send=del&TypID=<%=TypID%>&ModID=<%=ModID%>&ModNM=<%=ModNM%>','<%=Show_jsStr(TypName)%>')">      </td>
    </tr>
  </form>
  <%
  rs.movenext
  Loop
  end if
	  
	  rs.close()
	  set rs = nothing

FstID = Mid(ModID,4,2)
If TypID<>"" Then
  MinID = Left(TypID,Len(TypID)-2)&"12"
  DefID = Next_ID(TypID,MinID,4)
Else
  DefID = FstID&"1212"
End If
'Response.Write "<br>"&TypID&MinID
	  
	  %>

  <form name="ff" method="post" action="?">
    <tr bgcolor="#e8e8e8">
      <td align="right" bgcolor="#e8e8e8"><input name="send" type="hidden" id="send" value="ins">
        <input name="ModID" type="hidden" id="ModID" value="<%=ModID%>">
        新增 </td>
      <td bgcolor="#e8e8e8"><input name="TypID" type="text" id="TypID" value="<%=DefID%>" size="8" maxlength="12"></td>
      <td><input name="TypName" type="text" id="TypName" size="15" maxlength="24"></td>
      <td align="center"><input name="TypRem" type="text" id="TypRem" size="18" maxlength="255"></td>
      <td align="center"><input name="TypTop" type="text" id="TypTop" value="<%=TypTop%>" size="6" maxlength="3"></td>
      <td colspan="2"><input type="button" name="Button" value="新增类别" onClick="chkData()">
        <input name="ModNM" type="hidden" id="ModNM" value="<%=ModNM%>"></td>
    </tr>
  </form>
  <tr align="left" bgcolor="#FFFFFF">
    <td colspan="9">&nbsp;</td>
  </tr>
</table>
<script type="text/javascript">
function chkData()
{   errflag=1
        for(i=0;i<1;i++)
        {  
	 if (document.ff.TypID.value.length==0)
           { alert('[错误]\n 类别代码不能为空!');
             document.ff.TypID.focus();
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
