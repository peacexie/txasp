<!--#include file="config.asp"-->

<!--#include file="../../sadm/func2/func_sfile.asp" -->
<html>
<head>
<meta http-equiv="Pragma" content="no-cache">
<meta http-equiv="Expires" content="0">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title><%=Config_Name%>后台管理中心</title>
<link href="../../inc/adm_inc/adm_style.css" rel="stylesheet" type="text/css">
<script src="../../sadm/func1/Func_JS.js" type="text/javascript"></script>
</head>
<body>

<%

Set rs=Server.Createobject("Adodb.Recordset")
yAct = Request("yAct") 
Page = RequestS("Page","N",1)
MD = RequestS("MD","C",12)

    sql = " SELECT * FROM [AdmSyst] "
	sql =sql& " WHERE "
	If MD="" Then
	  sql =sql& " SysType IN('Info','Pics') "
	ElseIf Len(MD)<=3 Then
	  sql =sql& " SysType LIKE '"&MD&"%' "
	Else
	  sql =sql& " SysType='"&MD&"' "
	End If
	sql =sql& " ORDER BY SysType,SysTop,SysID " 
	
   rs.Open Sql,conn,1,1
   rs.PageSize = 18
   mRec = rs.RecordCount
If Int(Page)>rs.PageCount Or Int(Page)<1 Then
  Page = 1
End If

yAct = Request("yAct")
If yAct="Clear" Then
  Call rs_DoSql(conn,"DELETE FROM WebTyps WHERE TypMod='"&RequestS("ID","C",24)&"'")
  msg = "清空资料成功！"
  Call Add_Log(conn,Session("UsrID"),"清空类别---"&MD&"！","[reset_login]",Msg)
End If

%>
<br>
<table width="640" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="E0E0E0">
  <tr align="center" bgcolor="E0E0E0">
    <td height="27" colspan="10" align="center" bgcolor="f8f8f8"><table width="100%" border="0" cellpadding="3" cellspacing="0">
        <tr align="center" bgcolor="#FFFFFF">
          <td width="60%" align="center" bgcolor="#FFFFFF"><strong>信息类别</strong> (新闻/图片)</td>
          <td align="right" bgcolor="#FFFFFF"><font color="#FF0000"><%=msg%><a href="../type/type_List.asp?send=upd" target="_blank">刷新类别</a>&nbsp;
          <%If Config_Mode = "isExpert" Then%>
          | <a href="../type/type_data.asp?Group=<%=MD%>">数据处理</a>
          <%End If%>
           </font></td>
        </tr>
      </table></td>
  </tr>
  <tr align="center" bgcolor="E0E0E0">
    <td height="27" colspan="10" align="center" bgcolor="f8f8f8"><%= RS_Page(rs,Page,"?send=pag&TP="&TP&"",1)%></td>
  </tr>
  <tr align="center" bgcolor="E0E0E0">
    <td width="5%" height="27" align="center" nowrap>NO</td>
    <td height="27" align="center">名称</td>
    <td height="27" align="center" nowrap bgcolor="E0E0E0">编码</td>
    <td align="center" nowrap bgcolor="E0E0E0">Top</td>
    <td align="center" nowrap bgcolor="E0E0E0">N</td>
    <td align="center" nowrap bgcolor="E0E0E0">&nbsp;</td>
    <td align="center" nowrap bgcolor="E0E0E0">未审资料</td>
    <td align="center" nowrap bgcolor="E0E0E0">未审评论</td>
    <td align="center" nowrap bgcolor="E0E0E0">设置参数</td>
    <td height="27" align="center" nowrap bgcolor="E0E0E0">类别管理</td>
  </tr>
  <tr bgcolor="#333333">
    <td colspan="12" align="right" nowrap></td>
  </tr>
  <%
  If not rs.eof then
  rs.AbsolutePage = Page
  For i = 1 To rs.PageSize
		If i mod 2 = 1 Then
		  col = "#ffffff"
		Else
		  col = "#F8F8F8"
		End If
SysID = rs("SysID")
SysType = rs("SysType")
SysName = rs("SysName")
SysTop = rs("SysTop")
N = rs_Count(conn," WebTyps WHERE TypMod='"&SysID&"'")
M = rs_Count(conn," GboSend WHERE KeyMod='"&SysID&"' AND SetShow='N' ")
K = rs_Count(conn," "&rel_ModTab(SysID)&" WHERE KeyMod='"&SysID&"' AND SetShow='N' ")  
	  %>
  <form name="flist<%=i%>" method="post" action="?">
    <tr bgcolor="<%=col%>">
      <td align="right" nowrap><input type="hidden" name="hiddenField">
      <%=i%></td>
      <td><%=SysName%>
        <input name="yAct" type="hidden" id="yAct" value="edt">      </td>
      <td align="left" nowrap><%=SysID%>
        <input name="KeyID" type="hidden" id="SysID" value="<%=SysID%>"></td>
      <td align="center" nowrap bgcolor="<%=col%>"><%=SysTop%></td>
      <td align="center" nowrap bgcolor="<%=col%>"><%=N%></td>
      <td align="center" nowrap>
      <%If Config_Mode = "isExpert" Then%>
      <a onClick="Del_YN('?ID=<%=SysID%>&Page=<%=Page%>&yAct=Clear&MD=<%=MD%>','确认删除?小心操作哦！')" href="#" >清空</a>
	  <%End If%>
      </td>
      <td align="center" nowrap>
      <%If K>0 Then%>
      <A href='../../smod/info/info_list.asp?ModID=<%=SysID%>'><font color=red><%=K%>条</font></A>
      <%ElseIf inStr("(Inf,Pic)",Left(SysID,3))>0 Then%>
      <A href='../../smod/info/info_list.asp?ModID=<%=SysID%>'><%=K%>条</A>
      <%Else%>
      <font color=gray>------</font>
	  <%End If%>
      </td>
      <td align="center" nowrap>
      <%If M>0 Then%>
      <A href='../../smod/gbook/out_list.asp?ModID=<%=SysID%>'><font color=red><%=M%>条</font></A>
      <%ElseIf inStr("(Inf,Pic)",Left(SysID,3))>0 Then%>
      <A href='../../smod/gbook/out_list.asp?ModID=<%=SysID%>'><%=M%>条</A>
      <%Else%>
      <font color=gray>------</font>
	  <%End If%>
      </td>
      <td align="center" nowrap><a href="type_set.asp?ModID=<%=SysID%>">设置参数</a></td>
      <td align="center" nowrap><a href="../../smod/type/type_list.asp?ModID=<%=SysID%>">类别管理</a></td>
    </tr>
  </form>
  <%
  rs.Movenext
  If rs.Eof Then Exit For

  Next
%>
  <%
  End If
	  
	  rs.Close()
	  Set rs = Nothing
	  
	  If mRec=1 Then
	    Response.Redirect "../../smod/type/type_list.asp?ModID="&SysID&""
	  End If
	  %>
  <form name="fmDir" method="post" action="">
  <tr bgcolor="#999999">
    <td colspan="12" align="left" bgcolor="#F0F0F0"><input name="sDir" type="text" id="sDir" value="../../smod/type/type_list.asp?ModID=InfX123" size="45" maxlength="255">
      <input type="button" name="button" id="button" value="类别管理" onClick="goType()">
      &nbsp;(用于主菜单中未出现的类别管理) </td>
  </tr>
  </form>
  <form name="fmRep" method="post" action="?" target="_self">
    <tr align="center" bgcolor="#FFFFFF">
      <td colspan="12" align="left">../adupd/data_imp01.asp</td>
    </tr>
  </form>
</table>
<%

If yAct="ImgUpdate" Then 
for i=1 to 2
If i=1 Then iTab="InfoNews"
If i=2 Then iTab="InfoPics"
 sDir1=RequestS("sDir1","C",96) : sDir3=Replace(sDir1,"/","") ' /1/2/ - 1/2
 sDir2=RequestS("sDir2","C",96) : sDir4=Replace(sDir2,"/","") ' 
 If Len(sDir3)>=2 And Len(sDir1)-Len(sDir3)>=2 And Len(sDir4)>=2 And Len(sDir2)-Len(sDir4)>=2 Then
  SET rs=Server.CreateObject("Adodb.Recordset") 
  sql = "SELECT * FROM ["&iTab&"] WHERE InfCont LIKE '%"&sDir1&"%' "
  'Response.Write "<br>"&sql
  rs.Open sql,conn,1,3 
  if NOT rs.eof then 
   Do While NOT rs.EOF
    KeyID = rs("KeyID")
	InfCont = rs("InfCont")
	InfCont = Replace(InfCont,sDir1,sDir2)
	'InfCont = Replace(InfCont,uCase(sDir1),sDir2)
	'InfCont = Replace(InfCont,lCase(sDir1),sDir2)
	InfCont = Replace(InfCont,"'","''")
	Call rs_DoSql(conn,"UPDATE "&iTab&" SET InfCont='"&InfCont&"' WHERE KeyID='"&KeyID&"'")
	'rs("InfCont") = InfCont
	'Response.Write "<br>"&rs("InfSubj")
    rs.MoveNext()
   Loop
  End If
  rs.Close()
  SET rs=Nothing 
 End If
 Response.Write "<br>处理完成:"&Now()
next
End If

%>
<script type="text/javascript">
function goType()
{
  document.fmDir.action = document.fmDir.sDir.value;
  document.fmDir.submit();
}
</script>
</body>
</html>
