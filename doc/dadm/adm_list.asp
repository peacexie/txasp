<!--#include file="_config.asp"-->
<html>
<head>
<meta http-equiv="Pragma" content="no-cache">
<meta http-equiv="Expires" content="0">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title>公文管理</title>
<link href="../../inc/adm_inc/adm_style.css" rel="stylesheet" type="text/css">
<script src="../sadm/func1/Func_JS.js" type="text/javascript"></script>
</head>
<body>

<%
Set rs=Server.Createobject("Adodb.Recordset")

yAct = Request("yAct") 
Page = RequestS("Page","N",1)
KW = RequestS("KW",3,24)
TP = RequestS("TP",3,255)
	If KW&"" <> "" Then
	  sqlK = sqlK & " AND ( InfSubj LIKE '%"&KW&"%' "
	  'sqlK = sqlK & " OR InfKey LIKE '%"&KW&"%' "
	  sqlK = sqlK & " OR LogAUser LIKE '%"&KW&"%' "
	  sqlK = sqlK & " ) " 
	End If
	If TP&"" <> "" then
	  sqlK = sqlK & " AND (InfType LIKE '"&TP&"%') " 
	End If

cID = 0
sID = ""
  yVal = RequestS("yVal","C",24)
  For iy = 1 To Request.Form("yID").Count
    iID = RequestSafe(Request.Form("yID").item(iy),3,96) 
    If iID<>"" Then
	  sID = sID&"'"&iID&"',"
	  cID = cID + 1
	End If
  Next
  
If yAct="SetShow" Or yAct="SetHot" Or yAct="SetUBB" Or yAct="SetTop" Then
  If yAct="SetTop" Then
    yVal=RequestSafe(yVal,"N",888)
  Else
    yVal=Left(yVal,1)
  End If
  If sID<>"" Then
    sID = Replace(sID&"''",",''","")
	sql = " UPDATE DocsNews SET "&yAct&"='"&yVal&"' "
    sql = sql& " WHERE KeyID IN("&sID&")" 
	Call rs_DoSql(conn,sql)
  End If
    Msg = cID&" 条记录 设置成功!"
Elseif yAct="del_sel" then
	Call del_sfDir(ModTab,sID) 'Call rs_DFile("DocsNews",sID,"")
    Msg = cID&" 条记录删除成功!"
Elseif yAct="UpdList" then

    sql45 = "SELECT SysID,SysName FROM AdmSyst WHERE SysType='Inner' ORDER BY SysTop, SysID "
    rs.Open sql45,conn,1,1 
    s1="" : s2="" :s3="" : s4=""
    Do While Not rs.EOF 
      sysID = rs("SysID")
      sysNM = rs("SysName")
      s1 = s1&vbcrlf&"<td id='id_"&sysID&"'>"&sysNM&"</td>"
      '/////////////////////
      sql = "SELECT UsrID,UsrName FROM [AdmUser"&Adm_aUser&"] "
      sql = sql& " WHERE UsrType IN('"&sysID&"') ORDER BY UsrType,UsrID"
      Set rs3=Server.Createobject("Adodb.Recordset")
      rs3.Open Sql,conn,1,1
      s3 = ""
	  Do While NOT rs3.EOF 
        UsrName = rs3("UsrName")
        UsrID = rs3("UsrID")
          sc = ""
        If inStr(xPerm,""&UsrID&";")>0 Then
          sc = "checked"
        End If
        s3 = s3&vbcrlf&"  <li class='mInnID'><a href='?TG="&sysID&"&TU="&UsrID&"'>"&UsrName&"</a></li>"
        rs3.MoveNext
        i = i + 1
      Loop
      rs3.Close()
      Set rs3 = Nothing
	  s3 = vbcrlf&"<div id='div_"&sysID&"' style='visibility:xhidden;'>"&s3&vbcrlf&"</div>"
      '/////////////////////
	  s2 = s2&s3
	  s4 = s4&";"&sysID 
    rs.MoveNext
    Loop
    rs.Close()

    Call File_Add2("../../upfile/sys/doc/list_depart.asp",s1,"UTF-8")
	Call File_Add2("../../upfile/sys/doc/list_user.asp",s2,"UTF-8")
	s4 = "var sDeps = 'All__User"&s4&"';"
	Call File_Add2("../../upfile/sys/doc/list_depart.js",s4,"UTF-8")

    Response.Write vbcrlf&vbcrlf&"<hr>"&s1&"<hr>"&s2&"<hr>"&s4
    '//////////////////////////////////

    Response.Write " <hr>刷新成功!"
	Response.End()
End If
    sql = " SELECT DocsNews.* FROM [DocsNews] "
	sql =sql& " WHERE KeyMod='"&ModID&"' "&sqlK
	sql =sql& " ORDER BY LogATime DESC " 
   rs.Open Sql,conn,1,1
   rs.PageSize = 15 
If Int(Page)>rs.PageCount Or Int(Page)<1 Then
  Page = 1
End If

MDName = rs_Val("","SELECT SysName FROM AdmSyst WHERE SysID='"&ModID&"'")
'Response.Write sql
%>
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="E0E0E0">
  <tr align="center" bgcolor="E0E0E0">
    <td height="27" colspan="10" align="center" bgcolor="f8f8f8"><table width="100%" border="0" cellpadding="3" cellspacing="0">
        <tr align="center" bgcolor="#FFFFFF">
          <td width="30%" align="center" bgcolor="#FFFFFF"><strong>[<%=MDName%>]管理 </strong> | <a href="../../ext/login.asp" target="_blank">登陆&gt;&gt;</a></td>
          <td width="30%" align="center" bgcolor="#FFFFFF">&nbsp;&nbsp;&nbsp;<font color="#FF0000"><%=msg%></font></td>
          <form name="fm01" method="post" action="?">
            <td align="right" nowrap><a href="?yAct=UpdList" target="_blank">刷新</a>
              <input name="send" type="hidden" id="send" value="sch">
              <select class=form id=select name=TP >
			  <%=Get_rsTree(conn,"SELECT * FROM WebTyps WHERE TypMod='"&ModID&"' ORDER BY TypLayer ",InfType,"Lay") %>
              </select>
              <input name="KW" type="text" id="KW" value="<%=KW%>" size="12">
              <input type="submit" name="Submit" value="搜索">
            </td>
          </form>
        </tr>
      </table></td>
  </tr>
  <tr align="center" bgcolor="E0E0E0">
    <td height="27" colspan="10" align="center" bgcolor="f8f8f8"><%= RS_Page(rs,Page,"?send=pag&TP="&TP&"&KW="&KW&"",1)%></td>
  </tr>
  <tr align="center" bgcolor="E0E0E0">
    <td width="5%" height="27" align="center" nowrap>NO</td>
    <td width="40%" height="27" align="center">主题</td>
    <td height="27" align="center" nowrap>类别</td>
    <td nowrap>组别</td>
    <td nowrap>发布人</td>
    <td align="center" nowrap>时间</td>
    <td align="center" nowrap>阅读</td>
    <td align="center" nowrap>&nbsp;</td>
    <td height="27" align="center" nowrap>修改</td>
  <td align="center" nowrap>&nbsp;</td>
  </tr>
  <tr bgcolor="#333333">
    <td colspan="12" align="right" nowrap></td>
  </tr>
  <form name="flist" method="post" action="?">
    <%
  If not rs.eof then
  rs.AbsolutePage = Page
  For i = 1 To rs.PageSize
		If i mod 2 = 1 Then
		  col = "#ffffff"
		Else
		  col = "#F8F8F8"
		End If
KeyID = rs("KeyID")
KeyMod = rs("KeyMod")
KeyCode = rs("KeyCode")
InfType = rs("InfType")
InfTyp2 = rs("InfTyp2")
InfSubj = rs("InfSubj")
SetSubj = rs("SetSubj")
SetRead = rs("SetRead")
SetSubj = rs("SetSubj")
SetHot = rs("SetHot")
SetTop = rs("SetTop")
SetShow = rs("SetShow")
ImgName = rs("ImgName")
LogATime = rs("LogATime")
LogAUser = rs("LogAUser")

InfSubj = Show_sTitle(InfSubj,SetSubj)
ImgFlag = ""
If ImgName<>"" Then
  ImgFlag = "<img src='../dimg/tool/attfile.gif' border=0 align='absmiddle'>" 
End If

TypName = rs_Val("","SELECT TypName FROM WebTyps WHERE TypLayer='"&InfType&"'")
If TypName="" Then
  TypName = "<font color=gray>"&MDName&"</font>"
End If

InfTyp2 = rs_Val("","SELECT SysName FROM AdmSyst WHERE SysID='"&InfTyp2&"'")

LogAUBak = LogAUser
LogAUser = rs_Val("","SELECT UsrName FROM AdmUser"&Adm_aUser&" WHERE UsrID='"&LogAUser&"'")&"("&LogAUser&")"
'If LogAUser="" Then
  'LogAUser = rs("LogAUser")
'End If

fRead = rs_Val("","SELECT KeyN FROM DocsLogs WHERE KeyID='"&KeyID&"'")
If fRead&""="" Then
fRead = "<span style='color:red'>未读</span>"
End If

	  %>
    <tr bgcolor="<%=col%>">
      <td align="right" nowrap><%=i%>
      <input name="yID" type="checkbox" id="yID" value="<%=KeyID%>"></td>
      <td>                      <a href="../index.asp?Act=AdmLogin&InnID=<%=LogAUBak%>&ID=<%=KeyID%>&Dir=doc_view.asp" target="_blank"><%=InfSubj%></a><%=ImgFlag%>     </td>
      <td nowrap><%=TypName%><%=""&TypNam2%></td>
      <td align="center" nowrap><%=InfTyp2%></td>
      <td align="center" nowrap><a href="../index.asp?Act=AdmLogin&InnID=<%=LogAUBak%>" target="_blank"><%=LogAUser%></a></td>
      <td align="center" nowrap><%=LogATime%></td>
      <td align="center" nowrap><%=fRead%></td>
      <td align="center" nowrap>&nbsp;</td>
      <td align="center" nowrap><a href="../index.asp?Act=AdmLogin&InnID=<%=LogAUBak%>&ID=<%=KeyID%>&Dir=info_edit.asp" target="_blank">修改</a></td>
    <td align="center" nowrap>&nbsp;</td>
    </tr>
    <%
  rs.Movenext
  If rs.Eof Then Exit For

  Next
%>
    <tr bgcolor="E0E0E0">
      <td height="21" align="right" nowrap>
        <span id="yFlag" style="visibility:hidden ">N</span>
      <input name="yAll" type="checkbox" id="yAll" onClick="ySel()"></td>
      <td>全选
        <input name="Page" type="hidden" id="Page" value="<%=Page%>">
        <input name="TP" type="hidden" id="TP" value="<%=TP%>">
        <input name="TP2" type="hidden" id="TP2" value="<%=TP2%>">
        <input name="KW" type="hidden" id="KW" value="<%=KW%>">
</td>
      <td nowrap>&nbsp;</td>
      <td align="center" nowrap>&nbsp;</td>
      <td colspan="3" align="right" nowrap><select name="yAct" id="yAct" >
          <option value="del_sel">删除.所选</option>
          <!--
          <option value="SetShow">设置_显示</option>
          <option value="SetHot" selected>设置_推荐</option>
		  <option value="SetTop">设置_顺序</option>
          -->
        </select>
        <select name="yVal" id="yVal">
		  <option value="Y">Y</option>
          <option value="N">N</option>
		  <option value="X">X</option>
        </select>
      </td>
      <td colspan="3" align="left" nowrap><input type="submit" name="Submit" value="执行">      </td>
    </tr>
    <%  
  
  Else
  %>
    <tr align="center" bgcolor="#f4f4f4">
      <td colspan="12">无信息</td>
    </tr>
    <%
  End If
	  
	  rs.Close()
	  Set rs = Nothing
	  
	  %>
    <tr bgcolor="#999999">
      <td colspan="12" align="right"></td>
    </tr>
  </form>
</table>
<script type="text/javascript">
function ySel()
{
   var vFlag = yFlag.innerText;
   if (vFlag=="N"){
   yFlag.innerText = "Y";
   for(var i=0;i<document.flist.yID.length;i++)
   {document.flist.yID.item(i).checked=true;}
   }else{
   yFlag.innerText = "N";
   for(var i=0;i<document.flist.yID.length;i++)
   {document.flist.yID.item(i).checked=false;}
   }
}  

</script>

</body>
</html>
