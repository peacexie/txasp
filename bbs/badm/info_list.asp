<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title>后台管理中心</title>
<meta http-equiv="Pragma" content="no-cache">
<link href="../../inc/adm_inc/adm_style.css" rel="stylesheet" type="text/css">
</head>
<body>

<!--#include file="../../sadm/func1/func1.asp"-->
<!--#include file="../../sadm/func2/Func2.asp"-->
<!--#include file="../../sadm/func1/func_file.asp"-->
<!--#include file="../../sadm/func1/func_opt.asp"-->
<!--#include file="config.asp"-->
<!--#include file="conpub.asp"-->
<script src="../../sadm/func1/Func_JS.js" type="text/javascript"></script>

<%

KeyID   = RequestS("KeyID",3,48)
KeyID_Back = KeyID
TP = RequestS("TP",3,48)	
KW = RequestS("KW",3,12)
Page = RequestS("Page","N",1)
send = Request("send")
sqlK = ""
	if KW&"" <> "" then
	  sqlK = sqlK & " AND ( InfSubj LIKE '%"&KW&"%' "
	  sqlK = sqlK & " OR InfCont LIKE '%"&KW&"%' "
	  sqlK = sqlK & " ) " 
	end if
	  if TP&"" <> "" then
	    sqlK = sqlK & " AND ( KeyMod='"&TP&"' ) " 
	  end if

If send = "del" Then

	sql = "DELETE FROM BPKTitle WHERE KeyID='"&KeyID&"'"
	Call rs_DoSql(conn,sql)
	sql = "DELETE FROM BPKView WHERE KeyMod='"&KeyID&"'"
	Call rs_DoSql(conn,sql)
	sql = "DELETE FROM BPKVote WHERE KeyMod='"&KeyID&"'"
	Call rs_DoSql(conn,sql)
	Msg = "删除成功!"
Else

End If

    sql = " SELECT * FROM [BPKTitle] "
	sql =sql& " WHERE KeyMod='"&ModID&"'  "&sqlK
	sql =sql& " ORDER BY InfStart DESC" 
   Set rs=Server.Createobject("Adodb.Recordset")
   rs.Open Sql,conn,1,1
   rs.PageSize = 15 '2 '
if int(Page)>rs.PageCount or int(Page)<1 Then
  Page = 1
End If
%>
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="E0E0E0">
  <tr align="center" bgcolor="E0E0E0">
    <td height="27" colspan="8" align="center" bgcolor="f8f8f8"><table width="100%" border="0" cellpadding="3" cellspacing="0">
        <tr align="center" bgcolor="#FFFFFF">
          <td width="30%" align="center" bgcolor="#FFFFFF"><strong>辩论管理 &gt;&gt;</strong></td>
          <td width="30%" align="center" bgcolor="#FFFFFF">&nbsp;&nbsp;&nbsp;<font color="#FF0000"><%=msg%></font></td>
          <form name="fm01" method="post" action="?">
            <td align="right" nowrap><input name="send" type="hidden" id="send" value="sch">
              论坛类别:
              <select class=form id=select name=TP style="width:120px; ">
                <option value="">[所有类别]</option>
                
              </select>
              关键字:
              <input name="KW" type="text" id="KW" value="<%=KW%>" size="10">
              <input type="submit" name="Submit" value="搜索">
            </td>
          </form>
        </tr>
      </table></td>
  </tr>
  <tr bgcolor="#999999">
    <td colspan="8" align="right"></td>
  </tr>
  <tr align="center" bgcolor="E0E0E0">
    <td height="27" colspan="8" align="center" bgcolor="f8f8f8"><%if not rs.eof then
              Response.Write RS_Page(rs,Page,"?send=pag&TP="&TP&"&KW="&KW&"",1)
              end if%></td>
  </tr>
  <tr align="center" bgcolor="E0E0E0">
    <td height="27" align="center">NO</td>
    <td height="27" align="center">辩论标题(点击标题查看帖子)</td>
    <td height="27" align="center" nowrap>群组类别</td>
    <td align="center">&nbsp;</td>
    <td align="center">起止日期</td>
    <td align="center" nowrap>发起人</td>
    <td height="27" align="center">修改</td>
    <td height="27" align="center">删除</td>
  </tr>
  <tr bgcolor="#333333">
    <td colspan="10" align="right"></td>
  </tr>
  <%
  if not rs.eof then
  rs.AbsolutePage = Page
  for i = 1 to rs.PageSize
		if i mod 2 = 1 then
		  col = "#ffffff"
		else
		  col = "#F8F8F8"
		end if
KeyID = rs("KeyID")
InfType = rs("InfType")
TypName = rs_Val("","SELECT TypName FROM WebTyps WHERE TypID='"&InfType&"'")
InfSubj = rs("InfSubj")
InfCont = rs("InfCont")
InfOrg = rs("InfOrg")
InfOUrl = rs("InfOUrl")
InfMemb = rs("InfMemb")
InfStart = rs("InfStart")
InfEnd = rs("InfEnd")
InfView1 = rs("InfView1")
InfView2 = rs("InfView2")
InfVote1 = rs("InfVote1")
InfVote2 = rs("InfVote2")
SetRead = rs("SetRead")
SetHot = rs("SetHot")
SetShow = rs("SetShow")
ImgName = rs("ImgName")
LogAddIP = rs("LogAddIP")
LogAUser = rs("LogAUser")
LogATime = rs("LogATime")
	  %>

    <tr bgcolor="<%=col%>">
      <td align="right"><%=i%></td>
      <td> <a 
			  href="/pkview.asp?KeyID=<%=KeyID%>" target="_blank"><%=InfSubj%></a>
      </td>
      <td align="center" nowrap><%=TypName%></td>
      <td nowrap><br></td>
      <td nowrap><%=InfStart%>~<%=InfEnd%><br>        </td>
      <td align="center"><%=InfMemb%></td>
      <td align="center">        <a href="info_edit.asp?KeyID=<%=KeyID%>&TP=<%=TP%>&KW=<%=KW%>&Page=<%=Page%>">编辑</a> </td>
      <td align="center"><input type="button" name="Button" value="删除" 
			  onClick="Del_YN('?send=del&KeyID=<%=KeyID%>&ImgName=<%=ImgName%>','确认删除?')">
      </td>
    </tr>

  <%
  rs.movenext
  if rs.eof then exit for
  next
  else
  %>
  <tr align="center" bgcolor="#f4f4f4">
    <td colspan="10">无论坛</td>
  </tr>
  <%
  end if
	  
	  rs.close()
	  set rs = nothing
	  
	  %>
  <tr bgcolor="#999999">
    <td colspan="8" align="right"></td>
  </tr>
</table>
</body>
</html>
