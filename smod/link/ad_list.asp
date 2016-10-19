<!--#include file="../../sadm/func1/func1.asp"-->
<!--#include file="../../sadm/func2/Func2.asp"-->
<!--#include file="../../sadm/func1/func_opt.asp"-->
<html>
<head>
<meta http-equiv="Pragma" content="no-cache">
<meta http-equiv="Expires" content="0">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title><%=Config_Name%>后台管理中心</title>
<link href="../../inc/adm_inc/adm_style.css" rel="stylesheet" type="text/css">
</head>
<body>
<!--#include file="config.asp"-->
<script src="../../sadm/func1/Func_JS.js" type="text/javascript"></script>
<%

yAct = Request("yAct") 
Page = RequestS("Page","N",1)
TP = Request("TP") 
sqlK = " KeyMod='Advert' "

If TP<>"" Then 
  sqlK=sqlK&" AND InfType LIKE '%"&TP&"%' "
End If 

If yAct="del" Then
 sql = "DELETE FROM WebAdvert WHERE KeyID='"&RequestS("ID","C",48)&"'"
 Call rs_DoSql(conn,sql)
 Msg = cID&" 条记录删除成功!"
End If

 sql = " SELECT * FROM [WebAdvert] WHERE "&sqlK
 sql =sql& " ORDER BY InfType,KeyID DESC " 
 Set rs=Server.Createobject("Adodb.Recordset")
 rs.Open Sql,conn,1,1
 rs.PageSize = 12 
If Int(Page)>rs.PageCount Or Int(Page)<1 Then
  Page = 1
End If

%>
<br>
<table width="640" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="E0E0E0">
  <tr align="center" bgcolor="E0E0E0">
    <td height="27" colspan="7" align="center" bgcolor="f8f8f8"><table width="100%" border="0" cellpadding="3" cellspacing="0">
        <tr align="center" bgcolor="#FFFFFF">
          <td width="40%" align="center" bgcolor="#FFFFFF"><a href="ad_list.asp">浮动广告</a> - <a href="ad_pic.asp">图片广告</a> - <a href="ad_text.asp">文字广告</a><br>            <strong>浮动广告管理</strong></td>
          <td width="20%" align="center" nowrap bgcolor="#FFFFFF"><a href="../type/type_list.asp?ModID=HomAdv1">类别</a> | <a href="ad_add.asp">增加</a><br>            
          &nbsp;<font color="#FF0000"><%=msg%></font></td>
          <td align="center" nowrap>&nbsp;</td>
          <form name="fSearch" action="?">
            <td align="right" nowrap><select name="TP" style="width:120px; " id="TP">
                <option value="">[不限]</option>
                <%=Get_rsOpt(conn,"SELECT TypID,TypName FROM WebTyps WHERE TypMod='HomAdvert'",TP)%>
              </select>
              <input type="submit" name="Submit" value="搜索">
              <input type="submit" name="button" id="button" value="更新" onClick="javascript:window.open('ad_view.asp?TP='+document.fSearch.TP.value)"></td>
          </form>
        </tr>
      </table></td>
  </tr>
  <tr align="center" bgcolor="E0E0E0">
    <td height="27" colspan="7" align="center" bgcolor="f8f8f8"><%= RS_Page(rs,Page,"?send=pag&TP="&TP&"",1)%></td>
  </tr>
  <tr align="center" bgcolor="E0E0E0">
    <td width="5%" height="27" align="center" nowrap>NO</td>
    <td height="27" align="center">名称</td>
    <td height="27" align="center" nowrap bgcolor="E0E0E0">组别-模版-类别</td>
    <td align="center" nowrap bgcolor="E0E0E0">大小-坐标-左右</td>
    <td align="center" nowrap bgcolor="E0E0E0">修改</td>
    <td align="center" nowrap bgcolor="E0E0E0">删除</td>
    <td height="27" align="center" nowrap bgcolor="E0E0E0">查看</td>
  </tr>
  <tr bgcolor="#333333">
    <td colspan="9" align="right" nowrap></td>
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
KeyID = rs("KeyID")
InfType = rs("InfType")
InfName = rs("InfName")
InfPara = rs("InfPara")
LogAddIP = rs("LogAddIP")
LogAUser = rs("LogAUser")
LogATime = rs("LogATime")
TypA = Split(InfType,"|")
Typ1 = rs_Val("","SELECT TypName FROM WebTyps WHERE TypID='"&TypA(0)&"'")
Typ2 = TypA(1)
If Typ2="AdvPair" Then 
 Typ2 = "对联广告"
ElseIf Typ2="AdvJJCC" Then 
 Typ2 = "警警察察"
ElseIf Typ2="Float01" Then 
 Typ2 = "漂浮一"
ElseIf Typ2="Float02" Then 
 Typ2 = "漂浮二"
ElseIf Typ2="AdvLRXX" Then 
 Typ2 = "左右浮动"
ElseIf Typ2="AdvLRQQ" Then 
 Typ2 = "QQ浮动"
End If
Typ3 = TypA(2)
InfType = Typ1&" ("&Typ3&") "&Typ2
	  %>
      
  <form name="flist" method="post" action="?">
    <tr bgcolor="<%=col%>">
      <td align="right" nowrap><input type="hidden" name="hiddenField">
        <%=i%> </td>
      <td><%=InfName%></td>
      <td align="center" nowrap bgcolor="<%=col%>"><%=InfType%></td>
      <td align="center" nowrap><%=InfPara%></td>
      <td align="center" nowrap><a href="ad_edit.asp?ID=<%=KeyID%>&TP=<%=TP%>">修改</a></td>
      <td align="center" nowrap><a onClick="Del_YN('?ID=<%=KeyID%>&Page=<%=Page%>&yAct=del&TP=<%=TP%>','确认删除?')" href="#" >删除</a></td>
      <td align="center" nowrap><a href="ad_view.asp?ID=<%=KeyID%>&Act=View" target="_blank">查看</a></td>
    </tr>
  </form>
  <%
  rs.Movenext
  If rs.Eof Then Exit For

  Next
%>
  <%  
  
  Else
  %>
  <tr align="center" bgcolor="#f4f4f4">
    <td colspan="9">无信息</td>
  </tr>
  <%
  End If
	  
	  rs.Close()
	  Set rs = Nothing
	  
	  %>
  <tr bgcolor="#999999">
    <td colspan="9" align="right"></td>
  </tr>
</table>
</body>
</html>
