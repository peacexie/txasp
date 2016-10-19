<!--#include file="config.asp"-->
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

ModID = "UserCorp"

yAct = Request("yAct") 
Page = RequestS("Page","N",1)
KW = RequestS("KW",3,24)
TP = RequestS("TP",3,255)

	If KW&"" <> "" Then
	  sqlK = sqlK & " AND ( InfSubj LIKE '%"&KW&"%' "
	  sqlK = sqlK & " OR InfKey LIKE '%"&KW&"%' "
	  sqlK = sqlK & " OR InfCont LIKE '%"&KW&"%' "
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
  
If yAct="SetShow" Or yAct="SetHot" Or yAct="SetTop" Then
  If yAct="SetTop" Then
    yVal=RequestSafe(yVal,"N",888)
  Else
    yVal=Left(yVal,1)
  End If
  If sID<>"" Then
    sID = Replace(sID&"''",",''","")
	sql = " UPDATE TradeCorp SET "&yAct&"='"&yVal&"' "
    sql = sql& " WHERE KeyID IN("&sID&") " 
	Call rs_DoSql(conn,sql)
  End If
    Msg = cID&" 条记录 设置成功!"
Elseif yAct="del_sel" then
	Call rs_DFile("TradeCorp",sID,"")
    Msg = cID&" 条记录删除成功!"
End If
    sql = " SELECT TradeCorp.* FROM [TradeCorp] "
	sql =sql& " WHERE KeyMod IN('Corp','Privy','Gov','Org') "&sqlK
	sql =sql& " ORDER BY KeyID DESC" 
   Set rs=Server.Createobject("Adodb.Recordset")
   rs.Open Sql,conn,1,1
   rs.PageSize = 15 
If Int(Page)>rs.PageCount Or Int(Page)<1 Then
  Page = 1
End If

%>
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="E0E0E0">
  <tr align="center" bgcolor="E0E0E0">
    <td height="27" colspan="11" align="center" bgcolor="f8f8f8"><table width="100%" border="0" cellpadding="3" cellspacing="0">
        <tr align="center" bgcolor="#FFFFFF">
          <td width="30%" align="center" bgcolor="#FFFFFF"><strong>[企业资料]管理 </strong> | <a href="info_add.asp?TPU=<%=IDPerm%>">新增&gt;&gt;</a></td>
          <td width="30%" align="center" bgcolor="#FFFFFF">&nbsp;&nbsp;&nbsp;<font color="#FF0000"><%=msg%></font></td>
          <form name="fm01" method="post" action="?">
            <td align="right" nowrap><input name="send" type="hidden" id="send" value="sch">
              <select class=form id=select name=TP >
			  <option value="">[不限]</option>
			  <%=Get_rsOpt(conn,"SELECT TypID,TypName FROM TradeType WHERE TypMod='"&ModID&"' ORDER BY TypID ",TP) %>
              </select>
              <input name="KW" type="text" id="KW" value="<%=KW%>" size="12">
              <input type="submit" name="Submit" value="搜索">            </td>
          </form>
        </tr>
      </table></td>
  </tr>
  <tr align="center" bgcolor="E0E0E0">
    <td height="27" colspan="11" align="center" bgcolor="f8f8f8"><%= RS_Page(rs,Page,"?send=pag&TP="&TP&"&KW="&KW&"",1)%></td>
  </tr>
  <tr align="center" bgcolor="E0E0E0">
    <td width="5%" height="27" align="center" nowrap>NO</td>
    <td width="40%" height="27" align="center">主题</td>
    <td height="27" align="center" nowrap>类别</td>
    <td align="center" nowrap>行业</td>
    <td align="center" nowrap>发布</td>
    <td align="center" nowrap>User ID</td>
    <td align="center" nowrap>显序荐</td>
    <td align="center" nowrap>阅读</td>
    <td align="center" nowrap>&nbsp;</td>
    <td height="27" align="center" nowrap>修改</td>
  <td align="center" nowrap>&nbsp;</td>
  </tr>
  <tr bgcolor="#333333">
    <td colspan="13" align="right" nowrap></td>
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
TypName = rs_Val("","SELECT TypName FROM WebType WHERE TypID='"&InfType&"'")
InfTyp2 = rs("InfTyp2")
TypNam2 = rs_Val("","SELECT TypName FROM WebType WHERE TypID='"&InfTyp2&"'")
InfSubj = rs("InfSubj")
SetSubj = rs("SetSubj")
SetRead = rs("SetRead")
SetHot = rs("SetHot")
SetTop = rs("SetTop")
SetShow = rs("SetShow")
ImgName = rs("ImgName")
LogATime = rs("LogATime")
LogAUser = rs("LogAUser")
InfSubj = Show_sTitle(InfSubj,SetSubj)
ImgFlag = ""
If ImgName<>"" Then
  ImgFlag = "<img src='../../img/tool/attfile.gif' border=0 align='absmiddle'>" 
End If
	  %>
    <tr bgcolor="<%=col%>">
      <td align="right" nowrap><%=i%>
      <input 
			name="yID" type="checkbox" id="yID" value="<%=KeyID%>"></td>
      <td><a href="../mpub/adm_view.asp?ID=<%=KeyID%>" target="_blank"><%=InfSubj%></a><%=ImgFlag%>     </td>
      <td nowrap><%=TypName%></td>
      <td align="center" nowrap><%=""&TypNam2%></td>
      <td align="center" nowrap><%=LogATime%></td>
      <td align="center" nowrap><%=LogAUser%></td>
      <td align="center" nowrap><%=SetShow%><%=SetTop%><%=SetHot%></td>
      <td align="center" nowrap><%=SetRead%></td>
      <td align="center" nowrap>&nbsp;</td>
      <td align="center" nowrap><a href="info_edit.asp?ID=<%=KeyID%>&PG=<%=Page%>&KW=<%=KW%>&TP=<%=TP%>">修改</a></td>
    <td align="center" nowrap>&nbsp;</td>
    </tr>
    <%
  rs.Movenext
  If rs.Eof Then Exit For

  Next
%>
    <tr bgcolor="E0E0E0">
      <td height="21" align="right" nowrap>
        <span id="yFlag" style="visibility:hidden ">N</span>全选
      <input name="yAll" type="checkbox" id="yAll" onClick="ySel()"></td>
      <td><input name="Page" type="hidden" id="Page" value="<%=Page%>">
        <input name="TP" type="hidden" id="TP" value="<%=TP%>">
        <input name="TP2" type="hidden" id="TP2" value="<%=TP2%>">
        <input name="KW" type="hidden" id="KW" value="<%=KW%>"></td>
      <td nowrap>&nbsp;</td>
      <td align="center" nowrap>&nbsp;</td>
      <td align="center" nowrap><a href="info_top.asp?yTab=TradeCorp&ModID=<%=ModID%>">高级设置</a></td>
      <td colspan="3" align="right" nowrap><select name="yAct" id="yAct" >
          <option value="del_sel">删除.所选</option>
          <option value="SetShow">设置_显示</option>
          <option value="SetHot" selected>设置_推荐</option>
		  <option value="SetTop">设置_顺序</option>
        </select>
        <select name="yVal" id="yVal">
		  <option value="Y">Y</option>
          <option value="N">N</option>
		  <option value="X">X</option>
		  <%For i=0 To 9%>
		  <option value="<%=i%>" ><%=i%></option>
		  <%Next%>
          <option value="888">888</option>
        </select>      </td>
      <td colspan="3" align="left" nowrap><input type="submit" name="Submit" value="执行">      </td>
    </tr>
    <%  
  
  Else
  %>
    <tr align="center" bgcolor="#f4f4f4">
      <td colspan="13">无信息</td>
    </tr>
    <%
  End If
	  
	  rs.Close()
	  Set rs = Nothing
	  
	  %>
    <tr bgcolor="#999999">
      <td colspan="13" align="right"></td>
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
