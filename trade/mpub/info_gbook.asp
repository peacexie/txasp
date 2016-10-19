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

MD = RequestS("ModID",3,24) 
yAct = Request("yAct") 
Page = RequestS("Page","N",1)
KW = RequestS("KW",3,24)
TP = RequestS("TP",3,24)

sqlK = " WHERE KeyUser='"&Session("MemID")&"' "
If MD="TraG124" Then
  MDName = "在线留言"
  sqlK = sqlK & " AND KeyMod='"&MD&"' "
ElseIf MD="TraR124" Then
  MDName = "信息评论"
  sqlK = sqlK & " AND KeyMod='"&MD&"' "
ElseIf MD="TraA124" Then
  MDName = "申请职位"
  sqlK = sqlK & " AND KeyMod='"&MD&"' "
ElseIf MD="TraO124" Then
  MDName = "订购产品"
  sqlK = sqlK & " AND KeyMod='"&MD&"' "
Else
  MDName = "(所有)"
  'sqlK = sqlK & " "
End If

'If KW<>"" Then
 If TP="kUser" Then
  sqlK = sqlK & " AND ( KeyUser LIKE '"&KW&"%' ) " 
 ElseIf TP="kAdd" Then
  sqlK = sqlK & " AND ( LogAUser LIKE '"&KW&"%' ) " 
 ElseIf TP<>"" Then
  sqlK = sqlK & " AND (InfType='"&TP&"') " 
 Else
  sqlK = sqlK & " AND (InfSubj LIKE '%"&KW&"%') " 
 End If
'End If

cID = 0
sID = ""
If yAct="SetShow" Then
  yVal = RequestS("yVal","C",24)
  For iy = 1 To Request.Form("yID").Count
    iID = RequestSafe(Request.Form("yID").item(iy),3,96) 
    If iID<>"" Then
	  sID = sID&"'"&iID&"',"
	  cID = cID + 1
	End If
  Next
  sID = Replace(sID&"''",",''","")
  If sID<>"" Then
    sql = " UPDATE TradeGbook SET "&yAct&"='"&yVal&"' "
    sql = sql& " WHERE KeyID IN("&sID&")" ':Response.Write sql : Response.End()
	Call rs_DoSql(conn,sql)
  End If
    Msg = cID&" 条记录 设置成功!"
Elseif yAct="del_sel" then
  If Chk_URL3(Config_Path&"trade/mpub/info_gbook.asp")="eUrl" Then
    Response.End()
  End If
  For iy = 1 To Request.Form("yID").Count
    iID = RequestSafe(Request.Form("yID").item(iy),3,96) 
    If iID<>"" Then
	  cID = cID + 1
	  Call rs_DoSql(conn,"DELETE FROM TradeGbook WHERE KeyID='"&iID&"'")
	End If
  Next
    Msg = cID&" 条记录删除成功!"
ElseIf yAct="del_now" Then
 If KW&TP<>"" Then
  If Chk_URL3(Config_Path&"trade/mpub/info_gbook.asp")="eUrl" Then
    Response.End()
  End If
   	sqlK = " "&sqlK
	cID = rs_Count(conn," TradeGbook"&sqlK)
	Call rs_DoSql(conn,"DELETE FROM TradeGbook "&sqlK)
    Msg = cID&" 条记录删除成功!"
 End If
  sqlK="" : TP="" : KW=""
  Page=1 ' 清楚条件,重设第一页
End If

    Set rs=Server.Createobject("Adodb.Recordset")
	sql = " SELECT * FROM [TradeGbook] "&sqlK&sqlU
	sql =sql& " ORDER BY KeyID DESC"  ':Response.Write sql&SetTop
   rs.Open Sql,conn,1,1
   rs.PageSize = 15 
If Int(Page)>rs.PageCount Or Int(Page)<1 Then
  Page = 1
End If

%>
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="E0E0E0">
  <tr align="center" bgcolor="E0E0E0">
    <td height="27" colspan="10" align="center" bgcolor="f8f8f8"><table width="100%" border="0" cellpadding="3" cellspacing="0">
        <tr align="center" bgcolor="#FFFFFF">
          <td colspan="2" align="center" bgcolor="#FFFFFF">
          
| 
<a href='?ModID='>(All)</a> | 
<a href='?ModID=TraG124'>留言</a> | 
<a href='?ModID=TraR124'>评论</a> | 
<a href='?ModID=TraA124'>应聘</a> | 
<a href='?ModID=TraO124'>订购</a> | 
          
          </td>
          <td align="right" nowrap>&nbsp;</td>
        </tr>
        <tr align="center" bgcolor="#FFFFFF">
          <td width="30%" align="center" bgcolor="#FFFFFF"><strong>[<%=MDName%>]</strong></td>
          <td width="30%" align="center" bgcolor="#FFFFFF">&nbsp;&nbsp;&nbsp;<font color="#FF0000"><%=msg%></font></td>
          <form name="fm01" method="post" action="?">
            <td align="right" nowrap><input name="send" type="hidden" id="send" value="sch">
              <input name="ModID" type="hidden" id="ModID" value="<%=MD%>">
              <select name="TP" style="width:120px; " id="TP">
                <option value="">[不限]</option>
                <option value="kAdd">[uid]</option>
                <%=Get_rsOpt(conn,"SELECT TypID,TypName FROM WebType WHERE TypMod='BisU124'",TP)%>
              </select>              
              <input name="KW" type="text" id="KW" value="<%=KW%>" size="12">
              <input type="submit" name="Submit" value="搜索">            </td>
          </form>
        </tr>
      </table></td>
  </tr>
  <tr align="center" bgcolor="E0E0E0">
    <td height="27" colspan="10" align="center" bgcolor="f8f8f8"><%= RS_Page(rs,Page,"?send=pag&TP="&TP&"&KW="&KW&"&ModID="&ModID&"",1)%></td>
  </tr>
  <tr align="center" bgcolor="E0E0E0">
    <td width="5%" height="27" align="center" nowrap>NO</td>
    <td height="27" align="center">主题</td>
    <td width="5%" align="center" nowrap>&nbsp;</td>
    <td width="5%" align="center" nowrap>姓名</td>
    <td width="5%" height="27" align="center" nowrap>类别</td>
    <td width="5%" align="center" nowrap>发布</td>
    <td width="5%" align="center" nowrap>显</td>
    <td width="5%" align="center" nowrap>回复</td>
    <td width="5%" height="27" align="center" nowrap>修改</td>
    <td width="5%" align="center" nowrap>uid</td>
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
KeyUser = rs("KeyUser")
InfType = rs("InfType")
TypName = rs_Val("","SELECT TypName FROM WebType WHERE TypID='"&InfType&"'")
If TypName="" Then
  TypName = InfType
End If
InfSubj = Show_Text(rs("InfSubj")) 'rs("InfCont")
LnkName = rs("LnkName")
LnkEmail = rs("LnkEmail")
SetRead = rs("SetRead")
SetShow = rs("SetShow")
LogAddIP = rs("LogAddIP")
LogAUser = rs("LogAUser")
LogATime = rs("LogATime")

xxxCont =  Show_Text(rs("InfCont"))
xxxReply =  Show_Text(rs("InfReply"))
If KeyMod="TraG124" Then
  If xxxReply="" Then 
    xxxReply="<span style='color:#F00'>未回复</span>"
  Else
    xxxReply="<span style='color:#999'>已回复</span>"
  End If
Else
  xxyReply="<span style='color:#00F'>(原文)</span>"
  xxxReply = "<a href='../iview.asp?KeyID="&xxxReply&"' target='_blank'>"&xxyReply&"</a>"
End If

	  %>
    <tr bgcolor="<%=col%>">
      <td align="right" nowrap><%=i%>
        <input name="yID" type="checkbox" id="yID" value="<%=KeyID%>"></td>
      <td><a href="../madm/info_gbview.asp?KeyID=<%=KeyID%>" target="_blank"><%=InfSubj%></a> </td>
      <td align="center" nowrap>&nbsp;</td>
      <td align="center" nowrap><%=LnkName%></td>
      <td align="center" nowrap><%=TypName%></td>
      <td align="center" nowrap><%=LogATime%></td>
      <td align="center" nowrap><%=SetShow%></td>
      <td align="center" nowrap><%=xxxReply%></td>
      <td align="center" nowrap><a href="../madm/info_gbview.asp?KeyID=<%=KeyID%>&Act=Edit" target="_blank">修改</a></td>
      <td align="center" nowrap><%=LogAUser%></td>
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
        <input name="KW" type="hidden" id="KW" value="<%=KW%>">
        <input name="ModID" type="hidden" id="ModID" value="<%=ModID%>">
        <input name="PrmFlag" type="hidden" id="PrmFlag" value="<%=PrmFlag%>"></td>
      <td colspan="4" nowrap>
        <%If (ModID="MemB524" AND PrmFlag="(Inn)")OR(ModID="MemB224" AND PrmFlag="(Mem)") Then%>
        <%Else%>
        <select name="yAct" id="yAct" >
          <option value="del_sel">删除.所选</option>
          <option value="del_now">删除.当前</option>
          <option value="SetShow">设置_显示</option>
        </select>
        <select name="yVal" id="yVal">
		  <option value="Y" selected>Y</option>
          <option value="N">N</option>
		  <option value="X">X</option>
        </select> 
        <%End If%>
      </td>
      <td colspan="4" nowrap><input type="submit" name="Submit" value="执行">
&nbsp; </td>
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
<%
'Response.Write Timer()-Tim1
%>
</body>
</html>
