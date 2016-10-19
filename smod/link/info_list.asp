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
KW = RequestS("KW",3,24)
TP = RequestS("TP",3,24)

	If KW&"" <> "" Then
	  sqlK = sqlK & " AND ( InfSubj LIKE '%"&KW&"%' "
	  sqlK = sqlK & " OR InfURL LIKE '%"&KW&"%' "
	  sqlK = sqlK & " ) " 
	End If
	If TP&"" <> "" then
	  sqlK = sqlK & " AND (InfType LIKE '%"&TP&"%') " 
	  ordK = "SetTop,KeyID DESC"
	Else
	  ordK = "KeyID DESC"
	End If

cID = 0
sID = ""
If yAct="SetTop" Then
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
    sql = " UPDATE GboLink SET "&yAct&"='"&yVal&"' "
    sql = sql& " WHERE KeyID IN("&sID&")" ':Response.Write sql : Response.End()
	Call rs_DoSql(conn,sql)
  End If
    Msg = cID&" 条记录 设置成功!"
Elseif yAct="del_sel" then
  For iy = 1 To Request.Form("yID").Count
    iID = RequestSafe(Request.Form("yID").item(iy),3,96) 
    If iID<>"" Then
	  sID = sID&"'"&iID&"',"
	  cID = cID + 1
	End If
  Next
	sID = Replace(sID&"''",",''","")
	sql = "DELETE FROM GboLink WHERE 1=1 AND KeyID IN("&sID&")"
	Call rs_DoSql(conn,sql)
    Msg = cID&" 条记录删除成功!"
ElseIf yAct="del_now" Then
 If KW&TP<>"" Then
	cID = rs_Count(conn,"GboLink WHERE 1=1 "&sqlK)
	sql = "DELETE FROM GboLink WHERE LogAUser='"&Session("MemID")&"' "&sqlK
	Call rs_DoSql(conn,sql)
    Msg = cID&" 条记录删除成功!"
 End If
  sqlK="" : TP="" : KW=""
  Page=1 ' 清楚条件,重设第一页
End If
    ordK = "SetTop,KeyID DESC"
	sql = " SELECT * FROM [GboLink] "
	sql =sql& " WHERE 1=1 "&sqlK 
	sql =sql& " ORDER BY "&ordK&""  ':Response.Write sqlSetTop,
   Set rs=Server.Createobject("Adodb.Recordset")
   rs.Open Sql,conn,1,1
   rs.PageSize = 15 
If Int(Page)>rs.PageCount Or Int(Page)<1 Then
  Page = 1
End If

%>
<table width="640" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="E0E0E0">
  <tr align="center" bgcolor="E0E0E0">
    <td height="27" colspan="7" align="center" bgcolor="f8f8f8"><table 
	width="100%" border="0" cellpadding="3" cellspacing="0">
	<form name="fm01" method="post" action="?">
        <tr align="center" bgcolor="#FFFFFF">
          <td width="30%" rowspan="2" align="center" bgcolor="#FFFFFF"><strong>连接管理 | </strong>
          <a href="info_add.asp">《新增》</a></td>
          <td rowspan="2" align="center" bgcolor="#FFFFFF">&nbsp;&nbsp;&nbsp;</td>
          
            <td align="right" nowrap><font color="#FF0000"><%=msg%></font> &nbsp;<a href="../type/type_list.asp?ModID=HomLnk1">类别</a></td>
        </tr>
        <tr align="center" bgcolor="#FFFFFF">
          <td align="right" nowrap>
            <a href="info_upd.asp" target="_blank">刷新</a>              <input name="send" type="hidden" id="send" value="sch" >
            <select id=TP name=TP style="width:120; ">
              <%=Get_rsOpt(conn,"SELECT TypID,TypName FROM WebTyps WHERE TypMod='HomLnk1' ORDER BY TypID",InfType)%>
            </select>
            <input name="KW" type="text" id="KW" value="<%=KW%>" size="8">
<input type="submit" name="Submit" value="搜索">
&nbsp;</td>
        </tr></form>
      </table></td>
  </tr>
  <tr align="center" bgcolor="E0E0E0">
    <td height="27" colspan="7" align="center" bgcolor="f8f8f8"><%= RS_Page(rs,Page,"?send=pag&TP="&TP&"&KW="&KW&"",1)%></td>
  </tr>
  <tr align="center" bgcolor="E0E0E0">
    <td width="5%" height="27" align="center" nowrap>NO</td>
    <td width="40%" height="27" align="center">名称</td>
    <td align="center" nowrap>&nbsp;</td>
    <td height="27" align="center" nowrap>类别</td>
    <td align="center" nowrap>显序</td>
    <td height="27" align="center" nowrap>修改</td>
  </tr>
  <tr bgcolor="#333333">
    <td colspan="9" align="right" nowrap></td>
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
InfType = rs("InfType")
TypName = rs_Val("","SELECT TypName FROM WebTyps WHERE TypID='"&InfType&"'")
InfSubj = rs("InfSubj")
InfURL = rs("InfURL")
ImgName = rs("ImgName")
SetShow = rs("SetShow")
SetTop = rs("SetTop")
SetSubj = rs("SetSubj")
InfSubj = Show_sTitle(InfSubj,SetSubj)
	  %>
    <tr bgcolor="<%=col%>">
      <td align="right" nowrap><%=i%>
      <input name="yID" type="checkbox" id="yID" value="<%=KeyID%>"></td>
      <td><a href="<%=InfURL%>" target="_blank"><%=InfSubj%></a></td>
      <td align="center" nowrap>&nbsp;</td>
      <td align="center" nowrap><%=TypName%></td>
      <td align="center" nowrap><%=SetSAdm%><%=SetTop%></td>
      <td align="center" nowrap><a href="info_edit.asp?ID=<%=KeyID%>&PG=<%=Page%>&KW=<%=KW%>&TP=<%=TP%>&TP2=<%=TP2%>&TP3=<%=TP3%>">修改</a></td>
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
        <input name="TP3" type="hidden" id="TP3" value="<%=TP3%>">
        <input name="KW" type="hidden" id="KW" value="<%=KW%>">
        <input name="TP4" type="hidden" id="TP4" value="<%=TP4%>"></td>
      <td colspan="3" align="right" nowrap><select name="yAct" id="yAct" >
        <option value="del_sel">删除.所选</option>
        <!--
        <option value="del_now">删除.当前</option>
        <option value="SetSAdm">设置_显示</option>
        -->
        <option value="SetTop">设置_顺序</option>
      </select>
        <select name="yVal" id="yVal">
          <option value="Y" selected>Y</option>
          <option value="N">N</option>
          <option value="X">X</option>
          <%For i=0 To 9%>
          <option value="<%=i%>" ><%=i%></option>
          <%Next%>
          <option value="888" >888</option>
        </select>      </td>
      <td colspan="2" align="left" nowrap><input type="submit" name="Submit" value="执行"></td>
    </tr>
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
