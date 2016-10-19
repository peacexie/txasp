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

TPName = rs_Val("","SELECT SysName FROM AdmSyst WHERE SysID='"&ModID&"'")

yAct = Request("yAct")
Page = RequestS("Page","N",1)
KW = RequestS("KW",3,24)
KT = RequestS("KT",3,255)

If KW&"" <> "" Then
 If KT="kUser" Then
   sqlK = sqlK & " AND ( LogAUser LIKE '"&KW&"%' ) " 
 ElseIf KT="kCode" Then
   sqlK = sqlK & " AND ( TypID LIKE '"&KW&"%' ) " 
 Else 'kName
   sqlK = sqlK & " AND ( TypName LIKE '%"&KW&"%' ) " 
 End If
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

If yAct = "del_now" Then
 If KW<>"" Then
  sql = " DELETE FROM [TradeType] WHERE TypMod='"&ModID&"' "&sqlK
  Call rs_DoSql(conn,sql)
  Msg = "删除成功!"
 End If
ElseIf yAct = "del_sel" Then
 If sID<>"" Then
  sID = sID&"|"
  sID = Replace(sID,"',|","'")
  sql = " DELETE FROM [TradeType] WHERE TypID IN ("&sID&") "
  Call rs_DoSql(conn,sql)
  Msg = "删除成功!"
 End If
Else
  '
End If

TypID = ""

  sql = " SELECT * FROM [TradeType] WHERE TypMod='"&ModID&"' "&sqlK& " ORDER BY TypID DESC " 
  Set rs=Server.Createobject("Adodb.Recordset") 
  rs.Open Sql,conn,1,1
%>
<br>

<table width="720" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="#CCCCCC">
  <tr align="center" bgcolor="E0E0E0">
    <td height="21" colspan="6" align="center" bgcolor="f8f8f8"><table width="100%" border="0" cellpadding="1" cellspacing="0">
        <tr align="center">
          <td width="30%" rowspan="2" align="center"><strong>[<%=TPName%>]</strong> 类别列表</td>
          <td colspan="2" align="center" nowrap><!--#include file="set_menu.asp"--></td>
        </tr>
        <form name="fms21" method="post" action="?">
        <tr align="center">
          <td align="left" nowrap>&nbsp;<font color="#FF0000">提示:<%=msg%></font></td>
          <td width="40%" align="center" nowrap>
            <input name="KW" type="text" id="KW" value="<%=KW%>" size="12" maxlength="12">
            <select name="KT" id="KT">
              <option value="kUser">会员ID</option>
              <option value="kName">类别名称 </option>
              <option value="kCode">类别代码</option>
            </select>
            <input type="submit" name="button" id="button" value="提交">
            <input name="send" type="hidden" id="send" value="send">
            <input name="ModID" type="hidden" id="ModID" value="<%=ModID%>"></td>
        </tr>
        </form>
    </table></td>
  </tr>
  <tr align="center" bgcolor="E0E0E0">
    <td height="27" colspan="6" align="center" bgcolor="f8f8f8"><%= RS_Page(rs,Page,"?send=pag&KT="&KT&"&KW="&KW&"",1)%></td>
  </tr>

  <tr align="center" bgcolor="#FFFFFF">
    <td width="12%" height="27" align="center" bgcolor="#e8e8e8">NO    </td>
    <td width="20%" height="27" align="center" bgcolor="#e8e8e8">类别代码</td>
    <td height="27" align="center" bgcolor="#e8e8e8">类别名称</td>
    <td align="center" bgcolor="#e8e8e8">Flag</td>
    <td height="27" align="center" bgcolor="#e8e8e8">User ID</td>
    <td width="15%" height="27" align="center" bgcolor="#e8e8e8">时间</td>
  </tr>
<form name="flist" method="post" action="?">
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
		TypName = Show_Text(rs("TypName"))
		TypFlag = rs("TypFlag")
		LogAUser = rs("LogAUser")
		LogATime = rs("LogATime")
	  %>

    <tr bgcolor="<%=col%>">
      <td align="right"><%=i%>
      <input 
			name="yID" type="checkbox" id="yID" value="<%=TypID%>"></td>
      <td><%=TypID%></td>
      <td><%=TypName%>      </td>
      <td align="center"><%=TypFlag%></td>
      <td align="center"><%=LogAUser%></td>
      <td align="center"><%=LogATime%></td>
    </tr>

  <%
  rs.movenext
  Loop
%>
    <tr bgcolor="E0E0E0">
      <td height="21" align="right" nowrap>
        <span id="yFlag" style="visibility:hidden ">N</span>全选
      <input name="yAll" type="checkbox" id="yAll" onClick="ySel()"></td>
      <td><input name="Page" type="hidden" id="Page" value="<%=Page%>">
        <input name="TP" type="hidden" id="TP" value="<%=TP%>">
        <input name="KW" type="hidden" id="KW" value="<%=KW%>">
        <input name="ModID" type="hidden" id="ModID" value="<%=ModID%>"></td>
      <td colspan="3" align="right" nowrap><select name="yAct" id="yAct" >
          <option value="del_sel">删除.所选</option>
          <option value="del_now">删除.当前</option>
        </select>
      </td>
      <td align="left" nowrap><input type="submit" name="Submit" value="执行">      </td>
    </tr>
<%
  end if
	  
	  rs.close()
	  set rs = nothing

	  %>
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
