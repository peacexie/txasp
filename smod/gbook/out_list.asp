<!--#include file="config.asp"-->

<html>
<head>
<meta http-equiv="Pragma" content="no-cache">
<meta http-equiv="Expires" content="0">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title><%=Config_Name%>后台管理中心</title>
<link href="../../inc/adm_inc/adm_style.css" rel="stylesheet" type="text/css">
<script src="/sadm/func1/Func_JS.js" type="text/javascript"></script>
</head>
<body>

<%

MDFlag = Request("MDFlag") 
yAct = Request("yAct") 
Page = RequestS("Page","N",1)
yVal = RequestS("yVal","C",2)
KW = RequestS("KW",3,24)
TP = RequestS("TP",3,24)
If Left(KW,4)="uid=" Then
  sqlK = sqlK & " AND ( LogAUser LIKE '"&Mid(KW,5)&"%' ) " 
ElseIf inStr(KW,"=")>0 Then
   aKW = Split(KW,"=")
  sqlK = sqlK & " AND ( "&aKW(0)&" LIKE '"&aKW(1)&"%' ) " 
ElseIf KW&"" <> "" Then
  sqlK = sqlK & " AND ( InfSubj LIKE '%"&KW&"%' ) " 
End If

If MDFlag="JApp" Then
  ModTP = RequestS("ModTP","C",24)
  MDName = rs_Val("","SELECT TypName FROM WebTyps WHERE TypID='"&ModTP&"'")&" 职位申请"
  MDSubj = "姓名"
  MDType = "职位"
  MDSql = ModTP
ElseIf MDFlag="POrd" Then
  MDSubj = "主题"
  MDType = "产品"
Else
  MDName = rs_Val("","SELECT SysName FROM AdmSyst WHERE SysID='"&ModID&"'")&" 评论管理"
  MDSubj = "主题"
  MDType = "态度-姓名"
  MDTypx = "-"
  MDSql = ModID
End If

Set rs=Server.Createobject("Adodb.Recordset")
cID = 0
sID = ""
If yAct="del_sel" then
  For iy = 1 To Request.Form("yID").Count
    iID = RequestSafe(Request.Form("yID").item(iy),3,96) 
    If iID<>"" Then
	  cID = cID + 1
	  Call rs_DoSql(conn,"DELETE FROM GboSend WHERE KeyID='"&iID&"'")
	  Call del_sfCont("GboSend",iID)
	End If
  Next
    Msg = cID&" 条记录删除成功!"
	
ElseIf yAct="del_now" then

	sql = "SELECT KeyID FROM [GboSend] WHERE KeyMod='"&MDSql&"' "&sqlK 
	rs.Open sql,conn,1,1 
	cID = rs.RecordCount
	Do While NOT rs.EOF
 	 iID = rs("KeyID") 
	 Call del_sfCont("GboSend",iID)
	rs.MoveNext
	Loop
	rs.Close()
    Msg = cID&" 条记录删除成功!"

Elseif yAct="SetShow" Then

  For iy = 1 To Request.Form("yID").Count
    iID = RequestSafe(Request.Form("yID").item(iy),3,96) 
    If iID<>"" Then
	  cID = cID + 1
	  Call rs_DoSql(conn,"UPDATE GboSend SET "&yAct&"='"&yVal&"' WHERE KeyID='"&iID&"'")
	End If
  Next

    Msg = cID&" 条记录 设置成功!"
End If
    sql = " SELECT GboSend.* FROM [GboSend] "
	sql =sql& " WHERE KeyMod='"&MDSql&"' "&sqlK
	sql =sql& " ORDER BY KeyID DESC" 'SetTop,
   rs.Open Sql,conn,1,1
   rs.PageSize = 18 
If Int(Page)>rs.PageCount Or Int(Page)<1 Then
  Page = 1
End If



%>
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="E0E0E0">
  <tr align="center" bgcolor="E0E0E0">
    <td height="27" colspan="7" align="center" bgcolor="f8f8f8"><table width="100%" border="0" cellpadding="3" cellspacing="0">
        <tr align="center" bgcolor="#FFFFFF">
          <td width="30%" align="center" bgcolor="#FFFFFF"><strong>[<%=MDName%>]</strong></td>
          <td width="30%" align="center" bgcolor="#FFFFFF">&nbsp;&nbsp;&nbsp;<font color="#FF0000"><%=msg%></font></td>
          <form name="fm01" method="post" action="?">
            <td align="right" nowrap>
            <%
			sList = Replace(tmsListTab(12),"Jobs:","")
			aList = Split(sList,";")
			For i=0 To uBound(aList)-1
			  Response.Write "<a href='?ModTP="&aList(i)&"&MDFlag=JApp'>职位申请</a> "
			Next
			%>
            <input name="send" type="hidden" id="send" value="sch">
              <input name="ModTP" type="hidden" id="ModTP" value="<%=ModTP%>">
              <input name="MDFlag" type="hidden" id="MDFlag" value="<%=MDFlag%>">
关键字:
<input name="KW" type="text" id="KW" value="<%=KW%>" size="12">
              <input type="submit" name="Submit" value="搜索">            </td>
          </form>
        </tr>
      </table></td>
  </tr>
  <tr align="center" bgcolor="E0E0E0">
    <td height="27" colspan="7" align="center" bgcolor="f8f8f8"><%= RS_Page(rs,Page,"?send=pag&TP="&TP&"&KW="&KW&"&TUP="&ModID&"",1)%></td>
  </tr>
  <tr align="center" bgcolor="E0E0E0">
    <td width="5%" height="27" align="center" nowrap>NO</td>
    <td width="20%" height="27" align="center"><%=MDSubj%></td>
    <td width="30%" height="27" align="center" nowrap><%=MDType%></td>
    <td width="15%" align="center" nowrap>编号</td>
    <td width="15%" align="center" nowrap>日期</td>
    <td align="center" nowrap>审核</td>
    <td height="27" align="center" nowrap>查看</td>
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
KeyMod = rs("KeyMod")
KeyCode = rs("KeyCode")
InfType = rs("InfType")
InfSubj = rs("InfSubj")
LogATime = rs("LogATime")
LogAUser = rs("LogAUser")
SetShow = rs("SetShow")
InfSubj = Show_Text(InfSubj) 'Show_SetSubj(InfSubj,SetHot,SetSub)
If MDTypx = "-" Then
  InfType = "<font color='#0000ff'>"&InfType&"</font>-"&LogAUser
End If
If SetShow="" Or SetShow="-" Then SetShow="X"
SetShow = Get_State(SetShow,"Y;N;X","已审;未审;未知")
	  %>
    <tr bgcolor="<%=col%>">
      <td align="right" nowrap><%=i%>
      <input 
			name="yID" type="checkbox" id="yID" value="<%=KeyID%>"></td>
      <td><a href="out_view.asp?ID=<%=KeyID%>" target="_blank"><%=InfSubj%></a></td>
      <td align="center" nowrap><%=InfType%></td>
      <td align="center" nowrap><%=KeyCode%></td>
      <td align="center" nowrap><%=LogATime%></td>
      <td align="center" nowrap><%=SetShow%></td>
      <td align="center" nowrap><a href="out_view.asp?ID=<%=KeyID%>" target="_blank">查看</a></td>
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
        <input name="ModTP" type="hidden" id="ModTP" value="<%=ModTP%>">
        <input name="MDFlag" type="hidden" id="MDFlag" value="<%=MDFlag%>"></td>
      <td colspan="3" align="right" nowrap><select name="yAct" id="yAct" onChange="fAct()">
          <option value="del_sel">删除.所选</option>
          <option value="del_now">删除.当前</option>
          <option value="SetShow">设置_审核</option>
            </select>
        <select name="yVal" id="yVal">
		  <option value="D">删除</option>
        </select>
            </td>
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

var sAct = document.getElementById("yAct");
var sVal = document.getElementById("yVal");

function fAct(){
  sVal.options.length = 0;
  if(sAct.value=="SetShow"){
	sVal.options.add(new Option("已审","Y"));
	sVal.options.add(new Option("未审","N"));
  }else{
	sVal.options.add(new Option("删除","D"));  
  }
  //alert(XAct);
}
<%
If yAct<>"" Then 
  Response.Write("sAct.value='"&yAct&"'")
End If
%>
fAct();

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
