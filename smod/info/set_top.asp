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

yAct = Request("yAct") 
Page = RequestS("Page","N",1)
KW = RequestS("KW",3,24)
TP = RequestS("TP",3,255)
TP2= RequestS("TP2",3,48)
	If KW&"" <> "" Then
	  sqlK = sqlK & " AND ( InfSubj LIKE '%"&KW&"%' "
	  sqlK = sqlK & " OR InfCont LIKE '%"&KW&"%' "
	  sqlK = sqlK & " ) " 
	End If
	If TP&"" <> "" then
	  sqlK = sqlK & " AND (InfType LIKE '"&TP&"%') " 
	End If
	If TP2&"" <> "" then
	  sqlK = sqlK & " AND (InfTyp2='"&TP2&"') " 
	End If
cID = 0
sID = ""
  For iy = 1 To Request.Form("yID").Count
    iID = RequestSafe(Request.Form("yID").item(iy),3,96) 
    If iID<>"" Then
	  sID = sID&"'"&iID&"',"
	  cID = cID + 1
	End If
  Next
If yAct="del_sel" then
  If Chk_URL3(Config_Path&"smod/info/set_top.asp")="eUrl" Then
    Response.End()
  End If
	Call del_sfDir(ModTab,sID) 
    Msg = cID&" 条记录删除成功!"  
ElseIf yAct="LogATime" Or yAct="SetSubj" Or yAct="SetTop" Then
  If sID<>"" Then
    sID = Replace(sID&"''",",''","")
	If yAct="SetTop" Then
      yVal = RequestS("yVal","N",888)
	  sql = " UPDATE "&ModTab&" SET "&yAct&"="&yVal&" "
	ElseIf yAct="SetSubj" Then
      yVal = RequestS("yVal","C",12)
	  sql = " UPDATE "&ModTab&" SET "&yAct&"='"&yVal&"' " 
	ElseIf yAct="LogATime" Then
      yVal = RequestS("yVal","D",Now())
	  sql = " UPDATE "&ModTab&" SET "&yAct&"='"&yVal&"' " 
	End If
    sql = sql& " WHERE KeyID IN("&sID&")" 
	Call rs_DoSql(conn,sql)
  End If
    Msg = cID&" 条记录 设置成功!"
ElseIf yAct="Copy1" Then
    Msg = " 复制成功!"
End If

If yAct="SetTop" Then
  sqlO = "SetTop"
Else
  sqlO = "LogATime DESC"
End If

    sql = " SELECT "&ModTab&".* FROM ["&ModTab&"] "
	sql =sql& " WHERE KeyMod='"&ModID&"' "&sqlK
	sql =sql& " ORDER BY "&sqlO&"" 
   Set rs=Server.Createobject("Adodb.Recordset")
   rs.Open Sql,conn,1,1
   rs.PageSize = 15 
If Int(Page)>rs.PageCount Or Int(Page)<1 Then
  Page = 1
End If

%>
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="E0E0E0">
  <tr align="center" bgcolor="E0E0E0">
    <td height="27" colspan="7" align="center" bgcolor="f8f8f8"><table width="100%" border="0" cellpadding="3" cellspacing="0">
        <tr align="center" bgcolor="#FFFFFF">
          <td width="30%" align="center" bgcolor="#FFFFFF"><strong>[<%=rs_Val("","SELECT SysName FROM AdmSyst WHERE SysID='"&ModID&"'")%>] </strong> | 高级设置</td>
          <td width="30%" align="center" bgcolor="#FFFFFF">&nbsp;&nbsp;&nbsp;<font color="#FF0000"><%=msg%></font></td>
          <form name="fm01" method="post" action="?">
            <td align="right" nowrap><input name="send" type="hidden" id="send" value="sch">
              <select class=form id=select name=TP >
			  <%=Get_rsTree(conn,"SELECT * FROM WebTyps WHERE TypMod='"&ModID&"' ORDER BY TypLayer ",InfType,"Lay") %>
              </select>
              <input name="KW" type="text" id="KW" value="<%=KW%>" size="12">
              <input type="submit" name="Submit" value="搜索">            </td>
          </form>
        </tr>
      </table></td>
  </tr>
  <tr align="center" bgcolor="E0E0E0">
    <td height="27" colspan="7" align="center" bgcolor="f8f8f8"><%= RS_Page(rs,Page,"?send=pag&TP="&TP&"&KW="&KW&"&TP2="&TP2&"",1)%></td>
  </tr>
  <tr align="center" bgcolor="E0E0E0">
    <td width="5%" height="27" align="center" nowrap>NO</td>
    <td width="40%" height="27" align="center">主题</td>
    <td height="27" align="center" nowrap>类别</td>
    <td align="center" nowrap>发布</td>
    <td align="center" nowrap>顺序</td>
    <td height="27" align="center" nowrap>颜色</td>
    <td align="center" nowrap>复制</td>
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
TypName = rs_Val("","SELECT TypName FROM WebTyps WHERE TypLayer='"&InfType&"'")
InfSubj = rs("InfSubj")
SetSubj = rs("SetSubj")
SetTop = rs("SetTop")
LogATime = rs("LogATime")
ImgName = rs("ImgName")
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
      <td><a href="info_view.asp?ID=<%=KeyID%>" target="_blank"><%=InfSubj%></a><%=ImgFlag%>     </td>
      <td align="center" nowrap><%=TypName%><%=""&TypNam2%></td>
      <td align="center" nowrap><%=LogATime%></td>
      <td align="center" nowrap><%=SetTop%></td>
      <td align="center" nowrap><%=SetSubj%></td>
      <td align="center" nowrap><a href="set_copy.asp?Act=Copy1&ID=<%=KeyID%>&ModID=<%=ModID%>" target="_blank">复制</a></td>
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
        <input name="KW" type="hidden" id="KW" value="<%=KW%>">
        <span class="colF00">提示: <span id="nAct">&nbsp;</span></span></td>
      <td align="right" nowrap><select name="yAct" id="yAct" onChange="fAct()">
        <%=Get_SOpt("LogATime;SetTop;SetSubj","设置_时间;设置_顺序;设置_颜色",yAct,xFlag)%>
        <option value="del_sel">删除.所选</option>
      </select></td>
      <td colspan="2" align="center" nowrap><input name="yVal" type="text" id="yVal" value="<%=Now()%>" size="18" maxlength="20"></td>
      <td align="left" nowrap><input type="submit" name="Submit2" value="执行"></td>
      <td align="left" nowrap>&nbsp;</td>
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


function fAct()
{
  var frmID = document.flist;
  var vAct = frmID.yAct.value; 
  if(vAct=='SetTop'){
    nAct.innerHTML = "设置值在[100~999]之间";
	frmID.yVal.value = '888';
  }
  if(vAct=='SetSubj'){
    nAct.innerHTML = "设置格式[RRGGBB]";
	frmID.yVal.value = '000000';
  }
  if(vAct=='LogATime'){
    nAct.innerHTML = "设置格式[yyyy-mm-dd HH:mm:ss]";
	frmID.yVal.value = '<%=Now()%>';
  }
  //alert(XAct);
}
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
