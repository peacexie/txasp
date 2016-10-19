<!--#include file="../../sadm/func1/func1.asp"-->
<!--#include file="../../sadm/func2/Func2.asp"-->
<!--#include file="../../sadm/func2/func_sfile.asp" -->
<!--#include file="../../sadm/func1/func_file.asp"-->
<!--#include file="../../sadm/func1/func_opt.asp"-->
<!--#include file="config.asp"-->
<!--#include file="conpub.asp"-->
<html>
<head>
<meta http-equiv="Pragma" content="no-cache">
<meta http-equiv="Expires" content="0">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title><%=Config_Name%>后台管理中心</title>
<link href="../../inc/adm_inc/adm_style.css" rel="stylesheet" type="text/css">
</head>
<body>
<script src="../../sadm/func1/Func_JS.js" type="text/javascript"></script>
<%

ModID = RequestS("ModID",3,24)
ModTab = "BBSInfo"
yAct = Request("yAct") 
Page = RequestS("Page","N",1)
KW = RequestS("KW",3,24)
TP = RequestS("TP",3,240)
	If KW&"" <> "" Then
	  sqlK = sqlK & " AND ( InfSubj LIKE '%"&KW&"%' ) " 
	End If
	If TP&"" <> "" then
	  sqlK = sqlK & " AND (InfType LIKE '%"&TP&"%') " 
	End If
cID = 0
sID = ""
If yAct="SetShow" OR yAct="SetHot" Then
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
    sql = " UPDATE BBSInfo SET "&yAct&"='"&yVal&"' "
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
	  Call bbsDelID(iID,"") 
	End If
  Next
    Msg = cID&" 条记录删除成功!"
ElseIf yAct="del_now" Then
 If KW&TP<>"" Then
	 sql = "SELECT KeyID FROM BBSInfo WHERE 1=1 "&sqlK&" "
	 Set rs = Server.Createobject("Adodb.Recordset")
	 rs.Open sql,conn,1,1
	 cID = rs.RecordCount
	 Do While NOT rs.EOF
 	   iID = rs("KeyID") 
	   Call del_sfDir(ModTab,iID)
	 rs.MoveNext
	 Loop
	 rs.Close()
	 Set rs = Nothing
    Msg = cID&" 条记录删除成功!"
 End If
  sqlK="" : TP="" : KW=""
  Page=1 ' 清楚条件,重设第一页
ElseIf yAct="Clear" Then 
	sql = "DELETE FROM BBSInfo WHERE KeyMod='(BBS)' AND LEN(KeyRE)>12 AND KeyRe NOT IN(SELECT KeyID FROM BBSInfo) "
	Call rs_DoSql(conn,sql)
    Msg = " 清理成功!"
End If

'If ModID="(Member)" Then
  'sqlU = " AND (KeyMod='(MemTo)' OR KeyMod='(MFrom)') "
'Else
  sqlU = " AND KeyMod='"&ModID&"' "
'End If

    sql = " SELECT * FROM [BBSInfo] "
	sql =sql& " WHERE 1=1 "&sqlK&sqlU 'LogAUser='"&Session("MemID")&"' 1=1(Admin)
	sql =sql& " ORDER BY KeyID DESC"  ':Response.Write sql&SetTop
   Set rs=Server.Createobject("Adodb.Recordset")
   rs.Open Sql,conn,1,1
   rs.PageSize = 18 
If Int(Page)>rs.PageCount Or Int(Page)<1 Then
  Page = 1
End If

%>
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="E0E0E0">
  <tr align="center" bgcolor="E0E0E0">
    <td height="27" colspan="8" align="center" bgcolor="f8f8f8"><table width="100%" border="0" cellpadding="3" cellspacing="0">
        <tr align="center" bgcolor="#FFFFFF">
          <td width="30%" align="center" bgcolor="#FFFFFF"><strong>论坛帖子管理</strong></td>
          <td width="30%" align="center" bgcolor="#FFFFFF">&nbsp;&nbsp;&nbsp;<font color="#FF0000"><%=msg%></font></td>
          <form name="fm01" method="post" action="?">
            <td align="right" nowrap><input name="send" type="hidden" id="send" value="sch">
              <input name="ModID" type="hidden" id="ModID" value="<%=ModID%>">
              <select class="form" id="select" name="TP" >
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
    <td height="27" colspan="8" align="center" bgcolor="f8f8f8"><%= RS_Page(rs,Page,"?send=pag&TP="&TP&"&KW="&KW&"&ModID="&ModID&"",1)%></td>
  </tr>
  <tr align="center" bgcolor="E0E0E0">
    <td width="5%" height="27" align="center" nowrap>NO</td>
    <td height="27" align="center">主题</td>
    <td width="5%" height="27" align="center" nowrap>类别</td>
    <td width="5%" align="center" nowrap>发布</td>
    <td width="5%" align="center" nowrap>推荐</td>
    <td width="5%" align="center" nowrap>审核</td>
    <td width="5%" align="center" nowrap>回复</td>
    <td width="5%" height="27" align="center" nowrap>修改</td>
  </tr>
  <tr bgcolor="#333333">
    <td colspan="10" align="right" nowrap></td>
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
KeyRE = rs("KeyRE")&"" 
KeyMod = rs("KeyMod")
KeyFlag = rs("KeyFlag")
InfType = Show_Text(rs("InfType"))
InfSubj = Show_Text(rs("InfSubj"))
InfCont = rs("InfCont")
InfReply = rs("InfReply")&""
LnkName = rs("LnkName")
LnkEmail = rs("LnkEmail")
SetRead = rs("SetRead")
SetShow = rs("SetShow")
SetSAdm = rs("SetSAdm")
SetHot = rs("SetHot")
SetTop = rs("SetTop")
ImgName = rs("ImgName")
LogAddIP = rs("LogAddIP")
LogAUser = rs("LogAUser")
LogATime = rs("LogATime")
If KeyRe="" Then
 PID = KeyID
 InfReply = "主贴"
Else
 PID = KeyRe
 InfSubj = "<font color=gray>"&InfSubj&"</font>"
 InfReply = "<font color=gray>回复</font>"
End If
TypName = rs_Val("","SELECT TypName FROM WebTyps WHERE TypLayer LIKE '%"&InfType&"%'")
If SetShow="" Then SetShow="-"
SetShow = Get_State(SetShow,"Y;N;X;-","已审;未审;未过;未知")
If SetHot="" Then SetHot="-"
SetHot = Get_State(SetHot,"N;Y;-","一般;推荐;未知")
	  %>
    <tr bgcolor="<%=col%>">
      <td align="right" nowrap><%=i%>
      <input name="yID" type="checkbox" id="yID" value="<%=KeyID%>"></td>
      <td><a href="../../bbs/bview.asp?ID=<%=KeyID%>" target="_blank"><%=InfSubj%></a>      
      </td>
      <td nowrap><%=TypName%></td>
      <td align="center" nowrap><%=LogATime%></td>
      <td align="center" nowrap><%=SetHot%></td>
      <td align="center" nowrap><%=SetShow%></td>
      <td align="center" nowrap><%=InfReply%></td>
      <td align="center" nowrap><a href="../../bbs/bedit.asp?ID=<%=PID%>&eID=<%=KeyID%>" target="_blank">修改</a></td>
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
</td>
      <td colspan="2" nowrap><select name="yAct" id="yAct" onChange="fAct()">
          <option value="SetHot">设置_推荐</option>
          <option value="del_sel">删除.所选</option>
          <option value="del_now">删除.当前</option>
          <option value="SetShow">设置_审核</option>
          <option value="Clear">清理_垃圾</option>
        </select>
        <select name="yVal" id="yVal">
        </select>      </td>
      <td colspan="4" nowrap><input type="submit" name="Submit" value="执行">
&nbsp;提示 </td>
    </tr>
    <%  
  
  Else
  %>
    <tr align="center" bgcolor="#f4f4f4">
      <td colspan="10">无信息</td>
    </tr>
    <%
  End If
	  
	  rs.Close()
	  Set rs = Nothing
	  
	  %>
    <tr bgcolor="#999999">
      <td colspan="10" align="right"></td>
    </tr>
  </form>
</table>
<script type="text/javascript">

var sAct = document.getElementById("yAct");
var sVal = document.getElementById("yVal");

function fAct(){
  sVal.options.length = 0;
  if((sAct.value=="del_sel")||(sAct.value=="del_now")){
    sVal.options.add(new Option("删除","删除")); 
  }
  if(sAct.value=="SetShow"){
	sVal.options.add(new Option("已审","Y"));
	sVal.options.add(new Option("未审","N"));
	//sVal.options.add(new Option("未通过","X"));
  }
  if(sAct.value=="SetHot"){
	sVal.options.add(new Option("推荐","Y"));
	sVal.options.add(new Option("一般","N"));
  }
  if(sAct.value=="SetTop"){
    sVal.options.add(new Option("默认","888"));
    for(i=0;i<=9;i++){ sVal.options.add(new Option(i,i)); }
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
<%

%>
</body>
</html>
