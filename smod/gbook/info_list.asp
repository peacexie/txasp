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
Tim1=Timer()
Set rs=Server.Createobject("Adodb.Recordset")
yAct = Request("yAct") 
Page = RequestS("Page","N",1)
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
If TP&"" <> "" then
  sqlK = sqlK & " AND (InfType LIKE '%"&TP&"%') " 
End If
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
    sql = " UPDATE GboInfo SET "&yAct&"='"&yVal&"' "
    sql = sql& " WHERE KeyID IN("&sID&")" ':Response.Write sql : Response.End()
	Call rs_DoSql(conn,sql)
  End If
    Msg = cID&" 条记录 设置成功!"
Elseif yAct="del_sel" then
  If Chk_URL3(Config_Path&"smod/gbook/info_list.asp")="eUrl" Then
    Response.End()
  End If
  For iy = 1 To Request.Form("yID").Count
    iID = RequestSafe(Request.Form("yID").item(iy),3,96) 
    If iID<>"" Then
	  sID = sID&"'"&iID&"',"
	  cID = cID + 1
	End If
  Next
	Call del_sfCont("GboInfo",sID)
    Msg = cID&" 条记录删除成功!"
ElseIf yAct="del_now" Then
 If KW&TP<>"" Then
  If Chk_URL3(Config_Path&"smod/gbook/info_list.asp")="eUrl" Then
    Response.End()
  End If
	sql = "SELECT KeyID FROM [GboInfo] WHERE KeyMod='"&ModID&"' "&sqlK 
	rs.Open sql,conn,1,1 
	cID = rs.RecordCount
	Do While NOT rs.EOF
 	 iID = rs("KeyID") 
	 Call del_sfCont("GboInfo",iID)
	rs.MoveNext
	Loop
	rs.Close()
    Msg = cID&" 条记录删除成功!"
 End If
  sqlK="" : TP="" : KW=""
  Page=1 ' 清楚条件,重设第一页
End If

If ModID="GboU224" Then                     '管理员 个人笔记
  sqlU = " AND LogAUser='"&Session("UsrID")&"' "
ElseIf ModID="MemB324" AND PrmFlag="(Mem)" Then '会员 个人笔记
  sqlU = " AND LogAUser='"&Session("MemID")&"' "
ElseIf ModID="MemB224" AND PrmFlag="(Mem)" Then '系统通知 ------ 
  sqlU = " AND (LnkName='"&Session("MemID")&"' OR LnkName='(Public)' )"
ElseIf ModID="MemB124" AND PrmFlag="(Mem)" Then '会员 写给网管
  sqlU = " AND LogAUser='"&Session("MemID")&"' "
  
ElseIf ModID="MemB624" AND PrmFlag="(Inn)" Then '内部会员 个人笔记
  sqlU = " AND LogAUser='"&Session("InnID")&"' "
ElseIf ModID="MemB524" AND PrmFlag="(Inn)" Then '内部会员 系统通知 ------ 
  sqlU = " AND (LnkName='"&Session("InnID")&"' OR LnkName='(Public)' )"
ElseIf ModID="MemB424" AND PrmFlag="(Inn)" Then '内部会员 写给网管
  sqlU = " AND LogAUser='"&Session("InnID")&"' "
  
End If

    sql = " SELECT * FROM [GboInfo] "
	sql =sql& " WHERE KeyMod='"&ModID&"' "&sqlK&sqlU 'LogAUser='"&Session("MemID")&"' 1=1(Admin)
	sql =sql& " ORDER BY KeyID DESC"  ':Response.Write sql&SetTop
   rs.Open Sql,conn,1,1
   rs.PageSize = 150
If Int(Page)>rs.PageCount Or Int(Page)<1 Then
  Page = 1
End If

%>
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="E0E0E0">
  <tr align="center" bgcolor="E0E0E0">
    <td height="27" colspan="10" align="center" bgcolor="f8f8f8"><table width="100%" border="0" cellpadding="3" cellspacing="0">
        <tr align="center" bgcolor="#FFFFFF">
          <td width="30%" align="center" bgcolor="#FFFFFF"><strong>[<%=rs_Val("","SELECT SysName FROM AdmSyst WHERE SysID='"&ModID&"'")%>]&gt;&gt;</strong>
		  
		  <a href="info_add.asp?ModID=<%=ModID%>&PrmFlag=<%=PrmFlag%>">增加</a>
		  		  </td>
          <td width="30%" align="center" bgcolor="#FFFFFF">&nbsp;&nbsp;&nbsp;<font color="#FF0000"><%=msg%></font></td>
          <form name="fm01" method="post" action="?">
            <td align="right" nowrap><input name="send" type="hidden" id="send" value="sch">
              <input name="ModID" type="hidden" id="ModID" value="<%=ModID%>">
              <input name="PrmFlag" type="hidden" id="PrmFlag" value="<%=PrmFlag%>">
              <select name="TP" style="width:120px; " id="TP">
                <option value="">[不限栏目]</option>
                <%=Get_rsOpt(conn,"SELECT TypID,TypName FROM WebTyps WHERE TypMod='"&ModID&"' ORDER BY TypID",TP)%>
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
    <td width="5%" align="center" nowrap>审核</td>
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
KeyFlag = rs("KeyFlag")
InfType = rs("InfType")
TypName = rs_Val("","SELECT TypName FROM WebTyps WHERE TypID='"&InfType&"'")
If TypName="" Then
  TypName = MDName
End If
InfSubj = Show_Text(rs("InfSubj")) 'rs("InfCont")
LnkName = rs("LnkName")
LnkEmail = rs("LnkEmail")
SetRead = rs("SetRead")
SetShow = rs("SetShow")
ImgName = rs("ImgName")
LogAddIP = rs("LogAddIP")
LogAUser = rs("LogAUser")
LogATime = rs("LogATime")
xxxReply = rs("InfReply")
InfReply = Show_sfRead(KeyID,".rep.htm") 'Show_sfGbook(KeyID,".rep.htm")
'Response.Write InfReply
If InfReply&""="" Then
  InfReply = "<font color=gray>无回复</font>"
Else
  InfReply = "<font color='#ff00ff'>已回复</font>"
End If
isOwn = Flg_PUser(PrmFlag,LogAUser)
If SetShow="" Then SetShow="-"
SetShow = Get_State(SetShow,"Y;N;X;-","已审;未审;审核未通过;未知状态")
	  %>
    <tr bgcolor="<%=col%>">
      <td align="right" nowrap><%=i%>
        <%If isOwn Then%>
        <input name="yID" type="checkbox" id="yID" value="<%=KeyID%>">
        <%Else%>
        <input name="xID" type="checkbox" id="xID" value="---" disabled>
        <%End If%></td>
      <td><a href="info_view.asp?ID=<%=KeyID%>&PrmFlag=<%=PrmFlag%>" target="_blank"><%=InfSubj%></a> </td>
      <td align="center" nowrap>&nbsp;</td>
      <td align="center" nowrap><%=LnkName%></td>
      <td align="center" nowrap><%=TypName%></td>
      <td align="center" nowrap><%=LogATime%></td>
      <td align="center" nowrap><%=SetShow%></td>
      <td align="center" nowrap><%=InfReply%></td>
      <td align="center" nowrap>
      <% If isOwn Then %>
        <a href="info_edit.asp?ID=<%=KeyID%>&PG=<%=Page%>&KW=<%=KW%>&TP=<%=TP%>&CD=<%=CD%>&PrmFlag=<%=PrmFlag%>">修改</a>
      <%Else%>
        <span style="color:#CCC">修改</span>
      <%End If%>
      </td>
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
        <select name="yAct" id="yAct" onChange="fAct()">
          <option value="del_sel">删除.所选</option>
          <option value="del_now">删除.当前</option>
          <option value="SetShow">设置_审核</option>
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

var sAct = document.getElementById("yAct");
var sVal = document.getElementById("yVal");

function fAct(){
  sVal.options.length = 0;
  if((sAct.value=="del_sel")||(sAct.value=="del_now")){
    sVal.options.add(new Option("删除","删除")); 
  }
  if(sAct.value=="SetShow"){
    //sVal.options.add(new Option("审核状态","Y")); 
	sVal.options.add(new Option("已审","Y"));
	sVal.options.add(new Option("未审","N"));
	//sVal.options.add(new Option("审核未通过","X"));
  }
}
try{fAct();}
catch(err){}

function ySel(){
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
