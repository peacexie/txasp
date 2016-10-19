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
KW = RequestS("KW",3,60)
TP = RequestS("TP",3,255)
TP2= RequestS("TP2",3,48)&RequestS("InfTyp2",3,48)
If Left(KW,4)="uid=" Then
  sqlK = sqlK & " AND ( LogAUser LIKE '"&Mid(KW,5)&"%' ) " 
ElseIf inStr(KW,"=")>0 Then
   aKW = Split(KW,"=")
  sqlK = sqlK & " AND ( "&aKW(0)&" LIKE '"&aKW(1)&"%' ) " 
ElseIf KW&"" <> "" Then
  sqlK = sqlK & " AND ( InfSubj LIKE '%"&KW&"%' ) " 
End If
If TP&"" <> "" then
  sqlK = sqlK & " AND (InfType LIKE '"&TP&"%') " 
End If
If TP2&"" <> "" then
  sqlK = sqlK & " AND (InfTyp2='"&TP2&"') " 
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
  If Chk_URL3(Config_Path&"smod/info/info_list.asp")="eUrl" Then
    Response.End()
  End If
  If yAct="SetTop" Then
    yVal=RequestSafe(yVal,"N",888)
  Else
    yVal=Left(yVal,1)
  End If
  If sID<>"" Then
    sID = Replace(sID&"''",",''","")
	sql = " UPDATE "&ModTab&" SET "&yAct&"='"&yVal&"' "
    sql = sql& " WHERE KeyID IN("&sID&")" 
	Call rs_DoSql(conn,sql)
  End If
    Msg = cID&" 条记录 设置成功!"
Elseif yAct="del_sel" then
  If Chk_URL3(Config_Path&"smod/info/info_list.asp")="eUrl" Then
    Response.End()
  End If
	Call del_sfDir(ModTab,sID) 
    Msg = cID&" 条记录删除成功!"
End If

If PrmFlag="(Mem)" Then '会员 写给网管
  sqlU = " AND LogAUser='"&Session("MemID")&"' "
ElseIf PrmFlag="(Inn)" Then '内部会员 个人笔记
  sqlU = " AND LogAUser='"&Session("InnID")&"' "
ElseIf NOT Flg_PList() Then
  sqlU = " AND LogAUser='"&Session("UsrID")&"' "
End If

sqlOrd = "LogATime DESC"
'If TP&"" <> "" then
  'sqlOrd = " SetTop,LogATime DESC"
'End If 

    sql = " SELECT * FROM "&ModTab&" "
	sql =sql& " WHERE KeyMod='"&ModID&"' "&sqlK&sqlU 
	sql =sql& " ORDER BY "&sqlOrd&" " 
   Set rs=Server.Createobject("Adodb.Recordset")
   rs.Open Sql,conn,1,1
   rs.PageSize = 15 
If Int(Page)>rs.PageCount Or Int(Page)<1 Then
  Page = 1
End If

MDName = rs_Val("","SELECT SysName FROM AdmSyst WHERE SysID='"&ModID&"'")
MDRFlag = Eval(ModID&"ImgRelat")
%>
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="E0E0E0">
  <tr align="center" bgcolor="E0E0E0">
    <td height="27" colspan="11" align="center" bgcolor="f8f8f8"><table width="100%" border="0" cellpadding="3" cellspacing="0">
        <tr align="center" bgcolor="#FFFFFF">
          <td width="30%" align="center" bgcolor="#FFFFFF"><strong>[<%=MDName%>]管理 </strong>
        <%If get_ModCopy(ModID) Then%>
            <br>
        <a href="set_copy.asp?ModID=<%=ModID%>&Act=Copy" target="_blank" title="[此版本]与[中文版]信息同步"><font color="#0000FF">《同步》</font></a>
        <a href="set_copy.asp?ModID=<%=ModID%>&Act=Clear" target="_blank" title="清理[此版本]中与[中文版]没有同步的信息"><font color="#0000FF">《清理》</font></a>
        <%
		Else
		  Response.Write " | <a href='info_add.asp?ModID="&ModID&"&PrmFlag="&PrmFlag&"'>新增&gt;&gt;</a>"
		End If
		%>

          </td>
          <td width="30%" align="center" bgcolor="#FFFFFF">&nbsp;&nbsp;&nbsp;<font color="#FF0000"><%=msg%></font></td>
          <form name="fm01" method="post" action="?">
            <td align="right" nowrap><%=Get_TypeLable(ModID)%>
              <input name="send" type="hidden" id="send" value="sch">
              <input name="ModID" type="hidden" id="ModID" value="<%=ModID%>">
              <input name="PrmFlag" type="hidden" id="PrmFlag" value="<%=PrmFlag%>">
              <select class=form id=select name=TP >
			  <%=Get_TypeOpt(ModID,InfType)%>
              </select>
              <input name="KW" type="text" id="KW" value="<%=KW%>" size="12">
              &nbsp;
<%=Get_Typ2Opt(ModID,TP2) %>
              <input type="submit" name="Submit" value="搜索">
            </td>
          </form>
        </tr>
      </table></td>
  </tr>
  <tr align="center" bgcolor="E0E0E0">
    <td height="27" colspan="11" align="center" bgcolor="f8f8f8"><%= RS_Page(rs,Page,"?send=pag&TP="&TP&"&KW="&KW&"&TP2="&TP2&"",1)%></td>
  </tr>
  <tr align="center" bgcolor="E0E0E0">
    <td width="5%" height="27" align="center" nowrap>NO</td>
    <td width="40%" height="27" align="center">主题</td>
    <td height="27" align="center" nowrap><%=Get_TypeLable(ModID)%></td>
    <td align="center" nowrap>
  		<%
		tmpTyp2 = Eval(ModID&"Typ2Code")
		If tmpTyp2&""<>"" Then
		  Response.Write Eval(ModID&"Typ2Lable")
		End If
		%>
    </td>
    
   
    <td align="center" nowrap>审核</td>
    <td align="center" nowrap>推荐</td>
    <td height="27" align="center" nowrap>顺序</td>
    <td align="center" nowrap>修改</td>
    <td align="center" nowrap>阅读</td>
    <td align="center" nowrap>发布时间</td>
    
  <td align="center" nowrap>发布ID</td>
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
InfType = rs("InfType") : TypName = Get_TypeName(ModID,InfType) 
InfTyp2 = rs("InfTyp2") : InfTyp2 = Get_Typ2Name(ModID,InfTyp2)
InfSubj = rs("InfSubj")
SetSubj = rs("SetSubj")
SetRead = rs("SetRead")
SetSubj = rs("SetSubj")
SetHot = rs("SetHot")
SetTop = rs("SetTop")
SetShow = rs("SetShow")
ImgName = rs("ImgName")
LogATime = rs("LogATime")
LogAUser = rs("LogAUser")
InfSubj = Show_sTitle(InfSubj,SetSubj)
ImgFlag = ""
'If inStr(ImgName,".")>0 Then
If Len(ImgName)>15 Then
  ImgFlag = "<img src='../../img/tool/attfile.gif' border=0 align='absmiddle'>" 
End If
isOwn = Flg_PUser(PrmFlag,LogAUser)
isDep = Flg_PDep(InfTyp2,KeyMod)
If SetShow="" Then SetShow="-"
SetShow = Get_State(SetShow,"Y;N;X;-","已审;未审;未过;未知")
If SetHot="" Then SetHot="-"
SetHot = Get_State(SetHot,"N;Y;-","一般;推荐;未知")
	  %>
    <tr bgcolor="<%=col%>">
      <td align="right" nowrap><%=i%>
        <%If isOwn AND isDep Then%>
        <input name="yID" type="checkbox" id="yID" value="<%=KeyID%>">
        <%Else%>
        <input name="xID" type="checkbox" id="xID" value="---" disabled>
        <%End If%>
      </td>
      <td align="left">
      <%If MDRFlag Then%>
 <div style="float:right; border:1px solid #CCC; padding:0px 3px">
 <a href='../ibat/rpic_set.asp?IR=<%=KeyID%>&ModID=<%=ModID%>&PG=<%=Page%>&KW=<%=KW%>&TP=<%=TP%>&TP2=<%=TP2%>&PrmFlag=<%=PrmFlag%>'>相关图片</a>
 </div>
<%End If
	  'Response.Write ImgName&Len(ImgName)&"  ---- "
	  %>
      <a href="info_view.asp?ID=<%=KeyID%>&PrmFlag=<%=PrmFlag%>" target="_blank"><%=InfSubj%></a><%=ImgFlag%>
      <%=xxDep%>
      </td>
      <td nowrap><%=TypName%></td>
      <td align="center" nowrap><%=InfTyp2%></td>
      
      
      <td align="center" nowrap><%=SetShow%></td>
      <td align="center" nowrap><%=SetHot%></td>
      <td align="center" nowrap><%=SetTop%></td>
      
      <td align="center" nowrap>
      <% If isOwn And isDep Then %>
      <a href="info_edit.asp?ID=<%=KeyID%>&PG=<%=Page%>&KW=<%=KW%>&TP=<%=TP%>&TP2=<%=TP2%>&PrmFlag=<%=PrmFlag%>">修改</a>
      <%Else%>
      <font color="#CCCCCC">修改</font>
	  <%End If%></td>
      
      <td align="center" nowrap><%=SetRead%></td>
      <td align="center" nowrap><%=LogATime%></td>

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
      <td colspan="3" align="right"><input name="Page" type="hidden" id="Page" value="<%=Page%>">
        <input name="TP" type="hidden" id="TP" value="<%=TP%>">
        <input name="TP2" type="hidden" id="TP2" value="<%=TP2%>">
        <input name="KW" type="hidden" id="KW" value="<%=KW%>">
      <input name="PrmFlag" type="hidden" id="PrmFlag" value="<%=PrmFlag%>">      
      <a href="set_top.asp?ModID=<%=ModID%>">高级设置</a>
      | <a href="../ibat/bat_up.asp?ModID=<%=ModID%>">批量上传</a>
      | <a href="../ibat/imp_set.asp?ModID=<%=ModID%>">信息采集</a>
      </td>
      <td colspan="5" align="right" nowrap><select name="yAct" id="yAct" onChange="fAct()">
          <option value="del_sel">删除.所选</option>
          <%If Flg_PCheck() Then%>
          <option value="SetShow">设置_审核</option>
          <%End If%>
          <option value="SetHot" selected>设置_推荐</option>
		  <option value="SetTop">设置_顺序</option>
        </select>
        <select name="yVal" id="yVal">
        </select>
      </td>
      <td colspan="2" align="left" nowrap><input type="submit" name="Submit2" value="执行"></td>
    </tr>
    <%  
  
  Else
  %>
    <tr align="center" bgcolor="#f4f4f4">
      <td colspan="13">无信息 <a href="../ibat/bat_up.asp?ModID=<%=ModID%>">批量上传</a> </td>
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

var sAct = document.getElementById("yAct");
var sVal = document.getElementById("yVal");

function fAct(){
  sVal.options.length = 0;
  if(sAct.value=="del_sel"){
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
</body>
</html>
