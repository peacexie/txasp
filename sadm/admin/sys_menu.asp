<!--#include file="config.asp"-->
<html>
<head>
<meta http-equiv="Pragma" content="no-cache">
<meta http-equiv="Expires" content="0">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title><%=Config_Name%>后台管理中心</title>
<script src="../func1/Func_JS.js" type="text/javascript"></script>
<link href="../../inc/adm_inc/adm_style.css" rel="stylesheet" type="text/css">
<style type="text/css">
.btnA1 {
	font-size:12px;
	height:15px;
	padding:1px 0px;
	margin:0px;
	border-left:0px;
	border-top:0px;
	border-right:1px solid #999;
	border-bottom:1px solid #999;
}
</style>
</head>
<body>

<%

SET rs=Server.CreateObject("Adodb.Recordset") 


yAct = Request("yAct") 
yID = RequestS("yID","",48) : If yID="" Then yID="1"
yMod = "rXItem"&yID 'RequestS("yMod","",48) : If yMod="" Then yMod="rXItem1"
mCode = RequestS("mCode","",48)
mName = RequestS("mName","",96)
mUrl = RequestS("mUrl","",255)
mTop = RequestS("mTop","",8) : If mTop="" Then mTop="6666"

If yAct="Add" Then
  mTop = "0000"&mTop : mTop = Right(mTop,4)
  sql = " INSERT INTO [WebAdvert] (" 
  sql = sql& "  KeyID,KeyMod" 
  sql = sql& ",InfName,InfCont,InfPara,InfShow"  
  sql = sql& ",LogAddIP,LogAUser,LogATime" 
  sql = sql& ")VALUES(" 
  sql = sql& "  '" & Get_AutoID(24) &"', 'M" & yMod &"'" 
  sql = sql& ", '" & mName &"', '" & mUrl &"', '" & mTop &"', '" & RequestS("mShow","",2) &"'" 
  sql = sql& " ,'"&Get_CIP()&"','"&Session("UsrID")&"','"&Now()&"'"
  sql = sql& ")"
  if rs_Exist(conn,"SELECT KeyID FROM [WebAdvert] WHERE KeyMod='M"&yMod&"' AND InfName='"&mName&"' AND InfCont='"&mUrl&"'") = "YES" then
    msg = "增加失败!"&mName&" 已经存在!"
  elseif Len(mName)=0 Or Len(mUrl)=0 Then
    msg = "空值失败!"
  else
    Call rs_DoSql(conn,sql)
    msg = ""&mName&" 增加成功!"
  end if
ElseIf yAct="Edit" Then
  mTop = "0000"&mTop : mTop = Right(mTop,4)
  sql="UPDATE [WebAdvert] SET InfName='"&mName&"',InfCont='"&mUrl&"',InfPara='"&mTop&"',InfShow='"&RequestS("mShow","",2)&"'"
  sql=sql&" WHERE KeyID='"&mCode&"' " ':Response.Write sql
  Call rs_DoSql(conn,sql)
  msg = ""&mName&" 增加成功!"
ElseIf yAct = "Del" Then
    sql = " DELETE FROM [WebAdvert] WHERE KeyID='"&RequestS("KeyID","",48)&"' "
	Call rs_DoSql(conn,sql)
	Msg = "删除成功!"
ElseIf yAct = "Upd" Then
    dt = Request("dt")
	pf = "../../upfile/sys/config/sys_menu"&yID&".htm"
	Call File_Add2(pf,dt,"UTF-8")
	Msg = "更新成功!"
ElseIf yAct = "Rep" Then
  d1 = Request("sDir1")
  d2 = Request("sDir2")
  If d1<>"" And d2<>"" Then
	'update table1 set field1 = replace(field1, 'a ', 'b ') 
	sql = "SELECT KeyID,InfCont FROM WebAdvert WHERE KeyMod LIKE 'MrX%' "
    rs.Open sql,conn,1,1 
    Do While Not rs.EOF 
      KeyID = rs("KeyID")
      InfCont = rs("InfCont")
	  InfCont = Replace(InfCont,d1,d2)
      sql="UPDATE [WebAdvert] SET InfCont='"&InfCont&"' WHERE KeyID='"&KeyID&"'"
	  Call rs_DoSql(conn,sql)
    rs.MoveNext
    Loop
    rs.Close()
	Msg = "更新成功!"
  Else
	Msg = "请设置目录!"
  End If
End If


''rXTemp1(),rXTemp1X
sTmpA = rs_Val("","SELECT ParRem FROM AdmPara WHERE ParCode='rXTemp"&yID&"'")
sTmpB = rs_Val("","SELECT ParRem FROM AdmPara WHERE ParCode='rXTemp"&yID&"X'")
sMenu = ""

%>
<table width="600" border="0" align="center" cellpadding="1" cellspacing="2">
  <tr>
    <td valign="top"><fieldset style="padding:5px; text-align:center">
        <legend><strong>系统菜单切换</strong> - <a href="?yID=<%=yID%>">刷新</a></legend>
        <%If yID="1" Then%>
        <span class="colF00">头部菜单</span> - <a href="?yID=2">底部菜单</a>
        <%ElseIf yID="2" Then%>
        <a href="?yID=1">头部菜单</a> - <span class="colF00">底部菜单</span>
        <%End If%>
      </fieldset>
      <div style="line-height:8px;">&nbsp;</div>
      <fieldset style="padding:5px;">
        <legend><strong>系统菜单操作</strong></legend>
        <table width="100%" border="0" align="center" cellpadding="0" cellspacing="1">
          <form name="fm01" method="post" action="?">
            <tr>
              <td nowrap>选择</td>
              <td colspan="2"><select name="Opt1" id="Opt1" style="width:180px; " onChange="chOpt1()">
                  <option value="">(请选择)</option>
                  <%=amOpt(yMod)%>
                </select></td>
            </tr>
            <tr>
              <td nowrap>名称</td>
              <td colspan="2"><input name="mName" type="text" id="mName" size="24" maxlength="24"></td>
            </tr>
            <tr>
              <td nowrap>地址</td>
              <td colspan="2"><input name="mUrl" type="text" id="mUrl" size="24" maxlength="60"></td>
            </tr>
            <tr>
              <td nowrap>顺序</td>
              <td colspan="2"><input name="mTop" type="text" id="mTop" value="<%=mTop%>" size="8" maxlength="4" style="width:80px;">
                <select name="mShow" id="mShow" style="width:90px;">
                  <option value="-">本窗口</option>
                  <option value="c" <% If InfShow="c" Then Response.Write("selected")%>>新窗口</option>
                </select></td>
            </tr>
            <tr>
              <td nowrap><input name="yAct" type="hidden" id="yAct" value="Add"></td>
              <td><input name="SubAdd" type="submit" id="SubAdd" value="增加">
                &nbsp;
                <input name="SubEdit" type="submit" id="SubEdit" value="修改" disabled>
                <input type="hidden" name="mCode" id="mCode">
                <input name="yID" type="hidden" id="yID" value="<%=yID%>"></td>
              <td>&nbsp;</td>
            </tr>
          </form>
        </table>
      </fieldset>
      <div style="line-height:8px;">&nbsp;</div>
      <fieldset style="padding:5px; text-align:center">
        <legend><strong>相关提示</strong></legend>
        &nbsp; <span class="colF00"><%=msg%></span> &nbsp;
      </fieldset></td>
    <td width="50%" align="left" valign="top"><fieldset style="padding:5px;">
        <legend><strong>系统菜单管理</strong></legend>
        <div style="width:280px; height:240px; overflow:auto;">
          <table width="100%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="#EEEEEE">
            <%
    sql = " SELECT * FROM [WebAdvert] WHERE KeyMod='M"&yMod&"'"
	sql =sql& " ORDER BY InfPara,KeyID" 
   Set rs=Server.Createobject("Adodb.Recordset")
   rs.Open Sql,conn,1,1
   i = 0
   Do While NOT rs.EOF
		If i mod 2 = 1 Then
		  col = "#ffffff"
		Else
		  col = "#F8F8F8"
		End If
i = i + 1
KeyID = rs("KeyID")
InfName = rs("InfName")
InfCont = Show_Form(Left(rs("InfCont")&"",120))
InfPara = rs("InfPara")
InfShow = rs("InfShow")
iWins = "target='_self'"
If InfShow="c" Then iWins = "target='_blank'"
iTmpA = Replace(sTmpA,"<A ","<A "&iWins)
iTmpA = Replace(sTmpA,"<a ","<a "&iWins)
iTmpA = Replace(iTmpA,"($Name)",InfName)
iTmpA = Replace(iTmpA,"($Url)",InfCont)

sMenu = sMenu&vbcrlf&iTmpA

	  %>
            <tr bgcolor="<%=col%>">
              <td align="right" nowrap><%=i%></td>
              <td align="left"><a href="<%=InfCont%>" target="_blank"><%=InfName%></a></td>
              <td align="center"><a href="<%=InfCont%>" target="_blank"><%=InfPara%></a></td>
              <td align="center"><input class="btnA1" name="SubEdit2" type="button" id="SubEdit2" value="修改" onClick="edSet('<%=KeyID%>','<%=InfName%>','<%=InfCont%>','<%=InfPara%>','<%=InfShow%>')"></td>
              <td align="center"><input class="btnA1" name="SubEdit2" type="button" id="SubEdit2" value="删除" onClick="Del_YN('?yAct=Del&yID=<%=yID%>&KeyID=<%=KeyID%>','确认删除：<%=Show_jsStr(InfName)%>？')"></td>
            </tr>
            <%
  rs.Movenext
  Loop
  rs.Close()

sData = sMenu
sMenu = Replace(sTmpB,"($Menu)",sMenu) 

%>
          </table>
        </div>
      </fieldset></td>
  </tr>
  <form name="Rep" action="?" method="post">
    <tr>
      <td colspan="2" align="left" valign="top"><input name="sDir1" type="text" id="sDir1" value="/u/xtest21/" size="18" maxlength="255">
        -=&gt;
        <input name="sDir2" type="text" id="sDir2" value="/" size="18" maxlength="255">
        <input type="submit" name="button" id="button" value="替换目录">
        &nbsp;(用于重新配置网站目录后...)
        <input name="yAct" type="hidden" id="yAct" value="Rep"></td>
    </tr>
  </form>
  <tr>
    <td colspan="2" align="left" valign="top"><fieldset style="padding:5px;">
        <legend><strong>系统菜单演示</strong>
        <input name="SubEdit3" type="button" id="SubEdit3" value="更新" onClick="dtUpd()">
        </legend>
        <div style="width:580px; height:120px; overflow:auto;"> <%=sMenu%> </div>
      </fieldset></td>
  </tr>
</table>
<div id="sData" style="display:none"><%=sData%></div>
<form name="fmData" action="?" method="post">
  <input name="yID" type="hidden" value="<%=yID%>">
  <input name="yAct" type="hidden" value="Upd">
  <input name="dt" type="hidden" value="">
</form>
<%

Function amOpt(xMod)
Dim s1,a1,t1,a2
s1 = rs_Val("","SELECT ParRem FROM AdmPara WHERE ParCode='"&xMod&"' ")
a1 = Split(s1,vbcrlf)
s1 = "" '清空
For i=0 To uBound(a1)
  t1 = a1(i)
  t1 = Replace(t1,vbcr,"")
  t1 = Replace(t1,vblf,"")
  If t1<>"" Then
    a2 = Split(t1,"|")
	If uBound(a2)>=2 Then
	  If Left(a2(0),5)="(Url)" Then
	    s1 = s1&vbcrlf&"<option value='"&a2(2)&"'>URL-0:"&a2(1)&"</option>"
	  ElseIf Left(a2(0),5)="(Mod)" Then 'InfA999|($Config_Path)page/about.asp?ModID=($ModID)
	    s1 = s1&amMod(a2(1),a2(2))
	  ElseIf Left(a2(0),5)="(Typ)" Then '(Typ)|InfA124|page/about.asp?ModID=($ModID)&TypID=($TypID)&Flag=($Flag)
	    s1 = s1&amTyp(a2(1),a2(2))
	  ElseIf Left(a2(0),5)="(Out)" Then
	    s1 = s1&vbcrlf&"<option value='"&a2(2)&"'>Out-0:"&a2(1)&"</option>"
	  Else
	    Response.Write "<br>"&t1
	  End If
	End If
  End If
Next
amOpt = Replace(s1,"($Config_Path)",Config_Path)
End Function

Function amMod(xMod,xTmp)
Dim t,r,sql,sID,sNM,t2
  t = rs_Val("","SELECT [SysName] FROM AdmSyst WHERE SysID='"&xMod&"'")
  If t<>"" Then
    xTmp = Replace(xTmp,"($ModID)",xMod)
	r = vbcrlf&"<option value='"&xTmp&"'>Mod-1:"&t&"</option>"
  Else
	r = "" 'InfA999|($Config_Path)page/about.asp?ModID=($ModID)
	sql = "SELECT SysID,SysName FROM AdmSyst WHERE SysID LIKE '"&Left(xMod,4)&"%' ORDER BY SysTop, SysID "
    rs.Open sql,conn,1,1 
    Do While Not rs.EOF 
      sID = rs("SysID")
      sNM = rs("SysName")
      t2 = Replace(xTmp,"($ModID)",sID)
	  r = r&vbcrlf&"<option value='"&t2&"'>Mod-2:"&sNM&"</option>"
    rs.MoveNext
    Loop
    rs.Close()
  End If
amMod = r
End Function

Function amTyp(xMod,xTmp)
Dim t,r,sql,sID,sNM,sFL,t2
   ' (Typ)|InfA124|page/about.asp?ModID=($ModID)&TypID=($TypID)&Flag=($Flag)
	sql = "SELECT TypID,TypName,TypNam2 FROM WebTyps WHERE TypMod='"&xMod&"' AND TypDeep=1 ORDER BY TypID "
    rs.Open sql,conn,1,1 ':Response.Write "<br>"&sql
    Do While Not rs.EOF 
      sID = rs("TypID")
      sNM = rs("TypName")
	  sFL = rs("TypNam2")&""
      t2 = Replace(xTmp,"($ModID)",xMod)
	  t2 = Replace(t2,"($TypID)",sID)
	  t2 = Replace(t2,"($Flag)",sFL)
	  r = r&vbcrlf&"<option value='"&t2&"'>Typ-2:"&sNM&"</option>"
    rs.MoveNext
    Loop
    rs.Close()
amTyp = r
End Function


SET rs=Nothing

%>
<script type="text/javascript">
var oFrm = document.fm01;
var oOpt = oFrm.Opt1;
function chOpt1()
{
 iCode = oOpt.value;
 iName = oOpt.options[oOpt.selectedIndex].text;
 oFrm.mName.value = iName.substring(6);
 oFrm.mUrl.value = iCode;
}
function edSet(xCode,xName,xUrl,xTop,xShow)
{
 oFrm.mCode.value = xCode;
 oFrm.mName.value = xName;
 oFrm.mUrl.value = xUrl;
 oFrm.mTop.value = xTop;
 oFrm.mShow.value = xShow;
 oFrm.SubAdd.disabled = true; 
 oFrm.SubEdit.disabled = false; 
 oFrm.yAct.value = "Edit";
}
function dtUpd()
{
 var dFrm = document.fmData;
 var sDat = document.getElementById('sData').innerHTML;
 dFrm.dt.value = sDat; //alert(sDat);
 dFrm.submit();
}
</script>
</body>
</html>
