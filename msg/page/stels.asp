<!--#include file="_config.asp"-->
<%

Set rs=Server.Createobject("Adodb.Recordset")
Dim sTyps,sTypt :sTyps = "" :sTypt = ""

Function getTypeList(xMod,xSql)
 Dim TypID,TypName,s :s="" ': Response.Write xSql
 'If inStr(lCase(Left(xSql,8)),"and")>0 Then
   'xSql = 
 'End If
 rs.Open xSql,conn,1,1 '80040e0c, 没有为命令对象设置命令。
 Do While NOT rs.EOF
   TypID = rs(0)
   TypName = rs(1)
   s=s&"<a href='?tMod="&xMod&"&tTyp="&TypID&"'>"&TypName&"</a> | "
   sTyps = sTyps&TypID&";"
   sTypt = sTypt&TypName&";"
 rs.MoveNext
 Loop
 rs.Close()
 getTypeList = s
End Function

Function getTelsList(xSql)
 Dim TypID,TypName,s,si :s="" ': Response.Write xSql
 rs.Open xSql,conn,1,1 '80040e0c, 没有为命令对象设置命令。
 Do While NOT rs.EOF
   TypID = rs(0)&""
   TypName = rs(1)&""
   si = "<div class='itmName' title='"&TypID&"'><input name='itm' type='checkbox' id='itm' value='"&TypID&"["&TypName&"]'>"&TypName&"</div>"
   If Len(TypID)<5 Then
   si = "<div class='itmNamu' title='"&TypID&"'><input name='nul' type='checkbox' id='nul' value='null["&TypName&"]' disabled>"&TypName&"</div>"
   End If
   s=s&vbcrlf&si
 rs.MoveNext
 Loop
 rs.Close()
 If s="" Then s="<span class='fntCCC'>(无资料)</span>"
 getTelsList = s
End Function

tMod = RequestS("tMod",3,48) :If tMod="" Then tMod="SmsTels" '从电话薄
tTyp = RequestS("tTyp",3,48)
tPrm = RequestS("tPrm",3,48)
  sqlK = " "
If tTyp<>"" Then 
  sqlK = sqlK&" And TypID='"&tTyp&"' "
End If

If tMod="SmsTels" Then
  If tPrm="(Mem)" Then '会员自己 userTels
    sqlW1 = " TypMod='userTels' And LogAUser='"&Session("MemID")&"' "&sqlK
    sqlW2 = " TelMod='userTels' And LogAUser='"&Session("MemID")&"' "&sqlK
  Else '管理员 sysTels
    sqlW1 = " TypMod='sysTels' "&sqlK
    sqlW2 = " TelMod='sysTels' "&sqlK
  End If
  sqlT = " SELECT TypID,TypName FROM [SmsType] WHERE "&sqlW1&" Order BY TypID "
  sType = getTypeList(tMod,sqlT) 
  tName = "电话薄"
  sql  = " SELECT TelName,TelNum FROM [SmsTels] WHERE "&sqlW2&" "&sqlK&" Order BY TelType,TelID ASC "
  cTel = "TelNum"
  cName = "TelName"
  cTab = "SmsTels"
  cType = "TelType"
ElseIf tMod="UsrAdmin" Then
  sqlT = " SELECT SysID,SysName FROM AdmSyst WHERE SysType='Admin' ORDER BY SysTop,SysID "
  sType = getTypeList(tMod,sqlT) 
  tName = "管理员"
  sql  = " SELECT UsrName,UsrTel FROM [AdmUser"&Adm_aUser&"] WHERE UsrType LIKE 'Adm%' "&sqlK&" Order BY UsrType,UsrID ASC "
  cTel = "UsrTel"
  cName = "UsrName"
  cTab = "AdmUser"&Adm_aUser&""
  cType = "UsrType"
ElseIf tMod="InnAdmin" Then
  sqlT = " SELECT SysID,SysName FROM AdmSyst WHERE SysType='Inner' ORDER BY SysTop,SysID "
  sType = getTypeList(tMod,sqlT) 
  tName = "内部会员"
  sql  = " SELECT UsrName,UsrTel FROM [AdmUser"&Adm_aUser&"] WHERE UsrType LIKE 'Inn%' "&sqlK&" Order BY UsrType,UsrID ASC "
  cTel = "UsrTel"
  cName = "UsrName"
  cTab = "AdmUser"&Adm_aUser&""
  cType = "UsrType"
ElseIf tMod="SMember" Then
  'sqlT = " SELECT SysID,SysName FROM AdmSyst WHERE SysType='Inner' ORDER BY SysTop,SysID "
  sType = "" 'getTypeList(tMod,sqlT) 
  tName = "短信会员"
  sql  = " SELECT MemName,MemMobile FROM [SmsMember] WHERE 1=1 "&sqlK&" Order BY LogATime DESC "
  cTel = "MemMobile"
  cName = "MemName"
  cTab = "SmsMember"
  cType = "xxxType"
End If 


If tTyp<>"" Then
  sTypt = Get_SOpt(sTyps,sTypt,tTyp,"Val")
  sTyps = tTyp 
End If
sTWhr = " (cType)='aTyps(i)' "
If cType="xxxType" Then 
  sTWhr=" 1=1 "
  sTypt = "短信会员"
  sTyps = "SMember"
Else
  'sTWhr=" 1=1 AND "&sTWhr
End If
'Response.Write sTyps
'Response.Write sTypt

%>
<html>
<head>
<meta http-equiv="Pragma" content="no-cache">
<meta http-equiv="Expires" content="0">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title>短信号码选择</title>
<link href="../inc/spub.css" rel="stylesheet" type="text/css">
<script src="../../sadm/func1/Func_JS.js" type="text/javascript"></script>
<style type="text/css">
.bgCCC {
	background-color:#f0f0f0;
	padding:5px;
}
.btnB2 {
	font-size:12px;
	height:20px;
	padding:1px 0px 0px 0px;
	margin:0px;
	border-left:0px;
	border-top:0px;
	border-right:1px solid #999;
	border-bottom:1px solid #999;
}
input{
	vertical-align:middle;
}
</style>
</head>
<body>
<div style="line-height:8px;">&nbsp;</div>
<table width="610" border="1" align="center" cellpadding="3" cellspacing="5">
  <tr>
    <td align="center" colspan="2"><table width="100%" border="0" cellpadding="2" cellspacing="2" bgcolor="#FFFFFF">
        <tr class="bgCCC">
          <td width="20%" align="center"><strong><%=tName%></strong></td>
          <td align="left"><a href="?tMod=<%=tMod%>">(全部)</a>: <%=sType%></td>
        </tr>
      </table></td>
  </tr>
  <form action="" method="get" name="fItems" id="fItems">
  <%
  aTypt = Split(sTypt,";")
  aTyps = Split(sTyps,";")
  For i=0 To uBound(aTyps)
  If aTyps(i)<>"" Then
    iTWhr = Replace(Replace(sTWhr,"(cType)",cType),"aTyps(i)",aTyps(i)) '(cType)='aTyps(i)'
	'Response.Write iTWhr
  %>
  <tr>
    <td align="left" id="<%=aTyps(i)%>">
      <%=getTelsList("SELECT "&cTel&","&cName&" FROM "&cTab&" WHERE "&iTWhr&" ")%> 
    </td>
    <td width="10%" align="left" nowrap><input type="checkbox" name="grp" id="grp" onClick="setGroup(this,'<%=aTyps(i)%>')"><%=aTypt(i)%></td>
  </tr>
  <%
  End If
  Next
  %>
  <tr>
    <td align="center" colspan="2"><table width="100%" border="0" cellpadding="2" cellspacing="2" bgcolor="#FFFFFF">
        <tr class="bgCCC">
          <td width="50%" align="center"><input type="button" name="Button" value="替换(已经输入的号码)" class="btnB2" onClick="owMsgUpd1();" ></td>
          <td align="center"><input type="button" name="Button2" value="增加(已输入的号码不变)" class="btnB2" onClick="owMsgUpd2();" ></td>
        </tr>
      </table></td>
  </tr>

  </form>
</table>
<script type="text/javascript">

function owMsgUpd1() { 
	var e = document.fItems.itm;
	var s = ""; 
    for(var i=0;i<e.length;i++){
	   if(e.item(i).checked) s += e.item(i).value+"\n"; 
    }
	window.opener.telCont = s;
	window.opener.getTels();
    window.close();		
}
function owMsgUpd2() { 
	var e = document.fItems.itm;
	var s = ""; 
    for(var i=0;i<e.length;i++){
	   if(e.item(i).checked) s += e.item(i).value+"\n"; 
    }
	window.opener.telCont = s;
	window.opener.getTAdd();
    window.close();		
}

function setGroup(e,id){
	v = false; if(e.checked==true) v=true; //alert(f);
	var e2 = document.getElementById(id).getElementsByTagName("input"); 
	for(var i=0;i<e2.length;i++){ 
	  if(e2.item(i).disabled==false) e2.item(i).checked=v;
	}
}


</script>
<%
Set rs = Nothing
%>
</body>
</html>