<!--#include file="_config.asp"-->
<%

Set rs=Server.Createobject("Adodb.Recordset")

Function getTypeList(xMod,xUser)
 Dim sql,TypID,TypName,s :s=""
 sql = "SELECT TypID,TypName FROM SmsType WHERE TypMod='"&xMod&"' "
 If xUser<>"" Then sql=sql&" AND LogAUser='"&xUser&"' "
 rs.Open sql,conn,1,1 
 Do While NOT rs.EOF
   TypID = rs("TypID")
   TypName = rs("TypName")
   s=s&"<a href='?tMod="&xMod&"&tTyp="&TypID&"&PrmFlag="&PrmFlag&"'>"&TypName&"</a> | "
 rs.MoveNext
 Loop
 rs.Close()
 getTypeList = s
End Function

Function getShowSign(xStr)
  Dim sqlS,s
  If PrmFlag="(Mem)" Then
    sqlS = " TypMod='userSign' And LogAUser='"&Session("MemID")&"' "
  Else
    sqlS = " TypMod='sysSign' "
  End If
  'Response.Write sqlS
  s = rs_Val(conn,"SELECT TypName FROM SmsType WHERE "&sqlS)
  xStr = Replace(xStr,"(Sign)","("&s&")")
  xStr = Replace(xStr,"(Sign","("&s&"")
  getShowSign = xStr
End Function

tMod = RequestS("tMod",3,48) :If tMod="" Then tMod="sysTemp"
tTyp = RequestS("tTyp",3,48)
tPrm = RequestS("tPrm",3,48)
  sqlK = " TmpMod='"&tMod&"' "
If tTyp<>"" Then 
  sqlK = sqlK&" And TmpType='"&tTyp&"' "
End If

sType1 = getTypeList("sysTemp","") 
If Session("MemID")&""<>"" Then
  sType2 = getTypeList("userTemp",Session("MemID")) 
Else
  sType2 = "(无用户类别)"
End If

    tName = ""
If tTyp<>"" Then
    tName = rs_Val(conn,"SELECT TypName FROM SmsType WHERE TypID='"&tTyp&"' ")
End If
If tName="" Then
  If tMod="sysTemp" Then
    tName = "(系统范本)"
  Else
    tName = "(用户范本)"
	If Session("MemID")&""<>"" Then
	  sqlK = sqlK&" And LogAUser='"&Session("MemID")&"' "
	End If
	'If Session("UsrID")<>"" And Session("MemID")<>""
  End IF
End If

%>
<html>
<head>
<meta http-equiv="Pragma" content="no-cache">
<meta http-equiv="Expires" content="0">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title>短信模板选择</title>
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
	padding:3px 0px 0px 0px;
	margin:0px;
	border-left:0px;
	border-top:0px;
	border-right:1px solid #999;
	border-bottom:1px solid #999;
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
          <td align="left" class="fntCCC"><a href="?tMod=sysTemp&PrmFlag=<%=PrmFlag%>">系统范本</a>: <%=sType1%><br />
            <a href="?tMod=userTemp&PrmFlag=<%=PrmFlag%>">用户范本</a>: <%=sType2%></td>
        </tr>
      </table></td>
  </tr>
  <%

sql = " SELECT TOP 120 * FROM [SmsTemp] WHERE "&sqlK&" "
sql =sql& " ORDER BY TmpID " 
rs.Open Sql,conn,1,1
i = 0
If NOT rs.EOF Then
Do While NOT rs.EOF
  i = i + 1
  iCode = rs("TmpID")
  iCont = getShowSign(rs("TmpCont"))
%>
  <tr>
    <td width="93%" align="left" id="<%=iCode%>"><%=iCont%></td>
    <td align="center"><input type="button" name="Button" value="选择" class="btnB2" onClick="owMsgUpd('<%=iCode%>');" ></td>
  </tr>
  <%
  rs.movenext
Loop
Else
%>
  <tr>
    <td align="left" colspan="2"><fieldset style="padding:5px;">
        <legend> 提示： </legend>
        无 范本资料 供选择。
      </fieldset></td>
  </tr>
  <%
End If
rs.Close()
%>
  <tr>
    <td align="left" colspan="2" class="bgCCC fnt666">说明：把常用的短信编辑成范本，供发短信时选择，这样提高效率。</td>
  </tr>
</table>
<script type="text/javascript">

function owMsgUpd(id) 
{
	window.opener.tmpCont = document.getElementById(id).innerHTML;
	window.opener.getTemp();
    window.close();		
}
</script>
<%
Set rs = Nothing
%>
</body>
</html>