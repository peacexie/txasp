<!--#include file="stcfg.asp"-->
<%
act = Show_Text(Request("act")) :If act="" Then act="index"
smd = Show_Text(Request("smd")) :If smd="" Then smd="web"
usr = Show_Text(Request("usr")) :If usr="" Then usr="(Guest)"
ver = Show_Text(Request("ver")) :If ver="" Then ver="(unKnow)"
frm = Request.Servervariables("HTTP_REFERER")

j = 0
aCode = Split(Eval("sInfT124Code"),"|")
aName = Split(Eval("sInfT124Name"),"|")
sTime = Timer()

%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Pragma" content="no-cache">
<meta http-equiv="Expires" content="0">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>服务跟踪页@东莞网</title>
<script src="../../inc/home/jsInfo.js" type="text/javascript"></script>
<link rel="stylesheet" type="text/css" href="strace.css"/>
<style type="text/css">
.bdFull {
	border:0px solid #639;
}
.bdTop {
	border-top:1px solid #CCC;
}
.bdBot {
	border-bottom:1px solid #CCC;
}
.noBot {
	border-bottom:1px solid #F0F0FF;
}
.idCom {
	background-color:#F0F0FF;
	cursor:pointer;
}
.idAct {
	background-color:#FFF;
	border:1px solid #639;
	cursor:pointer;
}
.idRed {
	color:#F00;
}
.pdCell {
	padding:8px;
}
.ntSide {
	width:240px;
	background-color:#F0F0F0;
	font-size:12px;
}
</style>
</head>
<body>
<table width="99%" border="0" align="center" cellspacing="5" style="margin:5px; border:1px solid #669">
  <tr>
    <td colspan="2" valign="top" class="bdFull noBot"><table width="100%" border="0">
        <tr>
          <%
		  For i=0 to uBound(aCode)
		    If aCode(i)<>"" Then
			  j = i+1
		  %>
          <th width="16%" class="idCom" id="sw01_id<%=j%>" onMouseOver="msgReady('<%=j%>')" onmousedown="msgShow('<%=j%>')"><%=aName(i)%></th>
          <%
		    End If
		  Next
		  %>
        </tr>
      </table></td>
  </tr>
  <tr>
    <th class="bdBot" id="msgT">东莞网 服务跟踪页</th>
    <td rowspan="2" align="left" valign="top" class="pdCell ntSide" id="msgB">&nbsp;</td>
  </tr>
  <tr>
    <td valign="top" id="msgA" class="pdCell">&nbsp;</td>
  </tr>
</table>
<%
Response.Write "<div style='display:none'>"
For i=0 to uBound(aCode)
  If aCode(i)<>"" Then
	j = i+1
	sCont = rs_Val(conn,"SELECT InfCont FROM InfoNews WHERE InfType LIKE '"&aCode(i)&"%'")
	sCont = Replace(sCont,"($go.reset)",frm&"?act=reset")
	sCont = Replace(sCont,"($usr)",usr)
	sCont = Replace(sCont,"($smd)",smd)
	sCont = Replace(sCont,"($act)",act)
	sCont = Replace(sCont,"($ver)",ver)
	sCont = Replace(sCont,"($frm)",frm)
	sCont = Replace(sCont,"<HR","<hr")
	sCont = Replace(sCont,"<hr>","<hr />")
	sCont = Replace(sCont,"<hr/>","<hr />")
	aCont = Split(sCont&"<hr />","<hr />")
%>
<div id="box<%=j%>A"><%=aCont(0)%></div>
<div id="box<%=j%>B"><%=aCont(1)%></div>
<%
  End If
Next
Response.Write "</div>"

Response.Write "<!--"&Timer()-sTime&"-->"
%>
<script type="text/javascript">
var nMax = <%=j%>;
function msgReady(n){
  for(var i=1;i<=nMax;i++){ 
    swID = getElmID("sw01_id"+i+"");
	if(swID.className!="idAct idRed"){
	  if(n==i){ swID.className="idAct"; }
	  else{ swID.className="idCom"; } 
  } }
}
function msgShow(n){
  for(var i=1;i<=nMax;i++){ 
    swID = getElmID("sw01_id"+i+"");
	if(n==i){
      swID.className="idAct idRed";
	  getElmID("msgT").innerHTML = swID.innerHTML;
	  getElmID("msgA").innerHTML = getElmID("box"+i+"A").innerHTML;
	  getElmID("msgB").innerHTML = getElmID("box"+i+"B").innerHTML;
	}else{
	  swID.className="idCom";
  } }
}
msgShow(1);
</script>
</body>
</html>
