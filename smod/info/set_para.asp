<!--#include file="config.asp"-->
<html>
<head>
<meta http-equiv="Pragma" content="no-cache">
<meta http-equiv="Expires" content="0">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title><%=Config_Name%>后台管理中心</title>
<link href="../../inc/adm_inc/adm_style.css" rel="stylesheet" type="text/css">
<script type="text/javascript" src="../../inc/home/jsInfo.js"></script>
<style type="text/css">
tr, td {
	background-color:#FFF;
}
</style>
</head>
<body style="padding:8px">

<%
  Dim itm10(96)
  fmid = Request("fmid")
%>

<table border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="#CCCCCC">
  <tr>
    <td>项目</td>
    <td>选项(每行一个)</td>
    <td>+价格($)</td>
    <td>      +重量(Kg)</td>
  </tr>
  <tr>
    <td><input name="itm1001" type="text" id="itm1001"  value="<%=itm10(01)%>" size="12" maxlength="120"></td>
    <td><textarea name="itm1002" cols="24" rows="6" id="itm1002"><%=itm10(02)%></textarea></td>
    <td><textarea name="itm1003" cols="18" rows="6" id="itm1003"><%=itm10(03)%></textarea></td>
    <td><textarea name="itm1004" cols="18" rows="6" id="itm1004"><%=itm10(04)%></textarea></td>
  </tr>
  <tr>
    <td><input name="itm1006" type="text" id="itm1006"  value="<%=itm10(06)%>" size="12" maxlength="120"></td>
    <td><textarea name="itm1007" cols="24" rows="6" id="itm1007"><%=itm10(07)%></textarea></td>
    <td><textarea name="itm1008" cols="18" rows="6" id="itm1008"><%=itm10(08)%></textarea></td>
    <td><textarea name="itm1009" cols="18" rows="6" id="itm1009"><%=itm10(09)%></textarea></td>
  </tr>
  <tr>
    <td><input name="itm1011" type="text" id="itm1011"  value="<%=itm10(11)%>" size="12" maxlength="120"></td>
    <td><textarea name="itm1012" cols="24" rows="6" id="itm1012"><%=itm10(12)%></textarea></td>
    <td><textarea name="itm1013" cols="18" rows="6" id="itm1013"><%=itm10(13)%></textarea></td>
    <td><textarea name="itm1014" cols="18" rows="6" id="itm1014"><%=itm10(14)%></textarea></td>
  </tr>
  <tr>
    <td><input name="itm1016" type="text" id="itm1016"  value="<%=itm10(16)%>" size="12" maxlength="120"></td>
    <td><textarea name="itm1017" cols="24" rows="6" id="itm1017"><%=itm10(17)%></textarea></td>
    <td><textarea name="itm1018" cols="18" rows="6" id="itm1018"><%=itm10(18)%></textarea></td>
    <td><textarea name="itm1019" cols="18" rows="6" id="itm1019"><%=itm10(19)%></textarea></td>
  </tr>
  <tr>
    <td><input name="itm1021" type="text" id="itm1021"  value="<%=itm10(21)%>" size="12" maxlength="120"></td>
    <td><textarea name="itm1022" cols="24" rows="6" id="itm1022"><%=itm10(22)%></textarea></td>
    <td><textarea name="itm1023" cols="18" rows="6" id="itm1023"><%=itm10(23)%></textarea></td>
    <td><textarea name="itm1024" cols="18" rows="6" id="itm1024"><%=itm10(24)%></textarea></td>
  </tr>
  <tr>
    <td><input name="itm1026" type="text" id="itm1026"  value="<%=itm10(26)%>" size="12" maxlength="120"></td>
    <td><textarea name="itm1027" cols="24" rows="6" id="itm1027"><%=itm10(27)%></textarea></td>
    <td><textarea name="itm1028" cols="18" rows="6" id="itm1028"><%=itm10(28)%></textarea></td>
    <td><textarea name="itm1029" cols="18" rows="6" id="itm1029"><%=itm10(29)%></textarea></td>
  </tr>
  <tr>
    <td colspan="4" align="right"><a href="#" onClick="retPara()">返回&gt;&gt;</a>&nbsp;</td>
  </tr>
</table>

<script type="text/javascript">

// <a href="#" onClick="setPara()">大窗口模式</a>
// parent: function setPara(){ window.open("set_para.asp?fmid=<%=sys27_RVal%>","setPara"); }

var dPar = window.opener.document; //parent.document;
var pfm = dPar.fm01<%=fmid%>;
//var flagPage = new RegExp('<div style="page-break-after: always"></div>','i')

function retPara()
{
 for(i=1;i<96;i++){
  j = 1000+i;
  try{
    iv = document.getElementById("itm"+j).value;
	for(k=1;k<12;k++){ iv = iv.replace("\r\n","@"); }
    dPar.getElementById("itm"+j).value = iv;
  }catch(er) {}
 }
 window.close();	
} 

setPara();
function setPara()
{
 for(i=1;i<96;i++){
  j = 1000+i;
  try{
    iv = dPar.getElementById("itm"+j).value;
	for(k=1;k<12;k++){ iv = iv.replace("@","\r\n"); }
    document.getElementById("itm"+j).value = iv;
  }catch(er) {}
 }
} 



    //iv = dPar.getElementById("itm1001").value;
	//alert(iv);
    //document.getElementById("itm1001").value = iv;

</script>
</body>
</html>
