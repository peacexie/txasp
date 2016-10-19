<!--#include file="config.asp"-->
<%

tmpCont = rs_Val("","SELECT ParRem FROM AdmPara WHERE ParCode='tmp"&ModID&"'") 

ModImgCount = Eval(ModID&"ImgCount")
ModImgCount = RequestSafe(ModImgCount,"N",1)

prtID = Left(rs_AutID(conn,ModTab,"KeyID",upPart,"1",""),22)
codID = Get_FmtID("mdhnsx","")&"-"&Rnd_ID("KEY",4)
Dim Si,Ui :Ui="02" 
For i=1 To 240 ',96 
  Ui = Next_ID(Ui,"02",3)
  Si = Si&Ui&"|" 
Next

%>
<html>
<head>
<meta http-equiv="Pragma" content="no-cache">
<meta http-equiv="Expires" content="0">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title><%=Config_Name%>后台管理中心</title>
<link href="../../inc/adm_inc/adm_style.css" rel="stylesheet" type="text/css">
<script src="/sadm/func1/Func_JS.js" type="text/javascript"></script>
<script type="text/javascript" src="../../inc/home/jsInfo.js"></script>
<style type="text/css">
.iBox {
	border:1px solid #EEEEEE;
	padding:0px 0px 0px 0px;
	margin:0px 2px 5px 2px;
}
</style>
</head>
<body>
<%


If get_ModCopy(ModID) Then
  Response.Write js_Alert("注意：此类别(模块)信息，不需要再添加；\n只需要在列表页执行同步,并作适当编辑即可","Redir","../info/info_list.asp?ModID="&ModID) 
  Response.End()
End If
If ModImgCount<>1 Then
  Response.Write js_Alert("注意：此类别(模块)信息，需要的图片个数为("&ModImgCount&")个；\n本程序不适合！","Redir","../info/info_list.asp?ModID="&ModID) 
  Response.End()
End If

ModName = rs_Val("","SELECT SysName FROM AdmSyst WHERE SysID='"&ModID&"'")

%>
<table width="99%" border="0" align="center" cellpadding="0" cellspacing="1" style="border:1px solid #CCC; margin-top:2px;">
  <tr>
    <td style="padding-left:60px;"><div style="float:right"> <a href="?ModID=<%=ModID%>">重载本页</a> <a href="imp_set.asp?ModID=<%=ModID%>">信息采集</a> </div>
      <strong> <%=ModName%> 批量上传</strong> <%=bakID%></td>
    <td width="32%" rowspan="5" align="left" valign="top" style="padding:5px; background-color:#F0F0F0">说明：<br>
      <span class="col00F">***0.</span> 本 程序受启发于<a href="http://www.babytree.com/" target="_blank">宝宝树</a>照片批量上传而制作；<br>
      <span class="col00F">***1.</span> 本程序适合于：<span class="colF0F">有且仅有一个附图</span>(缩略图)，<span class="colF0F">详细内容为空或暂时不填</span>，且除了标题类别参数外<span class="colF0F">其它不需要设置</span>的一组图片批量上传；<br>
      <span class="col00F">***2.</span> 请先设置类别，再浏览图片；可用下方的(+n)按纽增加n个图片项目；一次最多可设置96个图片批量上传。<br>
      <span class="col00F">***3.</span> 本程序为增值程序，免费使用；请不要苛求它的功能；如不能满足您的需要，请用<a href="../info/info_add.asp?ModID=<%=ModID%>"><span class="col00F">普通方式添加资料</span></a>。<br>
      <span class="col00F">***4.</span> 建议把要上传的文件，放在同一文件夹中，<span class="colF0F">用标题作为图片名</span>(默认情况下,本系统把文件名作为信息的标题)；其文件名（除后缀外），不能用空格引号点等特殊字符； 建议全用英文半角的字母，数字或下划线；除图片名可用中文外，目录建议也不要用中文。
      <div id="msgBox" style="padding:1px; background-color:#FFFFCC"></div>
      </td>
  </tr>
  <tr>
    <td style="padding-left:120px">类别：
      <select name="InfType" id="InfType">
        <%=Get_TypeOpt(ModID,InfType)%>
      </select></td>
  </tr>
  <tr>
    <td id="batPics"></td>
  </tr>
  <tr>
    <td><div style="float:right">
        <input type="submit" name="btmSend" id="btmSend" value="确认提交" onClick="sendForms()">
      </div>
      <input type="button" name="btnA5" id="btnA5" value="+16" onClick="insBox(16)">
      <input type="button" name="btnA4" id="btnA4" value="+8" onClick="insBox(8)">
      <input type="button" name="btnA3" id="btnA3" value="+4" onClick="insBox(4)">
      <input type="button" name="btnA2" id="btnA2" value="+2" onClick="insBox(2)">
      <input type="button" name="btnA1" id="btnA1" value="+1" onClick="insBox(1)">
      <a href="?ModID=<%=ModID%>">重载本页</a></td>
  </tr>
</table>
<script type="text/javascript">

var Si = "<%=Si%>"; 
var Ai = Si.split("|"); 
var Ni=0; 

var sendNO = 0;
var sendOK = 0;
var sendNull = 0;
function sendForms()
{ 
  getElmID("btmSend").disabled = true;
  for(i=1;i<=5;i++){ getElmID("btnA"+i).disabled = true; }
  sendNO++; i = sendNO; 
  if(sendNO<=Ni) { 
	de = 13; if(chkBox(Ai[i])) { de=300; }
	setTimeout("sendForms()",de); // setInterval
  }else{
	sendNG = Ni-sendOK-sendNull;
	sendMsg = " 共 ["+sendOK+"] 个图片上传完毕！";
	if(sendNull>0) {sendMsg +=" <br>["+sendNull+"] 个项目被忽略(移除); ";}
	if(sendNG>0) {sendMsg +=" <br>["+sendNG+"] 个空项目未提交! ";}
	getElmID("msgBox").innerHTML = "<span class='col00F'>***5.</span>"+sendMsg;
	//alert(sendMsg);
  }
}
function chkBox(id)
{ 
  try{ 
    var sDoc = window.document.getElementById("iFrame"+id).contentWindow.document;
	var sFrm = sDoc.getElementById("iForm<%=Show_jsKey(prtID)%>"); 
	var sImg = sDoc.getElementById("ImgName1"); 
	//alert(sImg.value);
	if(sImg.value.length==0) {  getElmID("iBox"+id).innerHTML = " <div style='padding:3px 0px 0px 24px'><span class='colF0F'>空项目，未提交！</span></div>"; }
	else { sFrm.submit(); sendOK++; }
	return true;
  }catch(err){  
    sendNull++;
	return false;
  }
  
}


tmp = "";
tmp += "<div id='iBox_ID' class='iBox'>";
tmp += "<IFRAME id='iFrame_ID' src='bat_form.asp(Paras)' frameBorder=0 width='560' scrolling='no' height='50'></IFRAME>";
tmp += "</div>";
function insBox(n)
{
  var fmsObj = getElmID("batPics");
  for(i=1;i<=n;i++)	
  {
	j = i+Ni; if(j>96){ alert("最多96个"); return; }
	var p = "?no="+j;
	p += "&id1="+'<%=prtID%>';
	p += "&id2="+Ai[j];
	p += "&ModID=<%=ModID%>";
	p += "&codID=<%=codID%>";
	p += "&InfType="+getElmID("InfType").value;
	p += "&Now=<%=Now()%>";
	p += "&ModName=<%=ModName%>";
	ti = tmp.replace("(Paras)",p);
	ti = ti.replace(/_ID/g,Ai[j],"gi");
	var ep = getElmID("batPics");
	eBox = document.createElement("div");
	eBox.id = Ai[j];
	eBox.innerHTML = ti;
	fmsObj.appendChild(eBox); 
  }
  Ni=Ni+n; //alert(sn);
}
insBox(8);


</script>
</body>
</html>
