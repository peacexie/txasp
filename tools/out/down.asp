<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="../../sadm/func1/func1.asp"-->
<!--#include file="../../sadm/func2/Func2.asp"-->
<!--#include file="../../sadm/func1/func_opt.asp"-->
<!--#include file="../../sadm/func1/func_file.asp"-->
<!--#include file="../../sadm/func2/upremote.asp"-->
<%
Call Chk_Perm1("SysTools","") 
act = Request("act") 'act=DownOne
InfUrl = RequestS("InfUrl",3,255) 
InfSave = RequestS("InfSave",3,48)
If InfSave=""  Then 
  InfSave=Get_yyyymmdd("")
  InfSave=Left(InfSave,4)&"-"&Right(InfSave,4)
End If

'RemoteSaveFile(Config_Path&"upfile/tools/down/2010-1005/12.swf", iUrl, 9600, "(.jpg.gif.jpeg.png.tif.swf.flv.htm.html.txt)")
'http://www.pconline.com.cn/pcedu/specialtopic/vcpp/18.swf '\n

If act="DownOne" Then

  pRoot = Config_Path&"upfile/tools/down/"
  Call fold_add(pRoot,InfSave) 
  fName = Mid(InfUrl,inStrRev(InfUrl,"/")+1) 
  fNBak = fName
  fName = pRoot&InfSave&"/"&fName
  
  If RemoteSaveFile(fName, InfUrl, 96000, "(.jpg.gif.jpeg.png.tif.swf.flv.htm.html.txt)") Then
    f1 = "OK"
  Else
    f1 = "NG"
  End If
  
  'Response.Write fName
  Response.Write "---("&f1&"):"&fNBak
  Response.End()

ElseIf act="GetUrls" Then

  InfCSet  = RequestS("InfCSet",3,24)
  InfP1Str = RequestS("InfP1Str",3,255)
  InfP1    = RequestS("InfP1","N",0)
  InfP2Str = RequestS("InfP2Str",3,255)
  InfP2    = RequestS("InfP2","N",0)
  InfRep1  = RequestS("InfRep1",3,255)
  InfRep2  = RequestS("InfRep2",3,255)
  If InfP1Str="" Then InfP1Str="<body"
  If InfP2Str="" Then InfP2Str="</body"
  
  If act="GetUrls" And InfUrl<>"" Then

    sHtml=OutPage(InfUrl,InfCSet)&"" : sHbak = sHtml
    p1=OutSPos(sHtml,InfP1Str,InfP1) 
    p2=OutSPos(sHtml,InfP2Str,InfP2)
    If p1>0 And p2-p1>0 Then
      sHtml=Mid(sHtml,p1,p2-p1)
    Else
      sHtml="(Error)"
    End If
    If InfRep1<>"" Then
      sHtml = Get_Rep(sHtml,InfRep1,InfRep2)
    End If
    sHtml = Get_HLinks(sHtml,"<a[^>]*(href=[^>]*)[^>]*>([^<]*|<img[^<]*>)</a>")
	a = Split(sHtml,"||") : sUrl="" 
	For i=0 To uBound(a)
	  If a(i)<>"" Then
	  iUrl = Get_1Url(a(i),"href=") 
	  sUrl = sUrl&vblf&iUrl
	  End If
	Next

  End If

End If



%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>批量下载</title>
<script src="../../sadm/func1/Func_JS.js" type="text/javascript"></script>
<script src="../../inc/home/jsInfo.js" type="text/javascript"></script>
<style type="text/css">
body, td, th {
	font-size: 12px;
}
select {
	width:100px;
}
.DivTest {
	width:180px;
	height:18px;
	display:block;
	padding:3px;
	float:left;
	border:1px solid #CCCCCC;
	margin:1px;
}
.itmNO {
	width:30px;
	display:inline-block;
	text-align:center;
	border:1px solid #06F;
	padding:1px;
	margin:1px;
}
body {
	margin: 5px;
}
</style>
</head>
<body>
<table width="640" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="#CCCCCC">
  <form id="fmDownBat" name="fmDownBat" method="post" action="?">
    <tr>
      <td width="15%" align="center" bgcolor="#FFFFFF"><span id="RMsg">Url</span>: </td>
      <td bgcolor="#FFFFFF"><input name="InfUrl" type="text" id="InfUrl" value="<%=InfUrl%>" size="48" maxlength="120" xxcus="iMsg(51)" />
        <select name="InfCSet" id="InfCSet">
          <%=Get_SOpt("UTF-8;gb2312;big5;iso-8859-1","",InfCSet,"")%>
        </select></td>
    </tr>
    <tr id="R999" style="display:none">
      <td align="center" bgcolor="#FFFFFF">No:</td>
      <td bgcolor="#FFFFFF"><input name="InfN1" type="text" id="InfN1" value="<%=InfN1%>" size="6" />
        ~
        <input name="InfN2" type="text" id="InfN2" value="<%=InfN2%>" size="6" />
        &nbsp;&nbsp;Ext:
        <input name="InfExt" type="text" id="InfExt" value="<%=InfExt%>" size="6" />
        <input name="act2" type="hidden" id="act3" value="Down999" /></td>
    </tr>
    <tr id="RCom1">
      <td align="center" bgcolor="#FFFFFF">Pos1:</td>
      <td bgcolor="#FFFFFF"><input name="InfP1Str" type="text" id="InfP1Str" value="<%=InfP1Str%>" size="48" maxlength="48" />
        <input name="InfP1" type="text" id="InfP1" value="<%=InfP1%>" size="6" xxcus="iMsg(12)" /></td>
    </tr>
    <tr id="RCom2">
      <td align="center" bgcolor="#FFFFFF">Pos2:</td>
      <td bgcolor="#FFFFFF"><input name="InfP2Str" type="text" id="InfP2Str" value="<%=InfP2Str%>" size="48" maxlength="48" />
        <input name="InfP2" type="text" id="InfP2" value="<%=InfP2%>" size="6" xxcus="iMsg(14)" /></td>
    </tr>
    <tr id="RRep" style="display:none">
      <td align="center" bgcolor="#FFFFFF">Replace.:</td>
      <td nowrap="nowrap" bgcolor="#FFFFFF"><input name="InfRep1" type="text" id="InfRep1" value="<%=InfRep1%>" size="48" maxlength="120" />
        <br />
        <input name="InfRep2" type="text" id="InfRep2" value="<%=InfRep2%>" size="48" maxlength="120" />
        <input name="act" type="hidden" id="act" value="GetUrls" /></td>
    </tr>
    <tr>
      <td align="center" bgcolor="#FFFFFF">Action: </td>
      <td bgcolor="#FFFFFF"><div style="float:right">
          <input name="InfR1" type="radio" id="radio" value="comm" checked="checked" onclick="setOpt(1)" />
          Comm
          &nbsp;
          <input type="radio" name="InfR1" id="radio2" value="Rep" onclick="setOpt(2)" />
          Rep
          &nbsp;
          <input type="radio" name="InfR1" id="radio3" value="999" onclick="setOpt(3)" />
          999 </div>
        <input type="button" name="btmSend" id="btmSend" value="获取地址" onclick="GetUrls()" />
        <input type="button" name="btmList" id="btmList" value="生成地址" onclick="GetList()" style="display:none" />
        &nbsp;或在以下框中输入地址，每行一个</td>
    </tr>
    <tr>
      <td align="center" bgcolor="#FFFFFF">View</td>
      <td bgcolor="#FFFFFF"><textarea name="UrlList" cols="56" rows="12" id="UrlList" wrap="off"><%=sUrl%></textarea></td>
    </tr>
    <tr>
      <td align="center" bgcolor="#FFFFFF">Save: </td>
      <td bgcolor="#FFFFFF"><div style="float:right"> <a href="?<%=Now()%>">Reload</a> </div>
        /upfile/tools/down/
        <input name="InfSave" type="text" id="InfSave" value="<%=InfSave%>" size="12" />
        <input type="button" name="btmDown" id="btmDown" value="批量下载" onclick="downCheck()" /></td>
    </tr>
    <tr>
      <td align="center" bgcolor="#FFFFFF">Msg</td>
      <td bgcolor="#FFFFFF"><div> <span id="msgSend">[None]</span> <span id="msgEnd"></span> <span id='msgTest'> --- (msgTest)</span> </div>
        <div id="bugList"></div></td>
    </tr>
    <tr>
      <td align="center" bgcolor="#FFFFFF">注意</td>
      <td bgcolor="#FFFFFF"> 1. 本程序主要用于批量文件下载(如教程,文章等),支持格式为( .jpg .gif .jpeg .png.tif .swf .flv .htm .txt)；单个文件最大为(9.6MB)；要想下载上百M的视频等文件,不在本程序考虑范围内；<br />
        2. 
        使用后请清理(删除)<a href="../../smod/file/file_view.asp?yPath=tools/down/" target="_blank">下载的文件</a>,因使用不当而造成损失,本程序开发人员不负任何责任；</td>
    </tr>
  </form>
</table>
<script type="text/javascript">
var fm = document.fmDownBat;
var sUrls = "";

function GetList(){
  n1 = fm.InfN1.value;
  n2 = fm.InfN2.value;
  bUrl = fm.InfUrl.value;
  ext = fm.InfExt.value;
  sUrls = "";
  for(i=n1;i<=n2;i++){
	iUrl = bUrl+i+"."+ext;  
	sUrls += iUrl+"\n";
  }
  if(sUrls.length<12) sUrls="";
  fm.UrlList.value = sUrls;
}

function GetUrls(){
  fm.act.value = "GetUrls";
  fm.submit();
}

function downCheck(){
  getElmID("btmDown").disabled = true;
  sUrls = fm.UrlList.value; sUrls = sUrls.replace("\r","","gi");
  aUrls = sUrls.split("\n"); sUrls = "";
  var sItems = "", sTrys = "";
  for(i=0;i<aUrls.length;i++){
    if(aUrls[i].length>12) { 
	  sUrls += "\n"+aUrls[i];
	  sItems += "<span class='itmNO' id='itmID"+i+"'>"+i+"</span>";
	  sTrys += "<span id='tryID"+i+"'>0</span>, ";
	}
  } //alert(j); itmNO
  if(sUrls.length>12)
  {
	sUrls = sUrls.substring(1,sUrls.length);
	fm.UrlList.value = sUrls;
	getElmID("bugList").innerHTML = sItems+" --- "+sTrys+i+"";
	downMain();
  }
}

var sendOK = 0;
var sendNG = 0;
var nDelay = 500;
var nDoMax = 3;
var nTrys  = 10;
var TstUrl = "";
//var nWTime = 0;

function downMain(){
  
  var aUrls = sUrls.split("\n");
  nDoing = 0;
  iDoing = -1;
  nWait = 0;
  for(i=0;i<aUrls.length;i++){
	var iItem = getElmID("itmID"+i);
	var iTry = getElmID("tryID"+i);
	var vTry = 0;
	if(iItem.innerHTML=="...") { 
	  nDoing++; 
	  vTry = parseInt(iTry.innerHTML,10)+1; 
	  if(vTry >= nTrys) { iItem.innerHTML="--"; }
	  iTry.innerHTML = vTry;
	}
	if(iItem.innerHTML=="--") { 
	  nWait++; 
	}
	if(iDoing==-1){
	  var flg = isNaN(parseInt(iItem.innerHTML,10)); 
	  if(!flg) { iDoing=i; }
	}
  } 
  // id ... -- OK NG
  if(nDoing>=nDoMax){    // 等待 (含...)
     setTimeout("downMain()",nDelay);
	 getElmID("msgTest").innerHTML = " (...)等待["+nDoing+"]个！"; 
  }else if(iDoing>-1){  // 提交 (还有没有提交的)
	 downSend(iDoing,aUrls[iDoing]); 
	 setTimeout("downMain()",nDelay);
	 getElmID("msgTest").innerHTML = " (Send)["+iDoing+"]！"; 
  }else if(nWait>0){   // 等待 (含--)
	 setTimeout("downMain()",nDelay);
	 getElmID("msgTest").innerHTML = " (--)等待["+nWait+"]个！"; 
  }else{
	 getElmID("msgSend").innerHTML = "";
	 getElmID("msgEnd").innerHTML = "成功["+sendOK+"]个,失败["+sendNG+"]个！ ";
	 getElmID("msgTest").innerHTML = "<a href='"+TstUrl+"'>Test Link</a>"; 
	 alert('全部处理完'); return;
  }
}

function downSend(k,iUrl){
  getElmID("msgSend").innerHTML = " 第["+k+"]个已提交, ";
  getElmID("itmID"+k).innerHTML = "...";
  var url = "down.asp?act=DownOne";
	url += "&InfSave="+fm.InfSave.value;
	url += "&InfUrl="+iUrl;
	url += "&xRnd=<%=Rnd_ID("",12)%>"+(k+Math.random())+"";
	TstUrl = url; //window.open(url);
  var iHttp = getXmlHttp(); 
  iHttp.open("GET", url, true); 
  iHttp.onreadystatechange = function(){
	if (iHttp.readyState == 4) {
	  var rData = iHttp.responseText;
	  if(rData.indexOf("--(OK):")>0) { 
		sendOK++;
		getElmID("msgEnd").innerHTML = " 第["+k+""+rData+"]个已处理完;  ";
		getElmID("itmID"+k).innerHTML = "OK";
	  }else{
		sendNG++;
		getElmID("itmID"+k).innerHTML = "NG";
	  }
	}
  }
  iHttp.send(null); 
}

function setOpt(n){
  if(n==1){
	getElmID("RMsg").innerHTML = " Url ";
	setShow("btmSend;btmList;R999;RCom1;RCom2;RRep","Y;;;Y;Y;");
  }
  if(n==2){
	getElmID("RMsg").innerHTML = " Url ";
	setShow("btmSend;btmList;R999;RCom1;RCom2;RRep","Y;;;Y;Y;Y");
  }
  if(n==3){
	getElmID("RMsg").innerHTML = " Base Url ";
	setShow("btmSend;btmList;R999;RCom1;RCom2;RRep",";Y;Y;;;");
  }
}

function setShow(Elms,Flgs){
 aElm = Elms.split(";");
 aFlg = Flgs.split(";");
 for(var i=0;i<aElm.length;i++){ 
   var idElm = getElmID(aElm[i]);
   if (aFlg[i]=="Y"){
	  idElm.style.display='';
	  idElm.style.visibility='visible';
	}else{
	  idElm.style.display='none';
	  idElm.style.visibility='hidden';
   }
 }
}

</script>
</body>
</html>
