<!--#include file="iconfig.asp"-->
<%

'//Chk_URL("");
Call Chk_Perm9("MemFEditor","3") ' //MemFEditor ,FileInfo

ID = RequestS("ID",3,48) :pID = Replace(ID,"-","/")
EdtID = RequestS("EdtID",3,48)
pRoot = RequestS("pr",3,48)
pSub = RequestS("ps",3,48)
send = Request("send")

sExt = "(.js.xml.peace!db.peacbak.mdb.php.htm.stx.)"
upRoot = Config_Path&"upfile/" 'upRoot = Server.MapPath("/upfile/")
imgPath = Server.MapPath("/img/") 'idPath = upRoot&Replace(ID,"-","/")&"/"

If pRoot="imgIcon" Then
  pDisk = Config_Path&"img/" 
ElseIf pRoot<>"" Then
  pDisk = upRoot&""&pRoot&"/"
Else
  pDisk = upRoot&""&pID&"/"
End If
pView = pDisk
'Response.Write pDisk&",<br>"&pView 

Function getSubDirs(xPath2)
  Dim xPath,sNull,fso,dir,iFile ,s : s="" 
  xPath = xPath2&"/" :xPath = Replace(xPath,"//","/")
  sNull = "<span class='colCCC'>(无子目录)</span>"
  xPath = Server.MapPath(xPath) ':Response.Write xPath
  SET fso = Server.CreateObject("Scripting.FileSystemObject")
  If fso.FolderExists(xPath) THEN 
	SET dir = fso.GetFolder(xPath)
	  FOR EACH iFile IN dir.SubFolders
		iName = Mid(iFile,InStrRev(iFile,"\")+1,8)
		s=s& "<a href='?EdtID="&EdtID&"&ID="&ID&"&pr="&pRoot&"&ps="&iName&"'>"&iName&"</a>"
	  NEXT
	SET dir = Nothing
  Else
    s=sNull  
  End If
  SET fso = Nothing
  If s="" Then s=sNull
  getSubDirs = s
End Function

Function getFileType(xExt)
  Dim r :r = "unKnow"
  If inStr(fFiles,xExt)>0 Then r="Files"
  If inStr(fImage,xExt)>0 Then r="Image"
  If inStr(fDocus,xExt)>0 Then r="Docus"
  If inStr(fMedia,xExt)>0 Then r="Media"
  If xExt=".swf" Then r="Flash"
  If xExt=".flv" Then r="FLV"
  getFileType = r
End Function

If send="dFile" Then
  dFile = Request("dFile")
  fExt = Mid(dFile,InStrRev(dFile,"."),8)
  If NOT inStr(sExt,lCase(fExt))>0 Then
    Call fil_del(dFile)
  End If
End If
'if($send=='dFile'){ 
  '$dFile = myReq('dFile');
  '$fExt = strrchr($dFile,'.'); 
  'if(strpos($sExt,$fExt)>0){ }
  'else{ unlink($dFile); }  
'}

'//echo "$pDisk,<br>$pView";
  
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Pragma" content="no-cache">
<meta http-equiv="Expires" content="0">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title>附件管理中心</title>
<link href="../../inc/adm_inc/adm_style.css" rel="stylesheet" type="text/css">
<script src="../../inc/home/jsPlugs.js" type="text/javascript"></script>
<script src="../../inc/home/jsInfo.js" type="text/javascript"></script>
<style type="text/css">
table {
	margin:12px 0px 0px 0px;
	border:1px solid #033;
}
td.subLink a {
	width:80px;
	height:18px;
	overflow:hidden;
	text-align:center;
	display:inline-block;
	margin:1px;
}
.idShow {
	width:300px;
	height:200px;
	overflow:hidden;
	line-height:150%;
	position:absolute;
	padding:5px;
	background-color:#FFF;
	border:5px solid #999;
}
.idHidden {
	width:4px;
	height:5px;
	line-height:100%;
	overflow:hidden;
	position:absolute;
	padding:0px;
 background-color:;
	border:0px solid #F0F0F0;
}
</style>
</head>
<body>

<table width="680" border="0" align="center" cellpadding="3" cellspacing="1" bgcolor="#CCCCCC">
  <tr>
    <td width="10%" align="center" bgcolor="#FFFFFF">更改目录</td>
    <td width="10%" align="center" bgcolor="#FFFFFF"><a href="?EdtID=<%=EdtID%>&ID=<%=ID%>&pr=myfile">myFile</a></td>
    <td width="10%" align="center" bgcolor="#FFFFFF"><a href="?EdtID=<%=EdtID%>&ID=<%=ID%>&pr=myftp">myFTP</a></td>
    <td width="10%" align="center" bgcolor="#FFFFFF"><a href="?EdtID=<%=EdtID%>&ID=<%=ID%>&pr=imgIcon">imgIcon</a></td>
    <td align="center" bgcolor="#FFFFFF">&nbsp;</td>
    <td width="10%" align="center" bgcolor="#FFFFFF"><a href="?EdtID=<%=EdtID%>&ID=<%=ID%>&pr=(temp)">内容模版</a></td>
    <td width="10%" align="center" bgcolor="#FFFFFF"><a href="?EdtID=<%=EdtID%>&ID=<%=ID%>&pr=(char)">特殊字符</a></td>
    <td width="10%" align="center" bgcolor="#FFFFFF"><A href="../../upfile/readme.txt" target="_blank">目录规划</A></td>
    <td width="10%" align="center" bgcolor="#FFFFFF"><a href="?EdtID=<%=EdtID%>&ID=<%=ID%>">返回</a></td>
  </tr>
  <tr>
    <td colspan="9" align="left" bgcolor="#FFFFFF" class="subLink">
	<%
	If pRoot="(char)" Then
	  Response.Write "<a href='?EdtID="&EdtID&"&ID="&ID&"&pr=(char)&ps=peace'>Peace方案</a>"
	  Response.Write "<a href='?EdtID="&EdtID&"&ID="&ID&"&pr=(char)&ps=eweb'>eWeb方案</a>"
	  Response.Write "<a href='?EdtID="&EdtID&"&ID="&ID&"&pr=(char)&ps=baidu'>Baidu方案</a>"
	ElseIf pRoot="(temp)" Then
	  Response.Write "<a href='?EdtID="&EdtID&"&ID="&ID&"&pr=(temp)&ps=(def)'>(默认)方案</a>"
	  Response.Write "<a href='?EdtID="&EdtID&"&ID="&ID&"&pr=(temp)&ps=eweb'>eWeb方案</a>"
	Else
	  Response.Write getSubDirs(pDisk)
	End If
	%>
    </td>
  </tr>
</table>

<%
If pRoot="(char)" Or pRoot="(temp)" Then
  If pRoot&pSub="(char)" Then pSub="peace"
%>
  <!--#include file="edt_temp.asp"-->
<%
	idPath = ""
	vrPath = ""
Else
  If pSub<>"" Then
	idPath = pDisk&pSub&"/"
	vrPath = pView&pSub&"/"
  Else
	idPath = pDisk
	vrPath = pView
  End If

%>

<table width="680" border="0" align="center" cellpadding="3" cellspacing="1" bgcolor="#CCCCCC">
  <tr>
    <th colspan="2" align="center" bgcolor="#FFFFFF">选择/插入编辑器</th>
    <th bgcolor="#FFFFFF">附件列表 --- [<%=vrPath%>]</th>
    <th width="10%" align="center" bgcolor="#FFFFFF">打开</th>
    <th width="10%" align="center" bgcolor="#FFFFFF">大小</th>
    <th width="10%" align="center" bgcolor="#FFFFFF">删除</th>
  </tr>
 
  <%
    i = 0 : idPath = Server.MapPath(idPath) ':Response.Write "<br>"&idPath
	If fold_exist(idPath,"") Then
    set fso = CreateObject("Scripting.FileSystemObject")
    set dir = fso.GetFolder(file_fPath(idPath))
    for each f0 in dir.Files
	  i = i + 1
	  fName = f0.Name
	  fExt = Mid(fName,InStrRev(fName,"."),8)
	  fID = Replace(fName,fExt,"_"&Mid(fExt,2))
	  fID = Replace(Replace(fName,".","_"),"-","_")
	  fType = getFileType(lCase(fExt))
	  fullName = vrPath&fName
	  fSize = Get_BSize(f0.Size)
	  If inStr("(.jpg.jpeg.gif.)",fExt)>0 Then 
		jsAct = " onmouseover=""ImgShow(this,'"&fName&"','"&fID&"')"" onmouseout=""ImgHidden('"&fID&"')"" "
	  Else
	    jsAct = ""
	  End If
	  fParas = "'"&fName&"','"&fType&"'" '//url,Type
	  If fType="Image" Then
	    strShow = "<img src='"&fullName&"' width='40' height='20' border=0 onload='javascript:setImgSize(this);' />"
	  Else
	    strShow = fType
	  End If
  %>
  
  <tr>
    <td width="10%" align="center" bgcolor="#FFFFFF"<%=jsAct%>>
      <a href="#" onClick="onInsert(<%=fParas%>)">直接插入</a>
    </td>
    <td width="10%" align="center" bgcolor="#FFFFFF"<%=jsAct%>>
      <%If inStr("(Files,Docus)",fType)>0 Then%>
      <a xref="#" >设置插入</a>
      <%Else%>
      <a href="#setMark" onClick="onSetting(<%=fParas%>)">设置插入</a>
      <%End If%></td>
    <td bgcolor="#FFFFFF"><span class="idHidden" id='<%=fID%>'></span><%=fName%></td>
    <td align="center" bgcolor="#FFFFFF"<%=jsAct%>><a href="<%=fullName%>" target="_blank"><%=strShow%></a></td>
    <td align="right" bgcolor="#FFFFFF"><%=fSize%></td>
    <td align="center" bgcolor="#FFFFFF"><a href="?EdtID=<%=EdtID%>&ID=<%=ID%>&send=dFile&dFile=<%=vrPath&fName%>">删除</a></td>
  </tr>

  <%
	  'end if
	next 
	set dir = Nothing
	set fso = Nothing
	Else
	  If Request("ID")<>"" Then
		sPath = Request("ID")
		sPath = Replace(sPath,"-","/")
	  Else
		tDate = Now()
		sPath = "defup/"&Year(Now)&"/"&DatePart("m",tDate)
	  End If
	  Call fold_add9(Config_Path&"upfile/",sPath,0)	  
	End If
  %>
  <%If i=0 Then%>
  <tr>
    <td colspan="6" align="center" bgcolor="#FFFFFF" class="colF0F"><%=vrPath%><br>
      此目录下暂时无附件!</td>
  </tr>
  <%End If%>
</table>
<%End If%>
<a name="setMark"></a>
<table width="680" border='0' align="center" cellpadding='3' cellspacing='1' bgcolor="#CCCCCC">
  <form name="ffimg1" id="ffimg1" action="file_up.asp" enctype="multipart/form-data" method="post">
    <tr bgcolor="#FFFFFF" id="fmInsert" style="display:none;">
      <td nowrap>设置： 文件
        <input name="fmUrl" type="text" id="fmUrl" maxlength="4" style="width:180px; background-color:#F0F0F0" disabled>
        &nbsp;
        <label for="fmW"></label>
        宽x高:
        <input name="fmW" type="text" id="fmW" style="width:40px" value="550" size="4" maxlength="4">
        x
        <input name="fmH" type="text" id="fmH" style="width:40px" value="400" size="4" maxlength="4">
        &nbsp;对齐
        <label for="fmAlign"></label>
        <select name="fmAlign" id="fmAlign">
          <option value="">(不设置)</option>
          <option value="left">靠左</option>
          <option value="right">靠右</option>
          <option value="center">中间</option>
        </select>
        &nbsp;
        <input name="btUpload2" type=button id="btUpload2" onClick="onSetted()" value="插入">
        <input name="fmType" type="hidden" id="fmType"></td>
    </tr>
    <tr bgcolor="#FFFFFF" id="fmUpload">
      <td nowrap> 上传：
        <input name='ImgName' type='file' id="ImgName" style="width:280px; ">
        <input name="yPath" type="hidden" id="yPath" value="<%=pID%>/">
        <input name="goPage" type="hidden" id="goPage" value="goRef">
        <select name="nAuto" id="nAuto">
          <option value="AutoName">自动命名</option>
          <option value="KeepOrg">原文件名</option>
        </select>
        <input name="btUpload" type=submit id="btUpload" value="上传"></td>
    </tr>
  </form>
  <tr bgcolor="#FFFFFF">
    <td>注意：<span class="colF00" id="SymMessage"></span><br>
      1-1:
      如果第二个表中有附件,可点"直接插入"即可把附件插入到编辑框;<br>
      1-2:
      也可点"设置插入"设置参数后再插入到编辑框,系统自动识别SWF,FLV,图片和其它文件等格式；<br>
      2-1: 如果第二个表中没有附件,可点第一个表中的连接,更改目录浏览附件,再按以上操作；<br>
      2-2: 也可在第三个表中上传文件附件(仅超管可用),大文件请用FTP上传，可参考[管理帮助] 或 [<a href="../../upfile/readme.txt" target="_blank">文件目录规划</a>] 相关文件；</td>
  </tr>
</table>

<script type="text/javascript">

var vrPath = '<%=vrPath%>';

function ImgHidden(id){
	var idImg = document.getElementById(id);
	idImg.innerHTML = ''; 	
	idImg.className="idHidden";
}

function ImgShow(td,url,id){
	var idImg = document.getElementById(id);
	idImg.innerHTML = '<img src="' + vrPath+url + '" border="0" />'; //<br>图片预浏:<br>'+url+'
	idImg.className="idShow";
}

function onInsert(xFile,xType){ //直接插入
	// if(strstr("(Flash,FLV,Media)",$fType)){
	cTab = "(Flash,FLV,Media)";
	if(cTab.indexOf(xType)>0){
	  edtInsert(xFile,'f'+xType);
	}else{
	  edtInsert(xFile,xType);
	}
}
function onSetting(xFile,xType){ //设置...
	if(document.getElementById('fmInsert').style.display=='none'){
	  ShowDiv('fmInsert');
	  ShowDiv('fmUpload');
	}
	document.getElementById('fmUrl').value = xFile;
	document.getElementById('fmType').value = xType;
	if(xType=='Image'){
      img = new Image(); img.src = vrPath+xFile; 
	  document.getElementById('fmW').value = img.width;
	  document.getElementById('fmH').value = img.height;
	}
	alert('请先在下方设置参数!\n设置好后点[插入]');
}
function onSetted(){ //设置完...
	sFile = document.getElementById('fmUrl').value;
	sType = document.getElementById('fmType').value;
	sW = document.getElementById('fmW').value;
	sH = document.getElementById('fmH').value;
	sAlign = document.getElementById('fmAlign').value;
	sParas = ''; 
	if(sW) sParas += " width='"+sW+"' "; 
	if(sH) sParas += " height='"+sH+"' ";
	if(!sAlign=='') sParas += " align='"+sAlign+"' ";
	edtInsert(sFile,sType,sParas);
	ShowDiv('fmInsert');
	ShowDiv('fmUpload'); 
	document.getElementById('fmUrl').value = '';
	document.getElementById('fmType').value = '';
}
function edtInsert(xFile,xType,xParas){ //插入Editor
	sUrl = vrPath+xFile; sParas = '';
	cfgPath = '<%=Config_Path%>';
	if(xParas) sParas = xParas; 
	switch(xType){ 
	  case 'FLV':	
	    sTemp = "<embed src='"+cfgPath+"inc/home/playflv.swf' FlashVars='vcastr_file="+sUrl+"' type='application/x-shockwave-flash' "+sParas+"></embed>";
		sPlay = "<object classid='clsid:D27CDB6E-AE6D-11cf-96B8-444553540000' "+sParas+">";
        sPlay += "<param name='movie' value='"+cfgPath+"inc/home/playflv.swf' />";
        sPlay += "<param name='quality' value='high' />";
        sPlay += "<param name='FlashVars' value='vcastr_file="+sUrl+"' />";
        sPlay += sTemp+"</object>";
		sTemp = sPlay; 
	  break;  
	  case 'Flash':	sTemp = "<embed src="+sUrl+" quality='high' type='application/x-shockwave-flash' "+sParas+"></embed>"; break;  
	  case 'Media':	sTemp = "<embed src="+sUrl+" border='0' "+sParas+" />"; break;  
	  case 'Image':	sTemp = "<img src="+sUrl+" border='0' "+sParas+" />"; break;  
	  case 'Files':	sTemp = "<a href='"+sUrl+"'><img src='"+cfgPath+"img/file/txt.gif' border='0' "+sParas+" /></a>"; break;  
	  case 'Docus':	sTemp = "<a href='"+sUrl+"'><img src='"+cfgPath+"img/file/docus.gif' border='0' "+sParas+" /></a>"; break; 
	  case 'fFLV':   sTemp = "<a href='"+sUrl+"'><img src='"+cfgPath+"img/file/swf.gif' border='0' "+sParas+" /></a>"; break; 
	  case 'fFlash': sTemp = "<a href='"+sUrl+"'><img src='"+cfgPath+"img/file/swf.gif' border='0' "+sParas+" /></a>"; break; 
	  case 'fMedia': sTemp = "<a href='"+sUrl+"'><img src='"+cfgPath+"img/file/mov.gif' border='0' "+sParas+" /></a>"; break; 
	  default:     sTemp = "<a href='"+sUrl+"'><img src='"+cfgPath+"img/file/unknow.gif' border='0' "+sParas+" /></a>";;
	}
	window.opener.apiInsert<%=EdtID%>(sTemp);
	alert('成功插入文件:\n['+xFile+"#"+sTemp+']！');
}

</script>
</body>
</html>
