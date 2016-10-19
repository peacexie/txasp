<!--#include file="fconfig.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Pragma" content="no-cache">
<meta http-equiv="Expires" content="0">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title>附件文件（视频，图片，附件等大文件）选择</title>
<script src="../../sadm/func1/Func_JS.js" type="text/javascript"></script>
<script src="../../inc/home/jsInfo.js" type="text/javascript"></script>
<style type="text/css">
<!--
body, td, th {
	font-size: 12px;
}
body {
	margin: 8px;
}
a:link {
	color: #0000FF;
	text-decoration: none;
}
a:visited {
	text-decoration: none;
	color: #0000FF;
}
a:hover {
	text-decoration: underline;
	color: #FF00FF;
}
a:active {
	text-decoration: none;
	color: #FF0000;
}
.idShow {
  width:240px; height:300px; 
  overflow:hidden;
  line-height:150%;
  position:absolute;
  padding:5px;
  background-color:#FFF;
  border:5px solid #999;
}
.idHidden {
  width:4px; height:0px; overflow:hidden;
  position:absolute;
  padding:0px;
  background-color:;
  border:1px solid #CCC;
}
-->
</style>
</head>
<body>
<%

yRoot = Config_Path&"upfile/"
yPath = Request("yPath") '#dbf# sys
If inStr(lCase(yPath),"#dbf#/")>0 Or inStr(lCase(yPath),"sys/")>0 Then
  Response.Write "Error!"
  Response.End()
End If
If yPath<>"" Then
  yRoot = Replace(yRoot&yPath&"/","//","/")
End If
NowPath = Server.MapPath(yRoot)
ParPath = "" 
DepPath = Len(yPath)-Len(Replace(yPath,"/",""))

If DepPath=0 Then
  ParPath = "<span style='color:gray;'>回跟目录</span>-<span style='color:gray;'>上级目录</span>-"
ElseIf inStr(yPath,"../")>0 Then
  Response.End()
ElseIf DepPath=1 Then
  ParPath = "<A HREF='?'>回跟目录</A>-<span style='color:gray;'>上级目录</span>-"
ElseIf DepPath>=2 AND DepPath<=5 Then
  tPath = Left(yPath,Len(yPath)-1)
  tPos = inStrRev(tPath,"/")
  ParPath = Left(yPath,tPos)
  ParPath = "<A HREF='?'>回跟目录</A>-<A HREF='?yPath="&ParPath&"'>上级目录</A>-"
Else
  Response.End()
End If
'Response.Write yPath

Act = Request("Act")
Dir = Request("Dir")
Fil = Request("Fil")

If Act="Dir" AND Dir<>"" Then
  Call fold_add(Config_Path&"upfile/"&yPath,Dir)
  'Response.Write Act&Dir
ElseIf Act="Del" AND Fil<>"" Then
  fExt = Mid(Fil,InStrRev(Fil,"."),8)
  If NOT inStr(".js.peace!db.peacbak.mdb.asp.",lCase(fExt))>0 Then 'xml.
    Call fil_del(Config_Path&"upfile/"&yPath&"/"&Fil)
  End If
End If

dDir = Mid(Get_yyyymmdd(""),5,4)

%>
<table width="680" border='0' align="center" cellpadding='3' cellspacing='1' bgcolor="#CCCCCC">
  <tr bgcolor="#FFFFFF">
    <th align="center" background="/img/tool/MTop.jpg" bgcolor="#FFFFFF">视频附件选择</th>
    <td colspan="6" align="center" background="/img/tool/MTop.jpg" bgcolor="#FFFFFF">请点击文件即选择并返回</td>
  </tr>
  <tr bgcolor="#FFFFFF">
    <td colspan="3" align="left"><%=yRoot%><span id="fsName" style="color:#00F;"></span></td>
    <td align="left"><div class="idHidden" id="idShow"></div></td>
    <td colspan='3' align="center" nowrap>
	
	<%=ParPath%><span id="ClearFile" onClick="ActDO('',0)" style="cursor:pointer; color:#0000FF;">清除文件</span></td>
  </tr>
  <%
numfld = 0
numfil = 0
sezfld = 0
sezfil = 0
i = 0
SET objFSO = Server.CreateObject("Scripting.FileSystemObject")
IF objFSO.FolderExists(NowPath) THEN 
%>
  <tr bgcolor="#CCCCFF">
    <th nowrap bgcolor="#E0E0E0">文件选择/目录列表</th>
    <th width="15%" align='center' nowrap bgcolor="#E0E0E0">预览</th>
    <th width="8%" align='center' nowrap bgcolor="#E0E0E0">选择</th>
    <th width="10%" align='center' nowrap bgcolor="#E0E0E0">大小[B]</th>
    <th width="18%" align="center" nowrap bgcolor="#E0E0E0">创建时间</th>
    <th width="10%" align="center" nowrap bgcolor="#E0E0E0">删除</th>
  </tr>
  <%
SET objFolder = objFSO.GetFolder(NowPath)
FOR EACH objFile IN objFolder.Files
i = i + 1
if i mod 2 = 0 then
    row_col = "bgcolor=#f4f4f4"
else 
    row_col = "bgcolor=#FFFFFF"
end if
numfld = numfld +1
sizfld = sizfld + objFile.Size
fPName = yRoot&objFile.Name 'Replace(yRoot&objFile.Name,Config_Path,"")
ObjExt = Mid(fPName,InStrRev(fPName,"."),8) 
  If NOT inStr(".js.peace!db.peacbak.mdb.asp.",lCase(ObjExt))>0 Then 'xml.
    jsAct = " onmouseover=""ImgShow(this,'"&fPName&"')"" onmouseout=""ImgHidden()"" "
	If inStr(".jpg.jpeg.gif.",lCase(ObjExt))>0 Then 
	  fSubj = "<img src='"&fPName&"' width='80' height='20' border=0 "
	  fSubj = fSubj& " onload='javascript:setImgSize(this);' "&jsAct&" />"
	  tdAct = jsAct
	Else
	  fSubj = "打开查看"
	  tdAct = ""
	End If
%>
  <tr <%=row_col%>>
    <td onClick="ActDO('<%=fPName%>','<%=objFile.Size%>');" title="选择文件" style="cursor:hand;color:#0000FF; " nowrap><img src="../../img/tool/attfile.gif" width="14" height="14" border="0" align="absmiddle"> <%=objFile.Name%></td>
    <td align='center' nowrap><a href="<%=fPName%>" target="_blank"><%=fSubj%></a></td>
    <td align="center" nowrap style="cursor:hand;color:#0000FF; " title="选择文件" <%=tdAct%> onClick="ActDO('<%=fPName%>','<%=objFile.Size%>');">选择</td>
    <td align='right' nowrap><%=objFile.Size%></td>
    <td align="center" nowrap class="txtSC"><%=objFile.DateLastModified%></td>
    <td align="center" nowrap class="txtSC"><%If yPath<>"" AND inStr(Session(UsrPStr),"{Admin}")>0 Then%>
      <a href="#" onClick="Del_YN('?yPath=<%=yPath%>&Act=Del&Fil=<%=objFile.Name%>','确认删除?小心操作哦！')">删除</a>
      <%Else%>
      <span style="color:#999">删除</span>
    <%End If%></td>
  </tr>
  <%
  End If
NEXT
 %>
  <tr bgcolor='#CCCCCC'>
    <td height="1" colspan="7" nowrap></td>
  </tr>
  <%
FOR EACH objFolder IN objFolder.SubFolders
'If objFolder.Name="_public" Or objFolder.Name=Session("UsrID") Then
i = i + 1
if i mod 2 = 0 then
    row_col = "bgcolor=#f4f4f4"
else 
    row_col = "bgcolor=#FFFFFF"
end if
numfil = numfil +1
sizfil = sizfil + objFolder.Size

 If NOT ( lCase(objFolder.Name)="#dbf#" OR lCase(objFolder.Name)="sys" ) Then
%>
  <tr <%=row_col%>>
    <td nowrap><img src="../../img/tool/folder2.gif" width="21" height="12" border="0" align="absmiddle"><A HREF='?yPath=<%=yPath%><%=fil_fld&objFolder.Name%>/'><%=Left(objFolder.Name,18)%>/</A></td>
    <td align='right' nowrap>&nbsp;</td>
    <td align='right' nowrap>&nbsp;</td>
    <td align='right' nowrap><%=objFolder.Size%></td>
    <td align="center" nowrap class="txtSC"><%=objFolder.DateLastModified%></td>
    <td align="center" nowrap class="txtSC">&nbsp;</td>
  </tr>
  <%
 End If
NEXT
%>
  <tr bgcolor='#CCCCCC'>
    <td colspan='7' nowrap bgcolor="#E0E0E0">&nbsp;文件:<font color="#FF0000"><%=numfld%>[<%=Get_BSize(sizfld)%>]</font> &nbsp;目录:<font color="#FF0000"><%=numfil%>[<%=Get_BSize(sizfil)%>]</font> &nbsp;所有:<font color="#FF0000">[<%=Get_BSize(sizfld+sizfil)%>]</font></td>
  </tr>
  <%ELSE%>
  <tr bgcolor='#CCCCCC'>
    <td colspan='7' nowrap bgcolor="#E0E0E0"> 暂时无 文件/目录列表...请上传!</td>
  </tr>
  <%
END IF
SET objFolder = Nothing
SET objFSO = Nothing
%>
</table>

<div style="line-height:8px;">&nbsp;</div>
<%
If inStr(Session(UsrPStr),"{Admin}")<=0 Then
  sDis = "" '禁止
Else
  sDis = "x-"
End If
%>
<table width="680" border='0' align="center" cellpadding='3' cellspacing='1' bgcolor="#CCCCCC">
  <%If DepPath=1 Then%>
  <form name="ffimg2" id="ffimg2" action="?" method="post">
    <tr bgcolor="#FFFFFF">
      <td nowrap>目录:
        <input name="Dir" type="text" id="Dir" value="<%=dDir%>" size="24" maxlength="12" Xreadonly>
        <input name=Button type=submit id="Button2" value="建立目录" <%=sDis%>disabled>
        <input name="Act" type="hidden" id="Act" value="Dir">
        <input name="yPath" type="hidden" id="yPath" value="<%=yPath%>"></td>
      <td width="180" align="left" nowrap>&nbsp;</td>
    </tr>
  </form>
  <%End If%>
  <%If DepPath>=2 Then%>
  <form name="ffimg1" id="ffimg1" action="file_up.asp" enctype="multipart/form-data" method="post">
    <tr bgcolor="#FFFFFF">
      <td colspan="2" nowrap> 上传:
        <input name='ImgName' type='file' id="ImgName" style="width:280px; ">
        <input name="yPath" type="hidden" id="yPath" value="<%=yPath%>">
        <select name="ImgAuto" id="ImgAuto">
          <option value="AutoName">自动命名</option>
          <option value="KeepOrg">原文件名</option>
        </select>
        <input name="btUpload" type=submit id="btUpload" value="上传" <%=sDis%>disabled></td>
    </tr>
  </form>
  <%End If%>
  <tr bgcolor="#FFFFFF">
    <td colspan='2' nowrap>注意：上传和删除文件，仅超管可用；大文件请用FTP上传，可参考[管理帮助] 或 [<a href="../../upfile/readme.txt" target="_blank">文件目录规划</a>] 相关文件；</td>
  </tr>
</table>

<script type="text/javascript">

var idImg = document.getElementById("idShow");

function ImgHidden()
{
	idImg.innerHTML = ''; 	
	idImg.className="idHidden";
}

function ImgShow(td,url)
{
	idImg.innerHTML = '图片预浏:<br>'+url+'<br><img src="' + url + '" border="0" />'; 
	idImg.className="idShow";
}

function chkData(){
  if(document.ffimg1.fName.value.length==0){
	 alert("请浏览文件..."); 
  } else {
	 document.ffimg1.submit(); 
  }
}

function ActShow(owFile){
  var fnPos = owFile.lastIndexOf('/')+1; 
  var fsSub = owFile.substring(fnPos,owFile.length);
  document.getElementById('fsName').innerHTML = fsSub;	
}


function ActDO(owFile,owSize){
	
  var fnPos = owFile.lastIndexOf('/')+1; 
  var fsSub = owFile.substring(fnPos,owFile.length);
  
  try{ // For ActFCK( owFile );
	window.top.opener.SetUrl(owFile) ;
	window.top.close() ;
	window.top.opener.focus() ;
  }catch(ex1){ 
  try{ // For Select Sin File
	window.opener.yFile.value = owFile;
  }catch(ex2){
  try{ // For Select Vdo
	window.opener.xFile = owFile;
	window.opener.xSize = owSize;
    window.opener.owFileGet();
  }catch(ex3){
  try{ // For CK3
	window.opener.CKEDITOR.tools.callFunction(2, owFile);
	window.close();
  }catch(exEnd){
    alert("无父窗口!");
    ActShow(owFile);
    return;
  } //ex1
  } //ex2
  } //ex3
  } //exEnd
  window.close();
  
}

</SCRIPT>
</body>
</html>
