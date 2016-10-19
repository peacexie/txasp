<!--#include file="../../sadm/func1/func1.asp"-->
<!--#include file="../../sadm/func2/Func2.asp"-->
<!--#include file="../../sadm/func1/func_opt.asp"-->
<!--#include file="../../sadm/func1/func_file.asp"-->
<%

Call Chk_Perm1("SysTools","") 

act = Request("act") 
p = Request("p")
i = Request("i")

InfBase = RequestS("InfBase",3,255) 
InfExts = RequestS("InfExts",3,255)
InfFill = RequestS("InfFill",3,255)
InfSubs = RequestS("InfSubs",3,255)
InfSave = RequestS("InfSave",3,48)
InfCSet = RequestS("InfCSet",3,48)

If InfBase=""  Then 
  InfBase = Server.MapPath("../../")
End If
If InfExts=""  Then 
  InfExts = ".asp|.htm|.js|.vb|.cs|.php|.java|.js"
End If
If InfFill=""  Then 
  InfFill = "editor|pcode"
End If
' bbs\ doc\ inc\ setup\ trade\ vote\
defSubs = "member|page|pfile|sadm|smod|tools"
If RequestS("InfBase",3,255)="" Then
  InfSubs = defSubs
Else
 If InfSubs="" Then 
  InfSubs = code_SDir(InfBase)
  InfSubs = Left(InfSubs,Len(InfSubs)-1)
 End If
End If

'Response.Write code_SDir(InfBase)
If InfSave=""  Then 
  InfSave = Get_yyyymmdd("")
  InfSave = Left(InfSave,4)&"-"&Right(InfSave,4)
End If


'/////////////////////////////////////////////////////////////////////////
'/////////////////////////////////////////////////////////////////////////

sPBase = InfBase&"\"
sFExts = InfExts ':Response.Write "d."&sFExts
sFills = InfFill
aFills = split(sFills,"|")
Dim tCont :tCont=""


fmtHead1 = vbcrlf&vbcrlf
fmtHead1 = fmtHead1&vbcrlf&"/*******************************************************************************"
fmtHead1 = fmtHead1&vbcrlf&" File : (fName) "
fmtHead1 = fmtHead1&vbcrlf&"*******************************************************************************/"
fmtHead1 = fmtHead1&vbcrlf&vbcrlf


If p<>"" Then
  
  pRoot = Config_Path&"upfile/tools/code/"
  If act="CodeOut" Then
    Call fold_add(pRoot,InfSave) 
    If p="(ROOT)" Then
      cTemp = code_Root(sPBase) 
	  fName = pRoot&InfSave&"/0_root.txt"
    Else
	  cTemp = code_Subs(sPBase&p&"\")
	  fName = pRoot&InfSave&"/"&i&"_"&p&".txt"
    End If
    Call File_Add2(fName,cTemp,InfCSet)
  End If  
  Msg = ""&p&":完成！"&Now()
  
End If

Function code_Root(path)
    dim fs, folder, file, item, url, iPath, iName
    set fs = CreateObject("Scripting.FileSystemObject")
    set folder = fs.GetFolder(path)
    for each item in folder.Files
		iPath = Replace(item.Path,sPBase,"(Root)\")
		iName = Mid(item.Name,InStrRev(item.Name,"."),8)
	   If inStr(sFExts,iName)>0 Then
		  iHead = Replace(fmtHead1,"(fName)",iPath)
		  iCont = File_Read(item.Path,InfCSet) ':iCont = Replace(iCont,"<","〈")
		  tCont = tCont&iHead&iCont
	   End If
    next 
	code_Root = tCont
End Function

Function code_Subs(path)
    dim fs, folder, file, item, url, iPath, iName,i,f
    set fs = CreateObject("Scripting.FileSystemObject")
	'Response.Write path
    set folder = fs.GetFolder(path)
    for each item in folder.SubFolders
      f = true '/////////////////////////////////
	  if sFills<>"" Then
	    for i=0 to uBound(aFills)
	      if aFills(i)<>"" then
		  if inStr(item.Path,aFills(i)&"\")>0 then 
	        f = false
			exit for ''i = uBound(aFills)+2 ''
	      end if
		  end if
	    next
	  end if '//////////////////////////////////
	  if(f) then
	    code_Subs(item.Path)
	  end if
    next   
    for each item in folder.Files
		iPath = Replace(item.Path,sPBase,"\")
		iName = Mid(item.Name,InStrRev(item.Name,"."),8)
	   If inStr(sFExts,iName)>0 Then
		  iHead = Replace(fmtHead1,"(fName)",iPath)
		  iCont = File_Read(item.Path,InfCSet) ':iCont = Replace(iCont,"<","〈")
		  tCont = tCont&iHead&iCont
	   End If
    next 
	code_Subs = tCont
End Function

Function code_SDir(path)
    dim fs, folder, file, item, url, iPath, iName,i,f,r
	r = ""
    set fs = CreateObject("Scripting.FileSystemObject")
    set folder = fs.GetFolder(path)
    for each item in folder.SubFolders
      f = true '/////////////////////////////////
	  if sFills<>"" Then
	    for i=0 to uBound(aFills)
	      if aFills(i)<>"" then
		  if inStr(item.Path,aFills(i)&"\")>0 then 
	        f = false
			exit for ''i = uBound(aFills)+2 ''
	      end if
		  end if
	    next
	  end if '//////////////////////////////////
	  if(f) then
	    r = r&item.Name&"|"
	  end if
    next   
	code_SDir = r
End Function

Sub code_List(xPath2)
    dim fs, folder, file, item, url, iPath, iName, f, i
	Dim xPath : xPath=xPath2 
    set fs = CreateObject("Scripting.FileSystemObject")
	If instr(xPath,":")<=0 Then xPath=Server.MapPath(xPath) 
    set folder = fs.GetFolder(xPath)
    Response.Write("<li><b>" & folder.Name & "</b> - " & folder.Files.Count & " files, " & folder.SubFolders.Count & " dirs." & vbCrLf & "<ul>" & vbCrLf)
    for each item in folder.SubFolders
      f = true '/////////////////////////////////
	  if sFills<>"" Then
	    for i=0 to uBound(aFills)
	      if aFills(i)<>"" then
		  if inStr(item.Path,aFills(i)&"\")>0 then 
	        f = false
			exit for ''i = uBound(aFills)+2 ''
	      end if
		  end if
	    next
	  end if '//////////////////////////////////
	  if(f) then
	    code_List(item.Path)
	  end if
    next
    for each item in folder.Files
        iPath = item.Path
		iPath = Replace(iPath,sPBase,"\")
		Response.Write("<li>" & iPath & "</li>" & vbCrLf)
    next
    Response.Write("</ul>" & vbCrLf)
    Response.Write("</li>" & vbCrLf)
End Sub

%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>代码导出</title>
<script src="../../sadm/func1/Func_JS.js" type="text/javascript"></script>
<script src="../../inc/home/jsInfo.js" type="text/javascript"></script>
<style type="text/css">
body, td, th {
	font-size: 12px;
}
.btnSub {
	width:180px;
	text-align:left;
	margin:1px;
}
body {
	margin: 5px;
}
</style>
</head>
<body>
<table width="640" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="#CCCCCC">
  <form id="fmCodeOut" name="fmCodeOut" method="post" action="?">
    <tr>
      <td width="15%" align="center" bgcolor="#FFFFFF"><span id="RMsg">Path</span>: </td>
      <td bgcolor="#FFFFFF"><input name="InfBase" type="text" id="InfBase" value="<%=InfBase%>" size="48" maxlength="120" onchange="setSubs()" />
        <select name="LangSet" id="LangCSet" onchange="setTyps(this)">
          <%=Get_SOpt("(默认);asp;php;jsp;htm;vb;java;c","",InfCSet,"")%>
        </select></td>
    </tr>
    <tr>
      <td align="center" bgcolor="#FFFFFF">Ext:</td>
      <td bgcolor="#FFFFFF"><input name="InfExts" type="text" id="InfExts" value="<%=InfExts%>" size="48" maxlength="48" />
        <select name="InfCSet" id="InfCSet" onchange="setTyps(this)">
          <%=Get_SOpt("UTF-8;gb2312;big5;iso-8859-1","",InfCSet,"")%>
        </select>
      <input name="i" type="hidden" id="i" /></td>
    </tr>
    <tr>
      <td align="center" bgcolor="#FFFFFF">Fill:</td>
      <td bgcolor="#FFFFFF"><input name="InfFill" type="text" id="InfFill" value="<%=InfFill%>" size="48" maxlength="48" />
      <input name="p" type="hidden" id="p" /></td>
    </tr>
    <tr>
      <td align="center" bgcolor="#FFFFFF">Subs:</td>
      <td nowrap="nowrap" bgcolor="#FFFFFF"><input name="InfSubs" type="text" id="InfSubs" value="<%=InfSubs%>" size="48" maxlength="120" />
        <input name="act" type="hidden" id="act" value="GetSubs" /></td>
    </tr>
    <tr>
      <td align="center" bgcolor="#FFFFFF">Action: </td>
      <td bgcolor="#FFFFFF"><div style="float:right"> <a href="?<%=Now()%>">Reload</a> </div>
        /upfile/tools/code/
          <input name="InfSave" type="text" id="InfSave" value="<%=InfSave%>" size="12" xxcus="iMsg(12)" />
        <input type="submit" name="btmDown" id="btmDown" value="保存设置&amp;显示子目录" /></td>
    </tr>
    <tr>
      <td align="center" bgcolor="#FFFFFF">Dirs:</td>
      <td bgcolor="#FFFFFF">
	    <div> 
        <input class='btnSub' name="btn0" type="button" id="btn0" value="导出:0-(ROOT)" onclick="sndSub(0,'(ROOT)')" />
        <input class='btnSub' name="btm0" type="button" id="btm0" value="列表:0-(ROOT)" onclick="sndList(0,'(ROOT)')" disabled="disabled" />
	    </div>
		<%
aPath = Split(InfSubs,"|")
For i=0 To uBound(aPath)
  j = i+1
  iFlg = aPath(i) 'Replace(sPBase&,"\","\\")
  Response.Write vbcrlf&"<div>"
  Response.Write vbcrlf&"<input class='btnSub' name='btn"&j&"' type='button' id='btn"&j&"' value='导出:"&j&"-"&aPath(i)&"' onclick=""sndSub("&j&",'"&iFlg&"')"" />"
  Response.Write vbcrlf&"<input class='btnSub' name='btm"&j&"' type='button' id='btm"&j&"' value='列表:"&j&"-"&aPath(i)&"' onclick=""sndList("&j&",'"&iFlg&"')"" />"
  Response.Write vbcrlf&"</div>"
Next 
	  %></td>
    </tr>
    <tr>
      <td align="center" bgcolor="#FFFFFF">Msg</td>
      <td bgcolor="#FFFFFF"><div id="msgBox" style="padding:1px; background-color:#FFFFCC; clear:both;">
        注意：<%=Msg%>
        &nbsp; <br />
      1. 本程序主要用于本地程序代码导出,使用本程序<span style="color:#F0F">可能暴露源代码等敏感信息,请谨慎使用</span>；<br />
      2. 注意操作的目录,要有NTFS权限；<br />
      3. 
      使用后请清理(删除)<a href="../../smod/file/file_view.asp?yPath=tools/code/" target="_blank">导出的文件</a>,因使用不当而造成损失,本程序开发人员不负任何责任；</div></td>
    </tr>
  </form>
</table>
<%
If p<>"" And act="CodeList" Then
    If p="(ROOT)" Then
      'cTemp = code_Root(sPBase) 
	  'fName = pRoot&InfSave&"/0_root.txt"
    Else
	  Call code_List(sPBase&p&"\")
    End If
End If
%>

</body>
</html>
<script type="text/javascript">
var fm = document.fmCodeOut;

function setSubs(e){
  fm.InfSubs.value='';	
  for(i=0;i<=<%=j%>;i++){
	//getElmID("").disab
	var idElm1 = getElmID("btn"+i);
	var idElm2 = getElmID("btm"+i);
	idElm1.style.display='none';
	idElm1.style.visibility='hidden';
	idElm2.style.display='none';
	idElm2.style.visibility='hidden';
  }
}


function setTyps(e){
  var v = e.value; //alert(v); asp;php;jsp;htm;vb;java;c
  fm.InfExts.value = ".asp|.htm|.js|.vb|.cs|.cpp|.php|.java|.js";
  if(v=="asp") { fm.InfExts.value = ".asp|.htm|.js"; } 
  if(v=="php") { fm.InfExts.value = ".php|.htm|.js"; }
  if(v=="jsp") { fm.InfExts.value = ".jsp|.htm|.js|.java"; }
  if(v=="htm") { fm.InfExts.value = ".htm|.stm|"; }
  if(v=="vb")  { fm.InfExts.value =  ".vb|.js|.htm"; }
  if(v=="java"){ fm.InfExts.value = ".java|.js|.htm"; }
  if(v=="c")   { fm.InfExts.value = ".c|.cpp|.h|.htm"; }
}
function sndSub(i,p){
  fm.i.value = i;
  fm.p.value = p;
  fm.act.value = "CodeOut";
  fm.submit();
}
function sndList(i,p){
  fm.i.value = i;
  fm.p.value = p;
  fm.act.value = "CodeList";
  fm.submit();
}
</script>

