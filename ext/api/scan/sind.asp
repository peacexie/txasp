<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="slib.asp"-->
<%
server.ScriptTimeout =90000

act=request("act")
if act="login" then
	if request.Form("pwd") = PASSWORD then session("ChkCode")="ok"
elseif act="out" then
    session("ChkCode")="Not Login"
end if
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title>Asp木马扫描器</title>
<script language="JavaScript" type="text/JavaScript">

function ConfirmDel()
{
   if(confirm("确认删除？并且不能恢复！"))
     return true;
   else
     return false;
	 
}
</script>
</head>

<body>
<div align="center"><h2>Asp木马扫描器</h2></div>
<hr>
<%
If Session("ChkCode") <> "ok" then
	call LoginForm()
else
	dim pathStr
	if request("path")<>"" then
		pathStr=request("path")
	else
		pathStr=server.MapPath("/")
	end if
	response.Write("<a href=""javascript:history.back();"">←返回</a> | <a href=""?act=out"">登出→</a><br>"&Chr(10))
	if act="scan" then		
		dim Suspect,ScanFileNum,ScanFolderNum,BeginTime,EndTime,TmpPath,Report		
		'ScanFileType = "asp,cer,asa,cdx"
		Suspect = 0
		ScanFileNum = 0
		ScanFolderNum =0		
		BeginTime = timer
		response.Write("<textarea name=""textarea"" style=""width:100%"" rows=""5"">"&Chr(10))
		response.Write("扫描日志："&vbcrlf)
		if(request.QueryString("file")<>"") then
			Call ScanFile(request.QueryString("file"),"")
		else
			Call ScanFolder(pathStr)
		end if
		response.Write("</textarea>")
		Call ShowResult()
		EndTime = timer		
		response.write "<br><font size=""2"">执行时间："&cstr(int(((EndTime-BeginTime)*10000 )+0.5)/10)&"毫秒</font>"	
	elseif act="del" then
		Call DelFile(request.QueryString("file"))
		response.Write("<br><a href="""&request.ServerVariables("HTTP_REFERER")&""">返回</a>")
	elseif act="down" then
		Call Download(request.QueryString("file"))
	else
		call FileList(pathStr)
		call ScanForm()
	end if
end if

%>
<hr>
</body>
</html>
<%
Sub LoginForm
%>
<form name="form1" method="post" action="../../../tools/scan/?act=login">
  <div align="center">Password: 
    <input name="pwd" type="password" size="15"> 
    <input type="submit" name="Submit" value="提交">
  </div>
</form>
<%
end Sub
Sub ScanForm
%>
<form action="../../../tools/scan/?act=scan" method="post">
	<input type="submit" value=" 全站扫描 " style="background:#fff;border:1px solid #999;padding:2px 2px 0px 2px;margin:4px;border-width:1px 3px 1px 3px" />
</form>
<%
end sub
'遍历处理path及其子目录所有文件
Sub FileList(Path)
	Set FSO = CreateObject("Scripting.FileSystemObject")
	if not fso.FolderExists(path) then exit sub
	Set folders = FSO.GetFolder(Path)'目录下所有对象
	Set files = folders.files
	Set subfolders = folders.SubFolders
	'列表文件夹
	For Each fl in subfolders
		response.Write("<a href=""?path="&Path&"\"&fl.name&"""><img src=""../../img/tool/folder2.gif"" border=""0"">"&fl.name&"</a>"&Chr(10))
		response.Write("<a href=""?act=scan&path="&Path&"\"&fl.name&""">扫描</a><br>"&Chr(10))
    Next
	'列表文件
	For Each file_f in files
		response.Write("<img src=""../../img/tool/inf_from.gif"">"&file_f.name&""&Chr(10))
		response.Write("<a href=""?act=scan&file="&Path&"\"&file_f.name&""">扫描</a><br>"&Chr(10))
	Next
	set folders=nothing
	set files=nothing
	set subfolders=nothing
	Set FSO = Nothing
End Sub
Sub ShowResult
%>
<table width="100%" border="0" cellpadding="0" cellspacing="0" class="CContent">
  <tr>
    <td class="CPanel" style="padding:5px;line-height:170%;clear:both;font-size:12px">        
扫描完毕！一共检查文件夹<font color="#FF0000"><%=ScanFolderNum%></font>个，文件<font color="#FF0000"><%=ScanFileNum%></font>个，发现可疑点<font color="#FF0000"><%=Suspect%></font>个	
</td></tr></table>
<table width="100%" border="0" cellpadding="0" cellspacing="1" style="padding:5px; background-color:#666666;line-height:18px;clear:both;font-size:12px">
	<tr>
		<td width="30%" bgcolor="#FFFFFF">文件名称</td>
		<td width="20%" bgcolor="#FFFFFF">特征码</td>
		<td width="30%" bgcolor="#FFFFFF">描述</td>
		<td width="20%" bgcolor="#FFFFFF">创建/修改时间</td>
	</tr>
	<p>
	<%=Report%>
	<br/>
	</p>
</table>
<%
end Sub
'遍历处理path及其子目录所有文件
Sub ScanFolder(Path)
	dim folders,files,subfolders
	ScanFolderNum = ScanFolderNum + 1
	Set FSO = CreateObject("Scripting.FileSystemObject")
	if not fso.FolderExists(path) then exit sub
	Set folders = FSO.GetFolder(Path)
	Set files = folders.files
	For Each myfile in files
		If CheckExt(FSO.GetExtensionName(path&"\"&myfile.name)) Then
			Call ScanFile(Path&"\"&myfile.name, "")			
		End If
	Next
	Set subfolders = folders.SubFolders
	For Each f1 in subfolders
		ScanFolder path&"\"&f1.name		
    Next
	set folders=nothing
	set files=nothing
	set subfolders=nothing
	Set FSO = Nothing
End Sub

'检测文件
Sub ScanFile(FilePath, InFile)
	dim FSOs,ofile,filetxt,fileUri,vi
	ScanFileNum = ScanFileNum + 1
	response.Write("扫描文件:"&FilePath&vbcrlf)
	response.Flush()
	If InFile <> "" Then
		Infiles = "该文件被<a href=""http://"&Request.Servervariables("server_name")&"\"&InFile&""" target=_blank>"& InFile & "</a>文件包含执行"
	End If
	Set FSOs = CreateObject("Scripting.FileSystemObject")
	on error resume next
	set ofile = fsos.OpenTextFile(FilePath)
	filetxt = Lcase(ofile.readall())
	If err Then Exit Sub end if
	if len(filetxt)>0 then		
		'特征码检查
		fileUri = "<a href=""http://"&Request.Servervariables("server_name")&":"&Request.ServerVariables("SERVER_PORT")&"\"&replace(FilePath,server.MapPath("\")&"\","",1,1,1)&""" target=_blank>"&replace(FilePath,server.MapPath("\")&"\","",1,1,1)&"</a><br>"
		fileUri=fileUri&"操作：&nbsp;<a href=""?act=del&file="&FilePath&""" onClick=""return ConfirmDel()"">删除</a>"
		fileUri=fileUri&"&nbsp;<a href=""?act=down&file="&FilePath&""">下载</a>"
		for vi=0 to ubound(virus,2)
			If instr(filetxt, Lcase(virus(0,vi))) then
				Report = Report&"<tr bgcolor=""#FFFFFF""><td>"&fileUri&"</td><td>"&virus(0,vi)&"</td><td>"&virus(1,vi)&infiles&"</td><td>创建:"&GetDateCreate(filepath)&"<br>修改:"&GetDateModify(filepath)&"</td></tr>"
				Suspect = Suspect + 1
			End if
		next			
		for vi=0 to ubound(virus_Regx,2)
			Set regEx = New RegExp
			regEx.IgnoreCase = True
			regEx.Global = True
			regEx.Pattern = virus_Regx(0,vi)
			If regEx.Test(filetxt) Then
				Report = Report&"<tr bgcolor=""#FFFFFF""><td>"&fileUri&"</td><td>"&virus_Regx(0,vi)&"</td><td>"&virus_Regx(1,vi)&infiles&"</td><td>创建:"&GetDateCreate(filepath)&"<br>修改:"&GetDateModify(filepath)&"</td></tr>"
				Suspect = Suspect + 1
			End If
		next		
			
		'Check include file
		Set regEx = New RegExp
		regEx.IgnoreCase = True
		regEx.Global = True
		regEx.Pattern = "<!--\s*#include\s*file\s*=\s*"".*"""
		Set Matches = regEx.Execute(filetxt)
		For Each Match in Matches			
			tFile = Replace(Mid(Match.Value, Instr(Match.Value, """") + 1, Len(Match.Value) - Instr(Match.Value, """") - 1),"/","\")
			If Not CheckExt(FSOs.GetExtensionName(tFile)) Then
				Call ScanFile( Mid(FilePath,1,InStrRev(FilePath,"\"))&tFile, replace(FilePath,server.MapPath("\")&"\","",1,1,1) )
				SumFiles = SumFiles + 1
			End If
		Next
		Set Matches = Nothing
		Set regEx = Nothing
		
		'Check include virtual
		Set regEx = New RegExp
		regEx.IgnoreCase = True
		regEx.Global = True
		regEx.Pattern = "<!--\s*#include\s*virtual\s*=\s*"".*"""
		Set Matches = regEx.Execute(filetxt)
		For Each Match in Matches
			tFile = Replace(Mid(Match.Value, Instr(Match.Value, """") + 1, Len(Match.Value) - Instr(Match.Value, """") - 1),"/","\")
			If Not CheckExt(FSOs.GetExtensionName(tFile)) Then
				Call ScanFile( Server.MapPath("\")&"\"&tFile, replace(FilePath,server.MapPath("\")&"\","",1,1,1) )				
			End If
		Next
		Set Matches = Nothing
		Set regEx = Nothing
		
		'Check Server&.Execute|Transfer
		Set regEx = New RegExp
		regEx.IgnoreCase = True
		regEx.Global = True
		regEx.Pattern = "Server.(Exec"&"ute|Transfer)([ \t]*|\()"".*"""
		Set Matches = regEx.Execute(filetxt)
		For Each Match in Matches
			tFile = Replace(Mid(Match.Value, Instr(Match.Value, """") + 1, Len(Match.Value) - Instr(Match.Value, """") - 1),"/","\")
			If Not CheckExt(FSOs.GetExtensionName(tFile)) Then
				Call ScanFile( Mid(FilePath,1,InStrRev(FilePath,"\"))&tFile, replace(FilePath,server.MapPath("\")&"\","",1,1,1) )				
			End If
		Next
		Set Matches = Nothing
		Set regEx = Nothing	
		
		
	end if
	set ofile = nothing
	set fsos = nothing
End Sub

'检查文件后缀，如果与预定的匹配即返回TRUE
Function CheckExt(FileExt)
	If ScanFileType = "*" Then CheckExt = True
	Ext = Split(ScanFileType,",")
	For i = 0 To Ubound(Ext)
		If Lcase(FileExt) = Ext(i) Then 
			CheckExt = True
			Exit Function
		End If
	Next
End Function
'删除文件
Sub DelFile(FilePath)
	Set fso = Server.CreateObject("Scripting.FileSystemObject")
 	if fso.FileExists(FilePath) then
 		fso.DeleteFile(FilePath)
 		Response.Write("<h2>成功删除文件:</h2>" &FilePath)
	else
		response.Write("<h2>删除失败!文件:"&FilePath&"没有找到!</2>")
 	end if
 	set fso=nothing
end Sub
'下载文件
sub Download(FilePath)
	dim oStream
	Set FSO = Server.CreateObject("Scripting.FileSystemObject")
	if FSO.FileExists(FilePath) then
		set oStream=Server.CreateObject("ADODB.Stream")
		oStream.Type=1
		oStream.Open
		on error resume next
		oStream.LoadFromFile(FilePath)
		if Err.Number=0 then
			Response.AddHeader "Content-Disposition", "attachment; filename=" & FSO.GetFileName(FilePath)
			Response.AddHeader "Content-Length", oStream.Size
			Response.ContentType="bad/type" 'yeu cau ie hien hop thoai save-as
			Response.BinaryWrite oStream.Read
		end if
		oStream.Close
		set oStream=nothing
	end if
	set FSO=nothing
end sub
Function GetDateModify(filepath)
	dim s,days
	Set fso = CreateObject("Scripting.FileSystemObject")
    Set f = fso.GetFile(filepath) 
	s = f.DateLastModified 
	set f = nothing
	set fso = nothing
	days=DateDiff("d",Cdate(s),now())
	if(days>-7 and days<7) then
		s="<font color=""red"">"&s&"</font>"
	end if
	GetDateModify = s
End Function

Function GetDateCreate(filepath)
	dim s,days
	Set fso = CreateObject("Scripting.FileSystemObject")
    Set f = fso.GetFile(filepath) 
	s = f.DateCreated 
	set f = nothing
	set fso = nothing
	days=DateDiff("d",Cdate(s),now())
	if(days>-7 and days<7) then
		s="<font color=""red"">"&s&"</font>"
	end if
	GetDateCreate = s
End Function

%>