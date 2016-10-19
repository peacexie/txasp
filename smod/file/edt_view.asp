<!--#include file="../../upfile/sys/pcfg/editor.asp"-->
<!--#include file="edt_config.asp"-->
<!--#include file="../../ext/jqplug/json.asp"-->
<%
sPerm = Chk_PEditor("") 
if sPerm<>"(Pass)" then 
  Call json_Alert(1,"("&sPerm&")NoPerm!","msg")
  Response.End()
end if 

Dim aspUrl, rootPath, rootUrl, fileTypes
Dim currentPath, currentUrl, currentDirPath, moveupDirPath
Dim path, order, dirName, fso, folder, dir, file, result
Dim fileExt, dirCount, fileCount, orderIndex, i, j
Dim dirList(), fileList(), isDir, hasFile, filesize, isPhoto, filetype, filename, datetime

aspUrl = Request.ServerVariables("SCRIPT_NAME")
aspUrl = left(aspUrl, InStrRev(aspUrl, "/"))

'根目录路径，可以指定绝对路径，比如 /var/www/attached/
rootPath = Config_Path&"upfile/" ':Response.Write rootPath
'根目录URL，可以指定绝对路径，比如 http://www.yoursite.com/attached/
rootUrl = Config_Path&"upfile/"
'图片扩展名
fileTypes = "gif,jpg,jpeg,png" ',bmp


currentPath = ""
currentUrl = ""
currentDirPath = ""
moveupDirPath = ""



Set fso = Server.CreateObject("Scripting.FileSystemObject")

'目录名
dirName = Request.QueryString("dir")
If Not isEmpty(dirName) Then
	If instr(lcase("image,flash,media,file"), dirName) < 1 Then
		Response.Write "Invalid Directory name."
		Response.End
	End If
	rootPath = rootPath & dirName & "/"
	rootUrl = rootUrl & dirName & "/"
	If Not fso.FolderExists(Server.mappath(rootPath)) Then
		fso.CreateFolder(Server.mappath(rootPath))
	End If
End If

'根据path参数，设置各路径和URL
path = Request.QueryString("path")
If path = "" Then
	currentPath = Server.MapPath(rootPath) & "\" 
	currentUrl = rootUrl
	currentDirPath = ""
	moveupDirPath = ""
Else
	currentPath = Server.MapPath(rootPath & path) & "\"
	currentUrl = rootUrl + path
	currentDirPath = path
	moveupDirPath = RegexReplace(currentDirPath, "(.*?)[^\/]+\/$", "$1")
End If
'Response.Write "<br>"&currentPath
'Response.Write "<br>"&currentUrl
'Response.Write "<br>"&currentDirPath
'Response.Write "<br>"&moveupDirPath

Set folder = fso.GetFolder(currentPath)

'排序形式，name or size or type
order = lcase(Request.QueryString("order"))
Select Case order
	Case "type" orderIndex = 4
	Case "size" orderIndex = 2
	Case Else  orderIndex = 5
End Select

'不允许使用..移动到上一级目录
If RegexIsMatch(path, "\.\.") Then
	Response.Write "Access is not allowed."
	Response.End
End If
'最后一个字符不是/
If path <> "" And Not RegexIsMatch(path, "\/$") Then
	Response.Write "Parameter is not allowed."
	Response.End
End If
'目录不存在或不是目录
If Not DirectoryExists(currentPath) Then
	Response.Write "Directory does not exist."
	Response.End
End If

Set result = jsObject()
'相对于根目录的上一级目录
result("moveup_dir_path") = moveupDirPath
'相对于根目录的当前目录
result("current_dir_path") = currentDirPath
'当前目录的URL
result("current_url") = currentUrl

'文件数
dirCount = folder.SubFolders.count
fileCount = folder.Files.count
result("total_count") = dirCount + fileCount

ReDim dirList(dirCount)
i = 0
For Each dir in folder.SubFolders
	isDir = True
	hasFile = (dir.Files.count > 0)
	filesize = 0
	isPhoto = False
	filetype = ""
	filename = dir.name
	datetime = FormatDate(dir.DateLastModified)
	'///// if(!(strstr('(index.php,#dbf#)',$filename))){
	dirList(i) = Array(isDir, hasFile, filesize, isPhoto, filetype, filename, datetime)
	i = i + 1
	'/////
Next
ReDim fileList(fileCount)
i = 0
For Each file in folder.Files
	fileExt = mid(file.name, InStrRev(file.name, ".") + 1)
	isDir = False
	hasFile = False
	filesize = file.size
	isPhoto = (instr(lcase(fileTypes), fileExt) > 0)
	filetype = fileExt
	filename = file.name
	datetime = FormatDate(file.DateLastModified)
	'/////
	fileList(i) = Array(isDir, hasFile, filesize, isPhoto, filetype, filename, datetime)
	i = i + 1
	'/////
Next

'排序
Dim minidx, temp
For i = 0 To dirCount - 2
	minidx = i
	For j = i + 1 To dirCount - 1
		If (dirList(minidx)(5) > dirList(j)(5)) Then
			minidx = j
		End If
	Next
	If minidx <> i Then
		temp = dirList(minidx)
		dirList(minidx) = dirList(i)
		dirList(i) = temp
	End If
Next
For i = 0 To fileCount - 2
	minidx = i
	For j = i + 1 To fileCount - 1
		If (fileList(minidx)(orderIndex) > fileList(j)(orderIndex)) Then
			minidx = j
		End If
	Next
	If minidx <> i Then
		temp = fileList(minidx)
		fileList(minidx) = fileList(i)
		fileList(i) = temp
	End If
Next

Set result("file_list") = jsArray()
For i = 0 To dirCount - 1
	Set result("file_list")(Null) = jsObject()
	result("file_list")(Null)("is_dir") = dirList(i)(0)
	result("file_list")(Null)("has_file") = dirList(i)(1)
	result("file_list")(Null)("filesize") = dirList(i)(2)
	result("file_list")(Null)("is_photo") = dirList(i)(3)
	result("file_list")(Null)("filetype") = dirList(i)(4)
	result("file_list")(Null)("filename") = dirList(i)(5)
	result("file_list")(Null)("datetime") = dirList(i)(6)
Next
For i = 0 To fileCount - 1
	Set result("file_list")(Null) = jsObject()
	result("file_list")(Null)("is_dir") = fileList(i)(0)
	result("file_list")(Null)("has_file") = fileList(i)(1)
	result("file_list")(Null)("filesize") = fileList(i)(2)
	result("file_list")(Null)("is_photo") = fileList(i)(3)
	result("file_list")(Null)("filetype") = fileList(i)(4)
	result("file_list")(Null)("filename") = fileList(i)(5)
	result("file_list")(Null)("datetime") = fileList(i)(6)
Next


'输出JSON字符串 
testStrJSON1 = "{""moveup_dir_path"":"""",""current_dir_path"":"""",""current_url"":""\/php\/upfile\/"",""total_count"":12,""file_list"":[{""is_dir"":true,""has_file"":false,""filesize"":0,""is_photo"":false,""filetype"":"""",""filename"":""defdt"",""datetime"":""2011-07-11 01:45:54""},{""is_dir"":true,""has_file"":false,""filesize"":0,""is_photo"":false,""filetype"":"""",""filename"":""defup"",""datetime"":""2011-08-27 03:39:55""},{""is_dir"":true,""has_file"":false,""filesize"":0,""is_photo"":false,""filetype"":"""",""filename"":""dtbbs"",""datetime"":""2011-07-11 01:45:54""},{""is_dir"":true,""has_file"":false,""filesize"":0,""is_photo"":false,""filetype"":"""",""filename"":""dtbus"",""datetime"":""2011-07-11 01:45:54""},{""is_dir"":true,""has_file"":true,""filesize"":0,""is_photo"":false,""filetype"":"""",""filename"":""dtinf"",""datetime"":""2011-08-29 02:36:25""},{""is_dir"":true,""has_file"":false,""filesize"":0,""is_photo"":false,""filetype"":"""",""filename"":""dtoab"",""datetime"":""2011-07-11 01:45:54""},{""is_dir"":true,""has_file"":false,""filesize"":0,""is_photo"":false,""filetype"":"""",""filename"":""dtoad"",""datetime"":""2011-07-11 01:45:54""},{""is_dir"":true,""has_file"":false,""filesize"":0,""is_photo"":false,""filetype"":"""",""filename"":""dtoau"",""datetime"":""2011-03-28 01:20:34""},{""is_dir"":true,""has_file"":true,""filesize"":0,""is_photo"":false,""filetype"":"""",""filename"":""dtpic"",""datetime"":""2011-08-29 05:52:46""},{""is_dir"":true,""has_file"":true,""filesize"":0,""dir_path"":"""",""is_photo"":false,""filetype"":"""",""filename"":""myfile"",""datetime"":""2011-05-12 02:44:39""},{""is_dir"":true,""has_file"":true,""filesize"":0,""is_photo"":false,""filetype"":"""",""filename"":""myftp"",""datetime"":""2011-07-15 06:59:01""},{""is_dir"":false,""has_file"":false,""filesize"":2864,""dir_path"":"""",""is_photo"":false,""filetype"":""txt"",""filename"":""readme.txt"",""datetime"":""2010-06-23 01:36:31""}]}"
testStrJSON2 = "{""moveup_dir_path"":"""",""current_dir_path"":"""",""current_url"":""\/asp\/upfile\/"",""total_count"":17,""file_list"":[{""is_dir"":true,""has_file"":true,""filesize"":0,""is_photo"":false,""filetype"":"""",""filename"":""#dbf#"",""datetime"":""2011-09-01 08:06:53""},{""is_dir"":true,""has_file"":false,""filesize"":0,""is_photo"":false,""filetype"":"""",""filename"":""defdt"",""datetime"":""2011-08-05 14:08:51""},{""is_dir"":true,""has_file"":false,""filesize"":0,""is_photo"":false,""filetype"":"""",""filename"":""defup"",""datetime"":""2011-08-05 14:08:51""},{""is_dir"":true,""has_file"":false,""filesize"":0,""is_photo"":false,""filetype"":"""",""filename"":""dtbbs"",""datetime"":""2011-08-05 14:08:51""},{""is_dir"":true,""has_file"":false,""filesize"":0,""is_photo"":false,""filetype"":"""",""filename"":""dtbus"",""datetime"":""2011-08-05 14:08:51""},{""is_dir"":true,""has_file"":false,""filesize"":0,""is_photo"":false,""filetype"":"""",""filename"":""dtdoc"",""datetime"":""2011-08-05 14:08:51""},{""is_dir"":true,""has_file"":false,""filesize"":0,""is_photo"":false,""filetype"":"""",""filename"":""dtinf"",""datetime"":""2011-08-31 14:09:17""},{""is_dir"":true,""has_file"":false,""filesize"":0,""is_photo"":false,""filetype"":"""",""filename"":""dtpic"",""datetime"":""2011-08-31 14:05:46""},{""is_dir"":true,""has_file"":false,""filesize"":0,""is_photo"":false,""filetype"":"""",""filename"":""file"",""datetime"":""2011-08-31 14:32:10""},{""is_dir"":true,""has_file"":false,""filesize"":0,""is_photo"":false,""filetype"":"""",""filename"":""flash"",""datetime"":""2011-08-31 14:32:18""},{""is_dir"":true,""has_file"":false,""filesize"":0,""is_photo"":false,""filetype"":"""",""filename"":""image"",""datetime"":""2011-08-31 14:23:35""},{""is_dir"":true,""has_file"":false,""filesize"":0,""is_photo"":false,""filetype"":"""",""filename"":""myfile"",""datetime"":""2011-02-16 10:20:18""},{""is_dir"":true,""has_file"":false,""filesize"":0,""is_photo"":false,""filetype"":"""",""filename"":""myftp"",""datetime"":""2010-11-15 13:37:05""},{""is_dir"":true,""has_file"":false,""filesize"":0,""is_photo"":false,""filetype"":"""",""filename"":""sys"",""datetime"":""2011-04-20 08:06:08""},{""is_dir"":true,""has_file"":false,""filesize"":0,""is_photo"":false,""filetype"":"""",""filename"":""tools"",""datetime"":""2010-11-11 16:44:01""},{""is_dir"":false,""has_file"":false,""filesize"":41,""is_photo"":false,""filetype"":""asp"",""filename"":""index.asp"",""datetime"":""2010-09-20 16:21:58""},{""is_dir"":false,""has_file"":false,""filesize"":2864,""is_photo"":false,""filetype"":""txt"",""filename"":""readme.txt"",""datetime"":""2010-06-23 09:36:31""}]}"

'// php: header('Content-type: application/json; charset=UTF-8');
'// Content-Type: text/plain, text/html, application/json
Response.AddHeader "Content-Type", "text/html; charset=UTF-8"
'Response.ContentType = "application/json; charset=UTF-8" 'application/vnd.ms-Excel;application/zip
'Response.Addheader "Content-Disposition","attachment;Filename=json.asp" 
'Response.Charset = "UTF-8" 

'Response.Write testStrJSON2
'Response.Write testStrJSON2
'Response.Write vbcrlf&vbcrlf&vbcrlf

result.Flush
Response.End()


'自定义函数
Function DirectoryExists(dirPath)
	Dim fso
	Set fso = Server.CreateObject("Scripting.FileSystemObject")
	DirectoryExists = fso.FolderExists(dirPath)
End Function

Function RegexIsMatch(subject, pattern)
	Dim reg
	Set reg = New RegExp
	reg.Global = True
	reg.MultiLine = True
	reg.Pattern = pattern
	RegexIsMatch = reg.Test(subject)
End Function

Function RegexReplace(subject, pattern, replacement)
	Dim reg
	Set reg = New RegExp
	reg.Global = True
	reg.MultiLine = True
	reg.Pattern = pattern
	RegexReplace = reg.Replace(subject, replacement)
End Function

Public Function FormatDate(datetime)
	Dim y, m, d, h, i, s
	y = CStr(Year(datetime))
	m = CStr(Month(datetime))
	If Len(m) = 1 Then m = "0" & m
	d = CStr(Day(datetime))
	If Len(d) = 1 Then d = "0" & d
	h = CStr(Hour(datetime))
	If Len(h) = 1 Then h = "0" & h
	i = CStr(Minute(datetime))
	If Len(i) = 1 Then i = "0" & i
	s = CStr(Second(datetime))
	If Len(s) = 1 Then s = "0" & s
	FormatDate = y & "-" & m & "-" & d & " " & h & ":" & i & ":" & s
End Function

%>
