<!--#include file="config.asp"-->
<!--#include file="../../sadm/func1/up_class.asp"-->
<html>
<head>
<title>导入查询数据</title>
<meta http-equiv="Pragma" content="no-cache">
<link href="../../inc/adm_inc/adm_style.css" rel="stylesheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<script src="../../sadm/func1/Func_JS.js" type="text/javascript"></script>
<script src="../../sadm/func1/WinFunc.js" type="text/javascript"></script>
</head>
<body>

<% 

set upload=new upload_Class

'If Left(StuTabCode,3)<>"Stu" Then Response.End()
'Response.End()
   
 fMsg = "<blockquote><br>"
 fp = "CrdFile" '"Img1"&fii
 set up_file = upload.File(fp)
 if up_file.FileSize>0 then
   save_name = Get_yyyymmdd("")&"-"&Get_hhmmss()&"-"&Rnd_ID("KEY",4)&Mid(up_file.FileName,InStrRev(up_file.FileName,"."),8) 
   save_Title= up_file.FileName
   upFlag = "OK"
   If up_file.FileSize>5*1024*1024 Then 
     fMsg=fMsg&"文件大于5M；"
	 upFlag = "NG"
   End If
   If inStr(uCase(up_file.FileName),".XLS")<=0 Then 
     fMsg=fMsg&"文件格式错误；"
	 upFlag = "NG"
   End If
   If upFlag = "OK" then
     up_file.SaveAs Server.Mappath(Config_Path&"upfile/#dbf#/"&save_name)
	 fMsg=fMsg&"文件上传成功；"
   End If
 end if

set up_file = Nothing
set upload = nothing

If upFlag = "OK" then
 'Dim Conn,Driver,DBPath,Rs
 
 Set cn = Server.CreateObject("ADODB.Connection")
 conExcel = "Driver={Microsoft Excel Driver (*.xls)};DBQ="&Server.MapPath(Config_Path&"upfile/#dbf#/"&save_name)
 cn.Open conExcel
 Sql="Select * From [Sheet1$] "
 Set rs=cn.Execute(Sql)

 IF rs.Eof And Rs.Bof Then
   fMsg=fMsg&"<br>没有找到您需要的数据!!"
 Else
  j=0 : k=0 : sData = ""
  Do While Not Rs.EOF
   'Response.write "<br>"&rs("StuName")
   f10 = inStr(rs(0),"代表正常时间上班")
   f21 = inStr(rs(0)&rs(2),"时间")
   f22 = inStr(rs(0)&rs(2),"星期")
   f31 = inStr(rs(0)&rs(2)&rs(3)&rs(4),"上午")
   f32 = inStr(rs(0)&rs(2)&rs(3)&rs(4),"下午")
    If f10>0 Or f21+f22>2 Or f31+f32>2 Or rs(0)&""="" Then
	  '///// 
	  j = j+1
	Else
	  k = k+1
	  sData = sData&vbcrlf& "<tr>"
	  For i=0	To 14 
		sData = sData&vbcrlf& "  <td>"&rs(i)&"</td>"
	  Next
	  sData = sData&vbcrlf& "</tr>"
    End If 
   'Response.write "<br>"&rs("StuName")
  rs.MoveNext
  Loop
 End IF
 rs.Close
 Set rs=nothing
 cn.Close
 Set cn=Nothing

Else
   ' 文件上传失败
End if


sData = Replace(sData,"<td></td>","<td>&nbsp;</td>")
Call File_Add2(Config_Path&"upfile/sys/cache/order.htm",sData,"utf-8")
Response.write "<table border=1 cellpadding=0 cellspacing=1>"&sData&vbcrlf&"</table>"


fMsg=fMsg&"<br>成功导入数据：["&k&"]笔；<br>导入失败数据：["&j&"]笔"
fMsg=fMsg&"<br>导入完毕！<a href='admimport.asp'>返回</a></blockquote>"
Response.write fMsg

%>

    