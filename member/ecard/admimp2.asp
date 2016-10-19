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
 Sql="Select * From [CrdDemo$] "
 Set rs=cn.Execute(Sql)

 IF rs.Eof And Rs.Bof Then
   fMsg=fMsg&"<br>没有找到您需要的数据!!"
 Else
  j=0 : k=0
  Do While Not Rs.EOF
CrdNO = RequestSafe(rs("CrdNO"),"C",48) :CrdNO=Replace(CrdNO," ","")
CrdNCode = RequestSafe(rs("CrdNCode"),"C",48) :CrdNCode=Replace(CrdNCode," ","")
CrdName = RequestSafe(rs("CrdName"),"C",48) :CrdName=Replace(CrdName," ","")
CrdTime = RequestSafe(rs("CrdTime"),"D",48)
CrdSpeci = RequestSafe(rs("CrdSpeci"),"C",48)
CrdQty = RequestSafe(rs("CrdQty"),"N",1)
CrdType = RequestSafe(rs("CrdType"),"C",48)
CrdRem = RequestSafe(rs("CrdRem"),"C",255)
CrdID = Get_AutoID(24) 
sql = " INSERT INTO [MemCard] (" 
sql = sql& "  CrdID" 
sql = sql& ", CrdNO" 
sql = sql& ", CrdNCode" 
sql = sql& ", CrdName" 
sql = sql& ", CrdType" 
sql = sql& ", CrdTime" 
sql = sql& ", CrdSpeci" 
sql = sql& ", CrdQty"
sql = sql& ", CrdRem" 
sql = sql& ", CrdC48" 
sql = sql& ", CrdDate" 
sql = sql& ", CrdInt"
sql = sql& ")VALUES(" 
sql = sql& "  '" & CrdID &"'" 
sql = sql& ", '" & CrdNO &"'" 
sql = sql& ", '" & CrdNCode &"'" 
sql = sql& ", '" & CrdName &"'" 
sql = sql& ", '" & CrdType &"'" 
sql = sql& ", '" & CrdTime &"'" 
sql = sql& ", '" & CrdSpeci &"'" 
sql = sql& ", " & CrdQty &"" 
sql = sql& ", '" & CrdRem &"'" 
sql = sql& ", '" & RequestSafe(CrdC48,"C",48) &"'" 
sql = sql& ", '" & RequestSafe(CrdDate,"D",255) &"'" 
sql = sql& ", " & RequestSafe(CrdInt,"N",255) &"" 
sql = sql& ")"
If rs_Exist(conn,"SELECT CrdNO FROM MemCard WHERE CrdNO='"&CrdNO&"' AND CrdType='"&CrdType&"' AND CrdName='"&CrdName&"' ")="EOF" Then
 Call rs_Dosql(conn,sql)
 'fMsg=fMsg&"<br> --- ：       ["&CrdNO&"]-["&CrdNamee&"]"
 k=k+1
Else
 'Call rs_Dosql(conn,sql)
 fMsg=fMsg&"<br> --- 导入失败：["&CrdNO&"]-["&CrdName&"]-["&CrdType&"]"
 j=j+1
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

fMsg=fMsg&"<br>成功导入数据：["&k&"]笔；<br>导入失败数据：["&j&"]笔"
fMsg=fMsg&"<br>导入完毕！<a href='admimport.asp'>返回</a></blockquote>"
Response.write fMsg

%>

    