<!--#include file="../../sadm/func1/func1.asp"-->
<!--#include file="../../sadm/func2/Func2.asp"-->
<!--#include file="../../sadm/func1/func_file.asp"-->
<!--#include file="../../sadm/func1/up_class.asp"-->
<!--#include file="config.asp"-->
<html>
<head>
<meta http-equiv="Pragma" content="no-cache">
<meta http-equiv="Expires" content="0">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title><%=Config_Name%>后台管理中心</title>
<link href="../../inc/adm_inc/adm_style.css" rel="stylesheet" type="text/css">
</head>
<body>

<%
 	
Call Chk_URL()	
DeMax = "1"'RequestSafe(rs_Val("","SELECT ParNum FROM AdmPara WHERE ParCode='n"&ModID&"'"),"N",1)

IP = Get_CIP()
set upload=new upload_Class

ReEnd = RequestSafe(upload.form("ReEnd"),3,24)
KeyID = Get_AutoID(24)
KeyCode = RequestSafe(upload.form("KeyCode"),3,48) 
InfSubj = RequestSafe(upload.form("InfSubj"),3,255) 
InfRem1 = upload.form("InfRem1")
InfRem1 = RequestSafe(InfRem1,3,9600) 
InfRem1 = Show_Html(InfRem1)
InfType = RequestSafe(upload.form("InfType"),3,255)
sql = " INSERT INTO [VoteInfo] (" 
sql = sql& "  KeyID" 
sql = sql& ", KeyMod" 
sql = sql& ", KeyCode" 
sql = sql& ", InfType" 
sql = sql& ", InfSubj" 
sql = sql& ", InfRem1" 

sql = sql& ", InfTime1" 
sql = sql& ", InfTime2" 
sql = sql& ", InfNum1" 
sql = sql& ", InfNum2" 
sql = sql& ", InfCard" 
sql = sql& ", InfVNum" 

sql = sql& ", ImgName" 
sql = sql& ", LogAddIP" 
sql = sql& ", LogAUser" 
sql = sql& ", LogATime" 
sql = sql& ")VALUES(" 
sql = sql& "  '" & KeyID &"'" 
sql = sql& ", '" & ModID &"'"  '
sql = sql& ", '" & KeyCode &"'" 
sql = sql& ", '" & InfType &"'"  'ModID
sql = sql& ", '" & InfSubj &"'" 
sql = sql& ", '" & InfRem1 &"'" 

sql = sql& ", '" & RequestSafe(upload.form("InfTime1"),"D",Date()) &"'" 
sql = sql& ", '" & RequestSafe(upload.form("InfTime2"),"D",Date()) &"'" 
sql = sql& ", " & RequestSafe(upload.form("InfNum1"),"N",1) &"" 
sql = sql& ", " & RequestSafe(upload.form("InfNum2"),"N",10) &"" 
sql = sql& ", '" & RequestSafe(upload.form("InfCard"),"C",2) &"'" 
sql = sql& ", '" & RequestSafe(upload.form("InfVNum"),"C",24) &"'" 

sql = sql& ", ''"  'ImgName
sql = sql& ", '" & IP &"'" 
sql = sql& ", '" & Session("UsrID") &"'" 
sql = sql& ", '" & Now() &"'" 
sql = sql& ")"
  
    MsgFlag = "N" : fMsg = ""
  If Trim(InfSubj)<>"" AND Trim(InfRem1)<>"" Then
    Call rs_Dosql(conn,sql)	
	Session("ChkCode") = "" '// 马上清除　
	MsgFlag = "Y"
	'///////////////////////////
	For i=1 To 5
	  ColID = "ImgNam"&i
	 If i=1 Then 
	  ColID = "ImgName"
	 End If
	'///////////////////////////
	set up_file = upload.File(ColID)
    if up_file.FileSize>0 And ImgTitle="" then
 	   save_name = Year(Now)&"-"&Get_FmtID("mdhnsx","")&"_"&Rnd_ID("KEY",5)&Mid(up_file.FileName,InStrRev(up_file.FileName,"."),8)
	   save_Title = up_file.FileName
	   fType = TypeFile&"/"&TypeImg
	   upFlag = chk_file(up_file,PathImg&save_name,fType,intPSize) 
       if upFlag = "OK" then
		  up_file.SaveAs Server.mappath(PathImg&save_name)
		  sql2 = "UPDATE [VoteInfo] SET "&ColID&"='"&save_name&"' WHERE KeyID='"&KeyID&"'"
		  Call rs_Dosql(conn,sql2)
	        fMsg = fMsg & "\n["&save_name&"] Upload OK!"
	   else
	      if upFlag = "Exist" then
	        fMsg = fMsg & "\n !!! ["&save_name&"] Exist!"
		  elseif upFlag = "Type" then
	        fMsg = fMsg & "\n !!! ["&save_name&"] Type Error!"
		  elseif upFlag = "Size" then
	        fMsg = fMsg & "\n !!! ["&save_name&"] Size Too Big"
		  end if
	   end if
     end if
     set up_file = Nothing
	 Next
  End If
  
set upload = nothing

If MsgFlag = "Y" Then
  msg = "信息发布成功!"
Else
  msg = "信息发布失败!"
End If

'Call rs_GetFile(cKeyID,"") 
Session("KeyID") = ""
ReDir = "info_list.asp"
If ReEnd="Y" Then ReDir="info_add.asp?ReEnd="&ReEnd&"&ReTyp="&InfType&""
Response.Write js_Alert(msg&fMsg,"Redir",ReDir) 
%>

</body>
</html>
