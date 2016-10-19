<!--#include file="../file/iconfig.asp"-->
<!--#include file="../../sadm/func1/up_class.asp"-->
<!--#include file="../../sadm/func2/func_obj.asp"-->
<!--#include file="../../upfile/sys/pcfg/wmark.asp"-->
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

set upload=new upload_Class

TabID = "InfoPhoto" 
ColID = "ImgName" 
IR = upload.form("IR")
bakID = upload.form("bakID")
bakExt=upload.form("bakExt")
KeyID=bakID&bakExt
upPath = Replace(IR,"-","/")

  Dim sys27_Rnd(10)
  sys27_RVal = upload.form(App27Random)
  If sys27_RVal&"" = "" Then
    Response.End()
  Else
    For i = 1 To 9
     sys27_Rnd(i)=Mid(sys27_RVal,i*6,Mid(sys27_RVal,i,1))
    Next
  End If
  
  InfSubj = RequestSafe(upload.form("InfSubj"&sys27_Rnd(1)),3,255) 
  InfCont = RequestSafe(upload.form("InfCont"&sys27_Rnd(4)),3,255) 
  SetTop = RequestSafe(upload.form("SetTop"),"N",888)

	fMsg = ""
	fp = "ImgName1" 
	set up_file = upload.File(fp)
    if up_file.FileSize>0 then
 	   save_name = upPath&"/"&Get_FmtID("mdhnsx","")&"_"&Rnd_ID("KEY",5)&Mid(up_file.FileName,InStrRev(up_file.FileName,"."),8) 
	   upFlag = chk_file(up_file,Config_Path&"upfile/"&save_name,fFiles&fImage&fDocus&fMedia,4123) 
       if upFlag = "OK" then
		  'Response.Write Config_Path&"upfile/"&save_name
		  Call fold_add9(Config_Path&"upfile/",upPath,0)
		  up_file.SaveAs Server.mappath(Config_Path&"upfile/"&save_name)
		  Call ImgMMark(Config_Path&"upfile/"&save_name,wmk_Mark)
		  save_name = Replace(Config_Path&"upfile/"&save_name,Config_Path&"upfile/","")
  '//////////////////////////////////////////////////////////////////////			  
	sql = " INSERT INTO InfoPhoto (KeyID, KeyRe, KeyMod"  
	sql = sql& ", InfSubj, InfCont, SetTop, ImgName" 
	sql = sql& ", LogAddIP, LogAUser, LogATime" 
	sql = sql& ")VALUES(" 
	sql = sql& "  '" & Get_AutoID(24) &"', '" & IR &"', '" & ModID &"'" 
	sql = sql& ", '" & InfSubj &"', '"&InfCont&"', " & SetTop &", '"&save_name&"'" 
	sql = sql& ", '" & Get_CIP() &"', '" & Session("UsrID") &"', '" & Now() &"'" 
	sql = sql& ")"
    Call rs_DoSql(conn,sql)
  '//////////////////////////////////////////////////////////////////////
	        fMsg = fMsg & "\n !!! ["&save_name&"]文件上传成功!"
	   else
	      if upFlag = "Exist" then
	        fMsg = fMsg & "\n !!! ["&save_name&"]文件上传失败,文件已经存在!"
		  elseif upFlag = "Type" then
	        fMsg = fMsg & "\n !!! ["&save_name&"]文件上传失败,文件类型禁止上传!"
		  elseif upFlag = "Size" then
	        fMsg = fMsg & "\n !!! ["&save_name&"]文件上传失败,文件太大!"
		  end if
	   end if
     end if

	 set up_file = Nothing
	 set upload = Nothing

  Response.Write "<div style='padding:3px 0px 0px 24px'>提示:["&bakExt&"]"&msg&Replace(fMsg,"\n","<br>")&"</div>"
%>
</body>
</html>
