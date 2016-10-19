<!--#include file="fconfig.asp"-->
<!--#include file="../../sadm/func1/up_class.asp"-->
<!--#include file="../../sadm/func2/func_obj.asp"-->
<!--#include file="../../upfile/sys/pcfg/wmark.asp"-->
<%

goREFERE = Request.ServerVariables("HTTP_REFERER")
Call Chk_URL()
'Response.Write goREFERE

set upload=new upload_Class

	yPath = upload.form("yPath")
	fAuto = upload.form("ImgAuto")
	goPage = upload.form("goPage")
	
	set up_file = upload.File("ImgName")
    if up_file.FileSize>0 then
 	   If fAuto="AutoName" Then
	     save_name = Get_FmtID("mdhnsx","")&"_"&Rnd_ID("KEY",5)&Mid(up_file.FileName,InStrRev(up_file.FileName,"."),8)
	   Else
		 save_name = up_file.FileName  
	   End If
	   upFlag = chk_file(up_file,Config_Path&"upfile/"&yPath&"/"&save_name,fFiles&fImage&fDocus&fMedia,4123) 
	   if upFlag = "OK" then
		    up_file.SaveAs Server.mappath(Config_Path&"upfile/"&yPath&""&save_name)
			Call ImgMMark(Config_Path&"upfile/"&yPath&""&save_name,wmk_Mark)
	        fMsg = fMsg & "\n !!! ["&save_name&"]文件上传成功!"
	   else
	      if upFlag = "Exist" then
		    fName = Replace(save_name,".","-"&Rnd_ID("KEY",3)&".")
			up_file.SaveAs Server.mappath(Config_Path&"upfile/"&yPath&""&fName)
	        fMsg = fMsg & "成功!\n重命名为:["&fName&"]\n !!! 文件上传成功!"
	        'fMsg = fMsg & "\n !!! ["&save_name&"]文件上传失败,文件已经存在!"
		  elseif upFlag = "Type" then
	        fMsg = fMsg & "\n !!! ["&save_name&"]文件上传失败,文件类型禁止上传!"
		  elseif upFlag = "Size" then
	        fMsg = fMsg & "\n !!! ["&save_name&"]文件上传失败,文件太大!"
		  end if
	   end if
     else
	       fMsg = "\n !!! [空文件],文件上传失败!"
	 end if

set up_file = Nothing
set upload = nothing

Response.Write "<meta http-equiv='Content-Type' content='text/html; charset=utf-8'>"
If goPage="goRef" Then
  goPage = goREFERE
Else
  goPage = "file_view.php?yPath="&yPath&""
End If
Response.Write js_Alert("文件上传_"&fMsg,"Redir",goPage)

%>
