<!--#include file="../../upfile/sys/pcfg/editor.asp"-->
<!--#include file="edt_config.asp"-->
<!--#include file="../../sadm/func1/up_class.asp"-->
<!--#include file="../../sadm/func2/func_obj.asp"-->
<!--#include file="../../upfile/sys/pcfg/wmark.asp"-->
<!--#include file="../../ext/jqplug/json.asp"-->
<% 
'require '_cofig.php'; 
'require CORE_ROOT.'ext/jqplug/json.php';
'require 'edt_config.php';

sPerm = Chk_PEditor("")
if sPerm="(Pass)" then   
  if(Session("upPart")<>"") Then
    upPart = Session("KeyID")
  else                     
    upPart = "defup"
  end if
  If Len(Session("KeyID"))>15 Then
	upPath = Replace(Session("KeyID"),"-","/")  
  Else
	tDate = Now()
	tSTab = "123456789ABCDEFGHJKMNPQRSTUVWXY"
	mStr = Mid(tSTab,DatePart("m",tDate),1)
	mStr = mStr&Mid(tSTab,DatePart("d",tDate),1)
	upPath = "defup/"&Year(Now)&"/"&mStr
  End If
  Call fold_add9(Config_Path&"upfile/",upPath,0) 
end if

If edcEditID="edck3" Then ' /// edck3 //////////////////////////

  if sPerm<>"(Pass)" Then 
    Response.Write "<"&"script"&" type='text/javascript'>"&vbcrlf
	Response.Write "alert('("&sPerm&")NoPerm!')"&vbcrlf
	Response.Write "</"&"script"&">"&vbcrlf
	Response.End()
  end if
  
  set upload=new upload_Class
  set up_file = upload.File("upload")
    if up_file.FileSize>0 then
	   save_ext = Mid(up_file.FileName,InStrRev(up_file.FileName,"."),8)
	   save_name = upPath&"/"&Get_FmtID("mdhnsx","")&"_"&Rnd_ID("KEY",5)&save_ext
	   If Session("UsrID")&""="" And Session("MemID")="guest" Then 
	     intPSize = intGPSize '180
	   End If
	   upFlag = chk_file(up_file,Config_Path&"upfile/"&save_name,fFiles&fImage&fDocus&fMedia,intPSize) 
	   if upFlag = "OK" then
		  up_file.SaveAs Server.mappath(Config_Path&"upfile/"&save_name)
		  Call ImgMMark(Config_Path&"upfile/"&save_name,wmk_Mark)
	      fMsg = ""'fMsg & "\n !!! ["&save_name&"]文件上传成功!"
		  fVal = "2"
		  fileUrl = Config_Path&"upfile/"&save_name
	   else
		  fVal = "2"
		  if upFlag = "Exist" then
	        fMsg = fMsg & "\n !!! ["&save_ext&"]文件重名,文件上传成功!"
		  elseif upFlag = "Type" then
	        fMsg = fMsg & "\n !!! ["&save_ext&"]文件上传失败,文件类型禁止上传!"
		  elseif upFlag = "Size" then
	        fMsg = fMsg & "\n !!! ["&save_ext&"]文件上传失败,文件太大!"
		  end if
	   end if
     else
		   fMsg = "\n !!! [空文件],文件上传失败!"
	 end if
  set up_file = Nothing
  set upload = nothing

  Response.Write "<"&"script"&" type='text/javascript'>"&vbcrlf
  Response.Write "window.parent.CKEDITOR.tools.callFunction('"&fVal&"', '"&fileUrl&"','"&fMsg&"');"&vbcrlf
  Response.Write "</"&"script"&">"&vbcrlf
  Response.End()
    'exit($str);
	
ElseIf edcEditID="edkind" Then ' /// edkind ///////////////////////////////////////////////////
  
  if sPerm<>"(Pass)" then 
    Call json_Alert(1,"("&sPerm&")NoPerm!","msg")
	Response.End()
  end if
  
  set upload=new upload_Class
  set up_file = upload.File("imgFile")
    if up_file.FileSize>0 then
	   save_name = upPath&"/"&Get_FmtID("mdhnsx","")&"_"&Rnd_ID("KEY",5)&Mid(up_file.FileName,InStrRev(up_file.FileName,"."),8)
	   If Session("UsrID")&""="" And Session("MemID")="guest" Then 
	     intPSize = intGPSize '180
	   End If
	   upFlag = chk_file(up_file,Config_Path&"upfile/"&save_name,fFiles&fImage&fDocus&fMedia,intPSize) 
	   if upFlag = "OK" then
		  up_file.SaveAs Server.mappath(Config_Path&"upfile/"&save_name)
		  Call ImgMMark(Config_Path&"upfile/"&save_name,wmk_Mark)
	      fMsg = fMsg & "\n !!! ["&save_name&"]文件上传成功!"
		  fileUrl = Config_Path&"upfile/"&save_name
		  Call json_Alert(0,fileUrl,"url")
	   else
		  if upFlag = "Exist" then
	        fMsg = fMsg & "\n !!! ["&save_name&"]文件重名,文件上传成功!"
		  elseif upFlag = "Type" then
	        fMsg = fMsg & "\n !!! ["&save_name&"]文件上传失败,文件类型禁止上传!"
		  elseif upFlag = "Size" then
	        fMsg = fMsg & "\n !!! ["&save_name&"]文件上传失败,文件太大!"
		  end if
		  Call json_Alert(1,fMsg,"msg")
	   end if
     else
	      Call json_Alert(1,"[空文件],文件上传失败!","msg")
		   'fMsg = "\n !!! [空文件],文件上传失败!"
	 end if
  set up_file = Nothing
  set upload = nothing

ElseIf edcEditID="edxh1" Then ' /// edck3 //////////////////////////

  if sPerm<>"(Pass)" then 
	Response.Write "{'err':'"+jsonString("("&sPerm&")NoPerm!")+"','msg':''}"
	Response.End()
  end if
  
  set upload=new upload_Class
  set up_file = upload.File("filedata")
    if up_file.FileSize>0 then
	   save_name = upPath&"/"&Get_FmtID("mdhnsx","")&"_"&Rnd_ID("KEY",5)&Mid(up_file.FileName,InStrRev(up_file.FileName,"."),8)
	   If Session("UsrID")&""="" And Session("MemID")="guest" Then 
	     intPSize = intGPSize '180
	   End If
	   upFlag = chk_file(up_file,Config_Path&"upfile/"&save_name,fFiles&fImage&fDocus&fMedia,intPSize) 
	   if upFlag = "OK" then
		  up_file.SaveAs Server.mappath(Config_Path&"upfile/"&save_name)
		  Call ImgMMark(Config_Path&"upfile/"&save_name,wmk_Mark)
	      fMsg = fMsg & "["&save_name&"]文件上传成功!"
		  fileUrl = Config_Path&"upfile/"&save_name
		  jMsg = "{'url':'"+jsonString(fileUrl)+"','localname':'"+jsonString(up_file.FileName)+"','id':'"&jsonString(save_name)&"'}"
		  Response.Write "{'err':'','msg':"+jMsg+"}"
		  Response.End()
	   else
		  if upFlag = "Exist" then
	        fMsg = fMsg & "\n !!! ["&save_name&"]文件重名,文件上传成功!"
		  elseif upFlag = "Type" then
	        fMsg = fMsg & "\n !!! ["&save_name&"]文件上传失败,文件类型禁止上传!"
		  elseif upFlag = "Size" then
	        fMsg = fMsg & "\n !!! ["&save_name&"]文件上传失败,文件太大!"
		  end if
		  Response.Write "{'err':'"+jsonString(""&fMsg&"")+"','msg':''}"
		  Response.End()
	   end if
     else
		  Response.Write "{'err':'"+jsonString("[空文件],文件上传失败!")+"','msg':''}"
		  Response.End()
	 end if
  set up_file = Nothing
  set upload = nothing
  Function jsonString(str)
	  str=replace(str,"\","\\")
	  str=replace(str,"/","\/")
	  str=replace(str,"'","\'")
	  jsonString=str
  End Function
  'msg="{'url':'"+target+"','localname':'"+jsonString(upfile.file(inputname).FileName)+"','id':'1'}"
  'Response.Write "{'err':'"+jsonString(err)+"','msg':"+msg+"}"
  
ElseIf edcEditID="edfck" Then ' /// edfck ///////////////////////////////////////////////////

  if sPerm<>"(Pass)" Then 
    Response.Write("("&sPerm&")NoPerm!")
	Response.End()
  end if
  
  set upload=new upload_Class
  set up_file = upload.File("NewFile")
    if up_file.FileSize>0 then
	   save_name = upPath&"/"&Get_FmtID("mdhnsx","")&"_"&Rnd_ID("KEY",5)&Mid(up_file.FileName,InStrRev(up_file.FileName,"."),8)
	   If Session("UsrID")&""="" And Session("MemID")="guest" Then 
	     intPSize = intGPSize '180
	   End If
	   upFlag = chk_file(up_file,Config_Path&"upfile/"&save_name,fFiles&fImage&fDocus&fMedia,intPSize) 
	   if upFlag = "OK" then
		  up_file.SaveAs Server.mappath(Config_Path&"upfile/"&save_name)
		  Call ImgMMark(Config_Path&"upfile/"&save_name,wmk_Mark)
	      fMsg = fMsg & "\n !!! ["&save_name&"]文件上传成功!"
		  fVal = 101
		  fileUrl = Config_Path&"upfile/"&save_name
	   else
	      vFal = "1"
		  fVal = "1"
		  if upFlag = "Exist" then
	        fMsg = fMsg & "\n !!! ["&save_name&"]文件重名,文件上传成功!"
		  elseif upFlag = "Type" then
	        fMsg = fMsg & "\n !!! ["&save_name&"]文件上传失败,文件类型禁止上传!"
		  elseif upFlag = "Size" then
	        fMsg = fMsg & "\n !!! ["&save_name&"]文件上传失败,文件太大!"
		  end if
	   end if
     else
	       vFal = "1"
		   fMsg = "\n !!! [空文件],文件上传失败!"
	 end if
  set up_file = Nothing
  set upload = nothing
  
  Response.Write "<"&"script"&" type='text/javascript'>"&vbcrlf
  Response.Write "(function(){var d=document.domain;while (true){try{var A=window.parent.document.domain;break;}catch(e) {};d=d.replace(/.*?(?:\.|$)/,'');if (d.length==0) break;try{document.domain=d;}catch (e){break;}}})();"&vbcrlf
  Response.Write "	window.parent.OnUploadCompleted("&fVal&",'"&fileUrl&"','(File Name)','"&fMsg&"') ;"&vbcrlf
  Response.Write "</"&"script"&">"&vbcrlf

Else

  Response.Write "unKnow("&edcEditID&")"&vbcrlf
  Response.End()
	
end if

%>
