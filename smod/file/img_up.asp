<!--#include file="iconfig.asp"-->
<!--#include file="../../sadm/func1/up_class.asp"-->
<!--#include file="../../sadm/func2/func_obj.asp"-->
<!--#include file="../../upfile/sys/pcfg/wmark.asp"-->
<%
Call Chk_URL()

set upload=new upload_Class

TabID= RequestSafe(upload.form("TabID"),3,48)
upPath= RequestSafe(upload.form("upPath"),3,120)

ID   = RequestSafe(upload.form("ID"),3,48)
KeyID= RequestSafe(upload.form("KeyID"),3,48)
ColID= RequestSafe(upload.form("ColID"),3,48)
NO = RequestSafe(upload.form("NO"),"N",0)
WW = upload.form("WW")

	fMsg = ""
	fp = ColID 
	set up_file = upload.File(fp)
    if up_file.FileSize>0 then
 	   save_name = upPath&"/"&Get_FmtID("mdhnsx","")&"_"&Rnd_ID("KEY",5)&Mid(up_file.FileName,InStrRev(up_file.FileName,"."),8) 
	   If PrmFlag="(Mem)" Or Session("UsrID")&""="" Then 
	     intPSize = intGPSize '180
	   End If
	   upFlag = chk_file(up_file,Config_Path&"upfile/"&save_name,fFiles&fImage&fDocus&fMedia,intPSize) 
       if upFlag = "OK" then
		  up_file.SaveAs Server.mappath(Config_Path&"upfile/"&save_name)
		  Call ImgMMark(Config_Path&"upfile/"&save_name,wmk_Mark)
		  save_name = Replace(fPath&save_name,Config_Path&"upfile/","/upfile/")
  '//////////////////////////////////////////////////////////////////////			  
  If NO<>"0" Then
    cName = rs_Val(""," SELECT "&ColID&" FROM "&TabID&" WHERE "&KeyID&"='"&ID&"'")
	aImg = Split(cName&"^^^^^^^^^^","^")
	sName = ""
	For i=1 To 9
	  If Int(NO)=i Then
	    p = InStrRev(save_name,"/")
	    save_name = Mid(save_name,p+1)
	    sName = sName&"^"&save_name
	  Else
	    sName = sName&"^"&aImg(i)
	  End If
	Next
	sql = "UPDATE "&TabID&" SET "&ColID&"='"&sName&"' WHERE "&KeyID&"='"&ID&"'"
    Call rs_DoSql(conn,sql)
  Else
	sName = save_name
	sql = "UPDATE "&TabID&" SET "&ColID&"='"&sName&"' WHERE "&KeyID&"='"&ID&"'"
    Call rs_DoSql(conn,sql)
  End if
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

Response.Write js_Alert("图片修改_"&fMsg,"Redir","img_set.asp?TabID="&TabID&"&upPath="&upPath&"&ID="&ID&"&KeyID="&KeyID&"&ColID="&ColID&"&NO="&NO&"&WW="&WW&"") 

%>
