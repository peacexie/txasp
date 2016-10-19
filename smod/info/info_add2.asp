<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="config.asp"-->
<!--#include file="../../sadm/func1/up_class.asp"-->
<!--#include file="../../sadm/func2/func_obj.asp"-->
<!--#include file="../../sadm/func2/upremote.asp"-->
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
 	
Call Chk_URL()	

IP = Get_CIP()
If Len(Session("KeyID"))<15 Then 
  KeyID = rs_AutID(conn,ModTab,"KeyID",upPart,"1","")
Else
  KeyID = Session("KeyID")
End If
'upPath = upRoot&Replace(KeyID,"-","/")&"/" 

set upload=new upload_Class

PrmFlag = upload.form("PrmFlag")
Call Chk_Perm9(ModID,PrmFlag)

bakID = upload.form("bakID")
If bakID<>"" Then
  bakExt=upload.form("bakExt")
  KeyID=bakID&bakExt '//Rnd_ID("bakExt",2) 
End If
upPath = upRoot&Replace(KeyID,"-","/")&"/"

  Dim sys27_Rnd(10)
  sys27_RVal = upload.form(App27Random)
  If sys27_RVal&"" = "" Then
    Response.End()
  Else
    For i = 1 To 9
     sys27_Rnd(i)=Mid(sys27_RVal,i*6,Mid(sys27_RVal,i,1))
    Next
  End If

KeyCode = RequestSafe(upload.form("KeyCode"),3,48)
InfSubj = RequestSafe(upload.form("InfSubj"&sys27_Rnd(1)),3,255) 
InfCont = upload.form("InfCont"&sys27_Rnd(4))
InfCont = RequestSafe(InfCont,3,960000) 
InfCont = Show_Html(InfCont)
If SwhRemSave="Y" Then
InfCont = RemoteReplaceUrl(InfCont, upRoot, KeyID)
End If
If Config_Cont="DB" Then
  xxxCont = InfCont
Else
  xxxCont = ""
End If
InfType = RequestSafe(upload.form("InfType"),3,255)
InfTyp2 = RequestSafe(upload.form("InfTyp2"),3,48)
ReEnd = upload.form("ReEnd")
InfPara = PrmFlag
For i=1 To 96 
  iPara = RequestSafe(upload.form("InfPara"&i),3,1200)
  iPara = Replace(iPara,"^","")
  InfPara = InfPara&"^"&iPara
Next
SetSubj = RequestSafe(upload.form("SetSubj"),3,12) 
SetHot = RequestSafe(upload.form("SetHot"),3,2) 
SetTop = RequestSafe(upload.form("SetTop"),3,12) 
SetShow = RequestSafe(upload.form("SetShow"),3,2) 
ModImgCount = RequestSafe(upload.form("ModImgCount"),"N",1)
LogATime = RequestSafe(upload.form("LogATime"),"D",Now())
ChkCode = uCase(upload.form("ChkCode")) 
sql = " INSERT INTO "&ModTab&" (" 
sql = sql& "  KeyID" 
sql = sql& ", KeyMod" 
sql = sql& ", KeyCode" 
sql = sql& ", InfType" 
sql = sql& ", InfTyp2" 
sql = sql& ", InfSubj"  
sql = sql& ", InfCont" 
sql = sql& ", InfPara"
sql = sql& ", SetSubj" 
sql = sql& ", SetRead" 
sql = sql& ", SetHot" 
sql = sql& ", SetTop" 
sql = sql& ", SetShow" 
sql = sql& ", ImgName" 
sql = sql& ", LogAddIP" 
sql = sql& ", LogAUser" 
sql = sql& ", LogATime" 
sql = sql& ")VALUES(" 
sql = sql& "  '" & KeyID &"'" 
sql = sql& ", '" & ModID &"'" 
sql = sql& ", '" & KeyCode &"'" 
sql = sql& ", '" & InfType &"'" 
sql = sql& ", '" & InfTyp2 &"'" 
sql = sql& ", '" & InfSubj &"'" 
sql = sql& ", '" & xxxCont &"'" 
sql = sql& ", '" & InfPara &"'" 
sql = sql& ", '" & SetSubj &"'" 
sql = sql& ", 0" 
sql = sql& ", '" & SetHot &"'" 
sql = sql& ", '" & SetTop &"'" 
sql = sql& ", '" & SetShow &"'" 
sql = sql& ", '" & ImgName & "'" 
sql = sql& ", '" & IP &"'" 
sql = sql& ", '" & Get_PUser(PrmFlag) &"'" 
sql = sql& ", '" & LogATime &"'" 
sql = sql& ")"

If PrmFlag="(Mem)" Then
 If Session("ChkCode")<>ChkCode Or ChkCode&""="" Then
  Response.Write "<h1 style='line-height:180%; text-align:center'>认证码错误！</h1>"
  set up_file = Nothing
  Response.End()
 End If
End If
  
    MsgFlag = "N" ': Response.Write sql
	ImgName = ""
  If Trim(InfSubj)<>"" Then 'AND Trim(InfCont)<>"" 
    'Response.Write KeyID'&sql
	Call rs_Dosql(conn,sql)	
	Session("ChkCode") = "" '// 马上清除　
	Session("KeyID") = ""
	MsgFlag = "Y"
	'Response.Write ModImgCount
	For i=1 To ModImgCount
	'Response.Write "<br>"&i&ModImgCount
	set up_file = upload.File("ImgName"&i)
    if up_file.FileSize>0 then 'And ImgTitle=""
 	   save_name = Get_FmtID("mdhnsx","")&"_"&i&Rnd_ID("KEY",4)&Mid(up_file.FileName,InStrRev(up_file.FileName,"."),8)
	   fType = TypeFile&"/"&TypeImg
	   If PrmFlag="(Mem)" Or Session("UsrID")&""="" Then 'If Session("UsrID")&""="" And Session("MemID")="guest" Then 
	     intPSize = intGPSize '180
	   End If
	   upFlag = chk_file(up_file,upPath&save_name,fType,intPSize) 
       if upFlag = "OK" then
	      Call fold_add9(upRoot,KeyID,0) ': Response.Write upPath&save_name
		  up_file.SaveAs Server.mappath(upPath&save_name)
		  Call ImgMMark(upPath&save_name,wmk_Mark)
		  ImgName = ImgName&"^"&save_name
	      fMsg = fMsg & "\n["&save_name&"] 文件上传成功!"
	   else
	      if upFlag = "Exist" then
	        fMsg = fMsg & "\n !!! ["&save_name&"] 文件上传失败,文件已经存在!"
		  elseif upFlag = "Type" then
	        fMsg = fMsg & "\n !!! ["&save_name&"] 文件上传失败,文件类型禁止上传!"
		  elseif upFlag = "Size" then
	        fMsg = fMsg & "\n !!! ["&save_name&"] 文件上传失败,文件太大"
		  end if
	   end if
     end if
     set up_file = Nothing
	 Next
	 If ImgName<>"" Then
        Call rs_Dosql(conn,"UPDATE ["&ModTab&"] SET ImgName='"&ImgName&"' WHERE KeyID='"&KeyID&"'")
     Else
	    If upload.form("ImgSCopy")="Y" Then
		  Call ImgSUpd(ModTab,KeyID,InfCont,upload.form("ImgSmall1"),upload.form("ImgSmall2"))
		End If
	 End If
  End If
  
set upload = nothing

Call add_sfFile()

If MsgFlag = "Y" Then
  msg = "信息发布成功!"
Else
  msg = "信息发布失败!"
End If

If bakID<>"" Then
  Response.Write "<div style='padding:3px 0px 0px 24px'>提示:["&bakExt&"]"&msg&Replace(fMsg,"\n","<br>")&"</div>"
Else
 
  msg = msg&fMsg :If Session("MemID")="guest" Then msg = msg&"\n信息等待审核!"
  If ReEnd="Y" Then 
    ReDir="info_add.asp?ReEnd="&ReEnd&"&InfType="&InfType&"&InfTyp2="&InfTyp2&"&PrmFlag="&PrmFlag&""
  ElseIf ReEnd="C" Then
	ReDir = "guest_login.asp?act=out&msg="&Server.URLEncode(msg)
	Response.Redirect ReDir
  Else
    ReDir = "info_list.asp?PrmFlag="&PrmFlag&""
  End If
  Response.Write js_Alert(msg,"Redir",ReDir)
  
End If
%>

</body>
</html>
