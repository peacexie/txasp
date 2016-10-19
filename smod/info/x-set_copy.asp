<!--#include file="config.asp"-->
<html>
<head>
<meta http-equiv="Pragma" content="no-cache">
<meta http-equiv="Expires" content="0">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title><%=Config_Name%>后台管理中心</title>
<link href="../../inc/adm_inc/adm_style.css" rel="stylesheet" type="text/css">
<script src="/sadm/func1/Func_JS.js" type="text/javascript"></script>
</head>
<body>
<%

Act=Request("Act")
ModID = RequestS("ModID","C",48) '// PicS224
ModIX = Left(ModID,4)&"1"&Right(ModID,2)
ModNO = Mid(ModID,5,1) '// PicS224 --- 2
i = 0 : j = 0
'echo Act&"."&ModIX&"."&ModNO
'id = "dtpic-2010-6W-D3AV.01DXN<br>"
'echo id
'echo cpKeyID(id,ModNO)&"<br>"
'echo Next_ID("12A","12A",3)
'echo "<hr>" :Response.End()




If Act="Clear" Then '' SELECT [KeyCode] ; DELETE
  sql = "DELETE FROM "&ModTab&" WHERE [KeyMod]='"&ModID&"' AND [KeyCode] NOT IN(SELECT [KeyCode] FROM "&ModTab&" WHERE [KeyMod]='"&ModIX&"') "
  Call rs_DoSql(conn,sql)
ElseIf Act="DImp" Then 
  dSql = "DELETE FROM "&ModTab&" WHERE KeyMod='"&ModID&"' AND LogAddIP='(Import_X1)'"
  Call rs_DoSql(conn,dSql)
ElseIf Act="Copy1" Then
  
 ID = RequestS("ID","C",48)
 sql = "SELECT * FROM "&ModTab&" WHERE KeyID='"&ID&"'"
 Set rs=Server.Createobject("Adodb.Recordset")
 rs.Open sql,conn,1,1 
 If NOT rs.EOF Then
   KeyID = rs("KeyID") 'dtinf-2010-6-AB513.JUSX2
   KeyCode = rs("KeyCode")
   InfType = rs("InfType")
   InfTyp2 = rs("InfTyp2")
   InfSubj = rs("InfSubj") : Response.Write InfSubj 
   InfPara = rs("InfPara") 
   InfCont = rs("InfCont") 
   SetRead = rs("SetRead")
   SetSubj = rs("SetSubj")
   SetHot = rs("SetHot")
   SetTop = rs("SetTop") :SetTop = RequestSafe(SetTop,"N",0)
   SetShow = rs("SetShow")
   ImgName = rs("ImgName")
   LogATime = rs("LogATime") 
   flag = "OK"
 End If
 rs.Close()
 Set rs = Nothing
 
 If flag = "OK" Then
 set rs2 = server.createobject("ADODB.Recordset")
 sql = "SELECT * FROM "&ModTab&""
 rs2.Open sql,conn,1,3
 rs2.Addnew
   KeyID = rs_AutID(conn,ModTab,"KeyID",upPart,"1","")
   rs2("KeyID") = KeyID
   rs2("KeyMod") = ModID
   rs2("KeyCode") = Get_FmtID("mdhnsx","")&"-"&Rnd_ID("KEY",6)
   rs2("InfType") = InfType
   rs2("InfTyp2") = InfTyp2
   rs2("InfSubj") = InfSubj 
   rs2("InfPara") = InfPara 
   rs2("InfCont") = InfCont 
   rs2("SetRead") = SetRead
   rs2("SetSubj") = SetSubj
   rs2("SetHot") = SetHot
   rs2("SetTop") = SetTop 
   rs2("SetShow") = SetShow
   rs2("ImgName") = ImgName
   rs2("LogAddIP") = Get_CIP()
   rs2("LogAUser") = Get_PUser(PrmFlag)
   rs2("LogATime") = Now()
 rs2.Update
 rs2.Close
 set rs2 = nothing
 
	  Call fold_copy(upRoot&Replace(ID,"-","/"),upRoot&Replace(KeyID,"-","/"))
	  'Call rep_TmpCont(OldID,KeyID)
	  tFile = upRoot&Replace(KeyID,"-","/")&"/fcont.htm"
	  tCont = File_Read(tFile,"utf-8")
	  tCont = Replace(tCont,ID,KeyID)
	  If tCont<>"" And tCont<>"(File Read Error!)" Then
	    Call File_Add2(tFile,tCont,"utf-8")
	  End If
	  Call cpRelPic(KeyID,ID,ModID)
 
  Response.Redirect "info_edit.asp?ID="&KeyID 'info_edit.asp?ID=
 Else
  Response.Write "失败！"
  Response.End()
 End If
 
  
ElseIf Act<>"" Then
  Call Mod_Import() '同步
End If

If Act<>"" Then
  Response.Write "完成！"&Now()&"<BR>"
End If

%>

<a href="?ModID=<%=ModID%>&Act=Copy">Copy</a> | <a href="?ModID=<%=ModID%>&Act=Clear">Clear</a> | <a href="?ModID=<%=ModID%>&Act=DImp">DelImport</a>
<%

Sub Mod_Import()

 sql6 = "SELECT * FROM "&ModTab&" WHERE KeyMod='"&ModIX&"'" 
 Set rs=Server.Createobject("Adodb.Recordset")
 rs.Open sql6,conn,1,1 
 Do While NOT rs.EOF
 
   i=i+1
   KeyID = rs("KeyID") 'dtinf-2010-6-AB513.JUSX2
   OldID = KeyID
   KeyID = cpKeyID(KeyID,ModNO)
   KeyMod = ModID
   KeyCode = rs("KeyCode")
   InfType = rs("InfType")
	  iChar = Left(InfType,2) '' S110048;S120048; --- S1
	  jChar = Left(InfType,1)&ModNO '' PicS124,PicS224 --- S+1,2
	  InfType = Replace(InfType,iChar,jChar) '' S110048;S120048;
   InfTyp2 = rs("InfTyp2")
   InfSubj = rs("InfSubj") :InfSubj = Replace(InfSubj&"","'","")
   InfPara = rs("InfPara") :InfPara = Replace(InfPara&"","'","") 

   InfCont = rs("InfCont") :InfCont = Replace(InfCont&"","'","")
   
   SetRead = rs("SetRead")
   SetSubj = rs("SetSubj")
   SetHot = rs("SetHot")
   SetTop = rs("SetTop") :SetTop = RequestSafe(SetTop,"N",0)
   SetShow = rs("SetShow")
   ImgName = rs("ImgName")
   LogATime = rs("LogATime") 

	If rs_Exist(conn,"SELECT KeyCode FROM "&ModTab&" WHERE KeyMod='"&ModID&"' AND KeyCode='"&KeyCode&"'")="EOF" Then
	  j = j + 1
      sql = " INSERT INTO "&ModTab&" (" 
      sql = sql& "  KeyID" 
      sql = sql& ", KeyMod" 
      sql = sql& ", KeyCode" 
      sql = sql& ", InfType" 
      sql = sql& ", InfSubj" 
	  sql = sql& ", InfPara"
      sql = sql& ", InfCont" 
      sql = sql& ", SetSubj" 
      sql = sql& ", SetRead" 
      sql = sql& ", SetHot" 
      sql = sql& ", SetTop" 
      sql = sql& ", SetShow" 
      sql = sql& ", ImgName" 
      sql = sql& ", LogAddIP" 
      sql = sql& ", LogAUser" 
      sql = sql& ", LogATime" 
	  'sql = sql& ", LogEditIP"
      sql = sql& ")VALUES(" 
      sql = sql& "  '" & KeyID &"'" 
      sql = sql& ", '" & ModID & "'"
      sql = sql& ", '" & KeyCode &"'" 
      sql = sql& ", '" & InfType &"'" 
      sql = sql& ", '" & InfSubj & "'"
	  sql = sql& ", '" & InfPara & "'"
      sql = sql& ", '" & InfCont &"'" 
      sql = sql& ", '" & SetSubj &"'" 
      sql = sql& ", " & SetRead & "" 
      sql = sql& ", '" & SetHot &"'" 
      sql = sql& ", " & SetTop & "" 
      sql = sql& ", '" & SetShow &"'" 
      sql = sql& ", '" & ImgName &"'" 
      sql = sql& ", '(Import_X1)'" 
      sql = sql& ", '" & Session("UsrID") &"'" 
      sql = sql& ", '" & LogATime &"'" 
	  'sql = sql& ", '" & OldID &"'" 
      sql = sql& ")" ':Response.Write sql
	  Call rs_DoSql(conn,sql)
	  Call fold_copy(upRoot&Replace(OldID,"-","/"),upRoot&Replace(KeyID,"-","/"))
	  'Call rep_TmpCont(OldID,KeyID)
	  tFile = upRoot&Replace(KeyID,"-","/")&"/fcont.htm"
	  tCont = File_Read(tFile,"utf-8")
	  tCont = Replace(tCont,OldID,KeyID)
	  If tCont<>"" And tCont<>"(File Read Error!)" Then
	    Call File_Add2(tFile,tCont,"utf-8")
	  End If
	  Call cpRelPic(KeyID,OldID,ModID)
	  Response.Write vbcrlf&"<br>OK ----- "&j&"-"&i&InfType&" - "&InfSubj
	Else
	  Response.Write vbcrlf&"<br>NG ##### "&j&"-"&i&InfType&" - "&InfSubj
	End If

 rs.MoveNext
 Loop
 rs.Close()
 Set rs = Nothing
 
End Sub

Function cpKeyID(xID,xModNO)
  Dim p,e,id 'dtinf-2010-6W-CEF1.01DEM
  p = inStr(xID,".") '后3位SessionID
  e = Mid(xID,p+3,2) ':echo e&"<br>" 'Rnd_ID("KEY",21-p-1)
  e = Next_ID(e,e,2) ':echo e&"<br>"  'e = Replace(e,"_","")
  id = Left(xID,p-1)&"."&Mid(xID,P+1,2)&e&xModNO
  If id=xID Then 
  id = Left(xID,p-1)&"."&Rnd_ID("KEY",8)
  End If 
  cpKeyID = Left(id,24)
End Function

Function cpRelPic(xID,xOld,xMod) '复制相关文件
 Dim sql,rs,iSubj,iCont,iImg,iKey
 sql = "SELECT * FROM InfoPhoto WHERE KeyRe='"&xOld&"'"
 Set rs=Server.Createobject("Adodb.Recordset")
 rs.Open sql,conn,1,1 
 Do While NOT rs.EOF
   iSubj = rs("InfSubj")
   iCont = rs("InfCont")
   iTop = rs("SetTop")
   iImg = rs("ImgName")

	iImg = Replace(xID,"-","/")&Mid(iImg,25)
	sql = " INSERT INTO InfoPhoto (KeyID, KeyRe, KeyMod"  
	sql = sql& ", InfSubj, InfCont, SetTop, ImgName" 
	sql = sql& ", LogAddIP, LogAUser, LogATime" 
	sql = sql& ")VALUES(" 
	sql = sql& "  '" & Get_AutoID(24) &"', '" & xID &"', '" & xMod &"'" 
	sql = sql& ", '" & iSubj &"', '"&Replace(iCont,"'","")&"', " & iTop &", '"&iImg&"'" 
	sql = sql& ", '" & Get_CIP() &"', '" & Session("UsrID") &"', '" & Now() &"'" 
	sql = sql& ")"
	Call rs_DoSql(conn,sql)

 rs.MoveNext
 Loop
 rs.Close()
 Set rs = Nothing
 
End Function

%>
</body>
</html>
