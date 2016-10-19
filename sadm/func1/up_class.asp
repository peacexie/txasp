<%

'<!-- Huajing Files Updoad(Modify); ---Peace(XieYongshun)2005/10/05 -->

dim Data_5xsoft 'upload_5xsoft
Class upload_Class 
dim objForm,objFile,Version

Public function Form(strForm)
   strForm=lcase(strForm)
   if not objForm.exists(strForm) then
     Form=""
   else
     Form=objForm(strForm)
   end if
End function

Public function File(strFile)
   strFile=lcase(strFile)
   if not objFile.exists(strFile) then
     set File=new FileInfo
   else
     set File=objFile(strFile)
   end if
End function

Private Sub Class_Initialize 
  dim RequestData,sStart,vbCrlf,sInfo,iInfoStart,iInfoEnd,tStream,iStart,theFile
  dim iFileSize,sFilePath,sFileType,sFormValue,sFileName
  dim iFindStart,iFindEnd
  dim iFormStart,iFormEnd,sFormName
  Version="化境HTTP上传程序 Version 2.0"
  set objForm=Server.CreateObject("Scripting.Dictionary")
  set objFile=Server.CreateObject("Scripting.Dictionary")
  if Request.TotalBytes<1 then Exit Sub
  set tStream = Server.CreateObject("adodb.stream")
  set Data_5xsoft = Server.CreateObject("adodb.stream")
  Data_5xsoft.Type = 1
  Data_5xsoft.Mode =3
  Data_5xsoft.Open
  Data_5xsoft.Write  Request.BinaryRead(Request.TotalBytes)
  Data_5xsoft.Position=0
  RequestData =Data_5xsoft.Read 

  iFormStart = 1
  iFormEnd = LenB(RequestData)
  vbCrlf = chrB(13) & chrB(10)
  sStart = MidB(RequestData,1, InStrB(iFormStart,RequestData,vbCrlf)-1)
  iStart = LenB (sStart)
  iFormStart=iFormStart+iStart+1
  while (iFormStart + 10) < iFormEnd 
	iInfoEnd = InStrB(iFormStart,RequestData,vbCrlf & vbCrlf)+3
	tStream.Type = 1
	tStream.Mode =3
	tStream.Open
	Data_5xsoft.Position = iFormStart
	Data_5xsoft.CopyTo tStream,iInfoEnd-iFormStart
	tStream.Position = 0
	tStream.Type = 2
	tStream.Charset ="utf-8"
	sInfo = tStream.ReadText
	tStream.Close
	'取得表单项目名称
	iFormStart = InStrB(iInfoEnd,RequestData,sStart)
	iFindStart = InStr(22,sInfo,"name=""",1)+6
	iFindEnd = InStr(iFindStart,sInfo,"""",1)
	sFormName = lcase(Mid (sinfo,iFindStart,iFindEnd-iFindStart))
	'如果是文件
	if InStr (45,sInfo,"filename=""",1) > 0 then
	    set theFile=new FileInfo
		'取得文件名
		iFindStart = InStr(iFindEnd,sInfo,"filename=""",1)+10
		iFindEnd = InStr(iFindStart,sInfo,"""",1)
		sFileName = Mid (sinfo,iFindStart,iFindEnd-iFindStart)
		theFile.FileName=getFileName(sFileName)
		theFile.FilePath=getFilePath(sFileName)
		'取得文件类型
		iFindStart = InStr(iFindEnd,sInfo,"Content-Type: ",1)+14
		iFindEnd = InStr(iFindStart,sInfo,vbCr)
		theFile.FileType =Mid (sinfo,iFindStart,iFindEnd-iFindStart)
		theFile.FileStart =iInfoEnd
		theFile.FileSize = iFormStart -iInfoEnd -3
		theFile.FormName=sFormName
		if not objFile.Exists(sFormName) then
		  objFile.add sFormName,theFile
		end if
	else
	'如果是表单项目
		tStream.Type =1
		tStream.Mode =3
		tStream.Open
		Data_5xsoft.Position = iInfoEnd 
		Data_5xsoft.CopyTo tStream,iFormStart-iInfoEnd-3
		tStream.Position = 0
		tStream.Type = 2
		tStream.Charset ="utf-8"
	        sFormValue = tStream.ReadText 
	        tStream.Close
		if objForm.Exists(sFormName) then
		  objForm(sFormName)=objForm(sFormName)&", "&sFormValue		  
		else
		  objForm.Add sFormName,sFormValue
		end if
	end if
	iFormStart=iFormStart+iStart+1
	wend
  RequestData=""
  set tStream =nothing
End Sub

Private Sub Class_Terminate  
 if Request.TotalBytes>0 then
	objForm.RemoveAll
	objFile.RemoveAll
	set objForm=nothing
	set objFile=nothing
	Data_5xsoft.Close
	set Data_5xsoft =nothing
 end if
End Sub
   
 
 Private function GetFilePath(FullPath)
  If FullPath <> "" Then
   GetFilePath = left(FullPath,InStrRev(FullPath, "\"))
  Else
   GetFilePath = ""
  End If
 End  function
 
 Private function GetFileName(FullPath)
  If FullPath <> "" Then
   GetFileName = mid(FullPath,InStrRev(FullPath, "\")+1)
  Else
   GetFileName = ""
  End If
 End  function
 
End Class

Class FileInfo
  dim FormName,FileName,FilePath,FileSize,FileType,FileStart
  Private Sub Class_Initialize 
    FileName = ""
    FilePath = ""
    FileSize = 0
    FileStart= 0
    FormName = ""
    FileType = ""
  End Sub
  
 Public function SaveAs(FullPath)
    dim dr,ErrorChar,i
    SaveAs=true
    if trim(fullpath)="" or FileStart=0 or FileName="" or right(fullpath,1)="/" then exit function
    set dr=CreateObject("Adodb.Stream")
    dr.Mode=3
    dr.Type=1
    dr.Open
    Data_5xsoft.position=FileStart
    Data_5xsoft.copyto dr,FileSize
    dr.SaveToFile FullPath,2
    dr.Close
    set dr=nothing 
	Call ChekFType(FullPath,FileSize,FileStart) ''////////////////////////////////////////////////////////
    SaveAs=false 
  end function
  
	''/////////////////////////////////////////////////////////////////////////////////
	Public Function ChekFType(xPFile,xSize,xStart)
	  Dim e,fMov,fLog,s,str,stm,nMax
	  e = Mid(xPFile,InStrRev(xPFile,"."),12) 
	  nMax = 640123 '320123(320K),640123(640K),1200123(1.2M),2400123(2.4M)
	  fLog = false 
	  '一般在之前:chk_file() 就检测出来了，这里再检查一次
	  If inStr(lCase(fATyps),lCase(e))<=0 Then
		fMov = "Type"
	  ElseIf inStr(Session(UsrPStr),"{Admin}")>0 Then 'Admin管理员不检查？
	    'Pass
		'Response.End()
	  Else
		'Dim t1 :t1 = Timer()
		't1 = Timer()-t1
		If xSize>nMax Then 'nMax 以上，不检查但记录！
		  fLog = true
		Else 'nMax 以下，检查 'txt,office,pdf不检查???
		  '几乎所有时间都花在这里.
		  'str = File_Read(xPFile,"iso-8859-1") 'lCase() '以下获得str比File_Read高效率???	
		  set stm = server.CreateObject("Adodb.Stream")
		  With stm
			.Type = 1 :.Mode = 3 :.Open
			Data_5xsoft.position = xStart
			Data_5xsoft.copyto stm,xSize
			.Position = 0 :.Type = 2 :.Charset = "iso-8859-1"
			str = stm.ReadText
			.Close
		  End With
		  set stm = nothing
		  fMov = ChkTrojan(str,"")
		End If
		'echo "<br>tRead: "&t1 :t1 = Timer()
		'Response.End()
	  End If
	  If fMov<>"" Or fLog Then
		s = Mid(xPFile,InStrRev(xPFile,"\"),48)
		s = Replace(s,e,e&".Peace!Del")
		If fMov<>"" Then 
		  Call fil_move(xPFile,Config_Path&"upfile/#dbf#"&s)
		  e = "fMov:"&fMov
		Else
		  e = "fLog:"&FormatNumber(xSize/(1024*1024),2)&"M"
		End If
		s = Replace(xPFile,Server.MapPath("/"),"")
		s = Now()&" ^ "&Get_CIP()&" ^ "&e&" ^ "&Request.ServerVariables("URL")&" ^ "&s&" ^ "&_
		Session("UsrID")&","&Session("InnID")&","&Session("MemID")&" ^ "&_
		Request.Servervariables("HTTP_USER_AGENT")
		Call fil_app(Config_Path&"upfile/#dbf#/"&Get_yyyymmdd("")&".txt",s)
	  End If
	End Function
	''/////////////////////////////////////////////////////////////////////////////////
  
End Class

%>
