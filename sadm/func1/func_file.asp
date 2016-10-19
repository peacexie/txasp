
<%

Dim fFiles,fImage,fDocus,fMedia,fATyps
fFiles = ".txt/.log/.rar/.zip/.pdf/.htm/.html/.gz/"
fImage = ".swf/.flv/.jpg/.gif/.jpeg/.png/.tif/.tif/.tiff/"
fDocus = ".doc/.xls/.ppt/.pps/.wps/.et/.dpt/"
fMedia = ".midi/.mid/.wav/.mp3/.aif/.aiff/.avi/.wmv/.mpg/.wma/.rm/.ram/.au/.mpg/.mpeg/.mov/.data/"
fATyps = fFiles&fImage&fDocus&fMedia
Dim fsoCalss,adsCalss
fsoCalss = "Scr"&"ipting.File"&"System"&"Object"
adsCalss = "AdoDB.Str"&"eam"
'// ../../upfile/dtinf/ Inf-2010-5-YGA30.6G7F0KM

Function fold_add9(xRoot,xCode,xSub)
  Dim a,s,i
  a = Split(Replace(xCode,"/","-"),"-")
  s = ""
  For i=0 To uBound(a)-xSub
    Call fold_add(xRoot&s,a(i))
	s = s&a(i)&"/"
  Next
End Function

Function file_fPath(xPath)
  Dim fPath : fPath = xPath
  fPath = Replace(fPath,"///","/")
  fPath = Replace(fPath,"//","/")
  If instr(fPath,":")<=0 Then 
    fPath = Server.MapPath(fPath)
  End If
 file_fPath = fPath
End Function

Function fil_exist(xPath)	
  Dim fp,FSO
  fp = file_fPath(xPath)
  SET FSO = Server.CreateObject(fsoCalss)
  fil_exist = FSO.FileExists(fp) 
  Set FSO = Nothing
End Function
	
Function fil_add(xPath,xCont)	 
  Dim fp,FSO,FIL
  fp = file_fPath(xPath)
  SET FSO = Server.CreateObject(fsoCalss)
  SET FIL = FSO.CreateTextFile(fp,True,False)
    FIL.WRITE(xCont)
  Set FIL = Nothing
  Set FSO = Nothing
End Function

Function fil_app(xPath,xCont)	 
  Dim fp,fs
  Set fs = Server.CreateObject(fsoCalss)
  Set fp = fs.OpenTextFile(Server.MapPath(xPath),8,True)
  fp.WriteLine xCont
  fp.Close
  Set fp = Nothing
  Set fs = Nothing
End Function

Function File_Add2(xPName2,xCont,xCSet)
  Dim st,xPName 
  xPName = file_fPath(xPName2)
  Set st = Server.CreateObject(adsCalss)   
    st.Type = 2   
    st.Mode = 3   
    st.Charset = xCSet  
    st.Open()   
    st.WriteText xCont   
    st.SaveToFile xPName,2   
    st.Close()  
  Set st = Nothing   
End Function

Function File_Read(xPName,xCSet)
  Dim str,stm
  if fil_exist(xPName) then
    set stm = server.CreateObject(adsCalss)
      stm.Type = 2 
      stm.Mode = 3 
      stm.Charset = xCSet
      stm.Open
      stm.loadfromfile file_fPath(xPName) 
      str = stm.ReadText
      stm.Close
    set stm = nothing
  else
    str = "(File Read Error!)"
  end if
  File_Read = str
End Function

Function fil_del(xPath) 
  Dim fp,FSO,FIL
  if fil_exist(xPath) then
    'Response.Write "<br>1"&xPath&xnam
    fp = file_fPath(xPath)
    SET FSO = Server.CreateObject(fsoCalss)
      FSO.DeleteFile fp,TRUE
    Set FSO = Nothing
  else
    'Response.Write "<br>2"&xnam
  end if
End Function

Function fil_copy(xfila,xfilb)	
  Dim fp,FSO,FIL,obj,org
    Set fso = Server.CreateObject(fsoCalss)	
      obj = file_fPath(xfilb)
      org = file_fPath(xfila)
	  'Response.Write obj&"<br>"&org
	  fso.CopyFile org,obj,True
    Set FSO = Nothing	
End Function
	
Function fil_move(xfila,xfilb)	
  Dim fp,FSO,FIL,obj,org
    Set fso = Server.CreateObject(fsoCalss)	
      obj = file_fPath(xfilb)
      org = file_fPath(xfila)
	  fso.MoveFile org,obj
    Set FSO = Nothing
End Function
	
Function files_move(xPath1,xPath2)
  dim fs,oFold,f,fName
  set fs = Server.CreateObject(fsoCalss)
  Set oFold = fs.GetFolder(file_fPath(xPath1))
  For Each f In oFold.Files
    fName = f.Name
    Call fil_move(xPath1&fName,xPath2&fName)
  Next
  set oFold = Nothing
  set fs = nothing
End Function
'Call files_move("xxx","yyy")

Function fold_exist(xpath,xnm)	
  Dim FSO,fp
  fp = file_fPath(xpath&"/"&xnm)
  SET FSO = Server.CreateObject(fsoCalss)
	fold_exist = FSO.FolderExists(fp)
  Set FSO = Nothing
End Function

Function fold_Size(xPath)	
  Dim fp,fs
  fp = xPath :If fp="" Then fp = "./"
  SET objFSO = Server.CreateObject(fsoCalss)
  SET objFolder = objFSO.GetFolder(file_fPath(fp))
    fs = objFolder.Size
  SET objFolder = Nothing
  SET objFSO = Nothing
  fold_Size = fs
End Function 

Function fold_add(xpath,xnm)	
  Dim fp,FSO,fold
  if not fold_exist(xpath,xnm) then
    SET FSO = Server.CreateObject(fsoCalss)
	  fp = FSO.BuildPath(file_fPath(xpath&"/"),xnm)
    SET fold = FSO.CreateFolder(fp) 
  end if		
  Set FSO = Nothing
  Set fold = Nothing
End Function  

Function fold_del(xpath)	
  Dim fp,FSO
  SET FSO = Server.CreateObject(fsoCalss)
  fp = file_fPath(xpath)
  If FSO.FolderExists(fp) Then
    FSO.DeleteFolder(fp)
  End If
  Set FSO = Nothing
End Function	

Function fold_copy(xfila,xfilb)	
  Dim fp,FSO,FIL,obj,org
  If fold_exist(xfila,"") Then
    Set fso = Server.CreateObject(fsoCalss)	
      obj = file_fPath(xfilb)
      org = file_fPath(xfila)
	  fso.copyFolder org,obj,True
    Set FSO = Nothing
  End If	
End Function

Function fil_read(xnam,xnline)	
  Dim fp,FSO,tso,fil_str
  SET FSO = Server.CreateObject(fsoCalss) 
  SET tso = FSO.OpenTextFile(file_fPath(xnam),1)
  fil_str = ""
  IF xnline = -1 THEN
    DO UNTIL tso.AtEndOfStream
      fil_str = fil_str&tso.ReadLine
    LOOP
  ELSE
    DO UNTIL tso.AtEndOfStream
      IF tso.Line = xnline THEN
        fil_str = fil_str&tso.ReadLine
      END IF
      tso.SkipLine
    Loop		
  END IF
  tso.Close
  Set FSO = Nothing
  fil_read = fil_str
End Function

Function chk_file(xFile,xPath,xType,xSize)
  Dim fName,rMsg,fExt
  fName = UCase(xFile.FileName)
  fExt = Mid(fName,InStrRev(fName,"."),12)
  rMsg = "OK"
  'Response.Write lCase(fATyps)&"-"&lCase(fExt)
  'Response.End()
  If inStr(lCase(fATyps),lCase(fExt))<=0 Then
    rMsg = "Type"
  End If
  If fil_exist(xPath) Then
    rMsg = "Exist"
  End If
  If inStr(lCase(xType),lCase(fExt))<=0 Then 'NOT (InStr(xType,right(fName,3))>0)
    rMsg = "Type"
  End If
  If xFile.FileSize > xSize*1024 Then
    rMsg = "Size"
  End If
  chk_file = rMsg
End Function


%>
