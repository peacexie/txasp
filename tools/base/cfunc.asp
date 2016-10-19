<%

Function dbMy_Fill(xVal,xType)
  If xType="N" Then
   If Not isNumeric(xVal) Then
     xVal = "0"
   End If
  ElseIf xType="D" Then
   If Not isDate(xVal) Then
     xVal = "1900-12-31"
   End If
  Else
   xVal = Replace(xVal,"\","\\")
   xVal = Replace(xVal,vbcrlf,"\r\n")
   'xVal = Replace(xVal,vbcr,"\r")
   'xVal = Replace(xVal,vblf,"\n")
   xVal = Replace(xVal,"'","\'")
   xVal = Replace(xVal,"""","\""")
  End If
  dbMy_Fill = xVal
End Function

Function dbMy_TabData(xTab)
  Dim a(255),b(255),rs,s,c,i,j,fName,fType,fSize,fKey,sVal
  Set rs = Server.Createobject("Adodb.Recordset")
  rs.Open "SELECT * FROM ["&xTab&"]",conDB,1,1
  
	  s = "" : c = "" 
	  s = s&"            INSERT INTO `"&lCase(xTab)&"` (c) VALUES "
  FOR i = 0 TO rs.Fields.Count-1
    fName = rs.Fields(i).Name
    fType = rs.Fields(i).Type
    fSize = rs.Fields(i).DefinedSize
	b(i) = fName
	a(i) = fType
	c = c&"`"&fName&"`, "
  NEXT
    c = Left(c,Len(c)-2) 
	s = Replace(s,"(c)","("&c&")")
  
  Do While NOT rs.EOF
	c = vbcrlf&"("
	for j=0 to rs.Fields.Count-1
	   sVal = rs(b(j))&""
	   If Int(a(j))<7 Then
	    c = c&""&dbMy_Fill(sVal,"N")&", "
	   ElseIf Int(a(j))=7 Then
	    c = c&"'"&dbMy_Fill(sVal,"D")&"', "
	   Else
	    c = c&"'"&dbMy_Fill(sVal,"C")&"', "
	   End If
	next 
	c = Left(c,Len(c)-2)&") "
	s = s&c
    rs.MoveNext()
  Loop
  rs.close()
  
  Set rs = Nothing
  dbMy_TabData = s
End Function

Function dbMy_TabStru(xTab)
  Dim rs,s,i,fName,fType,fSize,fKey
  Set rs = Server.Createobject("Adodb.Recordset")
  rs.Open " SELECT * FROM ["&xTab&"] ",conDB,1,1
	  s = "" 's = s&vbcrlf&""
	  s = s&vbcrlf&"DROP TABLE IF EXISTS `"&xTab&"`; "
	  s = s&vbcrlf&"CREATE TABLE IF NOT EXISTS `"&xTab&"` ( "
  FOR i = 0 TO rs.Fields.Count-1
    fName = rs.Fields(i).Name
    fType = rs.Fields(i).Type
    fSize = rs.Fields(i).DefinedSize
	  s = s&"`"&fName&"` "
	If cStr(fType="2") Then
	  s = s&"smallint(6) default '0', "
	End If
	If cStr(fType="3") Then
	  s = s&"int(11) default '0', "
	End If
	If cStr(fType="4") Then
	  s = s&"float default '0', "
	End If
	If cStr(fType="7") Then
	  s = s&"datetime default NULL, "
	End If
	If cStr(fType="202") Then
	 If i=0 Then
	  s = s&"varchar("&fSize&") NOT NULL, "
	 Else
	  s = s&"varchar("&fSize&") default NULL, " 
	 End If
	End If
	If cStr(fType="203") Then
	  s = s&"longtext, "
	End If
    If i=0 Then
      fKey = fName
    End If
  NEXT
      s = s&"PRIMARY KEY  (`"&fKey&"`) "
	  s = s&")ENGINE=MyISAM DEFAULT CHARSET=utf8; "
  rs.Close()
  Set rs = Nothing
  dbMy_TabStru = s
End Function

Function clr_fList(xPath2)
    dim fs, folder, file, item, url, iPath, iName
	Dim xPath : xPath=xPath2 
    set fs = CreateObject("Scripting.FileSystemObject")
	If instr(xPath,":")<=0 Then xPath=Server.MapPath(xPath) 
    set folder = fs.GetFolder(xPath)
    for each item in folder.SubFolders
        clr_fList(item.Path)
    next   
    Dim fExt,aPath,i,fPath
	for each item in folder.Files
	  iPath = item.Path 
	  fExt = lCase(Mid(iPath,InStrRev(iPath,"."),8)) 
	  If FileIgno="" Then
	    fPath = "Show"
	  Else
	    fPath = "Show"
		aPath = Split(FileIgno,"|")
	    For i=0 To uBound(aPath)
		  If inStr(lCase(iPath),aPath(i)) Then
		    fPath = "Hidd"
			Exit For
		  End If
		Next
	  End If
		If inStr(FileType,fExt)>0 And fPath="Show" Then
		  If item.Size>210123 Then '204800
		    Response.Write VBCRLF&"<br><font color=blue>&middot;"&Replace(iPath,FilePRep,"")&"</font>"
		  Else
		    FileLStr = FileLStr&iPath&"|"
		    FileLTim = FileLTim&item.DateLastModified&" ("&item.Size&")|"
		    Response.Write VBCRLF&"<br>&middot;"&Replace(iPath,FilePRep,"")
		  End If
		End If
    next 
End Function

Function clr_Numb(xVal,xDef)
  If not isNumeric(xVal) then 
    Chk_Numb = xDef 
  Else 
    Chk_Numb = xVal
  End if 
End Function

Function clr_Reset(xPara,xDef) 
  tVal = Request(xPara)
  If tVal<>"" Then
    clr_Reset = tVal
	Session(xPara) = tVal
  Else
    tVal = Session(xPara)&""
    If tVal<>"" Then
      clr_Reset = tVal
    Else
      clr_Reset = xDef
    End If
  End If
End Function 

Function clr_Text(xText) 
   xText = xText&""
   xText = Replace(xText, "<", "&lt;")
   xText = Replace(xText, ">", "&gt;")
   clr_Text = xText
End Function 

%>
