
<%

Function GetMemInfo(xUser)
Dim msql,s,MemID,MemName,MemTyp2,MemCard2,rs2
 msql = "SELECT MemID,MemName,MemTyp2,MemCard FROM Member_ABCDE WHERE MemID='"&xUser&"'"
 Set rs2=Server.CreateObject("Adodb.Recordset")
 rs2.Open msql,conn,1,1  ',MemTyp2,MemCard2
 Do While NOT rs2.EOF
   MemID = rs2("MemName")
   MemName = rs2("MemName")&""
   MemTyp2 = rs2("MemTyp2")&""
   MemCard = rs2("MemCard")&""
 rs2.MoveNext
 Loop
 rs2.Close()
 Set rs2=Nothing
 s = "帐号:"&MemID
 If MemName<>"" AND MemName<>MemID Then s = s& "<br>姓名:"&MemName&""
 If MemTyp2<>"" Then s = s& "<br>专业:"&MemTyp2&""
 If MemCard<>"" Then s = s& "<br>证件:"&MemCard&""
 GetMemInfo = s
End Function 

Function GetCurrOS()
'Windows XP    Windows NT 5.1  
'Windows 2003  Windows NT 5.2  
'Windows Vista Windows NT 6.0  Longhorn
'Windows 2008  Windows NT 6.0  
'Windows 7     Windows NT 6.1  
'Windows 8     Windows NT 7.0 
  dim temp,strUA 
  strUA = Request.ServerVariables("HTTP_USER_AGENT")
  strUA = lcase(strUA) 
  if Instr(strUA,"windows")>0 then temp="Windows" 
  'if Instr(strUA,"windows ce")>0 then temp="Windows CE" 
  'if Instr(strUA,"windows 95")>0 then temp="Windows 95" 
  'if Instr(strUA,"win98")>0 then temp="Windows 98" 
  'if Instr(strUA,"windows 98")>0 then temp="Windows 98" 
  if Instr(strUA,"windows 2000")>0 then temp="Windows 2000" 
  if Instr(strUA,"windows xp")>0 then temp="Windows XP" 
  if Instr(strUA,"windows nt")>0 then temp="Windows NT" 
  if Instr(strUA,"windows nt 5.0")>0 then temp="Windows 2000" 
  if Instr(strUA,"windows nt 5.1")>0 then temp="Windows XP" 
  if Instr(strUA,"windows nt 5.2")>0 then temp="Windows 2003" 
  if Instr(strUA,"windows nt 6.0")>0 then temp="Windows Vista" 
  if Instr(strUA,"windows nt 6.1")>0 then temp="Windows 7" 
  'if Instr(strUA,"windows nt 7.0")>0 then temp="Windows 8" 
  if Instr(strUA,"x11")>0 or Instr(strUA,"unix")>0 then temp="Unix" 
  if Instr(strUA,"sunos")>0 or Instr(strUA,"sun os")>0 then temp="SUN OS" 
  if Instr(strUA,"powerpc")>0 or Instr(strUA,"ppc")>0 then temp="PowerPC" 
  if Instr(strUA,"macintosh")>0 then temp="Mac" 
  if Instr(strUA,"mac osx")>0 then temp="MacOSX" 
  if Instr(strUA,"freebsd")>0 then temp="FreeBSD" 
  if Instr(strUA,"linux")>0 then temp="Linux" 
  if Instr(strUA,"palmsource")>0 or Instr(strUA,"palmos")>0 then temp="PalmOS" 
  if Instr(strUA,"wap")>0 then temp="WAP" 
  GetCurrOS = temp 
End Function 
'Response.Write GetCurrOS()&";"

Function GetCurrBrs()
Dim sAgent,sBrows,sVer,sTemp
 sAgent=uCase(Request.ServerVariables("HTTP_USER_AGENT"))
 'Response.Write sAgent
 'sAgent=Split(sAgent,";")
 If InStr(sAgent,"MSIE")>0 Then
   'sBrows="MSIE "
   sVer=Mid(sAgent,inStr(sAgent,"MSIE"),8) 'Trim(Left(Replace(sAgent,"MSIE",""),6))
 ElseIf InStr(sAgent,"FIREFOX")>0 Then 
   sVer="Firefox"
 ElseIf InStr(sAgent,"CHROME")>0 Then 
   sVer="Chrome" 
 ElseIf InStr(sAgent,"SAFARI")>0 Then 
   sVer="Safari"
 ElseIf InStr(sAgent,"Netscape")>0 Then 
   sBrows="Netscape "
   sTemp=Split(sAgent,"/")
   sVer=sTemp(UBound(sTemp))
 ElseIf InStr(sAgent,"rv:")>0 Then
   sBrows="Mozilla "
   sTemp=Split(sAgent," ")
   sVer=sTemp(UBound(sTemp))
 End If
 GetCurrBrs = sBrows&" "&sVer
End Function 
'Response.Write GetCurrBrs()
%>
