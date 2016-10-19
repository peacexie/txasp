<%

Set rs=Server.Createobject("Adodb.Recordset")
s1="" : s2="" : s3="" : s4="" : s5=""
s="" : t=""
idTabs = "'Groups','Depart','Para','Module','Admin','Inner','System'"
'Trade



sql = " SELECT * FROM [AdmSyst] "
sql =sql& " WHERE SysType NOT IN("&idTabs&") " 
'sql =sql& " WHERE SysID IN(SELECT TypMod FROM WebTyps) "
'sql =sql& " AND SysTop < 'u' "
sql =sql& " ORDER BY SysType,SysTop,SysID " 
rs.Open Sql,conn,1,1
Do While NOT rs.EOF 
  SysID = rs("SysID")
  SysName = rs("SysName")
  s1=s1&SysID&"|"
  s2=s2&SysName&"|"
rs.MoveNext
Loop
rs.Close()
s=s&vbcrlf&"sysModCode="""&s1&""""
s=s&vbcrlf&"sysModName="""&s2&""""



aMod = Split(s1,"|")
'aName = Split(s2,"|")
For i=0 To uBound(aMod)-1
   iMod = aMod(i) :s=s&vbcrlf
   s1=vbcrlf&"s"&iMod&"Code="""
   s2=vbcrlf&"s"&iMod&"Lay ="""
   s3=vbcrlf&"s"&iMod&"Name="""
   s4=vbcrlf&"s"&iMod&"Nam2="""
   s5=vbcrlf&"s"&iMod&"Img ="""
   sql7 = "SELECT * FROM [WebTyps] WHERE TypMod='"&aMod(i)&"' ORDER BY TypMod,TypLayer"
   rs.Open sql7,conn,1,1 '3
   Do While NOT rs.EOF 
	 TypID = rs("TypID")
	 TypLayer = rs("TypLayer")
	 TypName = rs("TypName")
	 TypNam2 = rs("TypNam2")
	 TypResume = rs("TypResume")&""
	 ImgName = rs("ImgName")&""
	 s1=s1&TypID&"|"
	 s2=s2&TypLayer&"|"
	 s3=s3&TypName&"|"
	 s4=s4&TypNam2&"|"
	 s5=s5&ImgName&"|"
	 rs.Movenext
   Loop 
   rs.Close()
   s=s&s1&""""
   s=s&s2&""""
   s=s&s3&""""
   s=s&s4&""""
   s=s&s5&""""
Next



Set rs = Nothing

s = "<"&"%"&s&vbcrlf&"%"&">" :echo s
Call File_Add2("../../upfile/sys/config/MTypList.asp",s,"UTF-8")
Response.Write "<hr>更新成功！"
Response.End()
	
%>