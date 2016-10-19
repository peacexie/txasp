<%

Function SysDBSwap(xKey,xCol,xTab,xWhr,xStr1,xStr2)
  Dim sql,fRS,j,iKey,iVal
  sql = " SELECT "&xKey&","&xCol&" FROM ["&xTab&"] "&xWhr
  rs.Open sql,conn,1,1 
  If rs.Eof Or rs.Bof Then
    fRS = "NG" 
  Else
    aRS = Rs.GetRows
	fRS = "OK"
  End If
  rs.Close() 
  
  if fRS = "OK" then 
  For j=0 To Ubound(aRS,2)
    iKey = aRS(0,j)
	iVal = aRS(1,j)
	iVal = Replace(iVal,xStr1,xStr2)
	iVal = Replace(iVal,"'","''")
	Call rs_DoSql(conn,"UPDATE "&xTab&" SET "&xCol&"='"&iVal&"' WHERE "&xKey&"='"&iKey&"'")
  Next
  End If
  SysDBSwap = Ubound(aRS,2)+1
End Function 


Function SysConfig()
  Dim s,sql,sDim,OldFlag,ParFlag : s=""
  sql = "SELECT * FROM [AdmPara] "
  sql = sql & " WHERE ParFlag IN ('ParConfig','ParTYN','ParText','ParNum') "
  sql = sql & " ORDER BY ParFlag,ParName,ParCode "
  rs.Open sql,conn,1,1 
  Do While Not rs.EOF 
    ParFlag = rs("ParFlag")
	ParName = rs("ParName")
    ParCode = rs("ParCode")
    ParText = rs("ParText")
    ParNum = rs("ParNum")
	If OldFlag<>ParFlag Then 
	  sDim = sDim&vbcrlf&"Dim "&ParCode
	  s = s&vbcrlf
	Else
	  sDim = sDim&", "&ParCode
	End If
    'If ParFlag="ParText" Then ParCode = Mid(ParCode,2)
    If ParFlag="ParNum" Then
      s = s&vbcrlf&ParCode&" = "&RequestSafe(ParNum,"N","0")&"    '"&ParName&""
    Else
      s = s&vbcrlf&ParCode&" = "&Chr(34)&ParText&Chr(34)&"    '"&ParName&""
    End If
	OldFlag = ParFlag
  rs.MoveNext
  Loop
  rs.Close()
  s = vbcrlf&"'    Peace Para: [Config:  配置 | 系统 | 数字 | 开关]"&"   "&s 
  SysConfig = sDim&vbcrlf&s
End Function 

Function SysParas()
  Dim sql,k,s,sDim
  pfSql = Split("ParEditor,ParSMS,ParEmail,ParWMark,ParDate,ParLink",",") 'ParForm,
  pfName = Split("editor,  sms,   jmail,   wmark,   date,   link",",") 'form,  
  For k=0 To uBound(pfName)
	s = "" :sDim = ""
	sql = "SELECT * FROM [AdmPara] "
	sql = sql & " WHERE ParFlag='"&pfSql(k)&"' "
	sql = sql & " ORDER BY ParName,ParCode "
	rs.Open sql,conn,1,1 
	Do While Not rs.EOF 
	ParFlag = rs("ParFlag")
	ParName = rs("ParName")
	ParCode = rs("ParCode")
	ParText = rs("ParText")
	ParDate = rs("ParDate")
	sDim = sDim&ParCode&", "
	  If ParFlag="ParDate" Then
		s = s&vbcrlf&ParCode&" = "&Chr(34)&RequestSafe(ParDate,"D","1900-12-31")&Chr(34)&"    '"&ParName&"" 
	  Else
		s = s&vbcrlf&ParCode&" = "&Chr(34)&ParText&Chr(34)&"    '"&ParName&""
	  End If
	rs.MoveNext
	Loop
	rs.Close()
	s = "Dim "&Replace(sDim&"|",", |","")&""&s
	s = "'    Peace Para: ["&Trim(pfName(k))&"]"&"   "&vbcrlf&s
	s = "<"&"%"&vbcrlf&s&vbcrlf&"%"&">"
	pf = "../../upfile/sys/pcfg/"&Trim(pfName(k))&".asp"
	Call File_Add2(pf,s,"UTF-8")
	Response.Write vbcrlf&" "&pf&"; "
  Next
  'SysParas = s
End Function 
  
%>

