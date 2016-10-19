
<%
Function uPag_JXML()

 sql = " SELECT TOP 24 * FROM [JobPos] "
 sql =sql& " WHERE 1=1 AND SetSAdm='Y' "
 sql =sql& " ORDER BY KeyID DESC" 
   rs.Open Sql,conn,1,1
   s=""
   Do While NOT rs.EOF
KeyID = rs("KeyID")
KeyUser = rs("KeyUser")
KeyCode = rs("KeyCode")
KeyMod = rs("KeyMod")
KeyFlag = rs("KeyFlag")&""
InfType = Show_Text(rs("InfType"))
InfName = Show_Text(rs("InfName"))
InfSex = rs("InfSex")
InfAge = rs("InfAge")
InfWed = rs("InfWed")
InfGrade = Show_Text(rs("InfGrade"))
InfFrom = Show_Text(rs("InfFrom"))
InfPay = Show_Text(rs("InfPay"))
InfPElse = Show_Text(rs("InfPElse"))
InfWork = Show_Text(rs("InfWork"))
InfRemark = Show_Text(rs("InfRemark"))
LnkName = Show_Text(rs("LnkName"))
LnkTel = Show_Text(rs("LnkTel"))
LnkMobile = Show_Text(rs("LnkMobile"))
LnkEmail = Show_Text(rs("LnkEmail"))
LnkAddr = Show_Text(rs("LnkAddr"))
LnkURL = Show_Text(rs("LnkURL"))
LnkElse = Show_Text(rs("LnkElse"))
LogATime = rs("LogATime")
LogETime = rs("LogETime")

  companyName = ""
If inStr(KeyFlag,"C")>0 Then
  companyName = rs_Val(conn,"SELECT InfName FROM ComUser WHERE KeyCode='"&KeyFlag&"'","InfName")
ElseIf inStr(KeyFlag,"H")>0 Then
  companyName = rs_Val(conn,"SELECT InfName FROM HomUser WHERE KeyCode='"&KeyFlag&"'","InfName")
End If

companyInfo = " "
workerNum = " "
fax = " "
post = " "

If Len(companyName)>0 And Len(LnkEmail&LnkTel&LnkMobile)>0 Then
s=s&vbcrlf&"<position>"
s=s&vbcrlf&"		<postionNO><![CDATA["&  KeyCode &"]]></postionNO>"
s=s&vbcrlf&"		<positionName><![CDATA["&  InfName &"]]></positionName>"
s=s&vbcrlf&"		<time><![CDATA["&  LogATime &"]]></time>"
s=s&vbcrlf&"		<area><![CDATA["&  InfFrom &"]]></area>"
s=s&vbcrlf&"		<companyName><![CDATA["&  companyName &"]]></companyName>"
s=s&vbcrlf&"		<positionInfo><![CDATA["&  InfWork &"]]></positionInfo>"
s=s&vbcrlf&"		<companyInfo><![CDATA["&  companyInfo &"]]></companyInfo>" '公司介绍
s=s&vbcrlf&"		<email><![CDATA["&  LnkEmail &"]]></email>"
s=s&vbcrlf&"		<url><![CDATA[http://www.txjia.com/job/pos.asp?CD="&  KeyCode &"]]></url>"
s=s&vbcrlf&"		<pay><![CDATA["&  InfPay &"]]></pay>"
s=s&vbcrlf&"		<degree><![CDATA["&  InfGrade &"]]></degree>"
s=s&vbcrlf&"		<jobAge><![CDATA["&  InfAge &"]]></jobAge>"
s=s&vbcrlf&"		<jobType><![CDATA["&  InfType &"]]></jobType>"
s=s&vbcrlf&"		<workerNum><![CDATA["&  workerNum &"]]></workerNum>" '公司规模
s=s&vbcrlf&"		<linkman><![CDATA["&  LnkName &"]]></linkman>"
s=s&vbcrlf&"		<phone><![CDATA["&  LnkTel &"]]></phone>"
s=s&vbcrlf&"		<fax><![CDATA["&  fax &"]]></fax>" '传真
s=s&vbcrlf&"		<address><![CDATA["&  LnkAddr &"]]></address>"
s=s&vbcrlf&"		<post><![CDATA["&  post &"]]></post>" '邮政编码
s=s&vbcrlf&"		<homePage><![CDATA["&  LnkURL &"]]></homePage>"
s=s&vbcrlf&"</position>"
End If

   rs.MoveNext
   Loop
   rs.Close()

s1=vbcrlf&"<?xml version=""1.0"" encoding=""GB2312""?>"
s1=s1&vbcrlf&"<jobs>"
s2=s2&vbcrlf&"</jobs>"
s=s1&s&s2 'http://www.jobui.com/dlzp.html

    set fil1 = fso.CreateTextFile(Server.MapPath("../../script/job/job.xml"),True)
    fil1.Write(s)
    fil1.Close
	'Response.Write vbcrlf&"<hr>"&s&""  

uPag_JXML = s
End Function  

%>

