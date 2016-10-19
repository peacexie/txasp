<!--#include file="config.asp"-->
<!--#include file="../admin/sys.funcs.asp"-->
<!--#include file="../func2/cch_Class.asp"-->
<head>
<meta http-equiv="Pragma" content="no-cache">
<meta http-equiv="Expires" content="0">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
</head>
<body>

<%
tx01 = Timer()
SET rs=Server.CreateObject("Adodb.Recordset") 



Response.Write vbcrlf&"<br> ------ Config:  配置 | 系统 | 数字 | 开关"

s = SysConfig()
s="<"&"%" &s&vbcrlf&vbcrlf& "%"&">" ':Response.Write sDim
pf="../../upfile/sys/config/Config.asp"
Call File_Add2(pf,s,"UTF-8")
Response.Write vbcrlf&" "&pf&"; "



Response.Write vbcrlf&"<br> ------ pfName(n):  Jmail | Editor | SMS | 水印 | 日期 | 联系"

Call SysParas()




Response.Write vbcrlf&"<br> ------ fa(3):"

Dim fa(3) : fa(0)="Groups" : fa(1)="Para" : fa(2)="Module"
For i=0 To 2 

 sqlKi = ""
 If fa(i)="Module" Then
 sqlKi = " AND SysTop<'U' "
 End If
 sql = " SELECT * FROM [AdmSyst] WHERE SysType='"&fa(i)&"' "&sqlKi&" ORDER BY SysTop,SysID " 
 s=""

 rs.Open Sql,conn,1,1
  SysTop = "0"
 Do While NOT rs.EOF
  SysID = rs("SysID")
  SysName = Trim(rs("SysName")&"")
  If Mid(SysName,3,1)="与" Then SysName=Replace(SysName,"与","")
  SysType = rs("SysType")
     'If rs("SysTop")>"9" And SysTop<="9" And SysType="Para" Then
	 If (rs("SysTop")="e" OR rs("SysTop")="h") And SysType="Para" Then
       s=s& "<br>"
     End If
     SysTop = rs("SysTop")

   If SysType="Groups" Or SysType="Module" Then
	If SysType="Groups" Then
	 SysID = Mid(SysID,2)
	 iUrl = "?ModID="&SysID
	Else
	 SysID = Mid(SysID,4)
	 iUrl = "?ModID="&SysID
	End If

   ElseIf inStr(SysName,".yno")>0 Then
     SysName = Replace(SysName,".yno","")
	 iUrl = "para_yno.asp?ModID="&SysID
   ElseIf inStr(SysName,".int")>0 Then
     SysName = Replace(SysName,".int","")
	 iUrl = "para_int.asp?ModID="&SysID
   ElseIf inStr(SysName,".date")>0 Then
     SysName = Replace(SysName,".date","")
	 iUrl = "para_date.asp?ModID="&SysID
   ElseIf inStr(SysName,".rem")>0 Then
     SysName = Replace(SysName,".rem","")
	 iUrl = "para_rem.asp?ModID="&SysID
   Else
	 iUrl = "para_text.asp?ModID="&SysID
   End If 
   If SysType<>"Groups" Then 
     SysName = Replace(SysName,"参数","")
   End If
   
   s=s& vbcrlf&"<a href="&iUrl&">"&SysName&"</a> | "
   rs.MoveNext()
 Loop
 rs.close()
 
 pf = "../../upfile/sys/config/sf_"&fa(i)&".htm"
 Call File_Add2(pf,s,"UTF-8")
 Response.Write vbcrlf&""&fa(i)&"; "

Next



Response.Write vbcrlf&"<br> ------ ParFlag IN('ParRem','ParBBS','ParTemp'): "

sql = "SELECT "
sql = sql & " ParName "
sql = sql & ",ParCode "
sql = sql & ",ParFlag "
sql = sql & ",ParRem "
sql = sql & " FROM [AdmPara] "
sql = sql & " WHERE ParFlag IN('ParRem','ParBBS','ParTemp') AND ParCode LIKE '%.%' "
sql = sql & " ORDER BY ParCode "

rs.Open sql,conn,1,1 
Do While Not rs.EOF 
ParName = rs("ParName")
ParCode = rs("ParCode")
ParRem = rs("ParRem")

dtCont = ParRem

ParCode2 = LCase(ParCode)
ParCode3 = Right(ParCode2,3)
if ParCode3 = "asp" then
  dtCont = "'    Peace Para: ["&ParCode&"]"&ParName&"   "&vbcrlf&dtCont
  dtCont = "<"&"%"&vbcrlf&dtCont&vbcrlf&"%"&">"
elseif ParCode3 = "htm" then
  If inStr(dtCont,"<head>")<=0 Then
  dtCont = vbcrlf&"<!-- Peace Para: ["&ParCode&"]"&ParName&"-->"&vbcrlf&dtCont
  End If
elseif ParCode3 = ".js" then
  dtCont = vbcrlf&"/*   Peace Para: ["&ParCode&"]"&ParName&"*/ "&vbcrlf&dtCont
elseif ParCode3 = "txt" then
  dtCont = dtCont
else
  dtCont = vbcrlf&"/*   Peace Para: ["&ParCode&"]"&ParName&"*/ "&vbcrlf&dtCont
end if

pf = "../../upfile/sys/para/"&ParCode2&""
Call File_Add2(pf,dtCont,"UTF-8")
Response.Write vbcrlf&" "&pf&"; "
rs.MoveNext
Loop
rs.Close()




Response.Write vbcrlf&"<br> ------ ParFil: "

sql = "SELECT "
sql = sql & " ParName "
sql = sql & ",ParCode "
sql = sql & ",ParFlag "
sql = sql & ",ParText "
sql = sql & ",ParRem "
sql = sql & " FROM [AdmPara] "
sql = sql & " WHERE 1=1 AND ParFlag='ParFil' "
sql = sql & " ORDER BY ParCode "

rs.Open sql,conn,1,1 
s="" : s2="" : s3=""  :sDim = ""
Do While Not rs.EOF 
ParName = rs("ParName")
ParCode = rs("ParCode")
ParText = rs("ParText")
ParRem = rs("ParRem")
ParRem = Replace(ParRem,vbcrlf,"")
ParRem = Replace(ParRem,vbcr,"")
ParRem = Replace(ParRem,vblf,"")
ParRem = Replace(ParRem,chr(34),"")
sDim = sDim&ParCode&","
s = s& vbcrlf&""&ParCode&" = """&ParRem&""""
s2 = s2 & "&"&ParCode&""
s3 = s3& vbcrlf&"<br><br><li><b>"&ParName&"</b><br> 【 "&ParRem&"】</li>"
rs.MoveNext
Loop
rs.Close()

s2 = Mid(s2,2)
s = "Dim "&sDim&"ParFilALLKeys"&s
s = s&vbcrlf&"ParFilALLKeys = "&s2
s = "<"&"%"&vbcrlf&s&vbcrlf&"%"&">"
pf = "../../upfile/sys/para/keywords.asp"
Call File_Add2(pf,s,"UTF-8")
Response.Write vbcrlf&" "&pf&"; "

pf = "../../upfile/sys/para/kwords2.asp"
a = "<"&"%" &vbcrlf& "If Session([MemID])&[]&Session([UsrID])<>[]&Session([InnID])<>[] Then" &vbcrlf& "%"&">"
a = Replace(a,"[",chr(34))
a = Replace(a,"]",chr(34))
b = "<"&"%"&vbcrlf&"End If"&vbcrlf&"%"&">"
s3 = a&s3&b
Call File_Add2(pf,s3,"UTF-8")
Response.Write vbcrlf&" "&pf&"; "



Response.Write vbcrlf&"<br> ------ ShowMenu('$ModID$'): "

s0=vbcrlf
s0=s0&vbcrlf&"<div>"
s0=s0&vbcrlf&"  <ul class='left_top' onClick=""ShowMenu('$ModID$')"" >"
s0=s0&vbcrlf&"    <li class='left_top_left left'>$ModName$</li>"
s0=s0&vbcrlf&"    <li class='left_top_right right'></li>"
s0=s0&vbcrlf&"  </ul>"
s0=s0&vbcrlf&"  <ul class='left_conten col999' id='sSub$ModID$'>"
s0=s0&vbcrlf&"$ModItem$"
s0=s0&vbcrlf&"  </ul>"
s0=s0&vbcrlf&"  <div class='left_end'></div>"
s0=s0&vbcrlf&"</div>"

sql = "SELECT SysID,[SysName] FROM [AdmSyst] WHERE SysType='Module' And SysTop<'U' ORDER BY SysTop,SysID "
rs.Open sql,conn,1,1 
s1="" : s2="" : s3="" :mStr=""
Do While Not rs.EOF 
  SysName = rs("SysName")
  SysID = rs("SysID")
  mStr = mStr&";"&SysID
  ModItem = rs_Val("","SELECT ParRem FROM AdmPara WHERE ParCode='rM"&Mid(SysID,4)&"'")
  s1=s0
  s1=Replace(s1,"$ModID$",SysID)
  s1=Replace(s1,"$ModName$",SysName)
  s1=Replace(s1,"$ModItem$",ModItem)
  s2=s2&s1
rs.MoveNext
Loop
rs.Close()

s2=s2&vbcrlf&"<script type='text/javascript'>var mStr='"&mStr&"';</script>"
pf = "../../upfile/sys/config/adm_menu.config"
Call File_Add2(pf,s2,"UTF-8")
Response.Write vbcrlf&" "&pf&"; "


Response.Write vbcrlf&"<br> ------ ParGroup: "
sql = "SELECT * FROM AdmPara WHERE ParFlag='ParGroup' AND ParName<'9' ORDER BY ParName, ParCode "
rs.Open sql,conn,1,1 
s1="" : s2="var nGroup = (n); "&vbcrlf&"var aGroup = new Array();" : i=0
Do While Not rs.EOF 
  pCode = rs("ParCode")
  pName = rs("ParName") : pNFlg = Mid(pName,1,2)
  pText = rs("ParText") : pName = Mid(pName,3)
  
  If(pNFlg<"20") Then
      s1 =s1&vbcrlf& " <a id='ModGroup"&i&"' onClick='ShowGroup("&i&")' class='ModGroupB'>"&pName&"</a>"
      s2 =s2&vbcrlf& " aGroup["&i&"] = '"&pText&"';"
	  i=i+1 '//echo $pCode;
  Else
      s2 =s2&vbcrlf& "var "&pCode&" = '"&pText&"';"
  End If
rs.MoveNext
Loop
rs.Close()

pf = "../../upfile/sys/config/sf_Admin.htm"
Call File_Add2(pf,s1,"UTF-8")
Response.Write vbcrlf&" "&pf&"; "

s2 = Replace(s2,"(n)",i)
pf = "../../upfile/sys/config/sf_Admin.js"
Call File_Add2(pf,s2,"UTF-8")
Response.Write vbcrlf&" "&pf&"; "



Response.Write vbcrlf&"<br> ------ IP限制(FIP,FUrl,FKey): "
s=""
s=s&vbcrlf&vbcrlf&rs_Val("","SELECT ParRem FROM AdmPara WHERE ParCode='rSafeCfg.asp'")&" "
s=s&vbcrlf&vbcrlf&rs_Val("","SELECT ParRem FROM AdmPara WHERE ParCode='rSafeChr.asp'")&" "
s=s&vbcrlf&vbcrlf&rs_Val("","SELECT ParRem FROM AdmPara WHERE ParCode='rSafeIP.asp'")&" "
s=s&vbcrlf&vbcrlf&rs_Val("","SELECT ParRem FROM AdmPara WHERE ParCode='rSafeUrl.asp'")&" "

s = "<"&"%"&s&vbcrlf&"%"&">"
pf = "../../upfile/sys/para/keysafe.asp"
Call File_Add2(pf,s,"UTF-8")
Response.Write vbcrlf&" "&pf&"; "



Response.Write vbcrlf&"<br> ------ Depart: "

'///////////////////////////////////
sql45 = "SELECT SysID,SysName,SysNEng FROM AdmSyst WHERE SysType='Depart' ORDER BY SysTop, SysID "
rs.Open sql45,conn,1,1 
s1="" : s2="" : s4=""
iPeac = 0
Do While Not rs.EOF 
  a1 = rs("SysID")
  a2 = rs("SysName")
  a3 = rs("SysNEng")
  s1 = s1&a1&"|"
  s2 = s2&a2&"|"
  s3 = s3&a3&"|"

rs.MoveNext
Loop
rs.Close()

s1 = vbcrlf&"strDepCode = """&s1&"""" 
s2 = vbcrlf&"strDepName = """&s2&"""" 
s3 = vbcrlf&"strDepFlag = """&s3&"""" 
sDat = "<"&"%"&sVer&s1&s2&s3&vbcrlf&"%"&">"
pf = "../../upfile/sys/config/_depart.asp"
Call File_Add2(pf,sDat,"UTF-8")
Response.Write vbcrlf&" "&pf&"; "
'//////////////////////////////////



Response.Write vbcrlf&"<br> ------ Inner: "

sql45 = "SELECT SysID,SysName,SysNEng FROM AdmSyst WHERE SysType='Inner' ORDER BY SysTop, SysID "
rs.Open sql45,conn,1,1 
s1="" : s2="" : s4=""
iPeac = 0
Do While Not rs.EOF 
  a1 = rs("SysID")
  a2 = rs("SysName")
  'a3 = rs("SysNEng")
  s1 = s1&a1&"|"
  s2 = s2&a2&"|"
  's3 = s3&a3&"|"

rs.MoveNext
Loop
rs.Close()

s1 = vbcrlf&"strInnCode = """&s1&"""" 
s2 = vbcrlf&"strInnName = """&s2&"""" 
's3 = vbcrlf&"strDepFlag = """&s3&"""" 
sDat = "<"&"%"&sVer&s1&s2&vbcrlf&"%"&">"
pf = "../../upfile/sys/config/_inner.asp"
Call File_Add2(pf,sDat,"UTF-8")
Response.Write vbcrlf&" "&pf&"; "
'//////////////////////////////////



Response.Write vbcrlf&"<br> ------ ParSEO: "

s0 = vbcrlf
sql = "SELECT * FROM [AdmPara] WHERE ParFlag='ParSEO' ORDER BY ParCode "
rs.Open sql,conn,1,1 
Do While Not rs.EOF 
  ParCode = rs("ParCode")
  ParRem = rs("ParRem")
  ParRem = Replace(ParRem,"(","""")
  ParRem = Replace(ParRem,")","""")
  ParRem = Replace(ParRem,"=",ParCode&"=")
  s0 = s0&vbcrlf&ParRem&vbcrlf
rs.MoveNext
Loop
rs.Close()

s = "<"&"%"&s0&vbcrlf&"%"&">"
pf = "../../upfile/sys/para/seopara.asp"
Call File_Add2(pf,s,"UTF-8")
Response.Write vbcrlf&" "&pf&"; "


SET rs=Nothing



Response.Write vbcrlf&"<br> ------ End 刷新OK!!! <br>"&vbcrlf
Call cchClear()
tx01 = Timer()-tx01
Response.Write vbcrlf&"<br>"& tx01
%>

</body>
</html>
