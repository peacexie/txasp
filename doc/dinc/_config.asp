<!--#include file="../../inc/home/func3.asp"-->
<!--#include file="../../sadm/func1/func1.asp"-->
<!--#include file="../../sadm/func2/func2.asp"-->
<!--#include file="../../sadm/func1/func_opt.asp"-->
<!--#include file="../../sadm/func1/func_file.asp"-->
<!--#include file="../../sadm/func2/func_sfile.asp" -->

<!--#include file="../../upfile/sys/config/_inner.asp"-->
<!--#include file="../../upfile/sys/para/tmstyp2.asp" -->

<%

sysName = "(xx公司 内部公文)"
ModID = "DocD124"
ModTab = "DocsNews"
upRoot = Config_Path&"upfile/"
upPart = "dtdoc"
Session("upPart") = upPart

PrmFile = Get_fName()
If inStr(";index.asp;dow_view.asp;",PrmFile&";")>0 Then
 If Request("Act")="AdmLogin" Then
  Call Chk_Perm1("ModInnDocs","")
 Else
  Call Chk_Perm3(ModID,Config_Path&"doc/")
 End If
Else
  Call Chk_Perm3(ModID,Config_Path&"doc/")
End If

'////////////////////////////

Set rs=Server.CreateObject("Adodb.Recordset")

'////////////////////////////

Function GetGList(xType,xDef) 
Dim aCode,aName,aLay,i,s
 aCode = Split(strInnCode,"|")
 aName = Split(strInnName,"|")
 'aFlag = Split(strDepFlag,"|")
 For i=0 To uBound(aCode)-1
  If xType="Val" Then 
    If xDef=aCode(i) Then 
	  GetGList=aName(i)
	  Exit Function 'For
    End If
  Else
	''If (inStr(Session(UsrPStr),"("&aCode(i))>0 AND Len(aCode(i))>4) Or (inStr(Session(UsrPStr),"{Admin}")>0) Then
      fSel=" " : If aCode(i)=xDef Then fSel=" selected "
	  s=s&vbcrlf&"<option value="&aCode(i)&fSel&">"&aName(i)&"</option>"
    ''End If  
  End If
 Next
GetGList = s
End Function


Function GetUList(xPerm)
Dim rs,aCode,aName,SysID,SysName,sql,s,s2,i,j,k : i=0 : j=101 : s="" : k=0
'////////////////////////////////

 aCode = Split(strInnCode,"|")
 aName = Split(strInnName,"|")
 For k=0 To uBound(aCode)-1
  SysName = aName(k)
  SysID = aCode(k)
  bgCol = "FFFFFF" : If k Mod 2=1 Then bgCol="F0F0FF"
  s=s&vbcrlf&"<tr bgcolor='#"&bgCol&"' OnMouseOver='this.bgColor=""#FFFFCC""' onMouseOut='this.bgColor=""#"&bgCol&"""'><td align='left' width='90%'>"
  i=0 : s2="" : j=j+1
'////////////////////////////////
sql = "SELECT UsrID,UsrName FROM [AdmUser"&Adm_aUser&"] "
sql = sql& " WHERE UsrType IN('"&SysID&"') ORDER BY UsrType,UsrID"
Set rs=Server.Createobject("Adodb.Recordset")
rs.Open Sql,conn,1,1
Do While NOT rs.EOF 
  UsrName = rs("UsrName")
  UsrID = rs("UsrID")
  sc = "" :If inStr(xPerm,""&UsrID&";")>0 Then sc = " checked "
  'sd = " disabled='disabled' " :If Len(xPerm)>3 Then sd = ""
  s = s&vbcrlf&"<li class='mInnID2'><input name='InfView' type='checkbox' id='InfView' value='"&UsrID&";' "&sc&sd&" >"&UsrName&"</li>"
  s2 = s2 &""&UsrID&";"
  rs.MoveNext
  i = i + 1
Loop
rs.Close()
Set rs = Nothing
'////////////////////////////////
  Dim tmpName2
  tmpName2 = "s"&j&""
  tmpName2 = "InfVGrps"
  s=s&"</td><td align='center'><input name='"&tmpName2&"' type='checkbox' id='"&tmpName2&"' xvalue='"&s2&"' onClick=""ChkS99(this,'"&s2&"')""></td>"
  s=s& "<td nowrap>"&SysName&"</td></tr>"
 Next
'////////////////////////////////
s = vbcrlf&vbcrlf&""&s
s = s&""&vbcrlf&vbcrlf
GetUList = s
End Function

%>
