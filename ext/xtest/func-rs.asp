<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>rs_AutID Demo</title>
</head>
<body>

<!--#include file="../../page/_config.asp"-->

<%

'echo strr srepeat("2",3)

echo rs_OrdID(conn,"yymmdd-r3",6,3)
echo rs_OrdID(conn,"yymmdd-hnsx-s1r1",6,4)
echo "<br>"

echo rs_OrdID(conn,"yymmdd-hnsx-r2s2",6,2)
echo rs_OrdID(conn,"yymmdd-hnsx-r3",6,3)
echo rs_OrdID(conn,"yymd-hnsx-r2",4,2)

rID = Get_IDRnd()
eID = Get_IDEnc(rID,0)
Response.Write echo(rID)
Response.Write echo(eID)

dim a0

a0 = Split("a;b;c",";")
a0 = ""

If a0="" Then
  Response.Write "<1>"
End If


a1 = rs_Row("","SELECT ParCode,ParName FROM AdmPara WHERE ParCode='Config_Name'")
For i=0 To uBound(a1)
  Response.Write "<br>"&a1(i)
Next


Response.Write "<hr>"


a1 = rs_Row("","SELECT ParCode,ParName,ParFlag FROM AdmPara WHERE ParCode='Config_Name'")
For i=0 To uBound(a1)
  Response.Write "<br>"&a1(i)
Next


Response.Write "<hr>"


'记录数:UBound(aVal,2)+1，栏位数:UBound(aVal, 1)+1。
a2 = rs_Tab("","SELECT ParCode,ParName,ParFlag FROM AdmPara WHERE ParCode LIKE 'Config%'")
For i=0 To UBound(a2,2)
For j=0 To UBound(a2,1)
  Response.Write " : "&a2(j,i)
Next
Response.Write "<br>"
Next

a2 = rs_Tab("","SELECT ParCode,ParName,ParFlag FROM AdmPara WHERE ParCode LIKE 'xxxConfig%'")
If a2 <> Null Then
For i=0 To UBound(a2,2)
For j=0 To UBound(a2,1)
  Response.Write " : "&a2(j,i)
Next
Response.Write "<br>"
Next
End If


%>

&nbsp; &nbsp;　
</body>
</html>
