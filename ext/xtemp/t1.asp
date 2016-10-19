<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="../../page/_config.asp"-->



<%

a = 4
b = 3+2
v1 = Split("a,b",",")
v2 = Split("c,d",",")
Response.Write "<br>"& IIf(a=b,"=","!=")
v = IIf(a=b,v1,v2)
Response.Write "<br>v0(vs)v1:"& v(0)&" vs "& v(1)

Response.Write "<br>"

	strsql="Select * From web_news Where nClass=17 and n_Title like '"&request("keyword")&"'"
	strsql=strsql & " Order by ID Desc"
	Response.Write strsql
	set rs = server.CreateObject("ADODB.RecordSet")
	rs.open strsql,conn,1,1

'Select * From web_news Where nClass=17 and n_Title like '请输入线路名' Order by ID Desc

'-----'123456789012345678901234
'-----'yyyymmddhhnnss9991234567
id1 = "201104290757309991234567"
id2 =  201104290757309991234567
id3 = cStr(id2) 
id4 = FormatNumber(id2,0,0,0,0)

Response.Write "<br>str:<a href='?id1="&id1&"'>"&id1&"</a>"
Response.Write "<br>num:<a href='?id2="&id2&"'>"&id2&"</a>"
Response.Write "<br>cStr<a href='?id3="&id3&"'>"&id3&"</a>"
Response.Write "<br>fmt:<a href='?id4="&id4&"'>"&id4&"</a>"

Response.Write "<br>id1:"&Request("id1")
Response.Write "<br>id2:"&Request("id2")
Response.Write "<br>id3:"&Request("id3")
Response.Write "<br>id4:"&Request("id4")

id5 = id1+"1234"
id6 = cStr(id1)+cStr(1234)
id7 = id1&"1234"
id8 = cStr(id1)&cStr(1234)
Response.Write "<br>id5:"&id5
Response.Write "<br>id6:"&id6
Response.Write "<br>id7:"&id7
Response.Write "<br>id8:"&id8

id9 = id1
Response.Write "<br>id9:"&id9
ida = id2&""
Response.Write "<br>ida:"&ida

'Response.Write "<br>"&Request("id1")

Response.End()

fn = "ysWeb_"&Config_Code&".Peace!DB"
'fn = "wj_jinjin.gif"
'fn = "index.asp"
fp = Server.MapPath(Config_Path&"upfile/#dbf#/"&fn )
'str = File_Read(fp,"iso-8859-1") 'lCase()
'fMov = ChkTrojan(str,"")

echo Now()
dt = fRead(fp) ':echo dt
re = fCheck(dt):echo re

Function fRead(xFile)
  Dim t1,dat
  t1 = Timer()
  dat= File_Read(xFile,"iso-8859-1")
  t1 = Timer()-t1
  echo "tRead:"&t1
  echo "tRLen:"&Len(dat)
  If Len(dat)>720123 Then
    dat = Left(dat,240123)&"..."&Right(dat,240123)
  End If
  fRead = dat
End Function

Function fCheck(xStr)
  Dim t1,res
  t1 = Timer()
  res = ChkTrojan(xStr,"")
  t1 = Timer()-t1
  echo "tCheck:"&t1
  fCheck = res
End Function
%>
