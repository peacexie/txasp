<!--#include file="../../sadm/func1/func1.asp"-->
<!--#include file="../../sadm/func2/func2.asp"-->
<!--#include file="../../sadm/func1/func_file.asp"-->
<%

Call Chk_URL()

ModTab = RequestS("ModTab","C",48)
ModID = RequestS("ModID","C",48)
KeyID = RequestS("KeyID","C",48)
NO = RequestS("NO","C",4)

pRoot = Config_Path&"upfile/"
pFile = Replace(KeyID,"-","/")&"/votemood.txt"

Call fold_add9(pRoot,KeyID,0)
If fil_exist(pRoot&pFile) Then
  sv = File_Read(pRoot&pFile,"utf-8")
  sa = Split(sv,"|")
  sv = "" : sn = 0
  For i=1 To 8
    iv = 0
	If cStr(i)=NO Then iv = 1
	sv = sv&"|"&sa(i)+iv
	sn = sn+sa(i)+iv
  Next
  sv = sn&sv
  'Response.Write "<br>f11"
Else
  sv = ""
  For i=1 To 8
    iv = 0
	If cStr(i)=NO Then iv = 1
	sv = sv&"|"&iv
  Next
  sv = "1"&sv
  'Response.Write "<br>f21"
End If

'Response.Write "<br>"&pRoot&pFile
Call File_Add2(pRoot&pFile,sv,"utf-8") 
Response.Write "vmShow('"&sv&"');"

%>