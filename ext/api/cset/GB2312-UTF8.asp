<% 
Function GBtoUTF8(szInput)   
Dim wch, uch, szRet   
Dim x   
Dim nAsc, nAsc2, nAsc3   
'����������Ϊ�գ����˳�����   
If szInput = "" Then  
GBtoUTF8= szInput   
Exit Function  
End If  
'��ʼת��   
For x = 1 To Len(szInput)   
wch = Mid(szInput, x, 1)   
nAsc = AscW(wch)   
If nAsc < 0 Then nAsc = nAsc + 65536   
If (nAsc And &HFF80) = 0 Then  
szRet = szRet & wch   
Else  
If (nAsc And &HF000) = 0 Then  
uch = "%" & Hex(((nAsc \ 2 ^ 6)) Or &HC0) & Hex(nAsc And &H3F Or &H80)   
szRet = szRet & uch   
Else  
uch = "%" & Hex((nAsc \ 2 ^ 12) Or &HE0) & "%" & _   
Hex((nAsc \ 2 ^ 6) And &H3F Or &H80) & "%" & _   
Hex(nAsc And &H3F Or &H80)   
szRet = szRet & uch   
End If  
End If  
Next  
GBtoUTF8= szRet   
End Function  
Response.write GBtoUTF8("��") 
Response.write "<BR>"
Response.write GBtoUTF8("������ͨ���Ź���123��abc")  
%>
