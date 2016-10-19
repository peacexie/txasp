
<%

''// 简易加密
Dim CfgEnc : CfgEnc = "19790913" 

Function Peace_EncCode(xStr)   
Dim s,c1,c2
   For i=1 To Len(xStr)
     c1 = Mid(xStr,i,1)
	 c2 = Mid(CfgEnc,(i Mod 8)+1,1)
	 s = s&cStr(Hex(Asc(c1)+Asc(c2)))
   Next
Peace_EncCode = s
End Function

Function Peace_unEnc(xStr)   
Dim s,j,c,n
   For i=1 To Len(xStr) Step 2
     j = (i+1)/2 'i = 1,3,5
	 c = Mid(CfgEnc,(j Mod 8)+1,1)
	 n = CInt("&h"&Mid(xStr,i,2)) - Asc(c)
	 s = s&Chr(n)
   Next
Peace_unEnc = s
End Function

Response.Write "<br>"&Peace_EncCode("peace")
Response.Write "<br>"&Peace_unEnc("AC98A991ACA4AA95")

%>

