<%

Dim EncOffNew : EncOffNew = "432628512372017297"
Dim EncOffset : EncOffset = EncOffNew


Function Peace_Enc(xPW,xID)
Dim bID,bPW,iLen,pLen,k,n1,n2,n0,oStr
  bID = xID : iLen = Len(bID)
  bPW = xPW : pLen = Len(bPW)
  n1 = 0 : n2 = 0 : oStr = ""
  If iLen<2 Or pLen<5 Or bPW="12345" Or bPW="123456" Or bPW="adm" Or bPW="admin" Then 
	oStr = "Enc_"&Now()&"_"&xPW&"_"&xID&"_Error" : pLen = 0
  End If
  If iLen<pLen Then
    For k=1 To pLen-iLen
	  bID = bID&" "
	Next
  End If
  For k=1 To iLen
	n2 = n2 + Asc(Mid(bID,k,1))
  Next
    n2 = n2 Mod 17
  For k=1 To pLen
	n1 = n2 + (Mid(EncOffset,(k Mod 17)+1,1))
	n0 = Asc(Mid(bPW,k,1)) + Asc(Mid(bID,k,1)) - n1
    oStr = oStr&cStr(Hex(n0))
  Next
Peace_Enc = oStr
End Function

Function Peace_Unc(xPW,xID)
Dim bID,bPW,iLen,pLen,k,n1,n2,n0,oStr
  bPW = xPW : pLen = Len(bPW)/2
  bID = xID : iLen = Len(bID)
  n1 = 0 : n2 = 0 : oStr = ""
  If iLen<2 Or pLen<5 Then 
	oStr = "Unc_"&Now()&"_"&xPW&"_"&xID&"_Error" : pLen = 0
  End If
  If iLen<pLen Then
    For k=1 To pLen-iLen
	  bID = bID&" "
	Next
  End If
  For k=1 To iLen
	n2 = n2 + Asc(Mid(bID,k,1))
  Next
    n2 = n2 Mod 17
  For k=1 To pLen
	n1 = n2 + (Mid(EncOffset,(k Mod 17)+1,1))
	cc = Mid(bPW,(k-1)*2+1,2)
	nc = Peace_HO(Left(cc,1))*16 + Peace_HO(Right(cc,1))
	nc = nc + n1 - Asc(Mid(bID,k,1))
    oStr = oStr&Chr(nc)
  Next
Peace_Unc = oStr
End Function


Function Peace_HO(xChar)
Dim n
  If xChar<"A" Then
    n = Asc(xChar)-48
  Else
    n = Asc(xChar)-65+10
  End If
Peace_HO=n
End Function


Function Peace_eSN(xPW,xID)
Dim bPW,bID,oSN,oPW,eStr
  bPW=xPW : bID=xID 
  oSN=EncOffset
  oPW=Peace_Unc(bPW,bID)
  EncOffset=EncOffNew
  eStr=Peace_Enc(oPW,bID)
  EncOffset=oSN
Peace_eSN = eStr
End Function


Response.Write "<br>"&Peace_Enc("pw123456","sa123")
Response.Write "<br>"&Peace_Enc("pw123456","sa124")
Response.Write "<br>"&Peace_Enc("pw123456","sa125")
Response.Write "<br>"&Peace_Enc("pw123456","sa115")
Response.Write "<br>"&Peace_Enc("pw123456","sa105")
Response.Write "<br>"&Peace_Unc("DBBB999E59484B","sa123")


%>
