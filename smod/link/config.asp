
<%

'user//index.asp,user_editpw.asp,...
PrmPath = LCase(Request.ServerVariables("PATH_INFO")) 
PrmFile = Mid(PrmPath,InStrRev(PrmPath,"/")+1) 
If Left(PrmFile,5)="info_" Then
  PrmID = "HomLnk1" 
ElseIf Left(PrmFile,4)="ad_p" Then
  PrmID = "HomAdvert" 'HomAdv2
ElseIf Left(PrmFile,4)="ad_t" Then
  PrmID = "HomAdvert" 'HomAdv3
ElseIf Left(PrmFile,3)="ad_" Then
  PrmID = "HomAdvert" 'HomAdv1
Else
  PrmID = "ModHome" '"System"
End If

'Response.Write PrmID
Call Chk_Perm1(PrmID,"")

%>