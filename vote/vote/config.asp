<%

If Request("ModID")<>"" Then
  ModID = RequestS("ModID","C",48)
  Session("ModID") = ModID
ElseIf Session("ModID")&""<>"" Then
  ModID = Session("ModID")
Else
  ModID = "BBSVA24"
End If

Call Chk_Perm1(ModID,"") 

PathImg = "../../upfile/vote/"
PathHtml = "../../upfile/vote/"
TypeFile = ""
TypeImg = ".JPG/.GIF/.JPEG"

%>
