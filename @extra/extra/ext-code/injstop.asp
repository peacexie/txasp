<%

' Paras:
Dim AppAJaxCode,AppAJaxData
AppAJaxCode = "aJax471442"
AppAJaxData = "Data9T38U8U"

yAct = Request("yAct")
If yAct="Set" Then
  Session(AppAJaxCode)=AppAJaxData
  Response.Write "document.write('<!--Error-->');"
ElseIf yAct="Show" Then
  Response.Write "Error"&vbcrlf
  Response.Write Session(AppAJaxCode)
ElseIf yAct="Clear" Then
  Session(AppAJaxCode)=""
  Response.Write "document.write('<!--Error-->');"
End If

%>