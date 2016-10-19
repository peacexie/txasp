
<%

If Request("ModID")<>"" Then
  ModID = RequestS("ModID","C",48)
  Session("ModID") = ModID
ElseIf Session("ModID")&""<>"" Then
  ModID = Session("ModID")
Else
  ModID = "InfN124"
End If

'Call Chk_Perm1(ModID,"") 

%>