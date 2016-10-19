<%


Function get_ModCopy(xMod)
  Dim flag,tChr,tMod
  flag = false
  If inStr(Config_PCopy,xMod)>0 Then
    tChr = Mid(xMod,5,1)
    tMod = Mid(xMod,1,4)&"1"&Mid(xMod,6)
    If CStr(tChr)<>"1" AND inStr(Config_PCopy,tMod)>0 Then
      flag = true 
    End If	
  End If	   
  get_ModCopy = flag
End Function


If Request("ModID")<>"" Then
  ModID = RequestS("ModID","C",48)
  Session("ModID") = ModID
ElseIf Session("ModID")&""<>"" Then
  ModID = Session("ModID")
Else
  ModID = ""
End If
  
  
PrmPath = LCase(Request.ServerVariables("PATH_INFO")) 
PrmFile = Mid(PrmPath,InStrRev(PrmPath,"/")+1) 
If PrmFile="dep_edit.asp" Then
  PrmID = ModID '"" 'ModID 
Else
  PrmID = ModID
End If


ModTab = rel_ModTab(ModID)
upPart = rel_TabPath(ModTab)
Session("upPart") = upPart
'Response.Write vbcrlf&ModTab&upPart


%>

