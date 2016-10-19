<%

TP = RequestS("TypID",3,240)
KW = RequestS("KeyWD",3,24)
Page = RequestS("Page","N",1)
Flag = Request("Flag") ' 列表显示 | 阵列显示 
TPLay = Request("TPLay") ' 折叠菜单（全部）
TypMD = Request("TypMD") ' 折叠菜单(部分)

MDName = GetMName(MD)
If TP="" Then
 TPName = vPMsg_TName
Else
 TPName = GetTName(MD,TP,"")
 'GetNLay(MD,TP,"","","")
End If
'Response.Write " :: "&MDName&" &gt;&gt; "&TPName

  sqlK = sqlK & " WHERE ( KeyMod='"&MD&"' ) " 
If KW&"" <> "" Then
  sqlK = sqlK & " AND ( InfSubj LIKE '%"&KW&"%' "
  sqlK = sqlK & " OR InfCont LIKE '%"&KW&"%' "
  sqlK = sqlK & " ) " 
End If
If TP&"" <> "" Then
  sqlK = sqlK & " AND ( InfType LIKE '%"&TP&"%' ) " 
End If
 
%>

