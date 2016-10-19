<%
DefMD = "InfN124" 
DefTP = "" 

MD = RequestS("ModID",3,48) : If MD="" Then MD="TraT124"
TP = RequestS("TypID",3,240)

KW = RequestS("KeyWD",3,24)
Page = RequestS("Page","N",1)
Flag = Request("Flag")
ID = RequestS("KeyID",3,48)

MDName = Get_SOpt("TraA124;TraN124;TraT124;TraJ124","公司介绍;新闻中心;供求信息;招聘信息",MD,"Val")
ModTab = "TradeInfo"

If MD="TraA124" AND TP="" Then
  TP = rs_Val("","SELECT TOP 1 TypID FROM TradeType WHERE TypMod='"&MD&"' AND LogAUser='"&US&"' ORDER BY TypID ")
  Flag = "Cont"
End If

sqlK = " WHERE ( KeyMod='"&MD&"' ) " 
If KW&"" <> "" Then
  sqlK = sqlK & " AND ( InfSubj LIKE '%"&KW&"%' ) " 
End If
If TP&"" <> "" Then
 If US&"" <> "" Then
  sqlK = sqlK & " AND ( InfTyp2='"&TP&"' ) " 
 Else
  sqlK = sqlK & " AND ( InfType='"&TP&"' ) " 
 End If
End If
If US&"" <> "" Then
  sqlK = sqlK & " AND ( LogAUser='"&US&"' ) " 
End If

'tmpList = "FAQ" 
'Response.Write sqlK&ModTab&tmpList&TypMD

%>
