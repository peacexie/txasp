
<%

US = RequestS("UsrID","C",48)
If US="" Then 
  Response.Redirect "index.asp"
End If

  rs.Open "SELECT * FROM [TradeCorp] WHERE LogAUser='"&US&"'",conn,1,1 
  If NOT rs.eof then 
MemType = rs("InfType")
MemTyp2 = rs("InfTyp2")
MemSubj = Show_Text(rs("InfSubj"))
MemCont = Show_Text(rs("InfCont"))
MemName = rs("LnkName")
MemMod = rs("KeyMod") '' UserCorp
MemMName = Get_SOpt(mCfgCode,mCfgName,MemType,"Val")
  Else
'Response.Redirect "ypage.asp"
  End If
  rs.Close()

%>
