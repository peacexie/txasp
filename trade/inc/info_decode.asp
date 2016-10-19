<%

ID = RequestS("KeyID",3,48) 
ModTab = rel_IDTab(ID)


Dim aImgs
rs.Open "SELECT * FROM "&ModTab&" WHERE KeyID='"&ID&"' ",conn,1,3 
  if NOT rs.eof then 
KeyID = rs("KeyID")
KeyMod = rs("KeyMod")
KeyCode = rs("KeyCode")
InfType = rs("InfType")
InfSubj = Show_Text(rs("InfSubj"))
InfPara = rs("InfPara")&"" : aPara = Split(InfPara,"^")
SetRead = rs("SetRead")+1
LogATime = rs("LogATime")
LogAUser = rs("LogAUser")
xxxCont = Show_Html(rs("InfCont"))
ImgName = rs("ImgName")&"" 
aImgs = Show_sfImgs(ImgName,KeyID)
rs("SetRead") = SetRead
rs.Update() 
  else
KeyID = ""
  end if 
rs.Close()


If KeyID = "" Then 
  Response.End() 
End If


MD = KeyMod '"PicS124"'
TP = InfType '"S110048;S120104;"'
TPLay = TP

MDName = Get_SOpt("TraN124;TraT124;TraJ124","新闻中心;供求信息;招聘信息",MD,"Val")
ModTab = "TradeInfo"

%>

