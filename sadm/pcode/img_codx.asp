<%

' 禁止缓存
'Response.Expires = -9999
Response.buffer=true
Response.ExpiresAbsolute = Now() - 1
Response.Expires = 0
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma","no-cache"
Response.AddHeader "cache-ctrol","no-cache"
Response.ContentType = "Image/BMP"

nRnd = (Timer*100) Mod 3
'Session("ChkCode") = Rnd_ID("K",3+nRnd) 'nRnd&
Config_PSess = Request("Config_PSess")
If Config_PSess="" Then '记录入Session
  Session("ChkCode") = Rnd_ID("K",3+nRnd)
Else
  Session(Config_PSess) = Rnd_ID("K",3+nRnd)
End If
pUrl="http://eweb.dg.gd.cn/sadm/pinc/ShowImg.aspx?ChkCode="&Session("ChkCode")&"&xRnd="
Response.BinaryWrite GetBody(pUrl)


Function GetBody(xUrl) 
  on error resume next
  Set Retrieval = CreateObject("Microsoft.XMLHTTP") 
  With Retrieval 
  .Open "Get", xUrl, False, "", "" 
  .Send 
  GetBody = .ResponseBody
  End With 
  Set Retrieval = Nothing 
End Function

Function Rnd_ID(xType,xLen2)
  Dim objStr,orgNum,orgCap,orgLow,orgSp1,orgSp2,oStr,rStr,rMax,ni,xn,ch,bLen
  orgNum = "0123456789"                   ' 10   xType = 0,A,C,K,X
  orgCap = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"   ' 26  
  orgKey = "34569ACEFGHJKMNPQRSWXY"       ' 8B,UV,T7 22   Ii  Ll  Oo  Zz
  'orgPeace = "玉刊示末未打巧正永扒功扔顺去甘世古本可丙左右石布平卡北占且旦目甲申叮号田由史央兄叼叫另叨四生失禾丘付仗仙白仔斥瓜乎令用甩印句匆外冬包主市立代半汁它讨必司尼民谢出加李召皮台矛幼式扛英寺吉扣考托老地耳共芒芝朽臣再西在有百存而匠灰列成至此尖光早曲同吃因吸帆回年朱先丢舌竹乒乓休伍伏伐延件任份仰仿自血向似行舟全合兆企肌朵旬旨各名多争冰亦次衣充妄羊并米州汗江池忙宇守宅字安那迅收防如好羽巡王井夫天元木五支不太犬尤友匹巨牙屯比互切瓦止少日中内水午牛手毛升仁什片仆化币仍斤爪反介父今分乏公月氏勿丹勾文六方火心尺引巴孔以允予幻鸽"
  orgPLen = Len(orgPeace)-10

  oStr = ""
  rStr = ""  
  If xType = "0" Then
    oStr = orgNum
	rMax = 10
  ElseIf uCase(xType) = "A" Then
    oStr = orgCap
	rMax = 26
  ElseIf uCase(xType) = "C" Then
    oStr = orgPeace
	rMax = orgPLen
  ElseIf uCase(xType) = "X" Then
    oStr = orgPeace&orgKey
	rMax = orgPLen+23
  Else 'K
    oStr = orgKey
	rMax = 22 
  End If 
      Randomize
	  objStr = ""
  bLen = xLen2
  if bLen Mod 2 = 0 then
    bLen = bLen + 1
  end if
	  Do While ni < bLen
	         xn = Int(rMax*Rnd()) ' 0 ~ (rMax-1)
			 objStr = objStr&Mid(oStr,xn+1,1) 
      ni = ni + 1 
	  Loop	  
  Rnd_ID = Left(objStr,xLen2) 
End Function

%>