<%

' ��ֹ����
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
If Config_PSess="" Then '��¼��Session
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
  'orgPeace = "��ʾĩδ���������ǹ���˳ȥ�����ű��ɱ�����ʯ��ƽ����ռ�ҵ�Ŀ���궣������ʷ���ֵ����߶����ʧ�������ɰ��г�Ϻ�����˦ӡ����ⶬ������������֭���ֱ�˾����л��������Ƥ̨ì��ʽ��Ӣ�¼��ۿ����ϵض���â֥�೼�������аٴ�������г����˼������ͬ���������������ȶ�����ƹ����������Ӽ��η�������Ѫ��������ȫ�����󼡶�Ѯּ��������������³��������ݺ�����æ����լ�ְ���Ѹ�շ������Ѳ��������Ԫľ��֧��̫Ȯ����ƥ�����ͱȻ�����ֹ��������ˮ��ţ��ë����ʲƬ�ͻ����Խ�צ���鸸��ַ��������𵤹����������ĳ����Ϳ�������ø�"
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