<!--#include file="config.asp"-->
<!--#include file="../../sadm/func2/func_obj.asp"-->
<!--#include file="../../sadm/func2/upremote.asp"-->
<%

tTimer1 = Timer()

sub appPara(sVal)
  Dim b,m
  sVal = Show_HText(sVal,120)
  sVal = Replace(sVal,"( +","(+")
  sVal = Replace(sVal," )",")")
  If inStr(sVal,"(+")>0 Then
	b = Split(sVal,"(+")
	sk2 =sk2& Trim(b(0)) &"@"
	For m=1 To 2
	If m<=uBound(b) Then
	 b(m) = Replace(b(m),")","")
	 If m=1 Then
	  If inStr(b(m),"$")>0 Then
		b(m) = Replace(b(m),"$","")
		sk3 =sk3& Trim(b(m)) &"@"
	  Else
		sk3 =sk3& "0@"
	  End If
	 Else
	  If inStr(b(m),"Kg")>0 Then
		b(m) = Replace(b(m),"Kg","")
		sk4 =sk4& Trim(b(m)) &"@"
	  Else
		sk4 =sk4& "0@"
	  End If
	 End If
	End If
	Next
  Else
	sk2 =sk2& Trim(jVal) &"@" 
	sk3 =sk3& "0@"
	sk4 =sk4& "0@"
  End If
		
End sub

'no=4&
'ID=0BA8C830-0D3809ABJ&
'ModID=PicS224&
'KeyCode=9SJ7T2-NEDJ0H&
'Url=http%3A%2F%2Fwww%2Ebmsbattery%2Ecom%2Findex%2Ephp%3Fmain%5Fpage%3Dproduct%5Finfo%26amp%3BcPath%3D1%5F10%5F11%26amp%3Bproducts%5Fid%3D9&id1=dtpic-2010-9S-J7T2.32J&
'id2=0H&
'Now=2010-9-25 16:36:22&
'Frm=www.bmsbattery.com&
'Img=images/LiFePO4_38120.jpg

  no = Request("no")
  id1 = Request("id1")
  id2 = Request("id2")
  Frm = Request("Frm")
  Typ = Request("Typ")
  KeyID = id1&id2
  sSubj = ""

        'http://www.bmsbattery.com/index.php?main_page=product_info&amp;cPath=1_10_11&amp;products_id=89
  tUrl = "http://www.bmsbattery.com/index.php?main_page=product_info&cPath=25&products_id=98"
  'Url = "http://www.bmsbattery.com/index.php?main_page=product_info&cPath=1_10_11&products_id=116"
  
  Url = Request("Url") : Url = Replace(Url,"&amp;","&")
  If Url="" Then Url=tUrl
  sHtml = OutPage(Url,"iso-8859-1")&"" : sHbak = sHtml 
  p1 = OutSPos(sHtml,"<!--eof Form start-->",InfQ1) 
  p2 = OutSPos(sHtml,"<!--bof Form close-->",InfQ2)
  If p1>0 And p2-p1>0 Then
    sHtml = Mid(sHtml,p1,p2-p1)
	sHtml = Replace(sHtml,"images/","http://www.bmsbattery.com/images/")
	f1 = "Y"
  Else
    sHtml=""
	f1 = "N"
  End If
  
  sSubj = OutSFlag(sHtml,"<!--bof Product Name-->","<!--eof Product Name-->")
  sSubj = Show_HText(sSubj,1200)
  
  sPrice = OutSFlag(sHtml,"<!--bof Product Price block -->","<!--eof Product Price block -->")
  sPrice = Show_HText(sPrice,1200)
  sPrice = Replace(sPrice,"Starting at: ","")
  sPrice = Replace(sPrice,"$","")
  
  sCont = OutSFlag(sHtml,"<div id=""productDescription_1"" class=""productGeneral biggerText"">","<!--#AttributeOptions#-->")
  'sCont = Show_HText(Cont,12000)
  If Config_Cont="DB" Then
    xxxCont = sCont
	xxxCont = Replace(sCont,"'","''")
  Else
    xxxCont = ""
  End If
  
  sPara = OutSFlag(sHtml,"<!--bof AttribsOnTab sc2-->","<!--eof AttribsOnTab sc2-->")
  'sPara = Show_HText(sPara,12000)
  
  sPar3 = OutSFlag(sHtml,"<!--bof DetailsOnTab sc3-->","<!--eof DetailsOnTab sc3-->")
  aPar3 = Split(sPar3,"<li>")
  For i=1 To uBound(aPar3)
  If inStr(aPar3(i),"</li>")>0 Then
    iP3 = Show_HText("<li>"&aPar3(i),1200)
	If inStr(iP3,"Model:")>0 Then
	  iCode = iP3
	  iCode = Replace(iCode,"Model: ","")
	End If 
	If inStr(iP3,"Shipping Weight:")>0 Then
	  iWeight = iP3
	  iWeight = Replace(iWeight,"Shipping Weight: ","")
	  iWeight = Replace(iWeight,"Kg","")
	End If
	If inStr(iP3,"Manufactured by:")>0 Then
	  iMade = iP3
	  iMade = Replace(iMade,"Manufactured by: ","")
	End If
  End If
  Next
  'sPar3 = Show_HText(sPar3,12000)
  
  sImg1 = OutSFlag(sHtml,"<!--bof Main Product Image -->","<!--eof Main Product Image-->")
  sImg1 = Server.HTMLEncode(Get_1Url(sImg1,"href="))
  'sImg1 = Get_HLinks(sImg1,"<a[^>]*(href=[^>]*)[^>]*>([^<]*)</a>")
  '<a[^>]*(href=[^>]*)[^>]*>(<img[^<]*>)</a>
  
  sImg8 = OutSFlag(sHtml,"<!--bof Additional Product Images -->","<!--eof Additional Product Images -->")
  'sImg8 = Get_HLinks(sImg8,"<a[^>]*(href=[^>]*)[^>]*>([^<]*)</a>") 
  
  aPara = Split(sPara,"<h4")
  Dim ak(96) :k = 0
  For i=1 To uBound(aPara)
	  aPara(i) = Replace(aPara(i),vbcrlf,vblf)
	  aPara(i) = Replace(aPara(i),vbcr,vblf)
	  bPara = Split(aPara(i),vblf)
	  k = (i-1)*5 +1
	  sk2 = ""
	  sk3 = ""
	  sk4 = ""
  For j=0 To uBound(bPara)  
    jVal = bPara(j)
    If jVal<>"" Then
      If inStr(jVal,"</h4>")>0 Then
		jVal = Show_HText("<h4"&jVal,120)
		ak(k) = Trim(jVal)
	  End If
      If inStr(jVal,"</option>")>0 Then '  Or inStr(jVal,"type=""radio""")>0
		Call appPara(jVal)
      End If
    End If
  Next
	  ak(k+1) = sk2
	  ak(k+2) = sk3
	  ak(k+3) = sk4
	  tVal = ak(k)&": "&ak(k+1)&ak(k+2)&ak(k+3)
	  'Response.Write vbcrlf&tVal
  Next
  sTyp2 = ""
  For i=1 To 96 
    sTyp2 = sTyp2&"^"&ak(i)
  Next
  
  KeyCode = iCode 'RequestS("KeyCode",3,24)
  SetTop = "888"
  SetHot = "N"
  SetShow = "Y"
  SetSubj = "000000"
  IP = Get_CIP()
  LogATime = DateAdd("s",no,Request("Now"))
  
  Dim ap1(30) : sp1 = ""
  ap1(3)=iWeight '"" '重量
  ap1(8)="888" '库存
  ap1(2)=iMade '"" '生产
  ap1(4)=sPrice '"" '价格
  ap1(5)="0" '原价
  For i=1 To 96 
    sp1 = sp1&"^"&ap1(i)
  Next
  sp1 = Replace(sp1,vbcrlf,"")
  sp1 = Replace(sp1,vbcr,"")
  sp1 = Replace(sp1,vblf,"")
  
  ImgName = ""
  If sImg1<>"" Then
    Call fold_add9(upRoot, KeyID, 0)
	SaveFileType = Mid(sImg1, InstrRev(sImg1, ".") + 1) 
	SaveFileName = Get_FmtID("mdhnsx","")&"_"&Rnd_ID("KEY",4)&"(0."&SaveFileType
	locPath = upRoot&Replace(KeyID,"-","/")&"/"&SaveFileName
	If RemoteSaveFile(locPath, sImg1, 9600, "(.jpg.gif.jpeg.png.tif.swf.flv)") Then
	  'sHtml = Replace(sHtml,Img,locPath)
	  ImgName = "^"&SaveFileName&"^^^^^^^^"
	  f2 = "Y"
	Else
	  f2 = "N"
	End If
  Else
      f2 = "-"
  End If
  
  sql = " INSERT INTO "&ModTab&" (" 
  sql = sql& "  KeyID" 
  sql = sql& ", KeyMod" 
  sql = sql& ", KeyCode" 
  sql = sql& ", InfType" 
  sql = sql& ", InfTyp2" 
  sql = sql& ", InfSubj"  
  sql = sql& ", InfCont" 
  sql = sql& ", InfPara"
  sql = sql& ", SetSubj" 
  sql = sql& ", SetRead" 
  sql = sql& ", SetHot" 
  sql = sql& ", SetTop" 
  sql = sql& ", SetShow" 
  sql = sql& ", ImgName" 
  sql = sql& ", LogAddIP" 
  sql = sql& ", LogAUser" 
  sql = sql& ", LogATime" 
  sql = sql& ", LogEditIP"
  sql = sql& ")VALUES(" 
  sql = sql& "  '" & KeyID &"'" 
  sql = sql& ", '" & ModID &"'" 
  sql = sql& ", '" & KeyCode &"'" 
  sql = sql& ", '" & Typ &"'" 
  sql = sql& ", '" & sTyp2 &"'" 
  sql = sql& ", '" & sSubj &"'" 
  sql = sql& ", '" & xxxCont &"'" 
  sql = sql& ", '" & RequestSafe(sp1,3,510) &"'" 
  sql = sql& ", '" & SetSubj &"'" 
  sql = sql& ", 0" 
  sql = sql& ", '" & SetHot &"'" 
  sql = sql& ", " & SetTop &"" 
  sql = sql& ", '" & SetShow &"'" 
  sql = sql& ", '" & ImgName & "'" 
  sql = sql& ", '" & IP &"'" 
  sql = sql& ", '" & Get_PUser(PrmFlag) &"'" 
  sql = sql& ", '" & LogATime &"'" 
  sql = sql& ", '(imp_do.asp)'" 
  sql = sql& ")"
  'Response.Write sql
  If sSubj <> "" Then
    f1 = "Y"
	Call rs_DoSql(conn,sql)
	Call add_sfFile()
  Else
    f1 = "N"
  End If

  aImg8 = Split(sImg8,"<a ") : noj = 12
  For i=0 To uBound(aImg8)
  noj = noj+i
  iImg1 = Get_1Url(aImg8(i),"href=") 
  If iImg1<>"" Then 
    Call fold_add9(upRoot, KeyID, 0)
	SaveFileType = Mid(iImg1, InstrRev(iImg1, ".") + 1) 
	SaveFileName = Get_FmtID("mdhnsx","")&"_"&Rnd_ID("KEY",4)&"("&noj&"."&SaveFileType
	locPath = upRoot&Replace(KeyID,"-","/")&"/"&SaveFileName
	If RemoteSaveFile(locPath, iImg1, 9600, "(.jpg.gif.jpeg.png.tif.swf.flv)") Then
	sql = " INSERT INTO InfoPhoto (KeyID, KeyRe, KeyMod"  
	sql = sql& ", InfSubj, InfCont, SetTop, ImgName" 
	sql = sql& ", LogAddIP, LogAUser, LogATime" 
	sql = sql& ")VALUES(" 
	sql = sql& "  '" & Get_AutoID(24) &"', '" & KeyID &"', '" & ModID &"'" 
	sql = sql& ", '" & iImg1 &"', '(Import)', " & SetTop &", '"&Replace(KeyID,"-","/")&"/"&SaveFileName&"'" 
	sql = sql& ", '" & Get_CIP() &"', '" & Session("UsrID") &"', '" & Now() &"'" 
	sql = sql& ")"
    Call rs_DoSql(conn,sql)
	End If
  End If
  Next

  Response.Write "("&id2&")OK("&f1&f2&")"
  Response.End()
  
  'Response.Write vbcrlf&"<hr>"'&sImg8
  aImg8 = Split(sImg8,"<a ")
  For i=0 To uBound(aImg8)
    iImg1 = Get_1Url(aImg8(i),"href=") 'Get_HLinks("<a "&aImg8(i),"<a[^>]*(href=[^>]*)[^>]*>([^<]*)</a>")
	'Response.Write vbcrlf&"<br>"&iImg1
  Next
  'Response.Write vbcrlf&"<hr>"&sImg1
  
  'Response.Write vbcrlf&"<hr>"&sPar3
  'Response.Write vbcrlf&"<hr>"&sPara
  
  'Response.Write vbcrlf&"<hr>"&sCont
  'Response.Write vbcrlf&"<hr>"&sPrice
  'Response.Write vbcrlf&"<hr>"&sSubj
  


tTimer1 = Timer()-tTimer1
'Response.Write vbcrlf&"<hr>"&tTimer1

%>


</body>
</html>
