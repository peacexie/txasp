

<%

AID = get_AutoID(15)
cnOld = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source="&Server.MapPath("/script/#dbf#/old.mdb" ) 

 sql2 = " SELECT * FROM [P_Product] ORDER BY P_ID " 
 Set rs2=Server.Createobject("Adodb.Recordset")
 rs2.Open Sql2,cnOld,1,1
 Do While Not rs2.eof 
   P_ID = rs2("P_ID")
   minimage = "2008/"&rs2("minimage")&""
   images = "2008/"&rs2("images")&""
   P_NAME = rs2("P_NAME")&""
   P_CAT = rs2("P_CATEGORY")
   P_DES = rs2("P_DESCRIPTION")
   P_ADD = RequestSafe(rs2("P_ADDATE"),"D","1900-12-31")

'TypID	TypMod	TypFlag	TypLayer	TypName	TypNam2	TypDeep	TypResume	ImgName	ImgTitle	ImgWidth	ImgHeight	ImgScale	LogAddIP	LogAUser	LogATime	LogEditIP	LogEUser	LogETime
's110048	Pics124		s110048;	铸压产品	Precision die-casting	1	The metal part of computer connector, lighting fixture, bicycle components, gear, tube adapter, kitchen hardware, architecture hardware, furniture hardware, shoe hardware	电子、通讯、电脑、DVI连接器、光纤通讯收发器、DVD光头、微电机零件、汽车零件等系列压铸产品					127.0.0.1	peace	2008-3-27 15:16:08	127.0.0.1	peace	2008-4-21 10:55:31
's110056	Pics124		s110056;	模具制品	Precision mold	1	We can design and produce hot runner injection mold with the hot runner system and components of DME, HUSKEY,Mold-Masters,YUDO etc.	热流道塑胶模具、锌铝压铸模具、电木模具等各式模具的专业设计与制造					127.0.0.1	peace	2008-3-27 15:16:12	127.0.0.1	peace	2008-4-21 10:55:35
's110064	Pics124		s110064;	冲压产品	Press the products	1	Type描述	五金冲压件					127.0.0.1	peace	2008-3-27 15:16:18	127.0.0.1	peace	2008-4-21 14:10:54
's110072	Pics124		s110072;	锂电池	TypeName	1	Type描述	聚合物锂离子电池的典型应用主要有手机、笔记本电脑、个人数字助理(PDA)、 MP3、蓝牙耳机、便携式DVD、电动自行车、电动工具、通信设备、航模等等					127.0.0.1	peace	2008-4-21 10:47:01	127.0.0.1	peace	2008-4-21 14:10:57

'198	0	压铸产品	.198.	/压铸产品/	0	0	
'204	0	模具制品	.204.	/模具制品/	0	0	
'214	0	冲压产品	.214.	/冲压产品/	0	0	
'215	0	锂电池	.215.	/锂电池/	0	0	

If CStr(P_CAT)="198" Then
 InfType="s110048;"
ElseIf CStr(P_CAT)="204" Then
 InfType="s110056;"
ElseIf CStr(P_CAT)="214" Then
 InfType="s110064;"
ElseIf CStr(P_CAT)="215" Then
 InfType="s110072;"
Else
 InfType="s110072;"
End If
If minimage="2008/" AND images = "2008/" Then
 minimage=""
 images=""
ElseIf images = "2008/" Then
 images=minimage
ElseIf minimage = "2008/" Then
 minimage=images
End If

KeyID = AID&Right("0000"&P_ID,5)&Rnd_ID("KEY",4)
IDModel = "Pics124"
KeyCode = rs_AutID(conn,"NewsPics","KeyCode","Pic-","1","")
SetUBB = "X"

sql = " INSERT INTO [NewsPics] (" 
sql = sql& "  KeyID" 
sql = sql& ", KeyRe" 
sql = sql& ", KeyMod" 
sql = sql& ", KeyCode" 
sql = sql& ", InfType" 
sql = sql& ", InfSubject" 
sql = sql& ", InfContent" 
sql = sql& ", SetSub" 
sql = sql& ", SetRead" 
sql = sql& ", SetHot" 
sql = sql& ", SetUBB" 
sql = sql& ", SetTop" 
sql = sql& ", SetShow" 
sql = sql& ", ImgName" 
sql = sql& ", ImgNam2" 
sql = sql& ", ImgAlign" 
sql = sql& ", ImgTitle" 
sql = sql& ", ImgWidth" 
sql = sql& ", ImgHeight" 
sql = sql& ", ImgScale" 
sql = sql& ", LogAddIP" 
sql = sql& ", LogAUser" 
sql = sql& ", LogATime" 
sql = sql& ")VALUES(" 
sql = sql& "  '" & KeyID &"'" 
sql = sql& ", '" & KeyID &"'" 
sql = sql& ", '" & IDModel &"'" 
sql = sql& ", '" & KeyCode &"'" 
sql = sql& ", '" & InfType &"'" 
sql = sql& ", '" & P_NAME &"'" 
sql = sql& ", '" & P_DES &"'" 
sql = sql& ", '-'" 
sql = sql& ", 0" 
sql = sql& ", 'N'" 
sql = sql& ", '" & SetUBB &"'" 
sql = sql& ", '6'" 
sql = sql& ", 'Y'" 
sql = sql& ", '"&minimage&"'" 
sql = sql& ", '"&images&"'"
sql = sql& ", ''" 
sql = sql& ", ''" 
sql = sql& ", 1" 
sql = sql& ", 1" 
sql = sql& ", 1" 
sql = sql& ", '(IP.INP)'" 
sql = sql& ", '" & Session("UsrID") &"'" 
sql = sql& ", '" & P_ADD &"'" 
sql = sql& ")" ':Response.Write slq
 If P_NAME<>"" Then
 'Call rs_Dosql(conn,sql)	
 End If
 
 rs2.movenext
 loop
 rs2.close()
 set rs2 = nothing
'/////////////////////////////////



%>