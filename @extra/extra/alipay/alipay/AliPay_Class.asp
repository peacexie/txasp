<%


'///////////////////////////////////////////////////////////////////////////////
'//              支付宝调用类
'// 类    名:    Qlwz_AliPay
'// 作    用:    asp通用utf-8，gb2312编码通用接口
'// 作    者:    情留メ蚊子
'// 技术支持:    qlwz@qq.com  QQ:540644769
'// 官方网站:    www.94qing.com
'///////////////////////////////////////////////////////////////////////////////
Class Qlwz_AliPay
	Private Q_PartnerID, Q_SellerEmail, Q_SignCode, Q_VarNotifyUrl, Q_NotifyUrl, Q_ReturnUrl, Q_Service, Q_Charset
	Private Q_Subject, Q_Body, Q_Out_Trade_No, Q_Price, Q_Quantity, Q_Discount, Q_Paymethod
	Private Q_Post_fee, Q_Post_payment, Q_Post_type, Q_Post_fee_1, Q_Post_payment_1, Q_Post_type_1, Q_Post_fee_2, Q_Post_payment_2, Q_Post_type_2
	Private Q_ShowUrl, Q_buyer_name, Q_buyer_address, Q_buyer_zip, Q_buyer_phone, Q_buyer_mobile, Q_buyer_email
	Private Q_siteHost '// 一定要加此行 --- Peace
	Private Sub Class_Initialize()
		'///////////////////////以下是你必须设置的//////////////////////////////////
		'siteHost       =   Request.ServerVariables("HTTP_HOST")  'localhost:301 ; 218.16.99.70:616 ; 
		Q_siteHost      =   Request.ServerVariables("HTTP_HOST")
		Q_PartnerID     =	"2088101707183363"      '填写对应支付宝账户的合作者身份ID
		Q_SellerEmail   =	"mtc616@yeah.net"    '请填写支付宝签约帐户
		Q_SignCode      =	"rcaj004vn57t6ndfb6s7r7b37y62oa7j"   '填写对应支付宝帐户的安全校验码
		Q_NotifyUrl     =	"http://"&Q_siteHost&"/sadm/alipay2/AliPay_Notify.asp"  '交易过程中服务器通知的页面,,,绝对路径，要求可以访问
		Q_ReturnUrl     =	"http://"&Q_siteHost&"/sadm/alipay2/AliPay_Return.asp"  '付完款后跳转的页面,,,绝对路径，要求可以访问
		Q_Service       =   "create_direct_pay_by_user"  'create_direct_pay_by_user  (即时到账demo1) ，  trade_create_by_buyer 标准双接口
		Q_Charset       =   "UTF-8" '编码
		'seller_email   =   "mtc616@yeah.net"					    '请设置成您自己的支付宝帐户
		'partner	    =   "2088101707183363"				    '支付宝的账户的合作者身份ID
		'key	        =   "rcaj004vn57t6ndfb6s7r7b37y62oa7j"    '支付宝的安全校验码
		'show_url       =   "http://www.mtc616.com"               '网站的网址
		'notify_url	    =   "http://"&siteHost&"/sadm/alipay/Alipay_reAdmin.asp"     '付完款后服务器通知的页面 要用 http://格式的完整路径
		'return_url	    =   "http://"&siteHost&"/sadm/alipay/alipay_reUser.asp"	  '付完款后跳转的页面 要用 http://格式的完整路径
		'///////////////////////以上是你必须设置的//////////////////////////////////
		Q_Subject	    =	"测123号Peace"	'商品名称
		Q_Body		    =	"45678描述Peace"		'商品描述
		Q_Out_Trade_No  =   Year(Now())&Month(Now())&Day(Now())&Hour(Now())&Minute(Now())&Second(Now())'按时间获取的订单号
		Q_Price		    =	"0.01"			'price商品单价	0.01～50000.00 , 注：不要出现3,000.00，价格不支持","号
		Q_Quantity      =   "1"             '商品数量,如果走购物车默认为1
		Q_Discount      =   "0"             '商品折扣
		Q_Paymethod     =   "directPay"     '赋值:bankPay(网银);cartoon(卡通); directPay(余额)
		'///////////////////////以下是标准双接口设置的//////////////////////////////
		'标准双接口无论如何第一组都要保留，不可以为空
		'第一组
		Q_Post_fee	    =	"0.00"			'物流配送费用  0.00
		Q_Post_payment  =	"BUYER_PAY"	   '物流配送费用付款方式：SELLER_PAY(卖家支付)  BUYER_PAY(买家支付)  BUYER_PAY_AFTER_RECEIVE(货到付款)
		Q_Post_type	    =	"EXPRESS"		'物流配送方式：POST(平邮)、EMS(EMS)  EXPRESS(其他快递)
		'如果需要多添加几组物流方式，可以增加第三组物流参数,如果不需要，可以为空 ---- 没有第二组不能有第三组
		Q_Post_fee_1    =	""			'物流配送费用  0.00
		Q_Post_payment_1=	""	'物流配送费用付款方式：SELLER_PAY(卖家支付)  BUYER_PAY(买家支付)  BUYER_PAY_AFTER_RECEIVE(货到付款)
		Q_Post_type_1	=	""		'物流配送方式：POST(平邮)  EMS(EMS)  EXPRESS(其他快递)
		'如果需要多添加几组物流方式，可以增加第三组物流参数,如果不需要，可以为空
		Q_Post_fee_2    =	""			'物流配送费用  0.00
		Q_Post_payment_2=	""	'物流配送费用付款方式：SELLER_PAY(卖家支付)  BUYER_PAY(买家支付)  BUYER_PAY_AFTER_RECEIVE(货到付款)
		Q_Post_type_2	=	""		'物流配送方式：POST(平邮)  EMS(EMS)  EXPRESS(其他快递)
 		'以下是可选的参数 如果没有可以为空。注意：姓名、联系地址和邮政编码 这三项要么都为空，要么都不能为空。
		Q_ShowUrl       =	"http://www.mtc616.com"    '商品的演示地址
		Q_buyer_name    =	""   '收货人姓名
		Q_buyer_address =	""   '收货人地址
		Q_buyer_zip     =	""   '邮编5 位戒6 位数字组成
		Q_buyer_phone   =	""   '收货人电话
		Q_buyer_mobile  =	""   '收货人手机 必须是11 位数字
		Q_buyer_email   =	""   '买家的支付宝账号
		'///////////////////////以上是标准双接口设置的//////////////////////////////
	End Sub

	'********************************
	'创建要发送的url
	'********************************
	Public Function GetUrl()
		Dim mystr, Sign, j, val
		If Q_Service = "create_direct_pay_by_user" Then
		    mystr = Array("service="&Q_Service,"partner="&Q_PartnerID,"subject="&Q_Subject,"body="&Q_Body,"out_trade_no="&Q_Out_Trade_No,"price="&Q_Price,"discount="&Q_Discount,"quantity="&Q_Quantity,"payment_type=1","seller_email="&Q_SellerEmail,"paymethod="&Q_Paymethod,"notify_url="&Q_NotifyUrl,"return_url="&Q_ReturnUrl,"_input_charset="&Q_Charset)
		Else
		    mystr = Array("service="&Q_Service,"partner="&Q_PartnerID,"subject="&Q_Subject,"body="&Q_Body,"out_trade_no="&Q_Out_Trade_No,"price="&Q_Price,"discount="&Q_Discount,"show_url="&Q_ShowUrl,"receive_name="&Q_buyer_name,"receive_address="&Q_buyer_address,"receive_zip="&Q_buyer_zip,"receive_phone="&Q_buyer_phone,"receive_mobile="&Q_buyer_mobile,"buyer_email="&Q_buyer_email,"quantity="&Q_Quantity,"payment_type=1","seller_email="&Q_SellerEmail,"paymethod="&Q_Paymethod,"notify_url="&Q_NotifyUrl,"return_url="&Q_ReturnUrl,"_input_charset="&Q_Charset,"logistics_type="&Q_Post_type,"logistics_fee="&Q_Post_fee,"logistics_payment="&Q_Post_payment,"logistics_type_1="&Q_Post_type_1,"logistics_fee_1="&Q_Post_fee_1,"logistics_payment_1="&Q_Post_payment_1,"logistics_type_2="&Q_Post_type_2,"logistics_fee_2="&Q_Post_fee_2,"logistics_payment_2="&Q_Post_payment_2)
		End If
		Sign = GetMySign(mystr)
		GetUrl	= "https://www.alipay.com/cooperate/gateway.do?" 
		For j = 0 To UBound(mystr) Step 1     
			val = Split(mystr(j), "=")
			If val(1) <> "" Then
			    If val(0) = "subject" Or val(0) = "body" Or val(0) = "receive_name" Or val(0) = "receive_address" Then
                	GetUrl= GetUrl & val(0) & "=" & server.URLEncode(val(1)) & "&"
				Else
                	GetUrl= GetUrl & mystr(j)&"&"
				End If
			End If
		Next
		GetUrl	= GetUrl&"sign="&Sign&"&sign_type=MD5"
	End Function

	'********************************
	'处理即时返回信息
	'********************************
	Public Function Return_Url()
		Dim varItem, mystr
		Return_Url = False
		For Each varItem in Request.QueryString
			mystr = varItem&"="&Request(varItem)&"^"&mystr
		Next 
		If mystr <> "" Then mystr = Left(mystr,Len(mystr)-1)
		mystr = Split(mystr, "^")
		If GetMySign(mystr) = Request.QueryString("sign") And CheckNotify(Request.QueryString("notify_id"))  Then Return_Url = True
	End Function

	'********************************
	'处理服务器通知消息
	'********************************
	Public Function Notify_Url()
		Dim varItem, mystr
		Notify_Url = False
		For Each varItem in Request.Form
			mystr = varItem&"="&Request.Form(varItem)&"^"&mystr
		Next 
		If mystr <> "" Then mystr = Left(mystr,Len(mystr)-1)
		mystr = Split(mystr, "^")
		If GetMySign(mystr) = Request.Form("sign") And CheckNotify(Request.Form("notify_id"))  Then Notify_Url = True
	End Function

	'********************************
	'验证消息是不是支付宝发送的
	'********************************
	Private Function CheckNotify(ByVal NotifyID)
		Dim AlipayNotifyURL, Retrieval
		AlipayNotifyURL = "http://notify.alipay.com/trade/notify_query.do?partner=" & Q_PartnerID & "&notify_id=" & NotifyID
		Set Retrieval = Server.CreateObject("Msxml2.ServerXMLHTTP.3.0")
		Retrieval.setOption 2, 13056 
		Retrieval.open "GET", AlipayNotifyURL, False, "", "" 
		Retrieval.send()
		CheckNotify = Retrieval.ResponseText
		Set Retrieval = Nothing
	End Function

	'********************************
	'根据参数数组获取签名的数据
	'********************************
	Private Function GetMySign(ByVal Ary)
		Dim CNum, i, j, minmax, minmaxSlot, mark, val, temp, md5str
		CNum = UBound(Ary)
		For i = CNum TO 0 Step -1
			minmax = Ary(0) : minmaxSlot = 0
			For j = 1 To i
				mark = (Ary(j) > minmax)
				If mark Then minmax = Ary(j) : minmaxSlot = j
			Next
			If minmaxSlot <> i Then 
        		temp = Ary(minmaxSlot)
        		Ary(minmaxSlot) = Ary(i)
        		Ary(i) = temp
    		End If
 		Next
		For j = 0 To CNum Step 1
			val = Split(Ary(j), "=")
			If val(1) <> "" And val(0) <> "sign" And val(0) <> "sign_type" Then
			    md5str = md5str&Ary(j)
				If j <> CNum Then md5str = md5str&"&"
			End If 
		Next
		GetMySign = Alipay_MD5(md5str&Q_SignCode,LCase(Q_Charset))
	End Function

	Public Property Let ShowUrl(ByVal vData)
		Q_ShowUrl = vData
	End Property
	Public Property Let receive_name(ByVal vData)
		Q_buyer_name = vData
	End Property
	Public Property Let receive_address(ByVal vData)
		Q_buyer_address = vData
	End Property
	Public Property Let receive_zip(ByVal vData)
		Q_buyer_zip = vData
	End Property
	Public Property Let receive_phone(ByVal vData)
		Q_buyer_phone = vData
	End Property
	Public Property Let receive_mobile(ByVal vData)
		Q_buyer_mobile = vData
	End Property
	Public Property Let buyer_email(ByVal vData)
		Q_buyer_email = vData
	End Property

	Public Property Let logistics_fee(ByVal vData)
		Q_Post_fee = vData
	End Property
	Public Property Let logistics_payment(ByVal vData)
		Q_Post_payment = vData
	End Property
	Public Property Let logistics_type(ByVal vData)
		Q_Post_type = vData
	End Property
	Public Property Let logistics_fee_1(ByVal vData)
		Q_Post_fee_1 = vData
	End Property
	Public Property Let logistics_payment_1(ByVal vData)
		Q_Post_payment_1 = vData
	End Property
	Public Property Let logistics_type_1(ByVal vData)
		Q_Post_type_1 = vData
	End Property
	Public Property Let logistics_fee_2(ByVal vData)
		Q_Post_fee_2 = vData
	End Property
	Public Property Let logistics_payment_2(ByVal vData)
		Q_Post_payment_2 = vData
	End Property
	Public Property Let logistics_type_2(ByVal vData)
		Q_Post_type_2 = vData
	End Property

	Public Property Let PartnerID(ByVal vData)
		Q_PartnerID = vData
	End Property
	Public Property Let SellerEmail(ByVal vData)
		Q_SellerEmail = vData
	End Property
	Public Property Let SignCode(ByVal vData)
		Q_SignCode = vData
	End Property
	Public Property Let NotifyUrl(ByVal vData)
		Q_NotifyUrl = vData
	End Property
	Public Property Let ReturnUrl(ByVal vData)
		Q_ReturnUrl = vData
	End Property
	Public Property Let Service(ByVal vData)
		Q_Service = vData
	End Property
	Public Property Let Charset(ByVal vData)
		Q_Charset = vData
	End Property

	Public Property Let Subject(ByVal vData)
		Q_Subject = vData
	End Property
	Public Property Let Body(ByVal vData)
		Q_Body = vData
	End Property
	Public Property Let Out_Trade_No(ByVal vData)
		Q_Out_Trade_No = vData
	End Property
	Public Property Let Price(ByVal vData)
		Q_Price = vData
	End Property
	Public Property Let Quantity(ByVal vData)
		Q_Quantity = vData
	End Property
	Public Property Let Discount(ByVal vData)
		Q_Discount = vData
	End Property
	Public Property Let Paymethod(ByVal vData)
		Q_Paymethod = vData
	End Property
End Class
%>
<%
Private Const BITS_TO_A_BYTE = 8
Private Const BYTES_TO_A_WORD = 4
Private Const BITS_TO_A_WORD = 32
Private m_lOnBits(30)
Private m_l2Power(30)
Private Function LShift(lValue, iShiftBits)
    If iShiftBits = 0 Then
        LShift = lValue
        Exit Function
    ElseIf iShiftBits = 31 Then
        If lValue And 1 Then
            LShift = &H80000000
        Else
            LShift = 0
        End If
        Exit Function
    ElseIf iShiftBits < 0 Or iShiftBits > 31 Then
        Err.Raise 6
    End If
    If (lValue And m_l2Power(31 - iShiftBits)) Then
        LShift = ((lValue And m_lOnBits(31 - (iShiftBits + 1))) * m_l2Power(iShiftBits)) Or &H80000000
    Else
        LShift = ((lValue And m_lOnBits(31 - iShiftBits)) * m_l2Power(iShiftBits))
    End If
End Function
Private Function str2bin(varstr)
     Dim i, varchar, varasc, varlow, varhigh
     str2bin = ""
     For i = 1 To Len(varstr)
         varchar = mid(varstr,i,1)
         varasc = Asc(varchar)
         If varasc < 0 Then varasc = varasc + 65535
         If varasc > 255 Then
            varlow = Left(Hex(Asc(varchar)),2)
            varhigh = right(Hex(Asc(varchar)),2)
            str2bin = str2bin & chrB("&H" & varlow) & chrB("&H" & varhigh)
         Else
            str2bin = str2bin & chrB(AscB(varchar))
         End If
     Next
End Function
Private Function str2bin_utf(varstr)
    Dim varchar, code, codearr, j, i
    str2bin_utf = ""
    For i=1 To Len(varstr)
        varchar = Mid(varstr,i,1)
        code = Server.UrlEncode(varchar)
		If(code="+") Then code="%20"
        If Len(code) = 1 Then
           str2bin_utf = str2bin_utf & chrB(AscB(code))
        Else
           codearr = Split(code,"%")
           For j = 1 to UBound(codearr)
              str2bin_utf = str2bin_utf & ChrB("&H" & codearr(j))
           Next
         End If
    Next
End Function
Private Function RShift(lValue, iShiftBits)
    If iShiftBits = 0 Then
        RShift = lValue
        Exit Function
    ElseIf iShiftBits = 31 Then
        If lValue And &H80000000 Then
            RShift = 1
        Else
            RShift = 0
        End If
        Exit Function
    ElseIf iShiftBits < 0 Or iShiftBits > 31 Then
        Err.Raise 6
    End If
    RShift = (lValue And &H7FFFFFFE) \ m_l2Power(iShiftBits)
    If (lValue And &H80000000) Then RShift = (RShift Or (&H40000000 \ m_l2Power(iShiftBits - 1)))
End Function
Private Function RotateLeft(lValue, iShiftBits)
    RotateLeft = LShift(lValue, iShiftBits) Or RShift(lValue, (32 - iShiftBits))
End Function
Private Function AddUnsigned(lX, lY)
    Dim lX4, lY4, lX8, lY8, lResult
    lX8 = lX And &H80000000
    lY8 = lY And &H80000000
    lX4 = lX And &H40000000
    lY4 = lY And &H40000000
    lResult = (lX And &H3FFFFFFF) + (lY And &H3FFFFFFF)
    If lX4 And lY4 Then
        lResult = lResult Xor &H80000000 Xor lX8 Xor lY8
    ElseIf lX4 Or lY4 Then
        If lResult And &H40000000 Then
            lResult = lResult Xor &HC0000000 Xor lX8 Xor lY8
        Else
            lResult = lResult Xor &H40000000 Xor lX8 Xor lY8
        End If
    Else
        lResult = lResult Xor lX8 Xor lY8
    End If
    AddUnsigned = lResult
End Function
Private Function md5_F(x, y, z)
    md5_F = (x And y) Or ((Not x) And z)
End Function
Private Function md5_G(x, y, z)
    md5_G = (x And z) Or (y And (Not z))
End Function
Private Function md5_H(x, y, z)
    md5_H = (x Xor y Xor z)
End Function
Private Function md5_I(x, y, z)
    md5_I = (y Xor (x Or (Not z)))
End Function
Private Sub md5_FF(a, b, c, d, x, s, ac)
    a = AddUnsigned(a, AddUnsigned(AddUnsigned(md5_F(b, c, d), x), ac))
    a = RotateLeft(a, s)
    a = AddUnsigned(a, b)
End Sub
Private Sub md5_GG(a, b, c, d, x, s, ac)
    a = AddUnsigned(a, AddUnsigned(AddUnsigned(md5_G(b, c, d), x), ac))
    a = RotateLeft(a, s)
    a = AddUnsigned(a, b)
End Sub
Private Sub md5_HH(a, b, c, d, x, s, ac)
    a = AddUnsigned(a, AddUnsigned(AddUnsigned(md5_H(b, c, d), x), ac))
    a = RotateLeft(a, s)
    a = AddUnsigned(a, b)
End Sub
Private Sub md5_II(a, b, c, d, x, s, ac)
    a = AddUnsigned(a, AddUnsigned(AddUnsigned(md5_I(b, c, d), x), ac))
    a = RotateLeft(a, s)
    a = AddUnsigned(a, b)
End Sub
Private Function ConvertToWordArray(sMessage)
    Dim lMessageLength, lNumberOfWords, lWordArray(), lBytePosition, lByteCount, lWordCount
    Const MODULUS_BITS = 512
    Const CONGRUENT_BITS = 448
    lMessageLength = LenB(sMessage)
    lNumberOfWords = (((lMessageLength + ((MODULUS_BITS - CONGRUENT_BITS) \ BITS_TO_A_BYTE)) \ (MODULUS_BITS \ BITS_TO_A_BYTE)) + 1) * (MODULUS_BITS \ BITS_TO_A_WORD)
    ReDim lWordArray(lNumberOfWords - 1)
    lBytePosition = 0
    lByteCount = 0
    Do Until lByteCount >= lMessageLength
        lWordCount = lByteCount \ BYTES_TO_A_WORD
        lBytePosition = (lByteCount Mod BYTES_TO_A_WORD) * BITS_TO_A_BYTE
        lWordArray(lWordCount) = lWordArray(lWordCount) Or LShift(AscB(MidB(sMessage, lByteCount + 1, 1)), lBytePosition)
        lByteCount = lByteCount + 1
    Loop
    lWordCount = lByteCount \ BYTES_TO_A_WORD
    lBytePosition = (lByteCount Mod BYTES_TO_A_WORD) * BITS_TO_A_BYTE
    lWordArray(lWordCount) = lWordArray(lWordCount) Or LShift(&H80, lBytePosition)
    lWordArray(lNumberOfWords - 2) = LShift(lMessageLength, 3)
    lWordArray(lNumberOfWords - 1) = RShift(lMessageLength, 29)
    ConvertToWordArray = lWordArray
End Function
Private Function WordToHex(lValue)
    Dim lByte, lCount
    For lCount = 0 To 3
        lByte = RShift(lValue, lCount * BITS_TO_A_BYTE) And m_lOnBits(BITS_TO_A_BYTE - 1)
        WordToHex = WordToHex & Right("0" & Hex(lByte), 2)
    Next
End Function
Public Function Alipay_MD5(sMessage, btype)
    m_lOnBits(0) = CLng(1)
    m_lOnBits(1) = CLng(3)
    m_lOnBits(2) = CLng(7)
    m_lOnBits(3) = CLng(15)
    m_lOnBits(4) = CLng(31)
    m_lOnBits(5) = CLng(63)
    m_lOnBits(6) = CLng(127)
    m_lOnBits(7) = CLng(255)
    m_lOnBits(8) = CLng(511)
    m_lOnBits(9) = CLng(1023)
    m_lOnBits(10) = CLng(2047)
    m_lOnBits(11) = CLng(4095)
    m_lOnBits(12) = CLng(8191)
    m_lOnBits(13) = CLng(16383)
    m_lOnBits(14) = CLng(32767)
    m_lOnBits(15) = CLng(65535)
    m_lOnBits(16) = CLng(131071)
    m_lOnBits(17) = CLng(262143)
    m_lOnBits(18) = CLng(524287)
    m_lOnBits(19) = CLng(1048575)
    m_lOnBits(20) = CLng(2097151)
    m_lOnBits(21) = CLng(4194303)
    m_lOnBits(22) = CLng(8388607)
    m_lOnBits(23) = CLng(16777215)
    m_lOnBits(24) = CLng(33554431)
    m_lOnBits(25) = CLng(67108863)
    m_lOnBits(26) = CLng(134217727)
    m_lOnBits(27) = CLng(268435455)
    m_lOnBits(28) = CLng(536870911)
    m_lOnBits(29) = CLng(1073741823)
    m_lOnBits(30) = CLng(2147483647)
    m_l2Power(0) = CLng(1)
    m_l2Power(1) = CLng(2)
    m_l2Power(2) = CLng(4)
    m_l2Power(3) = CLng(8)
    m_l2Power(4) = CLng(16)
    m_l2Power(5) = CLng(32)
    m_l2Power(6) = CLng(64)
    m_l2Power(7) = CLng(128)
    m_l2Power(8) = CLng(256)
    m_l2Power(9) = CLng(512)
    m_l2Power(10) = CLng(1024)
    m_l2Power(11) = CLng(2048)
    m_l2Power(12) = CLng(4096)
    m_l2Power(13) = CLng(8192)
    m_l2Power(14) = CLng(16384)
    m_l2Power(15) = CLng(32768)
    m_l2Power(16) = CLng(65536)
    m_l2Power(17) = CLng(131072)
    m_l2Power(18) = CLng(262144)
    m_l2Power(19) = CLng(524288)
    m_l2Power(20) = CLng(1048576)
    m_l2Power(21) = CLng(2097152)
    m_l2Power(22) = CLng(4194304)
    m_l2Power(23) = CLng(8388608)
    m_l2Power(24) = CLng(16777216)
    m_l2Power(25) = CLng(33554432)
    m_l2Power(26) = CLng(67108864)
    m_l2Power(27) = CLng(134217728)
    m_l2Power(28) = CLng(268435456)
    m_l2Power(29) = CLng(536870912)
    m_l2Power(30) = CLng(1073741824)
    Dim x, k, AA, BB, CC, DD, a, b, c, d
    Const S11 = 7
    Const S12 = 12
    Const S13 = 17
    Const S14 = 22
    Const S21 = 5
    Const S22 = 9
    Const S23 = 14
    Const S24 = 20
    Const S31 = 4
    Const S32 = 11
    Const S33 = 16
    Const S34 = 23
    Const S41 = 6
    Const S42 = 10
    Const S43 = 15
    Const S44 = 21
	If LCase(btype) = "utf-8" Then 
        x = ConvertToWordArray(str2bin_utf(sMessage))
	Else
        x = ConvertToWordArray(str2bin(sMessage))
	End If
    a = &H67452301
    b = &HEFCDAB89
    c = &H98BADCFE
    d = &H10325476
    For k = 0 To UBound(x) Step 16
        AA = a
        BB = b
        CC = c
        DD = d
        md5_FF a, b, c, d, x(k + 0), S11, &HD76AA478
        md5_FF d, a, b, c, x(k + 1), S12, &HE8C7B756
        md5_FF c, d, a, b, x(k + 2), S13, &H242070DB
        md5_FF b, c, d, a, x(k + 3), S14, &HC1BDCEEE
        md5_FF a, b, c, d, x(k + 4), S11, &HF57C0FAF
        md5_FF d, a, b, c, x(k + 5), S12, &H4787C62A
        md5_FF c, d, a, b, x(k + 6), S13, &HA8304613
        md5_FF b, c, d, a, x(k + 7), S14, &HFD469501
        md5_FF a, b, c, d, x(k + 8), S11, &H698098D8
        md5_FF d, a, b, c, x(k + 9), S12, &H8B44F7AF
        md5_FF c, d, a, b, x(k + 10), S13, &HFFFF5BB1
        md5_FF b, c, d, a, x(k + 11), S14, &H895CD7BE
        md5_FF a, b, c, d, x(k + 12), S11, &H6B901122
        md5_FF d, a, b, c, x(k + 13), S12, &HFD987193
        md5_FF c, d, a, b, x(k + 14), S13, &HA679438E
        md5_FF b, c, d, a, x(k + 15), S14, &H49B40821
        md5_GG a, b, c, d, x(k + 1), S21, &HF61E2562
        md5_GG d, a, b, c, x(k + 6), S22, &HC040B340
        md5_GG c, d, a, b, x(k + 11), S23, &H265E5A51
        md5_GG b, c, d, a, x(k + 0), S24, &HE9B6C7AA
        md5_GG a, b, c, d, x(k + 5), S21, &HD62F105D
        md5_GG d, a, b, c, x(k + 10), S22, &H2441453
        md5_GG c, d, a, b, x(k + 15), S23, &HD8A1E681
        md5_GG b, c, d, a, x(k + 4), S24, &HE7D3FBC8
        md5_GG a, b, c, d, x(k + 9), S21, &H21E1CDE6
        md5_GG d, a, b, c, x(k + 14), S22, &HC33707D6
        md5_GG c, d, a, b, x(k + 3), S23, &HF4D50D87
        md5_GG b, c, d, a, x(k + 8), S24, &H455A14ED
        md5_GG a, b, c, d, x(k + 13), S21, &HA9E3E905
        md5_GG d, a, b, c, x(k + 2), S22, &HFCEFA3F8
        md5_GG c, d, a, b, x(k + 7), S23, &H676F02D9
        md5_GG b, c, d, a, x(k + 12), S24, &H8D2A4C8A
        md5_HH a, b, c, d, x(k + 5), S31, &HFFFA3942
        md5_HH d, a, b, c, x(k + 8), S32, &H8771F681
        md5_HH c, d, a, b, x(k + 11), S33, &H6D9D6122
        md5_HH b, c, d, a, x(k + 14), S34, &HFDE5380C
        md5_HH a, b, c, d, x(k + 1), S31, &HA4BEEA44
        md5_HH d, a, b, c, x(k + 4), S32, &H4BDECFA9
        md5_HH c, d, a, b, x(k + 7), S33, &HF6BB4B60
        md5_HH b, c, d, a, x(k + 10), S34, &HBEBFBC70
        md5_HH a, b, c, d, x(k + 13), S31, &H289B7EC6
        md5_HH d, a, b, c, x(k + 0), S32, &HEAA127FA
        md5_HH c, d, a, b, x(k + 3), S33, &HD4EF3085
        md5_HH b, c, d, a, x(k + 6), S34, &H4881D05
        md5_HH a, b, c, d, x(k + 9), S31, &HD9D4D039
        md5_HH d, a, b, c, x(k + 12), S32, &HE6DB99E5
        md5_HH c, d, a, b, x(k + 15), S33, &H1FA27CF8
        md5_HH b, c, d, a, x(k + 2), S34, &HC4AC5665
        md5_II a, b, c, d, x(k + 0), S41, &HF4292244
        md5_II d, a, b, c, x(k + 7), S42, &H432AFF97
        md5_II c, d, a, b, x(k + 14), S43, &HAB9423A7
        md5_II b, c, d, a, x(k + 5), S44, &HFC93A039
        md5_II a, b, c, d, x(k + 12), S41, &H655B59C3
        md5_II d, a, b, c, x(k + 3), S42, &H8F0CCC92
        md5_II c, d, a, b, x(k + 10), S43, &HFFEFF47D
        md5_II b, c, d, a, x(k + 1), S44, &H85845DD1
        md5_II a, b, c, d, x(k + 8), S41, &H6FA87E4F
        md5_II d, a, b, c, x(k + 15), S42, &HFE2CE6E0
        md5_II c, d, a, b, x(k + 6), S43, &HA3014314
        md5_II b, c, d, a, x(k + 13), S44, &H4E0811A1
        md5_II a, b, c, d, x(k + 4), S41, &HF7537E82
        md5_II d, a, b, c, x(k + 11), S42, &HBD3AF235
        md5_II c, d, a, b, x(k + 2), S43, &H2AD7D2BB
        md5_II b, c, d, a, x(k + 9), S44, &HEB86D391
        a = AddUnsigned(a, AA)
        b = AddUnsigned(b, BB)
        c = AddUnsigned(c, CC)
        d = AddUnsigned(d, DD)
    Next
    Alipay_MD5 = LCase(WordToHex(a) & WordToHex(b) & WordToHex(c) & WordToHex(d))
End Function
%>