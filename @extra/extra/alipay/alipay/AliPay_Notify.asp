<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="AliPay_Class.asp"-->
<!--#include file="../func1/func1.asp"-->
<!--#include file="../func2/func2.asp"-->

<%

Session.CodePage = 65001
conn = cora
Set rs = Server.CreateObject("Adodb.Recordset")
Dim MemID : MemID="(Alipay)"


If Request("ActTest")="(PeaceTest)" Then '// 用于测试

    out_trade_no = RequestS("id","C",255)
    If out_trade_no="" Then  out_trade_no = "MTCyyyyMMDD9999" '"MTC200912290008" '根据情况修改
    fPeace = out_trade_no&"_交易测试"
    returnTxt	= "success"
	'Call rs_DoSql(conn,"UPDATE PRD_COMBIN SET TOTAL_QUANTITY=400,ADULT_QUANTITY=1,CHILD_QUANTITY=1,BABY_QUANTITY=1 WHERE PRODUCT_COM_ID=132704")
	Response.Write returnTxt
	'Response.Write out_trade_no

Else

    Dim alipay, mystr, trade_status, returnTxt, out_trade_no
	Dim fPeace : fPeace = "(Null)"
	Set alipay = New Qlwz_AliPay

	If alipay.Notify_Url Then
		For Each varItem in Request.Form
			mystr = varItem&"="&Request.Form(varItem)&Chr(13)&mystr
		Next 
		trade_status = request.Form("trade_status")
		out_trade_no = Request.Form("out_trade_no")
		If trade_status = "WAIT_BUYER_PAY" Then
			fPeace = out_trade_no&"_等待买家付款"
			returnTxt	= "success"	
		ElseIf trade_status = "WAIT_SELLER_SEND_GOODS" Then      
			fPeace = out_trade_no&"_买家付款成功_等待卖家发货"
			returnTxt	= "success"
		ElseIf trade_status = "WAIT_BUYER_CONFIRM_GOODS" Then    
			fPeace = out_trade_no&"_卖家已发货等待买家确认"
			returnTxt	= "success"	
		ElseIf trade_status = "TRADE_FINISHED" Then             
			fPeace = out_trade_no&"_交易成功结束"
			returnTxt	= "success"		
		Else                                                     
			fPeace = out_trade_no&"_其他交易状态通知情况"
			returnTxt	= "success"
		End If
	Else
		fPeace = out_trade_no&"_交易失败"
		returnTxt	= "fail"
	End If
	Response.Write returnTxt
	Set alipay = Nothing

End If

	
'/////////////////////////////////////////////////////////////////////
'/////////////////////////////////////////////////////////////////////
If returnTxt = "success" Then

	 ' *** 0-A *** 取出信息 SALE_ORDER
	 sql = "SELECT ORDER_ID,CUSTOMER_ID,STAFF_ID FROM SALE_ORDER WHERE ORDER_NUMBER='"&out_trade_no&"' " 
	 rs.Open sql,conn,1,1
	 if not rs.EOF then
  	   OrdID = rs("ORDER_ID")
	   CusID = rs("CUSTOMER_ID")
  	   StfID = rs("STAFF_ID")
	 else
  	   OrdID = 0
	   CusID = 0
  	   StfID = 0
	 end if
	 rs.Close()
	 ' *** 0-B *** 取出信息 SALE_GROUP_ORDER  '多张子定单???
	 sql = "SELECT PRODUCT_ID,ADULT_QUANTITY,CHILD_QUANTITY,BABY_QUANTITY,RECEIVABLE FROM SALE_GROUP_ORDER WHERE SUBORDER_NUMBER='"&out_trade_no&"TT01' " 
	 rs.Open sql,conn,1,1
	 if not rs.EOF then
  	   CID = rs("PRODUCT_ID")
	   Q1 = rs("ADULT_QUANTITY")
	   Q2 = rs("CHILD_QUANTITY")
	   Q3 = rs("BABY_QUANTITY")
	   tFee = rs("RECEIVABLE")
	 else
  	   CID = 0
	   Q1 = 0
	   Q2 = 0
	   Q3 = 0
	   tFee = 0
	 end if
	 rs.Close()
	 'Response.Write "<br>OrdID,CID,tFee: "&OrdID&" "&CID&" "&tFee&" "
	 'Response.Write "<br>Q1,Q2,Q3: "&Q1&" "&Q2&" "&Q3&" "
	 'Response.Write "<br>CusID,StfID: "&CusID&" "&StfID&" "&xxx&" "
	 ' 2009-12-29 测试，UPDATE_DATE=to_date('"&Now()&"','yyyy-mm-dd hh24:mi:ss') 不能通过
	 ' *** 0-C *** 取出信息完毕 
	 
	 ' *** 1 *** sale_order
	 sql = "UPDATE sale_order SET "
	 sql = sql& " STATE=2"
	 sql = sql& ",ISWEBPAY=1"
	 sql = sql& ",UPDATE_DATE=to_date('"&Now()&"','yyyy-mm-dd hh24:mi:ss') "
	 sql = sql& " WHERE ORDER_NUMBER='"&out_trade_no&"' "
	 Call rs_DoSql(conn,sql) 
	 
	 ' *** 2 *** Sale_webpay
	 ' 然后在sale_webpay表中插入相应订购信息，如下图： 
	 ' Sale_webpay表结构如下：
	 ' -----------------------------------------------------------------------------------
	 ' 'F.Type 200      200             131      135       200            131         131      200 
	 ' 'F.Size 20       20              19       16        30             19          19       30 
	 ' 'Fields PRIKEYID ORDERNUMBER     PAYMONEY DATE_PAY  PAYFLAG        CUSTOMER_ID STAFF_ID PAYORG 
	 ' '649    00659    WEB200909280011 305.76   2009-9-28 已网上支付成功 560072      11627    支付宝 
	 ' '658    00668    20090929105525  500      2009-9-29 网上支付失败   0           0        支付宝 
	 MaxPK = getMaxPK()
	 sql = " INSERT INTO Sale_webpay (" 
	 sql = sql& "  PRIKEYID" 
	 sql = sql& ", ORDERNUMBER"
	 sql = sql& ", PAYMONEY" 
	 sql = sql& ", DATE_PAY" 
	 sql = sql& ", PAYFLAG" 
	 sql = sql& ", CUSTOMER_ID" 
	 sql = sql& ", STAFF_ID"
	 sql = sql& ", PAYORG" 
	 sql = sql& ")VALUES(" 
	 sql = sql& "  '" & MaxPK &"'"  '???
	 sql = sql& ", '" & out_trade_no &"'" 
	 sql = sql& ", " & tFee &"" 
	 sql = sql& ", to_date('"&Date()&"','yyyy-mm-dd')" 
	 sql = sql& ", '已网上支付成功'" 'PAYFLAG
	 sql = sql& ", " & CusID &"" 
	 sql = sql& ", " & StfID &"" 	
	 sql = sql& ", '支付宝'"
	 sql = sql& ")"
	 Call rs_DoSql(conn,sql) '/// ???  
	 
	 ' *** 3 *** Sale_group_order receivable=999.99,State=2,Update_date=Now()
	 sql = "UPDATE Sale_group_order SET "
	 sql = sql& " STATE=2" '客户确认？ 
	 sql = sql& ",Receivable="&tFee&" " 'Receivable,REALREC
	 sql = sql& ",UPDATE_DATE=to_date('"&Now()&"','yyyy-mm-dd hh24:mi:ss') "
	 sql = sql& " WHERE SUBORDER_NUMBER LIKE '"&out_trade_no&"%' " '多张子定单???
	 Call rs_DoSql(conn,sql) 
	 
	 ' *** 4 *** Prd_combin 更新字段：Adult_quantity，Child_quantity，Baby_quantity
	 fAdult = rs_Val(conn,"SELECT Adult_quantity FROM Prd_combin WHERE PRODUCT_COM_ID= "&CID&"","Adult_quantity")
	 If fAdult&""="" Then
	 sql = "UPDATE Prd_combin SET "
	 sql = sql& " Adult_quantity="&Q1&"" '
	 sql = sql& ",Child_quantity="&Q2&"" '
	 sql = sql& ",Baby_quantity="&Q3&"" '
	 sql = sql& " WHERE PRODUCT_COM_ID= "&CID&" " 
	 Else
	 sql = "UPDATE Prd_combin SET "
	 sql = sql& " Adult_quantity=Adult_quantity+"&Q1&"" '
	 sql = sql& ",Child_quantity=Child_quantity+"&Q2&"" '
	 sql = sql& ",Baby_quantity=Baby_quantity+"&Q3&"" '
	 sql = sql& " WHERE PRODUCT_COM_ID= "&CID&" " 
	 End If

	 Call rs_DoSql(conn,sql)
	 Call Add_Log(connBak,MemID,"支付成功:("&out_trade_no&")","Alipay",fPeace)
Else
	 Call Add_Log(connBak,MemID,"支付失败:("&out_trade_no&")","Alipay",fPeace)
End If
'/////////////////////////////////////////////////////////////////////
'/////////////////////////////////////////////////////////////////////


Function getMaxPK()
 Dim sql,idn : sql = "" '00659 ' 哈哈，大于 99999 怎么办？！
  sql = " SELECT Max(PRIKEYID) AS MaxPK FROM Sale_webpay "
  rs.Open sql,conn,1,1
  if not rs.EOF then
    idn = rs("MaxPK")&""
	If idn = "" Then
        idn = "00001"
	Else
	    idt = CLng(idn) + 1
	  If Len(idt)<5 Then
	    idn = Right("0000"&cStr(idt),5)
	  Else
	    idn = idt
	  End If
	End If
  else
    idn = "00001"
  end if
  rs.Close()
getMaxPK = idn
End Function
Set rs = Nothing

	Public Function xWriteToFile(fileName,Text)
		Dim Stream 
		Set Stream = Server.CreateObject("Adodb.stream")
		Stream.charset="utf-8"
		Stream.Type = 2
		Stream.Mode = 3
		Stream.open()
		Stream.WriteText(Text)
		Stream.SaveToFile Server.MapPath(fileName),2
		Stream.close()
		Set Stream = Nothing
	End Function
%>