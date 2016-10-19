<% 
rTime = Timer()

Dim bDel,bIns,kNO
bDel = ""
bIns = ""
kNO = 0



 '////// RepItem
 for i=1 To uBound(fa)
 ra = Split(fa(i),"	")
 if uBound(ra)>=10 then
 
	  test_no = RequestSafe(ra(0),"C",48)
	  item_no = RequestSafe(ra(1),"N",1)
	  print_order = RequestSafe(ra(2),"N",0)
	  report_item_name = RequestSafe(ra(3),"C",48)
	  report_item_code = RequestSafe(ra(4),"N",0)
	  result = RequestSafe(ra(5),"C",48)
	  units = RequestSafe(ra(6),"C",48)
	  abnormal_indicator = RequestSafe(ra(7),"C",48)
	  instrument_id = RequestSafe(ra(8),"C",48)
	  result_date_time = RequestSafe(ra(9),"D","1900-12-31")
	  print_context = RequestSafe(ra(10),"C",48)


  Dim cm :Set cm = Server.CreateObject("ADODB.Command")
  With cm
    .ActiveConnection = conn 
    .CommandText = "spImpRep3" '指定存储过程名
    .CommandType = 4 '表明这是一个存储过程
    .Prepared = true '要求将SQL命令先行编译 true false

	.Parameters.Append .CreateParameter("@test_no",200,1,48,test_no)
	.Parameters.Append .CreateParameter("@item_no",3,1,4,item_no)
	.Parameters.Append .CreateParameter("@print_order",3,1,4,print_order)
	.Parameters.Append .CreateParameter("@report_item_name",200,1,48,report_item_name)
	.Parameters.Append .CreateParameter("@report_item_code",3,1,4,report_item_code)
	.Parameters.Append .CreateParameter("@result",200,1,48,result)
	.Parameters.Append .CreateParameter("@units",200,1,48,units)
	.Parameters.Append .CreateParameter("@abnormal_indicator",200,1,48,abnormal_indicator)
	.Parameters.Append .CreateParameter("@instrument_id",200,1,48,instrument_id)
	.Parameters.Append .CreateParameter("@result_date_time",135,1,5,result_date_time)
	.Parameters.Append .CreateParameter("@print_context",200,1,48,print_context)
	
    .Execute()
	'.Parameters.Clear()
	
   
	'.Parameters.Clear();
  End With
  Set cm = Nothing
	
 else
      ra = Split(fa(i),"	")
	  nSkip = nSkip+1
	  fMsg=fMsg&"<br> --- 忽略...：["&uBound(ra)&":"&fa(i)&"]"
 end if
 next
 



 '////// End Member
Response.Write Timer()-rTime
%>

    