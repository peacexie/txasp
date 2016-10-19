<% 
rTime = Timer()

Set rs = Server.CreateObject("ADODB.Recordset") 'new ActiveXObject("ADODB.Connection");
rs.Open "select top 1 * from MemRRes", conn, 1 ,4

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

	'If rs_Exist(conn,"SELECT test_no FROM MemRRes WHERE test_no='"&test_no&"' AND item_no="&item_no&"")="EOF" Then
	Call rs_Dosql(conn,"DELETE FROM MemRRes WHERE test_no='"&test_no&"' AND item_no="&item_no&"")
 
	rs.AddNew()
	rs.Fields("test_no").Value    = test_no
	rs.Fields("item_no").Value    = item_no
	rs.Fields("print_order").Value  = print_order
	rs.Fields("report_item_name").Value  = report_item_name
	rs.Fields("report_item_code").Value = report_item_code
	
	rs.Fields("result").Value  = result
	rs.Fields("units").Value  = units
	rs.Fields("abnormal_indicator").Value = abnormal_indicator
	
	rs.Fields("instrument_id").Value  = instrument_id
	rs.Fields("result_date_time").Value  = result_date_time
	rs.Fields("print_context").Value = print_context
	
	rs.Update()
	
	  nOK=nOK+1
	  'Response.Write "<br>"&sql
	'Else
	  'fMsg=fMsg&"<br> --- 导入失败：["&test_no&"]-["&report_item_name&"]-重复"
	  'nRep=nRep+1
	'End If
 else
      ra = Split(fa(i),"	")
	  nSkip = nSkip+1
	  fMsg=fMsg&"<br> --- 忽略...：["&uBound(ra)&":"&fa(i)&"]"
 end if
 next
 '////// End Member
 
 rs.updateBatch()
 
rs.Close()
Set rs = Nothing
Response.Write Timer()-rTime
 
%>

    