<% 
rTime = Timer()
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
 
	  sql = " INSERT INTO MemRRes (" 

	  sql = sql& "  test_no"
	  sql = sql& ", item_no"
	  sql = sql& ", print_order"
	  sql = sql& ", report_item_name"
	  sql = sql& ", report_item_code"
	  sql = sql& ", result"
	  sql = sql& ", units"
	  sql = sql& ", abnormal_indicator"
	  sql = sql& ", instrument_id"
	  sql = sql& ", result_date_time"
	  sql = sql& ", print_context"

	  sql = sql& ")VALUES(" 
	  
	  sql = sql& "  '" & test_no &"'"
	  sql = sql& ", '" & item_no &"'"
	  sql = sql& ", '" & print_order &"'"
	  sql = sql& ", '" & report_item_name &"'"
	  sql = sql& ", '" & report_item_code &"'"
	  sql = sql& ", '" & result &"'"
	  sql = sql& ", '" & units &"'"
	  sql = sql& ", '" & abnormal_indicator &"'"
	  sql = sql& ", '" & instrument_id &"'"
	  sql = sql& ", '" & result_date_time &"'"
	  sql = sql& ", '" & print_context &"'"
	  
	  sql = sql& ")"
	  Call rs_Dosql(conn,sql)
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
Response.Write Timer()-rTime
%>

    