
<%
If verNow="2" Then
  vPag_First = "First" '//////////////////
  vPag_Prev = "Prev"
  vPag_Next = "Next"
  vPag_Last = "Last"
  vPag_Recs = "Rec"
  vPag_Page = "Page"
Else
  vPag_First = "首页" '//////////////////
  vPag_Prev = "上页"
  vPag_Next = "下页"
  vPag_Last = "尾页"
  vPag_Recs = "记录"
  vPag_Page = "页码"
End If

Function RS_Page(xrs,xPage,xStr,xN) 
  Dim Fpag,Ppag,Npag,Lpag,Fpag2,Ppag2,Npag2,Lpag2,i
  Dim a0,a1,a2
Fpag ="<FONT face='Webdings'>9</FONT>"&vPag_First&""
Ppag ="<FONT face='Webdings'>7</FONT>"&vPag_Prev&""
Npag ="<FONT face='Webdings'>8</FONT>"&vPag_Next&""
Lpag ="<FONT face='Webdings'>:</FONT>"&vPag_Last&""
Fpag2=" <font class='SysPGary'><FONT face='Webdings'>9</FONT>"&vPag_First&"</font> "
Ppag2=" <font class='SysPGary'><FONT face='Webdings'>7</FONT>"&vPag_Prev&"</font> "
Npag2=" <font class='SysPGary'><FONT face='Webdings'>8</FONT>"&vPag_Next&"</font> "
Lpag2=" <font class='SysPGary'><FONT face='Webdings'>:</FONT>"&vPag_Last&"</font> "
Response.Write  "<TABLE width=100% cellpadding=0 cellspacing=0 border=0>"
Response.Write  "<TR><TD align=left nowrap class='SysPBar'>&nbsp; "
  	  If xrs.PageCount = 1 Then  
	Response.Write  Fpag2  
	Response.Write  Ppag2 
	Response.Write  Npag2
	Response.Write  Lpag2  
	  Elseif Int(xPage) = xrs.PageCount Then   ''' !!! int !!!
	Response.Write  " <A class='SysPBar' HREF='" &xStr& "&Page=1'>" &Fpag& "</A> "  
	Response.Write  " <A class='SysPBar' HREF='" &xStr& "&Page=" &int(xPage)-1& "'>" &Ppag& "</A> "
	Response.Write  Npag2	 
	Response.Write  Lpag2	  
	  Elseif int(xPage) = 1 Then 
	Response.Write  Fpag2 
	Response.Write  Ppag2
	Response.Write  " <A class='SysPBar' HREF='" &xStr& "&Page=" &int(xPage)+1&"'>" &Npag& "</A> "
	Response.Write  " <A class='SysPBar' HREF='" &xStr& "&Page=" &xrs.PageCount&"'>" &Lpag& "</A> "  	  
	  ELSE
	Response.Write  " <A class='SysPBar' HREF='" &xStr& "&Page=1'>" &Fpag&"</A>"
	Response.Write  " <A class='SysPBar' HREF='" &xStr& "&Page=" &int(xPage)-1 &"'>" &Ppag& "</A> "
	Response.Write  " <A class='SysPBar' HREF='" &xStr& "&Page=" &int(xPage)+1 &"'>" &Npag& "</A> "
	Response.Write  " <A class='SysPBar' HREF='" &xStr& "&Page=" &xrs.PageCount &"'>" &Lpag& "</A> "   
	  End If   
Response.Write "</TD>"
  IF xrs.PageCount > 1 Then 
    Response.Write "<form name=fmPDir action=" &xStr& " method=get><TD align=right nowrap width='1%' class='SysPBar'>"
	Response.Write "GO<input name='Page' type='text' size='5' maxlength='8' "
	Response.Write " onchange='submit()' value="&Page&" style='border:none; text-align:right;' class='SysPBar'> &nbsp;" 
    a0=Split(xStr&"?","?") '//////
    If Right(a0(1),1) <> "&" Then 
        a0(1) = a0(1) & "&"
    End If
    a1 = Split(a0(1),"&")
    For i = 0 To Ubound(a1)-1
        a2 = Split(a1(i)&"=","=")
        if a2(0) <> "Page" And Ubound(a2) = 2 Then 
            Response.Write "<input type=hidden name=" &a2(0)& " value=" &a2(1)& ">"   
        End If
    Next '////////////////////////////////////////////////////////////////////
    Response.Write "</TD></form>"
  End If
Response.Write "<TD align=right nowrap width='1%' class='SysPBar'>&nbsp;"
Response.Write "["&vPag_Page&":" &xPage& "/" &xRS.PageCount& "] &nbsp;"&vPag_Recs&":"&xrs.RecordCount&"&nbsp;</TD>" 
Response.Write "</TR></TABLE>"	 
End Function

Function RS_Page6(xPCount,xPage,xStr,xRecN) 
'[| Recordset PageBar v1.2; ---Peace(XieYongshun)2005/12/ |]
  Dim Fpag,Ppag,Npag,Lpag,Fpag2,Ppag2,Npag2,Lpag2,i
  Dim a0,a1,a2
Fpag ="<FONT face='Webdings'>9</FONT>首页"
Ppag ="<FONT face='Webdings'>7</FONT>上页"
Npag ="<FONT face='Webdings'>8</FONT>下页"
Lpag ="<FONT face='Webdings'>:</FONT>尾页"
Fpag2=" <font color=#999999><FONT face='Webdings'>9</FONT>首页</font> "
Ppag2=" <font color=#999999><FONT face='Webdings'>7</FONT>上页</font> "
Npag2=" <font color=#999999><FONT face='Webdings'>8</FONT>下页</font> "
Lpag2=" <font color=#999999><FONT face='Webdings'>:</FONT>尾页</font> "
Response.Write  "<TABLE width=100% cellpadding=0 cellspacing=0 border=0>"
Response.Write  "<TR><TD align=left nowrap>&nbsp; "
  	  If xPCount = 1 Then  
	Response.Write  Fpag2  
	Response.Write  Ppag2 
	Response.Write  Npag2
	Response.Write  Lpag2  
	  Elseif Int(xPage) = Int(xPCount) Then   ''' !!! int !!!
	Response.Write  " <A HREF='" &xStr& "&Page=1'>" &Fpag& "</A> "  
	Response.Write  " <A HREF='" &xStr& "&Page=" &int(xPage)-1& "'>" &Ppag& "</A> "
	Response.Write  Npag2	 
	Response.Write  Lpag2	  
	  Elseif int(xPage) = 1 Then 
	Response.Write  Fpag2 
	Response.Write  Ppag2
	Response.Write  " <A HREF='" &xStr& "&Page=" &int(xPage)+1&"'>" &Npag& "</A> "
	Response.Write  " <A HREF='" &xStr& "&Page=" &xPCount&"'>" &Lpag& "</A> "  	  
	  ELSE
	Response.Write  " <A HREF='" &xStr& "&Page=1'>" &Fpag&"</A>"
	Response.Write  " <A HREF='" &xStr& "&Page=" &int(xPage)-1 &"'>" &Ppag& "</A> "
	Response.Write  " <A HREF='" &xStr& "&Page=" &int(xPage)+1 &"'>" &Npag& "</A> "
	Response.Write  " <A HREF='" &xStr& "&Page=" &xPCount &"'>" &Lpag& "</A> "   
	  End If   
Response.Write "</TD>"
  IF xPCount > 1 Then 
    Response.Write "<form name=fmPDir action=" &xStr& " method=get><TD align=right nowrap width='1%'>"
	Response.Write "去<input name='Page' type='text' size='5' maxlength='8' "
	Response.Write " onchange='submit()' value="&Page&" style='xborder:none; text-align:right;'>页 &nbsp;" 
    a0=Split(xStr&"?","?") '//////
    If Right(a0(1),1) <> "&" Then 
        a0(1) = a0(1) & "&"
    End If
    a1 = Split(a0(1),"&")
    For i = 0 To Ubound(a1)-1
        a2 = Split(a1(i)&"=","=")
        if a2(0) <> "Page" And Ubound(a2) = 2 Then 
            Response.Write "<input type=hidden name=" &a2(0)& " value=" &a2(1)& ">"   
        End If
    Next '////////////////////////////////////////////////////////////////////
    Response.Write "</TD></form>"
  End If
Response.Write "<TD align=right nowrap width='1%'>&nbsp;"
Response.Write "[页码:" &xPage& "/" &xPCount& "] &nbsp;共"&CStr(xRecN)&"条记录&nbsp;</TD>" 
Response.Write "</TR></TABLE>"	 
End Function

%>
