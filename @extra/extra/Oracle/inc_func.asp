<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%Response.Charset="UTF-8"%>
<!--#include file="../../script/sys/config/Config.asp"-->
<%

'CDataLinks = "Server=anytrvl2;Uid=anytravel;Pwd=password;" 
'CDataTypes = "142" 
cora = "Driver={Microsoft ODBC for oracle};"&CDataLinks
conn = cora
PassID = "A9989D9D"

Function Enc_Peace(xStr)   
Dim CfgID,s,c1,c2
   CfgID = "19790913" 
   For i=1 To Len(xStr)
     c1 = Mid(xStr,i,1)
	 c2 = Mid(CfgID,(i Mod 8)+1,1)
	 s = s&cStr(Hex(Asc(c1)+Asc(c2)))
   Next
Enc_Peace = s
End Function 

Function rs_Count(xconn,xTab)   
   Dim rsCnt,ObjCnt,sql 
   sql  = " SELECT COUNT(*) AS rsPubCountExp FROM "&xTab&" " 
   Set rsCnt = Server.Createobject("Adodb.Recordset")
   rsCnt.open sql,xconn
   ObjCnt = rsCnt("rsPubCountExp")
   rsCnt.Close()
   SET rsCnt = Nothing
   rs_Count = ObjCnt
End Function 

Function rs_Val(xconn,xSql,xOut) 
   Dim rsVal,ObjVal
   Set rsVal = Server.Createobject("Adodb.Recordset")
   rsVal.Open xSql,xconn,1,1
   IF rsVal.EOF THEN
      ObjVal = ""
   ELSE
      ObjVal = rsVal(xOut)
   End IF
   rsVal.Close()
   SET rsVal = Nothing 
   rs_Val = ObjVal
End Function

Function RequestS(xPName,xPType,xDefault) 
  RequestS = RequestSafe(Request(xPName),xPType,xDefault)
End Function 
Function RequestSafe(xPName,xPType,xDefault) 
  Dim PValue :PValue = xPName&""
  If xPType = "N" then ' N/C/D/ID/PW/EM
	If not isNumeric(PValue) then 
      PValue = xDefault ' -1;0;1
    End if 
  ElseIf xPType = "D" then 
	if ( NOT isDate(PValue) ) AND ( isDate(xDefault) ) then
      PValue = xDefault
    elseif ( NOT isDate(PValue) ) AND ( NOT isDate(xDefault) ) then 
      PValue = "1900-12-31"
	end if
  ELSE  'C Save
	PValue = LeftB(PValue,xDefault)
	PValue = Replace(PValue, "'", "''")		
  End if 
RequestSafe = PValue
End Function 

Function RS_PageThis(xrs,xPage) 
  Dim Fpag,Ppag,Npag,Lpag,Fpag2,Ppag2,Npag2,Lpag2,i
Fpag ="<FONT face='Webdings'>9</FONT>首页"
Ppag ="<FONT face='Webdings'>7</FONT>上页"
Npag ="<FONT face='Webdings'>8</FONT>下页"
Lpag ="<FONT face='Webdings'>:</FONT>尾页"
Fpag2=" <font color=#999999><FONT face='Webdings'>9</FONT>首页</font> "
Ppag2=" <font color=#999999><FONT face='Webdings'>7</FONT>上页</font> "
Npag2=" <font color=#999999><FONT face='Webdings'>8</FONT>下页</font> "
Lpag2=" <font color=#999999><FONT face='Webdings'>:</FONT>尾页</font> "
Response.Write  "<TABLE cellpadding=0 cellspacing=0 border=0>"
Response.Write  "<TR><TD align=left nowrap>&nbsp; "
  	  If xrs.PageCount = 1 Then  
	Response.Write  Fpag2  
	Response.Write  Ppag2 
	Response.Write  Npag2
	Response.Write  Lpag2  
	  Elseif Int(xPage) = xrs.PageCount Then   ''' !!! int !!!
	Response.Write  " <A onClick=""chPage('1')"" HREF='#'>" &Fpag& "</A> "  
	Response.Write  " <A onClick=""chPage('"&int(xPage)-1&"')"" HREF='#'>" &Ppag& "</A> "
	Response.Write  Npag2	 
	Response.Write  Lpag2	  
	  Elseif int(xPage) = 1 Then 
	Response.Write  Fpag2 
	Response.Write  Ppag2
	Response.Write  " <A onClick=""chPage('"&int(xPage)+1&"')"" HREF='#'>" &Npag& "</A> "
	Response.Write  " <A onClick=""chPage('"&xrs.PageCount&"')"" HREF='#'>" &Lpag& "</A> "  	  
	  ELSE
	Response.Write  " <A onClick=""chPage('1')"" HREF='#'>" &Fpag&"</A>"
	Response.Write  " <A onClick=""chPage('"&int(xPage)-1&"')"" HREF='#'>" &Ppag& "</A> "
	Response.Write  " <A onClick=""chPage('"&int(xPage)+1&"')"" HREF='#'>" &Npag& "</A> "
	Response.Write  " <A onClick=""chPage('"&xrs.PageCount&"')"" HREF='#'>" &Lpag& "</A> "   
	  End If   
Response.Write "</TD>"
Response.Write "<TD align=right nowrap> &nbsp;共"&xrs.RecordCount&"条记录&nbsp;</TD>" 
Response.Write "</TR></TABLE>"	 
End Function

%>

