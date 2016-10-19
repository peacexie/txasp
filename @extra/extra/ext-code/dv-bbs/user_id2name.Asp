<!--#include file="Inc/Conn.Asp"-->

<%

'Option Explicit
'Response.Buffer = True
'Dim Conn,Db,StartTime,Timeset,SqlNum
'Timeset=0
'StartTime = Timer()
'SqlNum=0
'Db="Data/yxbbs.mdb"

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

Function RequestSafe(xPName,xPType,xDefault) 
  Dim PValue 
  PValue = xPName&""
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

Dim cbbs,uid,sql,uname
cbbs = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source="&Server.MapPath("/bbs/"&Db)
uid = RequestSafe("id","N",0) 
sql = "SELECT name FROM [YX_User] WHERE ID="&uid&""
uname = rs_Val(cbbs,sql,"name")
Response.Redirect "http://www.mtc666.com/bbs/Profile.Asp?name="&uname&""
'SELECT TOP 5 * FROM [YX_User] WHERE 1=1 ORDER BY EssayNum DESC
'


%>