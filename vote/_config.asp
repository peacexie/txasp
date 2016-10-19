<!--#include file="../sadm/func1/func1.asp"-->
<!--#include file="../sadm/func2/func2.asp"-->
<%

Set rs=Server.CreateObject("Adodb.Recordset")
Cfg_FTime = cfgTimeC '"#"

Function VotList(xMod,xTop,xTim,xLen)
'xTim:All,Now,Wait,End
Dim sqlK,s,sql,KeyID,InfSubj : s=""
If xTim="All" Then
 sqlK=" "
ElseIf xTim="Wait" Then
 sqlK=" "&Cfg_FTime&Date()&Cfg_FTime&"<InfTime1 "
ElseIf xTim="End" Then
 sqlK=" AND InfTime2<"&Cfg_FTime&Date()&Cfg_FTime&" "
Else 
 sqlK=" AND "&Cfg_FTime&Date()&Cfg_FTime&" BETWEEN InfTime1 AND InfTime2 "
End If
 sql = " SELECT TOP "&xTop&" * FROM [VoteInfo] "
 sql =sql& " WHERE KeyMod='"&xMod&"' " &sqlK& " ORDER BY InfTime1 DESC" 
 rs.Open Sql,conn,1,1 
 Do While NOT rs.EOF
KeyID = rs("KeyID")
InfSubj = rs("InfSubj")&""
InfSubj = Show_Text(Left(InfSubj,xLen))  'Left("dasd",3) '  '
'goUrl = "vote.asp"
'If DateDiff("d",InfTim2,Date())>0 Then
'goUrl = "vres.asp"
'End If
goUrl = "vote.asp"
If xMod="BBSVB24" Then
goUrl = "research.asp"
End If
s=s&"&nbsp;<span class='vSide'>&middot;<a href='"&goUrl&"?ID="&KeyID&"'>"&InfSubj&"</a></span><br>"
 rs.MoveNext()
 Loop
 rs.Close()
VotList = s
End Function


Function InsItem(xItems,xName,xCard)
Dim sql,aItem,i
 sql = " INSERT INTO [VoteLogs] (" 
 sql = sql& "  KeyID, KeyMod" 
 sql = sql& ", KeyItems, InfName, InfTel, InfCard, InfLuck" 
 sql = sql& ", LogAddIP, LogAUser, LogATime" 
 sql = sql& " )VALUES( "
 sql = sql& "  '" & Get_AutoID(24) &"'" 
 sql = sql& ", '" & ID &"'"  'ModID
 sql = sql& ", '" & xItems &"'" 
 sql = sql& ", '" & xName &"'" 
 sql = sql& ", '" & RequestS("InfTel",3,96) &"'" 
 sql = sql& ", '" & xCard &"', 'X'" 
 sql = sql& ", '" & IP &"'" 
 sql = sql& ", '" & Session("UsrID") &"'" 
 sql = sql& ", '" & Now() &"'" 
 sql = sql& ")"
 Call rs_DoSql(conn,sql) 
 aItem = Split(xItems,",")
 For i=0 To uBound(aItem)
   Call rs_DoSql(conn,"UPDATE VoteItem SET SetVote=SetVote+1 WHERE KeyID='"&aItem(i)&"' ") 
 Next
End Function


Function Show_Tel(xTel)
Dim n,s : s=xTel&"" : n=Len(s)
 If n<=6 Then 
  Show_Tel=s
 ElseIf n<15 Then '< 15:N-5,***,2
  Show_Tel=Left(s,n-5)&"***"&Right(s,2)
 Else             '>=15:N-6,***,2
  Show_Tel=Left(s,n-6)&"****"&Right(s,2)
 End If
End Function

%>

