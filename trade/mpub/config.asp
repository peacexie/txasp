<!--#include file="../../sadm/func1/func1.asp"-->
<!--#include file="../../sadm/func2/func2.asp"-->
<!--#include file="../../sadm/func1/func_file.asp"-->
<!--#include file="../../sadm/func1/func_opt.asp"-->
<%

If Request("ModID")<>"" Then
  ModID = RequestS("ModID","C",48)
  Session("ModID") = ModID
ElseIf Session("ModID")&""<>"" Then
  ModID = Session("ModID")
Else
  ModID = "TraA124"
End If

Call Chk_URL()
Call Chk_Perm2("","")

Function rs_TraTID() 
Dim tSql,tVal,sVal
  tVal = Get_AutoID(12)&Rnd_ID("KEY",3) 
  tSql = "SELECT ParCode FROM TradePara WHERE ParCode='"&tVal&"' "
  If rs_Exist(conn,tSql)="YES" Then
    rs_TraTID = rs_TraTID()
  End If
rs_TraTID = tVal
End Function

%>

