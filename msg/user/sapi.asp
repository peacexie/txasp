<!--#include file="_config.asp"-->
<!--#include file="../../sadm/func1/md5_func.asp"-->
<%


' ///////////// Paras
Dim tm,id,sn,tel,ct1,cs,nsn,act,re

tm = RequestS("tm","D","190-12-31")
id = RequestS("id","C",48)
sn = RequestS("sn","C",255) '(enc=md5(id+sn+tm+url))

'tel,ct1
cs  = RequestS("cs","C",12) '(gb2312,utf-8) 
nsn = RequestS("nsn","C",96)
act = Request("act") 'Error/Readme/Balance(Check)/EditSN/Send
                     'SendOut,SendODo
re  = RequestS("re","C",12) '返回:Xml(默认)/Report/goBack

' ///////////// Actions
If act="" Then 
  If Request("tel")>3 And Request("ct1")<>"" Then
    act = "SendMsg"
  ElseIf nsn<>"" Then
    act = "EditSN"
  Else
    act = "Error"
  End If
Else '+(Check)
  If act = "Readme" Then
    re = "Report" 
  ElseIf act="Send" Then
    'act="Send"
  ElseIf act="Balance" Then
    'act="Balance"
  ElseIf act="SendOut" Then
    'act="SendOut"
	re = "SOForm"
  ElseIf act="SendODo" Then
    'act="SendODo"
	re = "SOBack"
  Else
    act = "Error"
  End If
End If

' ///////////// Check SN...
Dim reState,reInt,reInfo :reState="" :reInt=0
Dim uPass,uUrl,uCharge,uMobile,uFlag,rs,sql,res
'echo act&tel&ct1 ':Response.End()

If smcDebug = "isDebug" Then '调试标记
  reState = "Debug"
  reInt = -40
  reInfo = "(Debug)接口目前正在调试，可能是对系统作一些小的调整升级等，请稍等一会再试!"  
  re = "Xml" 
ElseIf act = "Error" Then
  reState = "Error"
  reInt = -10
  reInfo = "Error Action!"
ElseIf Not outChkTime(act,tm) Then
  reState = "Error"
  reInt = -20
  reInfo = "Error Action!"
Else
  sql = "SELECT * FROM [SmsMember] WHERE MemID='"&id&"' "
  Set rs = Server.Createobject("Adodb.Recordset")
  rs.Open Sql,conn,1,1
  if NOT rs.EOF then
	uCode = rs("MemCode") '"0BA94C3D-9E07-00CA-1C16-9999999999
	uUrl =  rs("MemUrl") '"localhost:240/u/demo/msg/out_demo.asp"
	uBalance = RequestSafe(rs("MemBalance"),"N",0) 
	uMobile = rs("MemMobile") '""
	uFlag =  rs("MemFlag") '""
	reInt = outChkCode(sn,id,tm,uCode,uUrl,uFlag)
  else 'No Found User
    reInt = -31
  end if
  rs.Close()
  set rs = nothing
  If reInt < 0 Then
	reState = "Error"
	reInfo = "Error SN!"
  End If 
End If
'Response.Write "<br>aa3."&smcDebug
'Response.Write "<br>aa3."&reInt&uUrl&act

' ///////////// Error/Readme/Balance(Check)/EditSN/SendMsg
If reInt = 0 Then
  If act = "Readme" Then
	reState = "OK"
	reInt = 0
	reInfo = "Readme!"
  ElseIf act = "Balance" Then
	reState = "OK"
	reInt = uBalance
	reInfo = "Balance!"
  ElseIf act = "Send" Then 
    res = outSend(id,uBalance,"API")
	reState = res(0)
	reInt = res(1)
	reInfo = res(2)
  
  ElseIf act = "SendODo" Then 
	Call chkReload()
	res = outSend(id,uBalance,"Frm")
    msg = res(2)&" "&Now()

  ElseIf act = "EditSN" Then 
    res = outEdit(id,nsn)
	reState = res(0)
	reInt = res(1)
	reInfo = res(2)
  End If
End If

%>
<%
' Return .... 
reTime = Now() ':reInfo=reInfo&"测试..."
rePage = Request.Servervariables("HTTP_REFERER") ':Response.Write "<br>"&rePage
If inStr(rePage,"?")>0 Then rePage=Mid(rePage,1,inStr(rePage,"?")-1) ':Response.Write "<br>"&rePage
rePage = rePage&"?reAct=Return&reState="&reState&"&reInt="&reInt&"&reInfo="&outEncode(reInfo)&"&reTime="&reTime&""
%>
<%If re = "goBack" Then%>
<%Response.Redirect rePage%>
<%ElseIf re = "Report" Then%>
<!--#include file="../inc/sapi_res.asp"-->
<%ElseIf re = "SOForm" Or re = "SOBack" Then%>
<!--#include file="../inc/sapi_out.asp"-->
<%Else%>
<!--#include file="../inc/sapi_res.xml"-->
<%End If%>
