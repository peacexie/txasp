<!--#include file="_config.asp"-->
<%

Call Chk_Perm1("Member","") 
ChkCode = Rnd_ID("KEY",8)
Session("ChkMember") = ChkCode 

MemID  = RequestS("MemID",3,255)
MemPWA = rs_Val("","SELECT MemPW FROM [Member"&Mem_aMemb&"] WHERE MemID='"&MemID&"'")

Response.Redirect "../member/login.asp?send=send&"&app30Tab(1)&"="&MemID&"&"&app30Tab(3)&"=[12345]&MemPWA="&MemPWA&"&ChkCode="&ChkCode&""&url30Par&""


'http://localhost:240/u/demo/member/login.asp?
'send=send&
'MemID7768R=demo&
'MemPWJEBHXRMF=[12345]&
'MemPWA=FD02061BAB061BAB05FD0205DE590BFA040C18AD0D0DF50D4B4705DE5905&
'ChkCode=
'SX6QYMGS&=588577768RSJEBHXRMF3J8C9QYU69P1WEPSMREN2H4NSB2YW3QU5XEY0XA10G4CR9U1KH2205
%>
