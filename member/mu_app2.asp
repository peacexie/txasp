<!--#include file="_config.asp"-->
<!--#include file="../upfile/sys/para/rkeyid.asp"-->
<!--#include file="../upfile/sys/pcfg/jmail.asp"-->
<!--#include file="../upfile/sys/para/keywords.asp"-->
<%  

Call sf_Guard()
Call App25Set()
goPage = Request("goPage")

rDir = Config_Path&"member/mu_"&AppRand12&".asp"
If Chk_URL3(rDir)="eUrl" Then
  Response.Redirect rDir&"?goPage="&goPage
End If

If SwhMemApp<>"Y" And Session("UsrID")&""="" Then 
  Response.Write "<br><br><center>"&vMem_RD21&"</center><br><br>"
  Response.End()
End If

sys27_RVal = Request(App27Random)&""
Dim sys27_Rnd(10)
If sys27_RVal&"" = "" Then
  Response.End()
Else
  For i = 1 To 9
   sys27_Rnd(i)=Mid(sys27_RVal,i*6,Mid(sys27_RVal,i,1))
  Next
End If

ChkCode = uCase(RequestF(App21Code,"C",255))
App23Flag = App23Check(App23Date) 

MemID = LCase(RequestF(AppRMemID&sys27_Rnd(1),"C",128))
MemID = Replace(Server.HTMLEncode(MemID),"'","`")
MemPW = RequestF("MemPW"&sys27_Rnd(2),"C",36) :MemPWBak = MemPW
MemQu = RequestF("MemQu"&sys27_Rnd(3),"C",36)
MemAn = RequestF("MemAn"&sys27_Rnd(3),"C",36) :MemAnBak = MemAn
MemType = RequestF("MemType","C","24")
MemTyp2 = RequestF("MemTyp2","C","24")
MemTyp3 = RequestF("dType","C","24") ' 用于认证
MemGrade = RequestF("dGrade","C",255)
MemName = RequestF(AppRName&sys27_Rnd(4),"C",96)
MemNam2 = RequestF("MemNam2"&sys27_Rnd(4),"C",96)
MemCard = RequestF("MemCard","C","48")
MemBirth = RequestF("MemBirth","D","1900-12-31")
MemEmail = RequestF(app30Tab(9)&sys27_Rnd(5),"C","255")
MemTel = RequestF("MemTel"&sys27_Rnd(5),"C","48")
MemMobile = RequestF("MemMobile"&sys27_Rnd(5),"C","48")
App25AR = Request(App25A)
App25BR = Request(App25B)
'App26Str = ""
'App26Chk = ""
'App26Day = Get_AppDay()+47 '每天变化一次
'For i=1 To Len(App26Code)
  'App26Str = App26Str&Request(Mid(App26Code,i,1)) 'Chr(96+i)
  'App26Chk = App26Chk&Chr(App26Day+i)                   'Chr(96+i)
'Next
'Response.Write vbcrlf&"<br>"&App26Str&vbcrlf&"<br>"&App26Chk

MemFlag = SwhMemChk

If Session("ChkCode")&""="" OR ChkCode&""="" Then
   msg = vMem_AP2A1
ElseIf ChkCode <> Session("ChkCode") then
   msg = vMem_AP2A2
ElseIf MemID = MemPW then
   msg = vMem_AP2A3
ElseIf MemName = "" then
   msg = vMem_AP2A4
ElseIf Len(MemID)<4 Or Len(MemPW)<5 then
   msg = vMem_AP2A5
ElseIf Session("UsrID")&""="" And inStr(rKeyID&",",MemID&",")>0 then
   msg = vMem_AP2A6&rKeyID
ElseIf rs_Exist(conn,"SELECT * FROM [AdmUser"&Adm_aUser&"] WHERE UsrID='"&MemID&"' ")="YES" then
   msg = vMem_AP2A7
ElseIf Not App23Flag Then
   msg = vMem_AP2AA

ElseIF App25AR="" Or App25AR<>App25D Then
   msg = vMem_AP2AB
ElseIF App25BR="" Or App25BR<>App25C Then
   msg = vMem_AP2AC
'ElseIF App26Chk="" Or App26Str<>App26Chk Then
   'msg = vMem_AP2AD
   
ElseIF  Request("xxID")<>"(xxID)" Or Request("Name")<>"(Name)" Then
   msg = vMem_AP2AE
   Call Add_Log(conn,MemID,"xxID-Name[ERR]"&MemID,"(*|)Login",sql)
   
ElseIF Not Get_AHChk(Request(App28Code),30) Then 
   msg = "Timeout App28Code:"&Request(App28Code)
'ElseIf MBType="00Public" OR MBGrade="01.Error" then
  'Call Add_Log(con00,MBID,"申请帐号NG","mu_app2.asp","Pass:"&RequestS("MBPW",3,255))
  'Response.End()
'ElseIf MBType="ComAgency, ComAgency, ComBBS, ComBlog, ComFrieds, ComTrade" then
  '
ElseIf Get_A30Chk(get30Time,get30TSN,60,"")<>"" Then
    msg = "错误！超时错误 或 安全认证错误！"
	
'ElseIf MemEmail="" Then
    'msg = "错误！超时错误 或 安全认证错误！"

Else
  
   sqlE = "SELECT MemID FROM [Member"&Mem_aMemb&"] WHERE MemID='"&MemID&"' "
 If rs_Exist(conn,sqlE) = "EOF" Then
MemAn = MD5_Mem(MemAn&MemID)
MemPW = MD5_Mem(MemPW&MemID) 

C = ""
xtrakeTime = Request("_trakeTime")
xtrakeTime = DateDiff("s",xtrakeTime,Now())
sql = " INSERT INTO [Member"&Mem_aMemb&"] (" 
sql = sql& "  MemID" 
sql = sql& ", MemPW" 
sql = sql& ", MemQu" 
sql = sql& ", MemAn" 
sql = sql& ", MemType" 
sql = sql& ", MemTyp2" 
sql = sql& ", MemGrade" 
sql = sql& ", MemName" 
sql = sql& ", MemNam2" 
sql = sql& ", MemSex" 
sql = sql& ", MemCard" 
sql = sql& ", MemBirth" 
sql = sql& ", MemFrom" 
sql = sql& ", MemTel" 
sql = sql& ", MemMobile" 
sql = sql& ", MemEmail" 
sql = sql& ", MemExp" 
sql = sql& ", MemFlag" 
sql = sql& ", LogAddIP" 
sql = sql& ", LogAUser" 
sql = sql& ", LogATime" 
sql = sql& ", LogEditIP" 
sql = sql& ", LogEUser" 
sql = sql& ", LogETime" 
sql = sql& ")VALUES(" 
sql = sql& "  '" & MemID &"'" 
sql = sql& ", '" & MemPW &"'" 
sql = sql& ", '" & MemQu &"'" 
sql = sql& ", '" & MemAn &"'" 
sql = sql& ", '" & MemType &"'" 
sql = sql& ", '" & MemTyp2 &"'" 
sql = sql& ", '" & MemGrade &"'" 
sql = sql& ", '" & MemName &"'" 
sql = sql& ", '" & MemNam2 &"'" 
sql = sql& ", '" & RequestS("MemSex",C,2) &"'" 
sql = sql& ", '" & MemCard &"'" 
sql = sql& ", '" & MemBirth &"'" 
sql = sql& ", '" & RequestS("InfFrom",C,255) &"'" 
sql = sql& ", '" & MemTel &"'" 
sql = sql& ", '" & MemMobile &"'" 
sql = sql& ", '" & MemEmail &"'" 
sql = sql& ", '" & RequestS("MemExp","D","1900-12-31") &"'" 
sql = sql& ", '" & MemFlag & "'" 
sql = sql& ", '" & Get_CIP() &"'" 
sql = sql& ", '" & xtrakeTime & "'" 
sql = sql& ", '" & Now() &"'" 
sql = sql& ", '-'" 
sql = sql& ", '-'" 
sql = sql& ", '1900-12-31'" 
sql = sql& ")"
   Call rs_DoSql(conn,sql)
   Session("ChkCode") = ""
   Session(App22Code)=""
   Session("ChkApp24")=""
   msg = "OK"'申请成功！
	'/////////////////////////////
		'Session("FGRID")  = FGrade
		'Session("FPerm")  = gPermStrUser("com",Session("FGRID"))
		'Session("FMID")   = USID
	    'Session("FirmID") = FID		
	'/////////////////////////////
	'Call SndMailApp(MemID,MemPWBak,MemType,MemEmail)
 Else
   msg = vMem_AP2A8
 End If
End If

p = "MemID="&MemID&"&MemQu="&MemQu&"&MemAn="&MemAnBak&"&MemName="&MemName&"&MemNam2="&MemNam2&"&MemType="&MemType&"&MemTyp2="&MemTyp2&"&MemCard="&MemCard&"&MemFrom="&MemFrom&"&MemBirth="&MemBirth&"&MemTel="&MemTel&"&MemMobile="&MemMobile&"&MemEmail="&MemEmail&"&"

%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<HEAD>
<meta http-equiv="Pragma" content="no-cache">
<meta http-equiv="Expires" content="0">
<META http-equiv=Content-Type content="text/html; charset=utf-8">
<TITLE><%=vMem_Btn01%>-<%=vPMsg_WName%></TITLE>
<META http-equiv=Pragma content=no-cache>
<link rel="stylesheet" type="text/css" href="../pfile/pimg/style.css">
</HEAD>
<BODY>
<%TPName=vMem_Btn01%>
<!--#include file="_ftop2.asp"-->
<!-- // Content Start -->
<div class="line18">&nbsp;</div>
        <TABLE width="300" border=0 
                        align=center cellPadding=2 cellSpacing=1 bgcolor="f0f0f0">
          <TBODY>
            <TR>
              <TD height="27" align="center" bgcolor="f8f8f8">&nbsp;</TD>
            </TR>
            <%If msg="OK" Then%>
            <TR>
              <TD align="center" bgcolor="f8f8f8" class="SysCont">&nbsp;<%=vMem_AP2B1%>
              <IFRAME name=SmsFrame src="../msg/out/smsapi.asp?Peace_Sms_RndID=<%=Timer%>&acu=SMember&MemID=<%=MemID%>" frameBorder=0 width="12" scrolling=no height="8"></IFRAME>
              <br>
              <%=vMem_AP2B2%> <a href="login.asp?MemID=<%=MemID%>&goPage=<%=goPage%>&verMemb=<%=verMemb%>"><font color="#0000FF"><%=vMem_AP2B3%></font></a> <%=vMem_AP2B4%></TD>
            </TR>
            <%Else%>
            <TR>
              <TD align="center" bgcolor="f8f8f8" class="SysCont"><font color="#FF0000"><%=Msg%></font> 
              <br>
              <%=vMem_AP2B2%> <a href='mu_<%=AppRandom%>.asp?<%=p%>goPage=<%=goPage%>&verMemb=<%=verMemb%>'><font color="#0000FF"><%=vMem_AP2B5%></font></a> <%=vMem_AP2B6%></TD>
            </TR>
            <%End If%>
            <TR align="right">
              <TD>&nbsp;</TD>
            </TR>
          </TBODY>
        </TABLE>
      <!--#include file="_fbot2.asp"-->
</BODY>
</HTML>
