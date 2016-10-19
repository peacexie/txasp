<!--#include file="../../sadm/func1/func1.asp"-->
<!--#include file="../../sadm/func2/func2.asp"-->
<!--#include file="config.asp"-->
<!--#include file="../../sadm/func1/md5_func.asp"-->
<!--#include file="../../pfile/lang/vmemb.asp"-->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title>用户中心 (Member Center)</title>
<meta http-equiv="Pragma" content="no-cache">
<link href="../../inc/mem_inc/mem_style.css" rel="stylesheet" type="text/css">
<script src="../../sadm/func1/Func_JS.js" type="text/javascript"></script>
</head>
<body>



<%

send = Request("send")

if send = "send" then
If Session("chkCode")&"" <> Request("chkCode") then
  msg = vMDel_B1
Else

  MemPW0  = RequestS("PW0",3,48)
  MemPWE0 = MD5_Mem(MemPW0&Session("MemID")) 'Peace_Enc(MemPW0,Session("MemID")) 
  sql = "SELECT MemID FROM [Member"&Mem_aMemb&"] WHERE MemID='"&Session("MemID")&"' AND MemPW='"&MemPWE0&"'"
  If rs_Exist(conn,sql) = "YES" Then
    sql = "DELETE FROM [Member"&Mem_aMemb&"] WHERE MemID='"&Session("MemID")&"' AND MemPW='"&MemPWE0&"' "
	Call rs_DoSql(conn,sql)
	'// Address
	'Call rs_DoSql(conn,"DELETE FROM [BlgCompany] WHERE KeyID='"&Session("MemID")&"'")
	'Call rs_DoSql(conn,"DELETE FROM [BlgHome]    WHERE KeyID='"&Session("MemID")&"'")
	'Call rs_DoSql(conn,"DELETE FROM [JobResume]  WHERE KeyID='"&Session("MemID")&"'")
	'Call Add_Log(conn,Session("MemID"),"删除帐号","(*|)member_app2.asp",Session("MemID"))
	Session("chkCode") = ""
    Msg = vMDel_B2
		Response.Redirect "/member/login.asp?send=del"
  Else
    Msg = vMDel_B3
  End If
  
End If ' chk Code
End If ' Send

Session("chkCode") = Rnd_ID("KEY",6)

%>
<!--#include file="../../inc/mem_inc/mem_top.asp"-->
<br>
        <TABLE width="540" border=0 
                        align=center cellPadding=2 cellSpacing=1 bgcolor="ccccff">
          <FORM name=ff action="mu_del.asp" method=post>
            <TBODY>
              <TR background="/img/tool/top01.gif">
                <TD height="24" colspan="2" nowrap background="/img/tool/top01.gif"><strong><%=vMDel_A1%></strong></TD>
              </TR>
              <TR bgcolor="ffffff">
                <TD align="right"><%=vMDel_A2%>：<BR>
                </TD>
                <TD><INPUT name=MemID class=bu1 id="MemID" value="<%=Session("MemID")%>"  
                              size=24 
                               maxLength=15 readonly>                </TD>
              </TR>
              <TR bgcolor="ffffff">
                <TD align="right"><%=vMDel_A3%>：</TD>
                <TD><INPUT name=PW0 
                               type=password class=bu1 id="PW0" size=24 
                              maxLength=15 language=javascript>
                (<%=vMDel_A4%>) </TD>
              </TR>
              <TR bgcolor="ffffff">
                <TD align="right"><%=vMDel_A5%>：</TD>
                <TD><input name='chkCode' type='text' id="chkCode" size='12' maxlength='48'>
                <input type="submit" name="chkCode2" value="<%=Session("chkCode")%>"></TD>
              </TR>
              <TR bgcolor="ffffff">
                <TD width="25%" align="left"><input name="send" type="hidden" id="send" value="send">
                <input name="verMemb" type="hidden" id="verMemb" value="<%=verMemb%>"></TD>
                <TD><input id=Submit1 type=button value="<%=vMDel_A6%>" name=Submit1 onClick="chkData();">
&nbsp;
                <input id=Submit1 type=Reset value="<%=vMDel_A7%>" name=Submit1 > &nbsp;<strong>&nbsp;</strong> <font color="#FF0000"><%=Msg%></font> </TD>
              </TR>
            </TBODY>
          </FORM>
</TABLE>

        <script type="text/javascript">

function chkData()
{   errflag=1
        for(i=0;i<1;i++)
        {  

 var tmp = chkF_PW(document.ff.PW0,"<%=vMDel_C1%>")
 if ( (document.ff.PW0.value.length<4) || (tmp=="ER") )
   {   
     errflag=0; break;
   }

 if ( !(document.ff.chkCode.value==document.ff.chkCode2.value) )
   {   
     alert("<%=vMDel_C2%>");
     document.ff.chkCode.focus();
     errflag=0; break;
   }
       
        }
          if (errflag==1)
          {    document.ff.submit()
          }
}
</script>
<br>
<!--#include file="../../inc/mem_inc/mem_bot.asp"-->
</body>
</html>
