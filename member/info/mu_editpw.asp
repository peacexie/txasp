<!--#include file="../../sadm/func1/func1.asp"-->
<!--#include file="../../sadm/func2/func2.asp"-->
<!--#include file="../admin/mconfig.asp"-->
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
  MemQu  = RequestS("MemQu",3,255)
  MemAn  = RequestS("MemAn",3,255)
  MemAm  = RequestS("MemAm",3,255)
  MemPW0 = RequestS("PW0",3,48)
  MemPW1 = RequestS("PW1",3,48)
  MemAnEX = MD5_Mem(MemAn&Session("MemID"))  
  MemPWE0 = MD5_Mem(MemPW0&Session("MemID")) 
  MemPWE1 = MD5_Mem(MemPW1&Session("MemID"))  
  sql = "UPDATE [Member"&Mem_aMemb&"] SET "
sql = sql& " MemQu='"&MemQu&"' "
If MemAn<>"" And MemAn<>MemAm Then
sql = sql& ",MemAn='"&MemAnEX&"' "
Msg2 = vMPW_A1
Else
Msg2 = vMPW_A2
End If
sql = sql& ",MemPW='"&MemPWE1&"' "
sql = sql& ",MemType = '" & RequestS("MemType",3,24) &"' " 
'sql = sql& ",MemGrade = '" & MemGrade &"' " 
sql = sql& " WHERE MemID='"&Session("MemID")&"'"
  If Session("MemPW") = MemPWE0 Then
    Call rs_DoSql(conn,sql)
	Msg = vMPW_A3&Msg2
	'Response.Write js_Alert(vMPW_A4,"Redir","/member/login.asp?send=out") 
  Else
    Msg = vMPW_A5 '& 'Session("MemPW") &"-"& MemPWE0&MemPW0
  End If   
  'Response.Write "<br>"&MemAn&"-"&MemPWE0&"-"&MemPWE1
  'Response.Write "<br>"&MemAn&"-"&MemPW0&"-"&MemPW1&"-"&Session("MemID")
else
end if

sql = " SELECT * FROM [Member"&Mem_aMemb&"] WHERE MemID='"&Session("MemID")&"' "
SET rs=Server.CreateObject("Adodb.Recordset")  
rs.Open sql,conn,1,1
  MemQu = rs("MemQu")
  MemAn = rs("MemAn")
  MemPW = rs("MemPW")
  MemType = rs("MemType")
  Session("MemPW") = rs("MemPW")
rs.Close
set rs = nothing  

%>
        <!--#include file="../../inc/mem_inc/mem_top.asp"-->
        <br>
        <TABLE width="540" border=0 
                        align=center cellPadding=2 cellSpacing=1 bgcolor="ccccff">
          <FORM name=ff action="?" method=post>
            <TBODY>
              <TR>
                <TD height="24" colspan="2" nowrap background="/img/tool/top01.gif"><strong><%=vMPW_C1%> </strong></TD>
              </TR>
              <TR bgcolor="ffffff">
                <TD align="right"><%=vMPW_E1%>：<BR>
                </TD>
                <TD><INPUT name=MemID class=bu1 id="MemID" value="<%=Session("MemID")%>"  
                              size=24 
                               maxLength=15 readonly>                </TD>
              </TR>
              
          <tr>
            <td align="right" nowrap bgcolor="#FFFFFF"><%=vMPW_E2%>：</td>
            <td align="left" nowrap bgcolor="#FFFFFF">
              <select name="MemType" id="MemType" style="width:150px;">
                <option value="">(<%=vMPW_E8%>)</option>
                <%=Get_SOpt(mCfgCode,mCfgName,MemType,"")%>
              </select>
            </td>
          </tr>
              
              <TR bgcolor="ffffff">
                <TD align="right"><%=vMPW_E3%>：</TD>
                <TD><INPUT name=PW0 
                               type=password class=bu1 id="PW0" size=24 
                              maxLength=15 language=javascript>
                (<%=vMPW_C3%>) </TD>
              </TR>
              <TR bgcolor="ffffff">
                <TD align="right"><%=vMPW_E4%>：</TD>
                <TD><INPUT name=PW1 type=password id="PW1" size=24 maxLength=15 language=javascript>
                (<%=vMPW_C3%>) </TD>
              </TR>
              <TR bgcolor="ffffff">
                <TD align="right"><%=vMPW_E5%>：<BR>
                </TD>
                <TD><INPUT name=PW2 type=password id="PW2" size=24 maxLength=15 language=javascript>
                (<%=vMPW_C3%>) </TD>
              </TR>
              <TR bgcolor="ffffff">
                <TD align="right"><%=vMPW_E6%>：<BR>
                </TD>
                <TD><INPUT name=MemQu id="MemQu" value="<%=MemQu%>" size=24 maxLength=60 language=javascript>
                (<%=vMPW_C4%>) </TD>
              </TR>
              <TR bgcolor="ffffff">
                <TD align="right"><%=vMPW_E7%>：<BR>
                </TD>
                <TD><textarea name="MemAn" cols="24" rows="3" class="bu1" id="MemAn" onBlur="SetAn(this)" onFocus="ClrAn(this)" onClick="ClrAn(this)" language="javascript"><%=MemAn%></textarea>
                  <br>
                (<%=vMPW_C5%>) </TD>
              </TR>
              <TR bgcolor="ffffff">
                <TD width="25%" align="left"><input name="send" type="hidden" id="send" value="send">
                <input name="MemAm" type="hidden" id="MemAm" value="<%=MemAn%>">
                <input name="verMemb" type="hidden" id="verMemb" value="<%=verMemb%>"></TD>
                <TD><input id=Submit1 type=button value='<%=vMPW_D1%>' name=Submit1 onClick="chkData();">
&nbsp;
                <input id=Submit1 type=Reset value='<%=vMPW_D2%>' name=Submit1 > &nbsp;<strong>&nbsp;</strong>  </TD>
              </TR>
              <tr bgcolor="FFFFFF">
                <td colspan="2"><strong><%=vMPW_C2%></strong>: <font color="#FF0000"><%=Msg%> </font></td>
              </tr>
              
            </TBODY>
          </FORM>
</TABLE>

        <script type="text/javascript">

var defAn = "<%=MemAn%>";
function ClrAn(e)
{ 
  if(e.value==defAn){
    e.value = "";
  }
}
function SetAn(e)
{ 
  if(e.value.length==0){
   e.value = defAn;
  } else {
   if(e.value.length>18){
    e.value = e.value.substring(0,15);
    alert('18个字符以内！');
   }
  }
}

function chkData()
{   errflag=1
        for(i=0;i<1;i++)
        {  

 var tmp = chkF_PW(document.ff.PW1,"<%=vMPW_B1%>")
 if ( (document.ff.PW1.value.length<4) || (tmp=="ER") )
   {   
     errflag=0; break;
   }
		  if (document.ff.PW1.value!=document.ff.PW2.value)
           { alert('<%=vMPW_B2%>');
             document.ff.PW2.focus();
             errflag=0;
             break;
           }
		  if (document.ff.MemQu.value.length==0)
           { alert('<%=vMPW_B3%>');
             document.ff.MemQu.focus();
             errflag=0;
             break;
           }
		  /*
		  if (document.ff.MemAn.value.length==0)
           { alert('<%=vMPW_B4%>');
             document.ff.MemAn.focus();
             errflag=0;
             break;
           }*/

		  if (document.ff.PW1.value==document.ff.MemID.value)
           { alert('<%=vMPW_B5%>');
             document.ff.PW1.focus();
             errflag=0;
             break;
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
