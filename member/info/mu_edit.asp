<!--#include file="../../sadm/func1/func1.asp"-->
<!--#include file="../../sadm/func2/func2.asp"-->
<!--#include file="../../sadm/func1/func_opt.asp"-->
<!--#include file="../admin/mconfig.asp"-->
<!--#include file="config.asp"-->
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
MemFrom = RequestS("MemFrom",C,255) 'RequestS("AddrA","C",12)&"-"&RequestS("AddrB","C",12)
sql = " UPDATE [Member"&Mem_aMemb&"] SET "  
sql = sql& " MemTyp2 = '" & RequestS("MemTyp2",C,24) &"'" 
sql = sql& ",MemName = '" & RequestS("MemName",C,48) &"'" 
sql = sql& ",MemNam2 = '" & RequestS("MemNam2",C,48) &"'" 
sql = sql& ",MemSex = '" & RequestS("MemSex",C,12) &"'" 
sql = sql& ",MemCard = '" & RequestS("MemCard",C,48) &"'" 
sql = sql& ",MemBirth = '" & RequestS("MemBirth","D","1900-12-31") &"'" 
sql = sql& ",MemFrom = '" & MemFrom &"'" 
sql = sql& ",MemMobile = '" & RequestS("MemMobile",C,48) &"'" 
sql = sql& ",MemTel = '" & RequestS("MemTel",C,48) &"'" 
sql = sql& ",MemEmail = '" & RequestS("MemEmail",C,255) &"'" 
sql = sql& ",LogEditIP = '" & Get_CIP() &"'" 
sql = sql& ",LogEUser = '" & Session("MemID") &"'" 
sql = sql& ",LogETime = '" & Now() &"'" 
sql = sql& " WHERE MemID='"&Session("MemID")&"' "
Call rs_DoSql(conn,sql) 
  msg = vMem_EdtC4
end if

SET rs=Server.CreateObject("Adodb.Recordset") 
rs.Open "SELECT * FROM [Member"&Mem_aMemb&"] WHERE MemID='"&Session("MemID")&"'",conn,1,1 
if NOT rs.eof then  
MemID = rs("MemID")
MemPW = rs("MemPW")
MemQu = rs("MemQu")
MemAn = rs("MemAn")
MemType = rs("MemType")
MemTyp2 = rs("MemTyp2")
MemGrade = rs("MemGrade")
MemName = rs("MemName")
MemNam2 = rs("MemNam2")
MemSex = rs("MemSex")
MemCard = rs("MemCard")
MemBirth = rs("MemBirth")
MemFrom = rs("MemFrom")
MemMobile = rs("MemMobile")
MemTel = rs("MemTel")
MemEmail = rs("MemEmail")
MemExp = rs("MemExp")
MemFlag = rs("MemFlag")
LogAddIP = rs("LogAddIP")
LogAUser = rs("LogAUser")
LogATime = rs("LogATime")
LogEditIP = rs("LogEditIP")
LogEUser = rs("LogEUser")
LogETime = rs("LogETime")
end if 
rs.Close()
SET rs=Nothing

'MemTypeOpt = Get_rsOpt(conn,"SELECT MTPID,MTPName FROM MemType WHERE MTPID LIKE 'Mod%'",MemType)
'MemTypeOpt = "<select id='MemType' name='MemType'>"&MemTypeOpt&"</select>"

%>

   <!--#include file="../../inc/mem_inc/mem_top.asp"-->
   <br>
        <table align="center" cellpadding="2" cellspacing="1" bgcolor="CCCCCC">
          <form name='fm01' method='post' action='?'>
            <tr bgcolor="#FFFFFF">
              <td rowspan="121" valign="top" nowrap><table border="0" cellspacing="0" cellpadding="5">
                  <tr>
                    <th background="/img/tool/top01.gif" bgcolor="#FFFFFF"><%=vMem_EdtA1%></th>
                  </tr>
                  <%If verMemb="2" Then%>
                  <tr>
                    <td><p>Notes for edit infomation:</p>
                        <p>1.User ID 4~24 bit letters or number;</p>
                        <p>2.Password 5~24 bit letters or number;</p>
                        <p>3. Password Distinguish case.</p>
                    </td>
                  </tr>
                  <%Else%>
                  <tr>
                    <td><p>会员资料修改注意事项：</p>
                        <p>帐号为4~24位字符，可包含数字,<br>
                字母及._-三个特殊符号；</p>
                        <p>密码为4~24位数字,字母,特殊字符，<br>
                但除 &amp;&quot;'&lt;&gt;=/\八个特殊符号；</p>
                        <p>密码区分大小写；<br>
                帐号中的大写字母自动转化为小写；</p>
                    </td>
                  </tr>
                  <%End If%>
              </table></td>
              <td align="right" nowrap bgcolor="#FFFFFF"><FONT color=#ff0000>*</FONT><%=vMem_EdtA2%>:</td>
              <td nowrap bgcolor="#FFFFFF"><input name='MemID' type='text' value="<%=MemID%>" size='24' maxlength='64' disabled> 
                (<%=vMem_EdtA4%>)
              </td>
            </tr>
          <tr>
            <td align="right" nowrap bgcolor="#FFFFFF"><%=vMem_EdtA3%>:</td>
            <td align="left" nowrap bgcolor="#FFFFFF">
              <%=Get_SOpt(mCfgCode,mCfgName,MemType,"Val")%>
             (<%=vMem_EdtA5%>)</td>
          </tr>
          <%If 1=2 Then%>
          <tr>
            <td colspan="2" bgcolor="#FFFFEE">公司----------------------------</td>
          </tr>
          <%End If%>
          <%If MemType="Corp" Then%>

          <tr>
            <td align="right" nowrap bgcolor="#FFFFFF"><%=vMem_InfA1%>:</td>
            <td align="left" nowrap bgcolor="#FFFFFF"><input name='MemName' type='text' id="MemName" value="<%=MemName%>" size='36' maxlength='48'></td>
          </tr>
          <tr>
            <td align="right" nowrap bgcolor="#FFFFFF"><%=vMem_InfA2%>:</td>
            <td align="left" nowrap bgcolor="#FFFFFF"><select name="MemTyp2" id="MemTyp2" style="width:150px;">
                <%=Get_rsOpt(conn,"SELECT TypID,TypName FROM WebType WHERE TypMod='Fields'",MemTyp2)%>
              </select></td>
          </tr>
          <tr>
            <td align="right" nowrap bgcolor="#FFFFFF"><%=vMem_InfA3%>:</td>
            <td align="left" nowrap bgcolor="#FFFFFF"><input name='MemCard' type='text' id="MemCard" value="<%=MemCard%>" size='24' maxlength='24'> 
              &nbsp;<%=vMem_InfAB%></td>
          </tr>
          <tr>
            <td align="right" nowrap bgcolor="#FFFFFF"><%=vMem_InfA4%>:</td>
            <td align="left" nowrap bgcolor="#FFFFFF"><input name='MemBirth' type='text' value="<%=MemBirth%>" size='24' maxlength='10'>
              (<%=vMem_InfAF%>:1900-12-31)</td>
          </tr>
          <tr>
            <td align="right" nowrap bgcolor="#FFFFFF"><%=vMem_InfA5%>:</td>
            <td align="left" nowrap bgcolor="#FFFFFF"><input name='MemNam2' type='text' id="MemNam2" value="<%=MemNam2%>" size='24' maxlength='12'>
              <select name="MemSex" id="MemSex" style="width:80 ">
                <%=Get_SOpt("F;M;X",vMem_InfD5,MemSex,"")%>
              </select></td>
          </tr>
          <%If 1=2 Then%>
          <tr>
            <td colspan="2" bgcolor="#FFFFEE">个人</td>
          </tr>
          <%End If%>
          <%ElseIf MemType="Privy" Then%>
          <input type='hidden' name='MemNam2' id="MemNam2" size='24' value="<%=MemNam2%>">
          <tr>
            <td align="right" nowrap bgcolor="#FFFFFF"><%=vMem_InfB1%>:</td>
            <td align="left" nowrap bgcolor="#FFFFFF"><input name='MemName' type='text' id="MemName" value="<%=MemName%>" size='24' maxlength='24'>
              <select name="MemSex" id="MemSex" style="width:80 ">
                <%=Get_SOpt("F;M;X",vMem_InfD5,MemSex,"")%>
            </select></td>
          </tr>
          <tr>
            <td align="right" nowrap bgcolor="#FFFFFF"><%=vMem_InfB2%>:</td>
            <td align="left" nowrap bgcolor="#FFFFFF"><input name='MemCard' type='text' id="MemCard" value="<%=MemCard%>" size='24' maxlength='24'></td>
          </tr>
          <tr>
            <td align="right" nowrap bgcolor="#FFFFFF"><%=vMem_InfB3%>:</td>
            <td align="left" nowrap bgcolor="#FFFFFF"><input name='MemBirth' type='text' value="<%=MemBirth%>" size='24' maxlength='10'>
              (<%=vMem_InfAF%>:1900-12-31)</td>
          </tr>
          <%If 1=2 Then%>
          <tr>
            <td colspan="2" bgcolor="#FFFFEE">机构</td>
          </tr>
          <%End If%>
          <%ElseIf MemType="Gov" Then%>
          <tr>
            <td align="right" nowrap bgcolor="#FFFFFF"><%=vMem_InfC1%>:</td>
            <td align="left" nowrap bgcolor="#FFFFFF"><input name='MemName' type='text' id="MemName" value="<%=MemName%>" size='36' maxlength='48'></td>
          </tr>
          <tr>
            <td align="right" nowrap bgcolor="#FFFFFF"><%=vMem_InfC2%>:</td>
            <td align="left" nowrap bgcolor="#FFFFFF"><input name='MemCard' type='text' id="MemCard" value="<%=MemCard%>" size='24' maxlength='24'>
            &nbsp;<%=vMem_InfAB%></td>
          </tr>
          <tr>
            <td align="right" nowrap bgcolor="#FFFFFF"><%=vMem_InfC3%>:</td>
            <td align="left" nowrap bgcolor="#FFFFFF"><input name='MemNam2' type='text' id="MemNam2" value="<%=MemNam2%>" size='24' maxlength='12'>
              <select name="MemSex" id="MemSex" style="width:80 ">
                <%=Get_SOpt("F;M;X",vMem_InfD5,MemSex,"")%>
              </select></td>
          </tr>
          <tr>
            <td align="right" nowrap bgcolor="#FFFFFF"><%=vMem_InfC4%>:</td>
            <td align="left" nowrap bgcolor="#FFFFFF"><input name='MemBirth' type='text' value="<%=MemBirth%>" size='24' maxlength='10'>
              (<%=vMem_InfAF%>:1900-12-31)</td>
          </tr>

          <%If 1=2 Then%>
          <tr>
            <td colspan="2" bgcolor="#FFFFEE">团体</td>
          </tr>
          <%End If%>
          <%ElseIf MemType="Org" Then%>
          <tr>
            <td align="right" nowrap bgcolor="#FFFFFF"><%=vMem_InfD1%>:</td>
            <td align="left" nowrap bgcolor="#FFFFFF"><input name='MemName' type='text' id="MemName" value="<%=MemName%>" size='36' maxlength='48'></td>
          </tr>
          <tr>
            <td align="right" nowrap bgcolor="#FFFFFF"><%=vMem_InfD2%>:</td>
            <td align="left" nowrap bgcolor="#FFFFFF"><input name='MemCard' type='text' id="MemCard" value="<%=MemCard%>" size='24' maxlength='24'>
            &nbsp;<%=vMem_InfAB%></td>
          </tr>
          <tr>
            <td align="right" nowrap bgcolor="#FFFFFF"><%=vMem_InfD3%>:</td>
            <td align="left" nowrap bgcolor="#FFFFFF"><input name='MemBirth' type='text' value="<%=MemBirth%>" size='24' maxlength='10'>
              (<%=vMem_InfAF%>:1900-12-31)</td>
          </tr>
          <tr>
            <td align="right" nowrap bgcolor="#FFFFFF"><%=vMem_InfD4%>:</td>
            <td align="left" nowrap bgcolor="#FFFFFF"><input name='MemNam2' type='text' id="MemNam2" value="<%=MemNam2%>" size='24' maxlength='12'>
              <select name="MemSex" id="MemSex" style="width:80 ">
                <%=Get_SOpt("F;M;X",vMem_InfD5,MemSex,"")%>
              </select></td>
          </tr>
          
          <%End If%>
          <%If 1=2 Then%>
          <tr>
            <td colspan="2" bgcolor="#FFFFEE">联系----------------------------</td>
          </tr>
          <%End If%>
            <tr>
              <td align="right" nowrap bgcolor="#FFFFFF"><%=vMem_EdtAA%>:</td>
              <td nowrap bgcolor="#FFFFFF"><input name='MemTel' type='text' id="MemTel" value="<%=MemTel%>" size='24' maxlength='24'></td>
            </tr>
            <tr>
              <td align="right" nowrap bgcolor="#FFFFFF"><%=vMem_EdtAB%>:</td>
              <td nowrap bgcolor="#FFFFFF"><input name='MemMobile' type='text' id="MemMobile" value="<%=MemMobile%>" size='24' maxlength='24'></td>
            </tr>
            <tr>
              <td align="right" nowrap bgcolor="#FFFFFF"><%=vMem_EdtAC%>:</td>
              <td nowrap bgcolor="#FFFFFF"><input name='MemEmail' type='text' value="<%=MemEmail%>" size='24' maxlength='120'>
              </td>
            </tr>


            <tr>
              <td align="right" nowrap bgcolor="#FFFFFF"><%=vMem_EdtAD%>:</td>
              <td nowrap bgcolor="#FFFFFF"><input name='MemFrom' type='text' id="MemFrom" value="<%=MemFrom%>" size='24' maxlength='120'></td>
            </tr>
            <tr>
              <td align="right" nowrap bgcolor="#FFFFFF"><input name='chkCode' type='hidden' id='chkCode' value='<%=chkCode%>'>
              </td>
              <td nowrap bgcolor="#FFFFFF"><input type='button' name='Submit' value='<%=vMem_EdtB1%>' onClick='chkData()'>
&nbsp;<a href="?verMemb=<%=verMemb%>"><%=vMem_EdtB2%></a> &nbsp;
        <input type='reset' name='Reset' value='<%=vMem_EdtB3%>'>
        <input name='send' type='hidden' id='send' value='send'>
        <input name="verMemb" type="hidden" id="verMemb" value="<%=verMemb%>">
 <font color="#FF00FF"><%=msg%></font>             </td>
            </tr>
          </form>
</table>


        <script type="text/javascript">


 function chkData()
 {
       var eflag = 0;
       for(ii=0;ii<1;ii++)
         {  ////////// //////////////// Srart For ////////////////
 if (document.fm01.MemName.value.length==0) 
   {   
     alert("<%=vMem_EdtC3%>"); 
     document.fm01.MemName.focus();
     eflag = 1; break;
   }
 tmv = chkF_Date(document.fm01.MemBirth,"<%=vMem_EdtC1%>");
 if (tmv=='ER') 
   { 
     eflag = 1; break;
   }
 tmv = chkF_Mail(document.fm01.MemEmail,"<%=vMem_EdtC2%>");
 if (tmv=='ER') 
   { 
     eflag = 1; break;
   }

         }  ////////// //////////////// End For //////////////////
         if (eflag==0){ document.fm01.submit(); }
}</script>
<br>
<!--#include file="../../inc/mem_inc/mem_bot.asp"-->
</body>
</html>
