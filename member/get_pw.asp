<!--#include file="_config.asp"-->
<%
TPName=vMem_PGetPW
goPage = Request("goPage")
send = Request.Form("send")
MemID = RequestS("MemID",3,128)
Method = Request("Method")


If Get_A30Chk(get30Time,get30TSN,60,"")<>"" Then '60min
  echo "31.Start!" :Response.End()
End If


Dim sys27_Rnd(10)
If send = "<>" Then
  Call Chk_URL()
  sys27_RVal = Request(App27Random)&""
  If sys27_RVal&"" = "" Then
    Response.End()
  Else
    For i = 1 To 9
     sys27_Rnd(i)=Mid(sys27_RVal,i*6,Mid(sys27_RVal,i,1))
    Next
  End If
Else
  sys27_RVal = Rnd_Base("5678",9)&Rnd_Base("",64)
  For i = 1 To 9
    sys27_Rnd(i)=Mid(sys27_RVal,i*6,Mid(sys27_RVal,i,1))
  Next
End If

%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<HEAD>
<META http-equiv=Content-Type content="text/html; charset=utf-8">
<TITLE><%=vMem_PGetPW%>-<%=vPMsg_WName%></TITLE>
<META http-equiv=Pragma content=no-cache>
<link rel="stylesheet" type="text/css" href="../pfile/pimg/style.css">
</HEAD>
<BODY>
<!--#include file="_ftop2.asp"-->
      <div class="line30">&nbsp;</div>
      <%If send="" Then%>
      <TABLE width="360" border=0 align=center cellPadding=2 cellSpacing=1>
        <FORM name='fpw01<%=sys27_RVal%>' action="?" method=post>
          <TR>
            <TD align="right"><%=vMem_PGetPW%>：</TD>
            <TD><%=vMem_PW01%> <%=vbcrlf&"<!--"&vbcrlf&Config_Code&"-->"&vbcrlf%>&nbsp;</TD>
          </TR>
          <TR>
            <TD align="right"><%=vMem_MemID%>：<BR></TD>
            <TD><INPUT name=MemID id="MemID" value="<%=MemID%>" size=24 maxLength=64 language=javascript></TD>
          </TR>
          <TR>
            <TD align="right">Method：</TD>
            <TD nowrap="nowrap"><select name="Method" id="Method">
              <option value="Email">Send Password to E-mail</option>
              <option value="Page">Get Password on Page</option>
            </select></TD>
          </TR>
          <TR>
            <TD width="25%" align="left"><input name="send" type="hidden" id="send" value="Step2">
            <input name="verMemb" type="hidden" id="verMemb" value="<%=verMemb%>" /></TD>
            <TD><input type="submit" name="Submit" value=" <%=vMem_Btn14%> ">
              &nbsp;
              <input id=Submit1 type=Reset value=" <%=vMem_Btn12%> " name=Submit1 >
            <input name='goPage' type='hidden' id='goPage' value='<%=goPage%>'></TD>
          </TR>
          <TR align="right">
            <TD colspan="2">&nbsp;</TD>
          </TR>
          <input name="<%=App27Random%>" type="hidden" id="<%=App27Random%>" value="<%=sys27_RVal%>">
          <input name="<%=App30Code%>0" type="hidden" id="<%=App30Code%>0" value="<%=app30Arr(0)%>" />
          <input name="<%=App30Code%>1" type="hidden" id="<%=App30Code%>1" value="<%=app30Arr(1)%>" />
          <input name="<%=App30Code%>2" type="hidden" id="<%=App30Code%>2" value="<%=app30Arr(2)%>" />
        </FORM>
      </TABLE>
      <%End If%>
      <%If send="Step2" Then%>
      <%
  sqlE = "SELECT MemID FROM [Member"&Mem_aMemb&"] WHERE MemID='"&MemID&"' "
  If rs_Exist(conn,sqlE)="EOF" Or Len(MemID)<4 Then
    Flg = "EOF"
  Else
    sql = "SELECT * FROM [Member"&Mem_aMemb&"] WHERE MemID='"&MemID&"' "
    rs.Open sql,conn,1,1
    MemQu   = rs("MemQu")
    MemName = rs("MemName")
    rs.Close            
    Flag = "Next"
  End If
%>
      <table width="360" border="0" align="center" cellpadding="2" cellspacing="1">
        <FORM name='fpw02<%=sys27_RVal%>' action="?" method=post>
          <tr>
            <td width="25%" align="right"><%=vMem_PGetPW%>：</td>
            <td align="left"><%=vMem_PW02%></td>
          </tr>
          <%If Flag = "Next" Then%>
          <tr>
            <td align="right"><%=vMem_MemID%>：<br /></td>
            <td align="left"><input name="MemID" id="MemID" value="<%=MemID%>" size="24" maxlength="64" /></td>
          </tr>
          <tr>
            <td align="right"><%=vMem_PW22%>：</td>
            <td align="left"><input name="MemName" type="text" id="MemName" value="<%=MemName%>" size="24" maxlength="24" /></td>
          </tr>
          <tr>
            <td align="right"><%=vMem_PW23%>：<br /></td>
            <td align="left"><input name="MemQu" id="MemQu" value="<%=MemQu%>" size="24" maxlength="120" /></td>
          </tr>
          <tr>
            <td align="right"><%=vMem_PW24%>：<br /></td>
            <td align="left"><input name="MemAn" id="MemAn" size="24" maxlength="120" /></td>
          </tr>
          <tr>
            <td align="right"><%=vMem_PW25%>：</td>
            <td align="left"><input name="MemBirth" id="MemBirth" size="12" maxlength="10" />
              (<%=vMem_PW21%>:1900-12-31)</td>
          </tr>
          <tr>
            <td align="right"><%=vMem_PW26%>：</td>
            <td align="left"><input name="MemEmail" type="text" id="MemEmail" value="<%=MemEmail%>" size="24" maxlength="120" /></td>
          </tr>
          <tr>
            <td align="right" bgcolor="ffffff">&nbsp; &nbsp;
              <input name="send" type="hidden" id="send" value="Step3" />
              <input name='chkCode' type='hidden' id='chkCode' value='<%=chkCode%>' /></td>
            <td align="left" bgcolor="ffffff"><input id="Submit2" type="button" value=" <%=vMem_Btn14%> " name="Submit2" onclick="chkData();" />
              &nbsp;
              <input id="Submit2" type="reset" value=" <%=vMem_Btn12%> " name="Submit2" />
            <input name='goPage2' type='hidden' id='goPage2' value='<%=goPage%>' /></td>
          </tr>
          <%Else%>
          <tr>
            <td align="right" bgcolor="ffffff"><%=vMem_PW04%>：</td>
            <td align="left" bgcolor="ffffff"><%=vMem_PW27%></td>
          </tr>
          <tr align="center">
            <td height="60"><input name="verMemb" type="hidden" id="verMemb" value="<%=verMemb%>" /></td>
            <td height="60" align="left"><font color="#FF0000"> [<%=MemID%>] <%=vMem_PW28%> </font> <br />
              <%=vMem_PW2A%> <a href="?goPage=<%=goPage%>&amp;verMemb=<%=verMemb%>"><%=vMem_PW2B%></a> <%=vMem_PW2C%> <a href="mu_<%=AppRand12%>.asp?goPage=<%=goPage%>&amp;verMemb=<%=verMemb%>"><%=vMem_PW2D%></a></td>
          </tr>
          <%End If%>
          <tr align="right">
            <td colspan="2" nowrap="nowrap"></td>
          </tr>
          <input name="<%=App27Random%>" type="hidden" id="<%=App27Random%>" value="<%=sys27_RVal%>">
          <input name="<%=App30Code%>0" type="hidden" id="<%=App30Code%>0" value="<%=app30Arr(0)%>" />
          <input name="<%=App30Code%>1" type="hidden" id="<%=App30Code%>1" value="<%=app30Arr(1)%>" />
          <input name="<%=App30Code%>2" type="hidden" id="<%=App30Code%>2" value="<%=app30Arr(2)%>" />
        </form>
      </table>
      <script type="text/javascript">

var fm = document.fpw02<%=sys27_RVal%>;

function chkData()
{   errflag=1
        for(i=0;i<1;i++)
        {  

		  if (fm.MemAn.value.length==0)
           { alert('<%=vMem_PW41%>');
             fm.MemAn.focus();
             errflag=0;
             break;
           }

		   
		  if (fm.MemBirth.value.length==0)
           { alert('<%=vMem_PW42%>');
             fm.MemBirth.focus();
             errflag=0;
             break;
           }
		   
        }
          if (errflag==1)
          {    fm.submit()
          }
}
                  </script>
      <%End If%>
      <%If send="Step3" Then%>
      <%

				  chkCode   = RequestS("chkCode",3,48)
				  gMemID    = RequestS("MemID",3,48)
				  gMemAn    = RequestS("MemAn",3,255)
				  gMemAnBak = gMemAn
				  gMemBirth = RequestS("MemBirth","D","1900-12-31")
				  gMemEmail = RequestS("MemEmail",3,255)
				  gMemAn    = MD5_Mem(gMemAn&gMemID) 

If Len(gMemAnBak)=0 Or Len(gMemBirth)="1900-12-31" Or Len(gMemEmail)=0 Or inStr(gMemEmail,"@")<=0 then
   msg = vMem_PW43
   Flag = "ERROR"
Else
				    sql = "SELECT * FROM [Member"&Mem_aMemb&"] WHERE MemID='"&gMemID&"' " 
				    rs.Open sql,conn,1,1
					if NOT rs.EOF then
					MemAn    = Trim(""&rs("MemAn"))
					MemBirth = RequestSafe(rs("MemBirth"),"D","1900-12-31")
					MemEmail = Trim(""&rs("MemEmail"))
					MemPWO   = rs("MemPW")
					MemPW    = Rnd_ID("KEY",6) 
					MemPWE   = MD5_Mem(MemPW&gMemID) 
					'Response.Write gMemID&MemPW
					else
					  Response.Redirect "get_pw.asp"
					end if
				    rs.Close               
				  'Response.Write gMemAn&":"&MemAn&"<br>"&gMemBirth&":"&MemBirth&"<br>"&gMemEmail&":"&MemEmail
				  If gMemAn=MemAn and Get_yyyymmdd(gMemBirth)=Get_yyyymmdd(MemBirth) and gMemEmail=MemEmail Then
					If Method="Email" Then
					  Call rs_DoSql(conn,"UPDATE [Member"&Mem_aMemb&"] SET MemPW='"&MemPWE&"' WHERE MemID='"&gMemID&"' ")
					  Call Add_Log(conn,gMemID,"取回密码成功","get_pw3.asp",sql)
					  Session("chkCode") = ""
				      Flag = "OE"
					  emTo = gMemID :If inStr(gMemID,"@")<=0 Then emTo = MemEmail
					  emSubj = "Get password"
					  emCont = "Hi, you get a new password in elifebike.com just now!  " '"&MemPW&"
					  emCont = emCont& "<br>User ID: (<b style='color:#F00'>"&gMemID&"</b>) "
					  emCont = emCont& "<br>New password is: (<b style='color:#F00'>"&MemPWBak&"</b>) [the letters or numbers in bracket]"
					  emCont = emCont& "<br><a href='http://www.elifebike.com/member/' target='_blank'>please login in</a> the web site and set a new one. "
					  Call Send_jMail("elifebike.com",emTo,emSubj,emCont,"gb2312")
					Else
					  Call rs_DoSql(conn,"UPDATE [Member"&Mem_aMemb&"] SET MemPW='"&MemPWE&"' WHERE MemID='"&gMemID&"' ")
					  Call Add_Log(conn,gMemID,"取回密码成功","get_pw3.asp",sql)
					  Session("chkCode") = ""
				      Flag = "OK"
					End If
					
				  Else
				    Flag = "ERROR"
					Call Add_Log(conn,gMemID,"取回密码失败","get_pw3.asp",sql)
					Session("chkCode") = ""
				  End If
End If
%>
      <table width="360" border="0" align="center" cellpadding="2" cellspacing="1">
        <tr>
          <td width="25%" align="right"><%=vMem_PGetPW%>：</td>
          <td align="left"><%=vMem_PW03%></td>
        </tr>
        <%If Flag = "OE" Then%>
        <tr>
          <td width="25%" align="right">OK：</td>
          <td>Please incept E-mail:<br />
            <%=emTo%></td>
        </tr>
        <%ElseIf Flag = "OK" Then%>
        <tr>
          <td width="25%" align="right"><%=vMem_PW04%>：</td>
          <td><%=vMem_PW34%></td>
        </tr>
        <tr>
          <td align="right"><%=vMem_MemID%>：<br /></td>
          <td><input name="MemID" id="MemID" value="<%=gMemID%>" size="24" maxlength="64" language="javascript" /></td>
        </tr>
        <tr>
          <td align="right"><%=vMem_MemPW%>：</td>
          <td><input name="MemPW" type="text" id="MemPW" value="<%=MemPW%>" size="24" maxlength="12" />
            (<%=vMem_PW33%>)</td>
        </tr>
        <tr>
          <td align="right"><%=vMem_PW31%>：</td>
          <td><input name="MemPWE" type="text" id="MemPWE" value="<%=MemPWE%>" size="24" maxlength="12" />
            Encode</td>
        </tr>
        <tr>
          <td align="right"><%=vMem_PW32%>：</td>
          <td><input name="MemPWO" type="text" id="MemPWO" value="<%=MemPWO%>" size="24" maxlength="12" />
            Org</td>
        </tr>
        <%Else%>
        <tr>
          <td width="25%" align="right"><%=vMem_PW04%>：</td>
          <td><%=vMem_PW36%></td>
        </tr>
        <tr align="center" bordercolor="#D4D0C8">
          <td height="60">&nbsp;</td>
          <td height="60" align="left"><font color="#FF0000">[<%=gMemID%>] <%=vMem_PW36%> </font> <br />
          <%=vMem_PW2A%> <a href="?goPage=<%=goPage%>&amp;verMemb=<%=verMemb%>"><%=vMem_PW2B%></a> <%=vMem_PW2C%> <a href="mu_<%=AppRand12%>.asp?goPage=<%=goPage%>&amp;verMemb=<%=verMemb%>"><%=vMem_PW2D%></a>
          <%=msg%>
          </td>
        </tr>
        <%End If%>
      </table>
      <%End If%>
      <!--#include file="_fbot2.asp"-->
</BODY>
</HTML>
