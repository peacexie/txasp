<!--#include file="_config.asp"-->
<%

ID = RequestS("ID","C",48)
IP = Get_CIP()
send = Request.Form("send")



  rs.Open "SELECT * FROM [VoteInfo] WHERE KeyID='"&ID&"'",conn,1,1 
  if NOT rs.eof then 
KeyID = rs("KeyID") : js="VotCount=0;"
KeyMod = rs("KeyMod")
InfSubj = Show_Text(rs("InfSubj"))
InfRem1 = Show_Html(rs("InfRem1"))
InfTim1 = rs("InfTime1") :js=js&"VotTim1='"&InfTim1&"';"
InfTim2 = rs("InfTime2") :js=js&"VotTim2='"&InfTim2&"';"
InfCard = rs("InfCard")  :js=js&"InfCard='"&InfCard&"';" 'Y/N
InfVNum = rs("InfVNum")  :js=js&"InfVNum='"&InfVNum&"';" 'D1T1,DNT1,DXTN
InfNum1 = rs("InfNum1")  :js=js&"InfNum1="&InfNum1&";"
InfNum2 = rs("InfNum2")  :js=js&"InfNum2="&InfNum2&";"
ImgName = rs("ImgName")
  else
KeyID = ""  
  end if 
  rs.Close()
If KeyID = "" Then Response.Redirect("vlist.asp")

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML xmlns="http://www.w3.org/1999/xhtml">
<HEAD>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title>贺奖名单-<%=Config_Name%></title>
<style type="text/css">
<!--
.vH100 { line-height:100%; }
body {
	margin-left: 0px;
	margin-top: 8px;
	margin-right: 0px;
	margin-bottom: 8px;
}
body,td,th {
	font-size: 13px;
}
-->
</style>
<script src="vote.js" type="text/javascript"></script>
</HEAD>
<BODY>

                  <!--////////////////////////////////////////Start-->
                  <table width="720" border="1" align="center" cellpadding="0" cellspacing="5">
                    <form action="?" method="post" name="fm01">
                      <tr>
                        <td align="left"><FIELDSET style="margin:5px; padding:5px;">
                          <LEGEND align="center"> &nbsp;<b><%=InfSubj%></b> (贺奖名单) &nbsp; </LEGEND>
                          <%=InfRem1%>
                          </FIELDSET></td>
                      </tr>
                      <tr>
                        <td align="left" valign="bottom" style="padding:5px;"><table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#E0E0E0">
                            <tr>
                              <td align="center" nowrap bgcolor="#f0f0f0" style="padding-right:3px;">Top</td>
                              <td align="center" nowrap bgcolor="#f0f0f0">姓名</td>
                              <td align="center" nowrap bgcolor="#f0f0f0">电话</td>
                              <td align="center" nowrap bgcolor="#f0f0f0">身份证号</td>
                              <td align="center" nowrap bgcolor="#f0f0f0">IP</td>
                              <td align="center" nowrap bgcolor="#f0f0f0">时间</td>
                            </tr>

                            <%
	
sql = " SELECT * FROM [VoteLogs] WHERE KeyMod='"&ID&"' "
sql = sql & " ORDER BY InfLuck ASC,KeyID " 
rs.Open Sql,conn,1,1
i=0
Do While NOT rs.EOF
i=i+1
KeyID = rs("KeyID")
KeyMod = rs("KeyMod")
InfName = rs("InfName")
InfCard = Show_Tel(rs("InfCard"))
InfTel = Show_Tel(rs("InfTel"))
InfRate = rs("InfRate")
InfLuck = rs("InfLuck")&""
LogAddIP = rs("LogAddIP")
LogATime = rs("LogATime")
LogATime = Mid(LogATime,3,Len(LogATime)-5)
If NOT isNumeric(InfRate) Then InfRate=0
InfRate = FormatNumber(InfRate,4)
If InfLuck="0" Then
 sLuck = "<font color='#FF00FF'>特等奖</font>"
ElseIf InfLuck="1" Then
 sLuck = "<font color='#FF0000'>一等奖</font>"
ElseIf InfLuck="2" Then
 sLuck = "<font color='#CC0000'>二等奖</font>"
ElseIf InfLuck="3" Then
 sLuck = "<font color='#990000'>三等奖</font>"
ElseIf InfLuck="4" Then
 sLuck = "<font color='#999999'>参与奖</font>"
Else
 sLuck = "<font color='#CCCCCC'>"&InfLuck&"</font>"
End If
%>

                            <tr>
                              <td align="center" nowrap bgcolor="#FFFFFF">&nbsp;<%=sLuck%>&nbsp;</td>
                              <td align="center" nowrap bgcolor="#FFFFFF"><%=InfName%></td>
                              <td align="center" nowrap bgcolor="#FFFFFF"><%=InfTel%></td>
                              <td align="center" nowrap bgcolor="#FFFFFF"><%=InfCard%></td>
                              <td align="center" nowrap bgcolor="#FFFFFF"><%=LogAddIP%></td>
                              <td align="center" nowrap bgcolor="#FFFFFF"><%=LogATime%></td>
                            </tr>
                            <%
rs.MoveNext()
Loop
rs.Close()
%>
                        </table></td>
                      </tr>
                    </form>
</table>
<!--////////////////////////////////////////End-->

</BODY>
</HTML>
