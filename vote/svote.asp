<!--#include file="../page/_config.asp"-->
<%

ID = RequestS("ID","C",48)
IP = Get_CIP()
yAct = Request("yAct")
iRad = RequestS("iRad","C",48)

  rs.Open "SELECT * FROM [VoteInfo] WHERE KeyID='"&ID&"'",conn,1,1 
  if NOT rs.eof then 
KeyID = rs("KeyID")
KeyMod = rs("KeyMod")
InfSubj = Show_Text(rs("InfSubj"))
InfRem1 = Show_Html(rs("InfRem1"))
InfTim1 = rs("InfTime1")
InfTim2 = rs("InfTime2")
  else
KeyID = ""  
  end if 
  rs.Close()
If KeyID = "" Then Response.Redirect("slist.asp")


fVote = False 
If DateDiff("d",InfTim2,Date())>0 Then
   fVNot = "调查已结束!<br><a href='rview.asp?ID="&ID&"'>查看结果</a>"
   'Response.Redirect "vres.asp"
ElseIf DateDiff("d",Date(),InfTim1)>0 Then
   fVNot = "调查还未开始!"
Else
   fVNot = "调查进行中..." 
   fVote = True
   If yAct="Vote" And iRad<>"" AND Session("fVote")&""<>"NO" Then
      Call rs_DoSql(conn,"UPDATE VoteItem SET SetVote=SetVote+1 WHERE KeyID='"&iRad&"' AND KeyMod='"&ID&"' ")
	  Session("fVote") = "NO"
	  fVNot = fVNot& "感谢您投票！" 
   End If
End If
vSum = rs_Val("","SELECT SUM(SetVote) AS vSum FROM VoteItem WHERE KeyMod='"&ID&"'")
'Response.Write vSum


%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Pragma" content="no-cache">
<meta http-equiv="Expires" content="0">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title>网上调查-<%=vPMsg_WName%></title>
<link rel="stylesheet" type="text/css" href="../pfile/pimg/style.css">
<script src="../inc/home/jsInfo.js" type="text/javascript"></script>
</head>
<body>
<div style="padding:1px 8px">
  <!--////////////////////////////////////////Start主体-->
  <table width="98%" border="1" align="center" cellpadding="0" cellspacing="5">
    <form action="sview.asp" method="post" name="fm01">
      <tr>
        <td align="left"><FIELDSET style="margin:5px; padding:5px;">
            <LEGEND align="center"> &nbsp;<b><%=InfSubj%></b> (<%=InfTim1%> ~ <%=InfTim2%>) &nbsp; </LEGEND>
            <%
sql = " SELECT * FROM [VoteItem] WHERE KeyMod='"&ID&"' "
sql = sql & " ORDER BY SetTop,KeyID " 
rs.Open Sql,conn,1,1
i=0
Do While NOT rs.EOF
i=i+1
iKeyID = rs("KeyID")
iSubj = rs("InfSubj")

%>
            <input name="iVRad" type="radio" id="iVRad<%=i%>" value="<%=iKeyID%>">
            <%=i%>. <%=iSubj%> <br>
            <%
						   rs.MoveNext()
						   Loop
						   rs.Close()
						   
If Session("fVote") = "NO" Or fVote = False Then

Else

End If
						   
						   %>
          </FIELDSET></td>
      </tr>
      <tr>
        <td><TABLE width="98%" align="center">
            <%If Session("fVote")&"" <> "NO" OR fVote = True OR fVote = False Then%>
            <TR>
              <TD height="30" align="left" nowrap> 已有投票共：<%=vSum%>人次
                <input type="button" name="button" id="button" value="投票" <%=dis%> onclick="chkVData(this,'sview.asp?yAct=Vote&ID=<%=ID%>','Vote')">
                <input name="input" type="button" value="查看结果" onclick="chkVData(this,'sview.asp?ID=<%=ID%>','View')" />
                <input name="ID" type="hidden" id="ID" value="<%=ID%>" />
                <input name="yAct" type="hidden" id="yAct" value="Vote" /></TD>
            </TR>
            <%Else%>
            <TR align=middle>
              <TD height=30 align="center" bgcolor="ffffDD"><span style="float:right; padding-right:5%;">
                <input name="" type="button" value="查看结果" onclick="chkVData(this,'sview.asp?ID=<%=ID%>','View')" />
                </span>注意：<%=fVNot%> ,如果你已经投票，则不能再投票了！</TD>
            </TR>
            <%End If%>
          </TABLE></td>
      </tr>
    </form>
  </table>
  <!--////////////////////////////////////////End主体-->
</div>
<script src="svote.js" type="text/javascript"></script>
</BODY>
</HTML>
