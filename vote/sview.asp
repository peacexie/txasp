<!--#include file="../page/_config.asp"-->
<%

ID = RequestS("ID","C",48)
iRad = RequestS("iRad","C",48)
yAct = Request("yAct")

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
'If KeyID = "" Then Response.Redirect("slist.asp")


   If yAct="Vote" And iRad<>"" Then 'AND Session("fVote")&""<>"NO" 
      Call rs_DoSql(conn,"UPDATE VoteItem SET SetVote=SetVote+1 WHERE KeyID='"&iRad&"' AND KeyMod='"&ID&"' ")
	  Session("fVote") = "NO"
	  fVNot = fVNot& "感谢您投票！" 
   End If


vSum = rs_Val("","SELECT SUM(SetVote) AS vSum FROM VoteItem WHERE KeyMod='"&ID&"'")
If vSum="0" Then
  vSum = 1
End If

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
<!--////////////////////////////////////////Start主体-->
<div style="padding:8px;">
  <table width="100%" border="0" align="center" cellpadding="0" cellspacing="5">
    <form action="?" method="post" name="fm01">
      <tr>
        <td align="left"><FIELDSET style="margin:5px; padding:5px;">
            <LEGEND align="center"> &nbsp;<b><%=InfSubj%></b> (<%=InfTim1%> ~ <%=InfTim2%>) &nbsp; </LEGEND>
            说明：<%=fVNot%> 共有(<span class="fnt00F"><%=vSum%>人次</span>)调查记录
            。<span style="float:right; padding-right:5%;">
            <input name="" type="button" value="关闭" onclick="javascript:window.close();"/>
            <!--
          <input name="" type="button" value="返回调查页" onclick="javascript:location.href='svote.asp?ID=<%=ID%>';"/>
          -->
            </span>
          </FIELDSET></td>
      </tr>
      <tr>
        <td width="650" align="left" valign="bottom" style="padding:5px 2px 5px 5px;"><%
sql = " SELECT * FROM [VoteItem] WHERE KeyMod='"&ID&"' "
sql = sql & " ORDER BY SetTop,KeyID " 
rs.Open Sql,conn,1,1
i=0
Do While NOT rs.EOF
i=i+1
iKeyID = rs("KeyID")
iSubj = rs("InfSubj")
iVote = rs("SetVote")
iVote = RequestSafe(iVote,"N","0")
%>
          <div style="padding-bottom:12px;">
            <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
              <tr>
                <td align="left"><b><span id="Msg<%=iKeyID%>"><%=i%>. <%=iSubj%></span></b> <%=iImgExt%></td>
              </tr>
              <tr>
                <td><%
								iw1 = 100*(iVote/vSum)
								iw0 = FormatNumber(iw1,2)
								iw1 = Int(iw1)
								iw2 = 100-iw1
								%>
                  <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" style="margin-top:5px;">
                    <tr>
                      <td colspan="2"><span class="fnt00F"><%=iVote%>票 (<%=iw0%>%) </span></td>
                    </tr>
                    <tr>
                      <td width="<%=iw1%>%" background="../img/vote/vote52.gif"><img src="../img/vote/vote52.gif" width="1" height="15" align="absmiddle" /></td>
                      <td width="<%=iw2%>%" background="../img/vote/vote62.gif"><img src="../img/vote/vote62.gif" width="1" height="15" align="absmiddle" /></td>
                    </tr>
                  </table></td>
              </tr>
            </table>
          </div>
          <%
						   rs.MoveNext()
						   Loop
						   rs.Close()
						   %></td>
      </tr>
    </form>
  </table>
</div>
<!--////////////////////////////////////////End主体-->
</BODY>
</HTML>
