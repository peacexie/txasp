

<link href="inc/style.css" rel="stylesheet" type="text/css">
<TABLE cellSpacing=0 cellPadding=0 width="210" border=0>
  <TR>
    <TD align=center width=150 background='../pfile/pimg/qqimg21_12_01.gif'
          height=30 class="SysTitle"><%=MDName%></TD>
    <TD background='../pfile/pimg/qqimg21_12_02.gif'>&nbsp;</TD>
  </TR>
  <TR>
    <TD height=120 colSpan=2 align=left valign="top" style="padding:2px;">
	  <%
	  If US<>"" Then 
	  js = rs_Val("","SELECT ParRem FROM TradePara WHERE ParFlag='Tra_Typ' AND LogAUser='"&US&"' ")
	  'Response.Write js
	  %>
      <script type="text/javascript">
        <%=js%>
		try{
		var ModID = "<%=MD%>";
		var aCode = eval("s"+ModID+"Code").split("|");
		var aName = eval("s"+ModID+"Name").split("|");
		var aFlag = eval("s"+ModID+"Flag").split("|");
		for(i=0;i<aCode.length-1;i++)
		{
		  if(ModID!="TraA124") { flg = ""; }
		  else { flg = "&Flag="+aFlag[i]+""; }
		  document.write("<div class='SysI01 SysI00'><a href=cinf.asp?ModID="+ModID+"&TypID="+aCode[i]+"&UsrID=<%=US%>"+flg+">"+aName[i]+"</a></div>");
		}
		} catch (e) {}
      </script>
      <%
	  Else
		sql = "SELECT TypID,TypName FROM WebType WHERE TypMod='Fields' ORDER BY TypID "
		str = Get_rsOpt(conn,sql,"")
		str = Replace(str,"<option value=","<div class='SysI01 SysI00'><a href=info.asp?ModID="&MD&"&TypID=")
		str = Replace(str,"</option>","</a></div>")
		Response.Write str
	  End If
	  %>
      <div class="line05">&nbsp;</div>
    </TD>
  </TR>
  <TR>
    <TD align=center width=150 background='../pfile/pimg/qqimg21_12_01.gif'
          height=30 style="BORDER-TOP: #b6d9f7 1px solid;" class="SysTitle">测试信息</TD>
    <TD align="right" nowrap="nowrap" 
      background='../pfile/pimg/qqimg21_12_02.gif' style="BORDER-TOP: #b6d9f7 1px solid;">更多<IMG height=27 
            src="../pfile/pimg/qqimg21_13.gif" width=13 align=absMiddle></TD>
  </TR>
  <TR>
    <TD height=120 colSpan=2 align=left valign="top" style="padding:3px;"></TD>
  </TR>
</TABLE>
