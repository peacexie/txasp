<table width="100%" border="0" cellpadding="1" cellspacing="1">
  <tr>
    <td width="30%" align="center"><strong><%=actName%></strong></td>
    <td width="10%" align="right">&nbsp;</td>
    <td align="center" nowrap="nowrap">
    <a href="?act=">发送短信</a>
     | <a href="?act=Group">短信群发</a>
	<%If sndType="(SmsAPI)" Then%>
     | <a href="?act=Test">发送测试</a>
    <%Else%>
    User: by (<%=sndUser%>)
    <%End If%>
    </td>
  </tr>
</table>
