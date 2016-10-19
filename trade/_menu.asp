<table width="950" border="0" align="center" cellpadding="0" cellspacing="0" class="bgLine1">
  <tr>
    <td width="12" align="center">&nbsp;</td>
    <td height="33" align="left" style="text-align:center"><li class="SysVers" id="SysEng"><a href="corp.asp?UsrID=<%=US%>">公司首页</a></li>
      <li class="SysVers" id="SysEng"><a href="cinf.asp?UsrID=<%=US%>&ModID=TraA124">公司介绍</a></li>
      <li class="SysVers" id="SysEng"><a href="cinf.asp?UsrID=<%=US%>&ModID=TraT124">产品供求</a></li>
      <li class="SysVers" id="SysEng"><a href="cinf.asp?UsrID=<%=US%>&ModID=TraN124">行业新闻</a></li>
      <li class="SysVers" id="SysEng"><a href="cinf.asp?UsrID=<%=US%>&ModID=TraJ124">企业招聘</a></li>
      <li class="SysVers" id="SysEng"><a href="gbook.asp?UsrID=<%=US%>">在线留言</a></li>
      </td>
    <td width="12" align="center">&nbsp;</td>
    <td width="320" align="center" nowrap="nowrap"><table border="0" align="center" cellpadding="1" cellspacing="1">
        <form id="fms01" name="fms01" method="post" onsubmit="InnSearch(this)">
          <tr>
            <td nowrap="nowrap">站内搜索：</td>
            <td nowrap="nowrap"><input name="KW" type="text" class="fm" id="KW" size="12" /></td>
            <td nowrap="nowrap"><select name="TP" id="TP" style="width:90px;">
                <option value="trade.asp?m=s">[会员供求]</option>
                <option value="news.asp?m=s">[行业新闻]</option>
                <option value="jobs.asp?m=s">[企业招聘]</option>
              </select></td>
            <td nowrap="nowrap"><input type="submit" name="button" id="button" value="搜索" />
            <input name="ID" type="hidden" id="ID" value="<%=ID%>" /></td>
          </tr>
        </form>
      </table></td>
  </tr>
</table>
