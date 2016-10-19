<!--CSS:table.VoteMood X3-->
<script type="text/javascript">var ModID="<%=MD%>";var KeyID="<%=ID%>";var ModTab="<%=ModTab%>";</script>
<div class="line05">&nbsp;</div>
<TABLE border="0" align="center" cellPadding=1 cellSpacing=0 class="VoteMood">
  <TR>
    <th colspan="8" align=left scope=col>
       <div style="float:right">(<span id="vmMsg">选您认同的表情后才可显示结果</span>) </div>
       您看到此篇文章时的感受是：
      
      </th>
  </TR>
  <TR id="vmRateN" style="display:none;">
    <TD id="vmRate1" vAlign=bottom>&nbsp;</TD>
    <TD id="vmRate2" vAlign=bottom>&nbsp;</TD>
    <TD id="vmRate3" vAlign=bottom>&nbsp;</TD>
    <TD id="vmRate4" vAlign=bottom>&nbsp;</TD>
    <TD id="vmRate5" vAlign=bottom>&nbsp;</TD>
    <TD id="vmRate6" vAlign=bottom>&nbsp;</TD>
    <TD id="vmRate7" vAlign=bottom>&nbsp;</TD>
    <TD id="vmRate8" vAlign=bottom>&nbsp;</TD>
  </TR>
  <%If 1=2 Then%>
  <TR>
    <TD height=50><IMG height=40 src="../img/mood/A1.gif" width=40></TD>
    <TD><IMG height=40 src="../img/mood/A2.gif" width=40></TD>
    <TD><IMG height=40 src="../img/mood/A3.gif" width=40></TD>
    <TD><IMG height=40 src="../img/mood/A4.gif" width=40></TD>
    <TD><IMG height=40 src="../img/mood/A5.gif" width=40></TD>
    <TD><IMG height=40 src="../img/mood/A6.gif" width=40></TD>
    <TD><IMG height=40 src="../img/mood/A7.gif" width=40></TD>
    <TD><IMG height=40 src="../img/mood/A8.gif" width=40></TD>
  </TR>
  <TR>
    <TD>欠扁</TD>
    <TD>支持</TD>
    <TD>超赞</TD>
    <TD>难过</TD>
    <TD>搞笑</TD>
    <TD>扯淡</TD>
    <TD>不解</TD>
    <TD>头晕</TD>
  </TR>
  <%End If%>
  <TR>
    <TD height=50><IMG height=48 src="../img/mood/C1_icon.jpg" width=44></TD>
    <TD><IMG height=48 src="../img/mood/C2_icon.jpg" width=44></TD>
    <TD><IMG height=48 src="../img/mood/C3_icon.jpg" width=44></TD>
    <TD><IMG height=48 src="../img/mood/C4_icon.jpg" width=44></TD>
    <TD><IMG height=48 src="../img/mood/C5_icon.jpg" width=44></TD>
    <TD><IMG height=48 src="../img/mood/C6_icon.jpg" width=44></TD>
    <TD><IMG height=48 src="../img/mood/C7_icon.jpg" width=44></TD>
    <TD><IMG height=48 src="../img/mood/C8_icon.jpg" width=44></TD>
  </TR>
  <TR>
    <TD>支持</TD>
    <TD>愤怒</TD>
    <TD>无聊</TD>
    <TD>暴汗</TD>
    <TD>养眼</TD>
    <TD>炒作</TD>
    <TD>不解</TD>
    <TD>搞笑</TD>
  </TR>
  <%If 1=2 Then%>
  <TR>
    <TD height=50><IMG height=50 src="../img/mood/B1_jiong.gif" width=50></TD>
    <TD><IMG height=50 src="../img/mood/B2_tu.gif" width=50></TD>
    <TD><IMG height=50 src="../img/mood/B3_se.gif" width=50></TD>
    <TD><IMG height=50 src="../img/mood/B4_ku.gif" width=50></TD>
    <TD><IMG height=50 src="../img/mood/B5_bucuo.gif" width=50></TD>
    <TD><IMG height=50 src="../img/mood/B6_guanzhu.gif" width=50></TD>
    <TD>&nbsp;</TD>
    <TD>&nbsp;</TD>
  </TR>
  <TR>
    <TD>囧</TD>
    <TD>恶心</TD>
    <TD>期待</TD>
    <TD>难过</TD>
    <TD>不错</TD>
    <TD>关注</TD>
    <TD>&nbsp;</TD>
    <TD>&nbsp;</TD>
  </TR>
  <%End If%>
  <TR>
    <TD><INPUT id=vmItem1 onclick="vmSend(1);" type=radio></TD>
    <TD><INPUT id=vmItem2 onclick="vmSend(2);" type=radio></TD>
    <TD><INPUT id=vmItem3 onclick="vmSend(3);" type=radio></TD>
    <TD><INPUT id=vmItem4 onclick="vmSend(4);" type=radio></TD>
    <TD><INPUT id=vmItem5 onclick="vmSend(5);" type=radio></TD>
    <TD><INPUT id=vmItem6 onclick="vmSend(6);" type=radio></TD>
    <TD><INPUT id=vmItem7 onclick="vmSend(7);" type=radio></TD>
    <TD><INPUT id=vmItem8 onclick="vmSend(8);" type=radio></TD>
  </TR>
  <TR>
    <TD colspan="8" class="line05">&nbsp;</TD>
  </TR>
</TABLE>
<div class="line05">&nbsp;</div>
<script type="text/javascript" src="../tools/nmood/form.js"></script>
