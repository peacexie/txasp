<link rel="stylesheet" type="text/css" href="../pfile/pimg/style.css">
<TABLE cellSpacing=0 cellPadding=0 width="210" border=0>
  <TR>
    <TD align=center width=150 background=../pfile/pimg/qqimg21_12_01.gif 
          height=30 class="SysTitle"><%=MDName%></TD>
    <TD background=../pfile/pimg/qqimg21_12_02.gif>&nbsp;</TD>
  </TR>
  <TR>
    <TD height=60 colSpan=2 align=left valign="top" style="padding:2px;">
	<%=GetItemLays(MD,"info.asp?ModID="&MD&"&TypMD="&TypMD&"&DepID="&DP&"&Flag="&Flag&"&",TypMD)%>
    </TD>
  </TR>
  <TR>
    <TD align=center width=150 background=../pfile/pimg/qqimg21_12_01.gif 
          height=30 style="BORDER-TOP: #b6d9f7 1px solid;" class="SysTitle">类别显示-List</TD>
    <TD align="right" nowrap="nowrap" 
      background=../pfile/pimg/qqimg21_12_02.gif style="BORDER-TOP: #b6d9f7 1px solid;">更多<IMG height=27 
            src="../pfile/pimg/qqimg21_13.gif" width=13 align=absMiddle></TD>
  </TR>
  <TR>
    <TD height=60 colSpan=2 align=left valign="top" style="padding:3px;">
      <%=GetItemList(MD,"info.asp?ModID="&MD&"&TypMD="&TypMD&"&DepID="&DP&"&Flag="&Flag&"&",TypMD)%>
    </TD>
  </TR>
  <TR>
    <TD align=center width=150 background=../pfile/pimg/qqimg21_12_01.gif 
          height=30 style="BORDER-TOP: #b6d9f7 1px solid;" class="SysTitle">类别显示-Tree</TD>
    <TD align="right" nowrap="nowrap" 
      background=../pfile/pimg/qqimg21_12_02.gif style="BORDER-TOP: #b6d9f7 1px solid;">更多<IMG height=27 
            src="../pfile/pimg/qqimg21_13.gif" width=13 align=absMiddle></TD>
  </TR>
  <TR>
    <TD height=60 colSpan=2 align=left valign="top" style="padding:3px;">
      <%=GetItemTree(MD,"info.asp?ModID="&MD&"&TypMD="&TypMD&"&Flag="&Flag&"&",TypMD)%>
      <script type="text/javascript">ChkTree('<%=TPLay%>');</script>
    </TD>
  </TR>
  <TR>
    <TD height=60 colSpan=2 align=left valign="top" style="padding:2px;" bgcolor="#F0F0F0">
	<div class='SysI01'><a href='pic.asp?ModID=PicS124&TypMD=&Flag=&TypID=S120062'>轮船航班信息</a></div>
	<div class='SysI02'><a href='pic.asp?ModID=PicS124&TypMD=&Flag=&TypID=S120066'>火车航班信息</a></div>
	<div class='SysI03'><a href='pic.asp?ModID=PicS124&TypMD=&Flag=&TypID=S120070'>飞机航班信息</a></div>
    
    <div class="line05">&nbsp;</div>
    </TD>
  </TR>
</TABLE>
