<link rel="stylesheet" type="text/css" href="../pfile/pimg/style.css">
<TABLE cellSpacing=0 cellPadding=0 width="210" border=0>
  <TR>
    <TD 
          height=30 align=center background=../pfile/pimg/qqimg21_12_01.gif class="SysTitle">友情连接</TD>
  </TR>
  <TR>
    <TD height=120 align=left valign="top" style="padding:2px 0px 12px 0px;">

<div class='SysI01 SysI00'>---<a href="link.asp">友情连接</a>---</div>
<div class='SysI01 SysI00'>---<a href="map.asp">站点导航</a>---</div>


<%

   sMD = "HomLnk1"
   aCode = Split(Eval("s"&sMD&"Code"),"|")
   aName = Split(Eval("s"&sMD&"Name"),"|")
   aNam2 = Split(Eval("s"&sMD&"Nam2"),"|")
  For i=0 to uBound(aCode)
    If aCode(i)<>"" Then
      iFlag = aNam2(i)
%>          
        <div class='SysI01 SysI00'><a href="./?TypID=<%=aCode(i)%>" class="as01" ><%=aName(i)%></a></div>
<%

    End If
  Next
%> 


    </TD>
  </TR>
  <TR>
    <TD 
          height=30 align=center background=../pfile/pimg/qqimg21_12_01.gif class="SysTitle" style="BORDER-TOP: #b6d9f7 1px solid;">---</TD>
  </TR>
  <TR>
    <TD height=300 align=left valign="top" style="padding:2px 8px 12px 8px;" class="SysFoot">
    　　---。</TD>
  </TR>
</TABLE>
