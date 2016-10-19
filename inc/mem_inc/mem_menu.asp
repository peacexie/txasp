<link rel="stylesheet" href="../../inc/mem_img/default.css" type="text/css" />
<%
Call Chk_Perm9("","(Mem)")
%>
<div class="menu" id="TabPage3">
  <table width="190" border="0" align="center" cellpadding="0" cellspacing="0" style="height:20px;" id="TabPage2">
    <tr>
      <td align="center" valign="top" class="ULMenuC">&nbsp;</td>
      <td rowspan="18" valign="top" id="MnuList"><div id="left_menu_cnt">
          <ul id="dleft_tab1A1">
            <li style="text-align:center; font-weight:bold;"> &nbsp;  会员帐户资料  &nbsp; </li>
            <li><a href="../../member/info/?verMemb=<%=verMemb%>" target="UsrMain">后台首页</a>---<a href="../../member/info/mu_edit.asp" target="UsrMain">基本资料</a></li>
            <li><a href="../../member/info/mu_editpw.asp" target="UsrMain">修改密码</a>---<a href="../../member/info/mu_del.asp" target="UsrMain"><font color="#FF00FF">注销帐号</font></a></li>
            <li style="text-align:center; font-weight:bold;"> &nbsp;  帮助与退出  &nbsp; </li>
            <li> ------ <a href="../../tools/help/uhelp.asp" target="UsrMain">会员帮助</a> ------ </li>
            <li> ------ <a href="../../member/login.asp?send=out&verMemb=<%=verMemb%>" target="_top">安全退出</a> ------ </li>
          </ul>
          <ul id="dleft_tab1A2">
            <li style="text-align:center; font-weight:bold;"> &nbsp;  会员定单  &nbsp; </li>
            <li> ------ <a href="../../smod/info/ord_list.asp?PrmFlag=(Mem)" target="UsrMain">定单列表</a> ------ </li>
            <li style="text-align:center; font-weight:bold;"> &nbsp;  Sms  &nbsp; </li>
            <li><a href="../../msg/user/send.asp?PrmFlag=(Mem)" target="UsrMain">短信发送</a>---<a href="../../msg/user/send.asp?PrmFlag=(Mem)&ModID=userTels&act=Group" target="UsrMain">短信群发</a></li>
            <li><a href="../../msg/user/tels.asp?PrmFlag=(Mem)&ModID=userTels" target="UsrMain">(电话薄)</a>---<a href="../../msg/user/tset.asp?PrmFlag=(Mem)&ModID=userTels" target="UsrMain">群发组设置</a></li>
            <li><a href="../../msg/user/temp.asp?PrmFlag=(Mem)&ModID=userTemp" target="UsrMain">范本设置</a>---<a href="../../msg/user/type.asp?PrmFlag=(Mem)&ModID=userTels" target="UsrMain">类别设置</a></li>
          </ul>
          <ul id="dleft_tab1A6">
            <li style="text-align:center; font-weight:bold;"> &nbsp;  留言与笔记  &nbsp; </li>
            <li><a href="../../smod/gbook/info_list.asp?ModID=MemB324&PrmFlag=(Mem)" target="UsrMain">个人笔记</a>---<a href="../../smod/gbook/info_add.asp?ModID=MemB324&PrmFlag=(Mem)" target="UsrMain">增加笔记</a></li>
            <li> ------ <a href="../../smod/gbook/info_list.asp?ModID=MemB224&PrmFlag=(Mem)" target="UsrMain">系统通知</a> ------ </li>
            <li><a href="../../smod/gbook/info_list.asp?ModID=MemB124&PrmFlag=(Mem)" target="UsrMain">写给网管</a>---<a href="../../smod/gbook/info_add.asp?ModID=MemB124&PrmFlag=(Mem)" target="UsrMain">增加留言</a></li>
          </ul>
          <ul id="dleft_tab1B1">
            <li>---外部论坛</li>
            <li>---内部论坛</li>
            <li>:<a href="../../smod/info/info_add.asp?ModID=InfN124&PrmFlag=(Mem)" target="UsrMain">ADD</a></li>
            <li>:<a href="../../smod/info/info_list.asp?ModID=InfN124&PrmFlag=(Mem)" target="UsrMain">List</a></li>
          </ul>
          <ul id="dleft_tab1C1">
            <li style="text-align:center; font-weight:bold;"> &nbsp;  商务信息  &nbsp; </li>
            <li><a target='UsrMain' href='../../trade/mpub/info_corp.asp'>企业资料</a>--<a target='UsrMain' href='../../trade/mpub/set_temp.asp'>参数</a>--<a target='UsrMain' href='../../trade/mpub/set_type.asp?ModID=TraA124'>类别</a></li>
            <li><a target='UsrMain' href='../../smod/info/info_list.asp?ModID=TraA124&PrmFlag=(Mem)'>企业介绍</a>--<a target='UsrMain' href='../../smod/info/info_add.asp?ModID=TraA124&PrmFlag=(Mem)'>增加</a>--<a target='UsrMain' href='../../trade/mpub/set_type.asp?ModID=TraA124'>类别</a></li>
            <li><a target='UsrMain' href='../../smod/info/info_list.asp?ModID=TraN124&PrmFlag=(Mem)'>新闻中心</a>--<a target='UsrMain' href='../../smod/info/info_add.asp?ModID=TraN124&PrmFlag=(Mem)'>增加</a>--<a target='UsrMain' href='../../trade/mpub/set_type.asp?ModID=TraN124'>类别</a></li>
            <li><a target='UsrMain' href='../../smod/info/info_list.asp?ModID=TraT124&PrmFlag=(Mem)'>供求交易</a>-<a target='UsrMain' href='../../smod/info/info_add.asp?ModID=TraT124&PrmFlag=(Mem)'>增加</a>-<a target='UsrMain' href='../../trade/mpub/set_type.asp?ModID=TraT124'>类别</a></li>
            <li><a target='UsrMain' href='../../smod/info/info_list.asp?ModID=TraJ124&PrmFlag=(Mem)'>招聘信息</a>-<a target='UsrMain' href='../../smod/info/info_add.asp?ModID=TraJ124&PrmFlag=(Mem)'>增加</a>-<a target='UsrMain' href='../../trade/mpub/set_type.asp?ModID=TraJ124'>类别</a></li> 
            <li>

--            
<a target='UsrMain' href='../../trade/mpub/info_gbook.asp?ModID=TraG124'>在线留言</a>
--
<a target='UsrMain' href='../../trade/mpub/info_gbook.asp?ModID=TraR124'>新闻评论</a>
</li>
<li>
--
<a target='UsrMain' href='../../trade/mpub/info_gbook.asp?ModID=TraO124'>产品订购</a>
--
<a target='UsrMain' href='../../trade/mpub/info_gbook.asp?ModID=TraA124'>申请职位</a>
            
            </li> 
          </ul>
          <ul id="dleft_tab2A1">

            <li style="text-align:center; font-weight:bold;"> &nbsp;  Account Infomation  &nbsp; </li> 
            <li><a href="../../member/info/?verMemb=2"              target="UsrMain">User Center</a>
            <li><a href="../../member/info/mu_edit.asp?verMemb=2"   target="UsrMain"> Account Infomation </a></li>
            <li><a href="../../member/info/mu_editpw.asp?verMemb=2" target="UsrMain"> Change password </a></a></li>
            <li><a href="../../member/info/mu_del.asp?verMemb=2"    target="UsrMain"><font color="#FF00FF">Kill Account</font></a></li>
            
            <li style="text-align:center; font-weight:bold;"> &nbsp; Help and Exit  &nbsp; </li>
            <li> ------ <a href="../../tools/help/uhelp.asp?verMemb=2"      target="UsrMain">Member Help</a> ------ </li>
            <li> ------ <a href="../../member/login.asp?send=out&verMemb=2" target="_top">Member Exit</a> ------ </li>
            
          </ul>
          <ul id="dleft_tab2A2">
            <li style="text-align:center; font-weight:bold;"> &nbsp; Member Order  &nbsp; </li>
            <li> ------ <a href="../../smod/info/ord_list.asp?PrmFlag=(Mem)&verMemb=2" target="UsrMain">View Order List</a> ------ </li>

          </ul>
        </div>
        <IFRAME name=RefFrame src="../../tools/out/check.asp" frameBorder=0 width="2" scrolling=no height="10"></IFRAME>
      </td>
    </tr>
    <%If verMemb="2" Then%>
    <tr>
      <td align="center" valign="top" class="ULMenu2A" id="left_tab2A1" onClick="javascript:UsrMenu('left_tab2A1');"><div  style="width:15px; height:70px; background-image:url(../../inc/mem_img/member.gif);"></div></td>
    </tr>
    <tr>
      <td align="center" valign="top" class="ULMenu2B" id="left_tab2A2" onClick="javascript:UsrMenu('left_tab2A2');"><div  style="width:15px; height:70px; background-image:url(../../inc/mem_img/morder.gif);"></div></td>
    </tr>
    
    <tr>
      <th height="60px">&nbsp;</th>
    </tr>
    <%Else%>
    <tr>
      <td align="center" valign="top" class="ULMenuA" id="left_tab1A1" onClick="javascript:UsrMenu('left_tab1A1');">帐户资料</td>
    </tr>
    <tr>
      <td align="center" valign="top" class="ULMenuB" id="left_tab1A2" onClick="javascript:UsrMenu('left_tab1A2');">定单资料</td>
    </tr>
    
    <tr>
      <td align="center" valign="top" class="ULMenuB" id="left_tab1A6" onClick="javascript:UsrMenu('left_tab1A6');">留言笔记</td>
    </tr>
    <tr>
      <td align="center" valign="top" class="ULMenuB" id="left_tab1B1" onClick="javascript:UsrMenu('left_tab1B1');">论坛ＰＫ</td>
    </tr>
    <tr>
      <td align="center" valign="top" class="ULMenuB" id="left_tab1C1" onClick="javascript:UsrMenu('left_tab1C1');">供求商务</td>
    </tr>
    
    <tr>
      <th height="5px">&nbsp;</th>
    </tr>
    <%End If%>
  </table>
</div>
<script type="text/javascript">

function UsrMenu(xTabID){
  var oItem = document.getElementById('TabPage2').getElementsByTagName('td'); 
  for(var i=2; i<oItem.length; i++){
    var x = oItem[i];    
    x.className = "ULMenuB";
  }

  document.getElementById(xTabID).className = "ULMenuA";
  var dvs=document.getElementById("left_menu_cnt").getElementsByTagName("ul");
  for (var i=0;i<dvs.length;i++){
    if (dvs[i].id==('d'+xTabID)) { dvs[i].style.display='block'; }
    else                         { dvs[i].style.display='none'; }
  }
}
<%If verMemb="2" Then%>
UsrMenu('left_tab2A1');
<%Else%>
UsrMenu('left_tab1A1');
<%End If%>
</script>
