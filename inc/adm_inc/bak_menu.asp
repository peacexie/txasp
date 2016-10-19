<%
If Session("UsrID")&"" = "" Then 
  Response.End()
End If
%>
<style type="text/css">
#showpage {
	CLEAR: both;
	text-align: center;
	width: 100%;
}
</style>


  
  <!-- // Menu Start //////////// -->
  




<div>
  <ul class='left_top' onClick="ShowMenu('ModMember')" >
    <li class='left_top_left left'>会员与查询</li>
    <li class='left_top_right right'></li>
  </ul>
  <ul class='left_conten col999' id='sSubModMember'>
<li>
<A target='mainFrame' href='../../member/admin/member.asp'>会员管理</A>
<A target='mainFrame' href='../../member/admin/system.asp'>参数</A>
<A target='_blank' href='../../member/mu_app_123456.asp'>注册</A>
</li>
<li>
<A target='mainFrame' href='../../smod/gbook/info_list.asp?ModID=MemB124'>会员反馈</A>
<A target='mainFrame' href='../../smod/gbook/info_list.asp?ModID=MemB224'>通知</A>
<A target='mainFrame' href='../../smod/gbook/info_list.asp?ModID=MemB324'>笔记</A>
</li>
<li>
<A target="mainFrame" href="../../member/ecard/member.asp">档案查询</A>
<A target="mainFrame" href="../../smod/type/type_list.asp?ModID=MemC124">分类</A>
<A target="mainFrame" href="../../member/ecard/admimport.asp">导入</A>
</li>
  </ul>
  <div class='left_end'></div>
</div>

<div>
  <ul class='left_top' onClick="ShowMenu('ModBBS')" >
    <li class='left_top_left left'>论坛与投票</li>
    <li class='left_top_right right'></li>
  </ul>
  <ul class='left_conten col999' id='sSubModBBS'>
<li id='admliBBSP124'>
<A target='mainFrame' href='../../bbs/badm/bbs_list.asp?ModID=BBSP124'>公开论坛</A>
 - 
<A target='mainFrame' href='../../smod/type/type_list.asp?ModID=BBSP124'>分类</A>
</li><li id='admliBBSPK24'>
<A target='mainFrame' href='../../bbs/badm/info_list.asp?ModID=BBSPK24'>观点辩论</A>
<A target='mainFrame' href='../../bbs/badm/view_list.asp?ModID=BBSPK24'>观点</A>
<A target='mainFrame' href='../../smod/type/type_list.asp?ModID=BBSPK24'>类别</A>
</li><li id='admliBBSI124'>
<A target='mainFrame' href='../../bbs/badm/bbs_list.asp?ModID=BBSI124'>内部论坛</A>
 - 
<A target='mainFrame' href='../../smod/type/type_list.asp?ModID=BBSI124'>分类</A>
</li><li id='admliBBSVA24'>
<A target='mainFrame' href='../../smod/vote/info_list.asp?ModID=BBSVA24'>投票项目</A>
<A target='mainFrame' href='../../smod/type/type_list.asp?ModID=BBSVA24'>类别</A>
<A target='mainFrame' href='../../smod/vote/vote_logs.asp'>记录</A>
</li><li id='admliBBSVB24'>
<A target='mainFrame' href='../../smod/vote/ind_list.asp?ModID=BBSVB24'>问卷调查</A>
<A target='mainFrame' href='../../smod/type/type_list.asp?ModID=BBSVB24'>类别</A>
<A target='mainFrame' href='../../smod/vote/vote_logs.asp'>记录</A>
</li><li id='admliBBSVC24'>
<A target='mainFrame' href='../../smod/vote/inf2_list.asp?ModID=BBSVC24'>简易调查</A>
<A target='mainFrame' href='../../smod/type/type_list.asp?ModID=BBSVC24'>类别</A>
<A target='mainFrame' href='../../smod/vote/vote_logs.asp'>记录</A>
</li>
  </ul>
  <div class='left_end'></div>
</div>

<div>
  <ul class='left_top' onClick="ShowMenu('ModTrade')" >
    <li class='left_top_left left'>商务与供求</li>
    <li class='left_top_right right'></li>
  </ul>
  <ul class='left_conten col999' id='sSubModTrade'>
<li>
<A target='mainFrame' href='../../trade/madm/info_corp.asp'>企业资料</A>
<A target='mainFrame' href='../../trade/madm/set_temp.asp'>参数</A>
<A target='mainFrame' href='../../trade/madm/set_type.asp?ModID=TraA124'>类别</A>
</li>
<li>
<A target='mainFrame' href='../../smod/info/info_list.asp?ModID=TraA124'>企业介绍</A>
<A target='mainFrame' href='../../trade/madm/set_type.asp?ModID=TraA124'>类别</A>
<A target='mainFrame' href='../../trade/madm/info_gbook.asp?ModID=TraG124'>留言</A>
</li>
<li>
<A target='mainFrame' href='../../smod/info/info_list.asp?ModID=TraN124'>新闻中心</A> 
<A target='mainFrame' href='../../trade/madm/set_type.asp?ModID=TraN124'>类别</A>
<A target='mainFrame' href='../../trade/madm/info_gbook.asp?ModID=TraR124'>评论</A>
</li>
<li>
<A target='mainFrame' href='../../smod/info/info_list.asp?ModID=TraT124'>供求交易</A>
<A target='mainFrame' href='../../trade/madm/set_type.asp?ModID=TraT124'>类别</A>
<A target='mainFrame' href='../../trade/madm/info_gbook.asp?ModID=TraO124'>订购</A>
</li>
<li>
<A target='mainFrame' href='../../smod/info/info_list.asp?ModID=TraJ124'>招聘信息</A>
<A target='mainFrame' href='../../trade/madm/set_type.asp?ModID=TraA124'>类别</A>
<A target='mainFrame' href='../../trade/madm/info_gbook.asp?ModID=TraO124'>应聘</A>
</li>
  </ul>
  <div class='left_end'></div>
</div>

<div>
  <ul class='left_top' onClick="ShowMenu('ModDocs')" >
    <li class='left_top_left left'>内部公文</li>
    <li class='left_top_right right'></li>
  </ul>
  <ul class='left_conten col999' id='sSubModDocs'>
<li>
<A target='mainFrame' href='../../doc/dadm/adm_list.asp'>公文管理</A>
-
<A target='mainFrame' href='../../doc/dadm/adm_logs.asp?ModID=DocD124'>查阅记录</A>
</li>
<li>
<A target='mainFrame' href='../../doc/dadm/adm_rem.asp'>公文评论</A>
-
<A target='mainFrame' href='../../smod/type/type_list.asp?ModID=DocD124'>公文类别</A>
</li>
<li>
<A target='mainFrame' href='../../sadm/user/user.asp?ModID=Inner'>公文会员</A>
-
<A target='mainFrame' href='../../sadm/system/system.asp?ModID=Inner'>公文分组</A>
</li>
<li>
<A target='mainFrame' href='../../smod/gbook/info_list.asp?ModID=MemB424'>会员反馈</A>
<A target='mainFrame' href='../../smod/gbook/info_list.asp?ModID=MemB524'>通知</A>
<A target='mainFrame' href='../../smod/gbook/info_list.asp?ModID=MemB624'>笔记</A>
</li>

  </ul>
  <div class='left_end'></div>
</div>

<div>
  <ul class='left_top' onClick="ShowMenu('ModSms')" >
    <li class='left_top_left left'>短信系统</li>
    <li class='left_top_right right'></li>
  </ul>
  <ul class='left_conten col999' id='sSubModSms'>
<li>
<A target='mainFrame' href='../../msg/admin/info.asp?act=Balance'>接口余额</A>
---
<A target='mainFrame' href='../../msg/admin/info.asp?act=Charge'>接口充值</A>
</li>
<li>
<A target='mainFrame' href='../../msg/admin/send.asp?act='>发送短信</A>
---
<A target='mainFrame' href='../../msg/admin/send.asp?act=Test'>发送测试</A>
</li>
<li>
<A target='mainFrame' href='../../msg/admin/logr.asp'>充值记录</A>
---
<A target='mainFrame' href='../../msg/admin/logs.asp'>发送记录</A>
</li>

<li>
<A target='mainFrame' href='../../msg/user/member.asp'>会员</A>
<A target='mainFrame' href='../../msg/user/tels.asp'>电话薄</A>
<A target='mainFrame' href='../../msg/user/temp.asp'>范本</A>
</li>
  </ul>
  <div class='left_end'></div>
</div>

<div>
  <ul class='left_top' onClick="ShowMenu('ModInfo')" >
    <li class='left_top_left left'>新闻与介绍</li>
    <li class='left_top_right right'></li>
  </ul>
  <ul class='left_conten col999' id='sSubModInfo'>
<li id='admliInfA124'>
<A target='mainFrame' href='../../smod/info/info_list.asp?ModID=InfA124'>企业介绍</A>
-
<A target='mainFrame' href='../../smod/info/info_add.asp?ModID=InfA124'>增加</A>
-
<A target='mainFrame' href='../../smod/gbook/out_list.asp?ModID=InfA124'>评论</A>
</li><li id='admliInfD124'>
<A target='mainFrame' href='../../smod/info/info_list.asp?ModID=InfD124'>部门介绍</A>
-
<A target='mainFrame' href='../../sadm/system/system.asp?ModID=Depart'>设置</A>
-
<A target='mainFrame' href='../../smod/gbook/out_list.asp?ModID=InfD124'>评论</A>
</li><li id='admliInfC124'>
<A target='mainFrame' href='../../smod/info/info_list.asp?ModID=InfC124'>课程介绍</A>
-
<A target='mainFrame' href='../../smod/info/info_add.asp?ModID=InfC124'>增加</A>
-
<A target='mainFrame' href='../../smod/gbook/out_list.asp?ModID=InfC124'>评论</A>
</li><li id='admliInfN124'>
<A target='mainFrame' href='../../smod/info/info_list.asp?ModID=InfN124'>新闻中心</A>
-
<A target='mainFrame' href='../../smod/info/info_add.asp?ModID=InfN124'>增加</A>
-
<A target='mainFrame' href='../../smod/gbook/out_list.asp?ModID=InfN124'>评论</A>
</li><li id='admliInfN224'>
<A target='mainFrame' href='../../smod/info/info_list.asp?ModID=InfN224'>新闻(英)</A>
-
<A target='mainFrame' href='../../smod/info/info_add.asp?ModID=InfN224'>增加</A>
-
<A target='mainFrame' href='../../smod/gbook/out_list.asp?ModID=InfN224'>评论</A>
</li><li hdFlag>
<A target=mainFrame href=../../smod/info/head_list.asp>头条设置</A>
<A target=mainFrame href=../../smod/info/set_stat.asp>信息发布统计</A>
</li>

<li>
<A target='mainFrame' href='../../smod/type/type_center.asp?MD=Inf'> ** 信息模块类别管理 ** </A>
</li>
  </ul>
  <div class='left_end'></div>
</div>

<div>
  <ul class='left_top' onClick="ShowMenu('ModPics')" >
    <li class='left_top_left left'>图片与视频</li>
    <li class='left_top_right right'></li>
  </ul>
  <ul class='left_conten col999' id='sSubModPics'>
<li id='admliPicS124'>
<A target='mainFrame' href='../../smod/info/info_list.asp?ModID=PicS124'>产品图片</A>
-
<A target='mainFrame' href='../../smod/info/info_add.asp?ModID=PicS124'>增加</A>
-
<A target='mainFrame' href='../../smod/info/ord_list.asp?ModID=PicS124'>定单</A>
</li><li id='admliPicS224'>
<A target='mainFrame' href='../../smod/info/info_list.asp?ModID=PicS224'>产品(英)</A>
-
<A target='mainFrame' href='../../smod/info/info_add.asp?ModID=PicS224'>增加</A>
-
<A target='mainFrame' href='../../smod/info/ord_list.asp?ModID=PicS224'>定单</A>
</li><li id='admliPicR124'>
<A target='mainFrame' href='../../smod/info/info_list.asp?ModID=PicR124'>人物图片</A>
-
<A target='mainFrame' href='../../smod/info/info_add.asp?ModID=PicR124'>增加</A>
-
<A target='mainFrame' href='../../smod/gbook/out_list.asp?ModID=PicR124'>评论</A>
</li><li id='admliPicT124'>
<A target='mainFrame' href='../../smod/info/info_list.asp?ModID=PicT124'>其它图片</A>
-
<A target='mainFrame' href='../../smod/info/info_add.asp?ModID=PicT124'>增加</A>
-
<A target='mainFrame' href='../../smod/gbook/out_list.asp?ModID=PicT124'>评论</A>
</li><li id='admliPicV124'>
<A target='mainFrame' href='../../smod/info/info_list.asp?ModID=PicV124'>视频播放</A>
-
<A target='mainFrame' href='../../smod/info/info_add.asp?ModID=PicV124'>增加</A>
-
<A target='mainFrame' href='../../smod/gbook/out_list.asp?ModID=PicV124'>评论</A>
</li><li id='admliPicV125'>
<A target='mainFrame' href='../../smod/info/info_list.asp?ModID=PicV125'>文件下载</A>
-
<A target='mainFrame' href='../../smod/info/info_add.asp?ModID=PicV125'>增加</A>
-
<A target='mainFrame' href='../../smod/gbook/out_list.asp?ModID=PicV125'>评论</A>
</li><li id='admliPicTBak'>
<A target='mainFrame' href='../../smod/info/info_list.asp?ModID=PicTBak'>备用图片</A>
-
<A target='mainFrame' href='../../smod/info/info_add.asp?ModID=PicTBak'>增加</A>
-
<A target='mainFrame' href='../../smod/gbook/out_list.asp?ModID=PicTBak'>评论</A>
</li>
<li>
<A target='mainFrame' href='../../smod/info/order_set.asp' hdFlag>定单参数</A>
---
<A target='mainFrame' href='../../smod/type/type_center.asp?MD=Pic'>信息类别</A>
</li>
  </ul>
  <div class='left_end'></div>
</div>

<div>
  <ul class='left_top' onClick="ShowMenu('ModGbook')" >
    <li class='left_top_left left'>留言与笔记</li>
    <li class='left_top_right right'></li>
  </ul>
  <ul class='left_conten col999' id='sSubModGbook'>
<li id='admliGboK124'>
<A target='mainFrame' href='../../smod/gbook/info_list.asp?ModID=GboK124'>(中文)访客留言</A>
<A target='mainFrame' href='../../smod/type/type_list.asp?ModID=GboK124'>分类</A>
</li><li id='admliGboK224'>
<A target='mainFrame' href='../../smod/gbook/info_list.asp?ModID=GboK224'>(英文)访客留言</A>
<A target='mainFrame' href='../../smod/type/type_list.asp?ModID=GboK224'>分类</A>
</li><li id='admliGboU124'>
<A target='mainFrame' href='../../smod/gbook/info_list.asp?ModID=GboU124'>站务笔记</A>
<A target='mainFrame' href='../../smod/gbook/info_add.asp?ModID=GboU124'>增加</A>
<A target='mainFrame' href='../../smod/type/type_list.asp?ModID=GboU124'>分类</A>
</li><li id='admliGboU224'>
<A target='mainFrame' href='../../smod/gbook/info_list.asp?ModID=GboU224'>私人秘密</A>
<A target='mainFrame' href='../../smod/gbook/info_add.asp?ModID=GboU224'>增加</A>
<A target='mainFrame' href='../../smod/type/type_list.asp?ModID=GboU224'>分类</A>
</li>
  </ul>
  <div class='left_end'></div>
</div>

<div>
  <ul class='left_top' onClick="ShowMenu('ModHome')" >
    <li class='left_top_left left'>刷新与杂项</li>
    <li class='left_top_right right'></li>
  </ul>
  <ul class='left_conten col999' id='sSubModHome'>
<li id='HomUpdate'>
<A target='mainFrame' href='../../smod/adupd/upd.asp'>信息刷新</A>
 -- 
<A target='mainFrame' href='../../tools/help/xhelp.asp'>管理帮助</A>
</li><li id='admliHomSets'>
<A target='mainFrame' href='../../smod/adupd/para_list.asp'>
常用杂项 --- 快捷设置</A>
</li>
<li>
<A id='admaHomAdvert' target='mainFrame' href='../../smod/link/ad_pic.asp'>广告设置</A>
 -- 
<A id='admaHomCount' target='mainFrame' href='../../tools/base/omaster.asp'>统计计数</A>
</li><li id='admliHomLnk1'>
<A target='mainFrame' href='../../smod/link/info_list.asp'>友情连接</A>
<A target='mainFrame' href='../../smod/link/info_add.asp'>增加</A>
<A target='mainFrame' href='../../smod/type/type_list.asp?ModID=HomLnk1'>分类</A>
</li><li id='admliBBSVD24'>
<A target='mainFrame' href='../../smod/vote/inf3_list.asp?ModID=BBSVD24'>网上调查</A>
<A target='mainFrame' href='../../smod/vote/inf3_add.asp?ModID=BBSVD24'>增加</A>
<A target='mainFrame' href='../../smod/type/type_list.asp?ModID=BBSVD24'>分类</A>
</li>
  </ul>
  <div class='left_end'></div>
</div>

<div>
  <ul class='left_top' onClick="ShowMenu('ModSystem')" >
    <li class='left_top_left left'>系统与设置</li>
    <li class='left_top_right right'></li>
  </ul>
  <ul class='left_conten col999' id='sSubModSystem'>
<li id='admliSysMods'>
<A target='mainFrame' href='../../sadm/system/logs_list.asp'>系统记录</A>
<A target='mainFrame' href='../../sadm/system/system.asp'>模块</A>
<A target='mainFrame' href='../../sadm/system/para_yno.asp'>参数</A>
</li>
<li id='admliSysUList'>
<A target='mainFrame' href='../../sadm/user/user_editpw.asp'>个人密码</A>
<A target='mainFrame' href='../../sadm/user/user.asp'>管理员列表</A>
</li><li id='admliSysDB'>
<A target='mainFrame' href='../../sadm/admin/typs_list.asp'>信息类别</A> 
<A target='mainFrame' href='../../sadm/admin/sys_dbacc.asp'>数据库管理</A>
</li><li id='admliSysMenu'>
<A target='mainFrame' href='../../sadm/admin/sys_menu.asp'>系统菜单</A>
--
<A target='mainFrame' href='../../sadm/admin/sys_mpara.asp'>菜单参数</A>
</li><li id='admliSysConfig'>
<A target='mainFrame' href='../../sadm/admin/type_list.asp'>系统类别</A>
<A target='mainFrame' href='../../sadm/admin/sys_config.asp'>配置</A>
<A target='mainFrame' href='../../tools/tools.asp?Act=Admin'>工具</A>
</li>
  </ul>
  <div class='left_end'></div>
</div>
<script type='text/javascript'>var mStr=';ModMember;ModBBS;ModTrade;ModDocs;ModSms;ModInfo;ModPics;ModGbook;ModHome;ModSystem';</script>

  <!-- // Menu End /////////// -->
  
  
