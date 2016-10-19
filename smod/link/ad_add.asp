<!--#include file="../../sadm/func1/func1.asp"-->
<!--#include file="../../sadm/func2/Func2.asp"-->
<!--#include file="../../sadm/func1/func_opt.asp"-->
<%
send = Request("send")
If send="send" Then
  InfType = RequestS("InfType","C",24)&"|"&RequestS("InfTyp2","C",24)&"|"&RequestS("InfTyp3","C",24)
  InfName = RequestS("InfName","C",24)&"|"&RequestS("InfNam2","C",24)
  InfPath = RequestS("InfPath1","C",120)&"|"&RequestS("InfPath2","C",120)
  InfCont = RequestS("InfCont1","C",480)&"|"&RequestS("InfCont2","C",480)
  InfPara = RequestS("InfPara1","N",80)&"|"&RequestS("InfPara2","N",60)&"|"&RequestS("InfPara3","N",40)&"|"&RequestS("InfPara4","N",30)&"|"&RequestS("InfPara5","C",12)
sql = " INSERT INTO [WebAdvert] (" 
sql = sql& "  KeyID,KeyMod" 
sql = sql& ", InfType, InfName, InfCont" 
sql = sql& ", InfPath, InfPara" 
sql = sql& ", LogAddIP,LogAUser,LogATime" 
sql = sql& ")VALUES(" 
sql = sql& "  '" & Get_AutoID(24) &"','Advert'" 
sql = sql& ", '" & InfType &"'" 
sql = sql& ", '" & InfName &"'" 
sql = sql& ", '" & InfCont &"'"  
sql = sql& ", '" & InfPath &"'" 
sql = sql& ", '" & InfPara &"'" 
sql = sql& ", '" & Get_CIP() &"'" 
sql = sql& ", '" & Session("UsrID") &"'" 
sql = sql& ", '" & Now() &"'" 
sql = sql& ")" ':Response.Write sql
  Call rs_DoSql(conn,sql) 
  Response.Write js_Alert("增加成功！","Redir","ad_list.asp") 

End If

InfPara1=200 : InfPara2=120 : InfPara3=480 : InfPara4=360
InfTyp2="AdvLRXX" : InfNam2=Date() : InfPara5="Left"

%>
<html>
<head>
<meta http-equiv="Pragma" content="no-cache">
<meta http-equiv="Expires" content="0">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title><%=Config_Name%>后台管理中心</title>
<link href="../../inc/adm_inc/adm_style.css" rel="stylesheet" type="text/css">
</head>
<body>
<!--#include file="config.asp"-->
<script src="../../sadm/func1/Func_JS.js" type="text/javascript"></script>
<form action="?" method="post" name="fm01" id="fm01" style="margin:0px;">
  <br>
  <table width="640" border="1" align="center" cellpadding="0" cellspacing="0" bordercolorlight="#CCCCCC" bordercolordark="#FFFFFF" style="margin-top:-2px;">
   
    <tr>
      <td width="12%" height="30" align="center" bgcolor="#EFEFEF">标题</td>
      <td align="left" bgcolor="#EFEFEF"><input name="InfName" id="InfName" value="<%=sSubj%>" size="45" maxlength="12" />
        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;落款
      <input name="InfNam2" id="InfNam2" value="<%=InfNam2%>" size="12" maxlength="12" />
      (部分模版)</td>
    </tr>
    <tr>
      <td height="30" align="center" bgcolor="#EFEFEF">组别</td>
      <td align="left" bgcolor="#EFEFEF"><select name="InfType" style="width:120px; " id="InfType">
        <%=Get_rsOpt(conn,"SELECT TypID,TypName FROM WebTyps WHERE TypMod='HomAdvert'",InfType)%>
      </select>        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
        类别
        <select name="InfTyp2" id="InfTyp2" style="width:120px; ">
          <option value="AdvPair" <%If InfTyp2="AdvPair" Then Response.Write("selected")%>>对联广告</option>
          <option value="AdvJJCC" <%If InfTyp2="AdvJJCC" Then Response.Write("selected")%>>警警察察</option>
          <option value="Float01" <%If InfTyp2="Float01" Then Response.Write("selected")%>>漂浮一</option>
          <option value="Float02" <%If InfTyp2="Float02" Then Response.Write("selected")%>>漂浮二</option>
          <option value="AdvLRXX" <%If InfTyp2="AdvLRXX" Then Response.Write("selected")%>>左右浮动</option>
          <option value="AdvLRQQ" <%If InfTyp2="AdvLRQQ" Then Response.Write("selected")%>>QQ浮动</option>
        </select>
        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;模版
        <select name="InfTyp3" id="InfTyp3" style="width:120px; ">
          <option value="">[无模版]</option>
          <option value="Rnd_nid" <%If InfTyp3="Rnd_nid" Then Response.Write("selected")%>>蓝色细框</option>
          <option value="Rnd_n01" <%If InfTyp3="Rnd_n01" Then Response.Write("selected")%>>灰色圆角</option>
          <option value="Rnd_n02" <%If InfTyp3="Rnd_n02" Then Response.Write("selected")%>>粉红圆角</option>
          <option value="Rnd_n03" <%If InfTyp3="Rnd_n03" Then Response.Write("selected")%>>灰色圆角</option>
          <option value="Rnd_n04" <%If InfTyp3="Rnd_n04" Then Response.Write("selected")%>>蓝色圆角</option>
          <option value="Rnd_n05" <%If InfTyp3="Rnd_n05" Then Response.Write("selected")%>>绿色QQ</option>
          <option value="Rnd_n06" <%If InfTyp3="Rnd_n06" Then Response.Write("selected")%>>灰色QQ</option>
        </select>
        <span class="FontRed">
        <input name="Submit" type="button" class="ModShow" value="查看" onClick="javascript:window.open('ad_view.asp?Act=VMod&ModID='+document.fm01.InfTyp3.value)"/>
        </span> </td>
    </tr>
    <tr>
      <td height="30" align="center" bgcolor="#EFEFEF">宽高</td>
      <td align="left" bgcolor="#EFEFEF"><input name="InfPara1" id="InfPara1" value="<%=InfPara1%>" size="2" maxlength="3" />
        x
        <input name="InfPara2" id="InfPara2" value="<%=InfPara2%>" size="2" maxlength="3" />
        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;坐标
        <input name="InfPara3" id="InfPara3" value="<%=InfPara3%>" size="2" maxlength="3" />
        x
        <input name="InfPara4" id="sWidth5" value="<%=InfPara4%>" size="2" maxlength="3" />
        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;左右
        <select name="InfPara5" id="InfPara5" style="width:120px; ">
          <option value="">[无区分]</option>
          <option value="Right" <%If InfPara5="Right" Then Response.Write("selected")%>>右</option>
          <option value="Left"  <%If InfPara5="Left" Then Response.Write("selected")%>>左</option>
        </select></td>
    </tr>
    <tr>
      <td height="30" align="center" bgcolor="#FFFFFF">地址</td>
      <td align="left" bgcolor="#FFFFFF"><input name="InfPath1" id="sUrl4" value="<%=InfPath1%>" size="60" maxlength="60" />
        (对Flash无效) <br />
        <input name="InfPath2" id="InfPath2" value="<%=InfPath2%>" size="60" maxlength="60" /></td>
    </tr>
    <tr>
      <td height="30" align="center">内容<br>
        (240字内)</td>
      <td align="left"><textarea name="InfCont1" cols="60" rows="5" id="InfCont1"><%=sCont%></textarea>
        <br>
        <input name=Button type=button id="Button2" value="选择图片1" onClick="getRetObject('InfCont1');window.open('../file/file_view.asp?yPath=myfile/link/')">
        <input name=view2 type=button id="Button1" value="选择图片2" onClick="getRetObject('InfCont2');window.open('../file/file_view.asp?yPath=myfile/link/')">
        <br>
      <textarea name="InfCont2" cols="60" rows="5" id="InfCont2"><%=sCont%></textarea></td>
    </tr>
    <tr>
      <td height="30" align="center" bgcolor="#EFEFEF"><input name="send" type="hidden" id="send" value="send"></td>
      <td align="left" bgcolor="#EFEFEF"><span class="FontRed">
        <input name="Submit2" type="submit" class="ButtonShort" value="确 认" />
        &nbsp;&nbsp;<a href="ad_list.asp">返回 </a></span></td>
    </tr>
    <tr>
      <td height="30" colspan="2" align="left" bgcolor="#EFEFEF"><span class="col00F">说明</span>： 请仔细阅读本说明，<span class="colF0F">如果觉得太麻烦，那请不要用这个功能！</span><br>
        <span class="col00F">标题，落款</span>：仅有些模版显示，即使如果模版不显示此项目，强烈建议填写一个有意义的标题；<br>
        <span class="col00F">组别</span>：可以动态添加，但要在前台作相应的调用，一般用默认的即可；建议演示区的不要删除，把不要的放入（修改）垃圾，谨慎设置首页广告，设置好后并刷新；<br>
        <span class="col00F">类别与模版</span>：
        对联广告，
        <OPTION value="AdvJJCC">警警察察</OPTION>
        ：每页（组别）只能放一对；
        <OPTION value="Float01">漂浮一</OPTION>
        ，
        <OPTION value="AdvJJCC"></OPTION>
        <OPTION value="Float02">漂浮二：每页只能放一个；</OPTION>
        <OPTION value="AdvLRXX" selected>左右浮动</OPTION>
        <OPTION value="AdvLRQQ">QQ浮动</OPTION>
        ：每页只能放多个；<br>
        <span class="col00F">宽高,坐标与左右</span>：宽高制广告区大小，高一般用自动即设置的高无效；坐标指广告的位置，如果是对联广告，
        <OPTION value="AdvJJCC">警警察察</OPTION>
        为上边距和左右边距；如果是
        <OPTION value="Float01">漂浮一</OPTION>
，
<OPTION value="AdvJJCC"></OPTION>
<OPTION value="Float02">漂浮二</OPTION>
指他们的初始位置；
<OPTION value="AdvLRXX" selected>左右浮动</OPTION>
<OPTION value="AdvLRQQ">QQ浮动</OPTION>
指他们的上边距和左右边距；左右设置仅对
<OPTION value="AdvLRXX" selected>左右浮动</OPTION>
<OPTION value="AdvLRQQ">QQ浮动有效。</OPTION>
<br>
<span class="col00F">地址与内容</span>：仅对联广告和警警察察填两项，其他只填一项即可；地址指广告的连接地址，如果内容为Flash，则地址无效；内容可以直接填写<span class="colF0F">文本内容</span>，可填写<span class="colF0F">图片或Flash地址</span>如：http://www.xxx.com/path/file.swf|.jpg|.gif，可填写<span class="colF0F">图片或Flash的代码</span>如：&lt;embed src='/images/flash/file.swf' quality='high' type='application/x-shockwave-flash' width=240 height=180&gt;&lt;/embed&gt; 或 &lt;a href='/news.asp' target='_blank'&gt;&lt;img src='/images/logo.jpg' width=120 height=60 border='0'/&gt;&lt;/a&gt; 等</td>
    </tr>
  </table>
</form>
<script type="text/javascript">
var yFile;
function getRetObject(id)
{
	yFile = eval("document.fm01."+id);
}
</script>
</body>
</html>
