<!--#include file="config.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Pragma" content="no-cache">
<meta http-equiv="Expires" content="0">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title><%=Config_Name%>后台管理中心</title>
<link href="../../inc/adm_inc/adm_style.css" rel="stylesheet" type="text/css">
<script type="text/javascript" charset="utf-8" src="../../inc/home/jquery.js"></script>
<script type="text/javascript" charset="utf-8" src="../../inc/home/jsInfo.js"></script>
<script type="text/javascript" charset="utf-8" src="../../smod/file/edt_api.asp?edtAct=mainLoad"></script>
<script src="../../inc/home/jskeys.js" type="text/javascript"></script>
<style type="text/css">
tr, td {
	background-color:#FFF;
}
</style>
</head>
<body>
<%

If get_ModCopy(ModID) Then
  Response.Write js_Alert("注意：此类别(模块)信息，不需要再添加；\n只需要在列表页执行同步,并作适当编辑即可","Redir","info_list.asp?ModID="&ModID) 
  Response.End()
End If

send = Request("send")
ReEnd = Request("ReEnd")
InfType = RequestS("InfType","C",255)
InfTyp2 = RequestS("InfTyp2","C",24)
Session("KeyID") = rs_AutID(conn,ModTab,"KeyID",upPart,"1","")

tmrPara = get_TmrPara(true)

If PrmFlag="(Mem)" Then
  tmpCont = rs_Val("","SELECT ParRem FROM TradePara WHERE ParFlag='"&ModID&"' AND LogAUser='"&Session("MemID")&"' ")
Else
 'If Left(ModID,3)="Tra" Then
  tmpCont = rs_Val("","SELECT ParRem FROM AdmPara WHERE ParCode='tmp"&ModID&"'") 
 'End If
End If


ModImgCount = Eval(ModID&"ImgCount")
ModImgCount = RequestSafe(ModImgCount,"N",1)

'Response.Write Session("KeyID")

  Dim sys27_Rnd(10)
  sys27_RVal = Rnd_Base("5678",9)&Rnd_Base("",64)
  For i = 1 To 9
    sys27_Rnd(i)=Mid(sys27_RVal,i*6,Mid(sys27_RVal,i,1))
  Next

%>
<%If Session("MemID")="guest" Then%>
<!--#include file="guest_add.asp"-->
<%Else%>
<table width="100%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="F0F0F0">
  <form name="fm01<%=sys27_RVal%>" id="fm01<%=sys27_RVal%>" action="info_add2.asp?FrmFlag=<%=FrmFlag%>" enctype="multipart/form-data" method="post">
    <input name="<%=App27Random%>" type="hidden" id="<%=App27Random%>" value="<%=sys27_RVal%>">
    <tr>
      <td align="center"><p>&nbsp;</p></td>
      <td align="center"><strong>[<%=rs_Val("","SELECT SysName FROM AdmSyst WHERE SysID='"&ModID&"'")%>]增加</strong></td>
    </tr>
    <tr>
      <td align="center" nowrap>主题</td>
      <td><input name="InfSubj<%=sys27_Rnd(1)%>" type="text" id="InfSubj<%=sys27_Rnd(1)%>" size="60" maxlength="120">
        <select name='SetSubj' size=1 class="Input_Text" id="SetSubj" style="width:90 ">
          <OPTION style='COLOR:#000000; BACKGROUND-COLOR:#000000' value='000000'>-[默认]-</OPTION>
          <OPTION style='COLOR:#FF0000; BACKGROUND-COLOR:#FF0000' value='FF0000'>#FF0000</OPTION>
          <OPTION style='COLOR:#00FF00; BACKGROUND-COLOR:#00FF00' value='00FF00'>#00FF00</OPTION>
          <OPTION style='COLOR:#0000FF; BACKGROUND-COLOR:#0000FF' value='0000FF'>#0000FF</OPTION>
          <OPTION style='COLOR:#00FFFF; BACKGROUND-COLOR:#00FFFF' value='00FFFF'>#00FFFF</OPTION>
          <OPTION style='COLOR:#FF00FF; BACKGROUND-COLOR:#FF00FF' value='FF00FF'>#FF00FF</OPTION>
          <OPTION style='COLOR:#FFFF00; BACKGROUND-COLOR:#FFFF00' value='FFFF00'>#FFFF00</OPTION>
        </select></td>
    </tr>
    <tr>
      <td align="center" nowrap><%=Get_TypeLable(ModID)%></td>
      <td><select name="InfType" id="InfType">
          <%=Get_TypeOpt(ModID,InfType)%>
        </select>
        &nbsp; &nbsp;
  		<%=Get_Typ2Opt(ModID,InfTyp2) %>
        &nbsp; &nbsp;
        <input name=Button type=button id="Button2" value="选择模板" onClick="owTemp()" />
        </td>
    </tr>
    <%=tmrPara(0)%>
    <tr>
      <td align="center">内容</td>
      <td>
<textarea id="InfCont<%=sys27_Rnd(4)%>" name="InfCont<%=sys27_Rnd(4)%>" style="width:600px;height:360px;visibility:hidden;display:none"><%=Show_Form(tmpCont)%></textarea>
<script type="text/javascript" charset="utf-8" src="../../smod/file/edt_api.asp?edtID=EditID01<%=sys27_Rnd(4)%>&edtCont=InfCont<%=sys27_Rnd(4)%>"></script>
        </td>
    </tr>
    
    <tr>
      <td align="center">属性</td>
      <td><select name="SetHot" id="SetHot">
          <option value="N" selected >一般</option>
          <option value="Y" >推荐</option>
        </select>
        <select name="SetShow" id="SetShow">
          <%If Flg_PCheck() Then%>
          <option value="Y" selected >审核</option>
          <%End If%>
          <option value="N" >未审</option>
        </select>
        <select name="SetTop" id="SetTop">
          <option value="888" selected >顺序</option>
          <%
	  For i = 0 to 9
	  sel = ""
	  %>
          <option value="<%=i%>"><%=i%></option>
          <%Next%>
        </select>
        &nbsp; &nbsp; 
        时间
        <input name="LogATime" type="text" id="LogATime" value="<%=Now()%>" size="20" maxlength="20">
        &nbsp; <a href="#" onClick="owFile('../file/img_list.asp?send=PicList&ID=<%=Session("KeyID")%>&EdtID=EditID01<%=sys27_Rnd(4)%>')">附件管理</a></td>
    </tr>
    <%=tmrPara(1)%>
    <%
	For i=1 To ModImgCount
	If ModImgCount=1 Then
	  snPic = "附图"
	Else
	  snPic = "附图"&i
	End If
	%>
    <tr>
      <td align="center" nowrap><%=snPic%></td>
      <td><input type='file' name='ImgName<%=i%>' id="ImgName<%=i%>" onChange="chkPVal(this)" style="width:360px; ">
        (建议50K内) </td>
    </tr>
    <%Next%>
    <%
	If i=2 And Eval(ModID&"ImgSCopy")="Y" Then
	  sa = Split(Eval(ModID&"ImgSAtuo")&"x","x")
	  s1 = RequestSafe(sa(0),"N",0)
	  s2 = RequestSafe(sa(1),"N",0)
	%>
    <tr title="附图为空时,此选项有效">
      <td align="center" nowrap>提取</td>
      <td><input name="ImgSCopy" type="checkbox" id="ImgSCopy" value="Y">
        自动取内容第一个图为附图 &nbsp;&nbsp;自动缩略
        <input name="ImgSmall1" type="text" id="ImgSmall1" value="<%=s1%>" size="3" maxlength="4" style="width:30px">
        x
        <input name="ImgSmall2" type="text" id="ImgSmall2" value="<%=s2%>" size="3" maxlength="4" style="width:30px">
        0为不缩略</td>
    </tr>
    <%End If%>
    <tr>
      <td align="center" nowrap>返回</td>
      <td><input name="ReEnd" type="radio" id="ReEnd1" value="N" <%If ReEnd="N" Then Response.Write("checked")%>>
        添加资料后返回列表
        &nbsp;&nbsp;&nbsp;
        <input type="radio" name="ReEnd" id="ReEnd2" value="Y" <%If ReEnd="Y" Then Response.Write("checked")%>>
        添加资料后继续</td>
    </tr>
    <%If PrmFlag="(Mem)" Then%>
    <tr>
      <td align="center" nowrap>认证码</td>
      <td><input name="ChkCode" type="text" id="ChkCode" size="6" maxlength="12" xonfocus="javascript:PicReLoad('../../');"/>
        <img src="../../sadm/pcode/img_frnd.asp" alt="<%=vPMsg_RelCode%>" name="ChkCImg" align="absmiddle" id="ChkCImg" style="cursor:pointer;" onClick="PicReLoad('../../')"/></td>
    </tr>
    <%End If%>
    <tr>
      <td align="center" nowrap>&nbsp;</td>
      <td><input name=view type=button id="Button1" value="确    定" onClick="chkData()">
        <input name=view type=reset id="Button1" value="重    写">
        <input name="ModImgCount" type="hidden" id="ModImgCount" value="<%=ModImgCount%>">
        <input name="PrmFlag" type="hidden" id="PrmFlag" value="<%=PrmFlag%>">
      </td>
    </tr>
  </form>
</table>
<%End If%>
<script type="text/javascript">

var tmpCont = "";
var fm = document.fm01<%=sys27_RVal%>;

function setKeys(){
  try{
	str = apiGetValEditID01<%=sys27_Rnd(4)%>();
	if(str.length==0) { alert("请先填写内容!"); }
	fm.InfPara1.value = getKeyN(str,3).replace(" ... ","");
  }catch(objError){ }
}

function chkPVal(e)
{
  var v = e.value; // F:\temp\t_view\gua_08.jpg
  if(v.length>0){
	if(isIE){
	  var f1 = v.substring(1,3); 
	  if(f1!=":\\") { 
	    alert("附图地址错误<1>！"+e.value); 
	    e.outerHTML = e.outerHTML; 
	    return;
	  }
	  va = v.split("\\");
	  v = va[va.length-1];
	}
  }
  var ed=v.indexOf(".");
  var e4=v.substring(v.length-4,v.length);
  var e5=v.substring(v.length-5,v.length);
  //(e4!=".")&&(e5!=".")
  if((ed<0)||((e4+e5).indexOf(".")<0)){
    alert("附图地址错误<2>！"); 
	e.outerHTML = e.outerHTML; 
	return;
  }
}

function owFile(url){ 
  window.open(url,"winFiles",",width=720,height=560,scrollbars=yes");
}
function owTemp(){
  var opt = fm.InfType;
  var typ = fm.InfType.value;
  window.open("set_temp.asp?tMod=<%=ModID%>&tTyp="+typ+"");
}
function getTemp() 
{
	apiSetValEditID01<%=sys27_Rnd(4)%>(tmpCont);
}

var xFile,xSize;
function owFileGet()
{
	  var tFile = xFile; 
	  var nPos = tFile.indexOf('/upfile/');
	  var sSub = tFile.substring(0,nPos);
	  tFile = tFile.replace(sSub,''); 
	  fm.InfPara7.value = tFile;
	  nPos = tFile.indexOf('.');
	  sSub = tFile.substring(0,nPos);
	  tFile = tFile.replace(sSub,''); 
	  fm.InfPara3.value = tFile;
	  fm.InfPara4.value = xSize;
}

function chkData()
{
       
	   var eflag = 0;
       for(i=0;i<1;i++)
         {  ////////// //////////////// Start For ////////////////
 if (fm.InfSubj<%=sys27_Rnd(1)%>.value.length==0) 
   {   
     alert(" 主题 不能为空！"); 
     fm.InfSubj<%=sys27_Rnd(1)%>.focus();
     eflag = 1; break;
   }
   
   fm.InfCont<%=sys27_Rnd(4)%>.value = apiGetValEditID01<%=sys27_Rnd(4)%>(); 
 if (fm.InfCont<%=sys27_Rnd(4)%>.value.length==0) 
   {   
     //alert(" 内容 不能为空！"); 
     //eflag = 1; break;
   } 
   
 if (fm.InfCont<%=sys27_Rnd(4)%>.value.length>=480000) 
   {   
     alert(" 内容 不能超过 480 K字!"); 
     eflag = 1; break;
   }

         }  ////////// //////////////// End For //////////////////
         if (eflag==0){ 
		 fm.submit();
		 }
}

</script>
</body>
</html>
