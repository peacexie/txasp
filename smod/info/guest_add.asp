<%
ReEnd = Request("ReEnd") :If ReEnd="" Then ReEnd="C"
%>
<table width="100%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="F0F0F0">
  <form name="fm01<%=sys27_RVal%>" id="fm01<%=sys27_RVal%>" action="info_add2.asp?FrmFlag=<%=FrmFlag%>" enctype="multipart/form-data" method="post">
    <input name="<%=App27Random%>" type="hidden" id="<%=App27Random%>" value="<%=sys27_RVal%>">
    <tr>
      <td align="center"><p>&nbsp;</p></td>
      <td align="center"><strong>[<%=rs_Val("","SELECT SysName FROM AdmSyst WHERE SysID='"&ModID&"'")%>]增加</strong></td>
    </tr>
    <tr>
      <td align="center" nowrap>文章标题</td>
      <td><input name="InfSubj<%=sys27_Rnd(1)%>" type="text" id="InfSubj<%=sys27_Rnd(1)%>" size="60" maxlength="120">
        <input name="SetSubj" type="hidden" id="SetSubj" value="000000" /></td>
    </tr>
    <tr>
      <td align="center" nowrap>文章类别</td>
      <td><select name="InfType" id="InfType">
          <option value="<%=InfType%>"><%=Get_TypeName(ModID,InfType)%></option>
		  
        </select>
        &nbsp; &nbsp; <%=Get_Typ2Opt(ModID,InfTyp2) %> &nbsp; &nbsp;
        <input name=Button type=button id="Button2" value="选择模板" onClick="owTemp()" style="display:none" /></td>
    </tr>
    <%=tmrPara(0)%>
    <tr>
      <td align="center">文章内容</td>
      <td><div style="width:580px">
          <script type="text/javascript">
var oFCKeditor = new FCKeditor( 'EditID01<%=sys27_Rnd(4)%>' ) ;
oFCKeditor.BasePath	= oFCKeditor.BasePath.replace("sadm/edfck/","<%=Config_Path%>sadm/edfck/"); 
oFCKeditor.Height	= 360 ;
oFCKeditor.Value	= '<%=Show_jsStr(tmpCont)%>' ;
oFCKeditor.Create() ; //Server.HTMLEncode(InfCont)
</script>
          <input name="InfCont<%=sys27_Rnd(4)%>" type="hidden" value="">
        </div></td>
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
    <tr>
      <td align="center" nowrap>返回</td>
      <td><input name="ReEnd" type="radio" id="ReEnd1" value="C" <%If ReEnd="C" Then Response.Write("checked")%>>
        添加资料后关闭
        &nbsp;&nbsp;&nbsp;
        <input type="radio" name="ReEnd" id="ReEnd2" value="Y" <%If ReEnd="Y" Then Response.Write("checked")%>>
        添加资料后继续 
        <input name="SetHot" type="hidden" id="SetHot" value="N" />
        <input name="SetShow" type="hidden" id="SetShow" value="N" />
        <input name="SetTop" type="hidden" id="SetTop" value="888" />
        <input name="LogATime" type="hidden" id="LogATime" value="Now()" /></td>
    </tr>

    <tr>
      <td align="center" nowrap>认证码</td>
      <td><input name="ChkCode" type="text" id="ChkCode" size="6" maxlength="12" xonfocus="javascript:PicReLoad('../../');"/>
        <img src="../../sadm/pcode/img_frnd.asp" alt="<%=vPMsg_RelCode%>" name="ChkCImg" align="absmiddle" id="ChkCImg" style="cursor:pointer;" onClick="PicReLoad('../../')"/></td>
    </tr>

    <tr>
      <td align="center" nowrap>&nbsp;</td>
      <td><input name=view type=button id="Button1" value="确    定" onClick="chkData()">
        <input name=view type=reset id="Button1" value="重    写">
        <input name="ModImgCount" type="hidden" id="ModImgCount" value="<%=ModImgCount%>">
        <input name="PrmFlag" type="hidden" id="PrmFlag" value="<%=PrmFlag%>"></td>
    </tr>
  </form>
</table>
<script type="text/javascript">
document.getElementById('InfPara2').value = "XX单位 XX人"; // 电话:
</script>