<script src="../sadm/func1/Func_JS.js" type="text/javascript"></script>
<div class="SysCPad SysCont">
<%
If aPara(4)="0" Then aPara(4)=vPic_Pric2
If aPara(5)="0" Then aPara(5)=vPic_Pric2
%>
  <table width="100%" border="0" align="center" cellpadding="0" cellspacing="1">
    <tr>
      <td align="left" valign="top"><table width="100%"  border="0" cellpadding="1" cellspacing="2">
        <form action="ocar.asp" method="post" name="fm01" id="fm01">
          <tr>
            <td width="20%" align="right"><%=vPic_Name%>:</td>
            <td><%=InfSubj%>
              <input name="Act" type="hidden" id="Act" value="PAdd" />
              <input name="ID" type="hidden" id="ID" value="<%=ID%>" />
              <input name="MD" type="hidden" id="MD" value="<%=MD%>" />
              <input name="iSubj" type="hidden" id="iSubj" value="<%=iSubj%>" /></td>
          </tr>
          <tr>
            <td align="right"><%=vPic_Code%>:</td>
            <td><%=KeyCode%></td>
          </tr>
          <tr>
            <td align="right"><%=vPic_Speci%>:</td>
            <td><%=aPara(3)%></td>
          </tr>
          <tr>
            <td align="right"><%=vPic_From%>:</td>
            <td><%=aPara(2)%></td>
          </tr>
          <tr>
            <td align="right"><%=vPic_Pric3%>:</td>
            <td><%=aPara(5)%></td>
          </tr>
          <%If Session("MemID")&""="xxxx" Then%>
          <tr>
            <td align="right"><%=vPic_Price%>:</td>
            <td><input name="InfPrice" type="text" id="InfPrice" value="<%=aPara(4)%>" readonly="readonly" /></td>
          </tr>
          <tr>
            <th align="center" nowrap="nowrap">&nbsp;</th>
            <td align="left"><input type="button" name="button" id="button" value="<%=vPic_OLogin%>" onclick="chkLogin()" /></td>
          </tr>
          <%Else%>
          <tr>
            <td align="right"><%=vPic_Price%>:</td>
            <td><%=aPara(4)%></td>
          </tr>
          <%If rs_Exist(conn,"SELECT * FROM [OrdItem] WHERE KeyCode='---' AND LogAUser='"&Session("MemID")&"' AND InfProID='"&ID&"' ")="YES" Then%>
          <tr>
            <th align="center" nowrap="nowrap">&nbsp;</th>
            <td align="left" class="cRed f14px"><%=vPic_OAdd2%> &nbsp; <a href="ocar.asp" class="cDRed"><%=vPic_OAdd3%></a></td>
          </tr>
          <%Else%>
          <tr>
            <td align="right"><%=vPic_Price%>:</td>
            <td><input name="InfPrice" type="text" id="InfPrice" value="<%=aPara(4)%>" readonly="readonly" /></td>
          </tr>
          <tr>
            <td align="right"><%=vPic_ONum%>:</td>
            <td><input name="OrdNum" type="text" id="OrdNum" value="1" onblur="caFee(this)" /></td>
          </tr>
          <tr>
            <td align="right"><%=vPic_OMoney%>:</td>
            <td><input name="OrdSum" type="text" id="OrdSum" value="<%=aPara(4)%>" readonly="readonly" /></td>
          </tr>
          <tr>
            <td align="right"><%=vPMsg_ChkCode%>: </td>
            <td nowrap="nowrap"><span style="width:40px; display:inline-block; padding:0px 3px; border:1px solid #F0F0F0"> <img src="../sadm/pcode/img_frnd.asp?Config_PCode=hij" alt="<%=vPMsg_RelCode%>" name="ChkCImg" align="absmiddle" id="ChkCImg" style="cursor:pointer;" onclick="PicReLoad('../','hij')"/></span>
              <input name='ChkCode' type='text' id="ChkCode" size='6' maxlength='8' /></td>
          </tr>
          <%If rs_Exist(conn,"SELECT * FROM [OrdItem] WHERE KeyCode='---' AND LogAUser='("&Session.SessionID&")' AND InfProID='"&ID&"' ")="YES" Then%>
          <tr>
            <th align="center" nowrap="nowrap">&nbsp;</th>
            <td align="left" class="cRed f14px"><%=vPic_OAdd2%> &nbsp; <a href="ocar.asp" class="cDRed"><%=vPic_OAdd3%></a></td>
          </tr>
          <%Else%>
          
          <tr>
            <th align="center" nowrap="nowrap">&nbsp;</th>
            <td align="left"><input type="button" name="button" id="button" value="<%=vPic_OAdd%>" onclick="chkData()" /></td>
          </tr>
          <%End If%>
          <%End If%>
          <%End If%>
          <input name="goPage" type="hidden" value="goRef" />
        </form>
      </table></td>
      <td width="50%" height="200" align="center" style="padding:8px"><a href='<%=aImgs(1)%>' target="_blank"><img src='<%=aImgs(1)%>' width="300" height="240" border="0" onload="javascript:setImgSize(this);" /></a></td>
    </tr>
    <tr>
      <td align="left" valign="top" colspan="2"><%Call Show_sfData(ID,"fcont.htm")%></td>
    </tr>
  </table>
</div>
<script type="text/javascript">

var fm = document.fm01;

function chkLogin() 
{
	fm.action = "../member/login.asp";
	fm.submit();
}


function chkData(xID) 
{
       var eflag = 0;
       for(i=0;i<1;i++)
         {  ////////// //////////////// Start For ////////////////
 if (fm.ChkCode.value.length==0) 
   {   
     alert("<%=vPMsg_ChkCErr%>"); 
     fm.ChkCode.focus();
     eflag = 1; break;
   }
         }  ////////// //////////////// End For //////////////////
         if (eflag==0){ fm.submit(); }
}

function caFee(e)
{
  
  var ev = e.value;
  var vp = fm.InfPrice.value;
  if ( isNaN(ev) || ev.indexOf(".")>=0 ){
    e.value = 1; ev = 1;
	alert("<%=vPic_jsInt%>");
  }
  
  if ((chkF_Int(fm.OrdNum,'')=='ER')||(fm.OrdNum.value<=0)) {
    e.value = 1; ev = 1;
	alert("<%=vPic_jsInt%>");
  }
  
  if( isNaN(vp) ){
	 fm.OrdSum.value = "<%=vPic_Pric2%>";  
  }else{
	 var v = ev*vp;
	 fm.OrdSum.value = v.toFixed(2);
  }

}
</script>
<div style="clear:both"></div>