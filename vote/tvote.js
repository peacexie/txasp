
function CheckItem(obj)
{
    //obj.checked ? VotCount++ : VotCount--;
    if (fVote=='False')
    {
        alert('无效！此IP可能已经有投票!');
		return false;
    }
    return true;
}

function chkData()
{
 var fmObj = document.fm01;	
 var eflag = 0; var  iVal, sVal='';
 var aItem = jsItem.split('|');
 var aITyp = jsITyp.split('|');
 
 
 for (var i=0;i<aItem.length-1;i++)
 {
   iMsg = document.getElementById('Msg'+aItem[i]).innerHTML;
   iVal = ''; /////////////////////////////// 清空当前项目值;
	
	
   if(aITyp[i]=='Select')
   {
     var fmBox = fmObj.elements['iRad'+aItem[i]];
	 j=0;
     for(var k=0;k<fmBox.length;k++)   
     {   
       if(fmBox[k].checked){
		 iVal += '1'; j++; 
	   }else{
		 iVal += '0';
	   }
     }
	 if(j==0) {
		eflag=1; 
		alert(iMsg+'\n为必须回答(选择)项目!');
		break;
	 }
   }
	 
   if(aITyp[i]=='FBlank')
   {
     var fmTxt = fmObj.elements['iTxt'+aItem[i]];
	 iVal = fmTxt.value;
	 if(iVal.length==0) {
		eflag=1; 
		alert(iMsg+'\n为必须回答(填写)项目!');
		fmTxt.focus();
		break;
	 }
	 iVal = iVal.replace(';','；');
   }
   
   if(aITyp[i]=='GBlank')
   {
     var fmTxt = fmObj.elements['iTxt'+aItem[i]];
	 iVal = fmTxt.value.replace(';','；');
   }
   
   if( (aITyp[i]=='FBlank')||(aITyp[i]=='GBlank') )
   {
   	 if(iVal.length>120) {
		eflag=1; 
		alert(iMsg+'\n填写内容请在120字以内!');
		fmTxt.focus();
		break;
	 }
	 iVal = ';'+iVal+';';
   }
   
   sVal += iVal;  //////////////////////// 连接当前项目值; 
 }  ////////////////////////////////////////// End For 
   sVal = sVal.replace(';;',';');
 

   if (eflag==1){ return; }
   
 if(InfCard=='Y')
 {  
   if (fmObj.InfName.value.length<2)
    {   
      alert(" 姓名 至少2个字！"); 
      fmObj.InfName.focus();
      eflag = 1; return; //break;
    }
	
   if (fmObj.InfTel.value.length<7)
    {   
      alert(" 电话证 至少7位！"); 
      fmObj.InfTel.focus();
      eflag = 1; return; //break;
    }
	
   if (fmObj.InfCard.value.length<15)
    {   
      alert(" 身份证 至少15位！"); 
      fmObj.InfCard.focus();
      eflag = 1; return; //break;
    }
	
   if (fmObj.ChkCode.value.length<2)
    {   
      alert(" 请填写 认证码！"); 
      fmObj.ChkCode.focus();
      eflag = 1; return; //break;
    }
	
 }

 //alert(sVal); // 测试使用
 if (eflag==0){ fmObj.vItems.value=sVal; fmObj.submit(); }
} //////////////////////////////////////////// End chkData()


