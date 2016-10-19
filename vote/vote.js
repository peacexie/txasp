
function setImgSize(xImg,xWidth,xHeight)
{
 var MaxWidth=xWidth;          var MaxHeight=xHeight; 
 var ObjImg=new Image();       ObjImg.src=xImg.src; 
 var ImgWidth=ObjImg.width;    var ImgHeight=ObjImg.height; 
 var RatHW=ImgHeight/ImgWidth; var RatWH=ImgWidth/ImgHeight;
 if(ImgWidth>MaxWidth){
    ImgWidth=MaxWidth;   ImgHeight=MaxWidth*RatHW;
 }
 if(ImgHeight>MaxHeight){
    ImgHeight=MaxHeight; ImgWidth=MaxHeight*RatWH;
 }
    xImg.width=ImgWidth; xImg.height=ImgHeight;
}

/*/////////////////*/

function VoteItem(obj)
{
    obj.checked ? VotCount++ : VotCount--;
    if (!VoteCMax())
    {
        VotCount--;
        return false;
    }
    return true;
}
function VoteCMax()
{
    if (VotCount > InfNum2)
    {
        alert('对不起，最多只能选择 ' + InfNum2 + ' 位候选人。');
        return false;
    }
    return true;
}

function chkData()
{
       var fmObj = document.fm01;	
	   var eflag = 0;
       for(i=0;i<1;i++)
         {  ////////// //////////////// Start For ////////////////
 if (VotCount<InfNum1) 
   {   
     alert('对不起，最少要选择 ' + InfNum1 + ' 位候选人。');
     eflag = 1; break;
   }
 if(InfCard=='Y')
  {   
    if (fmObj.InfCard.value.length<15)
    {   
      alert(" 身份证 至少15位！"); 
      fmObj.InfCard.focus();
      eflag = 1; break;
    }
    if (fmObj.InfTel.value.length<7)
    {   
      alert(" 电话证 至少7位！"); 
      fmObj.InfTel.focus();
      eflag = 1; break;
    }
    if (fmObj.InfName.value.length<2)
    {   
      alert(" 姓名 至少2个字！"); 
      fmObj.InfName.focus();
      eflag = 1; break;
    }
	
   if (fmObj.ChkCode.value.length<2)
    {   
      alert(" 请填写 认证码！"); 
      fmObj.ChkCode.focus();
      eflag = 1; return; //break;
    }
	
   }
         }  ////////// //////////////// End For //////////////////
         if (eflag==0){ fmObj.submit(); }
}