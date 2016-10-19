// JavaScript Functions v1.5; ---Peace(XieYongshun)2005/10 

function getXmlHttp() {
  var xmlHttp = false;
  try {
     xmlHttp = new XMLHttpRequest();
  } catch (trymicrosoft) {
    try {
       xmlHttp = new ActiveXObject("Msxml2.XMLHTTP");
    } catch (othermicrosoft) {
      try {
        xmlHttp = new ActiveXObject("Microsoft.XMLHTTP");
      } catch (failed) {
        xmlHttp = false;
      }
    }
  }
  if (!xmlHttp){
    alert("无法创建 XMLHttpRequest 对象！");
  }
  return xmlHttp;
}

function chkF_Mail(MObj,msg)
{
    var Mstr,Mi,Mj,Mk,Mkk,Mjj,Mlen;
    Mstr = MObj.value;
    Mat  = Mstr.indexOf( "@" );
    Mdot = Mstr.indexOf( ".", Mi );
    Md2  = Mstr.indexOf( "," );
    Mblk = Mstr.indexOf( " " );
    Mext = Mstr.lastIndexOf( "." ) + 1;
    Mlen = Mstr.length;
    if ( (Mat<=0)||(Mdot<=2)||(Md2!=-1)||(Mblk!=-1)||(Mlen-Mext<2)||(Mlen-Mext>3) )
	{   jsFlag = "ER";
	}else{
		jsFlag = "OK";};
    if( (jsFlag=="ER")&&(msg!="") )
	{  
	    alert(msg);
		MObj.focus();
	}
	return jsFlag;
}

function chkF_Date(MMObj,msg)
{  
	var MYY,MMM,MDD,MObj ;
	jsFlag = "OK";
	MObj = MMObj.value;
	MObj = MObj.replace("/", "-"); 
	MObj = MObj.replace(".", "-"); 
	strArr=MObj.split("-");
    YY = strArr[0]; //MObj.substr(0,4); //alert(YY);
	MM = strArr[1]; //MObj.substr(5,2); //alert(MM);
	DD = strArr[2]; //MObj.substr(8,2); //alert(DD);
	//if ( !(MObj.length==10)) { jsFlag = "ER"; }
	if ( isNaN(YY) ) { jsFlag = "ER"; 
	} else {
	   if ( (YY<1000) || (YY>9999) ) { jsFlag = "ER";}   }
	if ( isNaN(MM) ) { jsFlag = "ER"; 
	} else {
	   if ( (MM<1) || (MM>12) ) { jsFlag = "ER";}   }
	if ( isNaN(DD) ) { jsFlag = "ER"; 
	} else {
	   if ( (DD<1) || (DD>31) ) { jsFlag = "ER";}   }
	if (MM==2) {
	  if (  ((0 == YY % 4) && (0 != (YY % 100))) || (0 == YY % 400)  ) {
		  if (DD>29) {  jsFlag = "ER";}
	  }else{
		  if (DD>28) {  jsFlag = "ER";}
	  }    
	}else{
	  if (  (MM==4)||(MM==6)||(MM==9)||(MM==11)  ) {
		  if (DD>30) {  jsFlag = "ER";}
	  }else{
		  if (DD>31) {  jsFlag = "ER";}
	  }    
	} 
    if( (jsFlag=="ER")&&(msg!="") )
	{  
	    alert(msg);
		MMObj.focus();
	}
	return jsFlag;
}

function chkF_DCmp(DO1,DO2,msg)
{  
    var dt1,dt2,a1,a2,y1,y2,m1,m2,d1,d2;
	jsFlag = "OK";
	dt1 = DO1.value; 
	dt2 = DO2.value;
	dt1 = dt1.replace("/", "-"); 
	dt1 = dt1.replace(".", "-"); 	
	dt2 = dt2.replace("/", "-"); 
	dt2 = dt2.replace(".", "-"); 
	a1 = dt1.split("-");
	a2 = dt2.split("-");
    y1 = a1[0]; 
	m1 = a1[1]; if (m1.length==1)  {  m1 = "0" + m1; }
	d1 = a1[2]; if (d1.length==1)  {  d1 = "0" + d1; }
    y2 = a2[0]; 
	m2 = a2[1]; if (m2.length==1)  {  m2 = "0" + m2; }
	d2 = a2[2]; if (d2.length==1)  {  d2 = "0" + d2; }
    if (y1>y2) {  jsFlag = "ER"; }
	else       {  m1+d1>m2+d2;   } 
	if ( (y1==y2) && (m1+d1>m2+d2) ) {  jsFlag = "ER"; }
	//alert(y1+m1+d1+' '+y2+m2+d2);
    if( (jsFlag=="ER")&&(msg!="") )
	{  
	    alert(msg);
		DO1.focus();
	}
	return jsFlag;
}

function chkF_ID(PKObj,msg)
{
   var PKLen,PKStr,PKi,PKChr;
   PKLen = PKObj.value.length;
   PKStr = "";
   jsFlag = "OK";
   if ( PKLen==0 ) { jsFlag = "ER" ;}
   for (PKi=0;PKi<PKLen;PKi++)
   {
      PKChr = PKObj.value.substring(PKi,PKi+1);
      if ( ((PKChr>='A')&&(PKChr<='Z')) || ((PKChr>='a')&&(PKChr<='z')) || ((PKChr>='0')&&(PKChr<='9')) || (PKChr=='.') || (PKChr=='_') || (PKChr=='-') || (PKChr=='@') )
      {         PKStr += PKChr;
        }else{ jsFlag = "ER";break;     
      }
   }
   PKChr = PKObj.value.substring(0,1);
   PKObj.value = PKStr; 
   if(!(  ((PKChr>='A')&&(PKChr<='Z')) || ((PKChr>='a')&&(PKChr<='z')) || ((PKChr>='0')&&(PKChr<='9'))  )) 
	  { jsFlag = "ER"; }
    if( (jsFlag=="ER")&&(msg!="") )
	{  
	    alert(msg);
		PKObj.focus();
	}
   return jsFlag;
} 

function chkF_Tel(TObj,msg)
{
   var TLen,TStr,Ti,TChr,chNum;
   TLen = TObj.value.length;
   TStr = "";
   jsFlag = "OK";
   chNum  = 0;
   if ( TLen==0 ) { jsFlag = "ER" ;}
   for (Ti=0;Ti<TLen;Ti++)
   {
      TChr = TObj.value.substring(Ti,Ti+1);
      if ( ((TChr>='0')&&(TChr<='9')) || (TChr=='.') || (TChr=='_') || (TChr=='-') || (TChr=='(') || (TChr==')') )
      {         TStr += TChr;    if ((TChr>='0')&&(TChr<='9')) {chNum++;}
        }else{ jsFlag = "ER";break;     
      }
   }
   if (chNum<5) {jsFlag = "ER";}
   TObj.value = TStr; 
    if( (jsFlag=="ER")&&(msg!="") )
	{  
	    alert(msg);
		TObj.focus();
	}
   return jsFlag;
}

function chkF_PW(PKObj,msg)
{
   var PKLen,PKStr,PKi,PKChr;
   PKLen = PKObj.value.length;
   PKStr = "";
   jsFlag = "OK";  // _  "  '  /  \  ?  &  <  >
   if ( PKLen==0 ) { jsFlag = "ER" ;}
   for (PKi=0;PKi<PKLen;PKi++)
   {
      PKChr = PKObj.value.substring(PKi,PKi+1);
      if (  (PKChr==' ') || (PKChr=='&')  || (PKChr=='<') || (PKChr=='>') || (PKChr=='=') || (PKChr=='\'') || (PKChr=='\"')   )
	  {         jsFlag = "ER"; break; 
        }else{ jsFlag = "OK"; PKStr += PKChr;    
      }
   }
   PKObj.value = PKStr; 
    if( (jsFlag=="ER")&&(msg!="") )
	{  
	    alert(msg);
		PKObj.focus();
	}
   return jsFlag;
} 

function chkF_Dot(NDObj,msg)
{  jsFlag = "OK";
   if ( isNaN(NDObj.value) )
   {  jsFlag = "ER";
      NDObj.value = 0.0;
   }  
    if( (jsFlag=="ER")&&(msg!="") )
	{  
	    alert(msg);
		NDObj.focus();
	}
   return jsFlag;
} 

function chkF_Int(NIObj,msg)
{  jsFlag = "OK";
   if ( isNaN(NIObj.value) || NIObj.value.indexOf(".")>=0 )
   {  jsFlag = "ER";
      NIObj.value = 0;
   }  
    if( (jsFlag=="ER")&&(msg!="") )
	{  
	    alert(msg);
		NIObj.focus();
	}
   return jsFlag;
} 

function chkF_Blank(GObj,msg)
{  jsFlag = "OK";
   if ( GObj.value.length==0 ){ F = "ER";}
    if( (jsFlag=="ER")&&(msg!="") )
	{  
	    alert(msg);
		GObj.focus();
	}
   return jsFlag;
}

function chkF_Len(LNObj,LNLen,msg)
{  jsFlag = "OK";
   if ( LNObj.value.length>LNLen )
   {  jsFlag = "ER";
      LNObj.value = LNObj.value.substr(0,LNLen-3)+"..."
   }   
    if( (jsFlag=="ER")&&(msg!="") )
	{  
	    alert(msg);
		LNObj.focus();
	}
   return jsFlag;
}

function chkF_Pub(frmObj,frmIteN,frmMsg,frmItem)
{  
    frmflag=1;
    frm_Item = frmItem.split(";");
    frm_Msg  = frmMsg.split(";");    
    frmn = frmIteN ;//n = frm_Msg.dimensions();//ubound();
    for(frmi=0;frmi<frmn;frmi++)
    {
       if (frm_Item[frmi].length==0)
       {
          alert(frm_Msg[frmi]);
          frmflag=0;
          break;
       } 
    }
   if (frmflag==1) frmObj.submit()
}

function Del_YN(YNaddr,msg)
{
    if(confirm(msg))
     {
         location.href = YNaddr;
         return true;
     }
         return false;
}

function Dir_Addr(Direct)
{
    location.href = Direct; //YNaddr;
}


function GetUrlPara(qs)
{
    var s = location.href;
    s = s.replace("?","?&").split("&");
    var re = "";
    for(i=1;i<s.length;i++)
        if(s[i].indexOf(qs+"=")==0)
            re = s[i].replace(qs+"=","");
    return re;
}