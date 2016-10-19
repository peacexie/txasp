
function chkSOut(){   
  var nTels=0,nCont,ta,eflag=1;
  for(k=0;k<1;k++){  
   //手机号码
  if (fm.tNumb.value.length==0){ 
     alert('[错误]\n 请输入 手机号码!');
     fm.tNumb.focus();
     eflag=0; break;
  }
  fmtTels();
  ta = String(fm.tNumb.value).split("\n");
  if(telMaxs<ta.length){ alert("号码太多,不能超过"+telMaxs+"个"); eflag=0; break; }
  for(j=0;j<ta.length;j++){ 
	if(String(ta[j]).length>5) { nTels++; }
  }
  //One短信内容
   if (fm.tCont1.value.length==0){ 
	 alert('[错误]\n 请输入 短信内容!');
	 fm.tCont1.focus();
	 eflag=0; break;
   }
   if (fm.tCont1.value.length>70){ 
	 alert('[错误]\n 短信内容 请少于70字!');
	 fm.tCont1.focus();
	 eflag=0; break;
   }
   nCont = 1;
  //检查余额
  nMax = fm.tBalance.value;
  if (nTels*nCont>nMax){ 
     alert("[错误] 余额不足!\n当前余额为:["+nMax+"]\n需要余额为:["+nTels+"]");
	 eflag=0; break;
  }
  //检查内容非法关键字
  var k = chkCont1(fm.tCont1.value);
  if(k!=""){ 
    alert("短信内容 含有敏感关键字:"+k+"");
	eflag=0; break;
  }

  }
  if (eflag==1){ 
    fm.tel.value = fm.tNumb.value;
	fm.ct1.value = fm.tCont1.value;
    fm.submit(); 
  }
}

var nowSign = "";
function setSign(e){
  if(e.value=="") {return;} //(选择签名)空值退出
  var vt = e.options[e.selectedIndex].text;
  var v1 = " ("+vt+")";
  var v2 = "("+vt+")";
  var v3 = "("+vt+"";
  var ec = fm.tCont1;
  var m = 70;
  if(fm01.tMode.value=="More"){ 
    var n = objNO.innerHTML-1;
	ec = document.getElementById("tCont"+n);
  }
  if(fm01.tMode.value=="Long"){ 
    m = 255;
  }  
  if(nowSign.length>0){ //替换旧的签名
    var p = ec.value.lastIndexOf(nowSign);
	if(p>0){
	  var t = ec.value.substring(p);
	  if(t==nowSign) ec.value = ec.value.substring(0,p);	
	} 
  }else{
	//第一个签名
	var p = ec.value.lastIndexOf("("+e.options[1].text);
	if(p>0){
	   var t = ec.value.substring(p); 
	   ec.value = ec.value.substring(0,p);	
	   if((t==v1)||(t==v2)||(t==v3)){ nowSign = t; }
	}
  }
  if(ec.value.length>0){ 
	oldSign = nowSign;
	if(ec.value.length<=m-v1.length){
	  nowSign = v1; ec.value += nowSign; 
	}else if(ec.value.length<=m-v2.length){
	  nowSign = v2; ec.value += nowSign; 
	}else if(ec.value.length<=m-v3.length){
	  nowSign = v3; ec.value += nowSign; 
	}else{
	  var msg = "内容太长,此签名不适用!";
	      msg+= "\n短信内容最多:"+m+"字,";
		  msg+= "\n当前内容长度:"+ec.value.length+"字,";
		  msg+= "\n签名至少需要:"+v3.length+"字.";
	  alert(msg);	
	}
  }else{
	alert('请先填写内容,再选签名!');  
  }
}

function chkGroup(){ 
  ; // ...
  var nTels=0,nCont,ta,eflag=1;
  if (eflag==1){ fm.submit(); }
}

function chkSend(){   
  var nTels=0,nCont,ta,eflag=1;
  for(k=0;k<1;k++){  
   //手机号码
  if (fm.tNumb.value.length==0){ 
     alert('[错误]\n 请输入 手机号码!');
     fm.tNumb.focus();
     eflag=0; break;
  }
  fmtTels();
  ta = String(fm.tNumb.value).split("\n");
  if(telMaxs<ta.length){ alert("号码太多,不能超过"+telMaxs+"个"); eflag=0; break; }
  for(j=0;j<ta.length;j++){ 
	if(String(ta[j]).length>5) { nTels++; }
  }
  
  // 短信内容
  if(!chkSCnts(nTels)){
	eflag=0; break;  
  }

  }
  if (eflag==1){ fm.submit(); }
}
function chkSCnts(nTels){ 
  //One短信内容
  eflag = true;
  var m = fm01.tMode.value;
  if(m=="One"){ 
	 if (fm.tCont1.value.length==0){ 
	   alert('[错误]\n 请输入 短信内容!');
	   fm.tCont1.focus();
	   eflag=false;
	 }
	 if (fm.tCont1.value.length>70){ 
	   alert('[错误]\n 短信内容 请少于70字!');
	   fm.tCont1.focus();
	   eflag=false;
	 }
	 nCont = 1;
  }
  //More短信内容
  if(m=="More"){ 
	 var n = objNO.innerHTML; 
	 var t = 0;
	 for(i=1;i<n;i++){
	   var ct = document.getElementById("tCont"+i);
	   if (ct.value.length>70){ 
		 alert('[错误]\n 短信内容 请少于70字!');
		 ct.focus();
		 eflag=false;
	   }else{
		 if(ct.value.length===0) t++;
	   }
	   if(t==n-1){ 
		 alert('[错误]\n 请输入 短信内容!');
		 fm.tCont1.focus();
		 eflag=false;
	   }
	 }
	 nCont = n-1-t;
  }
  //Long短信内容
  if(m=="Long"){ 
	 if (fm.tCont1.value.length==0){ 
	   alert('[错误]\n 请输入 短信内容!');
	   fm.tCont1.focus();
	   eflag=false;
	 }
	 if (fm.tCont1.value.length>255){ 
	   alert('[错误]\n 短信内容 请少于255字!');
	   fm.tCont1.focus();
	   eflag=false;
	 }
	 nCont = fm.tCont1.value.length; //alert(nCont);
	 nCont = Math.ceil(nCont/65); 
  }
  //检查余额
  nMax = fm.tBalance.value;
  if (nTels*nCont>nMax){ 
     alert("[错误] 余额不足!\n当前余额为:["+nMax+"]\n需要余额为:["+nTels*nCont+"]=["+nTels+"个号码]x["+nCont+"条信息]");
	 eflag=false;
  }
  //检查内容非法关键字
  var k = chkCont1(fm.tCont1.value);
  if(k!=""){ 
    alert("短信内容 含有敏感关键字:"+k+"");
	eflag=false;
  }
  return eflag;
}

function chkCont1(str){ 
  f = ""; if(str.length==0) return "";
  a = keyCont.split(";");
  for(i=1;i<=a.length;i++){
    var k = String(a[i]);
	if(k.length>1){
	  if(str.indexOf(k)>=0)	{
		f = k;
		break;   
	  }
	}
  }
  return f;
}

function openWin(xUrl) 
{
   var rWin = String(Math.random()).replace(".","");
   window.open(xUrl,"Win_"+rWin+"",'height=480,width=640,toolbar=no,menubar=no,scrollbars=yes,resizable=yes,location=no,status=no');
}

function fmtTels() {
	s = fm.tNumb.value;
    var e1 = new RegExp("\ |\t",'gi'); //删除空格
	var e2 = new RegExp("\r",'gi'); //转化\r
	var e3 = new RegExp("\,|\;|\，|\；",'gi');
	var e4 = new RegExp("\n\n\n\n|\n\n\n|\n\n",'gi'); //多个\n
	s = s.replace(e1,"");
	s = s.replace(e2,"\n");
	s = s.replace(e3,"\n");
	s = s.replace(e4,"\n");
	fm.tNumb.value = s;
}

function getTAdd() {
	fm.tNumb.value += "\n"+telCont;
	fmtTels();
}
function getTels() {
	fm.tNumb.value = telCont;
	fmtTels();
}

function getTemp() {
	bakCont(4);
	fm.tCont1.value = tmpCont;
	chkChars(fm.tCont1,'nowChars',70);
	rebCont(4);
}

// 还原内容框
function rebCont(N){
  var f = 0;
  for(i=1;i<=N;i++){
	var no = i;
	try{
	  var t = document.getElementById("tCont"+no);
	  var b = document.getElementById("bCont"+no);
	  if((t.value.length==0)){
		  t.value = b.innerText.replace(/\[备份内容[1|2|3|4]\] /,"");
		  f++; 
	  }
	}catch(e){  }
  }
  //if(f>0){alert("注意：已经输入的短信内容 已经备份；\n请在 说明/提示 后面查看[备份内容]！");}
}

// 备份内容框
function bakCont(N){
  var f = 0;
  for(i=1;i<=N;i++){
	var no = i;
	try{
	  var t = document.getElementById("tCont"+no);
	  var b = document.getElementById("bCont"+no);
	  if(t.value.length>0){
		  b.innerHTML = "<span class='fntF0F'>[备份内容"+no+"]</span> "+t.value;
		  f++; 
	  }
	}catch(e){}
  }
  //if(f>0){alert("注意：已经输入的短信内容 已经备份；\n请在 说明/提示 后面查看[备份内容]！");}
}

// 增加内容框
function addCont(){
  bakCont(4);
  n = objNO.innerHTML;
  setConts(n);
  rebCont(4);
}

// 内容模式
function setMode(){
  bakCont(4);
  var nVal = fm01.tMode.value;
  objBar.style.display = 'none';
  if(nVal=="One"){ 
    objConts.innerHTML = tmp.replace("内容(n): ","")
  }
  if(nVal=="More"){ 
	setConts(2);  
	objBar.style.display = '';
  }
  if(nVal=="Long"){ 
	var s = tmp.replace("4","10");
	objConts.innerHTML = s.replace("内容(n): ","").replace("70","255").replace("70","255");
  }
  rebCont(4);
}

// 设置N个内容框
function setConts(n){
  objConts.innerHTML = "";
  for(i=0;i<n;i++){
	j = i+1; 
	si = tmp.replace("tCont1","tCont"+j).replace("tCont1","tCont"+j).replace("(n)","("+j+")");
	objConts.innerHTML += si;
  }
  objNO.innerHTML = j+1; 
  if(j+1>cnMax) objBar.style.display = 'none';
}

function chkChars(tObj,cShow,xMax){
  var n = tObj.value.length;
  document.getElementById(cShow).innerHTML = "已经输入["+n+"]个字";
  if (n>xMax) { alert("当前内容为["+n+"]个字，\n多于"+xMax+"，请重新编辑！"); }
}


