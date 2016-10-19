
  showPage(0);
  
function showPage(n)
{
  // Set Para    
  var divData = getElmID('pageData'); var sData = divData.innerHTML; 
  var divPOne = getElmID('pageOne');  //alert(sData);
  var divPBox = getElmID('pageBox');
  var divPBar = getElmID('pageBar');
  //*
  // var flagPage = new RegExp('<P style="display:none;">.dg.gd.cn.Pager.</P>','i'); for dg.gd.cn
  var flagPage = new RegExp('<div style="page-break-after: always"><span style="display: none">&nbsp;</span></div>','i')
  sData = sData.replace('style="page-break-after: always;"','style="page-break-after: always"',"gi")
  sData = sData.replace('style="display: none;"'           ,'style="display: none"',"gi");
  if(flagPage.test(sData))
  {
	var aData = sData.split(flagPage);
	var sBar = ''; //alert(sData);
	divData.style.display = 'none';
	divPOne.style.display = '';
	divPBox.style.display = '';
	divPOne.innerHTML = aData[n];
	// Show Pager Bar &lt;&gt;
    //for(i=0;i<aData.length;i++){
	m = aData.length;
	if(m<12) {
	  n1 = 0; 
	  n2 = m-1; 
	} else {
	 if(n>=m-5) { n1=m-11; n2=m-1; }
	 else if(n<=5) { n1=0; n2=10; }
	 else{ n1=n-5; n2=n+5; }
	}
	for(i=n1;i<=n2;i++){	
	 if((i>=0)&&(i<m)){
	  j = i+1;  //页码 
	  if(n==i) { sBar += '<SPAN class=pageNow>'+j+'</SPAN>'; } 
	  else     { sBar += '<SPAN><A class=pageLink onclick="showPWait('+i+')" href="#_InfSubj">'+j+'</A></SPAN>'; }
	 }
	}
	i = m - 1;  // n=[0~i]
	j = n-1; // 上1页
	if(j>=0) { sBar = '<SPAN><A class=pageLink onclick="showPWait('+j+')" href="#_InfSubj">上一页</A></SPAN>'+sBar; } 
	else     { sBar = '<SPAN class=pageGray>上一页</SPAN>'+sBar; }
	j = n+1; // 下1页
	if(j<=i) { sBar = sBar+'<SPAN><A class=pageLink onclick="showPWait('+j+')" href="#_InfSubj">下一页</A></SPAN>'; } 
	else     { sBar = sBar+'<SPAN class=pageGray>下一页</SPAN>'; }
	if(m>11) {
	  if(n==0) { sBar = '<SPAN class=pageGray>第1页</SPAN>'+sBar; } 
	  else     { sBar = '<SPAN><A class=pageLink onclick="showPWait(0)" href="#_InfSubj">第1页</A></SPAN>'+sBar; }
	  j = m;
	  if(n==i) { sBar = sBar+'<SPAN class=pageGray>第'+j+'页</SPAN>'; } 
	  else     { sBar = sBar+'<SPAN><A class=pageLink onclick="showPWait('+i+')" href="#_InfSubj">第'+j+'页</A></SPAN>'; }
	}
	//sBar = '<div style="width:3px; float:left;"></div>'+sBar;  <A href="#_InfSubj" ></A>
	divPBar.innerHTML = sBar;
  }
  //*/

}

function showPWait(n)
{  
  setTimeout("showPage("+n+")",10); 
}

