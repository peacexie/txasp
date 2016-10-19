
//****************** 选项切换代码-主菜单 */
function mTopOver(e){
  var mID = e.id.toString();
  if(mID.length>5)
    if(e.className=="mTopComm"){ 
	  e.className="mTopAct";
	  if(isFirefox||isIE8||isIE9){
		e.setAttribute("onmouseout","this.className='mTopComm';");
	  }else{
		e.setAttribute("onmouseout",function(){eval("this.className='mTopComm';")});
	  }
	}
}
function mTypOver(e){
  var mID = e.id.toString().substring(3); // id_InnStop-=> InnStop
  if(mID.length>5){
	a = sDeps.split(";");
	for(var i=0;i<a.length;i++){ 
	  var idBox = getElmID("div_" + a[i]);
	  var idTyp = getElmID("id_" + a[i]);
	  if (mID==a[i]){
		idBox.style.display='';
		idTyp.className = 'mTypAct';
	  }else{
		idBox.style.display='none';
		idTyp.className = 'mTypComm';
	} }
  }
}
//****************** 选项切换代码-部门 */
function mTypChang(xDiv)
{
  s = sDeps;
  a = s.split(";");
  for(var i=0;i<a.length;i++){ 
    var idBox = getElmID("div_" + a[i]);
	var idTyp = getElmID("id_" + a[i]);
	if (xDiv==a[i]){
	  idBox.style.display='';
	  idTyp.className = 'mTypAct';
	}else{
	  idBox.style.display='none';
	  idTyp.className = 'mTypComm';
  } }
}


function ChkAll(xTyp)
{
 // Pub; All; Nul
 var e = document.fm01.InfView;
 var v = true; var f = false;
 if(xTyp=='Pub'){ v=false; f = true; }
 if(xTyp=='All'){ v=true;  f = false; }
 if(xTyp=='Nul'){ v=false; f = false; }
 for(var i=0;i<e.length;i++)
 {
   e.item(i).checked = v;
   e.item(i).disabled = f;
 }
  var e2 = document.fm01.InfVGrps; 
  if(xTyp=='Pub'){
    v = false; f = true;
  }else{
	v = false; f = false;
  }
  for(var i=0;i<e2.length;i++){
      e2.item(i).checked = v;
      e2.item(i).disabled = f; 
  }
}
function ChkS99(e,xPStr)
{
 var v = true; 
 if(e.checked==false){ v=false; }
 var e2 = document.fm01.InfView;
 for(var i=0;i<e2.length;i++)
 {
   if (xPStr.indexOf(e2[i].value)>=0)
   e2.item(i).checked=v;
 }
}

