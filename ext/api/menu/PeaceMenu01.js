function menu01Show(xID){
  //return;
  var a = "Home;About;Xiag;Guide".split(";");
  for(var i=0;i<a.length;i++){
  try{
    ItemID = getElmID("menu01"+a[i]+"");
	IDivID = getElmID("menu02"+a[i]+"");
	if (xID=="menu01"+a[i]){
	  IDivID.style.display='';
	  IDivID.style.visibility='visible';
      ItemID.className="menu01Act";
	  //getElmID("idTest").innerHTML = xID;
	}else{
	  IDivID.style.display='none';
	  IDivID.style.visibility='hidden';
	  ItemID.className="";
    } 
  }catch(objError){}}
}

function menu01Page(){ // ´¦Àí²Ëµ¥CSS£»	 
  setEvent("onmouseover","muOver","menu01Tags","td");
  var muThisUrl = self.location; //alert(muThisUrl);
  //try{
	muThisUrl = ""+muThisUrl+""; //alert(muTishUrl);
	if(muThisUrl.indexOf("/page/index.asp")>0){
	  getElmID("menu01Home").className = "menu01Act";
	}
	if(muThisUrl.indexOf("/mreg/")>0){
	  menu01Show("InfN152");
	}
	
	if(muThisUrl.indexOf("/page/imod.asp")>0){
	  menu01Show(MD); //alert('1');
	}
	if(muThisUrl.indexOf("/page/info.asp")>0){
	  menu01Show(MD);
	}
	if(muThisUrl.indexOf("/page/iview.asp")>0){
	  menu01Show(MD);
	}
  
	if(muThisUrl.indexOf("/page/gbook.asp")>0){
	  menu01Show(MD); //alert('1');
	}
	if(muThisUrl.indexOf("/ext/r")>0){
	  menu01Show("BBSVB24"); //alert('1');
	}
  //}catch(objError){ alert("m"+MD); }  
}
