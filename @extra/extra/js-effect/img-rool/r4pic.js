
function getElmID(xID){
  return document.getElementById(xID);
}
function ShowDiv(xID)
{
 nDiv = getElmID(xID);
 if(nDiv.style.display=='none') { nDiv.style.display = ''; }
 else { nDiv.style.display = 'none'; } 
}  

function RoolNOver()
{
  RoolNFlag = "Stop";
}
function RoolNOut()
{
  RoolNFlag = "Start";
}
function RoolNSet(n)
{
  RoolNNow = n;
}

function RoolNSwh(n)
{
   
  if(RoolNFlag=="Start"){
    RoolNNow++; 
    if(RoolNNow==RoolNID.split("|").length) { RoolNNow = 1; }
    RoolNShow(RoolNNow);
  }
  
  //getElmID("tstID").innerHTML = getElmID("tstID").innerHTML + RoolNNow;
  setTimeout("RoolNSwh("+RoolNNow+")",3000); 
  
}
function RoolNOpen(n) {
  aID = RoolNID.split("|");
  window.open("nview.asp?KeyID="+aID[u-1]+"","winRoolNOpen","");
}
function RoolNShow(n) {
  aID   = RoolNID.split("|");
  aPic  = RoolNPic.split("|");
  aSubj = RoolNSubj.split("|");
  for(i=1; i<RoolNLen; i++) {
	  if(i==n){
	    getElmID("RoolNP"+i).className="RoolNa";
		var img = "<img src='"+aPic[i-1]+"' alt='"+aSubj[i-1]+"' width='240' height='180' border='0' />";
	    getElmID("RoolNImg").innerHTML = "<a href='nview.asp?KeyID="+aID[i-1]+"' target='_blank'>"+img+"</a>"; 
	  } else {
	    getElmID("RoolNP"+i).className="RoolNb";
	  }
  }
}