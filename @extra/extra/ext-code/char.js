function CharDel(xa,xb) 
{
  for(var i=0;i<xb.length;i++)
  {
    xa=xa.replace(xb.substring(i,i+1),'');
  }
  return xa;
}

function CharComp(xa,xb) 
{
  var s='',p=0,c='';
  for(var i=0;i<xb.length;i++)
  {
    c=xb.substring(i,i+1);
	p=xa.indexOf(c);
	if(p>=0){ s+=xa.substring(p,p+1);}
  }
  return s;
}

function CharAdd(xa,xb) 
{
  var s=xa,p=0,c='';
  for(var i=0;i<xb.length;i++)
  {
    c=xb.substring(i,i+1);
	p=s.indexOf(c);
	if(p<0){ s+=c;}
  }
  return s;
}