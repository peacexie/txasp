
var Play02Speed = 2123; //轮播间隔时间
var Play02Link = true; // true/false 有连接/无连接 --- line:102 
var Play021234 = true; // true/false 有1234/无1234 --- line:036 
// for Test alert(filterStyle.length);
//document.getElementById("test").innerHTML = "4.6";

var filterStyle = getFilterStyle();
var baseSpacStyle = "clear:both; display:block; width:23px;line-height:18px; font-size:12px; FONT-FAMILY:'宋体'; opacity: 0.6;";
    baseSpacStyle += "border:1px solid #fff;border-right:0;border-bottom:0;color:#fff;text-align:center; cursor:pointer; ";
var liStyle = "margin:0;list-style-type: none; margin:0;padding:0; float:right;";

var jsPlay02 = {
	
 _timer : null,
 _items : [],
 _container : null,
 _index : 0,
 _imgs : [],
 intervalTime : Play02Speed, 
 
 init : function( objID, w, h, time ){
  this.intervalTime = time || this.intervalTime;
  this._container = document.getElementById( objID );
  this._container.style.display = "block";
  this._container.style.width = w + "px";
  this._container.style.height = h + "px";
  this._container.style.position = "relative";
  this._container.style.overflow  = "hidden";
  this._container.style.border = "0px solid #F0F0F0";
  var linkStyle = "display: block; TEXT-DECORATION: none;";
  if( document.all ){
   linkStyle += filterStyle;
  } 
  var ulStyle = "margin:0;width:"+w+"px;position:absolute;z-index:999;right:5px;FILTER:Alpha(Opacity=30,FinishOpacity=90, Style=1);overflow: hidden;bottom:-1px;height:16px; border-right:1px solid #fff;";
  if(!Play021234) { ulStyle = "display:none;"+ulStyle; }
  var ulHTML = "";
  for(var i = this._items.length -1; i >= 0; i--){
   var spanStyle = "";
   if( i==this._index ){ spanStyle = baseSpacStyle + "background:#ff0000;"; } 
   else{                 spanStyle = baseSpacStyle + "background:#000;"; }
   ulHTML += getUlHtml(this,i,liStyle,spanStyle);
  }
  var urlPeace = getALink(this,linkStyle); 
  this._container.innerHTML = urlPeace + "<ul style=\""+ulStyle+"\">"+ulHTML+"</ul>";
  var link = this._container.getElementsByTagName("A")[0]; 
  link.style.width =  w + "px";
  link.style.height = h + "px";
  link.style.background = 'url(' + this._items[0].img + ') no-repeat center center';
  this._timer = setInterval( "jsPlay02.play()", this.intervalTime );
 },
 
 addItem : function( _title, _link, _imgURL ){
  this._items.push ( {title:_title, link:_link, img:_imgURL } );
  var img = new Image();
  img.src = _imgURL;
  this._imgs.push( img );
 },
 
 play : function( index ){
  if( index!=null ){
   this._index = index;
   clearInterval( this._timer );
   this._timer = setInterval( "jsPlay02.play()", this.intervalTime );
  } else {
   this._index = this._index<this._items.length-1 ? this._index+1 : 0;
  }
  var link = this._container.getElementsByTagName("A")[0]; 
  if(link.filters){
   var ren = Math.floor(Math.random()*(link.filters.length));
   link.filters[ren].Apply();
   link.filters[ren].play();
  }
  link.href = this._items[this._index].link;
  link.title = this._items[this._index].title;
  link.style.background = 'url(' + this._items[this._index].img + ') no-repeat center center';
  var ulHTML = "";
  for(var i = this._items.length -1; i >= 0; i--){
   var spanStyle = "";
   if( i==this._index ){ spanStyle = baseSpacStyle + "background:#ff0000;"; } 
   else{                 spanStyle = baseSpacStyle + "background:#000;"; }
   ulHTML += getUlHtml(this,i,liStyle,spanStyle);
  }
  this._container.getElementsByTagName("UL")[0].innerHTML = ulHTML; 
 },
 
 mouseOver : function(obj){
  var i = parseInt( obj.innerHTML );
  if( this._index!=i-1){
   obj.style.color = "#ff0000";
  }
 },
 
 mouseOut : function(obj){
  obj.style.color = "#fff";
 }

}

function getALink(e,linkStyle){
 var lnk = e._items[e._index].link; 
 if(Play02Link){ lnk = "<a href='../../../inc/home/"+lnk+"' target='_blank' title=\""+e._items[e._index].title+"\" style=\""+linkStyle+"\"></a>";}
 else {          lnk = "<a href='#'       target='_self'  title=\""+e._items[e._index].title+"\" style=\""+linkStyle+"\"></a>";}
 return lnk;
}

function getUlHtml(e,i,liStyle,spanStyle){
 var html = "";
 html += "<li style=\""+liStyle+"\">";
 html += "<span onmouseover=\"jsPlay02.mouseOver(this);\" onmouseout=\"jsPlay02.mouseOut(this);\" style=\""+spanStyle+"\" onclick=\"jsPlay02.play("+i+");return false;\" herf=\"javascript:;\" title=\"" + e._items[i].title + "\">" + (i+1) + "</span>";
 html += "</li>";
 return html;
}

function getFilterStyle(){
 var sty = "FILTER:";
 sty += "progid:DXImageTransform.Microsoft.Barn(duration=0.5, motion='out', orientation='vertical') ";
 sty += "progid:DXImageTransform.Microsoft.Barn ( duration=0.5,motion='out',orientation='horizontal') ";
 sty += "progid:DXImageTransform.Microsoft.Blinds ( duration=0.5,bands=10,Direction='down' )";
 sty += "progid:DXImageTransform.Microsoft.CheckerBoard()";
 sty += "progid:DXImageTransform.Microsoft.Fade(duration=0.5,overlap=0)";
 sty += "progid:DXImageTransform.Microsoft.GradientWipe ( duration=1,gradientSize=1.0,motion='reverse' )";
 sty += "progid:DXImageTransform.Microsoft.Inset ()";
 sty += "progid:DXImageTransform.Microsoft.Iris ( duration=1,irisStyle=PLUS,motion=out )";
 sty += "progid:DXImageTransform.Microsoft.Iris ( duration=1,irisStyle=PLUS,motion=in )";
 sty += "progid:DXImageTransform.Microsoft.Iris ( duration=1,irisStyle=DIAMOND,motion=in )";
 sty += "progid:DXImageTransform.Microsoft.Iris ( duration=1,irisStyle=SQUARE,motion=in )";
 sty += "progid:DXImageTransform.Microsoft.Iris ( duration=0.5,irisStyle=STAR,motion=in )";
 sty += "progid:DXImageTransform.Microsoft.RadialWipe ( duration=0.5,wipeStyle=CLOCK )";
 sty += "progid:DXImageTransform.Microsoft.RadialWipe ( duration=0.5,wipeStyle=WEDGE )";
 sty += "progid:DXImageTransform.Microsoft.RandomBars ( duration=0.5,orientation=horizontal )";
 sty += "progid:DXImageTransform.Microsoft.RandomBars ( duration=0.5,orientation=vertical )";
 sty += "progid:DXImageTransform.Microsoft.RandomDissolve ()";
 sty += "progid:DXImageTransform.Microsoft.Spiral ( duration=0.5,gridSizeX=16,gridSizeY=16 )";
 sty += "progid:DXImageTransform.Microsoft.Stretch ( duration=0.5,stretchStyle=PUSH )";
 sty += "progid:DXImageTransform.Microsoft.Strips ( duration=0.5,motion=rightdown )";
 sty += "progid:DXImageTransform.Microsoft.Wheel ( duration=0.5,spokes=8 )";
 sty += "progid:DXImageTransform.Microsoft.Zigzag ( duration=0.5,gridSizeX=4,gridSizeY=40 ); width: 100%; height: 100%";
 return sty;
}

/**************************************************
名称: 图片轮播类
创建时间: 2007-11-12
示例:
 页面中已经存在名为advPlay02A(或者别的ID也行)的节点.
 jsPlay02.addItem( "Play02Adv01", "http://www.dgchr.com/", "http://www.dgchr.com/img/img20/logo.jpg");
 jsPlay02.addItem( "Play02Adv02", "http://www.dg.gd.cn/", "http://www.dg.gd.cn/sys/pinc/pimg/dgnetlogo.gif");
 jsPlay02.addItem( "Play02Adv03", "http://www.txjia.com/", "http://www.txjia.com/img/logo/logo088031.jpg");
 setTimeout("jsPlay02.init('advPlay02A',360,120)",87);
备注:
 适用于一个页面只有一个图片轮播的地方.
***************************************************/

