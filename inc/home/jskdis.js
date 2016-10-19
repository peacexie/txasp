// JavaScript Document

document.write("<noscript><iframe src='*.PeacePage'></iframe></noscript>");
//禁止保存：<noscript><iframe src="*.htm"></iframe></noscript>，放在head里面。
//禁止粘贴：<input type=text onpaste="return false">
//关闭输入法：<input style="ime-mode:disabled">

document.oncontextmenu = function(){event.returnValue=false;} //屏蔽鼠标右键
document.onselectstart = function(){event.returnValue=false;} //屏蔽鼠标选择
window.onhelp = function() {return false;} //屏蔽F1帮助
document.onmousewheel = function(){ //屏蔽Shift+滚轮,Ctrl+滚轮
  if(event.shiftKey || event.ctrlKey){
    event.keyCode=0;
    event.returnValue=false;   
  }
}

//ondragstart="return false" 
//onselectstart="return false" 
//oncopy="document.selection.empty()" 
//onbeforecopy="return false"
document.ondragstart = function(){event.returnValue=false;}
document.onselectstart = function(){event.returnValue=false;}
document.onbeforecopy = function(){document.selection.empty();}
document.oncopy = function(){event.returnValue=false;}

document.onmousedown = function(){ 
  //alert(event.button); // for Test
  if(event.button==4){alert('非法操作;');} //屏蔽鼠标中键
} 

document.onkeydown = function(){
  //alert(event.keyCode); // for Test
  if(event.keyCode==27){event.keyCode=0;event.returnValue=false;}  //屏蔽ESC
  if(event.keyCode==114){event.keyCode=0;event.returnValue=false;}  //屏蔽F3
  if(event.keyCode==116){event.keyCode=0;event.returnValue=false;}  //屏蔽F5
  if(event.keyCode==122){event.keyCode=0;event.returnValue=false;}  //屏蔽F11
  if(event.ctrlKey && event.keyCode==67){event.keyCode=0;event.returnValue=false;} //屏蔽 Ctrl+c
  if(event.ctrlKey && event.keyCode==86){event.keyCode=0;event.returnValue=false;} //屏蔽 Ctrl+v
  if(event.ctrlKey && event.keyCode==70){event.keyCode=0;event.returnValue=false;} //屏蔽 Ctrl+f
  if(event.ctrlKey && event.keyCode==87){event.keyCode=0;event.returnValue=false;} //屏蔽 Ctrl+w
  if(event.ctrlKey && event.keyCode==69){event.keyCode=0;event.returnValue=false;} //屏蔽 Ctrl+e
  if(event.ctrlKey && event.keyCode==72){event.keyCode=0;event.returnValue=false;} //屏蔽 Ctrl+h
  if(event.ctrlKey && event.keyCode==73){event.keyCode=0;event.returnValue=false;} //屏蔽 Ctrl+i
  if(event.ctrlKey && event.keyCode==79){event.keyCode=0;event.returnValue=false;} //屏蔽 Ctrl+o
  if(event.ctrlKey && event.keyCode==76){event.keyCode=0;event.returnValue=false;} //屏蔽 Ctrl+l
  if(event.ctrlKey && event.keyCode==80){event.keyCode=0;event.returnValue=false;} //屏蔽 Ctrl+p
  if(event.ctrlKey && event.keyCode==66){event.keyCode=0;event.returnValue=false;} //屏蔽 Ctrl+b
  if(event.ctrlKey && event.keyCode==78){event.keyCode=0;event.returnValue=false;}  //屏蔽 Ctrl+n
  if(event.shiftKey && event.keyCode==121){event.keyCode=0;event.returnValue=false;}  //屏蔽 shift+F10
  
  if(window.event.srcElement.tagName == "A" && window.event.shiftKey) 
    {event.keyCode=0;event.returnValue=false;} //屏蔽 shift 加鼠标左键新开一网页
	
  with(event){
    if((event.ctrlKey && event.keyCode==82)){
      //alert(event.keyCode);
	  event.keyCode = 0;
      event.cancelBubble = true; 
	  return false;
    }
  }
	
  //屏蔽 Alt+ 方向键 ← ,  Alt+ 方向键 →
  if((window.event.altKey)&&((window.event.keyCode==37)||(window.event.keyCode==39))){
     alert('不准你使用[Alt+方向键]前进或后退网页!');
     event.returnValue=false;
  }
  //注：这还不是真正地屏蔽 Alt+ 方向键，因为 Alt+ 方向键弹出警告框时，按住 Alt 键不放，
  //用鼠标点掉警告框，这种屏蔽方法就失效了。以后若有哪位高手有真正屏蔽 Alt 键的方法，请告知。
 
  //屏蔽退格删除键,F5 刷新键,Ctrl+R
  if((event.keyCode==8)||(event.keyCode==116)||(event.ctrlKey && event.keyCode==82)){ 
     event.keyCode=0;
     event.returnValue=false; 
  }
  if((window.event.altKey)&&(window.event.keyCode==115)){ //屏蔽Alt+F4
      window.showModelessDialog("about:blank","","dialogWidth:1px;dialogheight:1px");
      return false;
  }
	
}

