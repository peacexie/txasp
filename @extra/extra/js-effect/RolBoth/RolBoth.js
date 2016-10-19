//var RollFlag = new Array(4);
function clickButtonShow(idx)
{
   //menu=eval("document.all.block"+idx+".style");
   //if(RollFlag[idx])
   {
     //menu.display = "none";
     //RollFlag[idx] = 0;
   //}else{
     //menu.display = "block";
     //RollFlag[idx] = 1;
   }
}

function RollColor(ob){   
   ob.filters.gray.enabled=false; 
  } 
function RollCFix(ob){ 
   ob.filters.gray.enabled=false;  
  } 
function RollCFade(ob){ 
   ob.filters.gray.enabled=true;   
  } 

function Roll_Change(source){    
    document.getElementById("ImgShow090519").src=source;
	setImgSiz2(document.getElementById("ImgShow090519"),200,100);
}

function RollMain(obj, oWidth, oHeight, direction, drag, zoom, speed) 
{
  var scrollDiv    = obj
  var scrollContent  = document.createElement("span")
  scrollContent.style.position = "absolute"
  scrollDiv.applyElement(scrollContent, "inside")
  var displayWidth  = (oWidth != "auto" && oWidth ) ? oWidth : scrollContent.offsetWidth + parseInt(scrollDiv.currentStyle.borderRightWidth)
  var displayHeight  = (oHeight != "auto" && oHeight) ? oHeight : scrollContent.offsetHeight + parseInt(scrollDiv.currentStyle.borderBottomWidth)
  var contentWidth  = scrollContent.offsetWidth
  var contentHeight  = scrollContent.offsetHeight
  var scrollXItems  = Math.ceil(displayWidth / contentWidth) + 1
  var scrollYItems  = Math.ceil(displayHeight / contentHeight) + 1
  scrollDiv.style.width   = displayWidth
  scrollDiv.style.height   = displayHeight
  scrollDiv.style.overflow = "hidden"
  scrollDiv.setAttribute("state", "stop")
  scrollDiv.setAttribute("drag", drag ? drag : "horizontal")
  scrollDiv.setAttribute("direction", direction ? direction : "stop")
  scrollDiv.setAttribute("zoom", zoom ? zoom : 1)
  scrollContent.style.zoom = scrollDiv.zoom
  var scroll_script =  "var scrollDiv = " + scrollDiv.uniqueID                    +"\n"+
        "var scrollObj = " + scrollContent.uniqueID                  +"\n"+
        "var contentWidth = " + contentWidth + " * (scrollObj.runtimeStyle.zoom ? scrollObj.runtimeStyle.zoom : 1)"  +"\n"+
        "var contentHeight = " + contentHeight + " * (scrollObj.runtimeStyle.zoom ? scrollObj.runtimeStyle.zoom : 1)"  +"\n"+
        "var scrollx = scrollObj.runtimeStyle.pixelLeft"                +"\n"+
        "var scrolly = scrollObj.runtimeStyle.pixelTop"                  +"\n"+
        "switch (scrollDiv.state.toLowerCase())"                +"\n"+
        "{"                          +"\n"+
          "case ('scroll')  :"                  +"\n"+
            "switch (scrollDiv.direction)"                +"\n"+
            "{"                      +"\n"+
              "case ('left')    :"              +"\n"+
                "scrollx = (--scrollx) % contentWidth"          +"\n"+
                "break"                  +"\n"+
              "case ('right')  :"                +"\n"+
                "scrollx = -contentWidth + (++scrollx) % contentWidth"      +"\n"+
                "break"                  +"\n"+
              "case ('up')  :"                +"\n"+
                "scrolly = (--scrolly) % contentHeight"          +"\n"+
                "break"                  +"\n"+
              "case ('down')  :"                +"\n"+
                "scrolly = -contentHeight + (++scrolly) % contentHeight"    +"\n"+
                "break"                  +"\n"+
              "case ('left_up')  :"              +"\n"+
                "scrollx = (--scrollx) % contentWidth"          +"\n"+
                "scrolly = (--scrolly) % contentHeight"          +"\n"+
                "break"                  +"\n"+
              "case ('left_down')  :"              +"\n"+
                "scrollx = (--scrollx) % contentWidth"          +"\n"+
                "scrolly = -contentHeight + (++scrolly) % contentHeight"    +"\n"+
                "break"                  +"\n"+
              "case ('right_up')  :"              +"\n"+
                "scrollx = -contentWidth + (++scrollx) % contentWidth"      +"\n"+
                "scrolly = (--scrolly) % contentHeight"          +"\n"+

                "break"                  +"\n"+
              "case ('right_down')  :"              +"\n"+
                "scrollx = -contentWidth + (++scrollx) % contentWidth"      +"\n"+
                "scrolly = -contentHeight + (++scrolly) % contentHeight"    +"\n"+
                "break"                  +"\n"+
              "default    :"              +"\n"+
                "return"                +"\n"+
            "}"                      +"\n"+
            "scrollObj.runtimeStyle.left = scrollx"              +"\n"+
            "scrollObj.runtimeStyle.top = scrolly"              +"\n"+
            "break"                      +"\n"+ 
          "case ('stop')  :"                    +"\n"+
          "case ('drag')  :"                    +"\n"+
          "default  :"                    +"\n"+
            "return"                    +"\n"+
        "}" 
  var contentNode = document.createElement("span")
  contentNode.runtimeStyle.position = "absolute"
  contentNode.runtimeStyle.width = contentWidth
  scrollContent.applyElement(contentNode, "inside")
  for (var i=0; i <= scrollXItems; i++)
  {
    for (var j=0; j <= scrollYItems ; j++)
    {
      if (i+j == 0)  continue
      var tempNode = contentNode.cloneNode(true)
      var contentLeft, contentTop
      scrollContent.insertBefore(tempNode)
      contentLeft = contentWidth * i
      contentTop = contentHeight * j
      tempNode.runtimeStyle.left = contentLeft
      tempNode.runtimeStyle.top = contentTop
    }
  }
  scrollDiv.onpropertychange = function()
  {
    var propertyName = window.event.propertyName
    var propertyValue = eval("this." + propertyName)  
    switch(propertyName)
    {
      case "zoom"    :
        var scrollObj = this.children[0]
        scrollObj.runtimeStyle.zoom = propertyValue
        var content_width = scrollObj.children[0].offsetWidth * propertyValue
        var content_height = scrollObj.children[0].offsetHeight * propertyValue
        scrollObj.runtimeStyle.left = -content_width + (scrollObj.runtimeStyle.pixelLeft % content_width)
        scrollObj.runtimeStyle.top = -content_height + (scrollObj.runtimeStyle.pixelTop % content_height)
        break
    }
  }
  scrollDiv.onlosecapture = function()
  {
    this.state = this.tempState ? this.tempState : this.state
    this.runtimeStyle.cursor = ""
    this.removeAttribute("tempState")
    this.removeAttribute("start_x")
    this.removeAttribute("start_y")
    this.removeAttribute("default_left")
    this.removeAttribute("default_top")
  } 
  scrollDiv.onmousedown = function()
  {
    if (this.state != "drag")  this.setAttribute("tempState", this.state)
    this.state = "drag"
    this.runtimeStyle.cursor = "default"
    this.setAttribute("start_x", event.clientX)
    this.setAttribute("start_y", event.clientY)
    this.setAttribute("default_left", this.children[0].style.pixelLeft)
    this.setAttribute("default_top", this.children[0].style.pixelTop)
    this.setCapture()
  }
scrollDiv.onmousemove = function()
  {
    if (this.state != "drag")  return
    var scrollx = scrolly = 0
    var zoomValue = this.children[0].style.zoom ? this.children[0].style.zoom : 1
    var content_width = this.children[0].children[0].offsetWidth * zoomValue
    var content_Height = this.children[0].children[0].offsetHeight * zoomValue
    if (this.drag == "horizontal" || this.drag == "both")
    {
      scrollx = this.default_left + event.clientX - this.start_x
      scrollx = -content_width + scrollx % content_width
      this.children[0].runtimeStyle.left = scrollx
    }
    if (this.drag == "vertical" || this.drag == "both")
    {
      scrolly = this.default_top + event.clientY - this.start_y
      scrolly = -content_Height + scrolly % content_Height
      this.children[0].runtimeStyle.top = scrolly
    }
  }
  scrollDiv.onmouseup = function()
  {
    this.releaseCapture()
  }
  
  scrollDiv.state = "scroll";
  setInterval(scroll_script, speed ? speed : 10)
}