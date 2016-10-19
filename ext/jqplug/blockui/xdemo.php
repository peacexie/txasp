<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Demo - Jquery中BlockUI的详解(转)</title>
<script src="/inc/home/jquery.js" type="text/javascript"></script>
<SCRIPT type=text/javascript src="blockui.js"></SCRIPT>
<link href="blockui.css" rel="stylesheet" type="text/css"/>
</head>
<body>
BlockUI基本使用 <br />
<a href="http://blog.sina.com.cn/s/blog_67aaf4440100rnjd.html" target="_blank">http://blog.sina.com.cn/s/blog_67aaf4440100rnjd.html</a>
<hr>
<div id='appRead' style="width:120px; border:1px solid #06C;CURSOR: pointer; padding:2px;">查看注册协议</div>
<script type=text/javascript>
$(document).ready(function(){
	$("#appRead").click(function(){
		$.ajax({
				type: "POST",
				url: "xinc.php",
				data: "act=appRead&membertypeid=1",
				success: function(msg){
					$('#frmWindow').remove(); 
					$("body").append("<div id='frmWindow'></div>");
					$('#frmWindow').append('<div class="topBar">会员注册协议<div class="pwClose">[X]关闭</div></div><div class="border"><div class="windowcontent"><div class="ntc">'+msg+'</div></div></div>');
					$.blockUI({
						message:$('#frmWindow'),
						css:{width:'600px',top:'80px'},
						overlayCSS:{backgroundColor:"gray",opacity:"0.5"}
					}); 
					$('.pwClose').click(function() { 
						$.unblockUI(); 
					}); 
				} //cuccess
		 }); //ajax
	});
});
</script>
<hr>
<?php require('xinc.php'); ?>
<hr />
<p>&nbsp;</p>
</body>
</html>
