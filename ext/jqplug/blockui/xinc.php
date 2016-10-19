<p>Jquery中BlockUI的详解(转)</p>
<p><strong>BlockUI插件需要那个jQuery版本的支持？</strong><br />
  BlockUI兼容jQuery   v1.2.3以上的版本</p>
<p><strong>BlockUI插件的V2版本有那些变化？</strong><br />
  解除锁定的时候，用于提示信息的元素不会从DOM中移除<br />
  默认的遮罩层为黑色<br />
  可用的选项设置进行了统一和清理<br />
  设置插件选项的方法改变了<br />
  放弃了对Opera   8的支持<br />
  提高了源代码的可读性<br />
  移除了displayBox功能 (其他 plugins会做的更好)<br />
</p>
<p><strong>我的原代码中的blockUI插件与新的2.00版兼容么？</strong><br />
  不兼容，如果原代码改变了blockUI的默认属性，那么会出现兼容问题。如何设置选项的语法发生了细微的改变。请查看Options页来了解新版本的选项设置方法。<br />
</p>
<p><strong>BlockUI插件还依赖于其他的插件么？</strong><br />
  不依赖<br />
</p>
<p><strong>我可以改变页面锁定时默认的提示信息么？</strong><br />
  <br />
  可以。默认的提示信息储存在$.blockUI.defaults.message中。你可以以一个新的值来替换它，例如：<br />
  $.blockUI.defaults.message   = &quot;Please be patient...&quot;;</p>
<p><strong>我能够改变遮罩层的颜色和透明度么？</strong></p>
<p>可以。默认的遮罩层样式储存在 $.blockUI.defaults.overlayCSS中。你可以指定一个不同的颜色和透明度，</p>
<p>例如<br />
  // 使用黄色遮罩层<br />
  $.blockUI.defaults.overlayCSS.backgroundColor =   '#ff0';<br />
  // 使遮罩层更透明<br />
  $.blockUI.defaults.overlayCSS.opacity = '.2';</p>
<p><strong>BlockUI支持Opera   8么？<br />
  </strong>不支持<br />
  <br />
  <strong>在linux的FF上我为什么看不到遮罩层？</strong></p>
<p>有几个人告诉我，在FF/Linux上整个页面的透明度渲染慢的让人发疯，所以默认情况下，在这些平台上遮罩层不透明。你可以重设applyPlatformOpacityRules值来启用透明度。例如:</p>
<p>// 在FF/Linux下启用遮罩层透明$.blockUI.defaults.applyPlatformOpacityRules =   false;<br />
</p>
<p><strong>BlockUI基本使用</strong><br />
  //   当有ajax请求时，当加载信息较慢时，会显示该等待信息，带来良好的用户体验<br />
  $(document).ajaxStart(function ()   {<br />
  $.blockUI({</p>
<p>                  // $.blockUI.defaults.message =   &quot;请稍候&quot;;(不写在$.blockUI({})里，写在外面)<br />
  message: '&lt;span   style=&quot;font-size:13px;font-weight:bolder&quot;&gt;请稍候&lt;/span&gt;',          </p>
<p>                  //   指的是提示框的css<br />
  css:   {<br />
  width: &quot;45px&quot;,   //   宽度小一点<br />
  top:   &quot;50%&quot;,<br />
  left:   &quot;50%&quot;<br />
  },           </p>
<p>                  //   遮光罩的css<br />
  // 等价$.blockUI.defaults.overlayCSS.backgroundColor =   &quot;#E4E7EC&quot;;<br />
  overlayCSS:   {<br />
  backgroundColor:   &quot;yellow&quot;,<br />
  opacity:&quot;0.8&quot;<br />
  }<br />
  });<br />
  });</p>
<p><strong>下载</strong><br />
  新版本的blockUI   v2.00可以在这里得到: jquery.blockUI.js.</p>
<p>旧版本的blockUI仍然可以在这里得到： <a href="http://jqueryjs.googlecode.com/svn/trunk/plugins/blockUI/jquery.blockUI.js">http://jqueryjs.googlecode.com/svn/trunk/plugins/blockUI/jquery.blockUI.js</a>.   旧版本的文档在这里.</p>
<p>原帖地址：<a href="http://www.cssrain.cn/demo/blockUI-V2/jQuery/blockUI/jQueryBlockUI.html">http://www.cssrain.cn/demo/blockUI-V2/jQuery/blockUI/jQueryBlockUI.</a><a href="http://www.cssrain.cn/demo/blockUI-V2/jQuery/blockUI/jQueryBlockUI.html">html</a></p>
