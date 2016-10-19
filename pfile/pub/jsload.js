
var flagScript = 'script';
var strCssLink = '<link href="(url)" rel="stylesheet" type="text/css">';
var strScript = '<'+flagScript+' src="(url)" type="text/javascript"></'+flagScript+'>';
document.write(strCssLink.replace('(url)','../img/rnd_nid/box_nid.css'));
document.write(strScript.replace('(url)','../inc/home/jsInfo.js'));
document.write(strScript.replace('(url)','../inc/home/jsadv.js'));
//document.write(strScript.replace('(url)','../inc/home/convtab.js'));
//document.write(strScript.replace('(url)','../pfile/pub/jshome.js'));
//document.write(strScript.replace('(url)','../inc/home/jsPager.js'));

//LoadLink("../pfile/pimg/style.css");
//LoadLink("../pfile/spub.css");
