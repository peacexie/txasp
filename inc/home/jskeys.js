
function getKeys(str){
	
  var reSpc1  = /，|、|。|．|；|：|？|！|︰|…|‥|′|‵|々|～|‖|ˇ|ˉ|﹐|﹑|﹒|·|﹔|﹕|﹖|﹗|｜|-|︱|-|︳|︴|﹏|（|）|︵|︶|｛|｝|︷|︸|〔|〕|︹|︺|【|】|︻|︼|《|》|︽|︾|〈|〉|︿|﹀|「|」|﹁|﹂|『|』|﹃|﹄|﹙|﹚|﹛|﹜|﹝|﹞|〝|〞|ˋ|ˊ/gi;
  var reSpc2  = /!|\"|\#|\$|\%|\&|\'|\(|\)|\*|\+|\,|\-|\.|\/|\:|\;|\<|\=|\>|\?|\@|\[|\\|\]|\^|\_|\`|\{|\||\}|\~/gi; 
  var reSpcA  = /\n\r|\r\n|\n|\r|\t|\f|\b/gi; // !"#$%&'()*+,-./:;<=>?@[\]^_`{|}~

  var reWord1 = / after | all | also | an | and | another | any | are | as | at | be | because | been | before | being | between | both | but | by | came | can | come | could | did | do | each | for | from | get | got | had | has | have | he | her | here | him | himself | his | how | if | in | into | is | it | like | make | many | me | might | more | most | much | must | my | never | now | of | on | only | or | other | our | out | over | said | same | see | should | since | some | still | such | take | than | that | the | their | them | then | there | these | they | this | those | through | to | too | under | up | very | was | way | we | well | were | what | where | which | while | who | with | would | you | your /gi;
  var reWord2 = /的|一|不|在|人|有|是|为|以|于|上|他|而|后|之|来|及|了|因|下|可|到|由|这|与|也|此|但|并|个|其|已|无|小|我|们|起|最|再|今|去|好|只|又|或|很|亦|某|把|那|你|乃|它/gi;

  // 专业词汇！
  var reWordA = /www |\ com|\ net|\ org|\ cn|\ gov/gi; 
  var reWordB = / gi | js | icp /gi;
  //var reWordD=/table|tbody|tr|td|div|span/gi;
  // 自定义！
  var reWordE = / dg | gd | 0769 /gi; 
  var reWordF = /自定义|mySelf|首页|主页/gi;
	
  str = str.replace(/<[^>]*>/gi," "); //过滤HTML代码,取前后各1000个字
  if(str.length>980){ str = str.substring(0,480)+str.substring(str.length-480); }
  
  str = str.replace(reSpc1," "); //过滤 特殊符号 
  str = str.replace(reSpc2," "); 
  str = str.replace(reSpcA," ");
  
  str = str.replace(reWord1," "); //过滤 噪音字符 from MSSQL2005
  str = str.replace(reWord2," "); 
  
  str = str.replace(reWordA," "); //过滤 专业词汇！
  str = str.replace(reWordB," ");
  str = str.replace(reWordE," "); //过滤 自定义字符
  str = str.replace(reWordF," ");
  
  for(i=0;i<8;i++){ // 多个空格合并
    str = str.replace(/\ \ \ \ \ \ \ \ /g," ");
	str = str.replace(/\ \ \ \ /g," ");
	str = str.replace(/\ \ /g," ");
  }
  
  str = str.replace(/\ /g,","); 
  a = str.split(","); str = "";
  for(i=0;i<a.length;i++){
  if(a[i].length>1){
  if(str.indexOf(a[i])<0) {
    if(a[i].length==2){
	  if(isNaN(a[i])){ 
	    str += a[i]+", "; // 过滤 00~99
	  }	 
    }else{ 
	    str += a[i]+", "; 
	}
  }}}
	
  return str;
  
}

function getKeyN(str,N){

  str = getKeys(str);
  if(str.indexOf(",")<0) {
	return str;
  }
  a = str.split(","); // 过滤重复:如智通,智通东莞 过滤后者
  t = 0;
  for(i=0;i<a.length;i++){
  for(j=0;j<a.length;j++){
  if(i!=j){
	 if((a[i].length>1)&&(a[j].length>1)){
	 if(a[j].indexOf(a[i])>=0){ //((a[i].indexOf(a[j])>=0)||) {
		a[j] = ""; //break;	
		t++;
	 }}
  }}}
  
  str = ""; // 新str
  for(i=0;i<a.length;i++){
  if(a[i].length>1){
  if(str.indexOf(a[i])<0) {
	str += a[i]+","; 
  }}}
  
  a = str.split(","); // 判断是否>2N
  s1=""; s2="";
  if(a.length>2*N+2){
    for(i=0;i<N;i++){
	  s1 += a[i]+",";
	  s2 += a[a.length-N-1+i]+",";
    }
	str = s1+" ... "+s2;  
  }
  str = str.substring(0,str.length-1);
  return str; //+"("+t+")";
  
}