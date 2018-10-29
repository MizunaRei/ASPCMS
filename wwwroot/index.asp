<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<!--#include file="inc/syscode_article.asp"-->
<%

const ChannelID=1
Const ShowRunTime="Yes"
MaxPerPage=20
SkinID=0
PageTitle="首页"
Set rsArticle= Server.CreateObject("ADODB.Recordset")
Set rsPic= Server.CreateObject("ADODB.Recordset")
%>
<html>
<head>
<TITLE><%=strPageTitle%></TITLE>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK 
href="imags97/DefaultSkin.css" type=text/css rel=stylesheet>
<LINK 
href="SkinIndex/DefaultSkin.css" type=text/css rel=stylesheet>
<SCRIPT language="JavaScript" src="SkinIndex/menu.js" 
type="text/JavaScript"></SCRIPT>
<META content=o7FhrjMKBn/3XGgcDXmGdE4BkAxwd6a97bpMEXpOURY= name=verify-v1>
<META content="MSHTML 6.00.2900.3395" name=GENERATOR>
<!--标题菜单 样式表 代码 begin-->
<link href="/templets/style/dedecms.css" rel="stylesheet" media="screen" type="text/css" />
<script language="javascript" type="text/javascript" src="/include/dedeajax2.js"></script>
<script src="/images/js/j.js" language="javascript" type="text/javascript"></script>
<style type="text/css">
.Layer1 {
	position:absolute;
	left:-4218px;
	top:138px;
	width:148px;
	height:75px;
	z-index:2;
}
</style>
<!--
	$(function(){
		$("dl.tbox dt span.label a[_for]").mouseover(function(){
			$(this).parents("span.label").children("a[_for]").removeClass("thisclass").parents("dl.tbox").children("dd").hide();
			$(this).addClass("thisclass").blur();
			$("#"+$(this).attr("_for")).show();
		});
		$("a[_for=uc_member]").mouseover();
	});
	
	function CheckLogin(){
	  var taget_obj = document.getElementById('_userlogin');
	  myajax = new DedeAjax(taget_obj,false,false,'','','');
	  myajax.SendGet2("/member/ajax_loginsta.php");
	  DedeXHTTP = null;
	}
-->
</script>
<!--标题菜单 样式表 代码 end-->
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<!--#include file="inc/Skin_CSS.asp"-->
<%call MenuJS()%>
</head>
<body bgcolor="#FFFFFF"  style="BACKGROUND-COLOR: #ffffff"  leftmargin="0" topmargin="0" marginwidth="0" marginheight="0"       onload=changeimage()  >
<!-- ImageReady Slices (首页_slice2.psd) -->
<div align="center" >
  <table id="__01" width="959"  border="0" cellpadding="0" cellspacing="0" >
    <tr>
      <td colspan="20" align="center"><img   src="Image20100223/TopPic20100227.jpg"  alt="红桥网" width="959" ></td>
    </tr>
  </table>
  <!--top-->
</div>
<!--标题菜单 开始-->
<div align="center" >
  <!--<table>
    <tr>
      <td align="center" valign="middle">-->
      <div    class="module blue mT10 wrapper w963" style="margin:0 auto;">
          <div class="top">
            <div class="t_l"></div>
            <div class="t_r"></div>
            <!-- //如果不使用currentstyle，可以在channel标签加入 cacheid='channeltoplist' 属性提升性能 -->
            <div id="navMenu"  >
              <table ><tr><td     valign="bottom">
              <ul >
                <!--资料中心  ‖  理论动态  ‖  时事新闻  ‖  学生作品  ‖  文    集  ‖  留    言-->
                
                <li  ><a href='/'>主页</a></li>
                <!--<li  ><a href='/'>近现代史</a></li>
                <li><a href='/'>思修法基</a></li> 
                <li> <a href='/'>马义原理</a></li>  
                <li>   <a href='/'>毛思邓三</a></li>-->  
                
                
                <!-- <li><a href='Article_Class2.asp?ClassID=2' >资料中心</a></li>
        <li><a href='/Vdt/'  rel='dropmenu4'>理论动态</a></li>
        <li><a href='/Vxm/'  rel='dropmenu5'>时事新闻</a></li>
        <li><a href='/Vpphd/'  rel='dropmenu25'>学生作品</a></li>
        <li><a href='/Vllyz/' >文  集</a></li>
        <li><a href='/plus/list.php?tid=9' >互动平台</a></li>-->
                <!--  <li><a href='/plus/list.php?tid=6' >风采</a></li>
        <li><a href='/Vpx/' >资料下载</a></li>
        <li><a href='/Vblog/'  rel='dropmenu8'>博客</a></li>
        <li><a href='/bbs/' >论坛</a></li>
        <li><a href='/ask/' >问卷</a></li>-->
                <!--为了动态增删课程，必须查询数据库-->
                
                 <%
	  dim rsArticleSpecialList , sqlArticleSpecialList
	  sqlArticleSpecialList="select  SpecialID , SpecialAbbreviation , OrderID from  Special order by  OrderID asc"
	  set rsArticleSpecialList=Server.CreateObject("Adodb.RecordSet")
	  rsArticleSpecialList.Open sqlArticleSpecialList,conn,1,3
	  do while not rsArticleSpecialList.EOF
	  	Response.Write("<li> <a   href='Article_Special.asp?SpecialID=" & rsArticleSpecialList("SpecialID")  & "' >"  &  rsArticleSpecialList("SpecialAbbreviation")  &   "</a></li>")
	  
	  rsArticleSpecialList.MoveNext
	  Loop
	  
	  rsArticleSpecialList.Close
set	  rsArticleSpecialList=Nothing
	  
	  
	  
	  
	  
	  
	  %>
                
                <!--为了动态增删栏目，必须查询数据库-->
                
                
                <%
	  dim rsArticleClassList , sqlArticleClassList
	  sqlArticleClassList="select  ClassID , ClassName , RootID from ArticleClass order by RootID asc"
	  set rsArticleClassList=Server.CreateObject("Adodb.RecordSet")
	  rsArticleClassList.Open sqlArticleClassList,conn,1,3
	  do while not rsArticleClassList.EOF
	  	Response.Write("<li> <a   href='Article_Class2.asp?ClassID=" & rsArticleClassList("ClassID")  & "' >"  &  rsArticleClassList("ClassName")  &   "</a></li>")
	  
	  rsArticleClassList.MoveNext
	  Loop
	  
	  rsArticleClassList.Close
set	  rsArticleClassList=Nothing
	  
	  
	  
	  
	  
	  
	  %>
              </ul>
              </td></tr></table>
            </div>
          </div>
          </div>
        <!--</td>
    </tr>
  </table>-->
</div>
<!--标题菜单 结束-->
<div align="center">
  <!--main body-->
  <table  bgcolor="#FFFFFF" width="989" border="0" cellpadding="0" cellspacing="3" >
    <tr>
      <!--First Row-->
      <td width="280"  valign="top"><!--User Login -->
        <!--<fieldset>
				
                <img  src="Image20100223/User_login_Title20100228.jpg"  /><br />
				<%' call ShowUserLogin() %>
            </fieldset>-->
        <table cellpadding="0" cellspacing="0" border="0" width="280">
          <tr>
            <td background="Image20100223/User_Login_Top.jpg" width="280" height="5"><!--<img src="Image20100223/User_Login_Top.jpg" width="280" />--></td>
          </tr>
          <tr  >
            <td   width="280" height="24" background="Image20100223/User_Login_Title.jpg"><!--<img src="Image20100223/User_Login_Title.jpg" />--></td>
          </tr>
          <tr>
            <td   width="280"  background="Image20100223/User_Login_Middle.jpg"><% call ShowUserLogin() %></td>
          </tr>
          <tr>
            <td width="280" height="5" background="Image20100223/User_Login_Bottom.jpg"    ><!--<img src="Image20100223/User_Login_Bottom.jpg" width="280" />--></td>
          </tr>
        </table></td>
      <!--User Login -->
      <td rowspan="2" valign="top" width="429" ><!--hot articles begin-->
        <table cellpadding="0" cellspacing="0" border="0" width="429">
          <tr>
            <td  valign="bottom" width="429" height="36"  background="Image20100223/TDTopHotTitleBackground20100228.jpg"><!--<font color="#FF0000"  size="4" ><strong>&nbsp;&nbsp;&nbsp;&nbsp;推 荐 文 章</strong></font>--></td>
          </tr>
          <!-- 由列出推荐文章的函数描绘表格。-->
          <% call ShowhotAtIndex(11,40) %>
          <tr>
            <td   valign="top"  width="429" height="3" background="Image20100223/TDBottom20100302.jpg"><!--<img src="Image20100223/TDBottom20100302.jpg" width="429" height="3" /> --></td>
          </tr>
        </table>
        <!--  <fieldset >
			  <%' call ShowhotAtIndex(12,48) %>
              </fieldset>--></td>
      <!--hot articles end-->
      <td width="280" valign="top" ><!--anouncement begin-->
        <table width="280" cellpadding="0" cellspacing="0" border="0" >
          <tr>
            <td width="280"  height="23" background="Image20100223/TDTopAnnounceTitleBackground20100310.jpg"></td>
          </tr>
          <tr>
            <td  width="280"   height="160" background="Image20100223/TDMiddle280RedBackground20100302.jpg"><%call ShowAnnounce(1,1)%></td>
          </tr>
          <tr>
            <td width="280"  height="2" background="Image20100223/TDBottom20100302.jpg"></td>
          </tr>
        </table></td>
      <!--anouncement end-->
    </tr>
    <!--/first row-->
    <tr>
      <!--second row-->
      <!--Special List begin-->
      <!--<td width="280"> 
      			<table cellpadding="0" cellspacing="0" border="0" width="280">
      				<tr>
                    	<td   valign="top" width="280" height="23" background="Image20100223/TDTopSpecialTitleBackground20100310.jpg">
                        	
                        </td>
                        </tr>		
      			</table>	</td>-->
      <!--Special List end-->
      <!--  <OBJECT classid="clsid:CFCDAA03-8BE4-11cf-B84B-0020AFBBCCFA" height="265" width="500">    
<PARAM NAME="_ExtentX" VALUE="10372">    
<PARAM NAME="_ExtentY" VALUE="6456">    
<PARAM NAME="SRC" VALUE="UpLoadFiles/视频标准日本语初级00.rm">    
<PARAM NAME="AUTOSTART" VALUE="-1">    
<PARAM NAME="SHUFFLE" VALUE="0">    
<PARAM NAME="PREFETCH" VALUE="0">    
<PARAM NAME="NOLABELS" VALUE="0">    
<PARAM NAME="CONTROLS" VALUE="IMAGEWINDOW">    
<PARAM NAME="CONSOLE" VALUE="Clip528211525">    
<PARAM NAME="LOOP" VALUE="0">    
<PARAM NAME="NUMLOOP" VALUE="0">    
<PARAM NAME="CENTER" VALUE="0">    
<PARAM NAME="MAINTAINASPECT" VALUE="0">    
<PARAM NAME="BACKGROUNDCOLOR" VALUE="#000000">    
<embed _extentx="10372" _extenty="6456" autostart="0" src="UpLoadFiles/视频标准日本语初级00.rm" shuffle="0" prefetch="0" nolabels="0" controls="IMAGEWINDOW" console="Clip528211525" loop="0" numloop="0" center="0" maintainaspect="0" backgroundcolor="#000000">    
</embed>    
</OBJECT> -->
      <td  valign="top" width="280" >
      <!--Pictures Flash Animation-->
      <div class="flashnews"> 
     <table  border="0" cellpadding="0" cellspacing="0" align="center" width="280" height="240"> 
  <tr> 
    <td background="images/bo05-2.gif" width="5" height="4"></td> 
    <td background="images/bo04-2.gif" height="4"></td> 
    <td background="images/bo03-2.gif" width="5" height="4"></td> 
  </tr> 
  <tr> 
    <td  background="images/bo01.gif" width="5"></td> 
    <td align="center" valign="middle" width="270" height="232"> <!-- size: 280px * 192px --> 
      <script language='javascript'> 
linkarr = new Array();
picarr = new Array();
textarr = new Array();
var swf_width=270;
var swf_height=232;
var files = "";
var links = "";
var texts = "";
//这里设置调用标记
linkarr[1] = "Article_Show.asp?ArticleID=873";
picarr[1]  = "UploadFiles/20081211172357736.jpg";
textarr[1] = "马克思主义在中国的广泛传播";
linkarr[2] = "Article_Show.asp?ArticleID=884";
picarr[2]  = "UploadFiles/2008121801859289.jpg";
textarr[2] = "鲜花献给党";


 
for(i=1;i<picarr.length;i++){
  if(files=="") files = picarr[i];
  else files += "|"+picarr[i];
}
for(i=1;i<linkarr.length;i++){
  if(links=="") links = linkarr[i];
  else links += "|"+linkarr[i];
}
for(i=1;i<textarr.length;i++){
  if(texts=="") texts = textarr[i];
  else texts += "|"+textarr[i];
}
document.write('<object classid="clsid:d27cdb6e-ae6d-11cf-96b8-444553540000" codebase="http://fpdownload.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,0,0" width="'+ swf_width +'" height="'+ swf_height +'">');
document.write('<param name="movie" value="/templets/images/bcastr3.swf"><param name="quality" value="high">');
document.write('<param name="menu" value="false"><param name=wmode value="opaque">');
document.write('<param name="FlashVars" value="bcastr_file='+files+'&bcastr_link='+links+'&bcastr_title='+texts+'">');
document.write('<embed src="/templets/images/bcastr3.swf" wmode="opaque" FlashVars="bcastr_file='+files+'&bcastr_link='+links+'&bcastr_title='+texts+'& menu="false" quality="high" width="'+ swf_width +'" height="'+ swf_height +'" type="application/x-shockwave-flash" pluginspage="http://www.macromedia.com/go/getflashplayer" />'); document.write('</object>'); 
</script></td> 
    <td background="images/bo02.gif" width="5"></td> 
  </tr> 
  <tr> 
    <td background="images/bo03.gif" width="5" height="4"></td> 
    <td background="images/bo04.gif" height="4"></td> 
    <td background="images/bo05.gif" width="5" height="4"></td> 
  </tr> 
</table> 
    </div><!--end Pictures Flash Animation-->
      
       </td>
      
      
    <!--Windows Media Player成功调用-->  <!--<object align=middle class=OBJECT classid=CLSID:22d6f312-b0f6-11d0-94ab-0080c74c7e95  id=MediaPlayer >
          <param name="ShowStatusBar" value="-1">
          <param name="Filename" value="UploadFiles/H.avi">
          <embed type=application/x-oleobject codebase=http://activex.microsoft.com/activex/controls/mplayer/en/nsmp2inf.cab#Version=5,1,52,701 flename=mp src="UploadFiles/H.avi" width=280 height=210> </embed>
        </object>--><!--Windows Media Player成功调用-->
        
       
        
        
      <!--Windows Media Player成功调用-->
      <!--<embed src="UploadFiles/视频标准日本语初级00.rm">-->
      <!-- <object classid=clsid:CFCDAA03-8BE4-11cf-B84B-0020AFBBCCFA height=285 id=RAOCX name=rmplay width=356>    
<param name="SRC" value="影片地址">    
<param name="CONSOLE" value="Clip1">    
<param name="CONTROLS" value="imagewindow">    
<param name="AUTOSTART" value="true">    
<embed src="UploadFiles/视频标准日本语初级00.rm" autostart="true" controls="ImageWindow" console="Clip1" pluginspage="http://www.real.com" target="_blank"  width="356" height="285">    
</embed>    
</object>-->
      <!--<OBJECT ID=RVOCX CLASSID="clsid:CFCDAA03-8BE4-11cf-B84B-0020AFBBCCFA"WIDTH=300 HEIGHT=134 src="UploadFiles/视频标准日本语初级00.rm"><PARAM NAME="src" VALUE="UploadFiles/视频标准日本语初级00.rm"></OBJECT>-->
      <!--Show Guest begin-->
      <td width="280" valign="top"><table width="280" cellpadding="0" cellspacing="0" border="0"  >
          <tr>
            <td  width="280" height="23"   background="Image20100223/TDTopGuestTitleBackground20100310.jpg" ></td>
          </tr>
          <% call showGuestAtIndex(28,11) %>
          <tr>
            <td height="2" background="Image20100223/TDBottom20100302.jpg" ></td>
          </tr>
        </table></td>
      <!--Show Guest End-->
    </tr>
    <!--/second row-->
  </table>
  <!--main body-->
</div>
<div align="center">
  <table >
    <tr>
      <td colspan="12" width="989" height="50" background="Image20100223/BottomPic20100228.jpg"><P align=center><B>| <SPAN style="CURSOR: hand" 
      onclick="var strHref=window.location.href;this.style.behavior='url(#default#homepage)';this.setHomePage('http://renwen.university.edu.cn');">设为首页</SPAN> | <SPAN title='两课教学网' style="CURSOR: hand" 
      onclick="window.external.addFavorite('http://renwen.university.edu.cn','两课教学网')">收藏本站</SPAN> | <A  href="mailto:86277298@QQ.COM">联系站长</A> | <A  
      href="http://renwen.university.edu.cn/FriendSite/Index.asp" target=_blank>友情链接</A> | <A  href="http://renwen.university.edu.cn/Copyright.asp" 
      target=_blank>版权申明</A> | </B><br>
          本网站由<font color="#3300FF"><a href="http://renwen.university.edu.cn/">university人文社会科学学院</a></font>主办、维护</P></td>
    </tr>
  </table>
</div>
<!-- End ImageReady Slices -->
</body>
</html>