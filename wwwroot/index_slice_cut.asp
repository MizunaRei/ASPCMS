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
<SCRIPT language=JavaScript src="SkinIndex/menu.js" 
type=text/JavaScript></SCRIPT>
<META 
content="<%=strPageTitle%>:资源免费，更新快，资源全，提供本科、硕士研究生的各种思想政治课和马克思主义理论课教学资源，栏目有：理论动态、资料中心、时事新闻、学生作品。" 
name=Keywords>
<META 
content="<%=strPageTitle%>:资源免费，更新快，资源全，提供本科、硕士研究生的各种思想政治课和马克思主义理论课教学资源，栏目有：理论动态、资料中心、时事新闻、学生作品。" 
name=Description>
<META content=o7FhrjMKBn/3XGgcDXmGdE4BkAxwd6a97bpMEXpOURY= name=verify-v1>
<META content="MSHTML 6.00.2900.3395" name=GENERATOR>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<!--#include file="inc/Skin_CSS.asp"-->
<%call MenuJS()%>


</head>
<body bgcolor="#FFFFFF" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<!-- ImageReady Slices (首页_slice2.psd) -->
<table id="__01" width="1025" height="1001" border="0" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
	<!--<tr>
		<td colspan="14">
			<img src="images/首页_slice2_01.jpg" width="1024" height="142" alt=""></td>
		<td>
			<img src="images/分隔符.gif" width="1" height="142" alt=""></td>
	</tr>-->
	<tr>
		<td rowspan="17">
			<img src="images/首页_slice2_02.jpg" width="17" height="858" alt=""></td>
		<td colspan="12">
			<img src="images/首页_slice2_03.jpg" width="989" height="188" alt=""></td>
		<td rowspan="17">
			<img src="images/首页_slice2_04.jpg" width="18" height="858" alt=""></td>
		<td>
			<img src="images/分隔符.gif" width="1" height="188" alt=""></td>
	</tr>
	<tr>
		<!--<td colspan="12">
			<img src="images/首页_slice2_05.jpg" width="989" height="25" alt=""></td>-->
            <td colspan="12"  background="images/首页_slice_05.jpg" width="989" height="25">
	    <!--<img src="images/首页_slice_05.jpg" width="989" height="25" alt="">-->  &nbsp; <a href="index.asp"><!--<font color="#000000" >-->首页<!--</font>--></a>    &nbsp; <a href="Article_Class2.asp?ClassID=1"><!--<font color="#000000">-->资料中心<!--</font>--></a>   &nbsp;<a href="Article_Class2.asp?ClassID=2"><!--<font color="#000000">-->理论动态<!--</font>--></a>  &nbsp; <a href="Article_Class2.asp?ClassID=3"><!--<font color="#000000">-->时事新闻<!--</font>--></a>   &nbsp;<a href="Article_Class2.asp?ClassID=58"><!--<font color="#000000">-->学生作品<!--</font>--></a>     &nbsp;      <a href="userlist.asp"><!--<font color="#000000">-->文集<!--</font>--></a>   &nbsp;<a href="guestbook.asp"><!--<font color="#000000">-->留言<!--</font>--></a> 
         <% call ShowSearchForm("Article_Search.asp",1) %></td>
		<td>
			<img src="images/分隔符.gif" width="1" height="25" alt=""></td>
	</tr>
	<tr>
		<td colspan="12">
			<img src="images/首页_slice2_06.jpg" width="989" height="3" alt=""></td>
		<td>
			<img src="images/分隔符.gif" width="1" height="3" alt=""></td>
	</tr>
	<tr>
		<td colspan="3" rowspan="3">
			<!--<img src="images/首页_slice2_07.jpg" width="231" height="116" alt="">-->
            <fieldset><legend><!--<font color="#666666">-->用户登录</legend><% call ShowUserLogin() %></fieldset>
            </td>
		<td colspan="5" rowspan="2" width="529" height="24" background="Images/首页_slice2_08.jpg">
			<!--<img src="images/首页_slice2_08.jpg" width="529" height="24" alt="">-->
            人气文章
            
            </td>
		<td colspan="4">
			<img src="images/首页_slice2_09.jpg" width="229" height="13" alt=""></td>
		<td>
			<img src="images/分隔符.gif" width="1" height="13" alt=""></td>
	</tr>
	<tr>
		<td rowspan="4">
			<img src="images/首页_slice2_10.jpg" width="10" height="202" alt=""></td>
		<td colspan="2" rowspan="3" width="210" height="186" background="images/首页_slice2_11.jpg">
			<!--<img src="images/首页_slice2_11.jpg" width="210" height="186" alt="">-->
            <%call ShowAnnounce(1,1)%>
            </td>
		<td rowspan="11">
			<img src="images/首页_slice2_12.jpg" width="9" height="382" alt=""></td>
		<td>
			<img src="images/分隔符.gif" width="1" height="11" alt=""></td>
	</tr>
	<tr>
		<td colspan="5" rowspan="2" width="529" height="175" background="images/首页_slice2_13.jpg">
			<!--<img src="images/首页_slice2_13.jpg" width="529" height="175" alt="">-->
             <% call Showhot(8,16) %>
            
            </td>
		<td>
			<img src="images/分隔符.gif" width="1" height="92" alt=""></td>
	</tr>
	<tr>
		<td rowspan="9">
			<img src="images/首页_slice2_14.jpg" width="2" height="279" alt=""></td>
		<td rowspan="5" width="210" height="127" >
        <table>
			<tr><td background="images/首页_slice2_15_top.jpg">
            <!--<img src="images/首页_slice2_15.jpg" width="210" height="127" alt="">-->&nbsp;课程列表</td></tr>
            <tr><td background="images/首页_slice2_15_middle.jpg">   <% call ShowSpecial(10) %>
            </td></tr>
            <tr><td background="images/首页_slice2_15_bottom.jpg"> &nbsp;</td></tr>
            </table>
            </td>
		<td rowspan="9">
			<img src="images/首页_slice2_16.jpg" width="19" height="279" alt=""></td>
		<td>
			<img src="images/分隔符.gif" width="1" height="83" alt=""></td>
	</tr>
	<tr>
		<td colspan="5">
			<img src="images/首页_slice2_17.jpg" width="529" height="16" alt=""></td>
		<td colspan="2">
			<img src="images/首页_slice2_18.jpg" width="210" height="16" alt=""></td>
		<td>
			<img src="images/分隔符.gif" width="1" height="16" alt=""></td>
	</tr>
	<tr>
		<td rowspan="7">
			<img src="images/首页_slice2_19.jpg" width="1" height="180" alt=""></td>
		<td colspan="7">
			<img src="images/首页_slice2_20.jpg" width="748" height="6" alt=""></td>
		<td>
			<img src="images/分隔符.gif" width="1" height="6" alt=""></td>
	</tr>
	<tr>
		<td rowspan="6">
			<img src="images/首页_slice2_21.jpg" width="14" height="174" alt=""></td>
		<td width="356" height="14" background="images/首页_slice2_22.jpg">
        <a href='Article_Class2.asp?ClassID=2'> &nbsp;&nbsp;资料中心 &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;更多>></a>
			<!--<img src="images/首页_slice2_22.jpg" width="356" height="14" alt="">--></td>
		<td rowspan="6">
			<img src="images/首页_slice2_23.jpg" width="9" height="174" alt=""></td>
		<td colspan="3" width="356" height="14" background="images/首页_slice2_24.jpg">
		<a href='Article_Class2.asp?ClassID=3'>	&nbsp;&nbsp;时事新闻&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;更多>></a>
            <!--<img src="images/首页_slice2_24.jpg" width="356" height="14" alt="">--></td>
		<td rowspan="6">
			<img src="images/首页_slice2_25.jpg" width="13" height="174" alt=""></td>
		<td>
			<img src="images/分隔符.gif" width="1" height="14" alt=""></td>
	</tr>
	<tr>
		<td rowspan="3">
			<!--<img src="images/首页_slice2_26.jpg" width="356" height="79" alt="">-->
          <%  call ShowArticle_Index(5,2,-1,10) %>
            </td>
		<td colspan="3" rowspan="3" width="356" height="79" background="images/首页_slice2_27.jpg">
			<!--<img src="images/首页_slice2_27.jpg" width="356" height="79" alt="">-->
            <%  call ShowArticle_Index(5,3,-1,10) %>
            </td>
		<td>
			<img src="images/分隔符.gif" width="1" height="8" alt=""></td>
	</tr>
	<tr>
		<td>
			<img src="images/首页_slice2_28.jpg" width="210" height="5" alt=""></td>
		<td>
			<img src="images/分隔符.gif" width="1" height="5" alt=""></td>
	</tr>
	<tr>
		<td rowspan="3" width="210" height="147" background="images/首页_slice2_29.jpg">
			<!--<img src="images/首页_slice2_29.jpg" width="210" height="147" alt="">-->
            留言板<br /><% call showGuest(20,10) %>
            </td>
		<td>
			<img src="images/分隔符.gif" width="1" height="66" alt=""></td>
	</tr>
	<tr>
		<td width="356" height="13" background="images/首页_slice2_30.jpg">
			<!--<img src="images/首页_slice2_30.jpg" width="356" height="13" alt="">-->
             
        <a href='Article_Class2.asp?ClassID=1'>   &nbsp; 理论动态&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;更多>></a></td>
		<td colspan="3" width="356" height="13" background="images/首页_slice2_31.jpg">
			<!--<img src="images/首页_slice2_31.jpg" width="356" height="13" alt="">-->
            <a href='Article_Class2.asp?ClassID=58'>  &nbsp;  学生作品&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;更多>></a>
            </td>
		<td>
			<img src="images/分隔符.gif" width="1" height="13" alt=""></td>
	</tr>
	<tr>
		<td width="356" height="68"  background="images/首页_slice2_32.jpg">
			<!--<img src="images/首页_slice2_32.jpg" width="356" height="68" alt="">-->
			<%  call ShowArticle_Index(5,1,-1,10) %></td>
		<td colspan="3"  width="356" height="68"  background="images/首页_slice2_33.jpg">
			<!--<img src="images/首页_slice2_33.jpg" width="356" height="68" alt="">-->
            
            </td>
		<td>
			<img src="images/分隔符.gif" width="1" height="68" alt=""></td>
	</tr>
	<tr>
		<td colspan="12">
			<img src="images/首页_slice2_34.jpg" width="989" height="6" alt=""></td>
		<td>
			<img src="images/分隔符.gif" width="1" height="6" alt=""></td>
	</tr>
	<tr>
		<td colspan="12" width="989" height="241" background="images/首页_slice2_35.jpg">
			<!--<img src="images/首页_slice2_35.jpg" width="989" height="241" alt="">-->
            <P align=center><B>| <SPAN style="CURSOR: hand" 
      onclick="var strHref=window.location.href;this.style.behavior='url(#default#homepage)';this.setHomePage('http://renwen.university.edu.cn');">设为首页</SPAN> | <SPAN title='两课教学网' style="CURSOR: hand" 
      onclick="window.external.addFavorite('http://renwen.university.edu.cn','两课教学网')">收藏本站</SPAN> | <A  href="mailto:86277298@QQ.COM">联系站长</A> | <A  
      href="http://renwen.university.edu.cn/FriendSite/Index.asp" target=_blank>友情链接</A> | <A  href="http://renwen.university.edu.cn/Copyright.asp" 
      target=_blank>版权申明</A> | </B></P>
      <p align="center">本网站由<font color="#3300FF"><a href="http://renwen.university.edu.cn/">university人文社会科学学院</a></font>主办、维护</p>
            </td>
		<td>
			<img src="images/分隔符.gif" width="1" height="241" alt=""></td>
	</tr>
	<tr>
		<td>
			<img src="images/分隔符.gif" width="17" height="1" alt=""></td>
		<td>
			<img src="images/分隔符.gif" width="2" height="1" alt=""></td>
		<td>
			<img src="images/分隔符.gif" width="210" height="1" alt=""></td>
		<td>
			<img src="images/分隔符.gif" width="19" height="1" alt=""></td>
		<td>
			<img src="images/分隔符.gif" width="1" height="1" alt=""></td>
		<td>
			<img src="images/分隔符.gif" width="14" height="1" alt=""></td>
		<td>
			<img src="images/分隔符.gif" width="356" height="1" alt=""></td>
		<td>
			<img src="images/分隔符.gif" width="9" height="1" alt=""></td>
		<td>
			<img src="images/分隔符.gif" width="149" height="1" alt=""></td>
		<td>
			<img src="images/分隔符.gif" width="10" height="1" alt=""></td>
		<td>
			<img src="images/分隔符.gif" width="197" height="1" alt=""></td>
		<td>
			<img src="images/分隔符.gif" width="13" height="1" alt=""></td>
		<td>
			<img src="images/分隔符.gif" width="9" height="1" alt=""></td>
		<td>
			<img src="images/分隔符.gif" width="18" height="1" alt=""></td>
		<td></td>
	</tr>
</table>
<!-- End ImageReady Slices -->
</body>
</html>