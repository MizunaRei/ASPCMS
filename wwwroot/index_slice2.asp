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
<!-- ImageReady Slices (首页_slice.psd) -->
<table id="__01" width="1025"  border="0" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF"> <!--height="1001"-->
	<!--<tr>
		<td colspan="14">
			<img src="images/首页_slice_01.jpg" width="1024" height="142" alt=""></td>
		<td>
			<img src="images/分隔符.gif" width="1" height="142" alt=""></td>
	</tr>-->
	<tr>
		<td rowspan="16">
			<img src="images/首页_slice_02.jpg" width="17" height="858" alt=""></td>
		<td colspan="12">
			<img src="images/首页_slice_03.jpg" width="989" height="188" alt=""></td>
		<td rowspan="16">
			<img src="images/首页_slice_04.jpg" width="18" height="858" alt=""></td>
		<td>
			<img src="images/分隔符.gif" width="1" height="188" alt=""></td>
	</tr>
	<tr>
		<td colspan="12"  background="images/首页_slice_05.jpg" width="989" height="25">
	    <!--<img src="images/首页_slice_05.jpg" width="989" height="25" alt="">-->  &nbsp; <a href="index.asp"><!--<font color="#000000" >-->首页<!--</font>--></a>    &nbsp; <a href="Article_Class2.asp?ClassID=1"><!--<font color="#000000">-->资料中心<!--</font>--></a>   &nbsp;<a href="Article_Class2.asp?ClassID=2"><!--<font color="#000000">-->理论动态<!--</font>--></a>  &nbsp; <a href="Article_Class2.asp?ClassID=3"><!--<font color="#000000">-->时事新闻<!--</font>--></a>   &nbsp;<a href="Article_Class2.asp?ClassID=58"><!--<font color="#000000">-->学生作品<!--</font>--></a>     &nbsp;      <a href="userlist.asp"><!--<font color="#000000">-->文集<!--</font>--></a>   &nbsp;<a href="guestbook.asp"><!--<font color="#000000">-->留言<!--</font>--></a> 
         <% call ShowSearchForm("Article_Search.asp",1) %></td>
		
        <td>
			<img src="images/分隔符.gif" width="1" height="25" alt=""></td>
	</tr>
	<tr>
		<td colspan="12">
			<img src="images/首页_slice_06.jpg" width="989" height="3" alt=""></td>
		<td>
			<img src="images/分隔符.gif" width="1" height="3" alt=""></td>
	</tr>
	<tr>
		<td colspan="3" rowspan="2"  width="231" height="96">
			<!--<img src="images/首页_slice_07.jpg" width="231" height="96" alt="">--><fieldset><legend><!--<font color="#666666">-->用户登录</legend><% call ShowUserLogin() %></fieldset></td>
		<td colspan="5" rowspan="3"  width="529" height="199" background="images/首页_slice_08.jpg">
			<!--<img src="images/首页_slice_08.jpg" width="529" height="199" alt="">-->人气文章</td>
		<td colspan="4">
			<img src="images/首页_slice_09.jpg" width="229" height="13" alt=""></td>
		<td>
			<img src="images/分隔符.gif" width="1" height="13" alt=""></td>
	</tr>
	<tr>
		<td rowspan="3">
			<img src="images/首页_slice_10.jpg" width="10" height="202" alt=""></td>
		<td colspan="2" rowspan="2">
			<img src="images/首页_slice_11.jpg" width="210" height="186" alt=""></td>
		<td rowspan="10">
			<img src="images/首页_slice_12.jpg" width="9" height="382" alt=""></td>
		<td>
			<img src="images/分隔符.gif" width="1" height="83" alt=""></td>
	</tr>
	<tr>
		<td rowspan="9">
			<img src="images/首页_slice_13.jpg" width="2" height="299" alt=""></td>
		<td rowspan="4">
			<img src="images/首页_slice_14.jpg" width="210" height="127" alt=""></td>
		<td rowspan="9">
			<img src="images/首页_slice_15.jpg" width="19" height="299" alt=""></td>
		<td>
			<img src="images/分隔符.gif" width="1" height="103" alt=""></td>
	</tr>
	<tr>
		<td colspan="5">
			<img src="images/首页_slice_16.jpg" width="529" height="16" alt=""></td>
		<td colspan="2">
			<img src="images/首页_slice_17.jpg" width="210" height="16" alt=""></td>
		<td>
			<img src="images/分隔符.gif" width="1" height="16" alt=""></td>
	</tr>
	<tr>
		<td rowspan="7">
			<img src="images/首页_slice_18.jpg" width="1" height="180" alt=""></td>
		<td colspan="7">
			<img src="images/首页_slice_19.jpg" width="748" height="6" alt=""></td>
		<td>
			<img src="images/分隔符.gif" width="1" height="6" alt=""></td>
	</tr>
	<tr>
		<td rowspan="6">
			<img src="images/首页_slice_20.jpg" width="14" height="174" alt=""></td>
		<td rowspan="2">
			<img src="images/首页_slice_21.jpg" width="356" height="14" alt=""></td>
		<td rowspan="6">
			<img src="images/首页_slice_22.jpg" width="9" height="174" alt=""></td>
		<td colspan="3" rowspan="2">
			<img src="images/首页_slice_23.jpg" width="356" height="14" alt=""></td>
		<td rowspan="6">
			<img src="images/首页_slice_24.jpg" width="13" height="174" alt=""></td>
		<td>
			<img src="images/分隔符.gif" width="1" height="2" alt=""></td>
	</tr>
	<tr>
		<td rowspan="2">
			<img src="images/首页_slice_25.jpg" width="210" height="25" alt=""></td>
		<td>
			<img src="images/分隔符.gif" width="1" height="12" alt=""></td>
	</tr>
	<tr>
		<td rowspan="2">
			<img src="images/首页_slice_26.jpg" width="356" height="79" alt=""></td>
		<td colspan="3" rowspan="2">
			<img src="images/首页_slice_27.jpg" width="356" height="79" alt=""></td>
		<td>
			<img src="images/分隔符.gif" width="1" height="13" alt=""></td>
	</tr>
	<tr>
		<td rowspan="3">
			<img src="images/首页_slice_28.jpg" width="210" height="147" alt=""></td>
		<td>
			<img src="images/分隔符.gif" width="1" height="66" alt=""></td>
	</tr>
	<tr>
		<td>
			<img src="images/首页_slice_29.jpg" width="356" height="13" alt=""></td>
		<td colspan="3">
			<img src="images/首页_slice_30.jpg" width="356" height="13" alt=""></td>
		<td>
			<img src="images/分隔符.gif" width="1" height="13" alt=""></td>
	</tr>
	<tr>
		<td>
			<img src="images/首页_slice_31.jpg" width="356" height="68" alt=""></td>
		<td colspan="3">
			<img src="images/首页_slice_32.jpg" width="356" height="68" alt=""></td>
		<td>
			<img src="images/分隔符.gif" width="1" height="68" alt=""></td>
	</tr>
	<tr>
		<td colspan="12">
			<img src="images/首页_slice_33.jpg" width="989" height="6" alt=""></td>
		<td>
			<img src="images/分隔符.gif" width="1" height="6" alt=""></td>
	</tr>
	<tr>
		<td colspan="12">
			<img src="images/首页_slice_34.jpg" width="989" height="241" alt=""></td>
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