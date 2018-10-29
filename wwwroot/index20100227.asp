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
<META content=o7FhrjMKBn/3XGgcDXmGdE4BkAxwd6a97bpMEXpOURY= name=verify-v1>
<META content="MSHTML 6.00.2900.3395" name=GENERATOR>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<!--#include file="inc/Skin_CSS.asp"-->
<%call MenuJS()%>
</head>
<body bgcolor="#FFFFFF"  style="BACKGROUND-COLOR: #ffffff"  leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<!-- ImageReady Slices (首页_slice2.psd) -->
<div align="center" ><table id="__01" width="989"  border="0" cellpadding="0" cellspacing="0" >
	<tr>
		<td colspan="20">
			<img   src="Image20100223/TopPic20100227.jpg"  alt="红桥网" width="989" ></td>
	</tr>
	
    
    </table><!--top--></div>
   <div align="center"><!--main body--><table  bgcolor="#FFFFFF" width="989">
   <tr><!--First Row-->
   		<td><!--User Login -->
        	<table >
            	<tr><td background="Image20100223/User_Login_Top.jpg" width="278" height="5"><!--<img src="Image20100223/User_Login_Top.jpg" width="280" />--></td></tr>
                <tr  ><td   width="278" height="24" background="Image20100223/User_Login_Title.jpg"><!--<img src="Image20100223/User_Login_Title.jpg" />--></td></tr>
                <tr><td   width="278" background="Image20100223/User_Login_Middle.jpg"><% call ShowUserLogin() %></td></tr>
                <tr><td width="278" height="5" background="Image20100223/User_Login_Bottom.jpg"    ><!--<img src="Image20100223/User_Login_Bottom.jpg" width="280" />--></td></tr>
        	</table>
            </td><!--User Login -->
   </tr>
    
	
	
	
	
	
	
	
	
	
	</table><!--main body--></div>
<div align="center">	<table >
	<tr>
		<td colspan="12" width="989" height="50" background="Image20100223/BottomPic20100228.jpg">
            <P align=center><B>| <SPAN style="CURSOR: hand" 
      onclick="var strHref=window.location.href;this.style.behavior='url(#default#homepage)';this.setHomePage('http://renwen.university.edu.cn');">设为首页</SPAN> | <SPAN title='两课教学网' style="CURSOR: hand" 
      onclick="window.external.addFavorite('http://renwen.university.edu.cn','两课教学网')">收藏本站</SPAN> | <A  href="mailto:86277298@QQ.COM">联系站长</A> | <A  
      href="http://renwen.university.edu.cn/FriendSite/Index.asp" target=_blank>友情链接</A> | <A  href="http://renwen.university.edu.cn/Copyright.asp" 
      target=_blank>版权申明</A> | </B><br>
      本网站由<font color="#3300FF"><a href="http://renwen.university.edu.cn/">university人文社会科学学院</a></font>主办、维护</P>
            </td>
	</tr>
</table></div>
<!-- End ImageReady Slices -->
</body>
</html>