<!--#include file="Inc/syscode_guest.asp"-->
<%
'请勿改动下面这三行代码
const ChannelID=5
Const ShowRunTime="Yes"
SkinID=0
%>
<html>
<head>
<title><%=strPageTitle & " >> " & PageTitle%></title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<!--#include file="inc/Skin_CSS.asp"-->
<%call MenuJS()%>
</head>
<body <%=Body_Label%> onmousemove='HideMenu()' bgcolor=#ffffff style="BACKGROUND-COLOR: #ffffff">
<div id=menuDiv style='Z-INDEX: 1000; VISIBILITY: hidden; WIDTH: 1px; POSITION: absolute; HEIGHT: 1px; BACKGROUND-COLOR: #9cc5f8'></div>
<!--标题栏-->
<table width="989"><tr><td> <%	call Top_noIndex() %> </td></tr>
<tr><td valign="top">
<div align="left"><!--main body-->
		<table bgcolor="#ffffff" >
		<tr><td valign="top"><!-- start left column-->
		<table width="231">
				<tr>
				<td>
            	<fieldset><legend>用户登录</legend><% call ShowUserLogin() %></fieldset>
                </td>
				</tr>
		<tr><td align="center"  background="Images/课程列表副本_06.jpg"  width="231" >&nbsp;&nbsp;&nbsp;&nbsp;<strong>站内导航</strong>
		</td></tr>
		<tr><td>
		  <ul>
			<li><a href="index.asp">【首&nbsp;&nbsp;&nbsp;&nbsp;页】</a>&nbsp;&nbsp;<a href="article_specialList.asp">【课程列表】</a></li>
		    <li><a href="Article_Class2.asp?ClassID=1">【理论动态】</a>&nbsp;&nbsp;<a href="Article_Class2.asp?ClassID=2">【资料中心】</a> </li>
		    <li><a href="Article_Class2.asp?ClassID=3">【时事新闻】</a>&nbsp;&nbsp;<a href="Article_Class2.asp?ClassID=58">【学生作品】</a> </li>
			</ul>
		</td>
		</tr>
		<tr><td background="Images/课程列表副本_06.jpg" align="center" width="231" > &nbsp;&nbsp;&nbsp;&nbsp;<strong>留言板</strong>
		</td></tr>
		<tr><td><% call GuestBook_Left() %></td></tr>
        <tr><td background="Images/课程列表副本_06.jpg" align="center" width="231" > &nbsp;&nbsp;&nbsp;&nbsp;<strong>留言搜索</strong>
        	</td></tr>
      	<tr><td>  <% call GuestBook_Search() %>
        </td></tr>
		</table>
		<!-- end left column--></td>
		<td><!--right column-->
        	<td valign="top"> 
      <table width="100%" border="0" cellpadding="0" cellspacing="0"  background="images/fcbg1_1.gif">
        <tr> 
          <td width="92"><strong><img src="Images/announce.gif" width="20" height="16" align="absmiddle">&nbsp;最新公告</strong></td>
          <td width="683"><div align="right"> 
        <MARQUEE scrollAmount=1 scrollDelay=4 width=480
            align="left" onMouseOver="this.stop()" onMouseOut="this.start()">
        <% call ShowAnnounce(2,5) %>
        </MARQUEE>
      </div></td>
        </tr>
      </table>
	  <%
		call showtip()
		call Guestbook()
	  %>
      <table width="100%" border="0" align="left" align="center" cellpadding="0" cellspacing="0" class="tdbg_rightall">
        <tr background="images/fcbg2.gif"> 
          <td> 
			<%
				call ShowGuestPage()
			%>
          </td>
        </tr>
      </table>
    </td>
        <!--end right column--></td>
		</tr><!--end main bocy-->
		</table><!--end main bocy-->
	</div>
</td><!--the great talbe--></tr>
<!--页面底部--> <tr><td>    <!--Bottom--><%  call Bottom_All()  %></td></tr>
    </table><!--the great talbe--><!--结束页面底部-->
<% call PopAnnouceWindow(400,300) %>
</body>
</html>
<%
call CloseConn()
%>