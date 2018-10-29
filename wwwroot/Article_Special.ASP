<!--#include file="Inc/syscode_article.asp"-->
<%
'请勿改动下面这三行代码
const ChannelID=2
Const ShowRunTime="Yes"
MaxPerPage=20
strFileName="Article_Special.asp?ClassID=" & ClassID & "&SpecialID=" & SpecialID
%>
<html>
<head>
<title><%=strPageTitle%></title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">

<!--#include file="inc/Skin_CSS.asp"-->
<%call MenuJS()%>
</head>
<body <%=Body_Label%> onmousemove='HideMenu()'  bgcolor=#ffffff style="BACKGROUND-COLOR: #ffffff">
<table><!--the great talbe--><tr><td>
<%	call Top_noIndex() %></td></tr>

<tr><td><div align="center"><!--main body-->
		<table bgcolor="#ffffff" >
		<tr><td><!-- start left column-->
		<table width="231">
				<tr>
				<td>
            	<fieldset><legend>用户登录</legend><% call ShowUserLogin() %></fieldset>
                </td>
				</tr>
		<tr><td   background="Images/课程列表副本_06.jpg"  width="231" >&nbsp;&nbsp;&nbsp;&nbsp;站内导航
		
		
		</td></tr>
		
		<tr><td>
		  <ul>
			<li><a href="index.asp">【首&nbsp;&nbsp;&nbsp;&nbsp;页】</a>&nbsp;&nbsp;<a href="article_specialList.asp">【课程列表】</a></li>
		    <li><a href="Article_Class2.asp?ClassID=1">【理论动态】</a>&nbsp;&nbsp;<a href="Article_Class2.asp?ClassID=2">【资料中心】</a> </li>
		    <li><a href="Article_Class2.asp?ClassID=3">【时事新闻】</a>&nbsp;&nbsp;<a href="Article_Class2.asp?ClassID=58">【学生作品】</a> </li>
			</ul>
		</td>
		</tr>
		<tr><td background="Images/课程列表副本_06.jpg"  width="231" > &nbsp;&nbsp;&nbsp;&nbsp;人气文章
		</td></tr>
		<tr><td><% call Showhot(8,16) %></td></tr>
        <tr><td background="Images/课程列表副本_06.jpg"  width="231" > &nbsp;&nbsp;&nbsp;&nbsp;推荐文章
        	</td></tr>
      	<tr><td>  <% call ShowElite(10,16) %>
        
        </td></tr>
		
		</table>
		
		
		
		<!-- end left column--></td>
		
		<td valign="top" width="758"><!--right column--> <table width="758" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td height="5"></td>
        </tr>
      </table>
      <table width="98%" border="0" align="center" cellpadding="0" cellspacing="0" class="title_main">
        <tr> 
          <td width="40">&nbsp;</td>
          <td><table width="100%" border="0" cellspacing="0" cellpadding="0" >
              <tr> 
                <td class="title_maintxt"><font face="黑体" color="#FF3300"><strong>[课程]<%=SpecialName%>文章列表</strong></font></td>
              </tr>
            </table></td>
        </tr>
      </table>
      <table width="98%" border="0" align="center" cellpadding="0" cellspacing="0" class="border">
        <tr> 
          <td valign="top"> <table width="100%" height="100%" border="0" cellpadding="0" cellspacing="5">
              <tr> 
                <td height="150" valign="top" align="left"> <%call ShowArticle(30)%> </td>
              </tr>
            </table></td>
        </tr>
        <tr class="tdbg"> 
          <td> <table width="99%" border="0" align="center" cellpadding="0" cellspacing="0">
              <tr class="tdbg_leftall"> 
                <td> <%
		  if totalput>0 then
		  	call showpage(strFileName,totalput,MaxPerPage,false,true,"篇文章")
		  end if
		  %> </td>
              </tr>
            </table></td>
        </tr>
      </table>
      <table width="98%" border="0" align="center" cellpadding="0" cellspacing="0">
        <tr> 
          <td height="13" Class="tdbg_left2"></td>
        </tr>
      </table>
      <table width='98%' border='0' align="center"cellpadding='2' cellspacing='0' class="tdbg_rightall">
        <tr class='tdbg_leftall'> 
          <td width="22%"> <div align="center"><img src="Images/checkarticle.gif" width="15" height="15" align="absmiddle">&nbsp;&nbsp;站内文章搜索：</div></td>
          <td width="78%"> <div align="center"> 
              <% call ShowSearchForm("Article_Search.asp",2) %>
            </div></td>
        </tr>
      </table></td>
		</tr><!--end main bocy-->
		
		
		</table><!--end main bocy-->
	
	</div>
    </td></tr>
    <tr><td><!--页面底部--> <%  call Bottom_All()  %><!--结束页面底部-->
	</td></tr>
	<% call PopAnnouceWindow(400,300) %>
    </table><!--the great talbe-->
</body>
</html>
<% call CloseConn() %>