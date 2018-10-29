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
<HEAD>
<TITLE><%=strPageTitle%></TITLE>
<META http-equiv=Content-Type content="text/html; charset=gb2312">
<META 
content="<%=strPageTitle%>:资源免费，更新快，资源全，提供本科、硕士研究生的各种思想政治课和马克思主义理论课教学资源，栏目有：理论动态、资料中心、时事新闻、学生作品。" 
name=Keywords>
<META 
content="<%=strPageTitle%>:资源免费，更新快，资源全，提供本科、硕士研究生的各种思想政治课和马克思主义理论课教学资源，栏目有：理论动态、资料中心、时事新闻、学生作品。" 
name=Description>
</HEAD>
<body><table>


	<!--body的中对齐--><div align="center">
    
    <!--	第一行，横幅--><tr>
		
		<!--第一行，第一格，横幅--><td>
			
			<img src="ImagesNew/TopBanner.jpg">

		</td><!--第一行，第一格，横幅-->
	</tr><!--	第一行，横幅-->
    <!--<br/>-->
 <!--<div align="left">-->
 <!--   第二行菜单和搜索栏-->
 
 <tr >
   <td align="left" background="ImagesNew/IconSpaceHolder.bmp"> <a href="index.asp"> <img src="ImagesNew/HomePageIcon.jpg"></a>  </td>
   <td align="left" background="ImagesNew/IconSpaceHolder.bmp"> <a href="Article_Class2.asp?ClassID=2"> <img src="ImagesNew/LibraryIcon.jpg"></a>  </td>
    <td align="left"> <a href="Article_Class2.asp?ClassID=1"> <img src="ImagesNew/TheoryIcon.jpg"></a>  </td>
    <td align="left"> <a href="Article_Class2.asp?ClassID=3"> <img src="ImagesNew/News.jpg"></a>  </td>
    <td align="left"> <a href="Article_Class2.asp?ClassID=58"> <img src="ImagesNew/StudentArticleIcons.jpg"></a>  </td>
    
     <td align="left"> <a href="userlist.asp"> <img src="ImagesNew/PersonalArticles.jpg"></a>
    <td align="left"> <a href="guestbook.asp"> <img src="ImagesNew/GuestBookIcon.jpg"></a>
    <td  width="40" background="ImagesNew/IconSpaceHolder.bmp"></td>
    <td  background="ImagesNew/IconSpaceHolder.bmp"><% call ShowSearchForm("Article_Search.asp",1) %> </td>
    <td  width="20" background="ImagesNew/IconSpaceHolder.bmp"></td>
  </tr>  <!--   第二行菜单和搜索栏-->
  <!--</div>-->
  <!--第三行--><tr >
  	<!--第三行第一列 大td-->
    <td  valign="top" align="left">
    <table ><!--第三行第一列 大table-->
  				<td  valign="top" align="left">
  					<fieldset><!--用户登录开始-->
            			<legend>用户登录</legend>
            				<% call ShowUserLogin() %>
            
            		</fieldset><!--用户登录结束-->
  
  						</td>
                        
         <!--第三行第一列 大table--> </table></td><!--第三行第一列 大td-->              
  </tr><!--第三行-->
   
   
   
   
    </div><!--body的中对齐-->
</table></body>
</html>