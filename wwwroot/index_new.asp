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
<body bgcolor="#FFFFFF"  style="BACKGROUND-COLOR: #ffffff"  leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<!-- ImageReady Slices (首页_slice2.psd) -->
<div align="center" ><table id="__01" width="989"  border="0" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
	<!--<tr>
		<td colspan="14">
			<img src="images/首页_slice2_01.jpg" width="1024" height="142" alt=""></td>
		<td>
			<img src="images/分隔符.gif" width="1" height="142" alt=""></td>
	</tr>-->
	<tr>
		<!--<td rowspan="17">
			<img src="images/首页_slice2_02.jpg" width="17" height="858" alt=""></td>-->
		<td colspan="12">
			<img   src="images/首页_slice2_03.jpg" width="989" height="188" alt=""></td>
		<!--<td rowspan="17">
			<img src="images/首页_slice2_04.jpg" width="18" height="858" alt=""></td>-->
		<!--<td>
			<img src="images/分隔符.gif" width="1" height="188" alt=""></td>-->
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
    </table><!--top--></div>
   
   
   
   <div align="center"><!--main body--><table  bgcolor="#FFFFFF" width="989">
    <tr><!--main body-->
    <td  align="left" valign="top"><!--left column user login-->
    	<table>
    	<tr><!--user login-->
        
        	<td>
            	<fieldset><legend><font  size="+1">用户登录</font></legend><% call ShowUserLogin() %></fieldset>
                </td>
        </tr>
        
        <tr><!--lessons-->
        	<td>
            	<table width="215">
			<tr><td background="images/首页_slice2_15_top.jpg" height="57" width="150">
            <!--<img src="images/首页_slice2_15.jpg" width="210" height="127" alt="">-->&nbsp;<strong><font size="+1">课程列表</font></strong></td></tr>
            <tr><td background="images/首页_slice2_15_middle.jpg">   <% call ShowSpecial(10) %>
            </td></tr>
            <!--<tr><td background="images/首页_slice2_15_bottom2.jpg"  width="150">&nbsp; </td></tr>-->
            </table>
            </td>
        </tr>
        
        <tr><!--board--><td>
        <table width="215">
        <tr ><td background="Images/首页_slice2_29_top.jpg" width="215"  height="44">&nbsp;<font size="+1">留言板</font>
        
        
        </td></tr>
        <tr><td background="Images/首页_slice2_29_middle.jpg" width="215">
        <% call showGuest(20,10) %>
        </td></tr>
			<tr><td  background="Images/首页_slice2_29_bottom.jpg" width="215" height="48">
        
        </td></tr>
        
        
        </table>
        
        </td></tr>
        
    	</table>
    </td>
    <td valign="top"  align="center"><!--right column articles-->
    	<table>
    		<tr><td valign="top" align="center">
            	<table width="529"><!--人气文章-->
                
                <tr>  
    				<td colspan="5" rowspan="2" width="529" height="40" background="Images/首页_slice2_08_2.jpg">
			<!--<img src="images/首页_slice2_08.jpg" width="529" height="24" alt="">-->
           <font size="+1"><strong><em> 人气文章</em></strong></font>
            
            </td>
   			  </tr>
              <tr><td>
              <% call Showhot(8,16) %>
              </td></tr></table>
              </td>
             <td  valign="top"  align="center">
             	<table> <!--公告-->
       			   	  <tr><td background="Images/首页_slice2_111.jpg"   width="210"  height="186">
       		 	  	   <%call ShowAnnounce(1,1)%>
     		   	 	    </td></tr>
             
             	</table></td>
              
              </tr>
    
    		<tr><!--四栏目-->
            <td  colspan="10">  
            	<table>
					<tr><!--首行-->
                    
                    
                    
                    
                    <td width="356" height="14" background="images/首页_slice2_200.jpg">
        <a href='Article_Class2.asp?ClassID=2'> &nbsp;&nbsp;资料中心 &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;更多>></a>
			<!--<img src="images/首页_slice2_22.jpg" width="356" height="14" alt="">--></td>
            
            <td colspan="3" width="356" height="14" background="images/首页_slice2_200.jpg">
		<a href='Article_Class2.asp?ClassID=3'>	&nbsp;&nbsp;时事新闻&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;更多>></a>
            <!--<img src="images/首页_slice2_24.jpg" width="356" height="14" alt="">--></td>
                    </tr>    
                    
                      <!--二行-->     
                      <tr>
                      <td><%  call ShowArticle_Index(5,2,-1,10) %></td>
                      <td><%  call ShowArticle_Index(5,3,-1,10) %></td>
              		        </tr> 
            		<!--三行-->
                    
                    <tr>
                    <td width="356" height="14" background="images/首页_slice2_200.jpg">
			<!--<img src="images/首页_slice2_30.jpg" width="356" height="13" alt="">-->
             
        <a href='Article_Class2.asp?ClassID=1'>   &nbsp; 理论动态&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;更多>></a></td>
		<td colspan="3" width="356" height="14" background="images/首页_slice2_200.jpg">
			<!--<img src="images/首页_slice2_31.jpg" width="356" height="13" alt="">-->
            <a href='Article_Class2.asp?ClassID=58'>  &nbsp;  学生作品&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;更多>></a>
            </td>
                    
                    </tr>
                    <tr><!--四行-->
                  <td>  <%  call ShowArticle_Index(5,1,-1,10) %></td>
                  <td><%  call ShowArticle_Index(5,58,-1,10) %></td>
                    </tr>
            	</table>
            
            </td>
            </tr>
    
   		  </table>
    
    </td>
    </tr>
	
	
	
	
	
	
	
	
	
	</table><!--main body--></div>
<%  call Bottom_All()  %>
<!-- End ImageReady Slices -->
</body>
</html>