<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">

<!--<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
-->
<!--<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">-->
<!--#include file="inc/syscode_article.asp"-->
<!--
Design by Free CSS Templates
http://renwen.university.edu.cn
Released for free under a Creative Commons Attribution 2.5 License
-->
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="content-type" content="text/html; charset=utf-8" />
<title>university两课教学网</title>
<meta name="keywords" content="" />
<META 
content="<%=strPageTitle%>:资源免费，更新快，资源全，提供本科、硕士研究生的各种思想政治课和马克思主义理论课教学资源，栏目有：理论动态、资料中心、时事新闻、学生作品。" 
name=Keywords />
<META 
content="<%=strPageTitle%>:资源免费，更新快，资源全，提供本科、硕士研究生的各种思想政治课和马克思主义理论课教学资源，栏目有：理论动态、资料中心、时事新闻、学生作品。" 
name=Description />
<meta name="description" content="" />
<link href="default_inflight.css" rel="stylesheet" type="text/css" />
</head>
<body >
<div id="logo">
	<h1><a href="#">university<span>两课教学网</span></a></h1>
	<h2><a href="http://renwen.university.edu.cn/">人文学院 主办维护</a></h2>
</div>
<div id="splash"><a href="#"><img src="images_inflight/forest.jpg" alt="" width="600" height="280" /></a></div>
<div id="menu">
	<ul>
		<%call ShowPath()%>
        <li class="first"><a href="#" title="">About Us</a></li>
		<li><a href="#" title="">Products</a></li>
		<li><a href="#" title="">Services</a></li>
		<li><a href="#" title="">Clients</a></li>
		<li><a href="#" title="">Support</a></li>
	</ul>
</div>
<div id="content">
	<div id="main">
		<div id="welcome" class="boxed">
			<h2 class="title">欢迎登陆两课教学网</h2>
			<div class="content">
				<p><strong>两课教学网</strong> 是一个由 <a href="http://renwen.university.edu.cn/">university人文社会科学学院</a> 主办并维护的……网站简介略  <em>..</em></p>
				<!--<p>An unordered list example:</p>
				<ul>
					<li>List item number one</li>
					<li>List item number two</li>
					<li>List item number three </li>
				</ul>-->
			</div>
		</div>
		<div id="example" class="boxed">
			<h2 class="title">栏目</h2>
			<div class="content">
				<p><h2>理论动态</h2></p><p><blockquote>理论动态简介及其他内容</blockquote></p>
				
					<p><% call ShowArticle_Index(10,1,-1,10)
					'第二个参数是CLASSID
					 %></p>
				
				<!--<h3>Heading Level Three</h3>
				<p>This is another example of a paragraph followed by an unordered list. In posuere  eleifend odio. Quisque semper augue mattis wisi. Maecenas ligula.  Pellentesque viverra vulputate enim. Aliquam erat volutpat lorem ipsum  dolorem.</p>
				<p>An ordered list example:</p>
				<ol>
					<li>List item number one</li>
					<li>List item number two</li>
					<li>List item number thre</li>
				</ol>-->
			</div>
            
            
            <div class="content">
				<p><h2>资料中心</h2></p><p><blockquote>资料中心简介及其他内容</blockquote></p>
				
					<p><% call ShowArticle_Index(10,2,-1,10)
					'第二个参数是CLASSID
					 %></p>
				
				<!--<h3>Heading Level Three</h3>
				<p>This is another example of a paragraph followed by an unordered list. In posuere  eleifend odio. Quisque semper augue mattis wisi. Maecenas ligula.  Pellentesque viverra vulputate enim. Aliquam erat volutpat lorem ipsum  dolorem.</p>
				<p>An ordered list example:</p>
				<ol>
					<li>List item number one</li>
					<li>List item number two</li>
					<li>List item number thre</li>
				</ol>-->
			</div>
            
            <div class="content">
				<p><h2>时事新闻</h2></p><p><blockquote>时事新闻简介及其他内容</blockquote></p>
				
					<p><% call ShowArticle_Index(10,3,-1,10)
					'第二个参数是CLASSID
					 %></p>
				
				<!--<h3>Heading Level Three</h3>
				<p>This is another example of a paragraph followed by an unordered list. In posuere  eleifend odio. Quisque semper augue mattis wisi. Maecenas ligula.  Pellentesque viverra vulputate enim. Aliquam erat volutpat lorem ipsum  dolorem.</p>
				<p>An ordered list example:</p>
				<ol>
					<li>List item number one</li>
					<li>List item number two</li>
					<li>List item number thre</li>
				</ol>-->
			</div>
            
            
            <div class="content">
				<p><h2>学生作品</h2></p><p><blockquote>学生作品简介及其他内容</blockquote></p>
				
					<p><% call ShowArticle_Index(10,58,-1,10)
					'第二个参数是CLASSID
					 %></p>
				
				<!--<h3>Heading Level Three</h3>
				<p>This is another example of a paragraph followed by an unordered list. In posuere  eleifend odio. Quisque semper augue mattis wisi. Maecenas ligula.  Pellentesque viverra vulputate enim. Aliquam erat volutpat lorem ipsum  dolorem.</p>
				<p>An ordered list example:</p>
				<ol>
					<li>List item number one</li>
					<li>List item number two</li>
					<li>List item number thre</li>
				</ol>-->
			</div>
            
            
            
            
            
            
            
            
		</div>
	</div>
	<div id="sidebar">
		<div id="login" class="boxed">
			<h2 class="title">用户登录</h2>
			<div class="content">
				<!--<form id="form1" method="post" action="#">-->
					<fieldset>
					<legend>用户登录</legend>
                    <%
					 call ShowUserLogin() 
					%>
					<!--<label for="inputtext1">Client ID:</label>
					<input id="inputtext1" type="text" name="inputtext1" value="" />
					<label for="inputtext2">Password:</label>
					<input id="inputtext2" type="password" name="inputtext2" value="" />
					<input id="inputsubmit1" type="submit" name="inputsubmit1" value="Sign In" />
					<p><a href="#">Forgot your password?</a></p>-->
					</fieldset>
				<!--</form>-->
			</div>
		</div>
        <div id="partners" class="boxed">
			<h2 class="title">公告</h2>
			<div class="content">
				<ul>
					<li><a href="#"><%
					'response.Write(ChannelID)
					'response.End()
					call ShowAnnounce_Index(1,1,0)
					%></a></li>
                    <!--<li><a href="#">Donec Dictum Metus</a></li>
                    
					<li><a href="#">Etiam Rhoncus Volutpat</a></li>
					<li><a href="#">Integer Gravida Nibh</a></li>
					<li><a href="#">Maecenas Luctus Lectus</a></li>
					<li><a href="#">Mauris Vulputate Dolor</a></li>
					<li><a href="#">Nulla Luctus Eleifend</a></li>
					<li><a href="#">Posuere Augue Sit Nisl</a></li>-->
				</ul>
			</div>
		</div>
		<div id="updates" class="boxed">
			<h2 class="title">课程列表</h2>
			<div class="content">
				<ul>
					<li>
						<h3><strong>马克思主义基本原理</strong></h3>
						<p><a href="#">课程简介或其他&#8230;</a></p>
					</li>
					<li>
						<h3>毛泽东思想、邓小平理论和“三个代表”重要思想</h3>
						<p><a href="#">课程简介或其他&#8230;</a></p>
					</li>
					<li>
						<h3>中国近现代史纲要</h3>
						<p><a href="#">课程简介或其他&#8230;</a></p>
					</li>
					<li>
						<h3>思想道德修养与法律基础</h3>
						<p><a href="#">课程简介或其他&#8230;</a></p>
					</li>
					<!--<li>
						<h3>February 20, 2007</h3>
						<p><a href="#">Vivamus fermentum nibh in augue. Praesent a lacus at urna congue rutrum. Nulla enim eros&#8230;</a></p>
					</li>-->
				</ul>
			</div>
		</div>
		
	</div>
</div>
<div id="footer" >
	<p id="legal">Copyleft &copy; 2008  <a href="http://renwen.university.edu.cn/">university人文社会科学学院</a>.</p>
	<p id="links"><a href="#">联系站长</a> | <a href="#">引用本站资源请注明出处</a></p>
</div>
</body>
</html>
