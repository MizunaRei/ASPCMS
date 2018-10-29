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
<HTML  xmlns="http://www.w3.org/1999/xhtml">
<HEAD>
<TITLE><%=strPageTitle%></TITLE>
<META http-equiv=Content-Type content="text/html; charset=gb2312">
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
</HEAD>
<BODY   style="BACKGROUND-IMAGE: url(SkinIndex/bg_all.gif)" bottomMargin=0 
leftMargin=0 background="SkinIndex/bg_all.gif" topMargin=0 rightMargin=0 
marginheight="0" marginwidth="0" <%=Body_Label%> onmousemove='HideMenu()' >
<DIV align=center>
  <!--标题栏-->
  <!--#include file="Top.asp"-->
</DIV>


<DIV align=center>
  <TABLE id=table2 height=300 cellSpacing=0 cellPadding=0 width=797 
background=SkinIndex/new_wow_43.jpg border=0>
    <TBODY>
    
    <TR>
      <TD width=8 rowSpan=7>　</TD>
      <TD width=781 bgColor=#ffffff colSpan=2></TD>
      <TD width=8 rowSpan=7>　</TD>
    </TR>
    <TR>
      <TD  width=741 background=SkinIndex/njyyjy_14.gif bgColor=#ffffff 
    height=5></TD>
    </TR>
    <TR>
      <TD width=781 bgColor=#ffffff colSpan=2 height=0></TD>
    </TR>
    <TR>
      <TD width=781 bgColor=#ffffff colSpan=2><TABLE id=table3 height=88 cellSpacing=0 cellPadding=0 width="100%" 
      border=0>
          <TBODY>
          
          <TR>
          
          <TD vAlign=top   align="left" width=197><DIV align=center>
           <!--用户登录-->   <TABLE id=table4 cellSpacing=0 cellPadding=0 width=197 border=0>
                <TBODY>
                  <TR>
                    <TD align=middle><img height=45 
                  src="SkinIndex/zdl_6.gif" width=197 border=0 >　</TD>
                  </TR>
                  <TR>
                    <TD><IMG height=24 src="SkinIndex/zin_r11_c2.gif" 
                  width=197 border=0></TD>
                  </TR>
                  <TR>
                    <!---用户登录代码---->
                    <TD align=center  valign="top" background=SkinIndex/zin_r13_c2.gif 
                height=100><% call ShowUserLogin() %>
                      </TD>
                  </TR>
                  <TR>
                    <TD><IMG height=80 src="SkinIndex/zin_r16_c1.gif" 
                  width=197 border=0></TD>
                  </TR>
                </TBODY>
              </TABLE><!--结束用户登录-->
            </DIV>
          <div align="center">
		  <!--人气文章--><TABLE id=table5 height=200 cellSpacing=0 cellPadding=0 width=197 
border=0>
              <TBODY>
                <TR>
                  <TD background=SkinIndex/zdl_8.gif height=35><P align=center><B>人 气 文 章</B></P></TD>
                </TR>
                <TR>
                  <TD height=24><IMG height=24 src="SkinIndex/zin_r11_c1.gif" 
            width=197 border=0></TD>
                </TR>
                
                
                <TR>
                  <TD  align="left" background=SkinIndex/zin_r13_c1.gif>
                   <% call Showhot(8,16) %>
                    </TD>
                </TR>
                
                <TR>
                  <TD height=24><IMG height=25 src="SkinIndex/zin_r18_c1.gif" 
            width=197 border=0></TD>
                </TR>
              </TBODY>
            </TABLE><!--结束人气文章-->
		  </div>
            
          <!--用户排行榜-->  <TABLE id=table5 height=200 cellSpacing=0 cellPadding=0 width=197 
            border=0>
                <TBODY>
                  <TR>
                    <TD background=SkinIndex/zdl_8.gif height=35><P align=center><a href="userlist.asp" target="_blank"><B>文 坛 争 峰</B></a></P></TD>
                  </TR>
                  <TR>
                    <TD height=24><IMG height=24 
                  src="SkinIndex/zin_r11_c1.gif" width=197 border=0></TD>
                  </TR>
                  <TR>
                    <TD align=center valign="top" background=SkinIndex/zin_r13_c1.gif><SPAN class=s1>
                        <% call ShowTopUser(10) %>
                        </SPAN>
                      </TD>
                  </TR>
                  <TR>
                    <TD height=24><IMG height=25 
                  src="SkinIndex/zin_r18_c1.gif" width=197 
              border=0></TD>
                  </TR>
                </TBODY>
              </TABLE><!--结束用户排行榜-->
              <!--课程列表--><TABLE id=table5 height=100 cellSpacing=0 cellPadding=0 width=191 border=0>
              <TBODY>
                <TR>
                  <TD background=SkinIndex/zdl_8.gif height=35><P align=center><B>课 程 列 表</B></P></TD>
                </TR>
                <TR>
                  <TD height=24><IMG height=24 src="SkinIndex/zin_r11_c1.gif" 
            width=197 border=0></TD>
                </TR>
                
                
                <TR>
                  <TD  align="center" background=SkinIndex/zin_r13_c1.gif>
                   <% call ShowSpecial(10) %>
                    </TD>
                </TR>
                
                <TR>
                  <TD height=24><IMG height=25 src="SkinIndex/zin_r18_c1.gif" 
            width=197 border=0></TD>
                </TR>
              </TBODY>
            </TABLE><!--结束课程列表-->
              
            </TD><!--结束左栏-->
          <!-- 用VBSCRIPT列出栏目及其文章标题,打开数据集 -->
          <TD  valign="top"  align="center" width=387><!--中间栏-->
              <%
          	sqlShowClassArticleListName="select ClassID,ClassName from ArticleClass Order by ClassID"
	Set rsShowClassArticleListName=Server.CreateObject("Adodb.RecordSet")
	rsShowClassArticleListName.Open sqlShowClassArticleListName,conn,1,1

	If rsShowClassArticleListName.BOF and rsShowClassArticleListName.EOF then
	
	Response.Write("没有任何可列出的栏目。请先添加栏目。")
	
	Else
	
			do while not rsShowClassArticleListName.EOF
			'开始画表格
			Response.Write("<TABLE id=table6 cellSpacing=0 cellPadding=0 width=387 border=0> <TBODY> <TR>  <TD background=SkinIndex/zin_r46_c21.jpg  height=48><B>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;")
	
			Response.Write("<a href='Article_Class2.asp?ClassID="  & rsShowClassArticleListName("ClassID")  & "'>"  )
			Response.Write(rsShowClassArticleListName("ClassName"))
			Response.Write("</a>")
			
			
			Response.Write("</B></TD> </TR> <TR>  <TD background=SkinIndex/zin_r42_c21_r3.jpg height=129><DIV align=center>  <TABLE id=table12 height=27 cellSpacing=0 cellPadding=0  width='95%' border=0>  <TBODY>  <TR>") 
			Response.Write(" <!--理论动态栏目内的文章列表-->   <TD><TABLE cellSpacing=0 cellPadding=0 width='100%'>  <TBODY>  <TR>")
			'结束画表格，开始列出文章
			'使用原网站的javascript列出文章
			
             ' <!--代码开始-->
            Response.Write("  <script language='JavaScript' type='text/JavaScript' src='article_js.asp?ClassID=" &  rsShowClassArticleListName("ClassID")  )
			 Response.Write("&IncludeChild=true&SpecialID=&ArticleNum=6&ShowType=1&ShowCols=1&ShowProperty=true&ShowClassName=false&ShowIncludePic=false&ShowTitle=true&ShowUpdateTime=false&ShowHits=false&ShowAuthor=true&ShowHot=false&ShowMore=false&TitleMaxLen=30&ContentMaxLen=200&Hot=false&Elite=false&DateNum=&OrderField=UpdateTime&OrderType=desc'>   </script>  ")
            '  <!--代码结束-->
              
			'结束列表，画出下半部分表格
			Response.Write("</TR> <TR></TR> </TBODY>  </TABLE></TD> </TR> </TBODY> </TABLE> </DIV></TD></TR>")
            Response.Write("<TR><TD><IMG height=20 src='SkinIndex/zin_r41_c21.jpg' width=387 border=0></TD> </TR> </TBODY> </TABLE>")
			'结束画表格
			
			rsShowClassArticleListName.movenext
			loop
	End If	

          
          %>
              
              <%
              rsShowClassArticleListName.close
				set  rsShowClassArticleListName=nothing
              %>
              <!-- 结束列出栏目及其文章标题代码,关闭数据集-->
              
              </TD><!--结束文章列表-->
            <TD width=197  align="right" valign="top"><!--右边栏-->
            <!--网站公告--><TABLE id=table5 height=200 cellSpacing=0 cellPadding=0 width=197 
            border=0>
                <TBODY>
                  <TR>
                    <TD background=SkinIndex/zdl_8.gif height=35><P align=center><B>本 网 公 告</B></P></TD>
                  </TR>
                  <TR>
                    <TD height=24><IMG height=24 
                  src="SkinIndex/zin_r11_c1.gif" width=197 border=0></TD>
                  </TR>
                  <TR>
                    <TD align=center background=SkinIndex/zin_r13_c1.gif><%call ShowAnnounce(1,1)%></TD>
                  </TR>
                  <TR>
                    <TD height=24><IMG height=25 
                  src="SkinIndex/zin_r18_c1.gif" width=197 
              border=0></TD>
                  </TR>
                </TBODY>
              </TABLE>
            <!--结束网站公告-->
              
         <!--留言板-->  <TABLE id=table5 height=200 cellSpacing=0 cellPadding=0 width=197 
            border=0>
              <TBODY>
                <TR>
                
                
                  <TD align=center  valign="middle" background=SkinIndex/zdl_8.gif 
                  height=35  width="197"><P align=center><B>留 言 板</B></p></TD>
                  </TR>
                  
                  <TR>
                    <TD height=24><IMG height=24 
                  src="SkinIndex/zin_r11_c1.gif" width=197 border=0></TD>
                  </TR>
                  <TR>
                    <TD align="left" background=SkinIndex/zin_r13_c1.gif ><% call showGuest(20,10) %></TD>
                  </TR>
                  <TR>
                    <TD height=24><IMG height=24 
                  src="SkinIndex/zin_r18_c1.gif" width=197 
                border=0></TD>
                
              </TR>
              
              </TBODY>
              
            </TABLE><!--结束留言板-->
             <!--死狗搜索--> <TABLE id=table5 height=100 cellSpacing=0 cellPadding=0 width=197 
            border=0>
                <TBODY>
                  <TR>
                    <TD background=SkinIndex/zdl_8.gif height=35><P align=center><B>Google 搜 索 </B></P></TD>
                  </TR>
                  <TR>
                    <TD height=24><IMG height=24 
                  src="SkinIndex/zin_r11_c1.gif" width=197 border=0></TD>
                  </TR>
                  <TR>
                    <TD background=SkinIndex/zin_r13_c1.gif><!-- Search Google -->
                      <CENTER>
                        <FORM action=http://www.google.cn/custom method=get 
                  target=google_window>
                          <TABLE bgColor=#fcfcf2>
                            <TBODY>
                              <TR>
                                <TD vAlign=top noWrap align=left height=32><A 
                        href="http://www.google.com/"><IMG alt=Google 
                        src="SkinIndex/Logo_25wht.gif" align=middle 
                        border=0></IMG></A> <BR>
                                  <LABEL style="DISPLAY: none" 
                        for=sbi>输入您的搜索字词</LABEL>
                                  <INPUT id=sbi maxLength=255 
                        size=25 name=q>
                                  </INPUT>
                                </TD>
                              </TR>
                              <TR>
                                <TD vAlign=top align=left><P align=center>
                                    <LABEL style="DISPLAY: none" 
                        for=sbb>提交搜索表单</LABEL>
                                    <INPUT id=sbb type=submit value=搜索 name=sa>
                                    </INPUT>
                                    <INPUT type=hidden value=pub-2209561258995469 
                        name=client>
                                    </INPUT>
                                    <INPUT type=hidden value=1 
                        name=forid>
                                    </INPUT>
                                    <INPUT type=hidden value=GB2312 
                        name=ie>
                                    </INPUT>
                                    <INPUT type=hidden value=GB2312 
                        name=oe>
                                    </INPUT>
                                    <INPUT type=hidden 
                        value=GALT:#008000;GL:1;DIV:#336699;VLC:663399;AH:center;BGC:FFFFFF;LBGC:336699;ALC:0000FF;LC:0000FF;T:000000;GFNT:0000FF;GIMP:0000FF;FORID:1 
                        name=cof>
                                    </INPUT>
                                    <INPUT type=hidden value=zh-CN 
                        name=hl>
                                    </INPUT>
                                  </P></TD>
                              </TR>
                            </TBODY>
                          </TABLE>
                        </FORM>
                      </CENTER>
                      <!-- Search Google --></TD>
                  </TR>
                  <TR>
                    <TD height=24><IMG height=25 
                  src="SkinIndex/zin_r18_c1.gif" width=197 
              border=0></TD>
                  </TR>
                </TBODY>
              </TABLE><!--结束死狗搜索-->
              <!--欢迎新生-->   <TABLE id=table5 height=200 cellSpacing=0 cellPadding=0 width=197 
border=0>
              <TBODY>
                <TR>
                  <TD background=SkinIndex/zdl_8.gif height=35><P align=center><B>欢 迎 加 入</B></P></TD>
                </TR>
                <TR>
                  <TD height=24><IMG height=24 src="SkinIndex/zin_r11_c1.gif" 
            width=197 border=0></TD>
                </TR>
                
                
                <TR>
                  <TD   align="center" background=SkinIndex/zin_r13_c1.gif>
				  <% call ShownewUser(5) %>
                  </TD>
                </TR>
                
                
                <TR>
                  <TD height=24><IMG height=25 src="SkinIndex/zin_r18_c1.gif" 
            width=197 border=0></TD>
                </TR>
              </TBODY>
            </TABLE><!--结束欢迎新生-->
			  </TD>
          </TR></TBODY>
          
        </TABLE></TD>
    </TR>
    <TR>
      <TD width=781 bgColor=#ffffff colSpan=2 height=1></TD>
    </TR>
    </TBODY>
    
  </TABLE>
  <!--分割网页-->
  
  
  
  </DIV>
  
  
 <!--页面底部--> <div align="center">
  
    <TABLE id=table7 height=70 cellSpacing=0 cellPadding=0 width=797 
background=SkinIndex/new_wow_43.jpg border=0>
      <TBODY>
        <TR>
          <TD width=8 rowSpan=2>　</TD>
          <TD width=781 bgColor=#c7b883 height=20><P align=center><B>| <SPAN style="CURSOR: hand" 
      onclick="var strHref=window.location.href;this.style.behavior='url(#default#homepage)';this.setHomePage('http://renwen.university.edu.cn');">设为首页</SPAN> | <SPAN title='两课教学网' style="CURSOR: hand" 
      onclick="window.external.addFavorite('http://renwen.university.edu.cn','两课教学网')">收藏本站</SPAN> | <A class=Bottom href="mailto:86277298@QQ.COM">联系站长</A> | <A class=Bottom 
      href="http://renwen.university.edu.cn/FriendSite/Index.asp" target=_blank>友情链接</A> | <A class=Bottom href="http://renwen.university.edu.cn/Copyright.asp" 
      target=_blank>版权申明</A> | </B></P></TD>
          <TD width=8 rowSpan=2>　</TD>
        </TR>
        <TR>
          <TD align=center width=781  valign="middle"
      bgColor=#ffffff>本网站由<font color="#3300FF"><a href="http://renwen.university.edu.cn/">university人文社会科学学院</a></font>主办、维护<BR>
            
             </TD>
        </TR>
      </TBODY>
    </TABLE>
  </div><!--结束页面底部-->
 
</BODY>
</HTML>
