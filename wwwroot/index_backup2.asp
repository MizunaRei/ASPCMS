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
leftMargin=0 background=SkinIndex/bg_all.gif topMargin=0 rightMargin=0 
marginheight="0" marginwidth="0" <%=Body_Label%> onmousemove='HideMenu()' >
<DIV align=center>
  <!--标题栏-->
  <!--#include file="Top.asp"-->
</DIV>
<DIV align=center>
  <TABLE id=table7 height=80 cellSpacing=0 cellPadding=0 width=797 
background=SkinIndex/new_wow_43.jpg border=0>
    <TBODY>
      <TR>
        <TD width=8>　</TD>
        <TD align=middle width=781 bgColor=#ffffff height=80><SCRIPT language=javascript src="SkinIndex/1.js"></SCRIPT>
        </TD>
        <TD width=8>　</TD>
      </TR>
    </TBODY>
  </TABLE>
</DIV>
<DIV align=center>
  <TABLE id=table7 cellSpacing=0 cellPadding=0 width=797 
background=SkinIndex/new_wow_43.jpg border=0>
    <TBODY>
      <TR>
        <TD width=8 rowSpan=3>　</TD>
        <TD width=781 bgColor=#ffffff height=22><IMG height=21 
      src="SkinIndex/zin_r39_c2.gif" width=781 border=0></TD>
        <TD width=8 rowSpan=3></TD>
      </TR>
      <TR>
        <TD align=middle width=781 background=SkinIndex/zin_r42_c12.gif 
    bgColor=#ffffff height=6></TD>
      </TR>
      <TR>
        <TD width=781 bgColor=#ffffff height=22><IMG height=19 
      src="SkinIndex/zin_r41_c2.gif" width=781 
border=0></TD>
      </TR>
    </TBODY>
  </TABLE>
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
      <TD align=right width=40 bgColor=#ffffff height=16><IMG height=31 
      src="SkinIndex/njyyjy_13.gif" width=38 border=0></TD>
      <TD width=741 background=SkinIndex/njyyjy_14.gif bgColor=#ffffff 
    height=16><TABLE id=table1 cellSpacing=0 cellPadding=0 width="100%" border=0>
          <TBODY>
            <TR>
              <TD align=middle>　</TD>
              <TD align=middle><B><A 
            href="http://www.zz6789.com/User/User_Article.asp?ChannelID=1&amp;Action=Add"><FONT 
            color=#0000ff>添加文章到【文章中心】</FONT></A></B></TD>
              <TD align=middle><B><A 
            href="http://www.zz6789.com/User/User_Article.asp?ChannelID=1001&amp;Action=Add"><FONT 
            color=#0000ff>添加文章到【高中课程】</FONT></A></B></TD>
              <TD align=middle><B><A 
            href="http://www.zz6789.com/User/User_Soft.asp?ChannelID=2&amp;Action=Add"><FONT 
            color=#0000ff>添加课件到【课件中心】</FONT></A></B></TD>
              <TD align=middle><B><A 
            href="http://www.zz6789.com/User/User_Photo.asp?ChannelID=3&amp;Action=Add"><FONT 
            color=#0000ff>添加图片到【图片中心】</FONT></A></B></TD>
              <TD align=middle>　</TD>
            </TR>
          </TBODY>
        </TABLE></TD>
    </TR>
    <TR>
      <TD width=781 bgColor=#ffffff colSpan=2 height=0></TD>
    </TR>
    <TR>
      <TD width=781 bgColor=#ffffff colSpan=2><TABLE id=table3 height=88 cellSpacing=0 cellPadding=0 width="100%" 
      border=0>
          <TBODY>
          
          <TR>
          
          <TD vAlign=top width=197><DIV align=center>
              <TABLE id=table4 cellSpacing=0 cellPadding=0 width=197 border=0>
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
                    <!---用户登录---->
                    <TD align=middle background=SkinIndex/zin_r13_c2.gif 
                height=172><% call ShowUserLogin() %>
                      </TD>
                  </TR>
                  <TR>
                    <TD><IMG height=80 src="SkinIndex/zin_r16_c1.gif" 
                  width=197 border=0></TD>
                  </TR>
                </TBODY>
              </TABLE>
            </DIV>
            <TABLE id=table5 height=200 cellSpacing=0 cellPadding=0 width=197 
            border=0>
              <TBODY>
                <TR>
                
                
                  <TD align=middle background=SkinIndex/zdl_8.gif 
                  height=35><Strong>留 言 板</Strong></TD>
                  </TR>
                  
                  <TR>
                    <TD height=24><IMG height=24 
                  src="SkinIndex/zin_r11_c1.gif" width=197 border=0></TD>
                  </TR>
                  <TR>
                    <TD vAlign=top align=middle 
                background=SkinIndex/zin_r13_c1.gif height=50><% call showGuest(20,10) %></TD>
                  </TR>
                  <TR>
                    <TD height=24><IMG height=25 
                  src="SkinIndex/zin_r18_c1.gif" width=197 
                border=0></TD>
                
              </TR>
              
              </TBODY>
              
            </TABLE></TD>
          <TD vAlign=top align=middle width=387><!-- 用VBSCRIPT列出栏目及其文章标题,打开数据集 -->
              <%
          	sqlShowClassArticleListName="select ClassID,ClassName from ArticleClass Order by ClassID"
	Set rsShowClassArticleListName=Server.CreateObject("Adodb.RecordSet")
	rsShowClassArticleListName.Open sqlShowClassArticleListName,conn,1,1

	If rsShowClassArticleListName.BOF and rsShowClassArticleListName.EOF then
	
	Response.Write("没有任何可列出的栏目。请先添加栏目。")
	
	Else
	
			do while not rsShowClassArticleListName.EOF
			'开始画表格
			Response.Write("<TABLE id=table6 cellSpacing=0 cellPadding=0 width=387 border=0> <TBODY> <TR>  <TD background=SkinIndex/zin_r46_c21.jpg  height=48><B>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;")
	
			Response.Write("<a href='Article_Class2.asp?ClassID="  & rsShowClassArticleListName("ClassID")  & "'>"  )
			Response.Write(rsShowClassArticleListName("ClassName"))
			Response.Write("</a>")
			
			
			Response.Write("</B></TD> </TR> <TR>  <TD background=SkinIndex/zin_r42_c21_r3.jpg height=93><DIV align=center>  <TABLE id=table12 height=27 cellSpacing=0 cellPadding=0  width='95%' border=0>  <TBODY>  <TR>") 
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
              <TABLE id=table6 cellSpacing=0 cellPadding=0 width=387 border=0>
                <TBODY>
                  <TR>
                    <TD background=SkinIndex/zin_r46_c21.jpg 
                  height=48><B>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;理论动态</B></TD>
                  </TR>
                  <TR>
                    <TD background=SkinIndex/zin_r42_c21_r3.jpg height=93><DIV align=center>
                        <TABLE id=table12 height=27 cellSpacing=0 cellPadding=0 
                  width="95%" border=0>
                          <TBODY>
                            <TR>
                              <!--理论动态栏目内的文章列表-->
                              <TD><TABLE cellSpacing=0 cellPadding=0 width="100%">
                                  <TBODY>
                                    <TR>
                                      <!--代码开始-->
                                      <script language='JavaScript' type='text/JavaScript' src='article_js.asp?ClassID=1&IncludeChild=true&SpecialID=&ArticleNum=6&ShowType=1&ShowCols=1&ShowProperty=true&ShowClassName=false&ShowIncludePic=false&ShowTitle=true&ShowUpdateTime=false&ShowHits=false&ShowAuthor=true&ShowHot=false&ShowMore=false&TitleMaxLen=30&ContentMaxLen=200&Hot=false&Elite=false&DateNum=&OrderField=UpdateTime&OrderType=desc'></script>
                                      <!--代码结束-->
                                    </TR>
                                    <TR></TR>
                                  </TBODY>
                                </TABLE></TD>
                            </TR>
                          </TBODY>
                        </TABLE>
                      </DIV></TD>
                  </TR>
                  <TR>
                    <TD><IMG height=20 src="SkinIndex/zin_r41_c21.jpg" 
                  width=387 border=0></TD>
                  </TR>
                </TBODY>
              </TABLE>
              <%
              rsShowClassArticleListName.close
				set  rsShowClassArticleListName=nothing
              %>
              <!-- 结束列出栏目及其文章标题代码,关闭数据集-->
              <TABLE id=table6 cellSpacing=0 cellPadding=0 width=387 border=0>
                <TBODY>
                  <TR>
                    <TD background=SkinIndex/zin_r46_c21.jpg 
                  height=48><B>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;推荐文章</B></TD>
                  </TR>
                  <TR>
                    <TD background=SkinIndex/zin_r42_c21_r3.jpg height=93><DIV align=center>
                        <TABLE id=table12 height=27 cellSpacing=0 cellPadding=0 
                  width="95%" border=0>
                          <TBODY>
                            <TR>
                              <!--推荐文章-->
                              <TD><TABLE cellSpacing=0 cellPadding=0 width="100%">
                                  <TBODY>
                                    <TR>
                                      <TD class=listbg>[<A class=listA 
                              href="http://www.zz6789.com/Article/C/C3/Index.html">初三级</A>]<A 
                              class=listA title=中考指导 
                              href="http://www.zz6789.com/Article/C/C3/200808/Article_20080819203053.html" 
                              target=_blank>中考指导</A></TD>
                                    </TR>
                                    <TR>
                                      <TD class=listbg2>[<A class=listA 
                              href="http://www.zz6789.com/Article/D/D5/Index.html">高二级</A>]<A 
                              class=listA title=新企业所得税法的哲学解析 
                              href="http://www.zz6789.com/Article/D/D5/200808/Article_20080816091805.html" 
                              target=_blank>新企业所得税法的哲学解析</A></TD>
                                    </TR>
                                    <TR>
                                      <TD class=listbg>[<A class=listA 
                              href="http://www.zz6789.com/Article/E/E8/Index.html">历界高考题</A>]<A 
                              class=listA title=2008年高考政治常识试题汇总 
                              href="http://www.zz6789.com/Article/E/E8/200808/Article_20080815091349.html" 
                              target=_blank>2008年高考政治常识试题汇总</A></TD>
                                    </TR>
                                    <TR>
                                      <TD class=listbg2>[<A class=listA 
                              href="http://www.zz6789.com/Article/E/E8/Index.html">历界高考题</A>]<A 
                              class=listA title=2008年高考哲学常识试题汇总 
                              href="http://www.zz6789.com/Article/E/E8/200808/Article_20080815091140.html" 
                              target=_blank>2008年高考哲学常识试题汇总</A></TD>
                                    </TR>
                                    <TR>
                                      <TD class=listbg>[<A class=listA 
                              href="http://www.zz6789.com/Article/E/E2/Index.html">高考分析</A>]<A 
                              class=listA title=2008年高考经济常识试题汇总 
                              href="http://www.zz6789.com/Article/E/E2/200808/Article_20080815090848.html" 
                              target=_blank>2008年高考经济常识试题汇总</A></TD>
                                    </TR>
                                    <TR>
                                      <TD class=listbg2>[<A class=listA 
                              href="http://www.zz6789.com/Article/J/J7/Index.html">时事政治</A>]<A 
                              class=listA title=2008年7月1—31日时事（部分转摘） 
                              href="http://www.zz6789.com/Article/J/J7/200808/Article_20080806171600.html" 
                              target=_blank>2008年7月1—31日时事（部分转摘）</A></TD>
                                    </TR>
                                    <TR>
                                      <TD class=listbg>[<A class=listA 
                              href="http://www.zz6789.com/Article/J/J7/Index.html">时事政治</A>]<A 
                              class=listA title=2008年7月时事政治 
                              href="http://www.zz6789.com/Article/J/J7/200808/Article_20080804083705.html" 
                              target=_blank>2008年7月时事政治</A></TD>
                                    </TR>
                                    <TR>
                                      <TD class=listbg2>[<A class=listA 
                              href="http://www.zz6789.com/Article/E/E3/Index.html">备考策略</A>]<A 
                              class=listA title=从08年高考文科综合全国卷看09年政治高考备考 
                              href="http://www.zz6789.com/Article/E/E3/200807/Article_20080725225113.html" 
                              target=_blank>从08年高考文科综合全国卷看09年政治高考备考</A></TD>
                                    </TR>
                                    <TR>
                                      <TD class=listbg>[<A class=listA 
                              href="http://www.zz6789.com/Article/D/D3/Index.html">初三级</A>]<A 
                              class=listA title=滑集中学时政总结 
                              href="http://www.zz6789.com/Article/D/D3/200806/Article_20080627100627.html" 
                              target=_blank>滑集中学时政总结</A></TD>
                                    </TR>
                                    <TR></TR>
                                  </TBODY>
                                </TABLE></TD>
                            </TR>
                          </TBODY>
                        </TABLE>
                      </DIV></TD>
                  </TR>
                  <TR>
                    <TD><IMG height=20 src="SkinIndex/zin_r41_c21.jpg" 
                  width=387 border=0></TD>
                  </TR>
                </TBODY>
              </TABLE>
              <TABLE id=table6 cellSpacing=0 cellPadding=0 width=387 border=0>
                <TBODY>
                  <TR>
                    <TD background=SkinIndex/zin_r46_c21.jpg height=48><P align=left><B>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
                        热点文章</B></P></TD>
                  </TR>
                  <TR>
                    <TD background=SkinIndex/zin_r42_c21_r3.jpg height=93><DIV align=center>
                        <TABLE id=table12 height=27 cellSpacing=0 cellPadding=0 
                  width="95%" border=0>
                          <TBODY>
                            <TR>
                              <!--热点文章-->
                              <TD><TABLE cellSpacing=0 cellPadding=0 width="100%">
                                  <TBODY>
                                    <TR>
                                      <TD class=listbg>[<A class=listA 
                              href="http://www.zz6789.com/Article/k/k3/Index.html">历史真相</A>]<A 
                              class=listA title=苏联人为何“不珍惜”苏联 
                              href="http://www.zz6789.com/Article/k/k3/200809/Article_20080908091703.html" 
                              target=_blank>苏联人为何“不珍惜”苏联</A></TD>
                                    </TR>
                                    <TR>
                                      <TD class=listbg2>[<A class=listA 
                              href="http://www.zz6789.com/Article/J/J7/Index.html">时事政治</A>]<A 
                              class=listA title=2008年8月1—31日时事 
                              href="http://www.zz6789.com/Article/J/J7/200809/Article_20080906143043.html" 
                              target=_blank>2008年8月1—31日时事</A></TD>
                                    </TR>
                                    <TR>
                                      <TD class=listbg>[<A class=listA 
                              href="http://www.zz6789.com/Article/H/H1/Index.html">教学心得</A>]<A 
                              class=listA title=让初中思想品德课教学生活化 
                              href="http://www.zz6789.com/Article/H/H1/200809/Article_20080905142959.html" 
                              target=_blank>让初中思想品德课教学生活化</A></TD>
                                    </TR>
                                    <TR>
                                      <TD class=listbg2>[<A class=listA 
                              href="http://www.zz6789.com/Article/C/C3/Index.html">初三级</A>]<A 
                              class=listA title=江苏省兴洪中学九年级第一单元思想品德测试 
                              href="http://www.zz6789.com/Article/C/C3/200809/Article_20080905081954.html" 
                              target=_blank>江苏省兴洪中学九年级第一单元思想品德测试</A></TD>
                                    </TR>
                                    <TR>
                                      <TD class=listbg>[<A class=listA 
                              href="http://www.zz6789.com/Article/I/I2/Index.html">新课标培训</A>]<A 
                              class=listA title=杜浪口中学教学模式与我的感悟 
                              href="http://www.zz6789.com/Article/I/I2/200808/Article_20080823095759.html" 
                              target=_blank>杜浪口中学教学模式与我的感悟</A></TD>
                                    </TR>
                                    <TR>
                                      <TD class=listbg2>[<A class=listA 
                              href="http://www.zz6789.com/Article/C/C3/Index.html">初三级</A>]<A 
                              class=listA title=中考指导 
                              href="http://www.zz6789.com/Article/C/C3/200808/Article_20080819203053.html" 
                              target=_blank>中考指导</A></TD>
                                    </TR>
                                    <TR>
                                      <TD class=listbg>[<A class=listA 
                              href="http://www.zz6789.com/Article/G/G3/Index.html">德育教育谈</A>]<A 
                              class=listA title=和风细雨暖人心 
                              href="http://www.zz6789.com/Article/G/G3/200808/Article_20080819161621.html" 
                              target=_blank>和风细雨暖人心</A></TD>
                                    </TR>
                                    <TR>
                                      <TD class=listbg2>[<A class=listA 
                              href="http://www.zz6789.com/Article/D/D5/Index.html">高二级</A>]<A 
                              class=listA title=新企业所得税法的哲学解析 
                              href="http://www.zz6789.com/Article/D/D5/200808/Article_20080816091805.html" 
                              target=_blank>新企业所得税法的哲学解析</A></TD>
                                    </TR>
                                    <TR>
                                      <TD class=listbg>[<A class=listA 
                              href="http://www.zz6789.com/Article/E/E8/Index.html">历界高考题</A>]<A 
                              class=listA title=2008年高考政治常识试题汇总 
                              href="http://www.zz6789.com/Article/E/E8/200808/Article_20080815091349.html" 
                              target=_blank>2008年高考政治常识试题汇总</A></TD>
                                    </TR>
                                    <TR></TR>
                                  </TBODY>
                                </TABLE></TD>
                            </TR>
                          </TBODY>
                        </TABLE>
                      </DIV></TD>
                  </TR>
                  <TR>
                    <TD><IMG height=20 src="SkinIndex/zin_r41_c21.jpg" 
                  width=387 border=0></TD>
                  </TR>
                </TBODY>
              </TABLE></TD>
            <TD width=197><TABLE id=table5 height=200 cellSpacing=0 cellPadding=0 width=197 
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
                    <TD align=middle background=SkinIndex/zin_r13_c1.gif><%call ShowAnnounce(1,1)%>
                    </TD>
                  </TR>
                  <TR>
                    <TD height=24><IMG height=25 
                  src="SkinIndex/zin_r18_c1.gif" width=197 
              border=0></TD>
                  </TR>
                </TBODY>
              </TABLE>
              <TABLE id=table5 height=200 cellSpacing=0 cellPadding=0 width=197 
            border=0>
                <TBODY>
                  <TR>
                    <TD background=SkinIndex/zdl_8.gif height=35><P align=center><B>用 户 排 行 榜</B></P></TD>
                  </TR>
                  <TR>
                    <TD height=24><IMG height=24 
                  src="SkinIndex/zin_r11_c1.gif" width=197 border=0></TD>
                  </TR>
                  <TR>
                    <TD align=middle background=SkinIndex/zin_r13_c1.gif><SPAN class=s1>
                        <% call ShowTopUser(5) %>
                        </SPAN>
                      <DIV align=right><A class=LinkTopUser 
                  href="http://www.zz6789.com/ShowUser.asp?Action=List&amp;ChannelID=0">more...</A>&nbsp;&nbsp;</DIV></TD>
                  </TR>
                  <TR>
                    <TD height=24><IMG height=25 
                  src="SkinIndex/zin_r18_c1.gif" width=197 
              border=0></TD>
                  </TR>
                </TBODY>
              </TABLE>
              <TABLE id=table5 height=100 cellSpacing=0 cellPadding=0 width=197 
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
              </TABLE></TD>
          </TR></TBODY>
          
        </TABLE></TD>
    </TR>
    <TR>
      <TD width=781 bgColor=#ffffff colSpan=2 height=1></TD>
    </TR>
    </TBODY>
    
  </TABLE>
  <DIV align=center>
    <TABLE id=table7 cellSpacing=0 cellPadding=0 width=797 
background=SkinIndex/new_wow_43.jpg border=0>
      <TBODY>
        <TR>
          <TD width=8 rowSpan=3>　</TD>
          <TD width=781 bgColor=#ffffff height=22><IMG height=21 
      src="SkinIndex/zin_r39_c2.gif" width=781 border=0></TD>
          <TD width=8 rowSpan=3></TD>
        </TR>
        <TR>
          <TD align=middle width=781 background=SkinIndex/zin_r42_c12.gif 
    bgColor=#ffffff height=6></TD>
        </TR>
        <TR>
          <TD width=781 bgColor=#ffffff height=22><IMG height=19 
      src="SkinIndex/zin_r41_c2.gif" width=781 
border=0></TD>
        </TR>
      </TBODY>
    </TABLE>
  </DIV>
  <DIV align=center>
    <TABLE id=table7 height=100 cellSpacing=0 cellPadding=0 width=797 
background=SkinIndex/new_wow_43.jpg border=0>
      <TBODY>
        <TR>
          <TD width=8 rowSpan=2>　</TD>
          <TD width=391 bgColor=#ffffff><P align=center>
            <TABLE id=table6 cellSpacing=0 cellPadding=0 width=387 border=0>
              <TBODY>
                <TR>
                  <TD background=SkinIndex/zin_r46_c21.jpg height=48><P align=left><B>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
                      时事新闻</B></P></TD>
                </TR>
                <TR>
                  <TD background=SkinIndex/zin_r42_c21_r3.jpg height=93><DIV align=center>
                      <TABLE id=table12 height=27 cellSpacing=0 cellPadding=0 width="95%" 
            border=0>
                        <TBODY>
                          <TR>
                            <!--时事新闻-->
                            <TD><TABLE cellSpacing=0 cellPadding=0 width="100%">
                                <TBODY>
                                  <TR>
                                    <TD class=listbg>[<A class=listA 
                        href="http://www.zz6789.com/Article/J/J7/Index.html">时事政治</A>]<A 
                        class=listA 
                        title="文章标题：2008年8月1—31日时事&#13;&#10;作&nbsp;&nbsp;&nbsp;&nbsp;者：董常祯&#13;&#10;更新时间：2008-9-6 14:30:43" 
                        href="http://www.zz6789.com/Article/J/J7/200809/Article_20080906143043.html" 
                        target=_blank>2008年8月1—31日时事</A></TD>
                                    <TD class=listbg align=right width=40><FONT 
                        color=red>09-06</FONT></TD>
                                  </TR>
                                  <TR>
                                    <TD class=listbg2>[<A class=listA 
                        href="http://www.zz6789.com/Article/J/J7/Index.html">时事政治</A>]<A 
                        class=listA 
                        title="文章标题：2008年7月1—31日时事（部分转摘）&#13;&#10;作&nbsp;&nbsp;&nbsp;&nbsp;者：董常祯&#13;&#10;更新时间：2008-8-6 17:16:00" 
                        href="http://www.zz6789.com/Article/J/J7/200808/Article_20080806171600.html" 
                        target=_blank>2008年7月1—31日时事（部分转摘）</A></TD>
                                    <TD class=listbg2 align=right width=40>08-06</TD>
                                  </TR>
                                  <TR>
                                    <TD class=listbg>[<A class=listA 
                        href="http://www.zz6789.com/Article/J/J7/Index.html">时事政治</A>]<A 
                        class=listA 
                        title="文章标题：2008年7月时事政治&#13;&#10;作&nbsp;&nbsp;&nbsp;&nbsp;者：fanyongxin&#13;&#10;更新时间：2008-8-4 8:37:05" 
                        href="http://www.zz6789.com/Article/J/J7/200808/Article_20080804083705.html" 
                        target=_blank>2008年7月时事政治</A></TD>
                                    <TD class=listbg align=right width=40>08-04</TD>
                                  </TR>
                                  <TR>
                                    <TD class=listbg2>[<A class=listA 
                        href="http://www.zz6789.com/Article/J/J7/Index.html">时事政治</A>]<A 
                        class=listA 
                        title="文章标题：2008年7月1—15日时事（原创）&#13;&#10;作&nbsp;&nbsp;&nbsp;&nbsp;者：董常祯&#13;&#10;更新时间：2008-7-16 21:48:41" 
                        href="http://www.zz6789.com/Article/J/J7/200807/Article_20080716214841.html" 
                        target=_blank>2008年7月1—15日时事（原创）</A></TD>
                                    <TD class=listbg2 align=right width=40>07-16</TD>
                                  </TR>
                                  <TR>
                                    <TD class=listbg>[<A class=listA 
                        href="http://www.zz6789.com/Article/J/J7/Index.html">时事政治</A>]<A 
                        class=listA 
                        title="文章标题：2008年6月16—30日时事（原创）&#13;&#10;作&nbsp;&nbsp;&nbsp;&nbsp;者：董常祯&#13;&#10;更新时间：2008-7-1 21:44:42" 
                        href="http://www.zz6789.com/Article/J/J7/200807/Article_20080701214442.html" 
                        target=_blank>2008年6月16—30日时事（原创）</A></TD>
                                    <TD class=listbg align=right width=40>07-01</TD>
                                  </TR>
                                  <TR>
                                    <TD class=listbg2>[<A class=listA 
                        href="http://www.zz6789.com/Article/J/j9/Index.html">时事透视</A>]<A 
                        class=listA 
                        title="文章标题：论抗震救灾精神&#13;&#10;作&nbsp;&nbsp;&nbsp;&nbsp;者：罗先树&#13;&#10;更新时间：2008-6-16 13:35:01" 
                        href="http://www.zz6789.com/Article/J/j9/200806/Article_20080616133501.html" 
                        target=_blank>论抗震救灾精神</A></TD>
                                    <TD class=listbg2 align=right width=40>06-16</TD>
                                  </TR>
                                  <TR>
                                    <TD class=listbg>[<A class=listA 
                        href="http://www.zz6789.com/Article/J/j9/Index.html">时事透视</A>]<A 
                        class=listA 
                        title="文章标题：西方媒体歪曲报道我国西藏“3&#8226;14”事件透视&#13;&#10;作&nbsp;&nbsp;&nbsp;&nbsp;者：罗先树&#13;&#10;更新时间：2008-6-16 13:26:56" 
                        href="http://www.zz6789.com/Article/J/j9/200806/Article_20080616132656.html" 
                        target=_blank>西方媒体歪曲报道我国西藏“3&#8226;14”事</A></TD>
                                    <TD class=listbg align=right width=40>06-16</TD>
                                  </TR>
                                  <TR>
                                    <TD class=listbg2>[<A class=listA 
                        href="http://www.zz6789.com/Article/J/J7/Index.html">时事政治</A>]<A 
                        class=listA 
                        title="文章标题：2008年6月1—15日时事（原创）&#13;&#10;作&nbsp;&nbsp;&nbsp;&nbsp;者：董常祯&#13;&#10;更新时间：2008-6-15 21:42:52" 
                        href="http://www.zz6789.com/Article/J/J7/200806/Article_20080615214252.html" 
                        target=_blank>2008年6月1—15日时事（原创）</A></TD>
                                    <TD class=listbg2 align=right width=40>06-15</TD>
                                  </TR>
                                  <TR>
                                    <TD class=listbg>[<A class=listA 
                        href="http://www.zz6789.com/Article/J/J6/Index.html">新闻专题</A>]<A 
                        class=listA 
                        title="文章标题：512国殇:请记住这100个瞬间(转）&#13;&#10;作&nbsp;&nbsp;&nbsp;&nbsp;者：佚名&#13;&#10;更新时间：2008-6-9 13:56:01" 
                        href="http://www.zz6789.com/Article/J/J6/200806/Article_20080609135601.html" 
                        target=_blank>512国殇:请记住这100个瞬间(转）</A></TD>
                                    <TD class=listbg align=right width=40>06-09</TD>
                                  </TR>
                                  <TR>
                                    <TD class=listbg2>[<A class=listA 
                        href="http://www.zz6789.com/Article/J/J2/Index.html">国际新闻</A>]<A 
                        class=listA 
                        title="文章标题：第三世界兄弟解囊相助：再困难也要帮中国&#13;&#10;作&nbsp;&nbsp;&nbsp;&nbsp;者：godeeg019&#13;&#10;更新时间：2008-6-2 13:57:21" 
                        href="http://www.zz6789.com/Article/J/J2/200806/Article_20080602135721.html" 
                        target=_blank>第三世界兄弟解囊相助：再困难也要帮中国</A></TD>
                                    <TD class=listbg2 align=right width=40>06-02</TD>
                                  </TR>
                                  <TR></TR>
                                </TBODY>
                              </TABLE></TD>
                          </TR>
                        </TBODY>
                      </TABLE>
                    </DIV></TD>
                </TR>
                <TR>
                  <TD><IMG height=20 src="SkinIndex/zin_r41_c21.jpg" width=387 
            border=0></TD>
                </TR>
              </TBODY>
            </TABLE>
            </P></TD>
          <TD width=390 bgColor=#ffffff><P align=center>
            <TABLE id=table6 cellSpacing=0 cellPadding=0 width=387 border=0>
              <TBODY>
                <TR>
                  <TD background=SkinIndex/zin_r46_c21.jpg height=48><P 
            align=left><B>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;最新文章</B></P></TD>
                </TR>
                <TR>
                  <TD background=SkinIndex/zin_r42_c21_r3.jpg height=93><DIV align=center>
                      <TABLE id=table12 height=27 cellSpacing=0 cellPadding=0 width="95%" 
            border=0>
                        <TBODY>
                          <TR>
                            <!--时事新闻-->
                            <TD><TABLE cellSpacing=0 cellPadding=0 width="100%">
                                <TBODY>
                                  <TR>
                                    <TD class=listbg>[<A class=listA 
                        href="http://www.zz6789.com/Article/C/C4/Index.html">高一级</A>]<A 
                        class=listA 
                        title="文章标题：高中会考&nbsp;第3框&nbsp;商品的价值量&nbsp;学案&#13;&#10;作&nbsp;&nbsp;&nbsp;&nbsp;者：封良占&#13;&#10;更新时间：2008-9-9 20:46:05" 
                        href="http://www.zz6789.com/Article/C/C4/200809/Article_20080909204605.html" 
                        target=_blank>高中会考&nbsp;第3框&nbsp;商品的价值量&nbsp;学案</A></TD>
                                    <TD class=listbg align=right width=40><FONT 
                        color=red>09-09</FONT></TD>
                                  </TR>
                                  <TR>
                                    <TD class=listbg2>[<A class=listA 
                        href="http://www.zz6789.com/Article/C/C4/Index.html">高一级</A>]<A 
                        class=listA 
                        title="文章标题：会考-经济常识第1、2框学案&#13;&#10;作&nbsp;&nbsp;&nbsp;&nbsp;者：封良占&#13;&#10;更新时间：2008-9-9 20:41:39" 
                        href="http://www.zz6789.com/Article/C/C4/200809/Article_20080909204139.html" 
                        target=_blank>会考-经济常识第1、2框学案</A></TD>
                                    <TD class=listbg2 align=right width=40><FONT 
                        color=red>09-09</FONT></TD>
                                  </TR>
                                  <TR>
                                    <TD class=listbg>[<A class=listA 
                        href="http://www.zz6789.com/Article/D/D3/Index.html">初三级</A>]<A 
                        class=listA 
                        title="文章标题：2&nbsp;0&nbsp;0&nbsp;8&nbsp;学&nbsp;年&nbsp;度&nbsp;第&nbsp;一&nbsp;学&nbsp;期第一次月考&#13;&#10;作&nbsp;&nbsp;&nbsp;&nbsp;者：zhuwenhui&#13;&#10;更新时间：2008-9-9 19:32:21" 
                        href="http://www.zz6789.com/Article/D/D3/200809/Article_20080909193221.html" 
                        target=_blank>2&nbsp;0&nbsp;0&nbsp;8&nbsp;学&nbsp;年&nbsp;度&nbsp;第&nbsp;一&nbsp;学&nbsp;期第一次月考</A></TD>
                                    <TD class=listbg align=right width=40><FONT 
                        color=red>09-09</FONT></TD>
                                  </TR>
                                  <TR>
                                    <TD class=listbg2>[<A class=listA 
                        href="http://www.zz6789.com/Article/A/A3/Index.html">初三级</A>]<A 
                        class=listA 
                        title="文章标题：初三月考试题&#13;&#10;作&nbsp;&nbsp;&nbsp;&nbsp;者：佚名&#13;&#10;更新时间：2008-9-9 14:09:59" 
                        href="http://www.zz6789.com/Article/A/A3/200809/Article_20080909140959.html" 
                        target=_blank>初三月考试题</A></TD>
                                    <TD class=listbg2 align=right width=40><FONT 
                        color=red>09-09</FONT></TD>
                                  </TR>
                                  <TR>
                                    <TD class=listbg>[<A class=listA 
                        href="http://www.zz6789.com/Article/H/H1/Index.html">教学心得</A>]<A 
                        class=listA 
                        title="文章标题：浅谈对学生的批评教育&#13;&#10;作&nbsp;&nbsp;&nbsp;&nbsp;者：河北&nbsp;鹿泉&nbsp;石坚&#13;&#10;更新时间：2008-9-8 21:19:24" 
                        href="http://www.zz6789.com/Article/H/H1/200809/Article_20080908211924.html" 
                        target=_blank>浅谈对学生的批评教育</A></TD>
                                    <TD class=listbg align=right width=40><FONT 
                        color=red>09-08</FONT></TD>
                                  </TR>
                                  <TR>
                                    <TD class=listbg2>[<A class=listA 
                        href="http://www.zz6789.com/Article/C/C1/Index.html">初一级</A>]<A 
                        class=listA 
                        title="文章标题：粤教版思想品德八年级（上册）期中试卷&#13;&#10;作&nbsp;&nbsp;&nbsp;&nbsp;者：佚名&#13;&#10;更新时间：2008-9-8 10:56:26" 
                        href="http://www.zz6789.com/Article/C/C1/200809/Article_20080908105626.html" 
                        target=_blank>粤教版思想品德八年级（上册）期中试卷</A></TD>
                                    <TD class=listbg2 align=right width=40><FONT 
                        color=red>09-08</FONT></TD>
                                  </TR>
                                  <TR>
                                    <TD class=listbg>[<A class=listA 
                        href="http://www.zz6789.com/Article/k/k3/Index.html">历史真相</A>]<A 
                        class=listA 
                        title="文章标题：苏联人为何“不珍惜”苏联&#13;&#10;作&nbsp;&nbsp;&nbsp;&nbsp;者：佚名&#13;&#10;更新时间：2008-9-8 9:17:03" 
                        href="http://www.zz6789.com/Article/k/k3/200809/Article_20080908091703.html" 
                        target=_blank>苏联人为何“不珍惜”苏联</A></TD>
                                    <TD class=listbg align=right width=40><FONT 
                        color=red>09-08</FONT></TD>
                                  </TR>
                                  <TR>
                                    <TD class=listbg2>[<A class=listA 
                        href="http://www.zz6789.com/Article/D/D5/Index.html">高二级</A>]<A 
                        class=listA 
                        title="文章标题：2008年与2007年《文化生活》教材第一单元改动情况对比&#13;&#10;作&nbsp;&nbsp;&nbsp;&nbsp;者：陈福芝&#13;&#10;更新时间：2008-9-7 23:58:31" 
                        href="http://www.zz6789.com/Article/D/D5/200809/Article_20080907235831.html" 
                        target=_blank>2008年与2007年《文化生活》教材第一单元改</A></TD>
                                    <TD class=listbg2 align=right width=40><FONT 
                        color=red>09-07</FONT></TD>
                                  </TR>
                                  <TR>
                                    <TD class=listbg>[<A class=listA 
                        href="http://www.zz6789.com/Article/J/J7/Index.html">时事政治</A>]<A 
                        class=listA 
                        title="文章标题：2008年8月1—31日时事&#13;&#10;作&nbsp;&nbsp;&nbsp;&nbsp;者：董常祯&#13;&#10;更新时间：2008-9-6 14:30:43" 
                        href="http://www.zz6789.com/Article/J/J7/200809/Article_20080906143043.html" 
                        target=_blank>2008年8月1—31日时事</A></TD>
                                    <TD class=listbg align=right width=40><FONT 
                        color=red>09-06</FONT></TD>
                                  </TR>
                                  <TR>
                                    <TD class=listbg2>[<A class=listA 
                        href="http://www.zz6789.com/Article/H/H1/Index.html">教学心得</A>]<A 
                        class=listA 
                        title="文章标题：让初中思想品德课教学生活化&#13;&#10;作&nbsp;&nbsp;&nbsp;&nbsp;者：瞿珍贵&#13;&#10;更新时间：2008-9-5 14:29:59" 
                        href="http://www.zz6789.com/Article/H/H1/200809/Article_20080905142959.html" 
                        target=_blank>让初中思想品德课教学生活化</A></TD>
                                    <TD class=listbg2 align=right width=40><FONT 
                        color=red>09-05</FONT></TD>
                                  </TR>
                                  <TR></TR>
                                </TBODY>
                              </TABLE></TD>
                          </TR>
                        </TBODY>
                      </TABLE>
                    </DIV></TD>
                </TR>
                <TR>
                  <TD><IMG height=20 src="SkinIndex/zin_r41_c21.jpg" width=387 
            border=0></TD>
                </TR>
              </TBODY>
            </TABLE>
            </P></TD>
          <TD width=8 rowSpan=2>　</TD>
        </TR>
        <TR>
          <TD width=391 bgColor=#ffffff><P align=center>
            <TABLE id=table6 cellSpacing=0 cellPadding=0 width=387 border=0>
              <TBODY>
                <TR>
                  <TD background=SkinIndex/zin_r46_c21.jpg height=48><P align=left><B>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
                      时事透视</B></P></TD>
                </TR>
                <TR>
                  <TD background=SkinIndex/zin_r42_c21_r3.jpg height=93><DIV align=center>
                      <TABLE id=table12 height=27 cellSpacing=0 cellPadding=0 width="95%" 
            border=0>
                        <TBODY>
                          <TR>
                            <!--时事新闻-->
                            <TD><TABLE cellSpacing=0 cellPadding=0 width="100%">
                                <TBODY>
                                  <TR>
                                    <TD class=listbg>[<A class=listA 
                        href="http://www.zz6789.com/Article/J/j9/Index.html">时事透视</A>]<A 
                        class=listA 
                        title="文章标题：论抗震救灾精神&#13;&#10;作&nbsp;&nbsp;&nbsp;&nbsp;者：罗先树&#13;&#10;更新时间：2008-6-16 13:35:01" 
                        href="http://www.zz6789.com/Article/J/j9/200806/Article_20080616133501.html" 
                        target=_blank>论抗震救灾精神</A></TD>
                                    <TD class=listbg align=right width=40>06-16</TD>
                                  </TR>
                                  <TR>
                                    <TD class=listbg2>[<A class=listA 
                        href="http://www.zz6789.com/Article/J/j9/Index.html">时事透视</A>]<A 
                        class=listA 
                        title="文章标题：西方媒体歪曲报道我国西藏“3&#8226;14”事件透视&#13;&#10;作&nbsp;&nbsp;&nbsp;&nbsp;者：罗先树&#13;&#10;更新时间：2008-6-16 13:26:56" 
                        href="http://www.zz6789.com/Article/J/j9/200806/Article_20080616132656.html" 
                        target=_blank>西方媒体歪曲报道我国西藏“3&#8226;14”事件透</A></TD>
                                    <TD class=listbg2 align=right width=40>06-16</TD>
                                  </TR>
                                  <TR>
                                    <TD class=listbg>[<A class=listA 
                        href="http://www.zz6789.com/Article/J/j9/Index.html">时事透视</A>]<A 
                        class=listA 
                        title="文章标题：冰雪记忆&#13;&#10;作&nbsp;&nbsp;&nbsp;&nbsp;者：河北省正定县第五中学政治组：刘玲花&#13;&#10;更新时间：2008-4-5 17:10:46" 
                        href="http://www.zz6789.com/Article/J/j9/200804/Article_20080405171046.html" 
                        target=_blank>冰雪记忆</A></TD>
                                    <TD class=listbg align=right width=40>04-05</TD>
                                  </TR>
                                  <TR>
                                    <TD class=listbg2>[<A class=listA 
                        href="http://www.zz6789.com/Article/J/j9/Index.html">时事透视</A>]<A 
                        class=listA 
                        title="文章标题：初中思想品德课和高中思想政治课&nbsp;贯彻党的十七大精神的指导意见&#13;&#10;作&nbsp;&nbsp;&nbsp;&nbsp;者：dyy730808&#13;&#10;更新时间：2008-1-31 0:03:48" 
                        href="http://www.zz6789.com/Article/J/j9/200801/Article_20080131000348.html" 
                        target=_blank>初中思想品德课和高中思想政治课&nbsp;贯彻党的十七大</A></TD>
                                    <TD class=listbg2 align=right width=40>01-31</TD>
                                  </TR>
                                  <TR>
                                    <TD class=listbg>[<A class=listA 
                        href="http://www.zz6789.com/Article/J/j9/Index.html">时事透视</A>]<A 
                        class=listA 
                        title="文章标题：黑作坊花炮转卖鞭炮厂&#13;&#10;作&nbsp;&nbsp;&nbsp;&nbsp;者：未知&#13;&#10;更新时间：2007-12-16 20:56:27" 
                        href="http://www.zz6789.com/Article/J/j9/200712/Article_20071216205627.html" 
                        target=_blank>黑作坊花炮转卖鞭炮厂</A></TD>
                                    <TD class=listbg align=right width=40>12-16</TD>
                                  </TR>
                                  <TR>
                                    <TD class=listbg2>[<A class=listA 
                        href="http://www.zz6789.com/Article/J/j9/Index.html">时事透视</A>]<A 
                        class=listA 
                        title="文章标题：新京报：重庆花炮作坊爆炸致12名童工死亡调查&#13;&#10;作&nbsp;&nbsp;&nbsp;&nbsp;者：未知&#13;&#10;更新时间：2007-12-16 20:53:59" 
                        href="http://www.zz6789.com/Article/J/j9/200712/Article_20071216205359.html" 
                        target=_blank>新京报：重庆花炮作坊爆炸致12名童工死亡调查</A></TD>
                                    <TD class=listbg2 align=right width=40>12-16</TD>
                                  </TR>
                                  <TR>
                                    <TD class=listbg>[<A class=listA 
                        href="http://www.zz6789.com/Article/J/j9/Index.html">时事透视</A>]<A 
                        class=listA 
                        title="文章标题：农村承包田多年未调引发矛盾&nbsp;部分农民租地耕种&#13;&#10;作&nbsp;&nbsp;&nbsp;&nbsp;者：未知&#13;&#10;更新时间：2007-12-16 20:52:41" 
                        href="http://www.zz6789.com/Article/J/j9/200712/Article_20071216205241.html" 
                        target=_blank>农村承包田多年未调引发矛盾&nbsp;部分农民租地耕种</A></TD>
                                    <TD class=listbg align=right width=40>12-16</TD>
                                  </TR>
                                  <TR>
                                    <TD class=listbg2>[<A class=listA 
                        href="http://www.zz6789.com/Article/J/j9/Index.html">时事透视</A>]<A 
                        class=listA 
                        title="文章标题：2007全国百强县评比取消发布背后&#13;&#10;作&nbsp;&nbsp;&nbsp;&nbsp;者：未知&#13;&#10;更新时间：2007-12-16 20:48:56" 
                        href="http://www.zz6789.com/Article/J/j9/200712/Article_20071216204856.html" 
                        target=_blank>2007全国百强县评比取消发布背后</A></TD>
                                    <TD class=listbg2 align=right width=40>12-16</TD>
                                  </TR>
                                  <TR>
                                    <TD class=listbg>[<A class=listA 
                        href="http://www.zz6789.com/Article/J/j9/Index.html">时事透视</A>]<A 
                        class=listA 
                        title="文章标题：广东阳江黑帮头目开赌场发家&#13;&#10;作&nbsp;&nbsp;&nbsp;&nbsp;者：未知&#13;&#10;更新时间：2007-12-16 20:44:26" 
                        href="http://www.zz6789.com/Article/J/j9/200712/Article_20071216204426.html" 
                        target=_blank>广东阳江黑帮头目开赌场发家</A></TD>
                                    <TD class=listbg align=right width=40>12-16</TD>
                                  </TR>
                                  <TR>
                                    <TD class=listbg2>[<A class=listA 
                        href="http://www.zz6789.com/Article/J/j9/Index.html">时事透视</A>]<A 
                        class=listA 
                        title="文章标题：伤痛，为了记忆还是忘却？&#13;&#10;作&nbsp;&nbsp;&nbsp;&nbsp;者：未知&#13;&#10;更新时间：2007-12-16 20:42:42" 
                        href="http://www.zz6789.com/Article/J/j9/200712/Article_20071216204242.html" 
                        target=_blank>伤痛，为了记忆还是忘却？</A></TD>
                                    <TD class=listbg2 align=right width=40>12-16</TD>
                                  </TR>
                                  <TR></TR>
                                </TBODY>
                              </TABLE></TD>
                          </TR>
                        </TBODY>
                      </TABLE>
                    </DIV></TD>
                </TR>
                <TR>
                  <TD><IMG height=20 src="SkinIndex/zin_r41_c21.jpg" width=387 
            border=0></TD>
                </TR>
              </TBODY>
            </TABLE>
            </P></TD>
          <TD width=390 bgColor=#ffffff><P align=center>
            <TABLE id=table6 cellSpacing=0 cellPadding=0 width=387 border=0>
              <TBODY>
                <TR>
                  <TD background=SkinIndex/zin_r46_c21.jpg height=48><P 
            align=left><B>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;最新课件</B></P></TD>
                </TR>
                <TR>
                  <TD background=SkinIndex/zin_r42_c21_r3.jpg height=93><DIV align=center>
                      <TABLE id=table12 height=27 cellSpacing=0 cellPadding=0 width="95%" 
            border=0>
                        <TBODY>
                          <TR>
                            <!--时事新闻-->
                            <TD><TABLE cellSpacing=0 cellPadding=0 width="100%">
                                <TBODY>
                                  <TR>
                                    <TD class=listbg>[<A class=listA 
                        href="http://www.zz6789.com/Soft/JuniorClass/c3/Index.html">初三级</A>]<A 
                        class=listA title='1.2实践中的"一国两制"' 
                        href="http://www.zz6789.com/Soft/JuniorClass/c3/200809/Soft_3839.html" 
                        target=_blank>1.2实践中的"一国两制"</A></TD>
                                    <TD class=listbg align=right width=40>09-03</TD>
                                  </TR>
                                  <TR>
                                    <TD class=listbg2>[<A class=listA 
                        href="http://www.zz6789.com/Soft/JuniorClass/c3/Index.html">初三级</A>]<A 
                        class=listA title=1.2人民当家作主的政治制度 
                        href="http://www.zz6789.com/Soft/JuniorClass/c3/200809/Soft_3838.html" 
                        target=_blank>1.2人民当家作主的政治制度</A></TD>
                                    <TD class=listbg2 align=right width=40>09-03</TD>
                                  </TR>
                                  <TR>
                                    <TD class=listbg>[<A class=listA 
                        href="http://www.zz6789.com/Soft/JuniorClass/c3/Index.html">初三级</A>]<A 
                        class=listA title=1.2富有活力的基本经济制度 
                        href="http://www.zz6789.com/Soft/JuniorClass/c3/200809/Soft_3837.html" 
                        target=_blank>1.2富有活力的基本经济制度</A></TD>
                                    <TD class=listbg align=right width=40>09-03</TD>
                                  </TR>
                                  <TR>
                                    <TD class=listbg2>[<A class=listA 
                        href="http://www.zz6789.com/Soft/JuniorClass/c3/Index.html">初三级</A>]<A 
                        class=listA title=1.1初级阶段的基本路线 
                        href="http://www.zz6789.com/Soft/JuniorClass/c3/200809/Soft_3836.html" 
                        target=_blank>1.1初级阶段的基本路线</A></TD>
                                    <TD class=listbg2 align=right width=40>09-03</TD>
                                  </TR>
                                  <TR>
                                    <TD class=listbg>[<A class=listA 
                        href="http://www.zz6789.com/Soft/JuniorClass/c3/Index.html">初三级</A>]<A 
                        class=listA title=1.1初级阶段的主要矛盾和根本任务 
                        href="http://www.zz6789.com/Soft/JuniorClass/c3/200809/Soft_3835.html" 
                        target=_blank>1.1初级阶段的主要矛盾和根本任务</A></TD>
                                    <TD class=listbg align=right width=40>09-03</TD>
                                  </TR>
                                  <TR>
                                    <TD class=listbg2>[<A class=listA 
                        href="http://www.zz6789.com/Soft/JuniorClass/c3/Index.html">初三级</A>]<A 
                        class=listA title=1.1初级阶段的社会主义 
                        href="http://www.zz6789.com/Soft/JuniorClass/c3/200809/Soft_3834.html" 
                        target=_blank>1.1初级阶段的社会主义</A></TD>
                                    <TD class=listbg2 align=right width=40>09-03</TD>
                                  </TR>
                                  <TR>
                                    <TD class=listbg>[<A class=listA 
                        href="http://www.zz6789.com/Soft/Highclass/Economy/Levy01/Index.html">第一课 
                                      神奇的货币</A>]<A class=listA title=神奇的货币 
                        href="http://www.zz6789.com/Soft/Highclass/Economy/Levy01/200809/Soft_3833.html" 
                        target=_blank>神奇的货币</A></TD>
                                    <TD class=listbg align=right width=40>09-02</TD>
                                  </TR>
                                  <TR>
                                    <TD class=listbg2>[<A class=listA 
                        href="http://www.zz6789.com/Soft/M/Index.html">复习专题课件</A>]<A 
                        class=listA title=政治生活复习第二课 
                        href="http://www.zz6789.com/Soft/M/200809/Soft_3831.html" 
                        target=_blank>政治生活复习第二课</A></TD>
                                    <TD class=listbg2 align=right width=40>09-02</TD>
                                  </TR>
                                  <TR>
                                    <TD class=listbg>[<A class=listA 
                        href="http://www.zz6789.com/Soft/N/Index.html">德育教育课件</A>]<A 
                        class=listA title=地震 
                        href="http://www.zz6789.com/Soft/N/200808/Soft_3828.html" 
                        target=_blank>地震</A></TD>
                                    <TD class=listbg align=right width=40>08-04</TD>
                                  </TR>
                                  <TR>
                                    <TD class=listbg2>[<A class=listA 
                        href="http://www.zz6789.com/Soft/JuniorClass/c2/Index.html">初二级</A>]<A 
                        class=listA title=平等待人（第二课时） 
                        href="http://www.zz6789.com/Soft/JuniorClass/c2/200807/Soft_3827.html" 
                        target=_blank>平等待人（第二课时）</A></TD>
                                    <TD class=listbg2 align=right width=40>07-24</TD>
                                  </TR>
                                  <TR></TR>
                                </TBODY>
                              </TABLE></TD>
                          </TR>
                        </TBODY>
                      </TABLE>
                    </DIV></TD>
                </TR>
                <TR>
                  <TD><IMG height=20 src="SkinIndex/zin_r41_c21.jpg" width=387 
            border=0></TD>
                </TR>
              </TBODY>
            </TABLE>
            </P></TD>
        </TR>
      </TBODY>
    </TABLE>
  </DIV>
  <DIV align=center>
    <TABLE id=table7 height=80 cellSpacing=0 cellPadding=0 width=797 
background=SkinIndex/new_wow_43.jpg border=0>
      <TBODY>
        <TR>
          <TD width=8>　</TD>
          <TD align=middle width=781 bgColor=#ffffff height=80><SCRIPT language=javascript src="SkinIndex/1.js"></SCRIPT>
          </TD>
          <TD width=8>　</TD>
        </TR>
      </TBODY>
    </TABLE>
  </DIV>
  <DIV align=center>
    <TABLE id=table7 height=100 cellSpacing=0 cellPadding=0 width=797 
background=SkinIndex/new_wow_43.jpg border=0>
      <TBODY>
        <TR>
          <TD width=8>　</TD>
          <TD vAlign=top align=middle width=197 bgColor=#ffffff><TABLE id=table5 height=200 cellSpacing=0 cellPadding=0 width=197 
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
                  <TD background=SkinIndex/zin_r13_c1.gif>
                   <% call Showhot(8,16) %>
                    </TD>
                </TR>
                
                <TR>
                  <TD height=24><IMG height=25 src="SkinIndex/zin_r18_c1.gif" 
            width=197 border=0></TD>
                </TR>
              </TBODY>
            </TABLE>
            <TABLE id=table5 height=200 cellSpacing=0 cellPadding=0 width=197 
border=0>
              <TBODY>
                <TR>
                  <TD background=SkinIndex/zdl_8.gif height=35><P align=center><B>欢 迎 新 生</B></P></TD>
                </TR>
                <TR>
                  <TD height=24><IMG height=24 src="SkinIndex/zin_r11_c1.gif" 
            width=197 border=0></TD>
                </TR>
                
                
                <TR>
                  <TD align=middle background=SkinIndex/zin_r13_c1.gif>
				  <% call ShownewUser(5) %>
                  </TD>
                </TR>
                
                
                <TR>
                  <TD height=24><IMG height=25 src="SkinIndex/zin_r18_c1.gif" 
            width=197 border=0></TD>
                </TR>
              </TBODY>
            </TABLE>
            <TABLE id=table5 height=200 cellSpacing=0 cellPadding=0 width=197 
border=0>
              <TBODY>
                <TR>
                  <TD background=SkinIndex/zdl_8.gif height=35><P align=center><B>本 站 调 查</B></P></TD>
                </TR>
                <TR>
                  <TD height=24><IMG height=24 src="SkinIndex/zin_r11_c1.gif" 
            width=197 border=0></TD>
                </TR>
                <TR>
                  		<TD align=middle background=SkinIndex/zin_r13_c1.gif  height=200>
            				<% call showvote() %>
           				 </TD><td bgcolor="
                </TR>
                <TR>
                  <TD height=24><IMG height=25 src="SkinIndex/zin_r18_c1.gif" 
            width=197 border=0></TD>
                </TR>
              </TBODY>
            </TABLE></TD>
          <TD vAlign=top align=middle width=584 bgColor=#ffffff><TABLE id=table8 cellSpacing=0 cellPadding=0 width=577 border=0>
              <TBODY>
                <TR>
                  <TD width=287 background=SkinIndex/zin_r48_c1.gif 
            height=48><B>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;教材分析</B></TD>
                  <TD width=3>　</TD>
                  <TD width=287 background=SkinIndex/zin_r48_c1.gif 
            height=48><B>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;高考研究</B></TD>
                </TR>
                <TR>
                  <TD background=SkinIndex/zin_r49_c1.gif height=200><DIV align=center>
                      <TABLE id=table12 height=27 cellSpacing=0 cellPadding=0 width="95%" 
            border=0>
                        <TBODY>
                          <TR>
                            <!--文章1-->
                            <TD><TABLE cellSpacing=0 cellPadding=0 width="100%">
                                <TBODY>
                                  <TR>
                                    <TD class=listbg vAlign=top width=10><IMG alt=普通文章 
                        src="SkinIndex/article_common.gif"></TD>
                                    <TD class=listbg>[<A class=listA 
                        href="http://www.zz6789.com/Article/A/A3/Index.html">初三级</A>]<A 
                        class=listA title=初三月考试题 
                        href="http://www.zz6789.com/Article/A/A3/200809/Article_20080909140959.html" 
                        target=_blank>初三月考试题</A></TD>
                                  </TR>
                                  <TR>
                                    <TD class=listbg2 vAlign=top width=10><IMG alt=普通文章 
                        src="SkinIndex/article_common.gif"></TD>
                                    <TD class=listbg2>[<A class=listA 
                        href="http://www.zz6789.com/Article/A/A1/Index.html">初一级</A>]<A 
                        class=listA title=教学过程如何体现学生的主体性. 
                        href="http://www.zz6789.com/Article/A/A1/200808/Article_20080813161106.html" 
                        target=_blank>教学过程如何体现学生的主体性.</A></TD>
                                  </TR>
                                  <TR>
                                    <TD class=listbg vAlign=top width=10><IMG alt=普通文章 
                        src="SkinIndex/article_common.gif"></TD>
                                    <TD class=listbg>[<A class=listA 
                        href="http://www.zz6789.com/Article/A/A3/Index.html">初三级</A>]<A 
                        class=listA title=2&nbsp;0&nbsp;0&nbsp;8年芜湖市初中毕业学业考试 
                        href="http://www.zz6789.com/Article/A/A3/200807/Article_20080702161758.html" 
                        target=_blank>2&nbsp;0&nbsp;0&nbsp;8年芜湖市初中毕业学业考试</A></TD>
                                  </TR>
                                  <TR>
                                    <TD class=listbg2 vAlign=top width=10><IMG alt=普通文章 
                        src="SkinIndex/article_common.gif"></TD>
                                    <TD class=listbg2>[<A class=listA 
                        href="http://www.zz6789.com/Article/A/A6/Index.html">高三级</A>]<A 
                        class=listA title=2008年文化生活热点问题及命题趋势训练题 
                        href="http://www.zz6789.com/Article/A/A6/200806/Article_20080628162241.html" 
                        target=_blank>2008年文化生活热点问题及命题趋</A></TD>
                                  </TR>
                                  <TR>
                                    <TD class=listbg vAlign=top width=10><IMG alt=普通文章 
                        src="SkinIndex/article_common.gif"></TD>
                                    <TD class=listbg>[<A class=listA 
                        href="http://www.zz6789.com/Article/A/A3/Index.html">初三级</A>]<A 
                        class=listA title=滑集中学2008届初三第二次月考思想品德 
                        href="http://www.zz6789.com/Article/A/A3/200804/Article_20080427141009.html" 
                        target=_blank>滑集中学2008届初三第二次月考思</A></TD>
                                  </TR>
                                  <TR>
                                    <TD class=listbg2 vAlign=top width=10><IMG alt=普通文章 
                        src="SkinIndex/article_common.gif"></TD>
                                    <TD class=listbg2>[<A class=listA 
                        href="http://www.zz6789.com/Article/A/A3/Index.html">初三级</A>]<A 
                        class=listA title=中国加入WTO有何积极意义？ 
                        href="http://www.zz6789.com/Article/A/A3/200804/Article_20080421210216.html" 
                        target=_blank>中国加入WTO有何积极意义？</A></TD>
                                  </TR>
                                  <TR>
                                    <TD class=listbg vAlign=top width=10><IMG alt=普通文章 
                        src="SkinIndex/article_common.gif"></TD>
                                    <TD class=listbg>[<A class=listA 
                        href="http://www.zz6789.com/Article/A/A3/Index.html">初三级</A>]<A 
                        class=listA title=2007-2008学年第一学期期末九年级思想政治试卷 
                        href="http://www.zz6789.com/Article/A/A3/200804/Article_20080421142028.html" 
                        target=_blank>2007-2008学年第一学期期末九年级</A></TD>
                                  </TR>
                                  <TR>
                                    <TD class=listbg2 vAlign=top width=10><IMG alt=普通文章 
                        src="SkinIndex/article_common.gif"></TD>
                                    <TD class=listbg2>[<A class=listA 
                        href="http://www.zz6789.com/Article/A/A3/Index.html">初三级</A>]<A 
                        class=listA title=初四上简答题 
                        href="http://www.zz6789.com/Article/A/A3/200804/Article_20080421141732.html" 
                        target=_blank>初四上简答题</A></TD>
                                  </TR>
                                  <TR>
                                    <TD class=listbg vAlign=top width=10><IMG alt=普通文章 
                        src="SkinIndex/article_common.gif"></TD>
                                    <TD class=listbg>[<A class=listA 
                        href="http://www.zz6789.com/Article/A/A3/Index.html">初三级</A>]<A 
                        class=listA title=高洞中学二00八年第一次模拟测思想品德试题 
                        href="http://www.zz6789.com/Article/A/A3/200804/Article_20080420084414.html" 
                        target=_blank>高洞中学二00八年第一次模拟测思</A></TD>
                                  </TR>
                                  <TR>
                                    <TD class=listbg2 vAlign=top width=10><IMG alt=推荐文章 
                        src="SkinIndex/article_elite.gif"></TD>
                                    <TD class=listbg2>[<A class=listA 
                        href="http://www.zz6789.com/Article/A/A6/Index.html">高三级</A>]<A 
                        class=listA title=东升学校（按指导意见）十七大与经济生活结合点专题四 
                        href="http://www.zz6789.com/Article/A/A6/200804/Article_20080413153656.html" 
                        target=_blank>东升学校（按指导意见）十七大与</A></TD>
                                  </TR>
                                  <TR></TR>
                                </TBODY>
                              </TABLE></TD>
                          </TR>
                        </TBODY>
                      </TABLE>
                    </DIV></TD>
                  <TD width=3>　</TD>
                  <TD background=SkinIndex/zin_r49_c1.gif height=200><DIV align=center>
                      <TABLE id=table12 height=27 cellSpacing=0 cellPadding=0 width="95%" 
            border=0>
                        <TBODY>
                          <TR>
                            <!--文章1-->
                            <TD><TABLE cellSpacing=0 cellPadding=0 width="100%">
                                <TBODY>
                                  <TR>
                                    <TD class=listbg vAlign=top width=10><IMG alt=普通文章 
                        src="SkinIndex/article_common.gif"></TD>
                                    <TD class=listbg>[<A class=listA 
                        href="http://www.zz6789.com/Article/E/E2/Index.html">高考分析</A>]<A 
                        class=listA 
                        title=高考热点专题一&nbsp;&nbsp;&nbsp;&nbsp;关注民生，共建和谐 
                        href="http://www.zz6789.com/Article/E/E2/200809/Article_20080901084735.html" 
                        target=_blank>高考热点专题一&nbsp;&nbsp;&nbsp;&nbsp;关注民生，共</A></TD>
                                  </TR>
                                  <TR>
                                    <TD class=listbg2 vAlign=top width=10><IMG alt=推荐文章 
                        src="SkinIndex/article_elite.gif"></TD>
                                    <TD class=listbg2>[<A class=listA 
                        href="http://www.zz6789.com/Article/E/E8/Index.html">历界高考题</A>]<A 
                        class=listA title=2008年高考政治常识试题汇总 
                        href="http://www.zz6789.com/Article/E/E8/200808/Article_20080815091349.html" 
                        target=_blank>2008年高考政治常识试题汇总</A></TD>
                                  </TR>
                                  <TR>
                                    <TD class=listbg vAlign=top width=10><IMG alt=推荐文章 
                        src="SkinIndex/article_elite.gif"></TD>
                                    <TD class=listbg>[<A class=listA 
                        href="http://www.zz6789.com/Article/E/E8/Index.html">历界高考题</A>]<A 
                        class=listA title=2008年高考哲学常识试题汇总 
                        href="http://www.zz6789.com/Article/E/E8/200808/Article_20080815091140.html" 
                        target=_blank>2008年高考哲学常识试题汇总</A></TD>
                                  </TR>
                                  <TR>
                                    <TD class=listbg2 vAlign=top width=10><IMG alt=推荐文章 
                        src="SkinIndex/article_elite.gif"></TD>
                                    <TD class=listbg2>[<A class=listA 
                        href="http://www.zz6789.com/Article/E/E2/Index.html">高考分析</A>]<A 
                        class=listA title=2008年高考经济常识试题汇总 
                        href="http://www.zz6789.com/Article/E/E2/200808/Article_20080815090848.html" 
                        target=_blank>2008年高考经济常识试题汇总</A></TD>
                                  </TR>
                                  <TR>
                                    <TD class=listbg vAlign=top width=10><IMG alt=推荐文章 
                        src="SkinIndex/article_elite.gif"></TD>
                                    <TD class=listbg>[<A class=listA 
                        href="http://www.zz6789.com/Article/E/E3/Index.html">备考策略</A>]<A 
                        class=listA title=从08年高考文科综合全国卷看09年政治高考备考 
                        href="http://www.zz6789.com/Article/E/E3/200807/Article_20080725225113.html" 
                        target=_blank>从08年高考文科综合全国卷看09年</A></TD>
                                  </TR>
                                  <TR>
                                    <TD class=listbg2 vAlign=top width=10><IMG alt=普通文章 
                        src="SkinIndex/article_common.gif"></TD>
                                    <TD class=listbg2>[<A class=listA 
                        href="http://www.zz6789.com/Article/E/E8/Index.html">历界高考题</A>]<A 
                        class=listA title=2008年广东高考政治试题答案(A卷相片版4) 
                        href="http://www.zz6789.com/Article/E/E8/200806/Article_20080617215719.html" 
                        target=_blank>2008年广东高考政治试题答案(A卷</A></TD>
                                  </TR>
                                  <TR>
                                    <TD class=listbg vAlign=top width=10><IMG alt=普通文章 
                        src="SkinIndex/article_common.gif"></TD>
                                    <TD class=listbg>[<A class=listA 
                        href="http://www.zz6789.com/Article/E/E8/Index.html">历界高考题</A>]<A 
                        class=listA title=2008年广东高考政治试题答案(A卷相片版3) 
                        href="http://www.zz6789.com/Article/E/E8/200806/Article_20080617215225.html" 
                        target=_blank>2008年广东高考政治试题答案(A卷</A></TD>
                                  </TR>
                                  <TR>
                                    <TD class=listbg2 vAlign=top width=10><IMG alt=普通文章 
                        src="SkinIndex/article_common.gif"></TD>
                                    <TD class=listbg2>[<A class=listA 
                        href="http://www.zz6789.com/Article/E/E8/Index.html">历界高考题</A>]<A 
                        class=listA title=2008年广东高考政治试题答案(A卷相片版2) 
                        href="http://www.zz6789.com/Article/E/E8/200806/Article_20080617214448.html" 
                        target=_blank>2008年广东高考政治试题答案(A卷</A></TD>
                                  </TR>
                                  <TR>
                                    <TD class=listbg vAlign=top width=10><IMG alt=普通文章 
                        src="SkinIndex/article_common.gif"></TD>
                                    <TD class=listbg>[<A class=listA 
                        href="http://www.zz6789.com/Article/E/E8/Index.html">历界高考题</A>]<A 
                        class=listA title=2008年广东高考政治试题答案(A卷相片版1) 
                        href="http://www.zz6789.com/Article/E/E8/200806/Article_20080617204930.html" 
                        target=_blank>2008年广东高考政治试题答案(A卷</A></TD>
                                  </TR>
                                  <TR>
                                    <TD class=listbg2 vAlign=top width=10><IMG alt=普通文章 
                        src="SkinIndex/article_common.gif"></TD>
                                    <TD class=listbg2>[<A class=listA 
                        href="http://www.zz6789.com/Article/E/E8/Index.html">历界高考题</A>]<A 
                        class=listA title=2008年广东高考政治试题(A卷word版) 
                        href="http://www.zz6789.com/Article/E/E8/200806/Article_20080617204107.html" 
                        target=_blank>2008年广东高考政治试题(A卷word</A></TD>
                                  </TR>
                                  <TR></TR>
                                </TBODY>
                              </TABLE></TD>
                          </TR>
                        </TBODY>
                      </TABLE>
                    </DIV></TD>
                </TR>
                <TR>
                  <TD height=17><IMG height=17 src="SkinIndex/zin_r50_c1.gif" 
            width=287 border=0></TD>
                  <TD width=3>　</TD>
                  <TD height=17><IMG height=17 src="SkinIndex/zin_r50_c1.gif" 
            width=287 border=0></TD>
                </TR>
              </TBODY>
            </TABLE>
            <TABLE id=table8 cellSpacing=0 cellPadding=0 width=577 border=0>
              <TBODY>
                <TR>
                  <TD width=287 background=SkinIndex/zin_r48_c1.gif 
            height=48><B>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;教案选编</B></TD>
                  <TD width=3>　</TD>
                  <TD width=287 background=SkinIndex/zin_r48_c1.gif 
            height=48><B>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;中考研究</B></TD>
                </TR>
                <TR>
                  <TD background=SkinIndex/zin_r49_c1.gif height=200><DIV align=center>
                      <TABLE id=table12 height=27 cellSpacing=0 cellPadding=0 width="95%" 
            border=0>
                        <TBODY>
                          <TR>
                            <!--文章1-->
                            <TD><TABLE cellSpacing=0 cellPadding=0 width="100%">
                                <TBODY>
                                  <TR>
                                    <TD class=listbg vAlign=top width=10><IMG alt=推荐文章 
                        src="SkinIndex/article_elite.gif"></TD>
                                    <TD class=listbg>[<A class=listA 
                        href="http://www.zz6789.com/Article/B/B3/Index.html">初三级</A>]<A 
                        class=listA title=2008年中考时政热点预测 
                        href="http://www.zz6789.com/Article/B/B3/200806/Article_20080619111204.html" 
                        target=_blank>2008年中考时政热点预测</A></TD>
                                  </TR>
                                  <TR>
                                    <TD class=listbg2 vAlign=top width=10><IMG alt=普通文章 
                        src="SkinIndex/article_common.gif"></TD>
                                    <TD class=listbg2>[<A class=listA 
                        href="http://www.zz6789.com/Article/B/B2/Index.html">初二级</A>]<A 
                        class=listA title=财产留给谁教学设计 
                        href="http://www.zz6789.com/Article/B/B2/200805/Article_20080524200913.html" 
                        target=_blank>财产留给谁教学设计</A></TD>
                                  </TR>
                                  <TR>
                                    <TD class=listbg vAlign=top width=10><IMG alt=普通文章 
                        src="SkinIndex/article_common.gif"></TD>
                                    <TD class=listbg>[<A class=listA 
                        href="http://www.zz6789.com/Article/B/B2/Index.html">初二级</A>]<A 
                        class=listA title=初二下思品课复习纲要 
                        href="http://www.zz6789.com/Article/B/B2/200804/Article_20080421142347.html" 
                        target=_blank>初二下思品课复习纲要</A></TD>
                                  </TR>
                                  <TR>
                                    <TD class=listbg2 vAlign=top width=10><IMG alt=普通文章 
                        src="SkinIndex/article_common.gif"></TD>
                                    <TD class=listbg2>[<A class=listA 
                        href="http://www.zz6789.com/Article/B/B5/Index.html">高二级</A>]<A 
                        class=listA title=事物运动时有规律的教案 
                        href="http://www.zz6789.com/Article/B/B5/200803/Article_20080330161846.html" 
                        target=_blank>事物运动时有规律的教案</A></TD>
                                  </TR>
                                  <TR>
                                    <TD class=listbg vAlign=top width=10><IMG alt=普通文章 
                        src="SkinIndex/article_common.gif"></TD>
                                    <TD class=listbg>[<A class=listA 
                        href="http://www.zz6789.com/Article/B/B3/Index.html">初三级</A>]<A 
                        class=listA 
                        title=信息化环境下的教学设计&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;《了解祖国&nbsp;&nbsp;爱我中华》单元教学案例 
                        href="http://www.zz6789.com/Article/B/B3/200803/Article_20080304102437.html" 
                        target=_blank>信息化环境下的教学设计&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</A></TD>
                                  </TR>
                                  <TR>
                                    <TD class=listbg2 vAlign=top width=10><IMG alt=普通文章 
                        src="SkinIndex/article_common.gif"></TD>
                                    <TD class=listbg2>[<A class=listA 
                        href="http://www.zz6789.com/Article/B/B1/Index.html">初一级</A>]<A 
                        class=listA title=青春发育 
                        href="http://www.zz6789.com/Article/B/B1/200801/Article_20080118092924.html" 
                        target=_blank>青春发育</A></TD>
                                  </TR>
                                  <TR>
                                    <TD class=listbg vAlign=top width=10><IMG alt=普通文章 
                        src="SkinIndex/article_common.gif"></TD>
                                    <TD class=listbg>[<A class=listA 
                        href="http://www.zz6789.com/Article/B/B4/Index.html">高一级</A>]<A 
                        class=listA title=税收及其种类 
                        href="http://www.zz6789.com/Article/B/B4/200801/Article_20080116220749.html" 
                        target=_blank>税收及其种类</A></TD>
                                  </TR>
                                  <TR>
                                    <TD class=listbg2 vAlign=top width=10><IMG alt=普通文章 
                        src="SkinIndex/article_common.gif"></TD>
                                    <TD class=listbg2>[<A class=listA 
                        href="http://www.zz6789.com/Article/B/B4/Index.html">高一级</A>]<A 
                        class=listA title=处理民族关系的原则 
                        href="http://www.zz6789.com/Article/B/B4/200801/Article_20080116220614.html" 
                        target=_blank>处理民族关系的原则</A></TD>
                                  </TR>
                                  <TR>
                                    <TD class=listbg vAlign=top width=10><IMG alt=普通文章 
                        src="SkinIndex/article_common.gif"></TD>
                                    <TD class=listbg>[<A class=listA 
                        href="http://www.zz6789.com/Article/B/B4/Index.html">高一级</A>]<A 
                        class=listA title=积极参与国际经济竞争与合作 
                        href="http://www.zz6789.com/Article/B/B4/200801/Article_20080116220438.html" 
                        target=_blank>积极参与国际经济竞争与合作</A></TD>
                                  </TR>
                                  <TR>
                                    <TD class=listbg2 vAlign=top width=10><IMG alt=普通文章 
                        src="SkinIndex/article_common.gif"></TD>
                                    <TD class=listbg2>[<A class=listA 
                        href="http://www.zz6789.com/Article/B/B2/Index.html">初二级</A>]<A 
                        class=listA title=法律规范经济行为 
                        href="http://www.zz6789.com/Article/B/B2/200801/Article_20080116125154.html" 
                        target=_blank>法律规范经济行为</A></TD>
                                  </TR>
                                  <TR></TR>
                                </TBODY>
                              </TABLE></TD>
                          </TR>
                        </TBODY>
                      </TABLE>
                    </DIV></TD>
                  <TD width=3>　</TD>
                  <TD background=SkinIndex/zin_r49_c1.gif height=200><DIV align=center>
                      <TABLE id=table12 height=27 cellSpacing=0 cellPadding=0 width="95%" 
            border=0>
                        <TBODY>
                          <TR>
                            <!--文章1-->
                            <TD><TABLE cellSpacing=0 cellPadding=0 width="100%">
                                <TBODY>
                                  <TR>
                                    <TD class=listbg vAlign=top width=10><IMG alt=普通文章 
                        src="SkinIndex/article_common.gif"></TD>
                                    <TD class=listbg>[<A class=listA 
                        href="http://www.zz6789.com/Article/F/F6/Index.html">历界中考题</A>]<A 
                        class=listA title=2008武汉市中考试题 
                        href="http://www.zz6789.com/Article/F/F6/200807/Article_20080708231409.html" 
                        target=_blank>2008武汉市中考试题</A></TD>
                                  </TR>
                                  <TR>
                                    <TD class=listbg2 vAlign=top width=10><IMG alt=普通文章 
                        src="SkinIndex/article_common.gif"></TD>
                                    <TD class=listbg2>[<A class=listA 
                        href="http://www.zz6789.com/Article/F/F6/Index.html">历界中考题</A>]<A 
                        class=listA title=2007武汉市中考试题 
                        href="http://www.zz6789.com/Article/F/F6/200807/Article_20080708231328.html" 
                        target=_blank>2007武汉市中考试题</A></TD>
                                  </TR>
                                  <TR>
                                    <TD class=listbg vAlign=top width=10><IMG alt=普通文章 
                        src="SkinIndex/article_common.gif"></TD>
                                    <TD class=listbg>[<A class=listA 
                        href="http://www.zz6789.com/Article/F/F6/Index.html">历界中考题</A>]<A 
                        class=listA title=2008辽宁省十二市中考思想品德试卷及答案 
                        href="http://www.zz6789.com/Article/F/F6/200807/Article_20080704001616.html" 
                        target=_blank>2008辽宁省十二市中考思想品德试</A></TD>
                                  </TR>
                                  <TR>
                                    <TD class=listbg2 vAlign=top width=10><IMG alt=普通文章 
                        src="SkinIndex/article_common.gif"></TD>
                                    <TD class=listbg2>[<A class=listA 
                        href="http://www.zz6789.com/Article/F/F1/Index.html">中考信息</A>]<A 
                        class=listA title=广西南宁市2008年中考政史试题 
                        href="http://www.zz6789.com/Article/F/F1/200807/Article_20080703102040.html" 
                        target=_blank>广西南宁市2008年中考政史试题</A></TD>
                                  </TR>
                                  <TR>
                                    <TD class=listbg vAlign=top width=10><IMG alt=普通文章 
                        src="SkinIndex/article_common.gif"></TD>
                                    <TD class=listbg>[<A class=listA 
                        href="http://www.zz6789.com/Article/F/F4/Index.html">中考资料</A>]<A 
                        class=listA title=抗震救灾中考政治题 
                        href="http://www.zz6789.com/Article/F/F4/200807/Article_20080703101429.html" 
                        target=_blank>抗震救灾中考政治题</A></TD>
                                  </TR>
                                  <TR>
                                    <TD class=listbg2 vAlign=top width=10><IMG alt=普通文章 
                        src="SkinIndex/article_common.gif"></TD>
                                    <TD class=listbg2>[<A class=listA 
                        href="http://www.zz6789.com/Article/F/F5/Index.html">中考模拟题</A>]<A 
                        class=listA title=2008年中考政治模拟试题 
                        href="http://www.zz6789.com/Article/F/F5/200807/Article_20080702075952.html" 
                        target=_blank>2008年中考政治模拟试题</A></TD>
                                  </TR>
                                  <TR>
                                    <TD class=listbg vAlign=top width=10><IMG alt=推荐文章 
                        src="SkinIndex/article_elite.gif"></TD>
                                    <TD class=listbg>[<A class=listA 
                        href="http://www.zz6789.com/Article/F/F4/Index.html">中考资料</A>]<A 
                        class=listA title=“5&#8226;12”地震中考专题 
                        href="http://www.zz6789.com/Article/F/F4/200806/Article_20080618105103.html" 
                        target=_blank>“5&#8226;12”地震中考专题</A></TD>
                                  </TR>
                                  <TR>
                                    <TD class=listbg2 vAlign=top width=10><IMG alt=推荐文章 
                        src="SkinIndex/article_elite.gif"></TD>
                                    <TD class=listbg2>[<A class=listA 
                        href="http://www.zz6789.com/Article/F/F4/Index.html">中考资料</A>]<A 
                        class=listA title=2008中考最新时事政治必备试题精选 
                        href="http://www.zz6789.com/Article/F/F4/200806/Article_20080615103602.html" 
                        target=_blank>2008中考最新时事政治必备试题精</A></TD>
                                  </TR>
                                  <TR>
                                    <TD class=listbg vAlign=top width=10><IMG alt=推荐文章 
                        src="SkinIndex/article_elite.gif"></TD>
                                    <TD class=listbg>[<A class=listA 
                        href="http://www.zz6789.com/Article/F/F4/Index.html">中考资料</A>]<A 
                        class=listA title=专题十一&nbsp;&nbsp;&nbsp;唱响红歌，收获感动 
                        href="http://www.zz6789.com/Article/F/F4/200806/Article_20080615103132.html" 
                        target=_blank>专题十一&nbsp;&nbsp;&nbsp;唱响红歌，收获感动</A></TD>
                                  </TR>
                                  <TR>
                                    <TD class=listbg2 vAlign=top width=10><IMG alt=推荐文章 
                        src="SkinIndex/article_elite.gif"></TD>
                                    <TD class=listbg2>[<A class=listA 
                        href="http://www.zz6789.com/Article/F/F4/Index.html">中考资料</A>]<A 
                        class=listA title=专题十&nbsp;&nbsp;纪念改革开放30周年 
                        href="http://www.zz6789.com/Article/F/F4/200806/Article_20080615103114.html" 
                        target=_blank>专题十&nbsp;&nbsp;纪念改革开放30周年</A></TD>
                                  </TR>
                                  <TR></TR>
                                </TBODY>
                              </TABLE></TD>
                          </TR>
                        </TBODY>
                      </TABLE>
                    </DIV></TD>
                </TR>
                <TR>
                  <TD height=17><IMG height=17 src="SkinIndex/zin_r50_c1.gif" 
            width=287 border=0></TD>
                  <TD width=3>　</TD>
                  <TD height=17><IMG height=17 src="SkinIndex/zin_r50_c1.gif" 
            width=287 border=0></TD>
                </TR>
              </TBODY>
            </TABLE>
            <TABLE id=table8 cellSpacing=0 cellPadding=0 width=577 border=0>
              <TBODY>
                <TR>
                  <TD width=287 background=SkinIndex/zin_r48_c1.gif 
            height=48><B>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;素质测试题</B></TD>
                  <TD width=3>　</TD>
                  <TD width=287 background=SkinIndex/zin_r48_c1.gif 
            height=48><B>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;教学研究</B></TD>
                </TR>
                <TR>
                  <TD background=SkinIndex/zin_r49_c1.gif height=200><DIV align=center>
                      <TABLE id=table12 height=27 cellSpacing=0 cellPadding=0 width="95%" 
            border=0>
                        <TBODY>
                          <TR>
                            <!--文章1-->
                            <TD><TABLE cellSpacing=0 cellPadding=0 width="100%">
                                <TBODY>
                                  <TR>
                                    <TD class=listbg vAlign=top width=10><IMG alt=普通文章 
                        src="SkinIndex/article_common.gif"></TD>
                                    <TD class=listbg>[<A class=listA 
                        href="http://www.zz6789.com/Article/C/C4/Index.html">高一级</A>]<A 
                        class=listA title=高中会考&nbsp;第3框&nbsp;商品的价值量&nbsp;学案 
                        href="http://www.zz6789.com/Article/C/C4/200809/Article_20080909204605.html" 
                        target=_blank>高中会考&nbsp;第3框&nbsp;商品的价值量&nbsp;学</A></TD>
                                  </TR>
                                  <TR>
                                    <TD class=listbg2 vAlign=top width=10><IMG alt=普通文章 
                        src="SkinIndex/article_common.gif"></TD>
                                    <TD class=listbg2>[<A class=listA 
                        href="http://www.zz6789.com/Article/C/C4/Index.html">高一级</A>]<A 
                        class=listA title=会考-经济常识第1、2框学案 
                        href="http://www.zz6789.com/Article/C/C4/200809/Article_20080909204139.html" 
                        target=_blank>会考-经济常识第1、2框学案</A></TD>
                                  </TR>
                                  <TR>
                                    <TD class=listbg vAlign=top width=10><IMG alt=普通文章 
                        src="SkinIndex/article_common.gif"></TD>
                                    <TD class=listbg>[<A class=listA 
                        href="http://www.zz6789.com/Article/C/C1/Index.html">初一级</A>]<A 
                        class=listA title=粤教版思想品德八年级（上册）期中试卷 
                        href="http://www.zz6789.com/Article/C/C1/200809/Article_20080908105626.html" 
                        target=_blank>粤教版思想品德八年级（上册）期</A></TD>
                                  </TR>
                                  <TR>
                                    <TD class=listbg2 vAlign=top width=10><IMG alt=普通文章 
                        src="SkinIndex/article_common.gif"></TD>
                                    <TD class=listbg2>[<A class=listA 
                        href="http://www.zz6789.com/Article/C/C3/Index.html">初三级</A>]<A 
                        class=listA title=江苏省兴洪中学九年级第一单元思想品德测试 
                        href="http://www.zz6789.com/Article/C/C3/200809/Article_20080905081954.html" 
                        target=_blank>江苏省兴洪中学九年级第一单元思</A></TD>
                                  </TR>
                                  <TR>
                                    <TD class=listbg vAlign=top width=10><IMG alt=普通文章 
                        src="SkinIndex/article_common.gif"></TD>
                                    <TD class=listbg>[<A class=listA 
                        href="http://www.zz6789.com/Article/C/C5/Index.html">高二级</A>]<A 
                        class=listA title=云南省2008年6月高中会考模拟考试 
                        href="http://www.zz6789.com/Article/C/C5/200808/Article_20080831115535.html" 
                        target=_blank>云南省2008年6月高中会考模拟考试</A></TD>
                                  </TR>
                                  <TR>
                                    <TD class=listbg2 vAlign=top width=10><IMG alt=推荐文章 
                        src="SkinIndex/article_elite.gif"></TD>
                                    <TD class=listbg2>[<A class=listA 
                        href="http://www.zz6789.com/Article/C/C3/Index.html">初三级</A>]<A 
                        class=listA title=中考指导 
                        href="http://www.zz6789.com/Article/C/C3/200808/Article_20080819203053.html" 
                        target=_blank>中考指导</A></TD>
                                  </TR>
                                  <TR>
                                    <TD class=listbg vAlign=top width=10><IMG alt=普通文章 
                        src="SkinIndex/article_common.gif"></TD>
                                    <TD class=listbg>[<A class=listA 
                        href="http://www.zz6789.com/Article/C/C3/Index.html">初三级</A>]<A 
                        class=listA title=2007年决战政治中考必背基础知识 
                        href="http://www.zz6789.com/Article/C/C3/200808/Article_20080806223009.html" 
                        target=_blank>2007年决战政治中考必背基础知识</A></TD>
                                  </TR>
                                  <TR>
                                    <TD class=listbg2 vAlign=top width=10><IMG alt=普通文章 
                        src="SkinIndex/article_common.gif"></TD>
                                    <TD class=listbg2>[<A class=listA 
                        href="http://www.zz6789.com/Article/C/C3/Index.html">初三级</A>]<A 
                        class=listA title=望奎三中期中考试政治试题参考答案 
                        href="http://www.zz6789.com/Article/C/C3/200807/Article_20080702203650.html" 
                        target=_blank>望奎三中期中考试政治试题参考答</A></TD>
                                  </TR>
                                  <TR>
                                    <TD class=listbg vAlign=top width=10><IMG alt=普通文章 
                        src="SkinIndex/article_common.gif"></TD>
                                    <TD class=listbg>[<A class=listA 
                        href="http://www.zz6789.com/Article/C/C3/Index.html">初三级</A>]<A 
                        class=listA title=期中考试初三政治试题 
                        href="http://www.zz6789.com/Article/C/C3/200807/Article_20080702203010.html" 
                        target=_blank>期中考试初三政治试题</A></TD>
                                  </TR>
                                  <TR>
                                    <TD class=listbg2 vAlign=top width=10><IMG alt=普通文章 
                        src="SkinIndex/article_common.gif"></TD>
                                    <TD class=listbg2>[<A class=listA 
                        href="http://www.zz6789.com/Article/C/C3/Index.html">初三级</A>]<A 
                        class=listA title=期中考试初三政治试题 
                        href="http://www.zz6789.com/Article/C/C3/200807/Article_20080702202743.html" 
                        target=_blank>期中考试初三政治试题</A></TD>
                                  </TR>
                                  <TR></TR>
                                </TBODY>
                              </TABLE></TD>
                          </TR>
                        </TBODY>
                      </TABLE>
                    </DIV></TD>
                  <TD width=3>　</TD>
                  <TD background=SkinIndex/zin_r49_c1.gif height=200><DIV align=center>
                      <TABLE id=table12 height=27 cellSpacing=0 cellPadding=0 width="95%" 
            border=0>
                        <TBODY>
                          <TR>
                            <!--文章1-->
                            <TD><TABLE cellSpacing=0 cellPadding=0 width="100%">
                                <TBODY>
                                  <TR>
                                    <TD class=listbg vAlign=top width=10><IMG alt=普通文章 
                        src="SkinIndex/article_common.gif"></TD>
                                    <TD class=listbg>[<A class=listA 
                        href="http://www.zz6789.com/Article/H/H1/Index.html">教学心得</A>]<A 
                        class=listA title=浅谈对学生的批评教育 
                        href="http://www.zz6789.com/Article/H/H1/200809/Article_20080908211924.html" 
                        target=_blank>浅谈对学生的批评教育</A></TD>
                                  </TR>
                                  <TR>
                                    <TD class=listbg2 vAlign=top width=10><IMG alt=普通文章 
                        src="SkinIndex/article_common.gif"></TD>
                                    <TD class=listbg2>[<A class=listA 
                        href="http://www.zz6789.com/Article/H/H1/Index.html">教学心得</A>]<A 
                        class=listA title=让初中思想品德课教学生活化 
                        href="http://www.zz6789.com/Article/H/H1/200809/Article_20080905142959.html" 
                        target=_blank>让初中思想品德课教学生活化</A></TD>
                                  </TR>
                                  <TR>
                                    <TD class=listbg vAlign=top width=10><IMG alt=普通文章 
                        src="SkinIndex/article_common.gif"></TD>
                                    <TD class=listbg>[<A class=listA 
                        href="http://www.zz6789.com/Article/H/H1/Index.html">教学心得</A>]<A 
                        class=listA title=错误的价值 
                        href="http://www.zz6789.com/Article/H/H1/200808/Article_20080803105903.html" 
                        target=_blank>错误的价值</A></TD>
                                  </TR>
                                  <TR>
                                    <TD class=listbg2 vAlign=top width=10><IMG alt=普通文章 
                        src="SkinIndex/article_common.gif"></TD>
                                    <TD class=listbg2>[<A class=listA 
                        href="http://www.zz6789.com/Article/H/H1/Index.html">教学心得</A>]<A 
                        class=listA title=浅谈思想政治课中的情感培养 
                        href="http://www.zz6789.com/Article/H/H1/200807/Article_20080730230051.html" 
                        target=_blank>浅谈思想政治课中的情感培养</A></TD>
                                  </TR>
                                  <TR>
                                    <TD class=listbg vAlign=top width=10><IMG alt=普通文章 
                        src="SkinIndex/article_common.gif"></TD>
                                    <TD class=listbg>[<A class=listA 
                        href="http://www.zz6789.com/Article/H/H1/Index.html">教学心得</A>]<A 
                        class=listA title=优化课堂结构，打造高效课堂 
                        href="http://www.zz6789.com/Article/H/H1/200807/Article_20080717154422.html" 
                        target=_blank>优化课堂结构，打造高效课堂</A></TD>
                                  </TR>
                                  <TR>
                                    <TD class=listbg2 vAlign=top width=10><IMG alt=普通文章 
                        src="SkinIndex/article_common.gif"></TD>
                                    <TD class=listbg2>[<A class=listA 
                        href="http://www.zz6789.com/Article/H/H1/Index.html">教学心得</A>]<A 
                        class=listA title=思想品德课堂延伸的形式 
                        href="http://www.zz6789.com/Article/H/H1/200807/Article_20080717145918.html" 
                        target=_blank>思想品德课堂延伸的形式</A></TD>
                                  </TR>
                                  <TR>
                                    <TD class=listbg vAlign=top width=10><IMG alt=普通文章 
                        src="SkinIndex/article_common.gif"></TD>
                                    <TD class=listbg>[<A class=listA 
                        href="http://www.zz6789.com/Article/H/H1/Index.html">教学心得</A>]<A 
                        class=listA title=思想政治课应真正成为思想品德教育课 
                        href="http://www.zz6789.com/Article/H/H1/200805/Article_20080521165423.html" 
                        target=_blank>思想政治课应真正成为思想品德教</A></TD>
                                  </TR>
                                  <TR>
                                    <TD class=listbg2 vAlign=top width=10><IMG alt=普通文章 
                        src="SkinIndex/article_common.gif"></TD>
                                    <TD class=listbg2>[<A class=listA 
                        href="http://www.zz6789.com/Article/H/H3/Index.html">研究成果</A>]<A 
                        class=listA title=中学思想品德课自主教学模式探索 
                        href="http://www.zz6789.com/Article/H/H3/200805/Article_20080518184417.html" 
                        target=_blank>中学思想品德课自主教学模式探索</A></TD>
                                  </TR>
                                  <TR>
                                    <TD class=listbg vAlign=top width=10><IMG alt=普通文章 
                        src="SkinIndex/article_common.gif"></TD>
                                    <TD class=listbg>[<A class=listA 
                        href="http://www.zz6789.com/Article/H/H1/Index.html">教学心得</A>]<A 
                        class=listA title=用激情和思考点燃政治课堂 
                        href="http://www.zz6789.com/Article/H/H1/200805/Article_20080518183930.html" 
                        target=_blank>用激情和思考点燃政治课堂</A></TD>
                                  </TR>
                                  <TR>
                                    <TD class=listbg2 vAlign=top width=10><IMG alt=普通文章 
                        src="SkinIndex/article_common.gif"></TD>
                                    <TD class=listbg2>[<A class=listA 
                        href="http://www.zz6789.com/Article/H/H3/Index.html">研究成果</A>]<A 
                        class=listA title=学习迁移理论及在教学中运用研究 
                        href="http://www.zz6789.com/Article/H/H3/200804/Article_20080420212903.html" 
                        target=_blank>学习迁移理论及在教学中运用研究</A></TD>
                                  </TR>
                                  <TR></TR>
                                </TBODY>
                              </TABLE></TD>
                          </TR>
                        </TBODY>
                      </TABLE>
                    </DIV></TD>
                </TR>
                <TR>
                  <TD height=17><IMG height=17 src="SkinIndex/zin_r50_c1.gif" 
            width=287 border=0></TD>
                  <TD width=3>　</TD>
                  <TD height=17><IMG height=17 src="SkinIndex/zin_r50_c1.gif" 
            width=287 border=0></TD>
                </TR>
              </TBODY>
            </TABLE></TD>
          <TD width=8>　</TD>
        </TR>
      </TBODY>
    </TABLE>
  </DIV>
  <DIV align=center>
    <TABLE id=table7 cellSpacing=0 cellPadding=0 width=797 
background=SkinIndex/new_wow_43.jpg border=0>
      <TBODY>
        <TR>
          <TD width=8 rowSpan=3>　</TD>
          <TD width=781 bgColor=#ffffff height=22><IMG height=21 
      src="SkinIndex/zin_r39_c2.gif" width=781 border=0></TD>
          <TD width=8 rowSpan=3>　</TD>
        </TR>
        <TR>
          <TD align=middle width=781 background=SkinIndex/zin_r42_c12.gif 
    bgColor=#ffffff height=60><TABLE cellSpacing=5 cellPadding=0 width="100%" align=center border=0>
              <TBODY>
                <TR vAlign=top>
                  <TD align=middle><A class="" 
            title="图片名称：1QW&#13;&#10;作&nbsp;&nbsp;&nbsp;&nbsp;者：letian1971&#13;&#10;更新时间：2008-8-4 15:58:20" 
            href="http://www.zz6789.com/Photo/Y/200808/Photo_20080804155820.html" 
            target=_blank><IMG class=pic3 height=90 
            src="SkinIndex/20080804155817692.jpg" width=110 
            border=0></A><BR>
                    <A class="" 
            title="图片名称：1QW&#13;&#10;作&nbsp;&nbsp;&nbsp;&nbsp;者：letian1971&#13;&#10;更新时间：2008-8-4 15:58:20" 
            href="http://www.zz6789.com/Photo/Y/200808/Photo_20080804155820.html" 
            target=_blank>1QW</A></TD>
                  <TD align=middle><A class="" 
            title="图片名称：123&#13;&#10;作&nbsp;&nbsp;&nbsp;&nbsp;者：佚名&#13;&#10;更新时间：2008-8-4 15:56:32" 
            href="http://www.zz6789.com/Photo/S/200808/Photo_20080804155632.html" 
            target=_blank><IMG class=pic3 height=90 
            src="SkinIndex/20080804155628425.jpg" width=110 
            border=0></A><BR>
                    <A class="" 
            title="图片名称：123&#13;&#10;作&nbsp;&nbsp;&nbsp;&nbsp;者：佚名&#13;&#10;更新时间：2008-8-4 15:56:32" 
            href="http://www.zz6789.com/Photo/S/200808/Photo_20080804155632.html" 
            target=_blank>123</A></TD>
                  <TD align=middle><A class="" 
            title="图片名称：123&#13;&#10;作&nbsp;&nbsp;&nbsp;&nbsp;者：佚名&#13;&#10;更新时间：2008-8-4 15:54:51" 
            href="http://www.zz6789.com/Photo/S/200808/Photo_20080804155451.html" 
            target=_blank><IMG class=pic3 height=90 
            src="SkinIndex/20080804155335962.jpg" width=110 
            border=0></A><BR>
                    <A class="" 
            title="图片名称：123&#13;&#10;作&nbsp;&nbsp;&nbsp;&nbsp;者：佚名&#13;&#10;更新时间：2008-8-4 15:54:51" 
            href="http://www.zz6789.com/Photo/S/200808/Photo_20080804155451.html" 
            target=_blank>123</A></TD>
                  <TD align=middle><A class="" 
            title="图片名称：感人的坚强！灾区志愿者拍到的一幕(组图)&#13;&#10;作&nbsp;&nbsp;&nbsp;&nbsp;者：新华网(北京)&#13;&#10;更新时间：2008-6-26 8:47:21" 
            href="http://www.zz6789.com/Photo/S/200806/Photo_20080626084721.html" 
            target=_blank><IMG class=pic3 height=90 
            src="SkinIndex/20080626084751345.jpg" width=110 
            border=0></A><BR>
                    <A class="" 
            title="图片名称：感人的坚强！灾区志愿者拍到的一幕(组图)&#13;&#10;作&nbsp;&nbsp;&nbsp;&nbsp;者：新华网(北京)&#13;&#10;更新时间：2008-6-26 8:47:21" 
            href="http://www.zz6789.com/Photo/S/200806/Photo_20080626084721.html" 
            target=_blank>感人的坚强！灾区</A></TD>
                  <TD align=middle><A class="" 
            title="图片名称：麻雀虽小，亲情一点不少(图)&#13;&#10;作&nbsp;&nbsp;&nbsp;&nbsp;者：宋峤&#13;&#10;更新时间：2008-6-26 8:43:52" 
            href="http://www.zz6789.com/Photo/Q/200806/Photo_20080626084352.html" 
            target=_blank><IMG class=pic3 height=90 
            src="SkinIndex/20080626084412185.jpg" width=110 
            border=0></A><BR>
                    <A class="" 
            title="图片名称：麻雀虽小，亲情一点不少(图)&#13;&#10;作&nbsp;&nbsp;&nbsp;&nbsp;者：宋峤&#13;&#10;更新时间：2008-6-26 8:43:52" 
            href="http://www.zz6789.com/Photo/Q/200806/Photo_20080626084352.html" 
            target=_blank>麻雀虽小，亲情一</A></TD>
                  <TD align=middle><A class="" 
            title="图片名称：美丽神奇的乌兰布统草原[组图]&#13;&#10;作&nbsp;&nbsp;&nbsp;&nbsp;者：邹宝良/摄影&#13;&#10;更新时间：2008-6-26 8:38:52" 
            href="http://www.zz6789.com/Photo/V/200806/Photo_20080626083852.html" 
            target=_blank><IMG class=pic3 height=90 
            src="SkinIndex/20080626084005909.jpg" width=110 
            border=0></A><BR>
                    <A class="" 
            title="图片名称：美丽神奇的乌兰布统草原[组图]&#13;&#10;作&nbsp;&nbsp;&nbsp;&nbsp;者：邹宝良/摄影&#13;&#10;更新时间：2008-6-26 8:38:52" 
            href="http://www.zz6789.com/Photo/V/200806/Photo_20080626083852.html" 
            target=_blank>美丽神奇的乌兰布</A></TD>
                </TR>
                <TR vAlign=top>
                  <TD align=middle><A class="" 
            title="图片名称：“6·26禁毒日”：走进戒毒所(组图)&#13;&#10;作&nbsp;&nbsp;&nbsp;&nbsp;者：本报记者 王强 摄影记者 杜海&#13;&#10;更新时间：2008-6-26 8:22:48" 
            href="http://www.zz6789.com/Photo/R/200806/Photo_20080626082248.html" 
            target=_blank><IMG class=pic3 height=90 
            src="SkinIndex/20080626082440715.jpg" width=110 
            border=0></A><BR>
                    <A class="" 
            title="图片名称：“6·26禁毒日”：走进戒毒所(组图)&#13;&#10;作&nbsp;&nbsp;&nbsp;&nbsp;者：本报记者 王强 摄影记者 杜海&#13;&#10;更新时间：2008-6-26 8:22:48" 
            href="http://www.zz6789.com/Photo/R/200806/Photo_20080626082248.html" 
            target=_blank>“6·26禁毒日”：</A></TD>
                  <TD align=middle><A class="" 
            title="图片名称：触目惊心！航拍南方水灾灾情(组图)&#13;&#10;作&nbsp;&nbsp;&nbsp;&nbsp;者：网友汉水&#13;&#10;更新时间：2008-6-26 8:15:17" 
            href="http://www.zz6789.com/Photo/T/200806/Photo_20080626081517.html" 
            target=_blank><IMG class=pic3 height=90 
            src="SkinIndex/20080626081941393.jpg" width=110 
            border=0></A><BR>
                    <A class="" 
            title="图片名称：触目惊心！航拍南方水灾灾情(组图)&#13;&#10;作&nbsp;&nbsp;&nbsp;&nbsp;者：网友汉水&#13;&#10;更新时间：2008-6-26 8:15:17" 
            href="http://www.zz6789.com/Photo/T/200806/Photo_20080626081517.html" 
            target=_blank>触目惊心！航拍南</A></TD>
                  <TD align=middle><A class="" 
            title="图片名称：顾秀莲勉力“最坚强的警花”蒋敏&#13;&#10;作&nbsp;&nbsp;&nbsp;&nbsp;者：中新社发 邹宪 摄&#13;&#10;更新时间：2005-12-30 8:53:19" 
            href="http://www.zz6789.com/Photo/W/200512/Photo_20051230085319.html" 
            target=_blank><IMG class=pic3 height=90 
            src="SkinIndex/20080530090603933.jpg" width=110 
            border=0></A><BR>
                    <A class="" 
            title="图片名称：顾秀莲勉力“最坚强的警花”蒋敏&#13;&#10;作&nbsp;&nbsp;&nbsp;&nbsp;者：中新社发 邹宪 摄&#13;&#10;更新时间：2005-12-30 8:53:19" 
            href="http://www.zz6789.com/Photo/W/200512/Photo_20051230085319.html" 
            target=_blank>顾秀莲勉力“最坚</A></TD>
                  <TD align=middle><A class="" 
            title="图片名称：台湾首富郭台铭出资六百万元认养成都大熊猫(图)&#13;&#10;作&nbsp;&nbsp;&nbsp;&nbsp;者：中新社发 杨菲菲 摄&#13;&#10;更新时间：2008-5-6 10:45:21" 
            href="http://www.zz6789.com/Photo/200805/Photo_20080506104521.html" 
            target=_blank><IMG class=pic3 height=90 
            src="SkinIndex/20080506104542644.jpg" width=110 
            border=0></A><BR>
                    <A class="" 
            title="图片名称：台湾首富郭台铭出资六百万元认养成都大熊猫(图)&#13;&#10;作&nbsp;&nbsp;&nbsp;&nbsp;者：中新社发 杨菲菲 摄&#13;&#10;更新时间：2008-5-6 10:45:21" 
            href="http://www.zz6789.com/Photo/200805/Photo_20080506104521.html" 
            target=_blank>台湾首富郭台铭出</A></TD>
                  <TD align=middle><A class="" 
            title="图片名称：连战参观访问三峡工程(组图)&#13;&#10;作&nbsp;&nbsp;&nbsp;&nbsp;者：中新网(北京)艾启平 摄&#13;&#10;更新时间：2008-5-6 10:10:24" 
            href="http://www.zz6789.com/Photo/U/200805/Photo_20080506101024.html" 
            target=_blank><IMG class=pic3 height=90 
            src="SkinIndex/20080506102938436.jpg" width=110 
            border=0></A><BR>
                    <A class="" 
            title="图片名称：连战参观访问三峡工程(组图)&#13;&#10;作&nbsp;&nbsp;&nbsp;&nbsp;者：中新网(北京)艾启平 摄&#13;&#10;更新时间：2008-5-6 10:10:24" 
            href="http://www.zz6789.com/Photo/U/200805/Photo_20080506101024.html" 
            target=_blank>连战参观访问三峡</A></TD>
                  <TD align=middle><A class="" 
            title="图片名称：广州市纪委在五一前发出廉政短信&#13;&#10;作&nbsp;&nbsp;&nbsp;&nbsp;者：大洋网(广州)&#13;&#10;更新时间：2008-5-6 9:39:56" 
            href="http://www.zz6789.com/Photo/U/200805/Photo_20080506093956.html" 
            target=_blank><IMG class=pic3 height=90 
            src="SkinIndex/20080506100822292.jpg" width=110 
            border=0></A><BR>
                    <A class="" 
            title="图片名称：广州市纪委在五一前发出廉政短信&#13;&#10;作&nbsp;&nbsp;&nbsp;&nbsp;者：大洋网(广州)&#13;&#10;更新时间：2008-5-6 9:39:56" 
            href="http://www.zz6789.com/Photo/U/200805/Photo_20080506093956.html" 
            target=_blank>广州市纪委在五一</A></TD>
                </TR>
              </TBODY>
            </TABLE></TD>
        </TR>
        <TR>
          <TD width=781 bgColor=#ffffff height=22><IMG height=19 
      src="SkinIndex/zin_r41_c2.gif" width=781 
border=0></TD>
        </TR>
      </TBODY>
    </TABLE>
  </DIV>
  <DIV align=center>
    <TABLE id=table7 height=100 cellSpacing=0 cellPadding=0 width=797 
background=SkinIndex/new_wow_43.jpg border=0>
      <TBODY>
        <TR>
          <TD width=8>　</TD>
          <TD vAlign=top align=middle width=197 bgColor=#ffffff><TABLE id=table5 height=200 cellSpacing=0 cellPadding=0 width=197 
border=0>
              <TBODY>
                <TR>
                  <TD background=SkinIndex/zdl_8.gif height=35><P align=center><B>本 站 统 计</B></P></TD>
                </TR>
                <TR>
                  <TD height=24><IMG height=24 src="SkinIndex/zin_r11_c1.gif" 
            width=197 border=0></TD>
                </TR>
                <TR>
                  <TD align=middle background=SkinIndex/zin_r13_c1.gif>高中课程：1046 
                    篇文章<BR>
                    文章中心：23721 篇文章<BR>
                    课件中心：2583 个课件<BR>
                    图片中心：6550 张图片<BR>
                    常用软件：0 
                    个软件<BR>
                    注册会员：156394位<BR>
                    <SCRIPT src="SkinIndex/CounterLink.htm"></SCRIPT>
                  </TD>
                </TR>
                <TR>
                  <TD height=24><IMG height=25 src="SkinIndex/zin_r18_c1.gif" 
            width=197 border=0></TD>
                </TR>
              </TBODY>
            </TABLE></TD>
          <TD vAlign=top align=middle width=584 bgColor=#ffffff><TABLE id=table8 cellSpacing=0 cellPadding=0 width=577 border=0>
              <TBODY>
                <TR>
                  <TD width=287 background=SkinIndex/zin_r48_c1.gif 
            height=48><B>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;推荐课件</B></TD>
                  <TD width=3>　</TD>
                  <TD width=287 background=SkinIndex/zin_r48_c1.gif 
            height=48><B>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;热点课件</B></TD>
                </TR>
                <TR>
                  <TD background=SkinIndex/zin_r49_c1.gif height=200><DIV align=center>
                      <TABLE id=table12 height=27 cellSpacing=0 cellPadding=0 width="95%" 
            border=0>
                        <TBODY>
                          <TR>
                            <!--推荐课件-->
                            <TD><TABLE cellSpacing=0 cellPadding=0 width="100%">
                                <TBODY>
                                  <TR>
                                    <TD class=listbg>[<A class=listA 
                        href="http://www.zz6789.com/Soft/Highclass/Politics/Levy05/Index.html">第五课 
                                      我国的人民代表大会制度</A>]<A class=listA title=人民代表大会制度 
                        href="http://www.zz6789.com/Soft/Highclass/Politics/Levy05/200807/Soft_3824.html" 
                        target=_blank>人民代表大会制度</A></TD>
                                  </TR>
                                  <TR>
                                    <TD class=listbg2>[<A class=listA 
                        href="http://www.zz6789.com/Soft/JuniorClass/c3/Index.html">初三级</A>]<A 
                        class=listA title=08专题复习 
                        href="http://www.zz6789.com/Soft/JuniorClass/c3/200806/Soft_3822.html" 
                        target=_blank>08专题复习</A></TD>
                                  </TR>
                                  <TR>
                                    <TD class=listbg>[<A class=listA 
                        href="http://www.zz6789.com/Soft/JuniorClass/c2/Index.html">初二级</A>]<A 
                        class=listA title=08专题复习 
                        href="http://www.zz6789.com/Soft/JuniorClass/c2/200806/Soft_3821.html" 
                        target=_blank>08专题复习</A></TD>
                                  </TR>
                                  <TR>
                                    <TD class=listbg2>[<A class=listA 
                        href="http://www.zz6789.com/Soft/Highclass/Politics/Levy03/Index.html">第三课 
                                      我国政府是人民的政府</A>]<A class=listA title=国旗为地震死难的平民而降 
                        href="http://www.zz6789.com/Soft/Highclass/Politics/Levy03/200805/Soft_3817.html" 
                        target=_blank>国旗为地震死难的平民而降</A></TD>
                                  </TR>
                                  <TR>
                                    <TD class=listbg>[<A class=listA 
                        href="http://www.zz6789.com/Soft/Highclass/Organize/Index.html">国家与国际组织</A>]<A 
                        class=listA title=国家的本质 
                        href="http://www.zz6789.com/Soft/Highclass/Organize/200805/Soft_3813.html" 
                        target=_blank>国家的本质</A></TD>
                                  </TR>
                                  <TR>
                                    <TD class=listbg2>[<A class=listA 
                        href="http://www.zz6789.com/Soft/Highclass/Economy/Levy12/Index.html">第十二课 
                                      经济全球化与对外开放</A>]<A class=listA title=积极与合作参与国际经济竞争PPT 
                        href="http://www.zz6789.com/Soft/Highclass/Economy/Levy12/200804/Soft_3807.html" 
                        target=_blank>积极与合作参与国际经济竞争PPT</A></TD>
                                  </TR>
                                  <TR>
                                    <TD class=listbg>[<A class=listA 
                        href="http://www.zz6789.com/Soft/N/Index.html">德育教育课件</A>]<A 
                        class=listA title=西藏问题 
                        href="http://www.zz6789.com/Soft/N/200804/Soft_3802.html" 
                        target=_blank>西藏问题</A></TD>
                                  </TR>
                                  <TR>
                                    <TD class=listbg2>[<A class=listA 
                        href="http://www.zz6789.com/Soft/M/Index.html">复习专题课件</A>]<A 
                        class=listA title=高中政治主观题答题技巧 
                        href="http://www.zz6789.com/Soft/M/200804/Soft_3799.html" 
                        target=_blank>高中政治主观题答题技巧</A></TD>
                                  </TR>
                                  <TR>
                                    <TD class=listbg>[<A class=listA 
                        href="http://www.zz6789.com/Soft/JuniorClass/c2/Index.html">初二级</A>]<A 
                        class=listA title=法不可违 
                        href="http://www.zz6789.com/Soft/JuniorClass/c2/200804/Soft_3797.html" 
                        target=_blank>法不可违</A></TD>
                                  </TR>
                                  <TR>
                                    <TD class=listbg2>[<A class=listA 
                        href="http://www.zz6789.com/Soft/M/Index.html">复习专题课件</A>]<A 
                        class=listA title=党的十七大报告解读 
                        href="http://www.zz6789.com/Soft/M/200803/Soft_3794.html" 
                        target=_blank>党的十七大报告解读</A></TD>
                                  </TR>
                                  <TR></TR>
                                </TBODY>
                              </TABLE></TD>
                          </TR>
                        </TBODY>
                      </TABLE>
                    </DIV></TD>
                  <TD width=3>　</TD>
                  <TD background=SkinIndex/zin_r49_c1.gif height=200><DIV align=center>
                      <TABLE id=table12 height=27 cellSpacing=0 cellPadding=0 width="95%" 
            border=0>
                        <TBODY>
                          <TR>
                            <!--热点课件-->
                            <TD><TABLE cellSpacing=0 cellPadding=0 width="100%">
                                <TBODY>
                                  <TR>
                                    <TD class=listbg>[<A class=listA 
                        href="http://www.zz6789.com/Soft/N/Index.html">德育教育课件</A>]<A 
                        class=listA title=地震 
                        href="http://www.zz6789.com/Soft/N/200808/Soft_3828.html" 
                        target=_blank>地震</A></TD>
                                  </TR>
                                  <TR>
                                    <TD class=listbg2>[<A class=listA 
                        href="http://www.zz6789.com/Soft/JuniorClass/c2/Index.html">初二级</A>]<A 
                        class=listA title=平等待人（第二课时） 
                        href="http://www.zz6789.com/Soft/JuniorClass/c2/200807/Soft_3827.html" 
                        target=_blank>平等待人（第二课时）</A></TD>
                                  </TR>
                                  <TR>
                                    <TD class=listbg>[<A class=listA 
                        href="http://www.zz6789.com/Soft/JuniorClass/c2/Index.html">初二级</A>]<A 
                        class=listA title=平等待人（第一课时） 
                        href="http://www.zz6789.com/Soft/JuniorClass/c2/200807/Soft_3826.html" 
                        target=_blank>平等待人（第一课时）</A></TD>
                                  </TR>
                                  <TR>
                                    <TD class=listbg2>[<A class=listA 
                        href="http://www.zz6789.com/Soft/JuniorClass/c2/Index.html">初二级</A>]<A 
                        class=listA title=自尊自信 
                        href="http://www.zz6789.com/Soft/JuniorClass/c2/200807/Soft_3825.html" 
                        target=_blank>自尊自信</A></TD>
                                  </TR>
                                  <TR>
                                    <TD class=listbg>[<A class=listA 
                        href="http://www.zz6789.com/Soft/Highclass/Politics/Levy05/Index.html">第五课 
                                      我国的人民代表大会制度</A>]<A class=listA title=人民代表大会制度 
                        href="http://www.zz6789.com/Soft/Highclass/Politics/Levy05/200807/Soft_3824.html" 
                        target=_blank>人民代表大会制度</A></TD>
                                  </TR>
                                  <TR>
                                    <TD class=listbg2>[<A class=listA 
                        href="http://www.zz6789.com/Soft/Highclass/Politics/Levy08/Index.html">第八课 
                                      走近国际社会</A>]<A class=listA title=联合国 
                        href="http://www.zz6789.com/Soft/Highclass/Politics/Levy08/200806/Soft_3823.html" 
                        target=_blank>联合国</A></TD>
                                  </TR>
                                  <TR>
                                    <TD class=listbg>[<A class=listA 
                        href="http://www.zz6789.com/Soft/JuniorClass/c3/Index.html">初三级</A>]<A 
                        class=listA title=08专题复习 
                        href="http://www.zz6789.com/Soft/JuniorClass/c3/200806/Soft_3822.html" 
                        target=_blank>08专题复习</A></TD>
                                  </TR>
                                  <TR>
                                    <TD class=listbg2>[<A class=listA 
                        href="http://www.zz6789.com/Soft/JuniorClass/c2/Index.html">初二级</A>]<A 
                        class=listA title=08专题复习 
                        href="http://www.zz6789.com/Soft/JuniorClass/c2/200806/Soft_3821.html" 
                        target=_blank>08专题复习</A></TD>
                                  </TR>
                                  <TR>
                                    <TD class=listbg>[<A class=listA 
                        href="http://www.zz6789.com/Soft/JuniorClass/c1/Index.html">初一级</A>]<A 
                        class=listA title=网络的诱惑 
                        href="http://www.zz6789.com/Soft/JuniorClass/c1/200806/Soft_3820.html" 
                        target=_blank>网络的诱惑</A></TD>
                                  </TR>
                                  <TR>
                                    <TD class=listbg2>[<A class=listA 
                        href="http://www.zz6789.com/Soft/JuniorClass/c1/Index.html">初一级</A>]<A 
                        class=listA title=主动控制情绪-做情绪的主人 
                        href="http://www.zz6789.com/Soft/JuniorClass/c1/200806/Soft_3819.html" 
                        target=_blank>主动控制情绪-做情绪的主人</A></TD>
                                  </TR>
                                  <TR></TR>
                                </TBODY>
                              </TABLE></TD>
                          </TR>
                        </TBODY>
                      </TABLE>
                    </DIV></TD>
                </TR>
                <TR>
                  <TD height=17><IMG height=17 src="SkinIndex/zin_r50_c1.gif" 
            width=287 border=0></TD>
                  <TD width=3>　</TD>
                  <TD height=17><IMG height=17 src="SkinIndex/zin_r50_c1.gif" 
            width=287 border=0></TD>
                </TR>
              </TBODY>
            </TABLE></TD>
          <TD width=8>　</TD>
        </TR>
      </TBODY>
    </TABLE>
  </DIV>
  <DIV align=center>
    <TABLE id=table7 height=80 cellSpacing=0 cellPadding=0 width=797 
background=SkinIndex/new_wow_43.jpg border=0>
      <TBODY>
        <TR>
          <TD width=8>　</TD>
          <TD align=middle width=781 bgColor=#ffffff height=80><SCRIPT language=javascript src="SkinIndex/1.js"></SCRIPT>
          </TD>
          <TD width=8>　</TD>
        </TR>
      </TBODY>
    </TABLE>
  </DIV>
  <DIV align=center>
    <TABLE id=table7 height=100 cellSpacing=0 cellPadding=0 width=797 
background=SkinIndex/new_wow_43.jpg border=0>
      <TBODY>
        <TR>
          <TD width=8 rowSpan=2>　</TD>
          <TD width=781 bgColor=#c7b883 height=20><P align=center><B>|<SPAN style="CURSOR: hand" 
      onclick="var strHref=window.location.href;this.style.behavior='url(#default#homepage)';this.setHomePage('http://www.zz6789.com');">设为首页</SPAN> | <SPAN title=中学思想政治教学网 style="CURSOR: hand" 
      onclick="window.external.addFavorite('http://www.zz6789.com','中学思想政治教学网')">收藏本站</SPAN> | <A class=Bottom href="mailto:16350310@QQ.COM">联系站长</A> | <A class=Bottom 
      href="http://www.zz6789.com/FriendSite/Index.asp" target=_blank>友情链接</A> | <A class=Bottom href="http://www.zz6789.com/Copyright.asp" 
      target=_blank>版权申明</A> | <A class=Bottom 
      href="http://www.zz6789.com/Adminzz/Admin_Index.asp" 
      target=_blank>管理登录</A>&nbsp;|&nbsp;</B></P></TD>
          <TD width=8 rowSpan=2>　</TD>
        </TR>
        <TR>
          <TD align=middle width=781 
      bgColor=#ffffff>本网由中学思想政治教学网研究室主办、维护<BR>
            粤ICP备05017562号<BR>
            版权所有Copyright(C)2000-2006 </TD>
        </TR>
      </TBODY>
    </TABLE>
  </DIV>
  <DIV align=center>
    <TABLE id=table7 height=50 cellSpacing=0 cellPadding=0 width=797 border=0>
      <TBODY>
        <TR>
          <TD width=8>　</TD>
          <TD width=781>　</TD>
          <TD width=8>　</TD>
        </TR>
      </TBODY>
    </TABLE>
  </DIV>
  <SCRIPT language=Javascript src=""></SCRIPT>
</DIV>
</BODY>
</HTML>
