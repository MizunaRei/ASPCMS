<!--#include file="inc/syscode_article.asp"-->
<%
'请勿改动下面这三行代码
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
<title><%=strPageTitle%></title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<!--#include file="inc/Skin_CSS.asp"-->
<%call MenuJS()%>
</head>
<body <%=Body_Label%> onmousemove='HideMenu()'>
<div id=menuDiv style='Z-INDEX: 1000; VISIBILITY: hidden; WIDTH: 1px; POSITION: absolute; HEIGHT: 1px; BACKGROUND-COLOR: #9cc5f8'></div>
<!--#include file="Top.asp" -->
<TABLE cellSpacing=0 cellPadding=0 width=770 align=center bgColor=#ffffff border=0>
  <TBODY>
    <TR> 
      <TD vAlign=top align=middle width=191> <div align="center"> 
          <table width="100%" border="0" align="center">
            <tr> 
              <td><img src="images/zuo1.gif" width="179" height="41"></td>
            </tr>
            <tr> 
              <td> 
                <% call ShowUserLogin() %>
              </td>
            </tr>
          </table>
          <TABLE width=179 border=0 align="center" cellPadding=0 cellSpacing=0>
            <TBODY>
              <TR> 
                <TD><IMG height=35 alt="" src="images/ss_1.gif" width=179></TD>
              </TR>
              <TR> 
                <TD class=s1 align=left bgColor=#00aace> 
                  <% call Showhot(8,16) %>
                </TD>
              </TR>
              <TR> 
                <TD><IMG height=9 alt="" src="images/ss_3.gif" 
        width=179></TD>
              </TR>
            </TBODY>
          </TABLE>
          <TABLE height=3 cellSpacing=0 cellPadding=0 width="100%" border=0>
            <TBODY>
              <TR> 
                <TD></TD>
              </TR>
            </TBODY>
          </TABLE>
          <TABLE width=179 border=0 align="center" cellPadding=0 cellSpacing=0>
            <TBODY>
              <TR> 
                <TD><IMG height=35 alt="" src="images/ss1_3.gif" 
width=179></TD>
              </TR>
              <TR> 
                <TD bgColor=#949693><SPAN class=s1> 
                  <% call ShowSpecial(5) %>
                  </SPAN></TD>
              </TR>
              <TR> 
                <TD><IMG 
            height=34 alt="" src="images/ss1_5.gif" width=179 
          border=0></TD>
              </TR>
              <TR> 
                <TD bgColor=#949693><SPAN class=s1> 
                  <% call ShownewUser(5) %>
                  </SPAN></TD>
              </TR>
              <TR> 
                <TD><IMG height=8 alt="" src="images/ss1_9.gif" 
        width=179></TD>
              </TR>
            </TBODY>
          </TABLE>
          <TABLE width=179 border=0 align="center" cellPadding=0 cellSpacing=0>
            <TBODY>
         
              <TR> 
                <TD><IMG 
            height=34 alt="" src="images/vote.gif" width=179 
          border=0></TD>
              </TR>
              <TR> 
                <TD bgColor=#949693><div align="center"><SPAN class=s1> 
                    <% call showvote() %>
                    </SPAN></div></TD>
              </TR>
              <TR> 
                <TD><IMG height=8 alt="" src="images/ss1_9.gif" 
        width=179></TD>
              </TR>
            </TBODY>
          </TABLE>
        </div></TD>
      <TD width=3 bgColor=#d7d7d7></TD>
      <TD width=5 bgColor=#979490></TD>
      <TD vAlign=top> <TABLE height=3 cellSpacing=0 cellPadding=0 width="100%" border=0>
          <TBODY>
            <TR> 
              <TD bgColor=#949694></TD>
            </TR>
          </TBODY>
        </TABLE>
        <TABLE height=2 cellSpacing=0 cellPadding=0 width="100%" bgColor=#949694 
      border=0>
          <TBODY>
            <TR> 
              <TD></TD>
            </TR>
          </TBODY>
        </TABLE>
        <TABLE cellSpacing=0 cellPadding=0 width="100%" border=0>
          <TBODY>
            <TR> 
              <TD vAlign=top width=379> <TABLE width="98%" border=0 align="center" cellSpacing=0>
                  <TBODY>
                    <TR> 
                      <TD width="50%"><a href="Article_Class2.asp?ClassID=1"><IMG 
                  height=33 src="images/wenxie-48.gif" width=145 
                  border=0></a></TD>
                      <TD vAlign=bottom width="50%"><DIV align=right><a href="Article_Class2.asp?ClassID=1"><IMG 
                  height=9 src="images/more.gif" width=33 
                  border=0></a></DIV></TD>
                    </TR>
                    <TR> 
                      <TD colSpan=2 height=6><IMG height=6 
                  src="images/wenxie-44.gif" width=376></TD>
                    </TR>
                    <TR> 
                      <TD colSpan=2><!--代码开始-->
<script language='JavaScript' type='text/JavaScript' src='article_js.asp?ClassID=1&IncludeChild=true&SpecialID=&ArticleNum=6&ShowType=1&ShowCols=1&ShowProperty=true&ShowClassName=false&ShowIncludePic=false&ShowTitle=true&ShowUpdateTime=false&ShowHits=false&ShowAuthor=true&ShowHot=false&ShowMore=false&TitleMaxLen=30&ContentMaxLen=200&Hot=false&Elite=false&DateNum=&OrderField=UpdateTime&OrderType=desc'></script>
<!--代码结束--></TD>
                    </TR>
                  </TBODY>
                </TABLE>
                <TABLE width="98%" border=0 align="center" cellSpacing=0>
                  <TBODY>
                    <TR> 
                      <TD width="50%"><a href="Article_Class2.asp?ClassID=2"><IMG src="NewImages/LiLunDongTai.png" alt="理论动态" width=33 
                  height=33 
                  border=0>理论动态</a></TD>
                      <TD vAlign=bottom width="50%"> <DIV align=right><a href="Article_Class2.asp?ClassID=2"><IMG 
                  height=9 src="images/more.gif" width=33 
                  border=0></a></DIV></TD>
                    </TR>
                    <TR> 
                      <TD colSpan=2 height=6><SPAN class=a3><IMG height=6 
                  src="images/wenxie-44.gif" width=376></SPAN></TD>
                    </TR>
                    <TR> 
                      <TD colSpan=2><!--代码开始-->
<script language='JavaScript' type='text/JavaScript' src='article_js.asp?ClassID=2&IncludeChild=true&SpecialID=&ArticleNum=6&ShowType=1&ShowCols=1&ShowProperty=true&ShowClassName=false&ShowIncludePic=false&ShowTitle=true&ShowUpdateTime=false&ShowHits=false&ShowAuthor=false&ShowHot=false&ShowMore=false&TitleMaxLen=30&ContentMaxLen=200&Hot=false&Elite=false&DateNum=&OrderField=UpdateTime&OrderType=desc'></script>
<!--代码结束--></TD>
                    </TR>
                  </TBODY>
                </TABLE>
                <TABLE width="98%" border=0 align="center" cellSpacing=0>
                  <TBODY>
                    <TR> 
                      <TD width="50%"><a href="Article_Class2.asp?ClassID=3"><IMG 
                  height=33 src="images/wenxie-51.gif" width=145 
                  border=0></a></TD>
                      <TD vAlign=bottom width="50%"> <DIV align=right><a href="Article_Class2.asp?ClassID=3"><IMG 
                  height=9 src="images/more.gif" width=33 
                  border=0></a></DIV></TD>
                    </TR>
                    <TR> 
                      <TD colSpan=2 height=6><SPAN class=a3><IMG height=6 
                  src="images/wenxie-44.gif" width=376></SPAN></TD>
                    </TR>
                    <TR> 
                      <TD colSpan=2><!--代码开始-->
<script language='JavaScript' type='text/JavaScript' src='article_js.asp?ClassID=3&IncludeChild=true&SpecialID=&ArticleNum=6&ShowType=1&ShowCols=1&ShowProperty=true&ShowClassName=false&ShowIncludePic=false&ShowTitle=true&ShowUpdateTime=false&ShowHits=false&ShowAuthor=true&ShowHot=false&ShowMore=false&TitleMaxLen=30&ContentMaxLen=200&Hot=false&Elite=false&DateNum=&OrderField=UpdateTime&OrderType=desc'></script>
<!--代码结束--></TD>
                    </TR>
                  </TBODY>
                </TABLE>
                <TABLE width="98%" border=0 align="center" cellSpacing=0>
                  <TBODY>
                    <TR> 
                      <TD width="50%"><a href="Article_Class2.asp?ClassID=58"><IMG src="NewImages/StudentArticleTop.png" alt="学生作品" width=33 
                  height=33 
                  border=0>学生作品</a></TD>
                      <TD vAlign=bottom width="50%"> <DIV align=right><a href="Article_Class2.asp?ClassID=58"><IMG 
                  height=9 src="images/more.gif" width=33 
                  border=0></a></DIV></TD>
                    </TR>
                    <TR> 
                      <TD colSpan=2 height=6><SPAN class=a3><IMG height=6 
                  src="images/wenxie-44.gif" width=376></SPAN></TD>
                    </TR>
                    <TR> 
                      <TD colSpan=2><!--代码开始-->
<script language='JavaScript' type='text/JavaScript' src='article_js.asp?ClassID=58&IncludeChild=true&SpecialID=&ArticleNum=6&ShowType=1&ShowCols=1&ShowProperty=true&ShowClassName=false&ShowIncludePic=false&ShowTitle=true&ShowUpdateTime=false&ShowHits=false&ShowAuthor=false&ShowHot=false&ShowMore=false&TitleMaxLen=30&ContentMaxLen=200&Hot=false&Elite=false&DateNum=&OrderField=UpdateTime&OrderType=desc'></script>
<!--代码结束--></TD>
                    </TR>
                  </TBODY>
                </TABLE></TD>
              <TD width=187 align=middle vAlign=top bgcolor="#f4f1ec"> <TABLE cellSpacing=0 width=181 border=0>
                  <TBODY>
                    <TR> 
                      <TD bgColor=#ffffff height=3></TD>
                    </TR>
                    <TR> 
                      <TD align=middle><SPAN class=a3><IMG height=38 
                  src="images/you1.gif" width=181></SPAN></TD>
                    </TR>
                    <TR> 
                      <TD height=100 vAlign=center bgColor=#f7efde> 
                        <%call ShowAnnounce(1,1)%>
                      </TD>
                    </TR>
                    <TR> 
                      <TD bgColor=#8cdf63><SPAN class=a3><IMG height=23 
                  src="images/wenxie-56.gif" width=102></SPAN></TD>
                    </TR>
                    <TR> 
                      <TD bgColor=#f4f1ec><SPAN class=s1>
                        <% call ShowTopUser(5) %>
                        </SPAN></TD>
                    </TR>
                    <TR> 
                      <TD bgColor=#8cdf63><SPAN class=a3><IMG height=23 
                  src="images/wenxie-57.gif" width=102></SPAN></TD>
                    </TR>
                    <tr> 
                      <td align="center" class=s6> <p><br>
                          <a href="Article_Show.asp?ArticleID=112" target="_blank">审核须知</a> 
                          | <a href="Article_Show.asp?ArticleID=111" target="_blank">投稿须知</a><br>
                          <a href="Article_Show.asp?ArticleID=113" target="_blank">组织章程</a> 
                          | <a href="guestbook.asp" target="_blank">抄袭举报</a><br>
                          <a href='mailto:<%=WebmasterEmail%>' target="_blank">主编信箱</a> 
                          | <a href="userlist.asp" target="_blank">个人文集</a> <br>
                          <br>
                      </td>
                    </tr>
                    <TR> 
                      <TD bgColor=#8cdf63><IMG height=23 
                  src="images/wenxie-59.gif" width=102></TD>
                    </TR>
                    <TR> 
                      <TD bgColor=#f4f1ec height=145> 
                        <% call showGuest(20,10) %>
                      </TD>
                    </TR>
                  </TBODY>
                </TABLE></TD>
            </TR>
          </TBODY>
        </TABLE></TD>
      <TD width=5 bgColor=#979490></TD>
    </TR>
  </TBODY>
</TABLE>
<table width="770" border="0" align="center" bgcolor="979490">
  <TR>
    <TD background=images/xia1.gif height=8></TD></TR>
</table>

<% call bottom() 
call CloseConn()  %>
</body>                         
</html>
