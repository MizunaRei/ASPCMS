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
<body <%=Body_Label%> onmousemove='HideMenu()'  style="BACKGROUND-IMAGE: url(SkinIndex/bg_all.gif)"  background="SkinIndex/bg_all.gif" >
<!--#include file="Top.asp"-->
<div align="center">
<table width="781" border="0" align="center" cellpadding="0" cellspacing="0" class="border2">
  <tr> 
  <!--左边栏-->  <TD width=191 align=left  valign="top" > 
    
      
        
           
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
                height=172><% call ShowUserLogin() %>
                      </TD>
                  </TR>
                  <TR>
                    <TD><IMG height=80 src="SkinIndex/zin_r16_c1.gif" 
                  width=197 border=0></TD>
                  </TR>
                </TBODY>
              </TABLE><!--结束用户登录-->
              <!--人气文章--><TABLE id=table5 height=100 cellSpacing=0 cellPadding=0 width=191 border=0>
              <TBODY>
                <TR>
                  <TD background=SkinIndex/zdl_8.gif height=35><P align=center><B>人 气 文 章</B></P></TD>
                </TR>
                <TR>
                  <TD height=24><IMG height=24 src="SkinIndex/zin_r11_c1.gif" 
            width=197 border=0></TD>
                </TR>
                
                
                <TR>
                  <TD  align="center" background=SkinIndex/zin_r13_c1.gif>
                   <% call Showhot(10,16) %>
                    </TD>
                </TR>
                
                <TR>
                  <TD height=24><IMG height=25 src="SkinIndex/zin_r18_c1.gif" 
            width=197 border=0></TD>
                </TR>
              </TBODY>
            </TABLE><!--结束人气文章-->
          
          
           <!--推荐文章--><TABLE id=table5 height=100 cellSpacing=0 cellPadding=0 width=191 
border=0>
              <TBODY>
                <TR>
                  <TD background=SkinIndex/zdl_8.gif height=35><P align=center><B>推 荐 文 章</B></P></TD>
                </TR>
                <TR>
                  <TD height=24><IMG height=24 src="SkinIndex/zin_r11_c1.gif" 
            width=197 border=0></TD>
                </TR>
                
                
                <TR>
                  <TD  align="center" background=SkinIndex/zin_r13_c1.gif>
                   <% call ShowElite(10,16) %>
                    </TD>
                </TR>
                
                <TR>
                  <TD height=24><IMG height=25 src="SkinIndex/zin_r18_c1.gif" 
            width=197 border=0></TD>
                </TR>
              </TBODY>
            </TABLE><!--结束推荐文章-->
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
     </TD><!--结束左边栏-->
    <td width="5" bgcolor="949693"></td>
    <td valign="top">&nbsp;</td>
    <td valign="top"> <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td height="5"></td>
        </tr>
      </table>
      <table width="98%" border="0" align="center" cellpadding="0" cellspacing="0" class="title_main">
        <tr> 
          <td width="40">&nbsp;</td>
          <td><table width="100%" border="0" cellspacing="0" cellpadding="0" >
              <tr> 
                <td class="title_maintxt"> <font color="#0000FF">[课程]</font><%=SpecialName%> 文章列表</td>
              </tr>
            </table></td>
        </tr>
      </table>
      <table width="98%" border="0" align="center" cellpadding="0" cellspacing="0" class="border">
        <tr> 
          <td valign="top"> <table width="100%" height="100%" border="0" cellpadding="0" cellspacing="5">
              <tr> 
                <td height="150" valign="top"> <%call ShowArticle(30)%> </td>
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
    <td valign="top"  width="5" bgcolor="#949693"></td>
  </tr>
</table>
<table width="770" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF" class="tdbg">
  <tr> 
    <td  height="13" align="center" valign="top"><table width="755" border="0" align="center" cellpadding="0" cellspacing="0">
        <tr> 
          <td height="13" Class="tdbg_left2"></td>
        </tr>
      </table></td>
  </tr>
</table>
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
<% call PopAnnouceWindow(400,300) %>
</div>
</body>
</html>
<% call CloseConn() %>