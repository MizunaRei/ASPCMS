<!--#include file="Inc/syscode_article.asp"-->
<%
'请勿改动下面这三行代码
const ChannelID=2
Const ShowRunTime="Yes"
MaxPerPage=20
strFileName="Article_Class.asp?ClassID=" & ClassID & "&SpecialID=" & SpecialID
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
<body <%=Body_Label%> onmousemove='HideMenu()' background="SkinIndex/bg_all.gif" style="BACKGROUND-IMAGE: url(SkinIndex/bg_all.gif)">
<!--#include file="Top.asp"-->
<%
dim sqlRoot,rsRoot,trs,arrClassID,TitleStr
sqlRoot="select C.ClassID,C.ClassName,C.RootID,L.LayoutFileName,L.LayoutID,C.Child,C.ParentPath From ArticleClass C inner join Layout L on C.LayoutID=L.LayoutID where C.ParentID=" & ClassID & " and C.IsElite=True and C.LinkUrl='' and C.BrowsePurview>=" & UserLevel & " order by C.OrderID"
Set rsRoot= Server.CreateObject("ADODB.Recordset")
rsRoot.open sqlRoot,conn,1,1
%>
<table width="781" border="0" align="center" cellpadding="0" cellspacing="0" class="border2" >
  <tr> 
 <!--左边栏-->   <TD width=191 align=left  valign="top" > 
    
      
        
           
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
    <td width="5" bgcolor="#949693"></td>
    <td width="575" valign="top"><table width="100%" border="0" cellpadding="0" cellspacing="0" valign="top" align="top">
        <tr> 
          <td height="393" valign="top"> <%
	if rsRoot.bof and rsRoot.eof then
	%> <table width="98%" border="0" valign="top" align="center" cellpadding="0" cellspacing="5" bgcolor="#F7EFDE" >
              <tr> 
                <td   valign="middle">&nbsp;&nbsp;&nbsp;&nbsp;<%=ClassName%> 文章列表</td>
              </tr>
            </table>
            <table width="98%" border="0" valign="top" align="center" cellpadding="0" cellspacing="5" bgcolor="#FFFFFF">
              <tr> 
                <td height="198" valign="top" background="images/fcbg2.gif"> <%call ShowArticle(30)%> </td>
              </tr>
            </table>
            <%
		  if totalput>0 then
		  	call showpage(strFileName,totalput,MaxPerPage,false,true,"篇文章")
		  end if
		  %> <%
	else
		do while not rsRoot.eof
%> <table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr> 
                <td><table width="98%" border="0" align="center" cellpadding="0" cellspacing="5" bgcolor="#F7EFDE">
                    <tr> 
                      <td>&nbsp;&nbsp;&nbsp;&nbsp; <%
				arrClassID=rsRoot(0)
				response.write "<a href='" & rsRoot(3) & "?ClassID=" & rsRoot(0) & "'>" & rsRoot(1) & "</a>"
				if rsRoot(5)>0 then
					response.write "："
					set trs=conn.execute("select top 4 C.ClassID,C.ClassName,C.RootID,L.LayoutFileName,L.LayoutID From ArticleClass C inner join Layout L on C.LayoutID=L.LayoutID where C.ParentID=" & rsRoot(0) & " and C.IsElite=True and C.LinkUrl=''  and C.BrowsePurview>=" & UserLevel & " order by C.OrderID")
					do while not trs.eof
						response.write "&nbsp;&nbsp;<a href='" & trs(3) & "?ClassID=" & trs(0) & "'>" & trs(1) & "</a>"
						trs.movenext
					loop
					set trs=conn.execute("select ClassID from ArticleClass where ParentID=" & rsRoot(0) & " or ParentPath like '%" & rsRoot(6) & "," & rsRoot(0) & ",%' and Child=0 and LinkUrl='' and BrowsePurview>=" & UserLevel)
					do while not trs.eof
						arrClassID=arrClassID & "," & trs(0)
						trs.movenext
					loop
				end if
				%> </td>
                    </tr>
                  </table></td>
              </tr>
              <tr> 
                <td height="57" align="center"><table bgcolor="#FFFFFF" width="98%" border="0" align="center" cellpadding="0" cellspacing="5">
                    <tr> 
                      <td height="24" colspan="2" valign="top"><table width="99%" border="0" align="center" cellpadding="0" cellspacing="0">
                          <tr> 
                            <td background="images/fcbg2.gif"> <%
sql="select top 5 A.ArticleID,A.ClassID,L.LayoutID,L.LayoutFileName,A.Title,A.Key,A.Author,A.CopyFrom,A.UpdateTime,A.Editor,A.TitleFontColor,A.TitleFontType,"
sql=sql & "A.Hits,A.OnTop,A.Hot,A.Elite,A.Passed,A.IncludePic,A.Stars,A.PaginationType,A.ReadLevel,A.ReadPoint,A.DefaultPicUrl from Article A"
sql=sql & " inner join Layout L on A.LayoutID=L.LayoutID where A.Deleted=False and A.Passed=True and A.ClassID in (" & arrClassID & ")  order by A.ArticleID desc"
rsArticle.open sql,conn,1,1
if rsArticle.bof and  rsArticle.eof then
	response.write "<li>没有任何文章</li>"
else
	call ArticleContent(30,True,True,True,1,True,True)
end if
rsArticle.close
				%> </td>
                          </tr>
                        </table></td>
                    </tr>
                    <tr> 
                      <td width="510" height="18" valign="top" background="images/fcbg2.gif"><div align="right"> 
                        </div></td>
                      <td width="53" valign="top" background="images/fcbg2.gif"> 
                        <%response.write "<a href='" & rsRoot(3) & "?ClassID=" & rsRoot(0) & "'>more...</a>"%> </td>
                    </tr>
                  </table></td>
              </tr>
            </table>
            <%
			rsRoot.movenext
		loop
	end if
	rsRoot.close
	set rsRoot=nothing
%> <table width='98%' border='0' align="center" cellpadding='0' cellspacing='5' bgcolor="#FFFFFF">
              <tr> 
                <td width="19%" height="18"> <div align="center">站内文章搜索</div></td>
                <td width="81%"戀彧敬瑦????	????????> <% call ShowSearchForm("Article_Search.asp",2) %> </td>
              </tr>
            </table></td>
        </tr>
      </table></td>
    <td width="5" valign="top" bgcolor="#949693">&nbsp;</td>
  </tr>
</table>
<!--一条无聊的分隔线<table width="770" border="0" align="center" cellpadding="0" cellspacing="0" class="tdbg">
  <tr> 
    <td  height="13" align="center" valign="top"><table width="770" border="0" align="center" cellpadding="0" cellspacing="0">
        
        <tr> 
          <td height="13" background="images/xia1.gif" ></td>
        </tr>
      </table></td>
  </tr>
</table>-->

<div align="center">
  
    <TABLE id=table7 height=100 cellSpacing=0 cellPadding=0 width=797 
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
          <TD align=center width=781 
      bgColor=#ffffff>本网站由<font color="#3300FF"><a href="http://renwen.university.edu.cn/">university人文社会科学学院</a></font>主办、维护<BR>
            
             </TD>
        </TR>
      </TBODY>
    </TABLE>
  </div>
</body>
</html>
<%
set rsArticle=nothing
set rsPic=nothing
call CloseConn()
%>