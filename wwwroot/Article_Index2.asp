<!--#include file="Inc/syscode_article.asp"-->
<%
'请勿改动下面这三行代码
const ChannelID=2
Const ShowRunTime="Yes"
MaxPerPage=20
SkinID=0
PageTitle="首页"
Set rsArticle= Server.CreateObject("ADODB.Recordset")
Set rsPic= Server.CreateObject("ADODB.Recordset")
%>
<html>
<head>
<title><%=strPageTitle & " >> " & PageTitle%></title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<!--#include file="inc/Skin_CSS.asp"-->
<%call MenuJS()%>
<script language="JavaScript" type="text/JavaScript">
function refreshMe()
{
    window.refresh;
}
</script>
</head>
<body <%=Body_Label%> onmousemove='HideMenu()' onLoad="javascipt:setTimeout('refreshMe()',1000);">
<!--#include file="top.asp"-->
<table  border="0" align="center" cellpadding="0" cellspacing="0" class="border2" width="760">
    <td width="180" valign="top" align="left" class="tdbg_leftall"> <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td height="5"></td>
        </tr>
      </table>
      <table width="180" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td background="Images/left01.gif"> <table width="100%" border="0" cellpadding="0" cellspacing="0">
              <tr class="title_left" > 
                <td class="title_lefttxt"><div align="center"><strong>用 户 登 录</strong></div></td>
              </tr>
            </table></td>
        </tr>
        <tr> 
          <td valign="top" class="tdbg_left"> <table width="100%" border="0" cellpadding="3">
              <tr> 
                <td> <% call ShowUserLogin() %> </td>
              </tr>
            </table></td>
        </tr>
        <tr> 
          <td class="title_left2"></td>
        </tr>
        <tr> 
          <td background="Images/left03.gif"><table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr class="title_left"> 
                <td class="title_lefttxt"><div align="center"><strong>专 题 栏 目</strong></div></td>
              </tr>
            </table></td>
        </tr>
        <tr> 
          <td height="50" valign="top" class="tdbg_left"> <table width="100%" border="0" cellpadding="8">
              <tr> 
                <td> <% call ShowSpecial(10) %> </td>
              </tr>
            </table></td>
        </tr>
        <tr> 
          <td class="title_left2"></td>
        </tr>
        <tr> 
          <td background="Images/left04.gif"><table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr class="title_left"> 
                <td class="title_lefttxt"><div align="center"><strong>本 站 统 计</strong></div></td>
              </tr>
            </table></td>
        </tr>
        <tr> 
          <td valign="top" class="tdbg_left"> <table width="100%" border="0" cellpadding="8">
              <tr> 
                <td> <% call ShowSiteCount() %> </td>
              </tr>
            </table></td>
        </tr>
        <tr> 
          <td class="title_left2"></td>
        </tr>
        <tr> 
          <td background="Images/left05.gif"><table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr class="title_left"> 
                <td class="title_lefttxt"><div align="center"><strong>用 户 排 行</strong></div></td>
              </tr>
            </table></td>
        </tr>
        <tr> 
          <td valign="top" class="tdbg_left"> <table width="100%" border="0" cellspacing="0" cellpadding="8">
              <tr> 
                <td> <% call ShowTopUser(8) %> </td>
              </tr>
            </table></td>
        </tr>
        <tr> 
          <td class="title_left2"></td>
        </tr>
        <tr> 
          <td background="Images/left06.gif"><table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr class="title_left"> 
                <td class="title_lefttxt"><div align="center"><strong>最 新 调 查</strong></div></td>
              </tr>
            </table></td>
        </tr>
        <tr> 
          <td valign="top" class="tdbg_left"> <table width="100%" border="0" cellpadding="8">
              <tr> 
                <td> <% call ShowVote() %> </td>
              </tr>
            </table></td>
        </tr>
        <tr> 
          <td class="title_left2"></td>
        </tr>
      </table></td>
    <td width="5"></td>

    <td width="410" valign="top">
      <table border="0" cellpadding="0" cellspacing="0" width="410">
<% 
	dim sqlRoot,rsRoot,trs,arrClassID,TitleStr 
	sqlRoot="select C.ClassID,C.ClassName,C.RootID,L.LayoutFileName,L.LayoutID,C.Child,C.Readme From ArticleClass C inner join Layout L on C.LayoutID=L.LayoutID where C.ParentID=0 and C.IsElite=True and C.LinkUrl='' and C.BrowsePurview>=" & UserLevel & " order by C.RootID" 
	Set rsRoot= Server.CreateObject("ADODB.Recordset") 
	rsRoot.open sqlRoot,conn,1,1 
	if rsRoot.bof and rsRoot.eof then  
		response.Write("还没有任何栏目，请首先添加栏目。") 
	else 
		do while not rsRoot.eof 
%> 
        <tr>
          <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr> 
                <td><table width="100%" border="0" cellpadding="0" cellspacing="0" class="title_main">
                    <tr> 
                      <td width="40">&nbsp;</td>
                      <td class="title_maintxt">
                        <%
				arrClassID=rsRoot(0)
				response.write "<a href='" & rsRoot(3) & "?ClassID=" & rsRoot(0) & "' title='" & rsRoot(6) & "'>" & rsRoot(1) & "</a>"
				if rsRoot(5)>0 then
					response.write "："
					set trs=conn.execute("select top 4 C.ClassID,C.ClassName,C.RootID,L.LayoutFileName,C.LinkUrl,C.Readme From ArticleClass C inner join Layout L on C.LayoutID=L.LayoutID where C.ParentID=" & rsRoot(0) & " and C.IsElite=True and C.LinkUrl='' and C.BrowsePurview>=" & UserLevel & " order by C.OrderID")
					do while not trs.eof
						if trs(4)<>"" then
							response.write "&nbsp;&nbsp;<a href='" & trs(4) & "' title='" & trs(5) & "'>" & trs(1) & "</a>"
						else
							response.write "&nbsp;&nbsp;<a href='" & trs(3) & "?ClassID=" & trs(0) & "' title='" & trs(5) & "'>" & trs(1) & "</a>"
						end if
						trs.movenext
					loop
					set trs=conn.execute("select ClassID from ArticleClass where RootID=" & rsRoot(2) & " and Child=0 and LinkUrl='' and BrowsePurview>=" & UserLevel)
					do while not trs.eof
						arrClassID=arrClassID & "," & trs(0)
						trs.movenext
					loop
				end if
				%>
                      </td>
                    </tr>
                  </table></td>
              </tr>
              <tr> 
                <td><table width="100%" border="0" cellpadding="3" cellspacing="0" class="border">
                    <tr> 
                      <td width="135" height="100" align="center" valign="top"> 
                        <%
sql="select top 1 A.ArticleID,A.ClassID,L.LayoutID,L.LayoutFileName,A.Title,A.Key,A.Author,A.CopyFrom,A.UpdateTime,A.Editor,A.TitleFontColor,A.TitleFontType,"
sql=sql & "A.Hits,A.OnTop,A.Hot,A.Elite,A.Passed,A.IncludePic,A.Stars,A.PaginationType,A.ReadLevel,A.ReadPoint,A.DefaultPicUrl from Article A"
sql=sql & " inner join Layout L on A.LayoutID=L.LayoutID where A.Deleted=False and A.Passed=True and A.ClassID in (" & arrClassID & ") and DefaultPicUrl<>'' order by A.OnTop,A.ArticleID desc"
rsPic.open sql,conn,1,1
if rsPic.bof and  rsPic.eof then
	response.write "<img src='images/NoPic.jpg' width=135 height=100 border=0><br>没有任何图片文章"
else
	strPic=""
	call GetPicArticleTitle(20,135,100)
	response.write strPic
end if
rsPic.close
				%> </td>
                      <td width="5">&nbsp;</td>
                      <td valign="top"> 
                        <table width="100%" border="0" cellspacing="0" cellpadding="0">
                          <tr valign="top"> 
                            <td height="100"> <%
sql="select top 5 A.ArticleID,A.ClassID,L.LayoutID,L.LayoutFileName,A.Title,A.Key,A.Author,A.CopyFrom,A.UpdateTime,A.Editor,A.TitleFontColor,A.TitleFontType,"
sql=sql & "A.Hits,A.OnTop,A.Hot,A.Elite,A.Passed,A.IncludePic,A.Stars,A.PaginationType,A.ReadLevel,A.ReadPoint,A.DefaultPicUrl from Article A"
sql=sql & " inner join Layout L on A.LayoutID=L.LayoutID where A.Deleted=False and A.Passed=True and A.ClassID in (" & arrClassID & ")  order by A.OnTop,A.ArticleID desc"
rsArticle.open sql,conn,1,1
if rsArticle.bof and  rsArticle.eof then
	response.write "<li>没有任何文章</li>"
else
	call ArticleContent(26,True,True,False,0,False,True)
end if
rsArticle.close
				%> </td>
                          </tr>
                          <tr> 
                            <td align="right"> <%response.write "<a href='" & rsRoot(3) & "?ClassID=" & rsRoot(0) & "'>更多&gt;&gt;&gt;</a>"%> </td>
                          </tr>
                        </table> </td>
                    </tr>
                  </table></td>
              </tr>
            </table>
            <table width="410" border="0" align="center" cellpadding="0" cellspacing="0">
              <tr> 
                <td  height="15" align="center" valign="top">
				  <table width="98%" border="0" align="center" cellpadding="0" cellspacing="0">
                    <tr> 
                      <td height="13" Class="tdbg_left2"></td>
                    </tr>
                  </table>
				</td>
			  </tr>
			</table>
		  </td>
        </tr> 
      <% 
			rsRoot.movenext 
		loop 
	end if 
	rsRoot.close 
	set rsRoot=nothing 
%> 
          </table>
    </td>
    <td width="5"></td>
    <td width="160" valign="top" class="tdbg_rightall">
      <table border="0" cellpadding="0" cellspacing="0" width="160">
        <tr class="title_right">
          <td align="center"><table width="100%" border="0" cellspacing="0" cellpadding="0"> 
              <tr>  
                <td align="center" class="title_righttxt"><strong>本 站 公 告</strong></td>                         
              </tr>                         
            </table></td>             
        </tr>             
        <tr class="tdbg_right">             
          <td align="center" height="80"><table width="100%" height="100%" border="0" cellpadding="0" cellspacing="5">                         
              <tr>                          
                <td valign="top">                         
<marquee id="scrollarea" direction="up" scrolldelay="10" scrollamount="1" width="150" height="80" onmouseover="this.stop();" onmouseout="this.start();">                          
            <% call ShowAnnounce(1,5) %>                         
                  </marquee></td>                         
              </tr>                         
            </table>
			</td>             
        </tr>             
        <tr class="title_right">             
          <td ><table width="100%" border="0" cellspacing="0" cellpadding="0">                         
              <tr>                          
                <td align="center" class="title_righttxt"><strong>文 章 搜 索</strong></td>                         
              </tr>                         
            </table></td>             
        </tr>             
        <tr class="tdbg_right">             
          <td height="55"><table width="100%" height="100%" border="0" cellpadding="0" cellspacing="5">                         
              <tr>                          
                <td valign="middle">                          
                  <div align="center">                         
                    <% call ShowSearchForm("Article_Search.asp",1) %>                         
                  </div></td>                         
              </tr>                         
            </table></td>             
        </tr>             
        <tr class="title_right" align="center">             
          <td><table width="100%" border="0" cellspacing="0" cellpadding="0">                         
              <tr>                          
                <td align="center" class="title_righttxt"><strong>最 新 热 门</strong></td>                         
              </tr>                         
            </table></td>             
        </tr>             
        <tr class="tdbg_right">             
          <td height="80"><table width="100%" height="100%" border="0" cellpadding="0" cellspacing="5">                         
              <tr>                          
                <td height="87" valign="top">                         
                  <% call ShowHot(5,14) %>                         
                </td>                         
              </tr>                         
            </table></td>             
        </tr>             
        <tr class="title_right">             
          <td align="center"><table width="100%" border="0" cellspacing="0" cellpadding="0">                         
              <tr>                          
                <td align="center" class="title_righttxt"><strong>最 新 推 荐</strong></td>                         
              </tr>                         
            </table></td>             
        </tr>             
        <tr class="tdbg_right">             
          <td height="80"><table width="100%" height="100%" border="0" cellpadding="0" cellspacing="5">                         
              <tr>                          
                <td height="87" valign="top">                         
                  <% call ShowElite(5,14) %>                         
                </td>                         
              </tr>                         
            </table></td>             
        </tr>             
        <tr class="title_right">             
          <td align="center"><table width="100%" border="0" cellspacing="0" cellpadding="0">                         
              <tr>                          
                <td align="center" class="title_righttxt"><strong>最 新 热 图</strong></td>                         
              </tr>                         
            </table></td>             
        </tr>             
        <tr class="tdbg_right">             
          <td valign="top" height="80"><% call ShowPicArticle(0,2,20,1,1,150,100,200,true,false) %></td>             
        </tr>             
		<tr class="title_right">             
          <td align="center"><table width="100%" border="0" cellspacing="0" cellpadding="0">                         
              <tr>                          
                <td align="center" class="title_righttxt"><strong>友 情 链 接</strong></td>                         
              </tr>                         
            </table></td>             
        </tr>            
        <tr class="tdbg_right">             
          <td valign="top" height="130"><table width="100%" height="100%" border="0" cellpadding="0" cellspacing="5">                         
              <tr>                          
                <td align="center" valign="top"><br>                          
                  <% call ShowFriendSite(1,10,1,1) %>                         
                  <br> <br>                          
                  <% call ShowFriendSite(2,10,1,3) %>                         
                  <br> <a href='FriendSiteReg.asp'> <br>                         
                  申请</a>&nbsp;&nbsp;<a href='FriendSite.asp'>更多&gt;&gt;&gt;</a>                          
                </td>                         
              </tr>                         
            </table></td>             
        </tr>            
      </table>             
                 
               
</table>             
<table width="760" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF" class="tdbg">
  <tr>                          
    <td  height="13" align="center" valign="top"><table width="755" border="0" align="center" cellpadding="0" cellspacing="0">                         
        <tr>                          
          <td height="13" Class="tdbg_left2"></td>                         
        </tr>                         
      </table></td>                         
  </tr>                         
</table>                         
<% call Bottom() %>                         
<% call PopAnnouceWindow(400,300) %>                         
</body>                         
</html>                         
<%                         
set rsArticle=nothing                         
set rsPic=nothing                         
call CloseConn()                         
%>