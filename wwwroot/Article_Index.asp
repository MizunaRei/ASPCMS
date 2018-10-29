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
</head>
<body <%=Body_Label%> onmousemove='HideMenu()'>
<!--#include file="top.asp"-->
<table width="760" border="0" align="center" cellpadding="0" cellspacing="0" class="border2">
  <tr> 
    <td width="180" rowspan="2" align="left" valign="top" class="tdbg_leftall"> 
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td height="5"></td>
        </tr>
      </table>
      <table width="180" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td background="Images/left01.gif"> 
            <table width="100%" border="0" cellpadding="0" cellspacing="0">
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
          <td background="Images/left02.gif"><table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr class="title_left"> 
                <td class="title_lefttxt"><div align="center"><strong>最 新 热 门</strong></div></td>
              </tr>
            </table></td>
        </tr>
        <tr>
          <td class="tdbg_left"><table width="100%" border="0" cellpadding="8">
              <tr> 
                <td>
                  <% call ShowHot(10,14) %>
                </td>
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
                <td> <% call ShowTopUser(5) %> </td>
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
        <tr> 
          <td background="Images/left07.gif"><table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr class="title_left"> 
                <td class="title_lefttxt"><div align="center"><strong>友 情 链 接</strong></div></td>
              </tr>
            </table></td>
        </tr>
        <tr> 
          <td align="center" valign="top" class="tdbg_left"><table width="100%" border="0" cellpadding="3">
              <tr> 
                <td> <div align="center"> 
                    <% call ShowFriendSite(1,10,1,1) %>
                    <br>
                    <% call ShowFriendSite(2,10,1,3) %>
                    <br>
                    <a href='FriendSiteReg.asp'> 申请</a>&nbsp;&nbsp;<a href='FriendSite.asp'>更多&gt;&gt;&gt;</a> 
                  </div></td>
              </tr>
            </table></td>
        </tr>
        <tr> 
          <td class="title_left2"></td>
        </tr>
      </table>
    </td>
    <td width="5"></td>
    <td width="575" valign="top"><table width="575" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td valign="top"><table width=575 height=15 border=0 align="center" cellPadding=0 cellSpacing=0>
              <tr> 
                <td width="20"><img src="Images/announce.gif" width="20" height="16"></td>
                <td width="64"><div align="center"><font color="#CC0000">本站公告：</font></div></td>
                <td width="491" height=15 align=center valign=middle> <div align="right"> 
                    <MARQUEE scrollAmount=1 scrollDelay=4 width=480
            align="left" onmouseover="this.stop()" onmouseout="this.start()">
                    <% call ShowAnnounce(2,5) %>
                    </MARQUEE>
                  </div></td>
              </tr>
            </table></td>
        </tr>
        <tr> 
          <td valign="top"><table width="100%" border="0" cellpadding="0" cellspacing="0">
              <tr> 
                <td Class="title_main2"><table border="0" cellpadding="0" cellspacing="0">
                    <tr> 
                      <td width="27" height="39">&nbsp;</td>
                      <td width="307" valign="bottom"> <table width="100%" border="0" cellspacing="5" cellpadding="0">
                          <tr> 
                            <td class="title_maintxt"><img src="Images/Star.gif" width="10" height="11" align="absmiddle"> 
                              <strong>最新推荐</strong></td>
                          </tr>
                        </table></td>
                      <td width="54">&nbsp;</td>
                      <td width="158" valign="bottom"> <table width="100%" border="0" cellspacing="5" cellpadding="0">
                          <tr> 
                            <td class="title_maintxt"><img src="Images/D_1.gif" width="13" height="13" align="absmiddle"> 
                              <strong><strong>最新热门</strong></strong></td>
                          </tr>
                        </table></td>
                      <td width="29">&nbsp;</td>
                    </tr>
                    <tr> 
                      <td height="139">&nbsp;</td>
                      <td valign="top"><table width="100%" border="0" cellpadding="0" cellspacing="5">
                          <tr> 
                            <td width="33%" align="center" valign="middle"> 
                              <% call ShowPicArticle(0,1,10,1,1,80,90,200,false,true) %>
                            </td>
                            <td width="67%" valign="top"> 
                              <% call ShowElite(7,20) %>
                            </td>
                          </tr>
                        </table></td>
                      <td>&nbsp;</td>
                      <td valign="top"><table width="100%" height="100%" border="0" cellpadding="0" cellspacing="5">
                          <tr> 
                            <td valign="top"> <marquee id="scrollarea" direction="up" scrolldelay="200" scrollamount="2" width="150" height="130" onmouseover="this.stop();" onmouseout="this.start();">
                              <% call ShowHot(10,14) %>
                              </marquee></td>
                          </tr>
                        </table></td>
                      <td>&nbsp;</td>
                    </tr>
                    <tr> 
                      <td colspan="5">&nbsp;</td>
                    </tr>
                  </table></td>
              </tr>
            </table></td>
        </tr>
        <tr> 
          <td height="8"></td>
        </tr>
        <tr>
          <td align="center" valign="middle"> 
            <% call Showad(1) %>
          </td>
        </tr>
        <tr> 
          <td colspan="10"><div align="center"></div></td>
        </tr>
        <tr> 
          <td valign="top"> <table width="99%" border="0" align="center" cellpadding="0" cellspacing="0">
              <tr> 
                <%
	dim sqlRoot,rsRoot,trs,arrClassID,TitleStr,ClassCount,iClassID
	sqlRoot="select C.ClassID,C.ClassName,C.RootID,L.LayoutFileName,L.LayoutID,C.Child,C.Readme From ArticleClass C inner join Layout L on C.LayoutID=L.LayoutID where C.ParentID=0 and IsElite=True and LinkUrl='' order by C.RootID"
	Set rsRoot= Server.CreateObject("ADODB.Recordset")
	rsRoot.open sqlRoot,conn,1,1

	dim sqlClassAD,rsClassAD,ClassAD
	sqlClassAD="select * from Advertisement where IsSelected=True"
	sqlClassAD=sqlClassAD & " and (ChannelID=0 or ChannelID=" & ChannelID & ")"
	sqlClassAD=sqlClassAD & " and ADType=2 order by ID Desc"
	set rsClassAD=server.createobject("adodb.recordset")
	rsClassAD.open sqlClassAD,conn,1,1
	
	if rsRoot.bof and rsRoot.eof then 
		response.Write("还没有任何栏目，请首先添加栏目。")
	else
		ClassCount=rsRoot.recordcount
		iClassID=0
		do while not rsRoot.eof
%>
                <td valign="top" width="282"> <table width="100%" border="0" cellspacing="0" cellpadding="0">
                    <tr> 
                      <td><table width="100%" border="0" cellpadding="0" cellspacing="0" class="title_main">
                          <tr> 
                            <td width="68">&nbsp;</td>
                            <td width="468" class="title_maintxt"><%
				arrClassID=rsRoot(0)
				response.write "<a href='" & rsRoot(3) & "?ClassID=" & rsRoot(0) & "' title='" & rsRoot(6) & "'>" & rsRoot(1) & "</a>"
				if rsRoot(5)>0 then
					set trs=conn.execute("select ClassID from ArticleClass where RootID=" & rsRoot(2) & " and Child=0 and LinkUrl=''")
					do while not trs.eof
						arrClassID=arrClassID & "," & trs(0)
						trs.movenext
					loop
				end if
				%></td>
                            <td width="39" class="title_maintxt"> <%response.write "<a href='" & rsRoot(3) & "?ClassID=" & rsRoot(0) & "'><font color='#666666'>more...</font></a>&nbsp;"%> </td>
                          </tr>
                        </table></td>
                    </tr>
                    <tr> 
                      <td><table width="100%" border="0" cellpadding="0" cellspacing="0" class="border">
                          <tr> 
                            <td height="100" valign="top"> <table width="100%" height="100%" border="0" cellpadding="0" cellspacing="5">
                                <tr> 
                                  <td height="152" valign="top"> 
                                    <%
sql="select top 8 A.ArticleID,A.ClassID,L.LayoutID,L.LayoutFileName,A.Title,A.Key,A.Author,A.CopyFrom,A.UpdateTime,A.Editor,A.TitleFontColor,A.TitleFontType,"
sql=sql & "A.Hits,A.OnTop,A.Hot,A.Elite,A.Passed,A.IncludePic,A.Stars,A.PaginationType,A.ReadLevel,A.ReadPoint,A.DefaultPicUrl from Article A"
sql=sql & " inner join Layout L on A.LayoutID=L.LayoutID where A.Deleted=False and A.Passed=True and A.ClassID in (" & arrClassID & ")  order by A.OnTop,A.ArticleID desc"
rsArticle.open sql,conn,1,1
if rsArticle.bof and  rsArticle.eof then
	response.write "<li>没有任何文章</li>"
else
	call ArticleContent(20,True,True,False,1,False,True)
end if
rsArticle.close
				%>
                                  </td>
                                </tr>
                              </table></td>
                          </tr>
                        </table></td>
                    </tr>
                    <tr> 
                      <td  height="15" align="center" valign="top"> <table width="98%" border="0" align="center" cellpadding="0" cellspacing="0">
                          <tr> 
                            <td height="13" Class="tdbg_left2"></td>
                          </tr>
                        </table></td>
                    </tr>
                  </table></td>
                <%
			iClassID=iClassID+1
			if iClassID mod 2=0 then
				response.write "</tr><tr><td colspan='3' align='center'>"
				if not rsClassAD.bof and not rsClassAD.eof then
					if rsClassAD("isflash")=true then
						ClassAD= "<object classid='clsid:D27CDB6E-AE6D-11cf-96B8-444553540000' codebase='http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=5,0,0,0'"
						if rsClassAD("ImgWidth")>0 then ClassAD = ClassAD & " width='" & rsClassAD("ImgWidth") & "'"
						if rsClassAD("ImgHeight")>0 then ClassAD = ClassAD & " height='" & rsClassAD("ImgHeight") & "'"
						ClassAD = ClassAD & "><param name='movie' value='" & rsClassAD("ImgUrl") & "'><param name='quality' value='high'><embed src='" & rsClassAD("ImgUrl") & "' pluginspage='http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash' type='application/x-shockwave-flash'"
						if rsClassAD("ImgWidth")>0 then ClassAD = ClassAD & " width='" & rsClassAD("ImgWidth") & "'"
						if rsClassAD("ImgHeight")>0 then ClassAD = ClassAD & " height='" & rsClassAD("ImgHeight") & "'"
						ClassAD = ClassAD & "></embed></object>"
					else
						ClassAD ="<a href='" & rsClassAD("SiteUrl") & "' target='_blank' title='" & rsClassAD("SiteName") & "：" & rsClassAD("SiteUrl") & vbcrlf & rsClassAD("SiteIntro") & "'><img src='" & rsClassAD("ImgUrl") & "'"
						if rsClassAD("ImgWidth")>0 then ClassAD = ClassAD & " width='" & rsClassAD("ImgWidth") & "'"
						if rsClassAD("ImgHeight")>0 then ClassAD = ClassAD & " height='" & rsClassAD("ImgHeight") & "'"
						ClassAD = ClassAD & " border='0'></a>"
					end if
					response.write ClassAD
					rsClassAD.movenext
				end if
				response.write "</td><tr><td height='5'></td></tr><tr>"
			else
				response.write "<td width='5'></td>"
			end if
			rsRoot.movenext
		loop
	end if
	rsRoot.close
	set rsRoot=nothing
%>
              </tr>
            </table></td>
        </tr>
        <tr> 
          <td><table width='100%' border='0'cellpadding='0' cellspacing='0'>
              <tr> 
                <td><table width="99%" border="0" align="center" cellpadding="0" cellspacing="0" class="title_right2">
                    <tr> 
                      <td width="36">&nbsp;</td>
                      <td width="418" class="title_maintxt">最新图片文章</td>
                      <td width="121" class="title_maintxt"><font color="#666666">www.asp163.net</font></td>
                    </tr>
                  </table></td>
              <tr> 
                <td height="80" valign="top"> <table width="99%" height="100%" border="0" align="center" cellpadding="0" cellspacing="5" class="border">
                    <tr> 
                      <td valign="top"> <% call ShowPicArticle(0,4,20,1,4,130,90,200,false,false) %> </td>
                    </tr>
                  </table></td>
              </tr>
              <tr> 
                <td  height="15" align="center" valign="top"> <table width="98%" border="0" align="center" cellpadding="0" cellspacing="0">
                    <tr> 
                      <td height="13" Class="tdbg_left2"></td>
                    </tr>
                  </table></td>
              </tr>
            </table></td>
        </tr>
        <tr> 
          <td><table width='100%' border='0'cellpadding='0' cellspacing='0'>
              <tr> 
                <td><table width="99%" border="0" align="center" cellpadding="0" cellspacing="0" class="tdbg_right2">
                    <tr> 
                      <td width="38">&nbsp;</td>
                      <td width="416" class="title_maintxt">栏目导航</td>
                      <td width="121" class="title_maintxt"><font color="#666666">www.asp163.net</font></td>
                    </tr>
                  </table></td>
              <tr> 
                <td height="80" valign="top"> <table width="99%" height="100%" border="0" align="center" cellpadding="0" cellspacing="5" class="border">
                    <tr> 
                      <td valign="top"> <% call ShowClassNavigation() %> </td>
                    </tr>
                  </table></td>
              </tr>
              <tr> 
                <td  height="15" align="center" valign="top"> <table width="98%" border="0" align="center" cellpadding="0" cellspacing="0">
                    <tr> 
                      <td height="13" Class="tdbg_left2"></td>
                    </tr>
                  </table></td>
              </tr>
            </table></td>
        </tr>
        <tr> 
          <td> <table width='99%' border='0' align="center"cellpadding='2' cellspacing='0' class="tdbg_rightall">
              <tr class='tdbg_leftall'> 
                <td width="22%"> <div align="center"><img src="Images/checkarticle.gif" width="15" height="15" align="absmiddle">&nbsp;&nbsp;站内文章搜索：</div></td>
                <td width="78%"> <div align="center"> 
                    <% call ShowSearchForm("Article_Search.asp",2) %>
                  </div></td>
              </tr>
            </table></td>
        </tr>
      </table>
	</td>
  </tr>
</table>
<table width="760" border="0" align="center" cellpadding="0" cellspacing="0" class="tdbg">
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
<% call ShowAD(0) %>                         
<% call ShowAD(4) %>                         
<% call ShowAD(5) %>                         
</body>
</html>
<%
set rsArticle=nothing
set rsPic=nothing
call CloseConn()
%>