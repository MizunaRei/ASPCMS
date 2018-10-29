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
<body <%=Body_Label%> onmousemove='HideMenu()'>
<!--#include file="Top.asp"-->
<%
dim sqlRoot,rsRoot,trs,arrClassID,TitleStr
sqlRoot="select C.ClassID,C.ClassName,C.RootID,L.LayoutFileName,L.LayoutID,C.Child,C.ParentPath From ArticleClass C inner join Layout L on C.LayoutID=L.LayoutID where C.ParentID=" & ClassID & " and C.IsElite=True and C.LinkUrl='' and C.BrowsePurview>=" & UserLevel & " order by C.OrderID"
Set rsRoot= Server.CreateObject("ADODB.Recordset")
rsRoot.open sqlRoot,conn,1,1
%>
<table width="770" border="0" align="center" cellpadding="0" cellspacing="0" class="border2">
  <tr> 
    <td width="180" align="left" valign="top" class="tdbg_leftall">
	 <TABLE cellSpacing=0 cellPadding=0 width="100%" border="0">
        <TR> 
          <TD align="center" background="Images/left01.gif"><table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr class="title_left"> 
                <td class="title_lefttxt"><div align="center"><strong>用 户 登 录</strong></div></td>
              </tr>
            </table></TD>
        </TR>
        <TR> 
          <TD height="40" valign="top" class="tdbg_left"> <table width="100%" height="100%" border="0" cellpadding="3">
              <tr> 
                <td valign="top"> <% call ShowUserLogin() %> </td>
              </tr>
            </table></TD>
        </TR>
        <TR> 
          <td class="title_left2"></td>
        </TR>
        <%if Child>0 then%>
        <tr> 
          <td background="Images/left12.gif"><table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr class="title_left"> 
                <td class="title_lefttxt"><div align="center"><strong>子栏目导航</strong></div></td>
              </tr>
            </table></td>
        </tr>
        <tr class='tdbg_left'> 
          <td height="50" valign="top"> <table width="100%" height="100%" border="0" cellpadding="8">
              <tr> 
                <td valign="top"> <% call ShowChildClass(1) %> </td>
              </tr>
            </table></td>
        </tr>
        <tr class='tdbg_left'> 
          <td class="title_left2"></td>
        </tr>
        <%end if%>
        <tr align='center'> 
          <td background="Images/left03.gif"><table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr class="title_left"> 
                <td class="title_lefttxt"><div align="center"><strong>专 题 栏 目</strong></div></td>
              </tr>
            </table></td>
        </tr>
        <tr class='tdbg_left'> 
          <td height="50" valign="top"> <table width="100%" height="100%" border="0" cellpadding="8">
              <tr> 
                <td valign="top"> <% call ShowSpecial(10) %> </td>
              </tr>
            </table></td>
        </tr>
        <tr class='tdbg_left'> 
          <td class="title_left2"></td>
        </tr>
        <tr align='center'> 
          <td background="Images/left02.gif"><table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr class="title_left"> 
                <td class="title_lefttxt"><div align="center"><b>最 新 热 门</b></div></td>
              </tr>
            </table></td>
        </tr>
        <tr class='tdbg_left'> 
          <td height="60" valign="top"> <table width="99%" height="100%" border="0" cellpadding="8">
              <tr> 
                <td valign="top"> <% call ShowHot(10,16) %> </td>
              </tr>
            </table></td>
        </tr>
        <tr class='tdbg_left'> 
          <td class="title_left2"></td>
        </tr>
        <tr align='center'> 
          <td background="Images/left08.gif"><table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr class="title_left"> 
                <td class="title_lefttxt"><div align="center"><b>最 新 推 荐</b></div></td>
              </tr>
            </table></td>
        </tr>
        <tr class='tdbg_left'> 
          <td height="67" valign="top"> <table width="100%" height="100%" border="0" cellpadding="8">
              <tr> 
                <td valign="top"> <% call ShowElite(10,16) %> </td>
              </tr>
            </table></td>
        </tr>
        <tr class='tdbg_left'> 
          <td class="title_left2"></td>
        </tr>
      </table>
    </td>
    <td width="5"></td>
    <td width="575" valign="top"><table width="100%" border="0" cellpadding="0" cellspacing="0"><tr>
          <td> 
            <table width="99%" border="0" align="center" cellpadding="0" cellspacing="0" class="tdbg_right2">
              <tr> 
                <td width="40">　</td>
                <td width="479" class="title_maintxt"><%=ClassName%> 文章列表</td>
                <td width="56" class="title_maintxt"><font color="#666666"><a href="Article_Class.asp?ClassID=1">more...</a></font></td>
              </tr>
            </table>
            <table width="99%" border="0" align="center" cellpadding="0" cellspacing="5" class="border">
              <tr> 
          <td height="200" valign="top"> 
            <script language='JavaScript' type='text/JavaScript' src='article_js.asp?ClassID=1&IncludeChild=true&SpecialID=&ArticleNum=20&ShowType=1&ShowCols=1&ShowProperty=true&ShowClassName=true&ShowIncludePic=true&ShowTitle=true&ShowUpdateTime=true&ShowHits=true&ShowAuthor=true&ShowHot=true&ShowMore=false&TitleMaxLen=40&ContentMaxLen=200&Hot=false&Elite=false&DateNum=&OrderField=UpdateTime&OrderType=desc'></script>
          </td>
        </tr>
      </table>
      <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
        <tr> 
          <td  height="15" align="center" valign="top"> <table width="98%" border="0" align="center" cellpadding="0" cellspacing="0">
              <tr> 
                <td height="13" Class="tdbg_left2"></td>
              </tr>
            </table></td>
        </tr>
      </table>
      <table width='100%' border='0'cellpadding='0' cellspacing='0'  >
        <tr align='center'> 
          <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr> 
                <td><table width="99%" border="0" align="center" cellpadding="0" cellspacing="0" class="title_right2">
                          <tr> 
                      <td width="40">　</td>
                            <td class="title_maintxt"><b>本栏最新图文</b></td>
                    </tr>
                  </table>
                  
                </td>
              </tr>
            </table></td>
        <tr> 
          <td height="80" valign="top"> <table width="99%" height="100%" border="0" align="center" cellpadding="0" cellspacing="5" class="border">
                    <tr> 
                      <td valign="top"> 
                        <% call ShowPicArticle(ClassID,4,14,2,2,110,95,78,false,false) %>
                      </td>
              </tr>
            </table></td>
        </tr>
      </table>
      
      <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
        <tr> 
          <td  height="15" align="center" valign="top"> <table width="98%" border="0" align="center" cellpadding="0" cellspacing="0">
              <tr> 
                <td height="13" Class="tdbg_left2"></td>
              </tr>
            </table></td>
        </tr>
      </table>
            <table width='99%' border='0' align="center"cellpadding='2' cellspacing='0' class="tdbg_rightall">
              <tr class='tdbg_leftall'> 
                <td width="22%"> <div align="center"><img src="Images/checkarticle.gif" width="15" height="15" align="absmiddle">&nbsp;&nbsp;站内文章搜索：</div></td>
                <td width="78%"> <div align="center"> 
                    <% call ShowSearchForm("Article_Search.asp",2) %>
                  </div></td>
              </tr>
            </table></td></tr></table>
	  </td>
  </tr>
</table>
<table width="770" border="0" align="center" cellpadding="0" cellspacing="0" class="tdbg">
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