<!--#include file="Inc/syscode_article.asp"-->
<%
const ChannelID=2
const ShowRunTime="Yes"
MaxPerPage=10
SkinID=0
PageTitle="搜索结果"
strFileName="Article_Search.asp?Field=" & strField & "&Keyword=" & keyword & "&ClassID=" & ClassID
%>
<html>
<head>
<title><%=SiteName%>--文章搜索结果</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<!--#include file="inc/Skin_CSS.asp"-->
<% call MenuJS() %>
</head>
<body  <%=Body_Label%>  onmousemove='HideMenu()'  bgcolor="#FFFFFF"  style="BACKGROUND-COLOR: #ffffff" >
<table width="989"><tr><td>
<div align="center" ><table id="__01" width="989"  border="0" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
	<tr>
		<td colspan="2">
			<img   src="images/首页_slice2_03.jpg" width="989" height="140" alt=""></td>
	</tr>
	<tr>
       <td align="left" background="images/首页_slice_05.jpg" width="84%" height="25">&nbsp;&nbsp;<a href="index.asp">首&nbsp;&nbsp;&nbsp;&nbsp;页</a>&nbsp;&nbsp;<font color="#FFFFFF">‖</font>&nbsp;&nbsp;<a href="Article_Class2.asp?ClassID=2">资料中心</a>&nbsp;&nbsp;<font color="#FFFFFF">‖</font>&nbsp;&nbsp;<a href="Article_Class2.asp?ClassID=1">理论动态</a>&nbsp;&nbsp;<font color="#FFFFFF">‖</font>&nbsp;&nbsp;<a href="Article_Class2.asp?ClassID=3">时事新闻</a>&nbsp;&nbsp;<font color="#FFFFFF">‖</font>&nbsp;&nbsp;<a href="Article_Class2.asp?ClassID=58">学生作品</a>&nbsp;&nbsp;<font color="#FFFFFF">‖</font>&nbsp;&nbsp;<a href="userlist.asp">文&nbsp;&nbsp;&nbsp;&nbsp;集</a>&nbsp;&nbsp;<font color="#FFFFFF">‖</font>&nbsp;&nbsp;<a href="guestbook.asp">留&nbsp;&nbsp;&nbsp;&nbsp;言</a></td><td background="images/首页_slice_05.jpg"  align="left"><% call ShowSearchForm("Article_Search.asp",1) %></td>
		<td colspan="2"><img src="images/分隔符.gif" width="1" height="25" alt=""></td>
	</tr>
    </table><!--top--></div></td></tr>
<tr><!--the great talbe--><Td>


<table width="989" border="0" align="center" cellpadding="0" cellspacing="0" class="border2">
  <tr> 
    <td  valign="top"> <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="title_main">
        <tr> 
          <td width="40">&nbsp;</td>
          <td><table width="100%" border="0" cellpadding="0" cellspacing="0" >
              <tr> 
                <td class="title_maintxt"><%
		if keyword="" then
			response.write "所有文章"
		else
			select case strField
				case "Title"
					response.write "文章标题含有 <font color=red>"&keyword&"</font> 的文章"
				case "Content"
					response.write "文章内容含有 <font color=red>"&keyword&"</font> 的文章"
				case "Author"
					response.write "作者姓名含有 <font color=red>"&keyword&"</font> 的文章"
				case "Editor"
					response.write "编辑姓名含有 <font color=red>"&keyword&"</font> 的文章"
				case else
					response.write "文章标题含有 <font color=red>"&keyword&"</font> 的文章"
			end select
		end if
%>
                </td>
              </tr>
            </table></td>
        </tr>
      </table>
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td height="5"></td>
        </tr>
      </table>
      <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="border">
        <tr> 
          <td valign="top"> <table width="100%" height="100%" border="0" cellpadding="0" cellspacing="5">
              <tr> 
                <td valign="top">
                  <%call ShowSearchResult()%>
                </td>
              </tr>
            </table></td>
        </tr>
        <tr class="tdbg"> 
          <td> <table width="100%" border="0" cellpadding="0" cellspacing="0">
              <tr class="tdbg_leftall"> 
                <td> 
                  <%
		  if totalput>0 then
		  	call showpage(strFileName,totalput,MaxPerPage,false,true,"篇文章")
		  end if
		  %>
                </td>
              </tr>
            </table></td>
        </tr>
      </table>
      <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
        <tr> 
          <td  height="15" align="center" valign="top"> <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
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
      </table>
      
    </td>
  </tr>
</table>
</Td><!--the great talbe--></tr><tr><!--the great talbe--><td>
<% call Bottom_all() %></td></tr></table>
</body>
</html>
<% call CloseConn() %>