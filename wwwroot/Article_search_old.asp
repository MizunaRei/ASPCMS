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
<body <%=Body_Label%> onmousemove='HideMenu()'>
 <table width="770" border="0" align="center" bgcolor="#FFFFFF">
    <tr > 
    <td colspan="3" height="3" bgColor=#e7ddd1></td>
  </tr>
</table>
<TABLE cellSpacing=0 cellPadding=0 width=770 align=center border=0 class="txt_css">
  <TBODY>
  <TR>
    <TD vAlign=bottom background=images/line-01.gif 
    height=26 ><%
	if ShowSiteChannel="Yes" then
		response.write strChannel
	else
		response.write "&nbsp;"
	end if
    	if ShowMyStyle="Yes" then
		response.write "<a href='#' onMouseOver='ShowMenu(menu_skin,100)'>自选风格&nbsp;</a>|"
	end if
	%></TD></TR>
  <TR>
    <TD bgColor=#000000 height=3></TD></TR></TBODY></TABLE>
<TABLE height=10 cellSpacing=0 width=770 align=center border=0 id="table1">
  <TBODY>
  <TR>
      <TD><IMG height=125 src="Images/wenxie-04.gif" 
width=772></TD>
    </TR></TBODY></TABLE>
<table width="770" height="30" border="0" align="center" cellpadding="0" cellspacing="0" background="images/titlebg2.jpg" >
  <tr valign="middle"> 
    <td ><IMG src="images/arr.gif" ></td>
    <td >
<%call ShowPath()%> </td>
<td>
<div align="center">
	<table width='100%' border='0' cellpadding='0' cellspacing='5'>
        <tr> 
                <td height="18"> 
                  <div align="center">站内搜索</div></td>
                <td width="81%"> 
                  <% call ShowSearchForm("Article_Search.asp",2) %>
                </td>
              </tr>
            </table></div>
	</td>
  </tr>
</table>



<table width="770" border="0" align="center" cellpadding="0" cellspacing="0" class="border2">
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
</body>
</html>
<% call CloseConn() %>