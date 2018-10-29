<!--#include file="Inc/syscode_guest.asp"-->
<%
'请勿改动下面这三行代码
const ChannelID=5
Const ShowRunTime="Yes"
SkinID=0
%>
<html>
<head>
<title><%=strPageTitle & " >> " & PageTitle%></title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<!--#include file="inc/Skin_CSS.asp"-->
<%call MenuJS()%>
</head>
<body <%=Body_Label%> onmousemove='HideMenu()'>
<div id=menuDiv style='Z-INDEX: 1000; VISIBILITY: hidden; WIDTH: 1px; POSITION: absolute; HEIGHT: 1px; BACKGROUND-COLOR: #9cc5f8'></div>
<table width="770" border="0" align="center" bgcolor="#FFFFFF">
  <tr >
    <td colspan="3" height="3" bgColor=#e7ddd1></td>
  </tr>
</table>
<TABLE cellSpacing=0 cellPadding=0 width=770 align=center border=0 class="txt_css">
  <TBODY>
  <TR>
    <TD vAlign=bottom background=images/line-01.gif 
    height=26><%
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
    <td ><div align="center"><IMG src="images/arr.gif" ></div></td>
    <td >
<%call ShowPath()%> </td>
  </tr>
</table>
<table width="770" border="0" align="center" cellpadding="0" cellspacing="0" class="border2">
  <tr> 
    <td width="180" align="left" valign="top" class="tdbg_leftall"> <TABLE cellSpacing=0 cellPadding=0 width="100%" border="0" style="word-break:break-all">
        <TR> 
          <TD align="center" background="Images/left01.gif"><table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr class="title_left"> 
                <td class="title_lefttxt"><div align="center"><strong>用 户 登 录</strong></div></td>
              </tr>
            </table></TD>
        </TR>
        <TR> 
          <TD height="40" valign="top" class="tdbg_left"> <table width="100%" height="100%" border="0" cellpadding="0" cellspacing="5">
              <tr> 
                <td valign="top"> <% call ShowUserLogin() %> </td>
              </tr>
            </table></TD>
        </TR>
        <TR>
          <td class="title_left2"></td>
        </TR>
      </table>
      <TABLE cellSpacing=0 cellPadding=0 width="100%" border="0" style="word-break:break-all">
        <TR> 
          <TD align="center" background="Images/left10.gif"><table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr class="title_left"> 
                <td class="title_lefttxt"> <div align="center"><b>留 言 功 能</b></div></td>
              </tr>
            </table></TD>
        </TR>
        <TR> 
          <TD height="40" valign="top" class="tdbg_left"> <table width="100%" height="100%" border="0" cellpadding="0" cellspacing="5">
              <tr> 
                <td valign="top"><% call GuestBook_Left() %></td>
              </tr>
            </table></TD>
        </TR>
        <TR>
          <td class="title_left2"></td>
        </TR>
      </table>
      <TABLE cellSpacing=0 cellPadding=0 width="100%" border="0" style="word-break:break-all">
        <TR> 
          <TD align="center" background="Images/left11.gif"><table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr class="title_left"> 
                <td class="title_lefttxt"> <div align="center"><b>留 言 搜 索</b></div></td>
              </tr>
            </table></TD>
        </TR>
        <TR> 
          <TD height="40" valign="top" class="tdbg_left"> <table width="100%" height="100%" border="0" cellpadding="0" cellspacing="5">
              <tr> 
                <td valign="top"><% call GuestBook_Search() %></td>
              </tr>
            </table></TD>
        </TR>
        <TR>
          <td class="title_left2"></td>
        </TR>
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
      </table>
    </td>
    <td width="5" align="left" valign="top">&nbsp;</td>
    <td valign="top"> 
      <table width="100%" border="0" cellpadding="0" cellspacing="0"  background="images/fcbg1_1.gif">
        <tr> 
          <td width="92"><strong><img src="Images/announce.gif" width="20" height="16" align="absmiddle">&nbsp;最新公告</strong></td>
          <td width="483"><div align="right"> 
        <MARQUEE scrollAmount=1 scrollDelay=4 width=480
            align="left" onmouseover="this.stop()" onmouseout="this.start()">
        <% call ShowAnnounce(2,5) %>
        </MARQUEE>
      </div></td>
        </tr>
      </table>
	  <%
		call showtip()
		call Guestbook()
	  %>
      <table width="99%" border="0" align="center" cellpadding="0" cellspacing="0" class="tdbg_rightall">
        <tr background="images/fcbg2.gif"> 
          <td> 
			<%
				call ShowGuestPage()
			%>
          </td>
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
<% call PopAnnouceWindow(400,300) %>
</body>
</html>
<%
call CloseConn()
%>