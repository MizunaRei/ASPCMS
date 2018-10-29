<DIV align=center>
<TABLE id=table1 height=100 cellSpacing=0 cellPadding=0 width=797 
 background=SkinIndex/new_wow_43.jpg border=0>
  <TBODY>
  <TR>
    <TD width=8 rowSpan=3></TD>
    <!--<TD vAlign=bottom align=right width=781 background=SkinIndex/bg001.jpg 
    height=2></TD>-->
    <TD width=8 rowSpan=3></TD></TR>
  <TR>
    <TD width=781 bgColor=#ffffff>
      <TABLE id=table10 height=100 cellSpacing=0 cellPadding=0 width="100%" 
      border=0>
        <TBODY>
        <TR>
          <!--<TD bgColor=#e1dcb9 height=2 background=SkinIndex/bg001.jpg></TD>--></TR>
        <TR>
          <TD align=middle width="100%" bgColor=#e1dcb9 background=SkinIndex/bg001.jpg>
            <TABLE height=80 cellSpacing=0 cellPadding=0 width=760 border=0>
              <TBODY>
              <TR>
                <META http-equiv=Content-Type content="text/html; charset=gb2312">
 <!--<table width="770" border="0" align="center" bgcolor="#FFFFFF">
    <tr > 
    <td colspan="3" height="3" bgColor=#e7ddd1></td>
  </tr>
</table>-->
<TABLE cellSpacing=0 cellPadding=0 width=780 align=center border=0 class="txt_css">
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
      <TD><IMG height=143 src="NewImages/Banner.jpg" 
width=782></TD>
    </TR></TBODY></TABLE>
<table width=780 height="30" border="0" align="center" cellpadding="0" cellspacing="0" background=SkinIndex/njyyjy_10.gif >
  <tr valign="middle" align="left"> 
    <td  valign="middle">&nbsp;<IMG src="images/arr.gif" ></td>
    <td align="left"><%call ShowPath()%> </td>
<td align="right">
<div align="center">
	<table width='100%' border='0' cellpadding='0' cellspacing='5'>
        <tr> 
                <td height="18"> 
                  <div align="center">搜索</div></td>
                <td width="81%"> 
                  <% call ShowSearchForm("Article_Search.asp",1) %>
                </td>
              </tr>
            </table></div>
	</td>
  </tr>
</table>
</TR></TBODY></TABLE></TD></TR>
        <!--<TR>
          <TD bgColor=#e1dcb9 height=17 background=SkinIndex/bg001.jpg></TD></TR>--></TBODY></TABLE></TD></TR>
  <!--<TR>
    <TD align=middle width=781 background=SkinIndex/njyyjy_10.gif 
    height=40></TD></TR>--></TBODY></TABLE>
</DIV>