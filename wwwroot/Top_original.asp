<META http-equiv=Content-Type content="text/html; charset=gb2312">
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
                  <% call ShowSearchForm("Article_Search.asp",1) %>
                </td>
              </tr>
            </table></div>
	</td>
  </tr>
</table>
