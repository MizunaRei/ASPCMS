<!--#include file="Inc/syscode_article.asp"-->
<%
const ChannelID=2
Const ShowRunTime="Yes"
dim tLayout,tUser
PageTitle="正文"
strFileName="Article_Show.asp"
if ArticleId<=0 or ArticleID="" then
	FoundErr=true
	ErrMsg=ErrMsg & "<br><li>请指定文章ID</li>"
end if
if FoundErr=True then
	call WriteErrMsg()
else
%>
<html>
<head>
<title><%=ArticleTitle%></title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<!--#include file="inc/Skin_CSS.asp"-->
<% call MenuJS() %>
<script language="JavaScript" type="text/JavaScript">
//双击鼠标滚动屏幕的代码
var currentpos,timer;
function initialize()
{
timer=setInterval ("scrollwindow ()",30);
}
function sc()
{
clearInterval(timer);
}
function scrollwindow()
{
currentpos=document.body.scrollTop;
window.scroll(0,++currentpos);
if (currentpos !=document.body.scrollTop)
sc();
}
document.onmousedown=sc
document.ondblclick=initialize
</script>
<script language="JavaScript" type="text/JavaScript">
<!--
function MM_reloadPage(init) {  //reloads the window if Nav4 resized
  if (init==true) with (navigator) {if ((appName=="Netscape")&&(parseInt(appVersion)==4)) {
    document.MM_pgW=innerWidth; document.MM_pgH=innerHeight; onresize=MM_reloadPage; }}
  else if (innerWidth!=document.MM_pgW || innerHeight!=document.MM_pgH) location.reload();
}
MM_reloadPage(true);
//-->
</script></head>
<body <%=Body_Label%> bgcolor=#ffffff style="BACKGROUND-COLOR: #ffffff" onmousemove='HideMenu()' oncontextmenu="return false" ondragstart="return false" onselectstart ="return false" onselect="document.selection.empty()" oncopy="document.selection.empty()" onbeforecopy="return false" onmouseup="document.selection.empty()"> 
<table><tr><td><%	call Top_noIndex() %></td></tr>
<tr><td><!--<table width="989" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr> 
    <td width="220" height="8" background="images/to_bj.gif"></td>
  </tr>
</table>-->
<table width="989" border="0" cellspacing="0" cellpadding="0" align="center" bgcolor="F7F6ED">
  <tr> 
    <td height="55" valign="bottom"> 
      <table width="100%" border="0" cellspacing="0" cellpadding="0" align="center">
        <tr> 
          <td>&nbsp; 
            <div align="center"><font size="4"><strong><%=rs("Title")%></strong></font>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
              
              <%
			  response.write "文 / "
		set tUser=Conn_User.execute("select " & db_User_ID & " from " & db_User_Table & " where " & db_User_Name & "='" & rs("Editor") & "'")
		if tUser.bof and tUser.eof then
			response.write rs("Editor")
		else
			response.write "<strong>" & rs("Editor") & "</strong>"
		end if
		%></a> 
              
            </div>
          </td>
        </tr>
        <tr>
          <td>
          </td>
        </tr>
        <tr> 
          <td align="right"> 
            <%
if rs("OnTop")=true then
	response.Write("<font color=blue>固顶</font>&nbsp;")
else
	response.write("&nbsp;&nbsp;&nbsp;")
end if
if rs("Hits")>=HitsOfHot then
	response.write("<font color=red>热门</font>&nbsp;")
else
	response.write("&nbsp;&nbsp;&nbsp;")
end if
if rs("Elite")=true then
	response.write("<font color=green>本站推荐</font>")
else
	response.write("&nbsp;&nbsp;")
end if
response.write "&nbsp;&nbsp;<font color='#009900'>" & string(rs("Stars"),"★") & "</font>"
%>
          </td>
        </tr>
        <tr> 
          <td height="5"></td>
        </tr>
      </table>
    </td>
  </tr>
</table>
<table width="989" border="0" cellspacing="0" cellpadding="0" align="center" bgcolor="#F7F6ED">
  <tr> 
    <td colspan="3">
      <table width="100%" border="0" cellspacing="0" cellpadding="0"开獕牥也浡???	??????屐?屐???>
        <tr> 
          <td background="images/title_line.gif">&nbsp;</td>
          <td width="20">&nbsp;</td>
        </tr>
      </table>
    </td>
  </tr>
  <tr> 
    <td width="197" height="200" valign="top"> 
     <!--作者文集--> <table width="197" border="0" cellspacing="0" cellpadding="0" height="100%">
        <tr> 
          <td width="197" height="190" valign="top"> 
            <table width="197" border="0" cellspacing="0" cellpadding="0" align="center" height="100%">
  
              <tr> 
                <td align="center" valign="top"> 
                  <TABLE id=table5 height=100 cellSpacing=0 cellPadding=0 width=197 
            border=0>
                <TBODY>
                  
                  <TR>
                    <TD  ></td></tr>
                    <tr><td height=24 background="images/课程列表副本_06.jpg" width=197 border=0>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<strong>作者其他作品</strong></TD>
                  </TR>
                  <TR>
                    <TD align="center"  valign="top" ><% call ShowCorrelative(10,16) %>
                    </TD>
                  </TR>
                  <TR>
                    <TD align="right"  valign="bottom" ><% call ShowAuthorName() %>&nbsp;&nbsp;
                    </TD>
                  </TR>
                  <TR>
                    <TD height=24></TD>
                  </TR>
                </TBODY>
              </TABLE>
                  
                  
                  
                  
                </td>
              </tr>
            </table>
          </td>
        </tr>
      </table><!--结束作者文集-->
    </td>
    <td valign="top" width="3" background="images/07bj.gif"></td>
    <td valign="top">
      <table width="92%" border="0" cellspacing="0" cellpadding="0" align="center">
        
        <tr> 
          <td align="center" height="3"></td>
        </tr>
        <tr> 
          <td class="editorword"> 
            <%call ShowArticleContent()%>
          </td>
        </tr>
        
      </table>
    </td>
  </tr>
</table>
<table width="989" border="0" cellspacing="0" cellpadding="0" align="center"  >
  <!--<tr> 
    <td background="images/04bj.gif" height="15" colspan="2">&nbsp;</td>
  </tr>-->
  <tr bgColor="#FFFFFF" > 
    
    <td  width="100%" align="center" colspan="10"> 
      <div align="center"><table width="100%" border="0" cellspacing="0" cellpadding="0" align="center" >
        <tr align="center"> 
          <td align="center" height="20" colspan="10"><div align="center"><%
		  response.write "发表时间[" & FormatDateTime(rs("UpdateTime"),2) & "]"
		  %> | 
            <%
		set tUser=Conn_User.execute("select " & db_User_ID & " from " & db_User_Table & " where " & db_User_Name & "='" & rs("Editor") & "'")
		if tUser.bof and tUser.eof then
			response.write rs("Editor")
		else
			response.write "<strong>" & rs("Editor") & "</strong>"
		end if
		%>
             | <a href="javascript:window.close();">关闭窗口</a> 
          </div></td>
        </tr>
     <!--   <tr> 
          <td> 
            <table width="100%" border="0" cellspacing="0" cellpadding="0" align="center">
              <tr> 
                <td height="2" background="images/06bj.gif"></td>
              </tr>
            </table>
          </td>
        </tr>-->
        <tr align="center"> 
          <td height="18" valign="bottom" align="center" colspan="10"> 
            <%
		response.write " 此文已被阅读 " & rs("Hits") & ""
		%>
            次 | <a href="Article_Comment.asp?ArticleID=<%=rs("ArticleID")%>" target="_blank">发表评论</a> 
            | <a href="SendMail.asp?ArticleID=<%=rs("ArticleID")%>" target="_blank"><font color="#FF0000"></font></a> 
            <% call ShowComment(10) %>
            | <a href="SendMail.asp?ArticleID=<%=rs("ArticleID")%>" target="_blank"><font color="#FF0000">将此文推荐给好友或媒体</font></a> 
            | <a href="Article_Print.asp?ArticleID=<%=rs("ArticleID")%>">打印此文</a></td>
        </tr>
      </table></div>
    </td>
  </tr>
  <!--<tr> 
    <td bgcolor="F7F6ED" colspan="2" height="2"  background="images/05bj.gif"></td>
  </tr>
  
  <tr> 
    <td colspan="2" height="1"  background="images/05bj.gif"></td>
  </tr>-->
  
</table></td><!--the great table--></tr><tr><!--the great table--><td><% call bottom_all() %></td><!--the great table--></tr><!--the great table--></table>

</body>
</html>
<%
end if
if not (ArticleId<=0 or ArticleID="") then
	 rs.close
	set rs=nothing
	end if 
call CloseConn()
%>