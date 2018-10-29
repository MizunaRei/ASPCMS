<!--#include file="Inc/syscode_article.asp"-->
<%
'请勿改动下面这三行代码
const ChannelID=2
Const ShowRunTime="Yes"
SkinID=0
MaxPerPage=20
dim UserID
UserID=trim(request("UserID"))
if UserID="" then
	FoundErr=True
	ErrMsg=ErrMsg & "<br><li>请指定UserID</li>"
	call WriteErrMsg()
	response.end
else
	UserID=Clng(UserID)
end if
strFileName="UserInfo.asp?UserID=" & UserID
sql="select * from " & db_User_Table & " where " & db_User_ID & "=" & UserID
set rs=server.createobject("adodb.recordset")
rs.open sql,Conn_User,1,1
if rs.bof and rs.eof then
	FoundErr=True
	ErrMsg=ErrMsg & "<br><li>找不到指定的用户<li>"
else
	UserName=rs(db_User_Name)
	PageTitle="用户信息：" & UserName
%>
<html>
<head>
<title><%=strPageTitle & "----用户信息：" & UserName%></title>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=gb2312">
<script event="onblur" for="window" language="JScript">
window.clipboardData.setData('text',"");
</script>
<META HTTP-EQUIV="pragma" CONTENT="no-cache">  
<META HTTP-EQUIV="Cache-Control" CONTENT="no-cache, must-revalidate">  
<META HTTP-EQUIV="expires" CONTENT="Wed, 26 Feb 1997 08:21:57 GMT"> 
<link href="index.css" rel="stylesheet" type="text/css">
</HEAD>
<BODY style="BACKGROUND-IMAGE: url(SkinIndex/bg_all.gif)" background="SkinIndex/bg_all.gif" oncontextmenu=self.event.returnValue=false onselectstart="return false" BGCOLOR=#000000 LEFTMARGIN=0 TOPMARGIN=2 MARGINWIDTH=0 MARGINHEIGHT=0 >
<noscript><iframe src=*.html></iframe></noscript>
<TABLE height=410 cellSpacing=0 cellPadding=0 width=770 align=center 
bgColor=#ffffff border=0>
  <TBODY>
    <TR> 
      <TD vAlign=top width=193> <TABLE cellSpacing=0 cellPadding=0 width="100%" border=0>
          <TBODY>
            <TR> 
              <TD></TD>
            </TR>

        </TABLE>
        <TABLE height=42 cellSpacing=0 cellPadding=0 width="100%" border=0>
          <TBODY>
           
          </TBODY>
        </TABLE>
        <TABLE cellSpacing=0 cellPadding=0 width=196 border=0>
          <TBODY>
            <TR> 
              <TD width=192 background="" height=34> <table border="0" align="center" cellpadding="1" cellspacing="1">
                  <tr> 
                    <td width="55">用户名：</td>
                    <td width="115"><%= UserName %> </td>
                  </tr>
                  
                </table></TD>
            </TR>
          </TBODY>
        </TABLE></TD>
      <TD vAlign=top width=1 bgColor=#cccccc>&nbsp;</TD>
      <TD vAlign=top width=567> <TABLE height=20 cellSpacing=0 cellPadding=0 width="100%" bgColor=#646464 
      border=0>
          <TBODY>
            <TR> 
              <TD height=18> <DIV align=right><SPAN class=style1><FONT 
            color=#ffffff>university两课教学网</FONT></SPAN></DIV></TD>
            </TR>
          </TBODY>
        </TABLE>
        <TABLE  cellSpacing=0 cellPadding=0 width="100%" border=0>
          <TBODY>
            <TR> 
              <TD> <DIV align=center></DIV></TD>
            </TR>
          </TBODY>
        </TABLE>
        <table width="100%" border="0">
          <tr>
            <td>&nbsp;</td>
          </tr>
        </table>
        <table width="100%" border="0">
          <tr>
            <td>
              <%call ShowUserArticle(30)%>
              <Br>
              <%
		  if totalput>0 then
		  	call showpage(strFileName,totalput,MaxPerPage,false,true,"篇文章")
		  end if
		  %>
            </td>
          </tr>
        </table>
        <DIV align=center> </DIV></TD>
    </TR>
  </TBODY>
</TABLE>
<%' call Bottom() %>
<!--页面底部--> <div align="center">
  
    <TABLE id=table7 height=70 cellSpacing=0 cellPadding=0 width=797 
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
          <TD align=center width=781  valign="middle"
      bgColor=#ffffff>本网站由<font color="#3300FF"><a href="http://renwen.university.edu.cn/">university人文社会科学学院</a></font>主办、维护<BR>
            
             </TD>
        </TR>
      </TBODY>
    </TABLE>
  </div><!--结束页面底部-->
<% call PopAnnouceWindow(400,300) %>
</body>
</html>
<%
end if
rs.close
set rs=nothing
call CloseConn()
%>