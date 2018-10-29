<!--#include file="Inc/conn.asp"-->
<%
dim ArticleID,sql,rs
ArticleID=trim(request("ArticleID"))
if ArticleId="" then
	response.Redirect("Default.asp")
end if
sql="select * from article where ArticleID=" & ArticleID & ""
Set rs= Server.CreateObject("ADODB.Recordset")
rs.open sql,conn,1,3
if rs.bof and rs.eof then
	response.Write("<p>找不到文章</p>")
else	
%>
<html>
<head>
<title><%=rs("Title")%></title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="STYLE.CSS" rel="stylesheet" type="text/css">
</head>
<body leftmargin="2" topmargin="0" marginwidth="0" marginheight="0" oncontextmenu="return false" ondragstart="return false" onselectstart ="return false" onselect="document.selection.empty()" oncopy="document.selection.empty()" onbeforecopy="return false" onmouseup="document.selection.empty()">
<noscript><iframe src=*></iframe></noscript>
<table width="99%" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF" class="border" style="word-break:break-all">
  <tr> 
    <td> <table width="100%" border="0" cellspacing="4" cellpadding="1">
        <tr> 
          <td><table width=100% border=0 cellpadding=0 cellspacing=0 class="title">
              <tr> 
                <td width=2% height="20"> <div align="center">&nbsp;<img src="Images/ARROW2.GIF" width="6" height="7" align="absmiddle">&nbsp;</div></td>
                <td width=87%>您要打印的文件是：<font color="#CC0000"><%=rs("Title")%></font> 
                </td>
                <td width="11%" align=right><div align="center"><a href="javascript:window.print()"><img src=Images/printpage.gif alt="打印" width="16" height="16" border="0" align=absmiddle></a> 
                    <a href=javascript:window.print()>打印本文</a></div></td>
              </tr>
            </table></td>
        </tr>
      </table>
      <p align="center"><strong><font size="3"><%=rs("Title")%></font></strong><br>
        <br>
        作者：<%=rs("Author")%>&nbsp;&nbsp;&nbsp;&nbsp;转贴自：<%=rs("CopyFrom")%>&nbsp;&nbsp;&nbsp;&nbsp;点击数：<%=rs("Hits")%></p>
      <table width="98%" border="0" align="center" cellpadding="0" cellspacing="0">
        <tr> 
          <td background="images/bj4.gif" height="1"></td>
        </tr>
      </table>
      <br>
      <table width="100%" border="0" cellspacing="6" cellpadding="4">
        <tr> 
          <td><%=replace(rs("Content"),chr(34),"")%></td>
        </tr>
      </table></td>
  </tr>
</table>
</body>
</html>
<%
end if
rs.close
set rs=nothing
call CloseConn
%>