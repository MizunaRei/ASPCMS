<!--#include file="inc/conn.asp"-->
<!--#include file="inc/ubbcode.asp"-->
<%
SkinID=0
%>
<html>
<head>
<title>留言预览</title>
<!--#include file="inc/Skin_CSS.asp"-->
</head>
<body>
<table width="100%" border="0" cellpadding="0" cellspacing="0" class="border">
<tr class="tdbg_leftall"> 
<tr > 
  <td align="center" valign="top"> 
	<table width="100%" border="0" cellspacing="0" cellpadding="0" height=22 class=title>
	  <tr> 
		<td ><font color=green>&nbsp;&nbsp;主题</font>:&nbsp;<%=request.form("title")%></td>
		<td width="200"><img src="<%=GuestPath%>images/guestbook/posttime.gif" width="11" height="11" align="absmiddle"><font color="#006633">&nbsp; 
		  <% =now()%>
		  </font> </td>
	  </tr>
	</table>
  </td>
</tr>
<tr class="tdbg_leftall"> 
  <td height="153" valign="top"> 
	<%
		response.write ubbcode(request.form("content")) 
	%>
  </td>
</tr>
</table>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
<tr> 
  <td  height="15" align="center" valign="top"> 
	<table width="98%" border="0" align="center" cellpadding="0" cellspacing="0">
	  <tr> 
		<td height="13" class="tdbg_left2"></td>
	  </tr>
	</table>
  </td>
</tr>
</table>
<div align="center">[<a href="javascript:window.close();">关闭窗口</a>] </div>