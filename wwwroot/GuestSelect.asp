<%
dim guestimagesnum,imagespath,emotnum,emotpath
imagespath="images/guestbook/"
emotpath="images/emot/"
guestimagesnum=23
emotnum=23

action=request("action")
select case action
	case "guestimages"
		PageTitle="请选择头像"
		call guestimages()
	case "emot"
		PageTitle="请选择表情"
		call emot()
end select

%>
<html>
<head>
<title><%=PageTitle%></title>
<script>
window.focus()
function changeimage(imagename)
{ 
	window.opener.document.formwrite.GuestImages.value=imagename;
	window.opener.document.formwrite.showimages.src="<%=imagespath%>"+imagename+".gif";
}
</script>
<link href="STYLE.CSS" rel="stylesheet" type="text/css">
</head>
<body>
<% sub guestimages()%>
<table align=center width=95% cellpadding=5><td>
<%

for i=1 to 9
	response.write "<img src='"&imagespath&"0"&i&".gif' border=0 onclick=""changeimage('0"&i&"') "" style=cursor:hand> "
next
for i=10 to guestimagesnum
	response.write "<img src='"&imagespath&""&i&".gif' border=0 onclick=""changeimage('"&i&"') "" style=cursor:hand> "
next
%>
</td></tr>
</table>
<%end sub%>

<% sub emot()%>
<table align=center width=95% cellpadding=5><td>
<%

for i=1 to emotnum
	response.write "<img src='"&emotpath&"emot"&i&".gif' border=0 onclick=""window.opener.document.formwrite.GuestContent.value+='[emot"&i&"]' "" style=cursor:hand> "
	if i mod 6 =0 then
		response.write "<br>"
	end if
next
%>
</td></tr></table>
<%end sub%>
<div align="center"><font size="2">[<a href="javascript:window.close();">关闭窗口</a>]</font></div>