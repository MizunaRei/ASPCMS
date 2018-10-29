<html>
<head>
<title>ÇëÑ¡ÔñÍ·Ïñ</title>
<script>
window.focus()
function changeimage(imagename)
{ 
	window.opener.document.formwrite.GuestImages.value=imagename;
	window.opener.document.formwrite.showimages.src="images/guestbook/"+imagename+".gif";
}
</script>
</head>
<body>
<table align=center width=95% class="table004" cellpadding=5><td class="table001">
<%

for i=1 to 9
	response.write "<img src='images/guestbook/0"&i&".gif' border=0 onclick=""changeimage('0"&i&"') "" style=cursor:hand> "
next
for i=10 to 22
	response.write "<img src='images/guestbook/"&i&".gif' border=0 onclick=""changeimage('"&i&"') "" style=cursor:hand> "
next
%>
</td></tr>

</table>
