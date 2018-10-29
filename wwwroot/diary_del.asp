<!--#include file="Inc/syscode_diary.asp"-->
<html>
<head>
<title>删除日记</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href='style.css' rel='stylesheet' type='text/css'>
<script language=javascript>
	function closedelwin()
	{
		opener.window.location.reload();
		self.close();
	}
</script>
<style type="text/css">
<!--
body {
	font-size: 10pt;
	line-height: 150%;
}
-->
</style>
</head>
<body bgcolor="#FFFFFF" leftmargin="0" topmargin="0">
<%
set rs=server.createobject("adodb.recordset")
if CurrentLoginUser=empty then
	errmsg="<br><br><li>你还没有登录，没有此操作权限！</li>"
	call writeerrmsg()
	response.end
end if
dim diaryID
diaryID=Cint(request("diaryID"))

sqlUser="select top 1 * from diary where diaryOwner='"&CurrentLoginUser&"' and ID="&diaryID
rs.open sqlUser,conn_User,1,3
if rs.eof and rs.bof then
	errmsg="<br><br><li>您的日记本中没有这则日记，操作失败！</li>"
	call writeerrmsg()
	response.end
else
	rs.delete
end if
rs.close

sqlUser="update [User] set diaryNum=diaryNum-1 WHERE Username='"&CurrentLoginUser&"'"
conn_user.execute(sqlUser)
response.write ("<br><div align=center><font color=red>删除日记成功！</font><br><br>[<a href=# onclick='closedelwin()'>关闭窗口</a>]")
Set rs=Nothing
conn_user.close
set conn_user=nothing
%>
</body>
</html>