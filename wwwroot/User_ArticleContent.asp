<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%
option explicit
response.buffer=true	
'强制浏览器重新访问服务器下载页面，而不是从缓存读取页面
Response.Buffer = True 
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1 
Response.Expires = 0 
Response.CacheControl = "no-cache" 
%>
<!--#include file="inc/conn.asp"-->
<!--#include file="inc/Conn_User.asp"-->
<!--#include file="inc/function.asp"-->
<%
if CheckUserLogined()=False then
	response.write "你还没有登录，不能使用此功能！"
	response.end
end if
%>
<html>
<head>
<meta http-equiv=Content-Type content="text/html; charset=gb2312">
<style type="text/css">body {font-size:	9pt}</style>
</head>
<BODY bgcolor="#FFFFFF" MONOSPACE>
<%
dim Action,FoundErr,ErrMsg
dim ArticleID,sql,rs
Action=trim(request("Action"))
ArticleID=trim(request("ArticleID"))
if Action="Modify" then
	if ArticleId="" then
		response.write "请指定要修改的文章ID"
	else
		ArticleID=Clng(ArticleID)
		sql="select * from article where Deleted=False and Passed=False and Editor='" & UserName & "' and ArticleID=" & ArticleID & ""
		Set rs= Server.CreateObject("ADODB.Recordset")
		rs.open sql,conn,1,1
		if rs.bof and rs.eof then
			response.write "找不到文章"
		else
			response.write "<p>" & rs("Content") & "</p>"
		end if
		rs.close
		set rs=nothing
	end if
end if
%>
</body>
</html>
<%
call CloseConn()
call CloseConn_User()
%>