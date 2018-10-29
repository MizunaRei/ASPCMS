<!--#include file="inc/conn.asp"-->
<!--#include file="inc/Conn_User.asp"-->
<%
dim PhotoID
PhotoID=trim(request("PhotoID"))
if PhotoId="" then
	call CloseConn()
	response.Redirect("index.asp")
else
	PhotoID=CLng(PhotoID)
end if
dim rsComment,sqlComment
sqlComment="select C.PhotoID,C.ClassID,C.UserType,C.UserName,C.Email,C.Oicq,C.Homepage,C.WriteTime,C.Score,C.Content,C.ReplyContent,C.ReplyName,C.ReplyTime,P.PhotoName,P.UpdateTime,P.PhotoID from PhotoComment C inner join Photo P on C.PhotoID=P.PhotoID where C.PhotoID=" & PhotoID & " order by C.CommentID desc"
Set rsComment= Server.CreateObject("ADODB.Recordset")
rsComment.open sqlComment,conn,1,1
%>
<html>
<head>
<title>所有评论</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="STYLE.CSS" rel="stylesheet" type="text/css">
</head>
<body leftmargin="2" topmargin="0" marginwidth="0" marginheight="0">
<%
if rsComment.bof and rsComment.eof then
	response.write "&nbsp;&nbsp;&nbsp;&nbsp;没有任何评论"
else
%>
<table width="760" border="0" align="center" cellpadding="5" cellspacing="0" class="border">
  <tr class="title">
    <td align="center"><%response.write "<b>" & rsComment("PhotoName")& "</b> ["& formatdatetime(rsComment("UpdateTime"),2) &"]"%></td>
  </tr>
  <tr class="title"> 
    <td align="center"><font color=red>（评论内容只代表网友观点，与本站立场无关！）</font></td>
  </tr>
  <tr class="tdbg"> 
    <td> <%
	response.write "<table width='100%' border='0' cellspacing='0' cellpadding='0'>"
	do while not rsComment.eof
		response.write "<tr><td width='70%'>"
		if rsComment("UserType")=1 then
			response.write "<li>会员"
			set rsCommentUser=Conn_User.execute("select " & db_User_ID & "," & db_User_Name & "," & db_User_Email & "," & db_User_QQ & "," & db_User_Homepage & " from " & db_User_Table & " where " & db_User_Name & "='" & rsComment("UserName") & "'")
			if rsCommentUser.bof and rsCommentUser.eof then
				response.write rsComment("UserName")
			else
				response.write "『<a href='UserInfo.asp?UserID=" & rsCommentUser(0) & "' title='姓名：" & rsCommentUser(1) & vbcrlf & "信箱：" & rsCommentUser(2) & vbcrlf & "Oicq：" & rsCommentUser(3) & vbcrlf & "主页：" &  rsCommentUser(4)&"'><font color='blue'>" & rsComment("UserName") & "</font></a>』"
			end if
		else
			response.write "<li>游客『<span title='姓名：" & rsComment("UserName") & vbcrlf & "信箱：" & rsComment("Email") & vbcrlf & "Oicq：" & rsComment("Oicq") & vbcrlf & "主页：" &  rsComment("Homepage")&"' style='cursor:hand'><font color='blue'>" & rsComment("UserName") & "</font></span>』"
		end if
		response.write "于" & rsComment("WriteTime") & "发表评论：</li>"
		response.write "</td><td align=right>评分："&rsComment("Score")&"分</td></tr>"
		response.write "<tr><td colspan='2'>"
		response.write "&nbsp;&nbsp;&nbsp;&nbsp;" & rsComment("Content") & "<br>"
		if rsComment("ReplyContent")<>"" then
			response.write "&nbsp;&nbsp;&nbsp;&nbsp;<font color='009900'>★</font>&nbsp;管理员『<font color='blue'>" & rsComment("ReplyName") & "</font>』于 " & rsComment("ReplyTime") & " 回复道：&nbsp;&nbsp;&nbsp;&nbsp;" & rsComment("ReplyContent") & "<br>"			
		end if
		response.write "<br></td></tr>"
		rsComment.movenext
	loop
	response.write "</td></tr></table>"
%> </td>
  </tr>
  <tr class="tdbg">
    <td align="center">【<a href='javascript:onclick=history.go(-1)'>返回图片内容页</a>】</td>
  </tr>
</table>
<%end if%>
</body>
</html>
<%
call CloseConn()
%>