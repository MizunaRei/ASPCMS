<!--#include file="inc/conn.asp"-->
<!--#include file="inc/Conn_User.asp"-->
<!--#include file="inc/config.asp"-->
<!--#include file="inc/function.asp"-->
<%
dim rs,sql
dim Action,FoundErr,ErrMsg
dim Subject,mailbody
dim ArticleID
dim MailToName,MailToAddress,FromName,MailFrom
dim ObjInstalled
ObjInstalled=IsObjInstalled("JMail.SMTPMail")
ArticleID=request("ArticleID")
Action=trim(request("Action"))
if ArticleID="" then
	foundErr = true
	ErrMsg=ErrMsg & "<br><li>请指定相关文章</li>"
else
	ArticleID=Clng(ArticleID)
end if
if CheckUserLogined()=False then 
	FoundErr=True
	ErrMsg=ErrMsg & "<br>&nbsp;&nbsp;&nbsp;&nbsp;你还没注册？或者没有登录？只有本站的注册用户才能使用“告诉好友”功能！<br><br>"
	ErrMsg=ErrMsg & "&nbsp;&nbsp;&nbsp;&nbsp;如果你还没注册，请赶紧<a href='User_Reg.asp'><font color=red>点此注册</font></a>吧！<br><br>"
	ErrMsg=ErrMsg & "&nbsp;&nbsp;&nbsp;&nbsp;如果你已经注册但还没登录，请赶紧<a href='User_Login.asp'><font color=red>点此登录</font></a>吧！<br><br>"
end if

if foundErr<>True then
	sql="Select Title,Content,UpdateTime,Author from article where articleid="&ArticleId&""
	set rs=server.createobject("adodb.recordset")
	rs.open sql,conn,1,1
	if rs.bof and rs.eof then
		FoundErr=True
		ErrMsg=ErrMsg & "找不到文章"
	else
		if Action="MailToFriend" then
			call MailToFriend()
		else
			call main()
		end if
	end if
	rs.close
	set rs=nothing
end if
if FoundErr=true then
	call WriteErrMsg()
end if
call CloseConn()

sub main()
%>
<HTML>
<HEAD>
<TITLE>告诉好友</TITLE>
<META content="text/html; charset=gb2312" http-equiv=Content-Type>
<link href="STYLE.CSS" rel="stylesheet" type="text/css">
</head>
<body>
<form action="sendmail.asp" method="post">
  <table cellpadding=2 cellspacing=1 border=0 width=400 class="border" align=center>
    <tr> 
      <td height="22" colspan=2 align=center valign=middle class="title"> <b>将本文告诉好友</b></td>
    </tr>
    <tr class="tdbg"> 
      <td width="120" align="right"><strong>收信人姓名：</strong></td>
      <td><input name="MailtoName" type="text" id="MailtoName" size="40" maxlength="20"> 
        <font color="#FF0000">*</font></td>
    </tr>
    <tr class="tdbg"> 
      <td width="120" align="right"><strong>收信人Email地址：</strong></td>
      <td><input name="MailToAddress" type=text id="MailToAddress" size="40" maxlength="100"> 
        <font color="#FF0000">*</font></td>
    </tr>
    <tr class="tdbg"> 
      <td width="120" height="60" align="right"><strong>文章信息：</strong></td>
      <td height="60">文章标题：<%= rs("Title") %><br>
        文章作者：<%= rs("Author") %> <br>
        发布时间：<%= rs("UpdateTime") %> </td>
    </tr>
    <tr class="tdbg"> 
      <td colspan=2 align=center><input name="Action" type="hidden" id="Action" value="MailToFriend"> 
        <input name="ArticleID" type="hidden" id="ArticleID" value="<%=request("ArticleID")%>"> 
        <input type=submit value=" 发 送 " name="Submit" <% If ObjInstalled=false Then response.write "disabled" end if%>> 
      </td>
    </tr>
    <%
If ObjInstalled=false Then
	Response.Write "<tr class='tdbg'><td height='40' colspan='2'><b><font color=red>对不起，因为服务器不支持 JMail组件! 所以不能使用本功能。</font></b></td></tr>"
End If
%>
  </table>
</form>
</BODY>
</HTML>
<%
end sub

sub MailToFriend()
	MailToName=trim(request.form("MailToName"))
	MailToAddress=trim(request.form("MailToAddress"))
	if MailToName="" then
		errmsg=errmsg & "<br><li>收信人姓名为空！</li>"
		founderr=true
	end if
	if IsValidEmail(MailToAddress)=false then
   		errmsg=errmsg & "<br><li>收信人的Email地址有错误！</li>"
   		founderr=true
	end if
				
	if founderr then
		exit sub
	end if
	
	call GetMailInfo()
	
	FromName=Request.Cookies("asp163")("UserName")
	dim trs
	set trs=conn_user.execute("select " & db_User_Name & "," & db_User_Email & " from " & db_User_Table & " where " & db_User_Name & "='" & FromName & "'")
	MailFrom=trs(1)
	set trs=nothing
	
	ErrMsg=SendMail(MailtoAddress,MailtoName,Subject,MailBody,FromName,MailFrom,3)
	if ErrMsg="" then
		call WriteSuccessMsg("已经成功将此文章发送给你的好友！")
	else
		FoundErr=True
	end if
end sub

sub GetMailInfo()
	Subject="您的朋友" & Addressee & "从" & SiteName & "给您发来的文章资料"

	mailbody=mailbody &"<style>A:visited {	TEXT-DECORATION: none	}"
	mailbody=mailbody &"A:active  {	TEXT-DECORATION: none	}"
	mailbody=mailbody &"A:hover   {	TEXT-DECORATION: underline overline	}"
	mailbody=mailbody &"A:link 	  {	text-decoration: none;}"
	mailbody=mailbody &"A:visited {	text-decoration: none;}"
	mailbody=mailbody &"A:active  {	TEXT-DECORATION: none;}"
	mailbody=mailbody &"A:hover   {	TEXT-DECORATION: underline overline}"
	mailbody=mailbody &"BODY   {	FONT-FAMILY: 宋体; FONT-SIZE: 9pt;}"
	mailbody=mailbody &"TD	   {	FONT-FAMILY: 宋体; FONT-SIZE: 9pt	}</style>"

	mailbody=mailbody &"<TABLE border=0 width='95%' align=center><TBODY><TR>"
	mailbody=mailbody &"<TD valign=middle align=top>"
	mailbody=mailbody &"--&nbsp;&nbsp;作者："&rs("Author")&"<br>"
	mailbody=mailbody &"--&nbsp;&nbsp;发布时间："&rs("UpdateTime")&"<br><br>"
	mailbody=mailbody &"--&nbsp;&nbsp;"&rs("title")&"<br>"
	mailbody=mailbody &""&rs("content")&""
	mailbody=mailbody &"</TD></TR></TBODY></TABLE>"

	mailbody=mailbody &"<center><a href='" & SiteUrl & "'>" & SiteName & "</a>"
	
end sub

%>