<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%
option explicit
response.buffer=true
Const PurviewLevel=2
Const CheckChannelID=2
Const PurviewLevel_Article=3
%>
<!--#include file="inc/conn.asp"-->
<!--#include file="inc/config.asp"-->
<!--#include file="Admin_ChkPurview.asp"-->
<!--#include file="inc/function.asp"-->
<!--#include file="inc/admin_code_Article.asp"-->
<%
dim ArticleID,Action,sql,rs,rsArticle,FoundErr,ErrMsg,PurviewChecked
ArticleID=trim(request("ArticleID"))
if ArticleID="" then
	FoundErr=True
	ErrMsg=ErrMsg & "<br><li>请指定要修改的文章ID</li>"
	call WriteErrMsg()
	call CloseConn()
	response.end
else
	ArticleID=Clng(ArticleID)
end if

sql="select * from article where ArticleID=" & ArticleID & ""
Set rsArticle= Server.CreateObject("ADODB.Recordset")
rsArticle.open sql,conn,1,1
if rsArticle.bof and rsArticle.eof then
	FoundErr=True
	ErrMsg=ErrMsg & "<br><li>找不到文章</li>"
else	
	if rsArticle("Passed")=True then
			FoundErr=True
			ErrMsg=ErrMsg & "<br><li>文章已经被审核通过，你不能再进行修改！</li>"
	end if
end if


if FoundErr=True then
	call WriteErrMsg()
else
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>退稿原因</title>
<link rel="stylesheet" type="text/css" href="Admin_style.css">
</head>
<body leftmargin="0" topmargin="0">
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="border">
  <tr class="topbg">
    <td height="22" colspan="2"  align="center">该文章被退稿的原因：</td>
  </tr>
<table>

<form method="POST" name="myform" action="Admin_ArticleProperty.asp">
  <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
    <tr align="center"> 
      <td height="81" class="tdbg"> 
        <table width="100%" border="0" cellpadding="0" cellspacing="0" align="center">
          <tr class="tdbg" align="center"> 
              
           
              <td width="502" align="center"><textarea name="nopasstxt" cols="70" rows="5" id="nopasstxt"><%=rsArticle("nopasstxt")%></textarea>
		  <input name="ArticleID" type="hidden" id="ArticleID" value="<%=ArticleID%>">
		  <input name="Action" type="hidden" id="Action" value="nopassed">
		</td>
            </tr>
	  </table>
       </td>
  </tr>
  <tr align="center"> 
  <td class="tdbg"><input name="Save" type="submit"  id="Save" value="确认" style="cursor:hand;">
		   
	</td></tr>
</table>
</form>
</body>
</html>
<%
end if
rsArticle.close
set rsArticle=nothing
call CloseConn()
%>