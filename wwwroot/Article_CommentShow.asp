<!--#include file="inc/conn.asp"-->
<!--#include file="inc/Conn_User.asp"-->
<%
dim ArticleID
ArticleID=trim(request("ArticleID"))
if ArticleId="" then
	call CloseConn()
	response.Redirect("index.asp")
else
	ArticleID=CLng(ArticleID)
end if
dim rsComment,sqlComment
sqlComment="select C.ArticleID,C.ClassID,C.UserType,C.UserName,C.Email,C.Oicq,C.Homepage,C.WriteTime,C.Score,C.Content,C.ReplyContent,C.ReplyName,C.ReplyTime,A.Title,A.UpdateTime,A.ArticleID from ArticleComment C inner join Article A on C.ArticleID=A.ArticleID where C.ArticleID=" & ArticleID & " order by C.CommentID desc"
Set rsComment= Server.CreateObject("ADODB.Recordset")
rsComment.open sqlComment,conn,1,1
%>
<html>
<head>
<title>所有评论</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="STYLE.CSS" rel="stylesheet" type="text/css">
</head>
<body leftmargin="2" topmargin="0" marginwidth="0" marginheight="0" style="BACKGROUND-IMAGE: url(SkinIndex/bg_all.gif)" background="SkinIndex/bg_all.gif" ><%
if rsComment.bof and rsComment.eof then
	response.write "&nbsp;&nbsp;&nbsp;&nbsp;没有任何评论"
else
%><table width="760" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr> 
    <td width="220" height="8" background="images/to_bj.gif"></td>
  </tr>
</table>
<table width="760" border="0" cellspacing="0" cellpadding="0" align="center" bgcolor="F7F6ED">
  <tr> 
    <td width="190" height="55" align="center">&nbsp;</td>
    <td height="55" valign="bottom"> 
      <table width="96%" border="0" cellspacing="0" cellpadding="0" align="center">
        <tr> 
          <td>&nbsp; 
            <div align="center">&nbsp;查看<font size="3" color="#FF0000"><strong> 
              <%response.write "<b>" & rsComment("Title")& "</b> ["& formatdatetime(rsComment("UpdateTime"),2) &"]"%>
              </strong></font>&nbsp;的所有评论</div>
          </td>
        </tr>
        <tr>
          <td>
          </td>
        </tr>
       
        <tr> 
          <td height="5"></td>
        </tr>
      </table>
    </td>
  </tr>
</table>
<table width="760" border="0" cellspacing="0" cellpadding="0" align="center" bgcolor="#F7F6ED">
  <tr> 
    <td colspan="3">
      <table width="100%" border="0" cellspacing="0" cellpadding="0"开獕牥也浡???	??????屐?屐???>
        <tr> 
          <td width="198">&nbsp;</td>
          <td background="images/title_line.gif">&nbsp;</td>
          <td width="20">&nbsp;</td>
        </tr>
      </table>
    </td>
  </tr>
  <tr> 
    <td width="187" height="200" valign="top"> 
      <table width="100%" border="0" cellspacing="0" cellpadding="0" height="100%">
        <tr> 
          <td width="187" height="190" valign="top"> 
            <table width="91%" border="0" cellspacing="0" cellpadding="0" align="center" height="100%">
              <tr> 
                <td valign="top" height="19" align="center"> <img src="NewImages/StudentArticleTop.png" width="150" height="150">                </td>
              </tr>
              <tr> 
                <td align="center" valign="top"> <br>
                  <br>
                </td>
              </tr>
            </table>
          </td>
        </tr>
      </table>
    </td>
    <td valign="top" width="3" background="images/07bj.gif"></td>
    <td valign="top">
      <table width="92%" border="0" cellspacing="0" cellpadding="0" align="center">
        
        <tr> 
          <td align="center" height="3"></td>
        </tr>
        <tr> 
          <td class="editorword"> 
            <%
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
%>
          </td>
        </tr>
        
      </table>
    </td>
  </tr>
</table>
<table width="760" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr> 
    <td background="images/04bj.gif" height="15" colspan="2">&nbsp;</td>
  </tr>
  <tr bgcolor="E9EAE5"> 
 
  <tr bgcolor="E2E3DE" align="center"> 
    <td colspan="2" height="22">【<a href='javascript:onclick=history.go(-1)'>返回文章内容页</a>】</td>
  </tr>
</table>
<%end if%>
</body>
</html>
<%
call CloseConn()
%>