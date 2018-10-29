<!--#include file="inc/syscode_diary.asp"-->
<%
Const PurviewLevel=2
Const CheckChannelID=0
Const PurviewLevel_Others="User"
%>
<!--#include file="Admin_ChkPurview.asp"-->
<%
set rs=server.createobject("adodb.recordset")
strFileName="diary_index.asp?DiaryOwner="&DiaryOwner
dim act,diaryID,nosecret
act=request("act")
diaryID=request("diaryID")
select case act
	case "delall1"
		sql="SELECT count(id) from diary where diaryOwner='"&DiaryOwner&"' and secret<=9"
		rs.open sql,conn_User,1,1
		i=rs(0)
		rs.close
		sqlUser="update [User] set diaryNum=diaryNum-"&i&" WHERE Username='"&DiaryOwner&"'"
		conn_user.execute(sqlUser)
		sqluser="delete from diary where diaryOwner='"&DiaryOwner&"' and secret<=9"
		conn_user.execute(sqlUser)

	case "delall2"
		sqlUser="update [User] set diaryNum=0 WHERE Username='"&DiaryOwner&"'"
		conn_user.execute(sqlUser)
		sqluser="delete from diary where diaryOwner='"&DiaryOwner&"'"
		conn_user.execute(sqlUser)
	case "del"
		sqlUser="update [User] set diaryNum=diaryNum-1 WHERE Username='"&DiaryOwner&"'"
		conn_user.execute(sqlUser)
		sqluser="delete from diary where diaryOwner='"&DiaryOwner&"' and id="&diaryID
		conn_user.execute(sqlUser)
	case "secret"
		sqluser="update diary set secret=999 where diaryOwner='"&DiaryOwner&"' and id="&diaryID
		conn_user.execute(sqlUser)
end select
sql="SELECT * from diary where diaryOwner='"&DiaryOwner&"' order by diaryDate desc"
rs.open sql,conn_User,1,1
if rs.eof then
	founderr=true
	errmsg="<br><br><li>该用户没有任何日记</li>"
end if
%>
<html>
<head>
<title>用户公开日记管理</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="Admin_Style.css" rel="stylesheet" type="text/css">
</head>
<body leftmargin="2" topmargin="0" marginwidth="0" marginheight="0">
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" Class="border">
  <tr class="topbg">
    <td height="22" colspan=2 align=center><strong>注 册 用 户 管 理</strong></td>
  </tr>
  <form name="form1" action="Admin_User.asp" method="get">
    <tr class="tdbg">
      <td width="100" height="30"><strong>快速查找用户：</strong></td>
      <td width="687" height="30"><select size=1 name="UserSearch" onChange="javascript:submit()">
          <option value="0">列出所有用户</option>
          <option value="1">文章最多TOP100</option>
          <option value="2">文章最少的100个用户</option>
          <option value="3">最近24小时内登录的用户</option>
          <option value="4">最近24小时内注册的用户</option>
          <option value="5">等待邮件验证的用户</option>
          <option value="6">等待管理员认证的用户</option>
          <option value="7">所有被锁住的用户</option>
          <option value="8">所有收费用户</option>
          <option value="9">所有VIP用户</option>
        </select>
        &nbsp;&nbsp;&nbsp;&nbsp;<a href="Admin_User.asp">用户管理首页</a>&nbsp;|&nbsp;<a href="Admin_User.asp?Action=Add">添加新用户</a></td>
    </tr>
  </form>
</table>
<br>
<table width="100%" border="0" cellspacing="1" cellpadding="2" class="border">
  <tr align="center" class="title">
    <td height="22"><strong>用&nbsp;户&nbsp;公&nbsp;开&nbsp;日&nbsp;记&nbsp;管&nbsp;理</strong></td>
  </tr>
  <tr class="tdbg">
    <td align="center">
        <%
		if founderr=true then
			call writeerrmsg()
			response.write("<br>&nbsp;")
		else
		%>
      <table width="50%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td height="5"></td>
        </tr>
      </table>
      <table border="0" width="80%" cellspacing="0" cellpadding="0">
        <tr>
          <td nowrap><img border="0" src="diary_images/08.GIF" width="11" height="10"><b>
            <%response.write(rs("diaryOwner")&"的公开日记")%>
            </b> &nbsp;&nbsp;&nbsp;&nbsp;[<a href=diary_userdiarylist.asp?diaryOwner=<%=diaryOwner%>&act=delall1 onclick='return confirm("您确认要删除该用户的全部公开日记吗？")'><font color=red>删除公开日记</font></a>]
            &nbsp;&nbsp;&nbsp;&nbsp;[<a href=diary_userdiarylist.asp?diaryOwner=<%=diaryOwner%>&act=delall2 onclick='return confirm("您确认要删除该用户的所有日记吗？")'><font color=red>删除全部日记</font></a>]
          </td>
          <td align="right" nowrap>&nbsp;
          </td>
        </tr>
      </table>
			<%if rs.eof and rs.bof then
				response.write "<p>还没有一则日记呢！</p>"
				totalput=0
			else
				totalput=rs.recordcount
				if currentpage<1 then
					currentpage=1
				end if
				if (currentpage-1)*MaxPerPage>totalput then
					if (totalPut mod MaxPerPage)=0 then
						currentpage= totalPut \ MaxPerPage
					else
						currentpage= totalPut \ MaxPerPage + 1
					end if
				end if

				if currentPage=1 then
					call ShowDiary()
				else
					if (currentPage-1)*MaxPerPage<totalPut then
						rs.move  (currentPage-1)*MaxPerPage
						dim bookmark
						bookmark=rs.bookmark
						call ShowDiary()
					else
						currentPage=1
						call ShowDiary()
					end if
				end if
			end if

			response.write ("<br>该用户共有<font color=red>&nbsp;"&totalput&"&nbsp;</font>则，公开日记<font color=red>&nbsp;"&nosecret&"&nbsp;</font>则，保密日记<font color=red>&nbsp;"&totalput-nosecret&"&nbsp;</font>则")

			if nosecret>MaxPerPage then
				call showpage(strFileName,totalput,MaxPerPage,false,false," 则日记")
			end if

			rs.close
			set rs=nothing
	end if%>
      <table width="420" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td background="diary_images/t-h_p-s.gif"><img src="diary_images/t-h_p-s.gif" width="26" height="12"></td>
        </tr>
      </table></td>
  </tr>
</table>
</body>
</html>
<%
sub ShowDiary()
	nosecret=0
	do while not rs.eof
		if rs("secret")<=9 then
			nosecret=nosecret+1%>
			<table width="90%" border="0" cellspacing="2" cellpadding="2">
			<tr>
			  <td><img src="diary_images/dia-b-icon.gif" width="21" height="21"><b>
			  <%response.write(FormatDateTime(rs("diaryDate"), 1))%>
				&nbsp;&nbsp;&nbsp;&nbsp;<%=rs("weather")%>&nbsp;&nbsp;&nbsp;&nbsp;<img src='diary_images/<%=rs("mood")%>'>
				</b>
				&nbsp;&nbsp;&nbsp;&nbsp;<a href=diary_userdiarylist.asp?diaryOwner=<%=diaryOwner%>&diaryID=<%=rs("ID")%>&act=del onclick='return confirm("您确认要删除本则日记吗？")'>删除</a>&nbsp;|&nbsp;<a href=diary_userdiarylist.asp?diaryOwner=<%=diaryOwner%>&diaryID=<%=rs("ID")%>&act=secret onclick='return confirm("您确认要把本则日记设为保密吗？")'>保密</a>
				</td>
			</tr>
			<tr>
			  <td><img src=diary_images/icon.gif height=8 width=24 border=0 hspace=0>
				<%=rs("diaryContent")%>
			  </td>
			</tr>
			</table>
			<table width="90%" border="0" cellspacing="0" cellpadding="0">
			<tr>
			  <td background="diary_images/t-h_p-s.gif"><img src="diary_images/t-h_p-s.gif"></td>
			</tr>
			</table>
			<%
		end if
		rs.movenext
		i=i+1
		if i>=MaxPerPage then exit do
	loop
end sub
%>