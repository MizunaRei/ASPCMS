<!--#include file="Inc/syscode_diary.asp"-->
<%
PageTitle="心情日记内容"
set rs=server.createobject("adodb.recordset")

dim seetimes,ismyself,diaryID
diaryID=cint(Request("diaryID"))
ismyself=true		'初始化为自己的日记本

call getRndBg()		'取得随机背景

sql="SELECT top 1 * from diary where ID="&diaryID&""
rs.open sql,conn_User,1,3
if rs.eof and rs.bof then
	founderr=true
	errmsg="<br><br><li>你要查看的日记不存在！</li>"
else
	if rs("diaryBg")<>"0" then strRndBg="diary_images/back/"&rs("diaryBg")
	if rs("diaryOwner")<>CurrentLoginUser then ismyself=false
end if

%>

<html>
<head>
<title><%=strPageTitle & " >> " & PageTitle & " >> 心情日记内容" %></title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<!--#include file="inc/Skin_CSS.asp"-->
<%call MenuJS()%>
<script language=javascript>
	function opendelwin(diaryID)
	{
	var delok=confirm("确实要删除这则日记吗？");
	if (delok)
		{
		window.open("diary_del.asp?diaryID="+diaryID,"windel","width=200,height=10,top=250,left=350");
		}
	return false;
	}
</script>
</head>
<body <%=Body_Label%> onmousemove='HideMenu()'>
<!--#include file="top.asp"-->
<table width="760" border="0" align="center" cellpadding="0" cellspacing="0" class="tdbg"
	style="BACKGROUND-ATTACHMENT: fixed; BACKGROUND-IMAGE: url(<%=strRndBg%>); BACKGROUND-POSITION:center center;  BACKGROUND-REPEAT: no-repeat;scrollbar-track-color:#ffffff; SCROLLBAR-FACE-COLOR: #ffffff; FONT-SIZE: 9pt; SCROLLBAR-HIGHLIGHT-COLOR: #ffffff; SCROLLBAR-SHADOW-COLOR: #dddddd;  SCROLLBAR-3DLIGHT-COLOR: #dddddd; SCROLLBAR-ARROW-COLOR: #dddddd; FONT-FAMILY: "Verdana"; SCROLLBAR-DARKSHADOW-COLOR: #ffffff">
    <tr>

    <td align="center"><img src="diary_images/dia-b-title.gif" width="101" height="36">
      <table width="50%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td height="5"></td>
        </tr>
      </table>
      <table border="0" width="80%" cellspacing="0" cellpadding="0">
        <tr>
          <td nowrap><img border="0" src="diary_images/08.GIF" width="11" height="10"></td>
          <td align="right" nowrap>
		  <!--#include file="diary_manageBar.asp"--></td>
        </tr>
      </table>
      <br> <table width="90%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td background="diary_images/t-h_p-s.gif"><img src="diary_images/t-h_p-s.gif" width="26" height="12"></td>
        </tr>
      </table>
		<%
		if founderr=true then
			call writeerrmsg()
		else
			if not rs.eof then
				call showdiary()
				response.write ("<font color=#888888 face=Arial>[共访问&nbsp;"&seetimes&"&nbsp;次]</font>")
			end if
		end if
		rs.close
		%>
      <table width="420" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td background="diary_images/t-h_p-s.gif"><img src="diary_images/t-h_p-s.gif" width="26" height="12"></td>
        </tr>
      </table>
	  </td>
    </tr>
  </table>
<%
call bottom()
set rs=nothing
conn_User.close
set conn_User=nothing
%>
</body>
</html>

<%
sub showdiary()
	%>
	<table width="90%" border="0" cellspacing="2" cellpadding="2">
	<tr>
	  <td><img src="diary_images/dia-b-icon.gif" width="21" height="21"><b>
	  <%
	   response.write("作者："&rs("diaryOwner")&"&nbsp;&nbsp;")
	   response.write(FormatDateTime(rs("diaryDate"), 1))
		if DateDiff("d",rs("addtime"),date())<2 then response.write("&nbsp;<font color=red>New!</font>")%>
		&nbsp;&nbsp;&nbsp;&nbsp;<%=rs("weather")%>&nbsp;&nbsp;&nbsp;&nbsp;<img src='diary_images/<%=rs("mood")%>'>
		</b>
		<%
		if ismyself=true then
			response.write("&nbsp;&nbsp;&nbsp;&nbsp;<a href='' onclick='return opendelwin("""&rs("ID")&""")'><img src=diary_images/del.gif alt=删除这则日记 border=0></a>&nbsp;&nbsp;<a href=diary_modify.asp?diaryID="&rs("ID")&"><img src=diary_images/edit.gif alt=修改这则日记 border=0></a>")
		end if
		%></td>
	</tr>
	<tr>
	  <td><img src=diary_images/icon.gif height=8 width=24 border=0 hspace=0>
		<%
			if CurrentLoginUser=rs("diaryOwner") then
				response.write(rs("diaryContent"))
			else
				select case cint(rs("secret"))
					case 0
						call showdiaryContent()
					case 9
						if CurrentLoginUser<>empty then
							call showdiaryContent()
						else
							response.write("<font color=red>本则日记只对用户公开！</font>")
						end if
					case 99
						if CurrentLoginUser<>empty and instr(rs("readers"),"|"&CurrentLoginUser&"|")>0 then
							call showdiaryContent()
						else
							response.write("<font color=red>本则日记只对部分朋友公开！</font>")
						end if
					case else
						response.write("<font color=red>本则日记完全保密！</font>")
				end select
			end if%>
	  </td>
	</tr>
	</table>
	<br> <table width="90%" border="0" cellspacing="0" cellpadding="0">
	<tr>
	  <td background="diary_images/t-h_p-s.gif"><img src="diary_images/t-h_p-s.gif" width="26" height="12"></td>
	</tr>
	</table>
	<%
end sub

sub showdiaryContent()
	response.write(rs("diaryContent"))
	rs("seetimes")=rs("seetimes")+1
	rs.update
	seetimes=rs("seetimes")
end sub%>