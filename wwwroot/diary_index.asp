<!--#include file="Inc/syscode_diary.asp"-->
<%
PageTitle="心情日记首页"
strFileName="diary_index.asp?DiaryOwner="&DiaryOwner&"&diaryDate="&diaryDate
set rs=server.createobject("adodb.recordset")

dim diaryTitle,diaryVisit,ismyself,diaryDate
diaryDate=request("diaryDate")
if not isdate(diaryDate) then diaryDate=""

ismyself=true				'是否是自己的日记本

if request("page")<>"" then
    currentPage=cint(request("page"))
else
	currentPage=1
end if

if DiaryOwner=empty then
	call publicDiary()
else
	call privateDiary()
end if

call getRndBg()
%>


<html>
<head>
<title>坠落凡尘原创文学家园|日记本</title>
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
<div id=menuDiv style='Z-INDEX: 1000; VISIBILITY: hidden; WIDTH: 1px; POSITION: absolute; HEIGHT: 1px; BACKGROUND-COLOR: #9cc5f8'></div>

<table width="760" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr> 
    <td background="images/to_bj.gif" height="8"></td>
  </tr>
  <tr> 
    <td background="images/to_bj01.gif" height="8" bgcolor="#D98145"></td>
  </tr>
  <tr> 
    <td background="images/m_bg.gif" height="62">
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td height="11" width="234" align="right"><img src="images/Logofc.gif" width="213" height="62"></td>
          <td width="526" height="78"  align="center" valign="middle">
                
            <% call ShowBanner() %></td>
        </tr>
      </table>
    </td>
  </tr>
  <tr>
    <td background="images/to_bj03.gif" height="8" bgcolor="#D98145"></td>
  </tr>
</table>
<table width="760" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr> 
    <td height="19" background="images/daohang1.gif"> 
      <table width="683" border="0" cellspacing="0" cellpadding="0" align="center">
        <tr> 
         
          <td height="24" background="images/daohang.gif" width="68" align="center"> 
            <a href="diary_index.asp" target="_blank"><font color="#FF0000">日记本</font></a></td>
          <td height="24" background="images/daohang.gif" width="68" align="center"><a href="Article_Class2.asp?ClassID=2">散 
            文</a></td>
          <td height="24" background="images/daohang.gif" width="68" align="center"><a href="Article_Class2.asp?ClassID=1">小 
            说</a></td>
          <td height="24" background="images/daohang.gif" width="68" align="center"><a href="Article_Class2.asp?ClassID=4">杂 
            文</a></td>
          <td height="24" background="images/daohang.gif" width="68" align="center"><a href="Article_Class2.asp?ClassID=3">诗 
            歌</a></td>
          <td height="24" background="images/daohang.gif" width="68" align="center"><a href="Article_Class2.asp?ClassID=5">新文学</a></td>
       <td height="24" background="images/daohang.gif" width="68" align="center"><a href="Article_Class2.asp?ClassID=43">文学联盟</a></td>
          <td height="24" background="images/daohang.gif" width="68" align="center"><a href="photo_index.asp">读 
            图</a></td>
          <td height="24" background="images/daohang.gif" width="68" align="center"><a href="soft_index.asp">疯狂下载</a></td>
          <td height="24" background="images/daohang.gif" width="68" align="center"><a href="http://bbs.fanchen.com" target="_blank"><font color="#FF0000">交流社区</font></a></td>
        </tr>
      </table>
    </td>
  </tr>
  <tr> 
    <td background="images/to_bj08.gif" height="10" bgcolor="#F4DBCA"></td>
  </tr>
  <tr> 
    <td background="images/to_bj06.gif" height="9"></td>
  </tr>
</table><script language="javascript">
function click() {
if (event.button==2) {  //改成button==2为禁止右键
alert('欢迎来到【凡尘文学网】本站刚起步，请多支持……')
}
}
document.onmousedown=click
</script>

<table width="760" border="0" align="center" cellpadding="0" cellspacing="0" class="tdbg"
	style="BACKGROUND-ATTACHMENT: fixed; BACKGROUND-IMAGE: url(<%=strRndBg%>); BACKGROUND-POSITION:center center;  BACKGROUND-REPEAT: no-repeat;scrollbar-track-color:#ffffff; SCROLLBAR-FACE-COLOR: #ffffff; FONT-SIZE: 9pt; SCROLLBAR-HIGHLIGHT-COLOR: #ffffff; SCROLLBAR-SHADOW-COLOR: #dddddd;  SCROLLBAR-3DLIGHT-COLOR: #dddddd; SCROLLBAR-ARROW-COLOR: #dddddd; FONT-FAMILY: "Verdana"; SCROLLBAR-DARKSHADOW-COLOR: #ffffff">
	<tr>
	 
    <td height="144" valign="top"> 
      <%
		if founderr=true then
			call writeerrmsg()
			response.write("<br>&nbsp;")
		else
		%>
      <img src="diary_images/dia-b-title.gif" width="101" height="36"> 
      <table width="50%" border="0" cellspacing="0" cellpadding="0" align="center">
			<tr>
			  <td height="5"></td>
			</tr>
		  </table>
		  <table border="0" width="90%" cellspacing="0" cellpadding="0" align="center">
			<tr>
			  <td nowrap><img border="0" src="diary_images/08.GIF" width="11" height="10"><b>
				<%
				response.write(diaryTitle)
				if isdate(diaryDate) then response.write("("&diaryDate&")")%></b>
				</td>
			  <td align="right" nowrap>
			  <!--#include file="diary_manageBar.asp"--></td>
			</tr>
		  </table>
		  <br> <table width="90%" border="0" cellspacing="0" cellpadding="0" align="center">
			<tr>
			  <td background="diary_images/t-h_p-s.gif"><img src="diary_images/t-h_p-s.gif" width="26" height="12"></td>
			</tr>
		  </table>
			<center><%if rs.eof and rs.bof then
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

			if totalput>0 then
				call showpage(strFileName,totalput,MaxPerPage,true,false," 则日记")
			end if

			if DiaryOwner<>"" then
				response.write ("<font color=#888888 face=Arial>[共访问&nbsp;"&diaryVisit&"&nbsp;次]</font>")
			end if

			rs.close
			'set rs=nothing
	end if%></center>
      <table width="420" border="0" cellspacing="0" cellpadding="0" align="center">
        <tr>
          <td height="16" background="diary_images/t-h_p-s.gif"><img src="diary_images/t-h_p-s.gif" width="26" height="12"></td>
        </tr>
      </table>
      <br>
	  </td>
    </tr>
  </table>
<%
call bottom()
conn_User.close
set conn_User=nothing

%>
</body>
</html>
<%
sub ShowDiary()
	do while not rs.eof
		%>
		<table width="90%" border="0" cellspacing="2" cellpadding="2">
		<tr>
		  <td><img src="diary_images/dia-b-icon.gif" width="21" height="21"><b>
		  <%if DiaryOwner=empty then response.write("作者："&rs("diaryOwner")&"&nbsp;&nbsp;")
		   response.write(FormatDateTime(rs("diaryDate"), 1))
			if DateDiff("d",rs("addtime"),date())<2 then response.write("&nbsp;<font color=red>New!</font>")%>
			&nbsp;&nbsp;&nbsp;&nbsp;<%=rs("weather")%>&nbsp;&nbsp;&nbsp;&nbsp;<img src='diary_images/<%=rs("mood")%>'>
			</b>
			<%
			if CurrentLoginUser=rs("diaryOwner") then
				response.write("&nbsp;&nbsp;&nbsp;&nbsp;<a href='' onclick='return opendelwin("""&rs("ID")&""")'><img src=diary_images/del.gif alt=删除这则日记 border=0></a>&nbsp;&nbsp;<a href=diary_modify.asp?diaryID="&rs("ID")&"><img src=diary_images/edit.gif alt=修改这则日记 border=0></a>")
			end if
			%></td>
		</tr>
		<tr>
		  <td><img src=diary_images/icon.gif height=8 width=24 border=0 hspace=0>
			<%
			if CurrentLoginUser=rs("diaryOwner") then
				response.write("<a href=diary_show.asp?diaryID="&rs("ID")&" title=点击查看全部内容>"&left(rs("diaryContent"),100)&"</a>……")
			else
				select case cint(rs("secret"))
					case 0
						response.write("<a href=diary_show.asp?diaryID="&rs("ID")&" title=点击查看全部内容>"&left(rs("diaryContent"),100)&"</a>……")
					case 9
						if CurrentLoginUser<>empty then
							response.write("<a href=diary_show.asp?diaryID="&rs("ID")&" title=点击查看全部内容>"&left(rs("diaryContent"),100)&"</a>……")
						else
							response.write("<font color=red>本则日记只对用户公开！</font>")
						end if
					case 99
						if CurrentLoginUser<>empty and instr(rs("readers"),"|"&CurrentLoginUser&"|")>0 then
							response.write("<a href=diary_show.asp?diaryID="&rs("ID")&" title=点击查看全部内容>"&left(rs("diaryContent"),100)&"</a>……")
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
		rs.movenext
		i=i+1
		if i>=MaxPerPage then exit do
	loop
end sub

sub privateDiary()
	dim rsUser,sqlUser
	set rsUser=server.createobject("adodb.recordset")
	sqlUser="SELECT * from [user] where UserName='"&DiaryOwner&"'"
	rsUser.open sqlUser,conn_User,1,1
	if not rsUser.eof then
		diaryVisit=rsUser("diaryVisit")

		if CurrentLoginUser<>DiaryOwner then
			sqlUser="update [user] set diaryvisit=diaryvisit+1 WHERE userName='"&DiaryOwner&"'"
			conn_User.execute(sqlUser)
			ismyself=false
		end if

		if ismyself=true then
			diaryTitle="我的日记本"
		else
			diaryTitle=DiaryOwner&"&nbsp;的公开日记本"
		end if

		if diaryDate="" then
			sql="SELECT * from diary where diaryOwner='"&DiaryOwner&"' order by diaryDate desc"
		else
			sql="SELECT * from diary where diaryOwner='"&DiaryOwner&"' and diaryDate=#"&diaryDate&"# order by diaryDate desc"
		end if

		rs.open sql,conn_User,1,1
	else
		founderr=true
		errmsg="<br><br><li>您访问的用户不存在!</li>"
	end if
	rsUser.close
	set rsUser=nothing
end sub

sub PublicDiary()
	ismyself=false
	if diaryDate="" then
		sql="SELECT * from diary order by diaryDate desc"
	else
		sql="SELECT * from diary where diaryDate=#"&diaryDate&"# order by diaryDate desc"
	end if
	diaryTitle="坠落凡尘文学家园个人日记本"
	rs.open sql,conn_User,1,1
end sub
%>