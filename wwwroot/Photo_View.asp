<!--#include file="Inc/syscode_Photo.asp"-->
<%
'请勿改动下面这三行代码
const ChannelID=4
PageTitle="查看图片"
FoundErr=False
if rs("PhotoLevel")<=999 then
	if UserLogined<>True then
		FoundErr=True
		ErrMsg=ErrMsg & "对不起，本图片为收费图片，要求至少是本站的注册用户才能欣赏！<br>您还没注册或者没有登录？所以不能欣赏本图片。请赶紧 <a href='User_Reg.asp'><font color=red><b>注册</b></font></a> 或 <a href='User_Login.asp'><font color=red><b>登录</a></font></a>吧！"
	else
		if UserLevel>rs("PhotoLevel") then
			FoundErr=True
			ErrMsg=ErrMsg & "对不起，本图片为收费图片，并且只有 <font color=blue>"
			if rs("PhotoLevel")=999 then
				ErrMsg=ErrMsg & "注册用户"
			elseif rs("PhotoLevel")=99 then
				ErrMsg=ErrMsg & "收费用户"
			elseif rs("PhotoLevel")=9 then
				ErrMsg=ErrMsg & "VIP用户"
			elseif rs("PhotoLevel")=5 then
				ErrMsg=ErrMsg & "管理员"
			end if
			ErrMsg=ErrMsg & "级别的用户</font> 才能欣赏。你目前的权限级别不够，所以不能欣赏。"
		else
			if ChargeType=1 and rs("PhotoPoint")>0 then
				if Request.Cookies("asp163")("Pay_Photo" & PhotoID)<>"yes" then
					if UserPoint<rs("PhotoPoint") then
						FoundErr=True
						ErrMsg=ErrMsg &"对不起，本图片为收费图片，并且欣赏本图片需要消耗 <b><font color=red>" & rs("PhotoPoint") & "</font></b> 点！"
						ErrMsg=ErrMsg &"而你目前只有 <b><font color=blue>" & UserPoint & "</font></b> 点可用。点数不足，无法欣赏本图片。请与我们联系进行充值。"
					else
						if lcase(trim(request("Pay")))="yes" then
							Conn_User.execute "update " & db_User_Table & " set " & db_User_UserPoint & "=" & db_User_UserPoint & "-" & rs("PhotoPoint") & " where " & db_User_Name & "='" & UserName & "'"
							response.Cookies("asp163")("Pay_Photo" & PhotoID)="yes"
						else
							FoundErr=True
							ErrMsg=ErrMsg & "<font color=red><b>注意</b></font>：欣赏本图片需要消耗 <font color=red><b>" & rs("PhotoPoint") & "</b></font>"
							ErrMsg=ErrMsg &"你目前尚有 <b><font color=blue>" & UserPoint & "</font></b> 点可用。阅读本文后，你将剩下 <b><font color=green>" & UserPoint-rs("PhotoPoint") & "</font></b> 点"
							ErrMsg=ErrMsg &"<br><br>你确实愿意花费 <b><font color=red>" & rs("PhotoPoint") & "</font></b> 点来欣赏本图片吗？"
							ErrMsg=ErrMsg &"<br><br><a href='Photo_View.asp?Pay=yes&UrlID=" & UrlID & "&PhotoID=" & PhotoID & "'>我愿意</a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href='index.asp'>我不愿意</a></p>"
						end if
					end if
				end if
			elseif ChargeType=2 then
				if ValidDays<=0 then
					FoundErr=True
					ErrMsg=ErrMsg & "<font color=red>对不起，本图片为收费图片，而您的有效期已经过期，所以无法欣赏本图片。请与我们联系进行充值。</font>"
				end if
			end if
		end if
	end if							
end if
if FoundErr=True then
	response.write ErrMsg
else
%>
<html>
<head>
<title><%=PhotoTitle%></title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<style type="text/css">
body {CURSOR: url('images/hmove.cur')}
</style>
<SCRIPT language=JavaScript>
drag = 0
move = 0
function init() {
    window.document.onmousemove = mouseMove
    window.document.onmousedown = mouseDown
    window.document.onmouseup = mouseUp
    window.document.ondragstart = mouseStop
}
function mouseDown() {
    if (drag) {
        clickleft = window.event.x - parseInt(dragObj.style.left)
        clicktop = window.event.y - parseInt(dragObj.style.top)
        dragObj.style.zIndex += 1
        move = 1
    }
}
function mouseStop() {
    window.event.returnValue = false
}
function mouseMove() {
    if (move) {
        dragObj.style.left = window.event.x - clickleft
        dragObj.style.top = window.event.y - clicktop
    }
}
function mouseUp() {
    move = 0
}
</SCRIPT>
<script language="JavaScript" type="text/JavaScript">
<!--
function MM_reloadPage(init) {  //reloads the window if Nav4 resized
  if (init==true) with (navigator) {if ((appName=="Netscape")&&(parseInt(appVersion)==4)) {
    document.MM_pgW=innerWidth; document.MM_pgH=innerHeight; onresize=MM_reloadPage; }}
  else if (innerWidth!=document.MM_pgW || innerHeight!=document.MM_pgH) location.reload();
}
MM_reloadPage(true);
//-->
</script>
</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" onload="init()">
<noscript><iframe src=*></iframe></noscript>
<%
dim UrlID
UrlID=trim(request("UrlID"))
if UrlID="" then
	UrlID=1
else
	UrlID=Cint(UrlID)
end if
if UrlID=1 then
	response.write "<div id='hiddenPic' style='position:absolute; left:433px; top:258px; width:77px; height:91px; z-index:1; visibility: hidden;'><img name='images2' src='" & rs("PhotoUrl") & "' border='0'></div>"
	response.write "<div id='block1' onmouseout='drag=0' onmouseover='dragObj=block1; drag=1;' style='z-index:10; height: 60; left: 0; position: absolute; top: 0; width: 120'><dd><img name='images1' src='" & rs("PhotoUrl") & "' border='0'></dd></div>"
else
	if UrlID>4 then
		response.write "地址参数错误！"
	else
		if rs("PhotoUrl" & UrlID)="" then
			response.write "地址错误!"
		else
			response.write "<div id='hiddenPic' style='position:absolute; left:433px; top:258px; width:77px; height:91px; z-index:1; visibility: hidden;'><img name='images2' src='" & rs("PhotoUrl" & UrlID) & "' border='0'></div>"
			response.write "<div id='block1' onmouseout='drag=0' onmouseover='dragObj=block1; drag=1;' style='z-index:10; height: 60; left: 0; position: absolute; top: 0; width: 120'><dd><img name='images1' src='" & rs("PhotoUrl" & UrlID) & "' border='0'></dd></div>"
		end if
	end if
end if
%>
</body>
</html>
<%
end if
rs.close
set rs=nothing
call CloseConn()
%>