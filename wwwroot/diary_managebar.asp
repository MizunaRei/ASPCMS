|&nbsp;<a href='' onclick='window.open("diary_calendar.asp?DiaryOwner=<%=DiaryOwner%>","win_user","width=480,height=330,left=200,top=120");return false;' title=按日历选择查看日记内容>选择日记时间</a>
|&nbsp;<a href="diary_index.asp?DiaryOwner=">凡尘公开日记</a>
<%if CurrentLoginUser<>empty then
	response.write("|&nbsp;<a href=diary_index.asp?DiaryOwner="&CurrentLoginUser&">查看我的日记</a>&nbsp;")
	response.write("|&nbsp;<a href=diary_add.asp title=记下自己的心情日记>写日记</a>")
end if%>
|&nbsp;<a href="userlist.asp" title=看看其他人的日记 target="_blank">个人日记列表</a>
<%if CurrentLoginUser=empty then
	response.write("|&nbsp;<a href="&sysRegFile&" title=注册申请你自己的日记本>还未注册凡尘</a>&nbsp;")
	response.write("|&nbsp;<a href="&sysLoginFile&" title=登录管理你自己的日记本>尚未登录</a>")
else
	response.write("|&nbsp;<a href="&sysLogoutFile&" title=退出自己的日记本>退出登录</a>")
end if%>
|