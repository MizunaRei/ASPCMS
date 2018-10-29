<%@language=vbscript codepage=936 %>
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
<!--#include file="conn.asp"-->
<!--#include file="Conn_User.asp"-->
<!--#include file="config.asp"-->
<!--#include file="function.asp"-->
<%
dim strChannel,sqlChannel,rsChannel,ChannelUrl,ChannelName
dim strFileName,MaxPerPage,totalPut,CurrentPage,TotalPages
dim BeginTime,EndTime,founderr,errmsg,i
dim rs,sql,rsUser,sqlUser
dim PageTitle,strPath,strPageTitle
dim SkinID,ClassID,AnnounceCount
'***********************************************************************************************
strPath= "&nbsp;您现在的位置：&nbsp;<a href='" & SiteUrl & "'>" & SiteName & "</a>"
strPageTitle= SiteTitle
if ShowSiteChannel="Yes" then
	strChannel= "|&nbsp;"
	sqlChannel="select * from Channel order by OrderID"
	set rsChannel=server.CreateObject("adodb.recordset")
	rsChannel.open sqlChannel,conn,1,1
	do while not rsChannel.eof
		if rsChannel("ChannelID")=ChannelID then
			ChannelUrl=rsChannel("LinkUrl")
			ChannelName=rsChannel("ChannelName")
			strChannel=strChannel & "<a href='" & ChannelUrl & "'><font color=red>" & ChannelName & "</font></a>&nbsp;|&nbsp;"
		else
			strChannel=strChannel & "<a href='" & rsChannel("LinkUrl") & "'>" & rsChannel("ChannelName") & "</a>&nbsp;|&nbsp;"
		end if
		rsChannel.movenext
	loop
	rsChannel.close
	set rsChannel=nothing
	if CurrentLoginUser<>empty then
			strChannel=strChannel & "<a href='' onclick='window.open(""userinfo_center.asp"",""win_user"",""width=180,height=300,left=580,top=120"");return false;'><font color=red>用户控制面板</font></a>&nbsp;|&nbsp;"
	end if
	strPath=strPath & "&nbsp;&gt;&gt;&nbsp;<a href='" & ChannelUrl & "'>" & ChannelName & "</a>"
	strPageTitle=strPageTitle & " >> " & ChannelName
end if
BeginTime=Timer
ClassID=0

' 请勿改动下面这三行代码
const ChannelID=8
Const ShowRunTime="Yes"
MaxPerPage=8
SkinID=0

'========== 山风多用户日记本插件系统参数，请根据你的系统实际情况修改 ============'
dim sysRegFile, sysLoginFile, sysLogoutFile, CurrentLoginUser, DiaryOwner
sysRegFile		= "user_reg.asp"												'定义用户注册文件
sysLoginFile	= "user_login.asp"												'定义用户登录文件
sysLogoutFile	= "user_logout.asp"												'定义用户退出文件
CurrentLoginUser=request.cookies("asp163")("username")							'当前已经登录用户名
Const bgNum=30																	'背景图片数量








'========= 山风多用户日记本插件系统代码，如果对ASP不是很熟悉，请勿更改 =========
DiaryOwner=trim(request("DiaryOwner"))
if instr(DiaryOwner,"'")>0 or instr(DiaryOwner," ")>0 then
	response.write ("用户名非法！")
	response.end
end if

'==================================
'过程名：getRndBg()
'功　能：获得随机背景
'==================================
dim strRndBg
function getRndBg()
	Randomize() '初始化随机数生成器。
	strRndBg = 1+Int((bgNum-1) * Rnd)
	if len(strRndBg)=1 then strRndBg="0"&strRndBg
	strRndBg="diary_images/back/"&strRndBg&".gif"
end function
%>

