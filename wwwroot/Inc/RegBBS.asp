<%
dim rsForumTitle,rsForumInfo,rsForumFace,Forum_info,Forum_Setting,Forum_user,Forum_userface
dim FU_UserClass,FU_TitlePic,FU_UserGroup,FU_UserGroupID,FU_Showre
dim FU_Face,FU_FaceWidth,FU_FaceHeight,FU_UserWealth,FU_UserEP,FU_UserCP

if UserTableType="Dvbbs6.0" or UserTableType="Dvbbs6.1" then
	if UserTableType="Dvbbs6.0" then
		set rsForumInfo=Conn_User.execute("select top 1 Forum_Info,Forum_setting,Forum_user,Forum_userface from config order by id asc")
		Forum_userface=split(rsForumInfo("Forum_userface"),"|")
	elseif UserTableType="Dvbbs6.1" then
		set rsForumInfo=Conn_User.execute("select Forum_Info,Forum_setting,Forum_user from dvbbs_info where active=1")
		set rsForumFace=Conn_User.execute("select Forum_userface from dvbbs_pic where active=1")
		Forum_userface=split(rsForumFace("Forum_userface"),"|")
	end if
	Forum_info=split(rsForumInfo("Forum_info"),",")
	Forum_Setting=split(rsForumInfo("Forum_setting"),",")
	Forum_user=split(rsForumInfo("Forum_user"),",")
	set rsForumTitle=Conn_User.execute("select usertitle,titlepic from usertitle where not minarticle=-1 order by minarticle")

	FU_UserClass = rsForumTitle(0)
	FU_TitlePic = rsForumTitle(1)
	if cint(Forum_Setting(25))=1 then
		FU_UserGroupID=5
	else
		FU_UserGroupID=4
	end if
	FU_Face = Forum_info(11)&Forum_userface(0)
	FU_FaceWidth = Forum_setting(38)
	FU_FaceHeight = Forum_setting(39)
	FU_UserWealth = Forum_user(0)
	FU_UserEP = Forum_user(5)
	FU_UserCP = Forum_user(10)
	FU_UserGroup = "нчценчеи"
	FU_Showre = 1
end if

sub UpdateUserNum(RegUserName)
	if UserTableType="Dvbbs6.0" then
		conn_user.execute("update config set usernum=usernum+1,lastuser='"&RegUserName&"'")
	elseif UserTableType="Dvbbs6.1" then
		conn_user.execute("update dvbbs_info set usernum=usernum+1,lastuser='"&RegUserName&"'")
	end if
end sub
%>