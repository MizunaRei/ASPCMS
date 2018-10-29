<%

dim AdminName,AdminPurview,PurviewPassed
dim AdminPurview_Article,AdminPurview_Soft,AdminPurview_Photo,AdminPurview_Guest,AdminPurview_Others
'两课网站代码
dim AdminPurview_Special,AdminPurview_SpecialID
'结束网课网站代码
dim rsGetAdmin,sqlGetAdmin
dim ComeUrl,cUrl
ComeUrl=lcase(trim(request.ServerVariables("HTTP_REFERER")))
'if ComeUrl="" then
'	response.write "<br><p align=center><font color='red'>对不起，为了系统安全，不允许直接输入地址访问本系统的后台管理页面。</font></p>"
'	response.end
'else
'	cUrl=trim("http://" & Request.ServerVariables("SERVER_NAME"))
'	if mid(ComeUrl,len(cUrl)+1,1)=":" then
'		cUrl=cUrl & ":" & Request.ServerVariables("SERVER_PORT")
'	end if
'	cUrl=lcase(cUrl & request.ServerVariables("SCRIPT_NAME"))
	'为了admin_index.asp的美观，放弃检查IP地址
	'if lcase(left(ComeUrl,instrrev(ComeUrl,"/")))<>lcase(left(cUrl,instrrev(cUrl,"/"))) then
'		response.write "<br><p align=center><font color='red'>对不起，为了系统安全，不允许从外部链接地址访问本系统的后台管理页面。</font></p>"
'		response.end
'	end if
'end if

AdminName=replace(session("AdminName"),"'","")
if AdminName="" then
	call CloseConn()
	response.redirect "Admin_login.asp"
end if
sqlGetAdmin="select * from Admin where UserName='" & AdminName & "'"
set rsGetAdmin=server.CreateObject("adodb.recordset")
rsGetAdmin.open sqlGetAdmin,conn,1,1
if rsGetAdmin.bof and rsGetAdmin.eof then
	rsGetAdmin.close
	set rsGetAdmin=nothing
	call CloseConn()
	response.redirect "Admin_login.asp"
end if
AdminPurview=rsGetAdmin("Purview")
AdminPurview_Article=rsGetAdmin("AdminPurview_Article")
AdminPurview_Soft=rsGetAdmin("AdminPurview_Soft")
AdminPurview_Photo=rsGetAdmin("AdminPurview_Photo")
AdminPurview_Guest=rsGetAdmin("AdminPurview_Guest")
AdminPurview_Others=rsGetAdmin("AdminPurview_Others")
'开始两课网站代码
AdminPurview_Special=rsGetAdmin("AdminPurview_Special")
AdminPurview_SpecialID=rsGetAdmin("AdminPurview_SpecialID")
'教师管理员的名字
session("AdminTrueName")=rsGetAdmin("TrueName")

'开始学生管理员的老师
			session("AdminPurview_SpecialID")=AdminPurview_SpecialID
			session("AdminTeacherName")=rsGetAdmin("TeacherName")
			'结束学生管理员的老师
'结束网课网站代码
rsGetAdmin.close
set rsGetAdmin=nothing
PurviewPassed=True
if PurviewLevel>0 then
	if ((AdminPurview - 1)>PurviewLevel) then
		PurviewPassed=False
	else
		if AdminPurview=3 then
			select case CheckChannelID
				case 0        '其他管理操作
					PurviewPassed=CheckPurview(AdminPurview_Others,PurviewLevel_Others)
				case 2        '文章频道
					if AdminPurview_Article>PurviewLevel_Article then
						PurviewPassed=False
					end if
				case 3       '下载频道
					if AdminPurview_Soft>PurviewLevel_Soft then
						PurviewPassed=False
					end if
				case 4       '图片频道
					if AdminPurview_Photo>PurviewLevel_Photo then
						PurviewPassed=False
					end if
				case 5       '留言板
					if AdminType=True then
						PurviewPassed=CheckPurview(AdminPurview_Guest,PurviewLevel_Guest)
					else
						PurviewPassed=True
					end if
			end select
		end if
	end if
end if
if PurviewPassed=False then
	response.write "<br><p align=center><font color='red'>对不起，你没有此项操作的权限。</font></p>"
	response.end
end if

function CheckPurview(AllPurviews,strPurview)
	if isNull(AllPurviews) or AllPurviews="" or strPurview="" then
		CheckPurview=False
		exit function
	end if
	CheckPurview=False
	if instr(AllPurviews,",")>0 then
		dim arrPurviews,i
		arrPurviews=split(AllPurviews,",")
		for i=0 to ubound(arrPurviews)
			if trim(arrPurviews(i))=strPurview then
				CheckPurview=True
				exit for
			end if
		next
	else
		if AllPurviews=strPurview then
			CheckPurview=True
		end if
	end if
end function

function CheckClassMaster(AllMaster,MasterName)
	if isNull(AllMaster) or AllMaster="" or MasterName="" then
		CheckClassMaster=False
		exit function
	end if
	CheckClassMaster=False
	if instr(AllMaster,"|")>0 then
		dim arrMaster,i
		arrMaster=split(AllMaster,"|")
		for i=0 to ubound(arrMaster)
			if trim(arrMaster(i))=MasterName then
				CheckClassMaster=True
				exit for
			end if
		next
	else
		if AllMaster=MasterName then
			CheckClassMaster=True
		end if
	end if
end function
%>
<!--#include file="Admin_PopMenu.asp"-->
