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
<!--#include file="inc/Conn_User.asp"-->
<!--#include file="inc/function.asp"-->
<!--#include file="inc/admin_code_article.asp"-->
<%
dim ArticleID,Action,sqlDel,rsDel,FoundErr,ErrMsg,PurviewChecked,ObjInstalled
dim ClassID,tClass,ClassName,RootID,ParentID,Depth,ParentPath,Child,ChildID,tID,tChild,ClassMaster
ArticleID=trim(request("ArticleID"))
Action=Trim(Request("Action"))
ObjInstalled=IsObjInstalled("Scripting.FileSystemObject")
FoundErr=False
if Action="" then
	FoundErr=True
	ErrMsg=ErrMsg & "<br><li>参数不足！</li>"
end if
if FoundErr=False then
	if Action="Del" then
		call DelArticle()
	elseif Action="ConfirmDel" then
		call ConfirmDel()
	elseif Action="ClearRecyclebin" then
		call ClearRecyclebin()
	elseif Action="Restore" then
		call Restore()
	elseif Action="RestoreAll" then
		call RestoreAll()
	elseif Action="DelFromSpecial" then
		call DelFromSpecial()
	end if
end if
if FoundErr=False then
	call CloseConn()
	response.Redirect ComeUrl
else
	call WriteErrMsg()
	call CloseConn()
end if

sub DelArticle()
	if ArticleID="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>请先选定文章！</li>"
		exit sub
	end if

	if instr(ArticleID,",")>0 then
		ArticleID=replace(ArticleID," ","")
		sqlDel="select * from Article where ArticleID in (" & ArticleID & ")"
	else
		ArticleID=Clng(ArticleID)
		sqlDel="select * from article where ArticleID=" & ArticleID
	end if
	Set rsDel= Server.CreateObject("ADODB.Recordset")
	rsDel.open sqlDel,conn,1,3
	do while not rsDel.eof
		PurviewChecked=False
		ClassID=rsDel("ClassID")
		if AdminPurview=1 or AdminPurview_Article<=2 or (rsDel("Editor")=AdminName and rsDel("Passed")=False) then
			PurviewChecked=True
		else
				set tClass=conn.execute("select ClassName,RootID,ParentID,Depth,ParentPath,Child,ClassMaster From ArticleClass where ClassID=" & ClassID)
				if tClass.bof and tClass.eof then
					founderr=True
					ErrMsg=ErrMsg & "<br><li>找不到指定的栏目</li>"
				else
					ClassName=tClass(0)
					RootID=tClass(1)
					ParentID=tClass(2)
					Depth=tClass(3)
					ParentPath=tClass(4)
					Child=tClass(5)
					ClassMaster=tClass(6)
					PurviewChecked=CheckClassMaster(ClassMaster,AdminName)
					if PurviewChecked=False and ParentID>0 then
						set tClass=conn.execute("select ClassMaster from ArticleClass where ClassID in (" & ParentPath & ")")
						do while not tClass.eof
							PurviewChecked=CheckClassMaster(tClass(0),AdminName)
							if PurviewChecked=True then exit do
							tClass.movenext
						loop
					end if
				end if
		end if
		if PurviewChecked=False then
			FoundErr=True
			ErrMsg=ErrMsg & "<br><li>删除" & rsDel("ArticleID") & "失败！原因：没有操作权限！</li>"
		else
			rsDel("Deleted")=True
			rsDel.update
			if rsDel("Passed")=True then
				Conn_User.execute("update " & db_User_Table & " set " & db_User_ArticleCount & "=" & db_User_ArticleCount & "-1," & db_User_ArticleChecked & "=" & db_User_ArticleChecked & "-1 where " & db_User_Name & "='" & rsDel("Editor") & "'")
			else
				Conn_User.execute("update " & db_User_Table & " set " & db_User_ArticleCount & "=" & db_User_ArticleCount & "-1 where " & db_User_Name & "='" & rsDel("Editor") & "'")
			end if
		end if
		rsDel.movenext
	loop
	rsDel.close
	set rsDel=nothing
end sub

sub ConfirmDel()
	if AdminPurview=2 and AdminPurview_Article>1 then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>对不起，你的权限不够！</li>"
		exit sub
	end if
	if ArticleID="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>请先选定文章！</li>"
		exit sub
	end if
	ArticleID=replace(ArticleID," ","")
	sqlDel="select UploadFiles from article where ArticleID in (" & ArticleID & ")"
	Set rsDel= Server.CreateObject("ADODB.Recordset")
	rsDel.open sqlDel,conn,1,3
	do while not rsDel.eof
		call DelFiles(rsDel(0) & "")
		rsDel.movenext
	loop
	rsDel.close
	set rsDel=nothing
	conn.execute("delete from Article where ArticleID in (" & ArticleID & ")")
	conn.execute("delete from ArticleComment where ArticleID in (" & ArticleID & ")")
end sub

sub ClearRecyclebin()
	if AdminPurview=2 and AdminPurview_Article>1 then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>对不起，你的权限不够！</li>"
		exit sub
	end if
	ArticleID=""
	sqlDel="select ArticleID,UploadFiles from article where Deleted=True"
	Set rsDel= Server.CreateObject("ADODB.Recordset")
	rsDel.open sqlDel,conn,1,3
	do while not rsDel.eof
		if ArticleID="" then
			ArticleID=rsDel(0)
		else
			ArticleID=ArticleID & "," & rsDel(0)
		end if
		call DelFiles(rsDel(1) & "")
		rsDel.movenext
	loop
	rsDel.close
	set rsDel=nothing
	if ArticleID<>"" then
		conn.execute("delete from Article where Deleted=True")
		conn.execute("delete from ArticleComment where ArticleID in (" & ArticleID & ")")
	end if
end sub

sub Restore()
	if AdminPurview=2 and AdminPurview_Article>1 then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>对不起，你的权限不够！</li>"
		exit sub
	end if
	if ArticleID="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>请先选定文章！</li>"
		exit sub
	end if
	ArticleID=replace(ArticleID," ","")
	conn.execute("update Article set Deleted=False where ArticleID in (" & ArticleID & ")")
	if instr(ArticleID,",")>0 then
		ArticleID=replace(ArticleID," ","")
		sqlDel="select * from Article where ArticleID in (" & ArticleID & ")"
	else
		ArticleID=Clng(ArticleID)
		sqlDel="select * from article where ArticleID=" & ArticleID
	end if
	Set rsDel= Server.CreateObject("ADODB.Recordset")
	rsDel.open sqlDel,conn,1,3
	do while not rsDel.eof
		rsDel("Deleted")=False
		rsDel.update
		if rsDel("Passed")=True then
			Conn_User.execute("update " & db_User_Table & " set " & db_User_ArticleCount & "=" & db_User_ArticleCount & "+1," & db_User_ArticleChecked & "=" & db_User_ArticleChecked & "+1 where " & db_User_Name & "='" & rsDel("Editor") & "'")
		else
			Conn_User.execute("update " & db_User_Table & " set " & db_User_ArticleCount & "=" & db_User_ArticleCount & "+1 where " & db_User_Name & "='" & rsDel("Editor") & "'")
		end if
		rsDel.movenext
	loop
	rsDel.close
	set rsDel=nothing
	
end sub

sub RestoreAll()
	if AdminPurview=2 and AdminPurview_Article>1 then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>对不起，你的权限不够！</li>"
		exit sub
	end if
	sqlDel="select * from Article where Deleted=True"
	Set rsDel= Server.CreateObject("ADODB.Recordset")
	rsDel.open sqlDel,conn,1,3
	do while not rsDel.eof
		rsDel("Deleted")=False
		rsDel.update
		if rsDel("Passed")=True then
			Conn_User.execute("update " & db_User_Table & " set " & db_User_ArticleCount & "=" & db_User_ArticleCount & "+1," & db_User_ArticleChecked & "=" & db_User_ArticleChecked & "+1 where " & db_User_Name & "='" & rsDel("Editor") & "'")
		else
			Conn_User.execute("update " & db_User_Table & " set " & db_User_ArticleCount & "=" & db_User_ArticleCount & "+1 where " & db_User_Name & "='" & rsDel("Editor") & "'")
		end if
		rsDel.movenext
	loop
	rsDel.close
	set rsDel=nothing
end sub

sub DelFromSpecial()
	if AdminPurview=2 and AdminPurview_Article>1 then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>对不起，你的权限不够！</li>"
		exit sub
	end if
	ArticleID=replace(ArticleID," ","")
	conn.execute("update Article set SpecialID=0 where ArticleID in (" & ArticleID & ")")
end sub

%>
