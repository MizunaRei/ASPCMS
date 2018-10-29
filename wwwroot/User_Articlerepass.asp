<!--#include file="inc/conn.asp"-->
<!--#include file="inc/Conn_User.asp"-->
<!--#include file="Inc/function.asp"-->
<!--#include file="inc/config.asp"-->
<%
if CheckUserLogined()=False then
	response.Redirect "User_Login.asp"
end if

dim ArticleID,Action,sqlrepass,rsrepass,FoundErr,ErrMsg,PurviewChecked,ObjInstalled
ArticleID=trim(request("ArticleID"))
FoundErr=False
ObjInstalled=IsObjInstalled("Scripting.FileSystemObject")

if ArticleId="" then
	FoundErr=True
	ErrMsg=ErrMsg & "<br><li>参数不足！</li>"
end if
if FoundErr=False then
	if instr(ArticleID,",")>0 then
		dim idarr,i
		idArr=split(ArticleID)
		for i = 0 to ubound(idArr)
			call Articlerepass(clng(idarr(i)))
		next
	else
		call Articlerepass(clng(ArticleID))
	end if
end if
call CloseConn()
if FoundErr=False then
	response.Redirect "User_Articlere.asp"
else
	call WriteErrMsg()
end if

sub Articlerepass(ID)
	PurviewChecked=False
	sqlrepass="select * from article where ArticleID=" & CLng(ID)
	Set rsrepass= Server.CreateObject("ADODB.Recordset")
	rsrepass.open sqlrepass,conn,1,3
	if rsrepass.bof and rsrepass.eof then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>找不到文章：" & rsrepass("Title") & " </li>"
	else
		if rsrepass("Editor")=Trim(Request.Cookies("asp163")("UserName")) then
			if rsrepass("Passed")=True or rsrepass("noPass")=False then
				FoundErr=True
				ErrMsg=ErrMsg & "<br><li>提交文章“" & rsrepass("Title") & "”失败。原因：此文章已经被审核通过，你不能再次提交！</li>"
			end if
		else
			FoundErr=True
			ErrMsg=ErrMsg & "<br><li>提交文章“" & rsrepass("Title") & "”失败。原因：此文章是其他网友发表的，你不能提交其他人的文章！</li>"
		end if
	end if
	if FoundErr=False then
		rsrepass("nopass")=False
		rsrepass.update
	end if
	rsrepass.close
	set rsrepass=nothing
end sub
%>
