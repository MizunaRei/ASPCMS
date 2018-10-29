<!--#include file="Inc/conn.asp"-->
<!--#include file="inc/Conn_User.asp"-->
<!--#include file="inc/config.asp"-->
<!--#include file="inc/ubbcode.asp"-->
<!--#include file="Inc/function.asp"-->
<%
dim CommentPurviewGrade,CommentUserGrade
UserLogined=CheckUserLogined()
if UserLevel="" then
	UserLevel=9999
else
	UserLevel=Cint(UserLevel)
end if

select case UserLevel
	case 9999
		CommentUserGrade="游客"
	case 999
		CommentUserGrade="注册用户"
	case 99
		CommentUserGrade="收费用户"
	case 9
		CommentUserGrade="VIP用户"
	case 5
		CommentUserGrade="管理员"
end select
select case CommentPurview
	case 9999
		CommentPurviewGrade="游客"
	case 999
		CommentPurviewGrade="注册用户"
	case 99
		CommentPurviewGrade="收费用户"
	case 9
		CommentPurviewGrade="VIP用户"
	case 5
		CommentPurviewGrade="管理员"
end select

if CommentPurview<UserLevel then 
	FoundErr=True
	ErrMsg=ErrMsg & "<br><br><li>对不起，只有本站的<font color=red>"
	ErrMsg=ErrMsg & CommentPurviewGrade
	ErrMsg=ErrMsg & "</font>才能发表评论！</li><br><br>"
	ErrMsg=ErrMsg & "<li>如果你还没注册，请赶紧<a href='User_Reg.asp'><font color=red>点此注册</font></a>吧！</li><br><br>"
	ErrMsg=ErrMsg & "<li>如果你已经注册但还没登录，请赶紧<a href='User_Login.asp'><font color=red>点此登录</font></a>吧！</li><br><br>"
end if

dim ArticleID,Action,ErrMsg,FoundErr
dim Commented,CommentedID,arrCommentedID,i
Action=trim(request("Action"))
ArticleID=trim(request("ArticleID"))
Commented=False
CommentedID=session("CommentedID")

if ArticleId="" then
	founderr=true
	errmsg=errmsg+"<li>请指定要评论的文章ID</li>"
else
	ArticleID=Clng(ArticleID)
end if
if CommentedID<>"" then
	if instr(CommentedID,"|")>0 then
		arrCommentedID=split(CommentedID,"|")
		for i=0 to ubound(arrCommentedID)
			if Clng(arrCommentedID(i))=ArticleID then
				Commented=True
				exit for
			end if
		next
	else
		if Clng(CommentedID)=ArticleID then
			Commented=True
		end if
	end if
end if
if Commented=True then
	FoundErr=True
	ErrMsg=ErrMsg & "<li>你已经对该篇文章发表过评论了！请勿连续对同一篇文章发表评论。</li>"
end if

if FoundErr<>True then
	if Action="Save" then
		call SaveComment()
	else
		call main()
	end if
end if
if FoundErr=True then
	call WriteErrMsg()
end if
call CloseConn()

sub main()
%>
<html>
<head>
<title>发表评论</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="STYLE.CSS" rel="stylesheet" type="text/css">
<script language="JavaScript" type="text/JavaScript">
function Check()
{
  if (document.form1.Name.value=="")
  {
    alert("请输入姓名！");
	document.form1.Name.focus();
	return false;
  }
  if (document.form1.Content.value=="")
  {
    alert("请输入评论内容！");
	document.form1.Content.focus();
	return false;
  }
  return true;  
}
</script>
</head>

<body leftmargin="2" topmargin="0" marginwidth="0" marginheight="0">
<form name="form1" method="post" action="Article_Comment.asp" onSubmit="return Check();">
  <table width="500" border="0" align="center" cellpadding="2" cellspacing="1" class="border">
    <tr class="title"> 
      <td height="22" colspan="4" align="center"><strong>发 表 评 论</strong> <font color=blue>（<%=CommentUserGrade%>）</font></td>
    </tr>
    <% if UserLogined=false then%>
    <tr class="tdbg"> 
      <td align="right" width="110">姓 &nbsp;名：</td>
      <td width="180"> 
        <input type="text" name="Name" maxlength="16" size="20">
        <font color=red>*</font> </td>
      <td width="60" align="right">Oicq：</td>
      <td width="170"> 
        <input type="text" name="Oicq" maxlength="15" size="20" >
      </td>
    </tr>
    <tr class="tdbg"> 
      <td align="right" width="110">性&nbsp;&nbsp;别：</td>
      <td width="180"> 
        <input type="radio" name="Sex" value="1" checked style="BORDER:0px;">
        男&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
        <input type="radio" name="Sex" value="0" style="BORDER:0px;">
        女 </td>
      <td width="60" align="right">Msn：</td>
      <td width="170"> 
        <input type="text" name="Msn" maxlength="40" size="20">
      </td>
    </tr>
    <tr class="tdbg"> 
      <td align="right" width="110">E-mail：</td>
      <td width="180"> 
        <input type="text" name="Email" maxlength="40" size="20">
      </td>
      <td width="60" align="right">Icq：</td>
      <td width="170"> 
        <input type="text" name="Icq" maxlength="15" size="20">
      </td>
    </tr>
    <tr class="tdbg"> 
      <td align="right" width="110">主&nbsp;&nbsp;页：</td>
      <td colspan="3"> 
        <input name="Homepage" type="text" id="Title" size="60" maxlength="60" value="http://">
      </td>
    </tr>
    <%else%>
    <input type="hidden" name="Name"  value="<%=UserName%>">
    <% end if %>
    <tr class="tdbg"> 
      <td align="right" width="110">评 分：</td>
      <td colspan="3"> 
        <input type="radio" name="Score" value="1">
        1分&nbsp;&nbsp;&nbsp;&nbsp; 
        <input type="radio" name="Score" value="2">
        2分&nbsp;&nbsp;&nbsp;&nbsp; 
        <input type="radio" name="Score" value="3" checked>
        3分&nbsp;&nbsp;&nbsp;&nbsp; 
        <input type="radio" name="Score" value="4">
        4分&nbsp;&nbsp;&nbsp;&nbsp; 
        <input type="radio" name="Score" value="5">
        5分 </td>
    </tr>
    <tr class="tdbg"> 
      <td align="right">评论内容：</td>
      <td colspan="3"> 
        <textarea name="Content" cols="50" rows="10" id="Content"></textarea>
      </td>
    </tr>
    <tr bgcolor="#DAE8CA" class="tdbg"> 
      <td colspan="4" align="center"> 
        <input name="Action" type="hidden" id="Action" value="Save">
        <input name="ArticleID" type="hidden" id="ArticleID" value="<%=ArticleID%>">
        <input type="submit" name="Submit" value=" 发 表 ">
      </td>
    </tr>
    <tr bgcolor="#DAE8CA" class="tdbg">
      <td colspan="4">
          <br>
		  
        <li> 请遵守<a href=bbs.htm target=_blank><font color=red>《互联网电子公告服务管理规定》</font></a>及中华人民共和国其他各项有关法律法规。<br>
          <li>严禁发表危害国家安全、损害国家利益、破坏民族团结、破坏国家宗教政策、破坏社会稳定、侮辱、诽谤、教唆、淫秽等内容的评论 。<br>
          <li>用户需对自己在使用本站服务过程中的行为承担法律责任（直接或间接导致的）。<br>
          <li>本站管理员有权保留或删除评论内容。<br>
          <li>评论内容只代表网友个人观点，与本网站立场无关。
      </td>
    </tr>
  </table>
  </form>

</body>
</html>
<%
end sub

sub SaveComment()
	dim rsComment,ClassID,tClass
	dim CommentUserType,CommentUserName,CommentUserSex,CommentUserEmail,CommentUserOicq
	dim CommentUserIcq,CommentUserMsn,CommentUserHomepage,CommentUserScore,CommentUserContent
	if UserLogined=false then
		CommentUserType=0
		CommentUserName=trim(request("Name"))
		if CommentUserName="" then
			founderr=true
			errmsg=errmsg & "<br><li>请输入姓名</li>"
		end if
		CommentUserSex=trim(request("Sex"))
		CommentUserOicq=trim(request("Oicq"))
		CommentUserIcq=trim(request("Icq"))
		CommentUserMsn=trim(request("Msn"))
		CommentUserEmail=trim(request("Email"))
		CommentUserHomepage=trim(request("Homepage"))
		if CommentUserHomepage="http://" or isnull(CommentUserHomepage) then CommentUserHomepage=""
	else
		CommentUserType=1
		CommentUserName=UserName
	end if
	CommentUserScore=Clng(request.Form("Score"))
	CommentUserContent=trim(request.Form("Content"))
	if CommentUserContent="" then
		founderr=true
		errmsg=errmsg & "<br><li>请输入评论内容</li>"
	end if
	CommentUserContent=DvHtmlEncode(CommentUserContent)

	set tClass=conn.execute("select ClassID from Article where ArticleID=" & ArticleID)
	if tClass.bof and tClass.eof then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>找不到要评论的文章，可能已经被删除！</li>"
	else
		ClassID=tClass(0)
	end if
	set tClass=nothing
	
	if founderr=true then
		exit sub
	end if

	set rsComment=server.createobject("adodb.recordset")
	sql="select * from ArticleComment"
	rsComment.open sql,conn,1,3
	rsComment.addnew
	rsComment("ClassID")=ClassID
	rsComment("ArticleID")=ArticleID
	rsComment("UserType")=CommentUserType
	rsComment("UserName")=CommentUserName
	rsComment("Sex")=CommentUserSex
	rsComment("Oicq")=CommentUserOicq
	rsComment("Icq")=CommentUserIcq
	rsComment("Msn")=CommentUserMsn
	rsComment("Email")=CommentUserEmail
	rsComment("Homepage")=CommentUserHomepage
	rsComment("IP")=Request.ServerVariables("REMOTE_ADDR")
	rsComment("Score")=CommentUserScore
	rsComment("Content")=CommentUserContent
	rsComment("WriteTime")=now()
	rsComment.update
	rsComment.close
	set rsComment=nothing

	if CommentedID="" then
		session("CommentedID")=ArticleID
	else
		session("CommentedID")=CommentedID & "|" & ArticleID
	end if
	call WriteSuccessMsg("发表评论成功！")
end sub
%>