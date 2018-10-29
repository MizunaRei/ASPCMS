<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%
option explicit
response.buffer=true	
Const PurviewLevel=2
Const CheckChannelID=0
Const PurviewLevel_Others="User"
%>
<!--#include file="inc/conn.asp"-->
<!--#include file="inc/config.asp"-->
<!--#include file="Admin_ChkPurview.asp"-->
<!--#include file="inc/function.asp"-->
<%
dim rs, sql
dim Action,FoundErr,ErrMsg
Action=trim(request("Action"))
%>
<html>
<head>
<title>用户等级管理</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="Admin_Style.css" rel="stylesheet" type="text/css">
</head>
<body leftmargin="2" topmargin="0" marginwidth="0" marginheight="0">
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" Class="border">
  <tr class="topbg"> 
    <td height="22" colspan=2 align=center><strong>用 户 等 级 管 理</strong></td>
  </tr>
    <tr class="tdbg"> 
      <td width="70" height="30"><strong>管理导航：</strong></td>
      
    <td height="30"> <a href="Admin_UserGrade.asp">用户等级管理首页</a>&nbsp;|&nbsp;<a href="Admin_UserGrade.asp?Action=Add">添加新用户等级</a></td>
    </tr>
</table>
<br>
<%
if Action="Add" then
	call Add()
elseif Action="SaveAdd" then
	call SaveAdd()
elseif Action="SaveModify" then
	call SaveModify()
elseif Action="Del" then
	call DelGrade()
else
	call main()
end if
if FoundErr=True then
	call WriteErrMsg()
end if
call CloseConn()

sub main()
	sql="select * from UserGrade order by Grade"
	Set rs=Server.CreateObject("Adodb.RecordSet")
	rs.open sql,conn,1,1
%>
<form name="form1" method="post" action="">
  <table width="100%" border="0" cellspacing="1" cellpadding="2" class="border">
    <tr align="center" class="title"> 
      <td height="22"><strong>用户等级</strong></td>
      <td height="22"><strong>等级名称</strong></td>
      <td height="22"><strong>图 片</strong></td>
      <td height="22"><strong>最少文章数</strong></td>
      <td height="22"><strong>每天限制发表文章数</strong></td>
      <td><strong>操作</strong></td>
    </tr>
    <%do while not rs.eof%>
    <tr align="center" class="tdbg" onmouseout="this.style.backgroundColor=''" onmouseover="this.style.backgroundColor='#BFDFFF'"> 
      <td><input name="GradeID" type="hidden" id="GradeID" value="<%=rs("GradeID")%>"> 
        <input name="Grade" type="text" id="Grade" value="<%=rs("Grade")%>" size="8" maxlength="5"></td>
      <td><input name="GradeName" type="text" id="GradeName" value="<%=rs("GradeName")%>" size="20" maxlength="50"></td>
      <td><input name="GradePic" type="text" id="GradePic" value="<%=rs("GradePic")%>" size="20" maxlength="50"></td>
      <td><input name="MinArticle" type="text" id="MinArticle" value="<%=rs("MinArticle")%>" size="10" maxlength="10"></td>
      <td><input name="LimitEveryDay" type="text" id="LimitEveryDay" value="<%=rs("LimitEveryDay")%>" size="10" maxlength="8"></td>
      <td><a href="Admin_UserGrade.asp?Action=Del&GradeID=<%=rs("GradeID")%>">删除</a></td>
    </tr>
    <%
		rs.movenext
	loop
	rs.close
	set rs=nothing
%>
    <tr align="center" class="tdbg"> 
      <td colspan="6"><input name="Action" type="hidden" id="Action" value="SaveModify"> 
        <input type="submit" name="Submit" value=" 保存修改结果 "></td>
    </tr>
  </table>
</form>
<%
end sub

sub Add()
%>
<form name="form1" method="post" action="Admin_UserGrade.asp">
<table width="100%" border="0" cellspacing="1" cellpadding="2" class="border">
  <tr align="center" class="title"> 
    <td height="22" colspan="2"><strong>新 增 用 户 等 级</strong></td>
  </tr>
  <tr class="tdbg"> 
      <td width="300"><strong>用户等级</strong><br>
        必须是数字</td>
    <td><input name="Grade" type="text" id="Grade" size="8" maxlength="5"></td>
  </tr>
  <tr class="tdbg"> 
    <td width="300"><strong>等级名称</strong></td>
    <td><input name="GradeName" type="text" id="GradeName" size="30" maxlength="50"></td>
  </tr>
  <tr class="tdbg"> 
      <td width="300"><strong>图片</strong><br>
        请首先将相应图片放到images目录中，然后在此直接输入文件名即可，不要输入路径</td>
    <td><input name="GradePic" type="text" id="GradePic" size="30" maxlength="50"></td>
  </tr>
  <tr class="tdbg"> 
      <td width="300"><strong>最少文章数</strong><br>
        必须是数字</td>
    <td><input name="MinArticle" type="text" id="MinArticle" size="8" maxlength="8"></td>
  </tr>
  <tr class="tdbg"> 
      <td width="300"><strong>每天限制发表文章数</strong><br>
        必须是数字</td>
    <td><input name="LimitEveryDay" type="text" id="LimitEveryDay" size="8" maxlength="8"></td>
  </tr>
  <tr align="center" class="tdbg"> 
    <td colspan="2"><input name="Action" type="hidden" id="Action" value="SaveAdd">
        <input type="submit" name="Submit2" value=" 添 加 "></td>
  </tr>
</table>
</form>
<%
end sub
%>
</body>
</html>
<%
sub SaveAdd()
	dim Grade,GradeName,GradePic,MinArticle,LimitEveryDay
	Grade=trim(request("Grade"))
	GradeName=trim(request("GradeName"))
	GradePic=trim(request("GradePic"))
	MinArticle=trim(request("MinArticle"))
	LimitEveryDay=trim(request("LimitEveryDay"))
	if Grade="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>请输入用户等级</li>"
	elseif not isnumeric(Grade) then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>用户等级必须是数字</li>"
	end if
	if GradeName="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>请输入等级名称</li>"
	end if
	if GradePic="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>请输入等级图片</li>"
	end if
	if MinArticle="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>请输入最少文章数</li>"
	elseif not isnumeric(MinArticle) then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>最少文章数必须是数字</li>"
	end if
	if LimitEveryDay="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>请输入每天限制发表文章数</li>"
	elseif not isnumeric(LimitEveryDay) then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>每天限制发表文章数必须是数字</li>"
	end if
	if FoundErr=True then
		exit sub
	end if
	sql="select * from UserGrade where Grade=" & Clng(Grade) & " or GradeName='" & GradeName & "'"
	set rs = server.CreateObject ("Adodb.recordset")
	rs.open sql,conn,1,3
	if rs.eof and rs.bof then
		rs.addnew
		rs("Grade")=Clng(Grade)
		rs("GradeName")=GradeName
		rs("GradePic")=GradePic
		rs("MinArticle")=MinArticle
		rs("LimitEveryDay")=LimitEveryDay
		rs.update
		rs.close
		set rs=nothing
		call CloseConn()
		response.redirect "Admin_UserGrade.asp"
	else
		FoundErr=True
		if rs("Grade")=Grade then
			ErrMsg=ErrMsg & "<br><li>已经存在用户等级：" & Grade & "<li>"
		end if
		if rs("GradeName")=GradeName then
			ErrMsg=ErrMsg & "<br><li>已经存在等级名称“" & GradeName & "”</li>"
		end if
	end if
	rs.close
	set rs=nothing
end sub

sub SaveModify()
	dim GradeID,Grade,GradeName,GradePic,MinArticle,LimitEveryDay,i
	GradeID=trim(request("GradeID"))
	Grade=trim(request("Grade"))
	GradeName=trim(request("GradeName"))
	GradePic=trim(request("GradePic"))
	MinArticle=trim(request("MinArticle"))
	LimitEveryDay=trim(request("LimitEveryDay"))
	if GradeID="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>请先添加用户等级</li>"
		exit sub
	end if
	if instr(GradeID,",")>0 then
		for i=1 to request.form("GradeID").count
			GradeID=replace(trim(request.form("GradeID")(i)),"'","")
			Grade=replace(trim(request.form("Grade")(i)),"'","")
			GradeName=replace(trim(request.form("GradeName")(i)),"'","")
			GradePic=replace(trim(request.form("GradePic")(i)),"'","")
			MinArticle=replace(trim(request.form("MinArticle")(i)),"'","")
			LimitEveryDay=replace(trim(request.form("LimitEveryDay")(i)),"'","")
			if isnumeric(GradeID) and isnumeric(Grade) and GradeName<>"" and GradePic<>"" and isnumeric(MinArticle) and isnumeric(LimitEveryDay) then
				conn.execute("update UserGrade set Grade=" & Clng(trim(Grade)) & ",GradeName='" & trim(GradeName) & "',GradePic='" & trim(GradePic) & "',MinArticle=" & Clng(Trim(MinArticle)) & ",LimitEveryDay=" & Clng(Trim(LimitEveryDay)) & " where GradeID=" & Clng(trim(GradeID)))	
			end if
		next
	else
		if isnumeric(GradeID) and isnumeric(Grade) and GradeName<>"" and GradePic<>"" and isnumeric(MinArticle) and isnumeric(LimitEveryDay) then
			conn.execute("update UserGrade set Grade=" & Clng(trim(Grade)) & ",GradeName='" & trim(GradeName) & "',GradePic='" & trim(GradePic) & "',MinArticle=" & Clng(Trim(MinArticle)) & ",LimitEveryDay=" & Clng(Trim(LimitEveryDay)) & " where GradeID=" & Clng(trim(GradeID)))	
		end if
	end if
	call WriteSuccessMsg("保存用户等级修改结果成功！")
end sub

sub DelGrade()
	dim GradeID,trs,tGrade
	GradeID=trim(request("GradeID"))
	if GradeID="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>请指定要删除的用户等级ID</li>"
		exit sub
	else
		GradeID=Clng(GradeID)
	end if
	set trs=conn.execute("select Grade from UserGrade where GradeID=" & GradeID)
	tGrade=trs(0)
	set trs=conn.execute("select GradeID from UserGrade where Grade<" & tGrade)
	if not (trs.bof and trs.eof) then
		conn.execute("update [User] set UserGrade=" & trs(0) & " where UserGrade=" & GradeID)
	else
		conn.execute("update [User] set UserGrade=1 where UserGrade=" & GradeID)
	end if
	conn.execute("delete from UserGrade where GradeID=" & GradeID)
	call WriteSuccessMsg("删除等级成功！同时已将属于该等级的用户降为下一级。")
end sub
%>
