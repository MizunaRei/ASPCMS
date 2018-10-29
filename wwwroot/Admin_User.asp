<%@language=vbscript codepage=936 %>
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
<!--#include file="inc/Conn_User.asp"-->
<!--#include file="inc/function.asp"-->
<!--#include file="inc/md5.asp"-->
<!--#include file="Inc/RegBBS.asp"-->
<%
const MaxPerPage=20
dim strFileName
dim totalPut,CurrentPage,TotalPages
dim rs, sql
dim UserID,UserSearch,Keyword,strField
dim Action,FoundErr,ErrMsg
dim tmpDays
keyword=trim(request("keyword"))
if keyword<>"" then 
	keyword=ReplaceBadChar(keyword)
end if
strField=trim(request("Field"))
UserSearch=trim(request("UserSearch"))
Action=trim(request("Action"))
UserID=trim(Request("UserID"))
ComeUrl=Request.ServerVariables("HTTP_REFERER")

if UserSearch="" then
	UserSearch=0
else
	UserSearch=Clng(UserSearch)
end if
strFileName="Admin_User.asp?UserSearch=" & UserSearch
if request("page")<>"" then
    currentPage=cint(request("page"))
else
	currentPage=1
end if

%>
<html>
<head>
<title>注册用户管理</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="Admin_Style.css" rel="stylesheet" type="text/css">
<SCRIPT language=javascript>
function unselectall()
{
    if(document.myform.chkAll.checked){
	document.myform.chkAll.checked = document.myform.chkAll.checked&0;
    } 	
}

function CheckAll(form)
{
  for (var i=0;i<form.elements.length;i++)
    {
    var e = form.elements[i];
    if (e.Name != "chkAll")
       e.checked = form.chkAll.checked;
    }
}
</SCRIPT>
</head>
<body leftmargin="2" topmargin="0" marginwidth="0" marginheight="0">
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" Class="border">
  <tr class="topbg"> 
    <td height="22" colspan=2 align=center><strong>注 册 用 户 管 理</strong></td>
  </tr>
  <form name="form1" action="Admin_User.asp" method="get">
    <tr class="tdbg"> 
      <td width="100" height="30"><strong>快速查找用户：</strong></td>
      <td width="687" height="30"><select size=1 name="UserSearch" onChange="javascript:submit()">
          <option value="0" <%if UserSearch=0 then response.write " selected"%>>列出所有用户</option>
          <option value="1" <%if UserSearch=1 then response.write " selected"%>>文章最多TOP100</option>
          <option value="2" <%if UserSearch=2 then response.write " selected"%>>文章最少的100个用户</option>
          <option value="3" <%if UserSearch=3 then response.write " selected"%>>最近24小时内登录的用户</option>
          <option value="4" <%if UserSearch=4 then response.write " selected"%>>最近24小时内注册的用户</option>
          <option value="5" <%if UserSearch=5 then response.write " selected"%>>等待邮件验证的用户</option>
          <option value="6" <%if UserSearch=6 then response.write " selected"%>>等待管理员认证的用户</option>
          <option value="7" <%if UserSearch=7 then response.write " selected"%>>所有被锁住的用户</option>
          <option value="8" <%if UserSearch=8 then response.write " selected"%>>所有收费用户</option>
          <option value="9" <%if UserSearch=9 then response.write " selected"%>>所有VIP用户</option>
        </select>
        &nbsp;&nbsp;&nbsp;&nbsp;<a href="Admin_User.asp">用户管理首页</a>&nbsp;|&nbsp;<a href="Admin_User.asp?Action=Add">添加新用户</a></td>
    </tr>
  </form>
</table>
<br>
<%
if Action="Add" then
	call AddUser()
elseif Action="SaveAdd" then
	call SaveAdd()
elseif Action="Modify" then
	call Modify()
elseif Action="SaveModify" then
	call SaveModify()
elseif Action="Del" then
	call DelUser()
elseif Action="Lock" then
	call LockUser()
elseif Action="UnLock" then
	call UnLockUser()
elseif Action="Move" then
	call MoveUser()
elseif Action="Update" then
	call UpdateUser()
elseif Action="DoUpdate" then
	call DoUpdate()
elseif Action="AddMoney" then
	call AddMoney()
elseif Action="SaveAddMoney" then
	call SaveAddMoney()
else
	call main()
end if
if FoundErr=True then
	call WriteErrMsg()
end if
call CloseConn_User()  

sub main()
	dim strGuide
	strGuide="<table width='100%'><tr><td align='left'>您现在的位置：<a href='Admin_User.asp'>注册用户管理</a>&nbsp;&gt;&gt;&nbsp;"
	select case UserSearch
		case 0
			sql="select * from " & db_User_Table & " order by " & db_User_ID & " desc"
			strGuide=strGuide & "所有用户"
		case 1
			sql="select top 100 * from " & db_User_Table & " order by " & db_User_ArticleChecked & " desc"
			strGuide=strGuide & "发表文章最多的前100个用户"
		case 2
			sql="select top 100 * from " & db_User_Table & " order by " & db_User_ArticleChecked & ""
			strGuide=strGuide & "发表文章最少的100个用户"
		case 3
			sql="select * from " & db_User_Table & " where datediff('h'," & db_User_LastLoginTime & ",Now())<25 order by " & db_User_LastLoginTime & " desc"
			strGuide=strGuide & "最近24小时内登录的用户"
		case 4
			sql="select * from " & db_User_Table & " where datediff('h'," & db_User_RegDate & ",Now())<25 order by " & db_User_RegDate & " desc"
			strGuide=strGuide & "最近24小时内注册的用户"
		case 5
			sql="select * from " & db_User_Table & " where " & db_User_UserLevel & "=3000 order by " & db_User_ID & " desc"
			strGuide=strGuide & "等待邮件验证的用户"
		case 6
			sql="select * from " & db_User_Table & " where " & db_User_UserLevel & "=2000 order by " & db_User_ID & " desc"
			strGuide=strGuide & "等待管理认证证的用户"
		case 7
			sql="select * from " & db_User_Table & " where " & db_User_LockUser & "=True order by " & db_User_ID & " desc"
			strGuide=strGuide & "所有被锁住的用户"
		case 8
			sql="select * from " & db_User_Table & " where " & db_User_UserLevel & "=99 order by " & db_User_ID & " desc"
			strGuide=strGuide & "所有收费用户"
		case 9
			sql="select * from " & db_User_Table & " where " & db_User_UserLevel & "=9 order by " & db_User_ID & " desc"
			strGuide=strGuide & "所有VIP用户"
		case 10
			if Keyword="" then
				sql="select * from " & db_User_Table & " order by " & db_User_ID & " desc"
				strGuide=strGuide & "所有用户"
			else
				select case strField
				case "UserID"
					if IsNumeric(Keyword)=False then
						FoundErr=True
						ErrMsg=ErrMsg & "<br><li>用户ID必须是整数</li>"
					else
						sql="select * from " & db_User_Table & " where " & db_User_ID & "=" & Clng(Keyword)
						strGuide=strGuide & "用户ID等于<font color=red> " & Clng(Keyword) & " </font>的用户"
					end if
				case "UserName"
					sql="select * from " & db_User_Table & " where " & db_User_Name & " like '%" & Keyword & "%' order by " & db_User_ID & " desc"
					strGuide=strGuide & "用户名中含有“ <font color=red>" & Keyword & "</font> ”的用户"
				end select
			end if
		case else
			FoundErr=True
			ErrMsg=ErrMsg & "<br><li>错误的参数！</li>"
	end select
	strGuide=strGuide & "</td><td align='right'>"
	if FoundErr=True then exit sub
	
	Set rs=Server.CreateObject("Adodb.RecordSet")
	rs.Open sql,Conn_User,1,1
  	if rs.eof and rs.bof then
		strGuide=strGuide & "共找到 <font color=red>0</font> 个用户</td></tr></table>"
		response.write strGuide
	else
    	totalPut=rs.recordcount
		strGuide=strGuide & "共找到 <font color=red>" & totalPut & "</font> 个用户</td></tr></table>"
		response.write strGuide
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
        	showContent
        	showpage strFileName,totalput,MaxPerPage,true,true,"个用户"
   	 	else
   	     	if (currentPage-1)*MaxPerPage<totalPut then
         	   	rs.move  (currentPage-1)*MaxPerPage
         		dim bookmark
           		bookmark=rs.bookmark
            	showContent
            	showpage strFileName,totalput,MaxPerPage,true,true,"个用户"
        	else
	        	currentPage=1
           		showContent
           		showpage strFileName,totalput,MaxPerPage,true,true,"个用户"
	    	end if
		end if
	end if
	rs.Close
	set rs=Nothing
	call ShowSearch()
end sub

sub showContent()
   	dim i
    i=0
%>
<table width='100%' border="0" cellpadding="0" cellspacing="0">
  <tr>
  <form name="myform" method="Post" action="Admin_User.asp" onSubmit="return confirm('确定要执行选定的操作吗？');">
      <td height="90"> 
        <table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="border">
          <tr class="title">
            <td width="30" align="center"><strong>选中</strong></td>
            <td width="30" align="center"><strong>ID</strong></td>
            <td width="80" height="22" align="center"><strong> 用户名</strong></td>
            <td width="84" height="22" align="center"><strong>所属用户组</strong></td>
            <td width="87" align="center"><strong>点数/天数</strong></td>
            <td width="82" height="24" align="center"><strong>最后登录IP</strong></td>
            <td width="97" align="center"><strong>最后登录时间</strong></td>
            <td width="60" height="22" align="center"><strong>登录次数</strong></td>
            <td width="40" height="22" align="center"><strong> 状态</strong></td>
<!-- ============================ 山风多用户日记本插件修改 01 开始 ==================================== -->
            <td width="59" height="20" align="center"><strong>日记数</strong></td>
<!-- ============================ 山风多用户日记本插件修改 01 结束 ==================================== -->
            <td width="60" align="center"><strong>操作</strong></td>
          </tr>
          <%do while not rs.EOF %>
          <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
            <td width="30" height="22" align="center"> 
              <input name='UserID' type='checkbox' onClick="unselectall()" id="UserID" value='<%=cstr(rs(db_User_ID))%>'></td>
            <td width="30" align="center"><%=rs(db_User_ID)%></td>
            <td width="80" align="center"><%
			response.write "<a href='Admin_User.asp?Action=Modify&UserID=" & rs(db_User_ID) & "' title=""======== 用 户 信 息 ========" & vbcrlf & "性别："
			if rs(db_User_Sex)=1 then
				response.write "男"
			else
				response.write "女"
			end if
			response.write vbcrlf & "信箱：" & rs(db_User_Email) & vbcrlf & "ＱＱ："
			if rs(db_User_QQ)<>"" then
				response.write rs(db_User_QQ)
			else
				response.write "未填"
			end if
			response.write vbcrlf & "MSN："
			if rs(db_User_Msn)<>"" then
			response.write rs(db_User_Msn)
			else
				response.write "未填"
			end if
			response.write vbcrlf & "主页："
			if rs(db_User_Homepage)<>"" then
				response.write rs(db_User_Homepage)
			else
				response.write "未填"
			end if
			response.write vbcrlf & "注册日期：" & rs(db_User_RegDate)

			response.write """>" & rs(db_User_Name) & "</a>"
			%> </td>
            <td align="center"> <%
			select case rs(db_User_UserLevel)
				case 3000
					response.write "<font color=green>等待邮件验证的用户</font>"
				case 2000
					response.write "<font color=green>等待管理员认证的用户</font>"
				case 999
					response.write "普通注册用户"
				case 99
					response.write "<font color=blue>收费用户</font>"
				case 9
					response.write "<font color=blue>VIP用户</font>"
				case else
					response.write "<font color=red>异常用户</font>"
			end select
			%> </td>
            <td align="center"> <%
	if rs(db_User_UserLevel)=99 or rs(db_User_UserLevel)=9 then
		if rs(db_User_ChargeType)=1 then
			if rs(db_User_UserPoint)<=0 then
				response.write "<font color=red>" & rs(db_User_UserPoint) & "</font> 点"
			else
				if rs(db_User_UserPoint)<=10 then
					response.write "<font color=blue>" & rs(db_User_UserPoint) & "</font> 点"
				else
					response.write rs(db_User_UserPoint) & " 点"
				end if
			end if
		else
		  if rs(db_User_Valid_Unit)=1 then
			ValidDays=rs(db_User_Valid_Num)
		  elseif rs(db_User_Valid_Unit)=2 then
			ValidDays=rs(db_User_Valid_Num)*30
		  elseif rs(db_User_Valid_Unit)=3 then
			ValidDays=rs(db_User_Valid_Num)*365
		  end if
		  tmpDays=ValidDays-DateDiff("D",rs(db_User_BeginDate),now())
		  if tmpDays<=0 then
			response.write "<font color=red>" & tmpDays & "</font> 天"
		  else
		  	if tmpDays<=10 then
				response.write "<font color=blue>" & tmpDays & "</font> 天"
		    else
				response.write tmpDays & " 天"
			end if
		  end if
		end if
	else
		response.write "&nbsp;"
	end if
		%></td>
            <td align="center"> <%
	if rs(db_User_LastLoginIP)<>"" then
		response.write rs(db_User_LastLoginIP)
	else
		response.write "&nbsp;"
	end if
	%> </td>
            <td align="center"> <%
	if rs(db_User_LastLoginTime)<>"" then
		response.write rs(db_User_LastLoginTime)
	else
		response.write "&nbsp;"
	end if
	%> </td>
            <td width="60" align="center"> <%
	if rs(db_User_LoginTimes)<>"" then
		response.write rs(db_User_LoginTimes)
	else
		response.write "0"
	end if
	%> </td>
            <td width="40" align="center"><%
	  if rs(db_User_LockUser)=true then
	  	response.write "<font color=red>已锁定</font>"
	  else
	  	response.write "正常"
	  end if
	  %></td>
<!-- ============================ 山风多用户日记本插件修改 02 开始 ==================================== -->
            <td align="center"><a href=diary_userdiarylist.asp?diaryOwner=<%=rs("UserName")%> title=点击查看管理该用户的公开日记><%=rs("diaryNum")%></a></td>
            <td align="center"> 
              <%
		response.write "<a href='Admin_User.asp?Action=Modify&UserID=" & rs(db_User_ID) & "'>改</a>&nbsp;"
		if rs(db_User_LockUser)=False then
			response.write "<a href='Admin_User.asp?Action=Lock&UserID=" & rs(db_User_ID) & "'>锁</a>&nbsp;"
		else
            response.write "<a href='Admin_User.asp?Action=UnLock&UserID=" & rs(db_User_ID) & "'>解</a>&nbsp;"
		end if
        response.write "<a href='Admin_User.asp?Action=Del&UserID=" & rs(db_User_ID) & "' onClick='return confirm(""确定要删除此用户吗？"");'>删</a>&nbsp;"
		if rs(db_User_UserLevel)=99 or rs(db_User_UserLevel)=9 then
			response.write "<a href='Admin_User.asp?Action=AddMoney&UserID=" & rs(db_User_ID) & "'>续费</a>"
		else
            response.write "&nbsp;&nbsp;&nbsp;&nbsp;"
		end if
		%>
            </td>
          </tr>
          <%
	i=i+1
	if i>=MaxPerPage then exit do
	rs.movenext
loop
%>
        </table>  
        <strong>操作</strong>
<table width="100%" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td width="200" height="30"><input name="chkAll" type="checkbox" id="chkAll" onclick=CheckAll(this.form) value="checkbox">
              选中本页显示的所有用户</td>
            <td> <strong>操作：</strong> 
              <input name="Action" type="radio" value="Del" checked onClick="document.myform.UserLevel.disabled=true">删除&nbsp;&nbsp;&nbsp;&nbsp;
              <input name="Action" type="radio" value="Lock" onClick="document.myform.UserLevel.disabled=true">锁定 &nbsp;&nbsp;&nbsp;
              <input name="Action" type="radio" value="UnLock" onClick="document.myform.UserLevel.disabled=true">解锁 &nbsp;&nbsp;&nbsp; 
              <input name="Action" type="radio" value="Move" onClick="document.myform.UserLevel.disabled=false">移动到
              <select name="UserLevel" id="UserLevel" disabled>
                <option value="3000">等待邮件认证的用户</option>
                <option value="2000">等待管理审核的用户</option>
                <option value="999">注册用户</option>
                <option value="99" selected>收费用户</option>
                <option value="9">VIP用户</option>
              </select>
              &nbsp;&nbsp; 
              <input type="submit" name="Submit" value=" 执 行 "> </td>
  </tr>
</table>
</td>
</form></tr></table>
<%
end sub

sub ShowSearch()
%>
<form name="form2" method="post" action="Admin_User.asp">
<table width="100%" border="0" cellspacing="1" cellpadding="2" class="border">
  <tr class="tdbg">
    <td width="120"><strong>用户高级查询：</strong></td>
    <td width="300">
      <select name="Field" id="Field">
      <option value="UserID" selected>用户ID</option>
      <option value="UserName">用户名</option>
      </select>
      <input name="Keyword" type="text" id="Keyword" size="20" maxlength="30">
      <input type="submit" name="Submit2" value=" 查 询 ">
      <input name="UserSearch" type="hidden" id="UserSearch" value="10">
	</td>
    <td>若为空，则查询所有用户</td>
  </tr>
</table>
    </form>
<%
end sub

sub AddUser()
%>
<form name="myform" action="Admin_User.asp" method="post">
  <table width=100% border=0 cellpadding=2 cellspacing=1 class="border">
    <TR align=center class='title'> 
      <TD height=22 colSpan=2><font class=en><b>添 加 新 用 户</b></font></TD>
    </TR>
    <TR class="tdbg" > 
      <TD width="40%"><b>用户名：</b><BR>
        不能超过14个字符（7个汉字）</TD>
      <TD width="60%"> <INPUT   maxLength=14 size=30 name=UserName> <font color="#FF0000">*</font></TD>
    </TR>
    <TR class="tdbg" > 
      <TD width="40%"><B>密码(至少6位)：</B><BR>
        请输入密码，区分大小写。 请不要使用任何类似 '*'、' ' 或 HTML 字符 </TD>
      <TD width="60%"> <INPUT   type=password maxLength=12 size=30 name=Password> 
        <font color="#FF0000">*</font> </TD>
    </TR>
    <TR class="tdbg" > 
      <TD width="40%"><strong>确认密码(至少6位)：</strong><BR>
        请再输一遍确认</TD>
      <TD width="60%"> <INPUT   type=password maxLength=12 size=30 name=PwdConfirm> 
        <font color="#FF0000">*</font> </TD>
    </TR>
    <TR class="tdbg" > 
      <TD width="40%"><strong>密码问题：</strong><BR>
        忘记密码的提示问题</TD>
      <TD width="60%"> <INPUT   type=text maxLength=50 size=30 name="Question"> 
        <font color="#FF0000">*</font> </TD>
    </TR>
    <TR class="tdbg" > 
      <TD width="40%"><strong>问题答案：</strong><BR>
        忘记密码的提示问题答案，用于取回密码</TD>
      <TD width="60%"> <INPUT   type=text maxLength=20 size=30 name="Answer"> 
        <font color="#FF0000">*</font> </TD>
    </TR>
    <TR class="tdbg" > 
      <TD width="40%"><strong>性别：</strong></TD>
      <TD width="60%"> <INPUT type=radio CHECKED value="1" name=sex>
        男 &nbsp;&nbsp; <INPUT type=radio value="0" name=sex>
        女</TD>
    </TR>
    <TR class="tdbg" > 
      <TD width="40%"><strong>Email地址：</strong></TD>
      <TD width="60%"> <INPUT   maxLength=50 size=30 name=Email> <font color="#FF0000">*</font></TD>
    </TR>
    <TR class="tdbg" > 
      <TD width="40%"><strong>主页：</strong></TD>
      <TD width="60%"> <INPUT   maxLength=100 size=30 name=homepage value="http://"></TD>
    </TR>
    <TR class="tdbg" > 
      <TD width="40%"><strong>OICQ号码：</strong></TD>
      <TD width="60%"> <INPUT maxLength=20 size=30 name=OICQ></TD>
    </TR>
    <TR class="tdbg" > 
      <TD width="40%"><strong>个人简介：</strong></TD>
      <TD width="60%"> <textarea name="msn" cols="30" rows="3"></textarea></TD>
    </TR>
    <TR class="tdbg" > 
      <TD width="40%"><strong>用户级别：</strong></TD>
      <TD width="60%"><select name="UserLevel" id="UserLevel">
          <option value="3000">等待邮件认证的用户</option>
          <option value="2000">等待管理审核的用户</option>
          <option value="999" selected>注册用户</option>
          <option value="99">收费用户</option>
          <option value="9">VIP用户</option>
        </select></TD>
    </TR>
    <TR class="tdbg" >
      <TD><strong>计费方式：</strong></TD>
      <TD><input name="ChargeType" type="radio" value="1" checked>
        扣点数<font color="#0000FF">（推荐）</font>：&nbsp;每阅读一篇收费文章，扣除相应点数。&nbsp;<br>
        <input type="radio" name="ChargeType" value="2">
        有效期：在有效期内，用户可以任意阅读收费内容</TD>
    </TR>
    <TR class="tdbg" >
      <TD><strong>用户点数：</strong><br>
        用于阅读需要“阅读点数”文章，在阅读文章时会减去相应的点数<br>
        此功能只有当计费方式为“扣点数”时才有效</TD>
      <TD><input name="UserPoint" type="text" id="UserPoint" value="500" size="10" maxlength="10">
        点</TD>
    </TR>
    <TR class="tdbg" >
      <TD><strong>有效期限：</strong><br>
        若超过此期限，则用户不能阅读收费内容<br>
        此功能只有当计费方式为“有效期限”时才有效</TD>
      <TD>开始日期：
        <input name="BeginDate" type="text" id="BeginDate" value="<%=FormatDateTime(now(),2)%>" size="20" maxlength="20">
      <br>
      有 效 期：
      <input name="Valid_Num" type="text" id="Valid_Num" value="1" size="10" maxlength="10">
      <select name="Valid_Unit" id="Valid_Unit">
      <option value="1">天</option>
      <option value="2">月</option>
      <option value="3" selected>年</option>
      </select>
      </TD>
    </TR>
    <TR class="tdbg" > 
      <TD width="40%"><strong>用户状态：</strong></TD>
      <TD width="60%"><input name="LockUser" type="radio" value="False" checked>
        正常&nbsp;&nbsp; <input type="radio" name="LockUser" value="True">
        锁定</TD>
    </TR>
    <TR align="center" class="tdbg" > 
      <TD colspan="2"><input name="Action" type="hidden" id="Action" value="SaveAdd">
        <input type="submit" name="Submit" value=" 添 加 "></TD>
    </TR>
  </TABLE>
</form>
<%
end sub

sub Modify()
	dim UserID
	dim rsUser,sqlUser
	UserID=trim(request("UserID"))
	if UserID="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>参数不足！</li>"
		exit sub
	else
		UserID=Clng(UserID)
	end if
	Set rsUser=Server.CreateObject("Adodb.RecordSet")
	sqlUser="select * from " & db_User_Table & " where " & db_User_ID & "=" & UserID
	rsUser.Open sqlUser,Conn_User,1,3
	if rsUser.bof and rsUser.eof then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>找不到指定的用户！</li>"
		rsUser.close
		set rsUser=nothing
		exit sub
	end if
%>
<FORM name="Form1" action="Admin_User.asp" method="post">
  <table width="100%" border="0" cellspacing="1" cellpadding="2" class="border">
    <TR class='title'> 
      <TD height=22 colSpan=2 align="center"><b>修改注册用户信息</b></TD>
    </TR>
    <TR class="tdbg" > 
      <TD width="40%"><b>用户名：</b></TD>
      <TD width="60%"><%=rsUser(db_User_Name)%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href="Admin_ArticleManage.asp?Field=Editor&Keyword=<%=rsUser(db_User_Name)%>">查看此用户发表的文章</a></TD>
    </TR>
    <TR class="tdbg" > 
      <TD width="40%"><B>密码(至少6位)：</B><BR>
        请输入密码，区分大小写。 请不要使用任何类似 '*'、' ' 或 HTML 字符 </TD>
      <TD width="60%"> <INPUT   type=password maxLength=16 size=30 name=Password> 
        <font color="#FF0000">如果不想修改，请留空</font> </TD>
    </TR>
    <TR class="tdbg" >
      <TD><strong>确认密码(至少6位)：</strong><br>
        请再输一遍确认</TD>
      <TD><INPUT name=PwdConfirm   type=password id="PwdConfirm" size=30 maxLength=12>
        <font color="#FF0000">如果不想修改，请留空</font> </TD>
    </TR>
    <TR class="tdbg" > 
      <TD width="40%"><strong>密码问题：</strong><br>
        忘记密码的提示问题</TD>
      <TD width="60%"> <INPUT name="Question"   type=text value="<%=rsUser(db_User_Question)%>" size=30> 
      </TD>
    </TR>
    <TR class="tdbg" > 
      <TD width="40%"><strong>问题答案：</strong><BR>
        忘记密码的提示问题答案，用于取回密码</TD>
      <TD width="60%"> <INPUT   type=text size=30 name="Answer"> <font color="#FF0000">如果不想修改，请留空</font></TD>
    </TR>
    <TR class="tdbg" > 
      <TD width="40%"><strong>性别：</strong></TD>
      <TD width="60%"> <INPUT type=radio value="1" name=sex <%if rsUser(db_User_Sex)=1 then response.write "CHECKED"%>>
        男 &nbsp;&nbsp; <INPUT type=radio value="0" name=sex <%if rsUser(db_User_Sex)=0 then response.write "CHECKED"%>>
        女</TD>
    </TR>
    <TR class="tdbg" > 
      <TD width="40%"><strong>Email地址：</strong></TD>
      <TD width="60%"> <INPUT name=Email value="<%=rsUser(db_User_Email)%>" size=30   maxLength=50>
        <a href="mailto:<%=rsUser(db_User_Email)%>">给此用户发一封电子邮件</a> </TD>
    </TR>
    <TR class="tdbg" > 
      <TD width="40%"><strong>主页：</strong></TD>
      <TD width="60%"> <INPUT   maxLength=100 size=30 name=homepage value="<%=rsUser(db_User_Homepage)%>"></TD>
    </TR>
    <TR class="tdbg" > 
      <TD width="40%"><strong>OICQ号码：</strong></TD>
      <TD width="60%"> <INPUT name=OICQ value="<%=rsUser(db_User_QQ)%>" size=30 maxLength=20></TD>
    </TR>
    <TR class="tdbg" > 
      <TD width="40%"><strong>个人简介：</strong></TD>
      <TD width="60%"> <textarea name="msn" cols="30" rows="3"><%=rsUser(db_User_Msn)%></textarea></TD>
    </TR>
    <TR class="tdbg" > 
      <TD width="40%"><strong>用户级别：</strong></TD>
      <TD width="60%"><select name="UserLevel" id="UserLevel">
          <option value="3000" <%if rsUser(db_User_UserLevel)=3000 then response.write " selected"%>>等待邮件认证的用户</option>
          <option value="2000" <%if rsUser(db_User_UserLevel)=2000 then response.write " selected"%>>等待管理审核的用户</option>
          <option value="999" <%if rsUser(db_User_UserLevel)=999 then response.write " selected"%>>注册用户</option>
          <option value="99" <%if rsUser(db_User_UserLevel)=99 then response.write " selected"%>>收费用户</option>
          <option value="9" <%if rsUser(db_User_UserLevel)=9 then response.write " selected"%>>VIP用户</option>
        </select></TD>
    </TR>
    <TR class="tdbg" >
      <TD><strong>计费方式：</strong></TD>
      <TD><input name="ChargeType" type="radio" value="1" <%if rsUser(db_User_ChargeType)=1 then response.write " checked"%>>
        扣点数<font color="#0000FF">（推荐）</font>：&nbsp;每阅读一篇收费文章，扣除相应点数。&nbsp;<br>
        <input type="radio" name="ChargeType" value="2" <%if rsUser(db_User_ChargeType)=2 then response.write " checked"%>>
        有效期：在有效期内，用户可以任意阅读收费内容</TD>
    </TR>
    <TR class="tdbg" >
      <TD><strong>用户点数：</strong><br>
        用于阅读需要“阅读点数”文章，在阅读文章时会减去相应的点数<br>
        此功能只有当计费方式为“扣点数”时才有效</TD>
      <TD><input name="UserPoint" type="text" id="UserPoint" value="<%=rsUser(db_User_UserPoint)%>" size="10" maxlength="10">
        点</TD>
    </TR>
    <TR class="tdbg" >
      <TD><strong>有效期限：</strong><br>
        若超过此期限，则用户不能阅读收费内容<br>
        此功能只有当计费方式为“有效期限”时才有效</TD>
      <TD>开始日期：
        <input name="BeginDate" type="text" id="BeginDate" value="<%=FormatDateTime(rsUser(db_User_BeginDate),2)%>" size="20" maxlength="20">
      <br>
      有 效 期：
      <input name="Valid_Num" type="text" id="Valid_Num" value="<%=rsUser(db_User_Valid_Num)%>" size="10" maxlength="10">
      <select name="Valid_Unit" id="Valid_Unit">
      <option value="1" <%if rsUser(db_User_Valid_Unit)=1 then response.write " selected"%>>天</option>
      <option value="2" <%if rsUser(db_User_Valid_Unit)=2 then response.write " selected"%>>月</option>
      <option value="3" <%if rsUser(db_User_Valid_Unit)=3 then response.write " selected"%>>年</option>
      </select>
      </TD>
    </TR>
    <TR class="tdbg" > 
      <TD width="40%"><strong>用户状态：</strong></TD>
      <TD width="60%"><input type="radio" name="LockUser" value="False" <%if rsUser(db_User_LockUser)=False then response.write "checked"%>>
        正常&nbsp;&nbsp; <input type="radio" name="LockUser" value="True" <%if rsUser(db_User_LockUser)=True then response.write "checked"%>>
        锁定</TD>
    </TR>
    <TR class="tdbg" > 
      <TD height="40" colspan="2" align="center"><input name="Action" type="hidden" id="Action" value="SaveModify"> 
        <input name=Submit   type=submit id="Submit" value="保存修改结果"> <input name="UserID" type="hidden" id="UserID" value="<%=rsUser(db_User_ID)%>"></TD>
    </TR>
  </TABLE>
</form>
<%
	rsUser.close
	set rsUser=nothing
end sub

sub AddMoney()
	dim UserID
	dim rsUser,sqlUser
	UserID=trim(request("UserID"))
	if UserID="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>参数不足！</li>"
		exit sub
	else
		UserID=Clng(UserID)
	end if
	Set rsUser=Server.CreateObject("Adodb.RecordSet")
	sqlUser="select * from " & db_User_Table & " where " & db_User_ID & "=" & UserID
	rsUser.Open sqlUser,Conn_User,1,3
	if rsUser.bof and rsUser.eof then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>找不到指定的用户！</li>"
		rsUser.close
		set rsUser=nothing
		exit sub
	end if
	if rsUser(db_User_UserLevel)>99 then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>此用户不是收费用户或VIP用户，无需续费！</li>"
		rsUser.close
		set rsUser=nothing
		exit sub
	end if
%>
<FORM name="Form1" action="Admin_User.asp" method="post">
  <table width="100%" border="0" cellspacing="1" cellpadding="2" class="border">
    <TR class='title'> 
      <TD height=22 colSpan=2 align="center"><b>用 户 续 费</b></TD>
    </TR>
    <TR class="tdbg" > 
      <TD width="40%"><b>用户名：</b></TD>
      <TD width="60%"><%=rsUser(db_User_Name)%></TD>
    </TR>
    <TR class="tdbg" > 
      <TD width="40%"><strong>用户级别：</strong></TD>
      <TD width="60%"><%
	  if rsUser(db_User_UserLevel)=99 then
	  	response.write "收费用户"      
	  elseif rsUser(db_User_UserLevel)=9 then
	  	response.write "VIP用户"
	  end if
      %></TD>
    </TR>
    <TR class="tdbg" >
      <TD><strong>计费方式：</strong></TD>
      <TD><%
	  if rsUser(db_User_ChargeType)=1 then
	   	response.write "扣点数：&nbsp;每阅读一篇收费文章，扣除相应点数。"
      else
	   	response.write "有效期：在有效期内，用户可以任意阅读收费内容"
	  end if
	  %>
        <input name="ChargeType" type="hidden" id="ChargeType" value="<%=rsUser(db_User_ChargeType)%>">
		</TD>
    </TR>
    <%if rsUser(db_User_ChargeType)=1 then%>
	<TR class="tdbg" >
      <TD><strong>目前的用户点数：</strong></TD>
      <TD><%=rsUser(db_User_UserPoint)%> 点</TD>
    </TR>
    <TR class="tdbg" >
      <TD><strong>追加点数：</strong></TD>
      <TD> <input name="UserPoint" type="text" id="UserPoint" value="100" size="10" maxlength="10">
      点</TD>
    </TR>
	<%else%>
    <TR class="tdbg" >
      <TD><strong>目前的有效期限信息：</strong></TD>
      <TD><%
	  response.write "开始计算日期" & FormatDateTime(rsUser(db_User_BeginDate),2) & "&nbsp;&nbsp;&nbsp;&nbsp;有 效 期：" & rsUser(db_User_Valid_Num)
	  if rsUser(db_User_Valid_Unit)=1 then
	  	ValidDays=rsUser(db_User_Valid_Num)
	  	response.write "天"
	  elseif rsUser(db_User_Valid_Unit)=2 then
	  	ValidDays=rsUser(db_User_Valid_Num)*30
	  	response.write "月"
	  elseif rsUser(db_User_Valid_Unit)=3 then
	  	ValidDays=rsUser(db_User_Valid_Num)*365
	  	response.write "年"
	  end if
	  response.write "<br>"
	  tmpDays=ValidDays-DateDiff("D",rsUser(db_User_BeginDate),now())
	  if tmpDays>=0 then
	  	response.write "尚有 <font color=blue>" & tmpDays & "</font> 天到期"
	  else
	  	response.write "已经过期 <font color=red>" & abs(tmpDays) & "</font> 天"
	  end if
	  %>
      </TD>
    </TR>
	<tr class="tdbg" >
	  <td><strong>追加天数：</strong><br>
	    若目前用户尚未到期，则追加相应天数<br>
	    若目前用户已经过了有效期，则有效期从续费之日起重新计数。</td>
	  <td>
      <input name="Valid_Num" type="text" id="Valid_Num" value="1" size="10" maxlength="10">
      <select name="Valid_Unit" id="Valid_Unit" <%if tmpDays>0 then response.write " disabled"%>>
        <option value="1" <%if rsUser(db_User_Valid_Unit)=1 then response.write " selected"%>>天</option>
        <option value="2" <%if rsUser(db_User_Valid_Unit)=2 then response.write " selected"%>>月</option>
        <option value="3" <%if rsUser(db_User_Valid_Unit)=3 then response.write " selected"%>>年</option>
      </select>
	  </td>
	</tr>
	<%end if%>
    <TR class="tdbg" > 
      <TD height="40" colspan="2" align="center"><input name="Action" type="hidden" id="Action" value="SaveAddMoney"> 
        <input name=Submit   type=submit id="Submit" value="保存续费结果"> <input name="UserID" type="hidden" id="UserID" value="<%=rsUser(db_User_ID)%>"></TD>
    </TR>
  </TABLE>
</form>
<%
	rsUser.close
	set rsUser=nothing
end sub

sub UpdateUser()
%>
<FORM name="Form1" action="Admin_User.asp" method="post">
<table width="100%" border="0" cellspacing="1" cellpadding="2" class="border">
  <tr align="center" class="title"> 
    <td height="22" colspan="2"><strong>更 新 用 户 数 据</strong></td>
  </tr>
  <tr class="tdbg"> 
      <td colspan="2"><p>说明：<br>
          1、本操作将重新计算用户的发表文章数。<br>
          2、本操作可能将非常消耗服务器资源，而且更新时间很长，请仔细确认每一步操作后执行。</p>
      </td>
  </tr>
  <tr class="tdbg"> 
    <td height="25">开始用户ID：</td>
    <td height="25"><input name="BeginID" type="text" id="BeginID" value="1" size="15" maxlength="10">
      用户ID，可以填写您想从哪一个ID号开始进行修复</td>
  </tr>
  <tr class="tdbg"> 
    <td height="25">结束用户ID：</td>
    <td height="25"><input name="EndID" type="text" id="EndID" value="1000" size="15" maxlength="10">
      将更新开始到结束ID之间的用户数据，之间的数值最好不要选择过大</td>
  </tr>
  <tr class="tdbg"> 
    <td height="25">&nbsp;</td>
    <td height="25"><input name="Submit" type="submit" id="Submit" value="更新用户数据"> <input name="Action" type="hidden" id="Action" value="DoUpdate"></td>
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
	dim UserName,Password,PwdConfirm,Question,Answer,Sex,Email,Homepage,OICQ,MSN,UserLevel,LockUser,ChargeType,UserPoint,BeginDate,Valid_Num,Valid_Unit
	UserName=trim(request("UserName"))
	Password=trim(request("Password"))
	PwdConfirm=trim(request("PwdConfirm"))
	Question=trim(request("Question"))
	Answer=trim(request("Answer"))
	Sex=trim(Request("Sex"))
	Email=trim(request("Email"))
	Homepage=trim(request("Homepage"))
	OICQ=trim(request("OICQ"))
	MSN=trim(request("MSN"))
	UserLevel=trim(request("UserLevel"))
	LockUser=trim(request("LockUser"))
	ChargeType=trim(request("ChargeType"))
	UserPoint=trim(request("UserPoint"))
	BeginDate=trim(request("BeginDate"))
	Valid_Num=trim(request("Valid_Num"))
	Valid_Unit=trim(request("Valid_Unit"))
	
	if UserName="" or strLength(UserName)>14 or strLength(UserName)<4 then
		founderr=true
		errmsg=errmsg & "<br><li>请输入用户名(不能大于14小于4)</li>"
	else
  		if Instr(UserName,"=")>0 or Instr(UserName,"%")>0 or Instr(UserName,chr(32))>0 or Instr(UserName,"?")>0 or Instr(UserName,"&")>0 or Instr(UserName,";")>0 or Instr(UserName,",")>0 or Instr(UserName,"'")>0 or Instr(UserName,",")>0 or Instr(UserName,chr(34))>0 or Instr(UserName,chr(9))>0 or Instr(UserName,"")>0 or Instr(UserName,"$")>0 then
			errmsg=errmsg+"<br><li>用户名中含有非法字符</li>"
			founderr=true
		end if
	end if
	if Password="" or strLength(Password)>12 or strLength(Password)<6 then
		founderr=true
		errmsg=errmsg & "<br><li>请输入密码(不能大于12小于6)</li>"
	else
		if Instr(Password,"=")>0 or Instr(Password,"%")>0 or Instr(Password,chr(32))>0 or Instr(Password,"?")>0 or Instr(Password,"&")>0 or Instr(Password,";")>0 or Instr(Password,",")>0 or Instr(Password,"'")>0 or Instr(Password,",")>0 or Instr(Password,chr(34))>0 or Instr(Password,chr(9))>0 or Instr(Password,"")>0 or Instr(Password,"$")>0 then
			errmsg=errmsg+"<br><li>密码中含有非法字符</li>"
			founderr=true
		end if
	end if
	if PwdConfirm="" then
		founderr=true
		errmsg=errmsg & "<br><li>请输入确认密码(不能大于12小于6)</li>"
	else
		if Password<>PwdConfirm then
			founderr=true
			errmsg=errmsg & "<br><li>密码和确认密码不一致</li>"
		end if
	end if
	if Question="" then
		founderr=true
		errmsg=errmsg & "<br><li>密码提示问题不能为空</li>"
	end if
	if Answer="" then
		founderr=true
		errmsg=errmsg & "<br><li>密码答案不能为空</li>"
	end if
	if Sex="" then
		founderr=true
		errmsg=errmsg & "<br><li>性别不能为空</li>"
	else
		sex=cint(sex)
		if Sex<>0 and Sex<>1 then
			Sex=1
		end if
	end if
	if Email="" then
		founderr=true
		errmsg=errmsg & "<br><li>Email不能为空</li>"
	else
		if IsValidEmail(Email)=false then
			errmsg=errmsg & "<br><li>您的Email有错误</li>"
   			founderr=true
		end if
	end if
	if OICQ<>"" then
		if not isnumeric(OICQ) or len(cstr(OICQ))>10 then
			errmsg=errmsg & "<br><li>OICQ号码只能是4-10位数字，您可以选择不输入。</li>"
			founderr=true
		end if
	end if
	
	if UserLevel="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>请指定用户级别！</li>"
	else
		UserLevel=CLng(UserLevel)
	end if
	if LockUser="True" then
		LockUser=True
	else
		LockUser=False
	end if
	if ChargeType="" then
		ChargeType=1
	else
		ChargeType=Clng(ChargeType)
	end if
	if UserPoint="" then
		UserPoint=0
	else
		UserPoint=Clng(UserPoint)
	end if
	if BeginDate="" then
		BeginDate=now()
	else
		BeginDate=Cdate(BeginDate)
	end if
	if Valid_Num="" then
		Valid_Num=0
	else
		Valid_Num=Clng(Valid_Num)
	end if
	if Valid_Unit="" then
		Valid_Unit=1
	else
		Valid_Unit=Clng(Valid_Unit)
	end if
	if (UserLevel=99 or UserLevel=9) then
		if ChargeType=1 and UserPoint=0 then
			FoundErr=True
			ErrMsg=ErrMsg & "<br><li>请输入用户点数！</li>"
		end if
		if ChargeType=2 and Valid_Num=0 then
			FoundErr=True
			ErrMsg=ErrMsg & "<br><li>请输入有效期限</li>"
		end if
	end if

	if founderr=true then
		exit sub
	end if
	
	dim sqlReg,rsReg
	sqlReg="select * from " & db_User_Table & " where " & db_User_Name & "='" & Username & "'"
	set rsReg=server.createobject("adodb.recordset")
	rsReg.open sqlReg,Conn_User,1,3
	if not(rsReg.bof and rsReg.eof) then
		founderr=true
		errmsg=errmsg & "<br><li>你注册的用户已经存在！请换一个用户名再试试！</li>"
	else
		rsReg.addnew
		rsReg(db_User_Name)=UserName
		rsReg(db_User_Password)=md5(Password)
		rsReg(db_User_Question)=Question
		rsReg(db_User_Answer)=md5(Answer)
		rsReg(db_User_Sex)=Sex
		rsReg(db_User_Email)=Email
		rsReg(db_User_Homepage)=Homepage
		rsReg(db_User_QQ)=OICQ
		rsReg(db_User_Msn)=MSN
		rsReg(db_User_UserLevel)=UserLevel
		rsReg(db_User_LockUser)=LockUser
		rsReg(db_User_RegDate)=Now()
		rsReg(db_User_ChargeType)=ChargeType
		rsReg(db_User_UserPoint)=UserPoint
		rsReg(db_User_BeginDate)=BeginDate
		rsReg(db_User_Valid_Num)=Valid_Num
		rsReg(db_User_Valid_Unit)=Valid_Unit

		if UserTableType="Dvbbs6.0" or UserTableType="Dvbbs6.1" then
			rsReg(db_User_UserClass) = FU_UserClass
			rsReg(db_User_TitlePic) = FU_TitlePic
			rsReg(db_User_UserGroupID) = FU_UserGroupID
			rsReg(db_User_Face) = FU_Face
			rsReg(db_User_FaceWidth) = FU_FaceWidth
			rsReg(db_User_FaceHeight) = FU_FaceHeight
			rsReg(db_User_UserWealth) = FU_UserWealth
			rsReg(db_User_UserEP) = FU_UserEP
			rsReg(db_User_UserCP) = FU_UserCP
			rsReg(db_User_UserGroup) = FU_UserGroup
			rsReg(db_User_Showre) = FU_Showre
		end if 

		rsReg.update
		call UpdateUserNum(UserName)

		founderr=false
	end if
	rsReg.close
	set rsReg=nothing
	call CloseConn_User()
	response.Redirect "Admin_User.asp"
end sub		

sub SaveModify()
	dim UserID,Password,PwdConfirm,Question,Answer,Sex,Email,Homepage,OICQ,MSN,UserLevel,LockUser,ChargeType,UserPoint,BeginDate,Valid_Num,Valid_Unit
	dim rsUser,sqlUser
	Action=trim(request("Action"))
	UserID=trim(request("UserID"))
	if UserID="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>参数不足！</li>"
		exit sub
	else
		UserID=Clng(UserID)
	end if
	Password=trim(request("Password"))
	PwdConfirm=trim(request("PwdConfirm"))
	Question=trim(request("Question"))
	Answer=trim(request("Answer"))
	Sex=trim(Request("Sex"))
	Email=trim(request("Email"))
	Homepage=trim(request("Homepage"))
	OICQ=trim(request("OICQ"))
	MSN=trim(request("MSN"))
	UserLevel=trim(request("UserLevel"))
	LockUser=trim(request("LockUser"))
	ChargeType=trim(request("ChargeType"))
	UserPoint=trim(request("UserPoint"))
	BeginDate=trim(request("BeginDate"))
	Valid_Num=trim(request("Valid_Num"))
	Valid_Unit=trim(request("Valid_Unit"))

	if Password<>"" then
		if strLength(Password)>12 or strLength(Password)<6 then
			founderr=true
			errmsg=errmsg & "<br><li>密码不能大于12小于6，如果你不想修改密码，请保持为空。</li>"
		end if
		if Instr(Password,"=")>0 or Instr(Password,"%")>0 or Instr(Password,chr(32))>0 or Instr(Password,"?")>0 or Instr(Password,"&")>0 or Instr(Password,";")>0 or Instr(Password,",")>0 or Instr(Password,"'")>0 or Instr(Password,",")>0 or Instr(Password,chr(34))>0 or Instr(Password,chr(9))>0 or Instr(Password,"")>0 or Instr(Password,"$")>0 then
			errmsg=errmsg+"<br><li>密码中含有非法字符，如果你不想修改密码，请保持为空。</li>"
			founderr=true
		end if
	end if
	if Password<>PwdConfirm then
		founderr=true
		errmsg=errmsg & "<br><li>密码和确认密码不一致</li>"
	end if
	if Question="" then
		founderr=true
		errmsg=errmsg & "<br><li>密码提示问题不能为空</li>"
	end if
	if Sex="" then
		founderr=true
		errmsg=errmsg & "<br><li>性别不能为空</li>"
	else
		sex=cint(sex)
		if Sex<>0 and Sex<>1 then
			Sex=1
		end if
	end if
	if Email="" then
		founderr=true
		errmsg=errmsg & "<br><li>Email不能为空</li>"
	else
		if IsValidEmail(Email)=false then
			errmsg=errmsg & "<br><li>您的Email有错误</li>"
   			founderr=true
		end if
	end if
	if OICQ<>"" then
		if not isnumeric(OICQ) or len(cstr(OICQ))>10 then
			errmsg=errmsg & "<br><li>OICQ号码只能是4-10位数字，您可以选择不输入。</li>"
			founderr=true
		end if
	end if
	
	if UserLevel="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>请指定用户级别！</li>"
	else
		UserLevel=CLng(UserLevel)
	end if
	if LockUser="True" then
		LockUser=True
	else
		LockUser=False
	end if
	if ChargeType="" then
		ChargeType=1
	else
		ChargeType=Clng(ChargeType)
	end if
	if UserPoint="" then
		UserPoint=0
	else
		UserPoint=Clng(UserPoint)
	end if
	if BeginDate="" then
		BeginDate=now()
	else
		BeginDate=Cdate(BeginDate)
	end if
	if Valid_Num="" then
		Valid_Num=0
	else
		Valid_Num=Clng(Valid_Num)
	end if
	if Valid_Unit="" then
		Valid_Unit=1
	else
		Valid_Unit=Clng(Valid_Unit)
	end if
	if (UserLevel=99 or UserLevel=9) then
		if ChargeType=1 and UserPoint=0 then
			FoundErr=True
			ErrMsg=ErrMsg & "<br><li>请输入用户点数！</li>"
		end if
		if ChargeType=2 and Valid_Num=0 then
			FoundErr=True
			ErrMsg=ErrMsg & "<br><li>请输入有效期限</li>"
		end if
	end if

	if founderr=true then
		exit sub
	end if
	
	Set rsUser=Server.CreateObject("Adodb.RecordSet")
	sqlUser="select * from " & db_User_Table & " where " & db_User_ID & "=" & UserID
	rsUser.Open sqlUser,Conn_User,1,3
	if rsUser.bof and rsUser.eof then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>找不到指定的用户！</li>"
		rsUser.close
		set rsUser=nothing
		exit sub
	end if
	if Password<>"" then
		rsUser(db_User_Password)=md5(Password)
	end if
	rsUser(db_User_Question)=Question
	if Answer<>"" then
		rsUser(db_User_Answer)=md5(Answer)
	end if
	rsUser(db_User_Sex)=Sex
	rsUser(db_User_Email)=Email
	rsUser(db_User_Homepage)=HomePage
	rsUser(db_User_QQ)=OICQ
	rsUser(db_User_Msn)=MSN
	rsUser(db_User_UserLevel)=UserLevel
	rsUser(db_User_LockUser)=LockUser
	rsUser(db_User_ChargeType)=ChargeType
	rsUser(db_User_UserPoint)=UserPoint
	rsUser(db_User_BeginDate)=BeginDate
	rsUser(db_User_Valid_Num)=Valid_Num
	rsUser(db_User_Valid_Unit)=Valid_Unit
	rsUser.update
	rsUser.Close
	set rsUser=nothing
	call CloseConn_User()
	response.redirect "Admin_User.asp"
end sub

sub SaveAddMoney()
	dim UserID,ChargeType,UserPoint,Valid_Num,Valid_Unit
	dim rsUser,sqlUser
	Action=trim(request("Action"))
	UserID=trim(request("UserID"))
	if UserID="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>参数不足！</li>"
		exit sub
	else
		UserID=Clng(UserID)
	end if
	ChargeType=trim(request("ChargeType"))
	UserPoint=trim(request("UserPoint"))
	Valid_Num=trim(request("Valid_Num"))
	Valid_Unit=trim(request("Valid_Unit"))

	if ChargeType="" then
		ChargeType=1
	else
		ChargeType=Clng(ChargeType)
	end if
	if UserPoint="" then
		UserPoint=0
	else
		UserPoint=Clng(UserPoint)
	end if
	if Valid_Num="" then
		Valid_Num=0
	else
		Valid_Num=Clng(Valid_Num)
	end if
		if Valid_Unit="" then
		Valid_Unit=1
	else
		Valid_Unit=Clng(Valid_Unit)
	end if
	
	if ChargeType=1 and UserPoint=0 then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>请输入要追加的用户点数！</li>"
	end if
	if ChargeType=2 and Valid_Num=0 then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>请输入要追加的天数</li>"
	end if

	if founderr=true then
		exit sub
	end if
	
	Set rsUser=Server.CreateObject("Adodb.RecordSet")
	sqlUser="select * from " & db_User_Table & " where " & db_User_ID & "=" & UserID
	rsUser.Open sqlUser,Conn_User,1,3
	if rsUser.bof and rsUser.eof then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>找不到指定的用户！</li>"
		rsUser.close
		set rsUser=nothing
		exit sub
	end if
	if ChargeType=1 then
		rsUser(db_User_UserPoint)=rsUser(db_User_UserPoint)+UserPoint
	else
		if rsUser(db_User_Valid_Unit)=1 then
			ValidDays=rsUser(db_User_Valid_Num)
		elseif rsUser(db_User_Valid_Unit)=2 then
			ValidDays=rsUser(db_User_Valid_Num)*30
		elseif rsUser(db_User_Valid_Unit)=3 then
			ValidDays=rsUser(db_User_Valid_Num)*365
		end if
		tmpDays=ValidDays-DateDiff("D",rsUser(db_User_BeginDate),now())
		if tmpDays>0 then
			rsUser(db_User_Valid_Num)=rsUser(db_User_Valid_Num)+Valid_Num
		else
			rsUser(db_User_BeginDate)=now()
			rsUser(db_User_Valid_Num)=Valid_Num
			rsUser(db_User_Valid_Unit)=Valid_Unit
		end if
	end if
	rsUser.update
	rsUser.Close
	set rsUser=nothing
	call CloseConn_User()
	response.redirect "Admin_User.asp"
end sub

sub DelUser()
	if UserID="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>请指定要删除的用户</li>"
		exit sub
	end if
	if instr(UserID,",")>0 then
		UserID=replace(UserID," ","")
		sql="delete from " & db_User_Table & " where " & db_User_ID & " in (" & UserID & ")"
	else
		sql="delete from " & db_User_Table & " where " & db_User_ID & "=" & Clng(UserID)
	end if
	Conn_User.Execute sql
	call CloseConn_User()      
	response.redirect ComeUrl
end sub

sub LockUser()
	if UserID="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>请选择要锁定的用户</li>"
		exit sub
	end if
	if instr(UserID,",")>0 then
		UserID=replace(UserID," ","")
		sql="Update " & db_User_Table & " set " & db_User_LockUser & "=true where " & db_User_ID & " in (" & UserID & ")"
	else
		sql="Update " & db_User_Table & " set " & db_User_LockUser & "=true where " & db_User_ID & "=" & CLng(UserID)
	end if
	Conn_User.Execute sql
	call CloseConn_User()      
	response.redirect ComeUrl
end sub

sub UnLockUser()
	if UserID="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>请指定要解锁的用户</li>"
		exit sub
	end if
	if instr(UserID,",")>0 then
		UserID=replace(UserID," ","")
		sql="Update " & db_User_Table & " set " & db_User_LockUser & "=False where " & db_User_ID & " in (" & UserID & ")"
	else
		sql="Update " & db_User_Table & " set " & db_User_LockUser & "=False where " & db_User_ID & "=" & CLng(UserID)
	end if
	Conn_User.Execute sql
	call CloseConn_User()      
	response.redirect ComeUrl
end sub

sub MoveUser()
	dim msg
	if UserID="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>请指定要移动的用户</li>"
		exit sub
	end if
	dim UserLevel
	UserLevel=trim(request("UserLevel"))
	if UserLevel="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>请指定目标用户组</li>"
		exit sub
	else
		UserLevel=Clng(UserLevel)
	end if
	if instr(UserID,",")>0 then
		UserID=replace(UserID," ","")
		if UserLevel=999 then
			msg="&nbsp;&nbsp;&nbsp;&nbsp;已经成功将选定用户设为“<font color=blue>普通注册用户</font>”！并且按照你在[网站选项---用户选项]中的预设值给设定了这些用户的计费方式、初始点数、有效期等数据。"
			msg=msg & "<br><br>计费方式："
			if ChargeType_999=1 then
				msg=msg & "扣点数<br>初始点数：" & UserPoint_999 & "点"
			else
				msg=msg & "有效期<br>开始日期：" & formatdatetime(now(),2) & "<br>有 效 期：" & ValidDays_999 & "天"
			end if
			sql="Update " & db_User_Table & " set " & db_User_UserLevel & "=" & UserLevel & "," & db_User_ChargeType & "=" & ChargeType_999 & "," & db_User_UserPoint & "=" & UserPoint_999 & "," & db_User_BeginDate & "=#" & formatdatetime(now(),2) & "#," & db_User_Valid_Num & "=" & ValidDays_999 & "," & db_User_Valid_Unit & "=1 where " & db_User_ID & " in (" & UserID & ")"
		elseif UserLevel=99 then
			msg="&nbsp;&nbsp;&nbsp;&nbsp;已经成功将选定用户设为“<font color=blue>收费用户</font>”！并且按照你在[网站选项---用户选项]中的预设值给设定了这些用户的计费方式、初始点数、有效期等数据。"
			msg=msg & "<br><br>计费方式："
			if ChargeType_99=1 then
				msg=msg & "扣点数<br>初始点数：" & UserPoint_99 & "点"
			else
				msg=msg & "有效期<br>开始日期：" & formatdatetime(now(),2) & "<br>有 效 期：" & ValidDays_99 & "天"
			end if
			sql="Update " & db_User_Table & " set " & db_User_UserLevel & "=" & UserLevel & "," & db_User_ChargeType & "=" & ChargeType_99 & "," & db_User_UserPoint & "=" & UserPoint_99 & "," & db_User_BeginDate & "=#" & formatdatetime(now(),2) & "#," & db_User_Valid_Num & "=" & ValidDays_99 & "," & db_User_Valid_Unit & "=1 where " & db_User_ID & " in (" & UserID & ")"
		elseif UserLevel=9 then
			msg="&nbsp;&nbsp;&nbsp;&nbsp;已经成功将选定用户设为“<font color=blue>VIP用户</font>”！并且按照你在[网站选项---用户选项]中的预设值给设定了这些用户的计费方式、初始点数、有效期等数据。"
			msg=msg & "<br><br>计费方式："
			if ChargeType_9=1 then
				msg=msg & "扣点数<br>初始点数：" & UserPoint_9 & "点"
			else
				msg=msg & "有效期<br>开始日期：" & formatdatetime(now(),2) & "<br>有 效 期：" & ValidDays_9 & "天"
			end if
			sql="Update " & db_User_Table & " set " & db_User_UserLevel & "=" & UserLevel & "," & db_User_ChargeType & "=" & ChargeType_9 & "," & db_User_UserPoint & "=" & UserPoint_9 & "," & db_User_BeginDate & "=#" & formatdatetime(now(),2) & "#," & db_User_Valid_Num & "=" & ValidDays_9 & "," & db_User_Valid_Unit & "=1 where " & db_User_ID & " in (" & UserID & ")"
		end if
	else
		if UserLevel=999 then
			msg="&nbsp;&nbsp;&nbsp;&nbsp;已经成功将选定用户设为“<font color=blue>普通注册用户</font>”！并且按照你在[网站选项---用户选项]中的预设值给设定了这些用户的计费方式、初始点数、有效期等数据。"
			msg=msg & "<br><br>计费方式："
			if ChargeType_999=1 then
				msg=msg & "扣点数<br>初始点数：" & UserPoint_999 & "点"
			else
				msg=msg & "有效期<br>开始日期：" & formatdatetime(now(),2) & "<br>有 效 期：" & ValidDays_999 & "天"
			end if
			sql="Update " & db_User_Table & " set " & db_User_UserLevel & "=" & UserLevel & "," & db_User_ChargeType & "=" & ChargeType_999 & "," & db_User_UserPoint & "=" & UserPoint_999 & "," & db_User_BeginDate & "=#" & formatdatetime(now(),2) & "#," & db_User_Valid_Num & "=" & ValidDays_999 & "," & db_User_Valid_Unit & "=1 where " & db_User_ID & "=" & CLng(UserID)
		elseif UserLevel=99 then
			msg="&nbsp;&nbsp;&nbsp;&nbsp;已经成功将选定用户设为“<font color=blue>收费用户</font>”！并且按照你在[网站选项---用户选项]中的预设值给设定了这些用户的计费方式、初始点数、有效期等数据。"
			msg=msg & "<br><br>计费方式："
			if ChargeType_99=1 then
				msg=msg & "扣点数<br>初始点数：" & UserPoint_99 & "点"
			else
				msg=msg & "有效期<br>开始日期：" & formatdatetime(now(),2) & "<br>有 效 期：" & ValidDays_99 & "天"
			end if
			sql="Update " & db_User_Table & " set " & db_User_UserLevel & "=" & UserLevel & "," & db_User_ChargeType & "=" & ChargeType_99 & "," & db_User_UserPoint & "=" & UserPoint_99 & "," & db_User_BeginDate & "=#" & formatdatetime(now(),2) & "#," & db_User_Valid_Num & "=" & ValidDays_99 & "," & db_User_Valid_Unit & "=1 where " & db_User_ID & "=" & CLng(UserID)
		elseif UserLevel=9 then
			msg="&nbsp;&nbsp;&nbsp;&nbsp;已经成功将选定用户设为“<font color=blue>VIP用户</font>”！并且按照你在[网站选项---用户选项]中的预设值给设定了这些用户的计费方式、初始点数、有效期等数据。"
			msg=msg & "<br><br>计费方式："
			if ChargeType_9=1 then
				msg=msg & "扣点数<br>初始点数：" & UserPoint_9 & "点"
			else
				msg=msg & "有效期<br>开始日期：" & formatdatetime(now(),2) & "<br>有 效 期：" & ValidDays_9 & "天"
			end if
			sql="Update " & db_User_Table & " set " & db_User_UserLevel & "=" & UserLevel & "," & db_User_ChargeType & "=" & ChargeType_9 & "," & db_User_UserPoint & "=" & UserPoint_9 & "," & db_User_BeginDate & "=#" & formatdatetime(now(),2) & "#," & db_User_Valid_Num & "=" & ValidDays_9 & "," & db_User_Valid_Unit & "=1 where " & db_User_ID & "=" & CLng(UserID)
		end if
	end if
	Conn_User.Execute sql
	
	call WriteSuccessMsg(msg)
end sub

sub DoUpdate()
	dim BeginID,EndID,sqlUser,rsUser,trs
	BeginID=trim(request("BeginID"))
	EndID=trim(request("EndID"))
	if BeginID="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>请指定开始ID</li>"
	else
		BeginID=Clng(BeginID)
	end if
	if EndID="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>请指定结束ID</li>"
	else
		EndID=Clng(EndID)
	end if
	
	if FoundErr=True then exit sub
	
	Set rsUser=Server.CreateObject("Adodb.RecordSet")
	sqlUser="select * from " & db_User_Table & " where " & db_User_ID & ">=" & BeginID & " and " & db_User_ID & "<=" & EndID
	rsUser.Open sqlUser,Conn_User,1,3
	do while not rsUser.eof
		set trs=Conn.execute("select count(ArticleID) from Article where Editor='" & rsUser(db_User_Name) & "'")
		if isNull(trs(0)) then
			rsUser(db_User_ArticleCount)=0
		else
			rsUser(db_User_ArticleCount)=trs(0)
		end if
		set trs=Conn.execute("select count(ArticleID) from Article where Passed=True and Editor='" & rsUser(db_User_Name) & "'")
		if isNull(trs(0)) then
			rsUser(db_User_ArticleChecked)=0
		else
			rsUser(db_User_ArticleChecked)=trs(0)
		end if
		rsUser.update
		rsUser.movenext
	loop
	rsUser.close
	set rsUser=nothing
	call WriteSuccessMsg("已经成功将用户数据进行了更新！")
end sub
%>