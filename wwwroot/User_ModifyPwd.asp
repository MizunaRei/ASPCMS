<!--#include file="inc/Conn_User.asp"-->
<!--#include file="inc/md5.asp"-->
<!--#include file="inc/function.asp"-->
<%
if CheckUserLogined()=False then
	response.Redirect "User_Login.asp"
end if

dim Action,FoundErr,ErrMsg
dim OldPassword,Password,PwdConfirm
dim rsUser,sqlUser
Action=trim(request("Action"))
UserName=trim(request("UserName"))
OldPassword=trim(request("OldPassword"))
Password=trim(request("Password"))
PwdConfirm=trim(request("PwdConfirm"))
if Action="Modify" and UserName<>"" then
	Set rsUser=Server.CreateObject("Adodb.RecordSet")
	sqlUser="select * from " & db_User_Table & " where " & db_User_Name & "='" & UserName & "'"
	rsUser.Open sqlUser,Conn_User,1,3
	if rsUser.bof and rsUser.eof then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>找不到指定的用户！</li>"
	else
		if OldPassword="" then
			FoundErr=True
			ErrMsg=ErrMsg & "<br><li>请输入旧密码！</li>"
		else
			if Instr(OldPassword,"=")>0 or Instr(OldPassword,"%")>0 or Instr(OldPassword,chr(32))>0 or Instr(OldPassword,"?")>0 or Instr(OldPassword,"&")>0 or Instr(OldPassword,";")>0 or Instr(OldPassword,",")>0 or Instr(OldPassword,"'")>0 or Instr(OldPassword,",")>0 or Instr(OldPassword,chr(34))>0 or Instr(OldPassword,chr(9))>0 or Instr(OldPassword,"")>0 or Instr(OldPassword,"$")>0 then
				errmsg=errmsg+"<br><li>旧密码中含有非法字符</li>"
				founderr=true
			else
				if md5(OldPassword)<>rsUser(db_User_Password) then
					FoundErr=True
					ErrMsg=ErrMsg & "<br><li>你输入的旧密码不对，没有权限修改！</li>"
				end if
			end if
		end if
		if strLength(Password)>12 or strLength(Password)<6 then
			founderr=true
			errmsg=errmsg & "<br><li>请输入新密码(不能大于12小于6)。</li>"
		else
			if Instr(Password,"=")>0 or Instr(Password,"%")>0 or Instr(Password,chr(32))>0 or Instr(Password,"?")>0 or Instr(Password,"&")>0 or Instr(Password,";")>0 or Instr(Password,",")>0 or Instr(Password,"'")>0 or Instr(Password,",")>0 or Instr(Password,chr(34))>0 or Instr(Password,chr(9))>0 or Instr(Password,"")>0 or Instr(Password,"$")>0 then
				errmsg=errmsg+"<br><li>新密码中含有非法字符</li>"
				founderr=true
			end if
		end if
		if PwdConfirm="" then
			FoundErr=True
			ErrMsg=ErrMsg & "<br><li>请输入确认密码！</li>"
		else
			if PwdConfirm<>Password then
				FoundErr=True
				ErrMsg=ErrMsg & "<br><li>确认密码与新密码不一致！</li>"
			end if
		end if
		if FoundErr<>true then
			Password=md5(Password)
			rsUser(db_User_Password)=Password
			rsUser.update
			Response.Cookies("asp163")("Password") = PassWord
		end if
	end if
	rsUser.close
	set rsUser=nothing
	if FoundErr=True then
		call WriteErrMsg()
	else
		call WriteSuccessMsg("成功修改密码！")
	end if
else
%>
<html>
<head>
<title>修改用户密码</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="Admin_STYLE.CSS" rel="stylesheet" type="text/css">
</head>
<body>
<br>
<FORM name="Form1" action="User_ModifyPwd.asp" method="post">
  <table width=400 border=0 align="center" cellpadding=2 cellspacing=1 class='border'>
    <TR align=center class='title'> 
      <TD height=22 colSpan=2><font class=en><b>修改用户密码</b></font></TD>
    </TR>
    <TR class="tdbg" > 
      <TD width="120" align="right"><b>用 户 名：</b></TD>
      <TD><%=Trim(Request.Cookies("asp163")("UserName"))%> <input name="UserName" type="hidden" value="<%=Trim(Request.Cookies("asp163")("UserName"))%>"></TD>
    </TR>
    <TR class="tdbg" > 
      <TD width="120" align="right"><B>旧 密 码：</B></TD>
      <TD> <INPUT   type=password maxLength=16 size=30 name=OldPassword> </TD>
    </TR>
    <TR class="tdbg" > 
      <TD width="120" align="right"><B>新 密 码：</B></TD>
      <TD> <INPUT   type=password maxLength=16 size=30 name=Password> </TD>
    </TR>
    <TR class="tdbg" > 
      <TD width="120" align="right"><strong>确认密码：</strong></TD>
      <TD><INPUT name=PwdConfirm   type=password id="PwdConfirm" size=30 maxLength=16> 
      </TD>
    </TR>
    <TR align="center" class="tdbg" > 
      <TD height="40" colspan="2"><input name="Action" type="hidden" id="Action" value="Modify"> 
        <input name=Submit   type=submit id="Submit" value=" 保 存 "> </TD>
    </TR>
  </TABLE>
</form>
</body>
</html>
<%
end if
call CloseConn_User()
%>