<%@language=vbscript codepage=936 %>
<!--#include file="inc/conn.asp"-->
<!--#include file="inc/Conn_User.asp"-->
<!--#include file="inc/function.asp"-->
<!--#include file="inc/md5.asp"-->
<!--#include file="inc/config.asp"-->
<%
dim Action,FoundErr,ErrMsg
dim rs,sql
Action=trim(request("Action"))
if Action="Check" then
	call CheckUser()
else
	call main()
end if
if FoundErr=true then
	call WriteErrMsg()
end if
call CloseConn()

sub main()
%>
<html>
<head>
<title>注册用户登录</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="style.CSS">
<script language=javascript>
function SetFocus()
{
if (document.Login.UserName.value=="")
	document.Login.UserName.focus();
else
	document.Login.UserName.select();
}
function CheckForm()
{
	if(document.Login.UserName.value=="")
	{
		alert("请输入用户名！");
		document.Login.UserName.focus();
		return false;
	}
	if(document.Login.Password.value == "")
	{
		alert("请输入密码！");
		document.Login.Password.focus();
		return false;
	}
	if(document.Login.CheckNum.value == "")
	{
		alert("请输入验证码！");
		document.Login.CheckNum.focus();
		return false;
	}
}
</script>
</head>
<body onLoad="SetFocus();">
<p>&nbsp;</p>
<form name="Login" action="User_RegCheck.asp" method="post" onSubmit="return CheckForm();">
    <table width="400" border="0" align="center" cellpadding="5" cellspacing="0" class="border" >
      <tr class="title"> 
        
      <td colspan="2" align="center"> <strong>注册用户认证</strong></td>
      </tr>
      
    <tr> 
      <td height="120" colspan="2" class="tdbg">请输入您注册时填写的用户名和密码，以及本站发给你的确认信中的随机验证码。必须完全正确后，你的帐户才会激活。 
        <table width="250" border="0" cellspacing="8" cellpadding="0" align="center">
          <tr> 
            <td align="right">用户名称：</td>
            <td><input name="UserName"  type="text"  id="UserName" size="23" maxlength="20"></td>
          </tr>
          <tr> 
            <td align="right">用户密码：</td>
            <td><input name="Password"  type="password" id="Password" size="23" maxlength="20"></td>
          </tr>
          <tr>
		    <td height='25' align='right'>随机验证码：</td>
			<td height='25'><input name="CheckNum" type="text" id="CheckNum" size="23" maxlength="6"></td>
		  </tr>
		  <tr align="center"> 
            <td colspan="2"> <input name="Action" type="hidden" id="Action" value="Check"> 
              <input   type="submit" name="Submit" value=" 确认 "> &nbsp; <input name="reset" type="reset"  id="reset" value=" 清除 "> 
            </td>
          </tr>
        </table>
        </td>
      </tr>
    </table>
</form>
</body>
</html>
<%
end sub

sub CheckUser()
	dim password,CheckNum
	username=replace(trim(request("username")),"'","")
	password=replace(trim(Request("password")),"'","")
	CheckNum=replace(trim(Request("CheckNum")),"'","")

	if UserName="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>用户名不能为空！</li>"
	end if
	if Password="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>密码不能为空！</li>"
	end if
	if CheckNum="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>密码不能为空！</li>"
	end if

	if FoundErr=True then
		exit sub
	end if
	
	password=md5(password)
	set rs=server.createobject("adodb.recordset")
	sql="select * from " & db_User_Table & " where " & db_User_LockUser & "=False and " & db_User_Name & "='" & username & "' and " & db_User_Password & "='" & password &"'"
	rs.open sql,Conn_User,1,3
	if rs.bof and rs.eof then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>用户名或密码错误！！！</li>"
	else
		if password<>rs(db_User_Password) then
			FoundErr=True
			ErrMsg=ErrMsg & "<br><li>用户名或密码错误！！！</li>"
		else
			if AdminCheckReg="Yes" then
				rs(db_User_UserLevel)=2000
				rs.update
				call WriteSuccessMsg("恭喜你通过了Email验证。请等待管理开通你的帐号。开通后，你就正式正为本站的一员了。")
			else
				rs(db_User_UserLevel)=999
				rs.update
				call WriteSuccessMsg("恭喜你正式成为本站的一员。")
				call SaveCookie_asp163()
			end if
		end if
	end if
	rs.close
	set rs=nothing
	
end sub

sub SaveCookie_asp163()
	Response.Cookies("asp163")("UserName")=rs(db_User_Name)
	Response.Cookies("asp163")("Password") = rs(db_User_Password)
	Response.Cookies("asp163")("UserLevel")=rs(db_User_UserLevel)
	Response.Cookies("asp163")("CookieDate") = 0
end sub

%>