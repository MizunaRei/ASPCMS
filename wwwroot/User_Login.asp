<%@language=vbscript codepage=936 %>
<%
option explicit
'强制浏览器重新访问服务器下载页面，而不是从缓存读取页面
Response.Buffer = True 
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1 
Response.Expires = 0 
Response.CacheControl = "no-cache" 
'主要是使随机出现的图片数字随机

dim ComeUrl
ComeUrl=trim(request("ComeUrl"))
if ComeUrl="" then
	ComeUrl=Request.ServerVariables("HTTP_REFERER")
end if
if ComeUrl="" then
	ComeUrl="index.asp"
end if
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
	//'验证码不兼容XP SP3及以后操作系统

	//if (document.Login.CheckCode.value=="")
//	{
//       		alert ("请输入您的验证码！");
//       		document.Login.CheckCode.focus();
//       		return(false);
//       }
}
function CheckBrowser() 
{
  var app=navigator.appName;
  var verStr=navigator.appVersion;
  if (app.indexOf('Netscape') != -1) {
    alert("友情提示：\n    你使用的是Netscape浏览器，可能会导致无法使用后台的部分功能。建议您使用 IE6.0 或以上版本。");
  } 
  else if (app.indexOf('Microsoft') != -1) {
    if (verStr.indexOf("MSIE 3.0")!=-1 || verStr.indexOf("MSIE 4.0") != -1 || verStr.indexOf("MSIE 5.0") != -1 || verStr.indexOf("MSIE 5.1") != -1)
      alert("友情提示：\n    您的浏览器版本太低，可能会导致无法使用后台的部分功能。建议您使用 IE6.0 或以上版本。");
  }
}
</script>
</head>
<body onLoad="SetFocus();">
<p>&nbsp;</p>
<form name="Login" action="User_ChkLogin.asp" method="post" onSubmit="return CheckForm();">
    <table width="300" border="0" align="center" cellpadding="5" cellspacing="0" class="border" >
      <tr class="title"> 
        <td colspan="2" align="center"> <strong>注册用户登录</strong></td>
      </tr>
      
    <tr> 
      <td height="120" colspan="2" class="tdbg">
<table width="250" border="0" cellspacing="8" cellpadding="0" align="center">
          <tr> 
            <td align="right">用户名称：</td>
            <td><input name="UserName"  type="text"  id="UserName" size="23" maxlength="20"></td>
          </tr>
          <tr> 
            <td align="right">用户密码：</td>
            <td><input name="Password"  type="password" id="Password" size="23" maxlength="20"></td>
          </tr>
	  <!--'验证码不兼容XP SP3及以后操作系统-->
      <!--<tr> 
            <td align="right">验 证 码：</td>
            <td><input name="CheckCode" size="6" maxlength="4"><img src="inc/checkcode.asp"></td>
          </tr>-->
          <tr>
		    <td height='25' align='right'>Cookie选项：</td>
			<td height='25'><select name=CookieDate><option selected value=0>不保存</option><option value=1>保存一天</option><option value=2>保存一月</option><option value=3>保存一年</option></select></td>
		  </tr>
		  <tr align="center"> 
            <td colspan="2"> <input name="ComeUrl" type="hidden" id="ComeUrl" value="<%=ComeUrl%>"> 
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
