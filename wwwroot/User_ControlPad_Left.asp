<!--#include file="inc/Conn_User.asp"-->
<!--#include file="inc/md5.asp"-->
<!--#include file="inc/function.asp"-->
<%
if CheckUserLogined()=False then
	response.Redirect "index.asp"
end if
%>
<html>
<head>
<title></title>
<style>
<!--
a:link {
	text-decoration: none;
	color: #000000;
	font-family: 宋体
}
a:visited {
	text-decoration: none;
	color: #000000;
	font-family: 宋体
}
a:hover {
	text-decoration: underline;
	color: #cc0000
}
body {
	font-family: "宋体";
	font-size: 9pt;
	line-height: 12pt;
}
table {
	font-family: "宋体";
	font-size: 9pt;
	line-height: 12pt;
}
-->
</style>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>
<body bgcolor="#E7DEAD" text="#000000" leftmargin="0" topmargin="0" background="NewImages/bgLightGreen.gif">
<br>
<table width="100%" border="0" cellpadding="0" cellspacing="0" height="100%">
  <tr>
    <td width="10" height="200">&nbsp;</td>
    <td width="8">&nbsp;</td>
    <td valign="top" width="1004"><script language="JavaScript">
function MenuClick(name)
{
if (name.style.display=="none")
{
xiaoshuo.style.display = "none"
zawen.style.display = "none"

xiaoxi.style.display = "none"
shige.style.display = "none"
name.style.display="block";

}
else
name.style.display="none";
}
</script>
      <br>
      <%
		response.write "欢 迎 你：" & UserName
		response.write "<br>您的等级："
		if UserLevel=999 then
			response.write "注册用户"
		elseif UserLevel=99 then
			response.write "收费用户"
		elseif UserLevel=9 then
			response.write "高级用户"
		end if
%>
      <br>
      <img name="Image5" border="0" src="NewImages/Textapps.png" width="28" height="28"><a href="javascript:MenuClick(shige);">投稿管理中心</a><br/>
      <div id="shige" style="display:none;"> &nbsp;&nbsp;&nbsp;<img src="NewImages/iecool_arrow_210.gif" width="20" height="16" border="0"><a href="User_ArticleAdd.asp" target="main"><font color="blue">发表新的文章</font></a><br>
        &nbsp;&nbsp;&nbsp;<img src="NewImages/iecool_arrow_210.gif" width="20" height="16" border="0"><a href="User_ArticleManage.asp" target="main"><font color="blue">查看文章状态</font></a><br>
        &nbsp;&nbsp;&nbsp;<img src="NewImages/iecool_arrow_210.gif" width="20" height="16" border="0"><a href="User_Articlere.asp" target="main"><font color="blue">被退回的文章</font></a></div>
      <br>
      <img name="Image5" border="0" src="NewImages/Favorites.png" width="28" height="28"><a href="javascript:MenuClick(xiaoshuo);">个人资料管理</a><br>
      <div id="xiaoshuo" style="display:none;"> &nbsp;&nbsp;&nbsp;<img src="NewImages/iecool_arrow_210.gif" width="20" height="16" border="0"><a href=User_ModifyPwd.asp target=main><font color="blue">修改登陆密码</font></a><br>
        &nbsp;&nbsp;&nbsp;<img src="NewImages/iecool_arrow_210.gif" width="20" height="16" border="0"><a href=User_ModifyInfo.asp target=main><font color="blue">修改个人资料</font></a></div>
      <br>
      <img name="Image5" border="0" src="NewImages/Mail.png" width="28" height="28"><a href="javascript:MenuClick(xiaoxi);">个人消息管理</a><br>
      <div id="xiaoxi" style="display:none;"> &nbsp;&nbsp;&nbsp;<img src="NewImages/iecool_arrow_210.gif" width="20" height="16" border="0"><a href="sms_main.asp?action=new" target="main"><font color="blue">发送短消息</font></a><br>
        &nbsp;&nbsp;&nbsp;<img src="NewImages/iecool_arrow_210.gif" width="20" height="16" border="0"><a href="sms_user.asp?action=inbox" target="main"><font color="blue">查看短消息</font></a></div>
      <br>
      <img name="Image5" border="0" src="NewImages/Adresses.png" width="28" height="28"><a href="javascript:MenuClick(zawen);">交流中心</a><br>
      <div id="zawen" style="display:none;"> &nbsp;&nbsp;&nbsp;<img src="NewImages/iecool_arrow_210.gif" width="20" height="16" border="0"><a href="guestbook.asp" target="_blank"><font color="#FF0000">给我们留言</font></a><br>        
      </div>
      <br>
      <img name="Image5" border="0" src="NewImages/Classic.png" width="28" height="28"><a href="User_ControlPad_main.asp" target="main">管理首页</a><br>
      <img name="Image5" border="0" src="NewImages/Xpired.png" width="28" height="28"><a href="User_Logout.asp" target="_top">退出登录</a><br>
      </div>
      </div>
      <div align="center"></div></td>
  </tr>
</table>
</body>
</html>
