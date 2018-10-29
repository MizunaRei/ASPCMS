<!--#include file="inc/conn.asp"-->
<!--#include file="inc/config.asp"-->
<!--#include file="inc/function.asp"-->
<!--#include file="inc/md5.asp"-->
<!--#include file="inc/ubbcode.asp"-->
<%
dim Action,FoundErr,ErrMsg
Action=trim(request("Action"))
if EnableLinkReg="No" then
	FoundErr=True
	ErrMsg=ErrMsg & "<br><li>管理员没有开放友情链接申请！</li>"
	call WriteErrMsg()
else
	if Action="Reg" then
		dim LinkType,LinkSiteName,LinkSiteUrl,LinkLogoUrl,LinkSiteAdmin,LinkEmail,LinkSitePassword,LinkSitePwdConfirm,LinkSiteIntro
		LinkType=trim(request("LinkType"))
		LinkSiteName=trim(request("SiteName"))
		LinkSiteUrl=trim(request("SiteUrl"))
		LinkLogoUrl=trim(request("LogoUrl"))
		LinkSiteAdmin=trim(request("SiteAdmin"))
		LinkEmail=trim(request("Email"))
		LinkSitePassword=trim(request("SitePassword"))
		LinkSitePwdConfirm=trim(request("SitePwdConfirm"))
		LInkSiteIntro=trim(request("SiteIntro"))
		if LinkType="" then
			FoundErr=True
			ErrMsg=ErrMsg & "<br><li>链接类型不能为空！</li>"
		else
			LinkType=Cint(LinkType)
			if LinkType=1 and (LinkLogoUrl="" or LinkLogoUrl="http://") then
				FoundErr=True
				ErrMsg=ErrMsg & "<br><li>请输入网站LOGO地址！</li>"
			end if
		end if
		if LinkSiteName="" then
			FoundErr=True
			ErrMsg=ErrMsg & "<br><li>网站名称不能为空！</li>"
		end if
		if LinkSiteUrl="" then
			FoundErr=True
			ErrMsg=ErrMsg & "<br><li>网站地址不能为空！</li>"
		end if
		if LinkSiteAdmin="" then
			FoundErr=True
			ErrMsg=ErrMsg & "<br><li>站长姓名不能为空！</li>"
		end if
		if LinkEmail="" then
			FoundErr=True
			ErrMsg=ErrMsg & "<br><li>Email不能为空！</li>"
		else
			if IsValidEmail(LinkEmail)=false then
				errmsg=errmsg & "<br><li>Email地址错误!</li>"
		   		founderr=true
			end if
		end if
		if LinkSitePassword="" then
			FoundErr=True
			ErrMsg=ErrMsg & "<br><li>网站密码不能为空！</li>"
		end if
		if LinkSitePwdConfirm="" then
			FoundErr=True
			ErrMsg=ErrMsg & "<br><li>确认密码不能为空！</li>"
		end if
		if LinkSitePwdConfirm<>LinkSitePassword then
			FoundErr=True
			ErrMsg=ErrMsg & "<br><li>网站密码与确认密码不一致！</li>"
		end if
		if LinkSiteIntro="" then
			FoundErr=True
			ErrMsg=ErrMsg & "<br><li>网站简介不能为空！</li>"
		end if
		if FoundErr<>True then
			dim sqlLink,rsLink
			sqlLink="select top 1 * from FriendSite where SiteName='" & dvHtmlEncode(LinkSiteName) & "' and SiteUrl='" & dvHtmlEncode(LinkSiteUrl) & "'"
			set rsLink=Server.CreateObject("Adodb.RecordSet")
			rsLink.open sqlLink,conn,1,3
			if not (rsLink.bof and rsLink.eof) then
				FoundErr=True
				ErrMsg=ErrMsg & "<br><li>你申请的网站已经存在！请不要重复申请！</li>"
			else
				rsLink.Addnew
				rsLink("LinkType")=LinkType
				rsLink("SiteName")=dvHtmlEncode(LinkSiteName)
				rsLink("SiteUrl")=dvHtmlEncode(LinkSiteUrl)
				rsLink("LogoUrl")=dvHtmlEncode(LinkLogoUrl)
				rsLink("SiteAdmin")=dvHtmlEncode(LinkSiteAdmin)
				rsLink("Email")=dvHtmlEncode(LinkEmail)
				rsLink("SitePassword")=md5(LinkSitePassword)
				rsLink("SiteIntro")=dvHtmlEncode(LinkSiteIntro)
				rsLink.update
				rsLink.close
				set rsLink=nothing
				call WriteSuccessMsg()
			end if
		end if
		if FoundErr=True then
			call WriteErrMsg()
		end if
	else
%>
<html>
<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href="Admin_style.css" rel="stylesheet">
<title>申请友情链接</title>
<script LANGUAGE="javascript">
<!--
function Check() {
if (document.AddLink.SiteName.value=="")
	{
	  alert("请输入网站名称！")
	  document.AddLink.SiteName.focus()
	  return false
	 }
if (document.AddLink.SiteUrl.value=="")
	{
	  alert("请输入网站地址！")
	  document.AddLink.SiteUrl.focus()
	  return false
	 }
if (document.AddLink.SiteUrl.value=="http://")
	{
	  alert("请输入网站地址！")
	  document.AddLink.SiteUrl.focus()
	  return false
	 }
if (document.AddLink.SiteAdmin.value=="")
	{
	  alert("请输入站长姓名！")
	  document.AddLink.SiteAdmin.focus()
	  return false
	 }
if (document.AddLink.Email.value=="")
	{
	  alert("请输入电子邮件地址！")
	  document.AddLink.Email.focus()
	  return false
	 }
if (document.AddLink.Email.value=="@")
	{
	  alert("请输入电子邮件地址！")
	  document.AddLink.Email.focus()
	  return false
	 }
if (document.AddLink.SitePassword.value=="")
	{
	  alert("请输入网站密码！")
	  document.AddLink.SitePassword.focus()
	  return false
	 }
if (document.AddLink.SitePwdConfirm.value=="")
	{
	  alert("请输入确认密码！")
	  document.AddLink.SitePwdConfirm.focus()
	  return false
	 }
if (document.AddLink.SitePwdConfirm.value!=document.AddLink.SitePassword.value)
	{
	  alert("网站密码与确认密码不一致！")
	  document.AddLink.SitePwdConfirm.focus()
	  document.AddLink.SitePwdConfirm.select()
	  return false
	 }
if (document.AddLink.SiteIntro.value=="")
	{
	  alert("请输入网站简介！")
	  document.AddLink.SiteIntro.focus()
	  return false
	 }
}

//-->
</script>
</head>
<body leftmargin="2" topmargin="0" marginwidth="0" marginheight="0"><br>
<table width="53%" border="0" align="center">
  <tr>
    <td>请您在申请连接后及时做好我们站的连接工作：）<br>我们的连接方式如下：<br>图片地址：http://www.fanchen.com/images/fclogo.gif<br>网站名字：坠落凡尘<br>网站地址：http://www.fanchen.com</td>
  </tr>
</table>
<form method="post" name="AddLink" onsubmit="return Check()" action="FriendSiteReg.asp">
  <table border="0" cellpadding="2" cellspacing="1" align="center" width="400" class="border">
    <tr align="center" class="title"> 
      <td height="22" colspan="2"><strong>申请友情链接</strong></td>
    </tr>
    <tr class="tdbg"> 
      <td width="100" height="25" align="right">链接类型：</td>
      <td height="25"><input name="LinkType" type="radio" value="1" checked>
        Logo链接&nbsp;&nbsp;&nbsp;&nbsp; <input type="radio" name="LinkType" value="2">
        文字链接</td>
    </tr>
    <tr class="tdbg"> 
      <td width="100" height="25" align="right" valign="middle">网站名称：</td>
      <td height="25"> <input name="SiteName" size="30"  maxlength="20" title="这里请输入您的网站名称，最多为20个汉字"> 
        <font color="#FF0000"> *</font></td>
    </tr>
    <tr class="tdbg"> 
      <td width="100" height="25" align="right">网站地址：</td>
      <td height="25"> <input name="SiteUrl" size="30"  maxlength="100" type="text"  value="http://" title="这里请输入您的网站地址，最多为50个字符，前面必须带http://"> 
        <font color="#FF0000">*</font></td>
    </tr>
    <tr class="tdbg"> 
      <td width="100" height="25" align="right">网站Logo：</td>
      <td height="25"> <input name="LogoUrl" size="30"  maxlength="100" type="text"  value="http://" title="这里请输入您的网站LogoUrl地址，最多为50个字符，如果您在第一选项选择的是文字链接，这项就不必填">
        大小必须是88*31 </td>
    </tr>
    <tr class="tdbg"> 
      <td width="100" height="25" align="right">站长姓名：</td>
      <td height="25"> <input name="SiteAdmin" size="30"  maxlength="20" type="text"  title="这里请输入您的大名了，不然我知道您是谁啊。最多为20个字符"> 
        <font color="#FF0000">*</font></td>
    </tr>
    <tr class="tdbg"> 
      <td width="100" height="25" align="right">电子邮件：</td>
      <td height="25"> <input name="Email" size="30"  maxlength="30" type="text"  value="@" title="这里请输入您的联系电子邮件，最多为30个字符"> 
        <font color="#FF0000">*</font></td>
    </tr>
    <tr class="tdbg"> 
      <td width="100" height="25" align="right">网站密码：</td>
      <td height="25"><input name="SitePassword" type="password" id="SitePassword" size="20" maxlength="20"> 
        <font color="#FF0000">*</font> 用于修改信息时用。</td>
    </tr>
    <tr class="tdbg">
      <td height="25" align="right">确认密码：</td>
      <td height="25"><input name="SitePwdConfirm" type="password" id="SitePwdConfirm" size="20" maxlength="20">
        <font color="#FF0000">*</font> </td>
    </tr>
    <tr class="tdbg"> 
      <td width="100" align="right">网站简介：</td>
      <td valign="middle"> <textarea name="SiteIntro" cols="40" rows="5" id="SiteIntro" title="这里请输入您的网站的简单介绍"></textarea></td>
    </tr>
    <tr class="tdbg"> 
      <td height="40" colspan="2" align="center"><input name="Action" type="hidden" id="Action" value="Reg"> 
        <input type="submit" value=" 确 定 " name="cmdOk"> &nbsp; <input type="reset" value=" 重 填 " name="cmdReset"> 
      </td>
    </tr>
  </table>
  </form>
</body>
</html>
<%
	end if
end if
call CloseConn()
sub WriteSuccessMsg()
	response.write "申请友情链接成功！请等待管理员审核通过。"
end sub
%>