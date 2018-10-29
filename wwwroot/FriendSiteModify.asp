<!--#include file="inc/conn.asp"-->
<!--#include file="inc/config.asp"-->
<!--#include file="inc/function.asp"-->
<!--#include file="inc/md5.asp"-->
<!--#include file="inc/ubbcode.asp"-->
<%
dim ID,Action,FoundErr,ErrMsg
ID=trim(request("ID"))
Action=trim(request("Action"))
if ID="" then
	call CloseConn()
	response.redirect "FriendSite.asp"
end if
dim sqlLink,rsLink
sqlLink="select * from FriendSite where ID=" & CLng(ID)
set rsLink=Server.CreateObject("Adodb.RecordSet")
rsLink.open sqlLink,conn,1,3
if rsLink.bof and rsLink.eof then
	FoundErr=True
	ErrMsg=ErrMsg & "<br><li>找不到站点！</li>"
else
  if Action="Modify" then
	dim LinkType,LinkSiteName,LinkSiteUrl,LinkLogoUrl,LinkSiteAdmin,LinkEmail,OldSitePassword,LinkSitePassword,LinkSitePwdConfirm,LinkSiteIntro
	LinkType=trim(request("LinkType"))
	LinkSiteName=trim(request("SiteName"))
	LinkSiteUrl=trim(request("SiteUrl"))
	LinkLogoUrl=trim(request("LogoUrl"))
	LinkSiteAdmin=trim(request("SiteAdmin"))
	LinkEmail=trim(request("Email"))
	OldSitePassword=trim(request("OldSitePassword"))
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
			ErrMsg=ErrMsg & "<br><li>网站LOGO不能为空！</li>"
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
	if OldSitePassword="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>旧网站密码不能为空！</li>"
	end if
	if md5(OldSitePassword)<>rsLink("SitePassword") then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>你输入的旧网站密码不对，没有权限修改！</li>"
	end if
	if LinkSitePwdConfirm<>LinkSitePassword then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>新网站密码与确认密码不一致！</li>"
	end if
	if LinkSiteIntro="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>网站简介不能为空！</li>"
	end if
	if LinkSiteIntro="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>网站简介不能为空！</li>"
	end if
	if FoundErr<>True then
			rsLink("LinkType")=LinkType
			rsLink("SiteName")=dvHtmlEncode(LinkSiteName)
			rsLink("SiteUrl")=dvHtmlEncode(LinkSiteUrl)
			rsLink("LogoUrl")=dvHtmlEncode(LinkLogoUrl)
			rsLink("SiteAdmin")=dvHtmlEncode(LinkSiteAdmin)
			rsLink("Email")=dvHtmlEncode(LinkEmail)
			if LinkSitePassword<>"" then
				rsLink("SitePassword")=md5(LinkSitePassword)
			end if
			rsLink("SiteIntro")=dvHtmlEncode(LinkSiteIntro)
			rsLink.update
			rsLink.close
			set rsLink=nothing
			call CloseConn()
			Response.Redirect "FriendSite.asp"
		end if
	end if
end if
if FoundErr=True then
	call WriteErrMsg()
else
%>
<html>
<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href="style.css" rel="stylesheet">
<title>修改友情链接信息</title>
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
if (document.AddLink.OldSitePassword.value=="")
	{
	  alert("请输入旧网站密码！")
	  document.AddLink.OldSitePassword.focus()
	  return false
	 }
if (document.AddLink.SitePwdConfirm.value!=document.AddLink.SitePassword.value)
	{
	  alert("新网站密码与确认密码不一致！")
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
<body leftmargin="2" topmargin="0" marginwidth="0" marginheight="0">
<form method="post" name="AddLink" onsubmit="return Check()" action="FriendSiteModify.asp">
  <table border="0" cellpadding="2" cellspacing="1" align="center" width="400" class="border">
    <tr align="center" class="title"> 
      <td height="22" colspan="2"><strong>修改友情链接信息</strong></td>
    </tr>
    <tr class="tdbg"> 
      <td width="100" height="25" align="right">链接类型：</td>
      <td height="25"><input name="LinkType" type="radio" value="1" <%if rsLink("LinkType")=1 then response.write "checked"%>>
        Logo链接&nbsp;&nbsp;&nbsp;&nbsp; <input type="radio" name="LinkType" value="2" <%if rsLink("LinkType")=2 then response.write "checked"%>>
        文字链接</td>
    </tr>
    <tr class="tdbg"> 
      <td width="100" height="25" align="right" valign="middle">网站名称：</td>
      <td height="25"> <input name="SiteName" title="这里请输入您的网站名称，最多为20个汉字" value="<%=rsLink("SiteName")%>" size="30"  maxlength="20"> 
        <font color="#FF0000"> *</font></td>
    </tr>
    <tr class="tdbg"> 
      <td width="100" height="25" align="right">网站地址：</td>
      <td height="25"> <input name="SiteUrl" size="30"  maxlength="100" type="text"  value="<%=rsLink("SiteUrl")%>" title="这里请输入您的网站地址，最多为50个字符，前面必须带http://"> 
        <font color="#FF0000">*</font></td>
    </tr>
    <tr class="tdbg"> 
      <td width="100" height="25" align="right">网站Logo：</td>
      <td height="25"> <input name="LogoUrl" size="30"  maxlength="100" type="text"  value="<%=rsLink("LogoUrl")%>" title="这里请输入您的网站LogoUrl地址，最多为50个字符，如果您在第一选项选择的是文字链接，这项就不必填"> 
      </td>
    </tr>
    <tr class="tdbg"> 
      <td width="100" height="25" align="right">站长姓名：</td>
      <td height="25"> <input name="SiteAdmin" type="text"  title="这里请输入您的大名了，不然我知道您是谁啊。最多为20个字符" value="<%=rsLink("SiteAdmin")%>" size="30"  maxlength="20"> 
        <font color="#FF0000">*</font></td>
    </tr>
    <tr class="tdbg"> 
      <td width="100" height="25" align="right">电子邮件：</td>
      <td height="25"> <input name="Email" size="30"  maxlength="30" type="text"  value="<%=rsLink("Email")%>" title="这里请输入您的联系电子邮件，最多为30个字符"> 
        <font color="#FF0000">*</font></td>
    </tr>
    <tr class="tdbg"> 
      <td width="100" height="25" align="right">原网站密码：</td>
      <td height="25"><input name="OldSitePassword" type="password" id="OldSitePassword" size="20" maxlength="20"> 
        <font color="#FF0000">* 必须输入</font></td>
    </tr>
    <tr class="tdbg"> 
      <td width="100" height="25" align="right">新网站密码：</td>
      <td height="25"><input name="SitePassword" type="password" id="SitePassword" size="20" maxlength="20"> 
        <font color="#0000FF">若不修改，请保持为空</font></td>
    </tr>
    <tr class="tdbg"> 
      <td height="25" align="right">确认密码：</td>
      <td height="25"><input name="SitePwdConfirm" type="password" id="SitePwdConfirm" size="20" maxlength="20"> 
      </td>
    </tr>
    <tr class="tdbg"> 
      <td width="100" align="right">网站简介：</td>
      <td valign="middle"> <textarea name="SiteIntro" cols="40" rows="5" id="SiteIntro" title="这里请输入您的网站的简单介绍"><%=rsLink("SiteIntro")%></textarea></td>
    </tr>
    <tr class="tdbg"> 
      <td height="40" colspan="2" align="center"><input name="ID" type="hidden" id="ID" value="<%=rsLink("ID")%>"> 
        <input name="Action" type="hidden" id="Action" value="Modify"> <input type="submit" value=" 确 定 " name="cmdOk"> 
      </td>
    </tr>
  </table>
  </form>
</body>
</html>
<%
end if
rsLink.close
set rsLink=nothing
call CloseConn()
%>