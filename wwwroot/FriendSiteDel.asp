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
  	if Action="Del" then
		dim OldSitePassword
		OldSitePassword=trim(request("OldSitePassword"))
		if OldSitePassword="" then
			FoundErr=True
			ErrMsg=ErrMsg & "<br><li>旧网站密码不能为空！</li>"
		end if
		if md5(OldSitePassword)<>rsLink("SitePassword") then
			FoundErr=True
			ErrMsg=ErrMsg & "<br><li>你输入的旧网站密码不对，没有权限删除！</li>"
		end if
		if FoundErr<>True then
			rsLink.delete
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
<title>删除友情链接信息</title>
<script LANGUAGE="javascript">
<!--
function Check() {
if (document.AddLink.OldSitePassword.value=="")
	{
	  alert("请输入旧网站密码！")
	  document.AddLink.OldSitePassword.focus()
	  return false
	 }
}

//-->
</script>
</head>
<body leftmargin="2" topmargin="0" marginwidth="0" marginheight="0">
<form method="post" name="AddLink" onsubmit="return Check()" action="FriendSiteDel.asp">
  <table border="0" cellpadding="2" cellspacing="1" align="center" width="400" class="border">
    <tr align="center" class="title"> 
      <td height="22" colspan="2"><strong>删除友情链接信息</strong></td>
    </tr>
    <tr class="tdbg"> 
      <td width="100" height="25" align="right">链接类型：</td>
      <td height="25">
	  <%
	  if rsLink("LinkType")=1 then
	  	response.write "Logo链接"
	  else
	  	response.write "文字链接"
	  end if
	  %></td>
    </tr>
    <tr class="tdbg"> 
      <td width="100" height="25" align="right" valign="middle">网站名称：</td>
      <td height="25"><%=rsLink("SiteName")%></td>
    </tr>
    <tr class="tdbg"> 
      <td width="100" height="25" align="right">网站地址：</td>
      <td height="25"><%=rsLink("SiteUrl")%></td>
    </tr>
    <tr class="tdbg"> 
      <td width="100" height="25" align="right">网站Logo：</td>
      <td height="25"><%=rsLink("LogoUrl")%></td>
    </tr>
    <tr class="tdbg"> 
      <td width="100" height="25" align="right">站长姓名：</td>
      <td height="25"><%=rsLink("SiteAdmin")%></td>
    </tr>
    <tr class="tdbg"> 
      <td width="100" height="25" align="right">电子邮件：</td>
      <td height="25"><%=rsLink("Email")%></td>
    </tr>
    <tr class="tdbg"> 
      <td width="100" align="right">网站简介：</td>
      <td valign="middle"><%=rsLink("SiteIntro")%></td>
    </tr>
    <tr class="tdbg"> 
      <td width="100" height="25" align="right">原网站密码：</td>
      <td height="25"><input name="OldSitePassword" type="password" id="OldSitePassword" size="20" maxlength="20"> 
        <font color="#FF0000">* 必须输入</font></td>
    </tr>
    <tr class="tdbg"> 
      <td height="40" colspan="2" align="center"><input name="ID" type="hidden" id="ID" value="<%=rsLink("ID")%>"> 
        <input name="Action" type="hidden" id="Action" value="Del"> <input type="submit" value=" 确 定 " name="cmdOk"> 
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