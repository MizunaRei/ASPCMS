<%@language=vbscript codepage=936 %>
<%
option explicit
response.buffer=true	
Const PurviewLevel=1
'response.write "此功能被WEBBOY暂时禁止了！"
'response.end
%>
<!--#include file="inc/conn.asp"-->
<!--#include file="inc/config.asp"-->
<!--#include file="Admin_ChkPurview.asp"-->
<!--#include file="inc/function.asp"-->
<%
dim ObjInstalled,Action,FoundErr,ErrMsg
ObjInstalled=IsObjInstalled("Scripting.FileSystemObject")
Action=trim(request("Action"))
if Action="" then
	Action="ShowInfo"
end if
%>
<html>
<head>
<title>专题管理</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="Admin_Style.css" rel="stylesheet" type="text/css">
</head>
<body leftmargin="2" topmargin="0" marginwidth="0" marginheight="0">
  <table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" Class="border">
    <tr class="topbg"> 
      
    <td height="22" colspan=2 align=center><b>网 站 配 置</b></td>
    </tr>
    <tr class="tdbg"> 
      <td width="70" height="30"><strong>管理导航：</strong></td>
      
    <td height="30"><a href="Admin_SiteConfig.asp">网站信息配置</a> | <a href="Admin_SiteConfig.asp">网站选项配置</a> 
      | <a href="#Email">邮件服务器选项</a> | <a href="#UpFile">上传文件选项</a></td>
    </tr>
  </table>
<%
if Action="SaveConfig" then
	call SaveConfig()
else
	call ShowConfig()
end if
if FoundErr=True then
	call WriteErrMsg()
end if
call CloseConn()

sub ShowConfig()
%>
<form method="POST" action="Admin_SiteConfig.asp" id="form1" name="form1">
  <table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" Class="border">
    <tr> 
      <td height="22" colspan="2" class="topbg"> <a name="SiteInfo"></a><strong>网站信息配置</strong></td>
    </tr>
    <tr> 
      <td width="400" height="25" class="tdbg"><strong>网站名称：</strong></td>
      <td width="368" height="25" class="tdbg"> 
        <input name="SiteName" type="text" id="SiteName" value="<%=SiteName%>" size="40" maxlength="50">
      </td>
    </tr>
    <tr> 
      <td width="400" height="25" class="tdbg"><strong>网站标题：</strong></td>
      <td width="368" height="25" class="tdbg"> 
        <input name="SiteTitle" type="text" id="SiteTitle" value="<%=SiteTitle%>" size="40" maxlength="50">
      </td>
    </tr>
    <tr> 
      <td width="400" height="25" class="tdbg"><strong>网站地址：</strong><br>
        请添写完整URL地址</td>
      <td width="368" height="25" class="tdbg"> 
        <input name="SiteUrl" type="text" id="SiteUrl" value="<%=SiteUrl%>" size="40" maxlength="255">
      </td>
    </tr>
    <tr> 
      <td width="400" height="25" class="tdbg"><strong>LOGO地址：</strong><br>
        请添写完整URL地址</td>
      <td width="368" height="25" class="tdbg"> 
        <input name="LogoUrl" type="text" id="LogoUrl" value="<%=LogoUrl%>" size="40" maxlength="255">
      </td>
    </tr>
    <tr> 
      <td width="400" height="25" class="tdbg"><strong>Banner地址：</strong><br>
        请添写完整URL地址。若为空，则在Banner处显示广告管理中预设的广告。</td>
      <td width="368" height="25" class="tdbg"> 
        <input name="BannerUrl" type="text" id="BannerUrl" value="<%=BannerUrl%>" size="40" maxlength="255">
      </td>
    </tr>
    <tr> 
      <td width="400" height="25" class="tdbg"><strong>站长姓名：</strong></td>
      <td width="368" height="25" class="tdbg"> 
        <input name="WebmasterName" type="text" id="WebmasterName" value="<%=WebmasterName%>" size="40" maxlength="20">
      </td>
    </tr>
    <tr> 
      <td width="400" height="25" class="tdbg"><strong>站长信箱：</strong></td>
      <td width="368" height="25" class="tdbg"> 
        <input name="WebmasterEmail" type="text" id="WebmasterEmail" value="<%=WebmasterEmail%>" size="40" maxlength="100">
      </td>
    </tr>
    <tr> 
      <td width="400" height="25" class="tdbg"><strong>版权信息：</strong><br>
        支持HTML标记，不能使用双引号</td>
      <td width="368" height="25" class="tdbg"> 
        <textarea name="Copyright" cols="32" rows="4" id="Copyright"><%=Copyright%></textarea>
      </td>
    </tr>
    <tr> 
      <td height="25" colspan="2" class="topbg"><a name="SiteOption"></a><strong>网站选项配置</strong></td>
    </tr>
    <tr> 
      <td width="400" height="25" class="tdbg"><strong>后台是否显示右键菜单：</strong></td>
      <td height="25" class="tdbg"> 
        <input type="radio" name="ShowPopMenu" value="Yes" <%if ShowPopMenu="Yes" then response.write "checked"%>>
        是 &nbsp;&nbsp;&nbsp;&nbsp; 
        <input type="radio" name="ShowPopMenu" value="No" <%if ShowPopMenu="No" then response.write "checked"%>>
        否</td>
    </tr>
    <tr> 
      <td width="400" height="25" class="tdbg"><strong>是否显示网站频道：</strong></td>
      <td height="25" class="tdbg"> 
        <input type="radio" name="ShowSiteChannel" value="Yes" <%if ShowSiteChannel="Yes" then response.write "checked"%>>
        是 &nbsp;&nbsp;&nbsp;&nbsp; 
        <input type="radio" name="ShowSiteChannel" value="No" <%if ShowSiteChannel="No" then response.write "checked"%>>
        否</td>
    </tr>
    <tr> 
      <td width="400" height="25" class="tdbg"><strong>是否显示自选风格：</strong></td>
      <td height="25" class="tdbg"> 
        <input type="radio" name="ShowMyStyle" value="Yes" <%if ShowMyStyle="Yes" then response.write "checked"%>>
        是 &nbsp;&nbsp;&nbsp;&nbsp; 
        <input type="radio" name="ShowMyStyle" value="No" <%if ShowMyStyle="No" then response.write "checked"%>>
        否</td>
    </tr>
    <tr> 
      <td width="400" height="25" class="tdbg"><strong>是否使用树状导航菜单：</strong></td>
      <td height="25" class="tdbg"> 
        <input type="radio" name="ShowClassTreeGuide" value="Yes" <%if ShowClassTreeGuide="Yes" then response.write "checked"%>>
        是 &nbsp;&nbsp;&nbsp;&nbsp; 
        <input type="radio" name="ShowClassTreeGuide" value="No" <%if ShowClassTreeGuide="No" then response.write "checked"%>>
        否</td>
    </tr>
    <tr> 
      <td width="400" height="25" class="tdbg"><strong>是否启用文章审核功能：</strong><br>
        如禁用文章审核功能，则新录入的文章将直接发表（通过审核）</td>
      <td height="25" class="tdbg"> 
        <input type="radio" name="EnableArticleCheck" value="Yes" <%if EnableArticleCheck="Yes" then response.write "checked"%>>
        是 &nbsp;&nbsp;&nbsp;&nbsp; 
        <input type="radio" name="EnableArticleCheck" value="No" <%if EnableArticleCheck="No" then response.write "checked"%>>
        否</td>
    </tr>
    <tr> 
      <td width="400" height="25" class="tdbg"><strong>是否启用留言审核功能：</strong><br>
        如禁用留言审核功能，则新发表的留言将直接显示（通过审核）</td>
      <td height="25" class="tdbg"> 
        <input type="radio" name="EnableGuestCheck" value="Yes" <%if EnableGuestCheck="Yes" then response.write "checked"%>>
        是 &nbsp;&nbsp;&nbsp;&nbsp; 
        <input type="radio" name="EnableGuestCheck" value="No" <%if EnableGuestCheck="No" then response.write "checked"%>>
        否</td>
    </tr>
    <tr> 
      <td width="400" height="25" class="tdbg"><strong>是否开放文件上传功能：</strong><br>
        添加/修改文章时是否可以上传文件</td>
      <td height="25" class="tdbg"> 
        <input type="radio" name="EnableUploadFile" value="Yes" <%if EnableUploadFile="Yes" then response.write "checked"%>>
        是 &nbsp;&nbsp;&nbsp;&nbsp; 
        <input type="radio" name="EnableUploadFile" value="No" <%if EnableUploadFile="No" then response.write "checked"%>>
        否</td>
    </tr>
    <tr>
      <td width="400" height="25" class="tdbg"><strong>是否保存远程图片到本地：
        </strong><br>
        如果从其它网站上复制的内容中包含图片，则将图片复制到本站服务器上</td>
      <td height="25" class="tdbg">
        <input type="radio" name="EnableSaveRemote" value="Yes" <%if EnableSaveRemote="Yes" then response.write "checked"%>>
        是 &nbsp;&nbsp;&nbsp;&nbsp; 
        <input type="radio" name="EnableSaveRemote" value="No" <%if EnableSaveRemote="No" then response.write "checked"%>>
        否</td>
    </tr>
    <tr> 
      <td width="400" height="25" class="tdbg"><strong>是否开放友情链接申请：</strong></td>
      <td height="25" class="tdbg"> 
        <input type="radio" name="EnableLinkReg" value="Yes" <%if EnableLinkReg="Yes" then response.write "checked"%>>
        是 &nbsp;&nbsp;&nbsp;&nbsp; 
        <input type="radio" name="EnableLinkReg" value="No" <%if EnableLinkReg="No" then response.write "checked"%>>
        否</td>
    </tr>
    <tr> 
      <td width="400" height="25" class="tdbg"><strong>是否弹出公告窗口：</strong></td>
      <td height="25" class="tdbg"> 
        <input type="radio" name="PopAnnounce" value="Yes" <%if PopAnnounce="Yes" then response.write "checked"%>>
        是 &nbsp;&nbsp;&nbsp;&nbsp; 
        <input name="PopAnnounce" type="radio" value="No" <%if PopAnnounce="No" then response.write "checked"%>>
        否</td>
    </tr>
    <tr> 
      <td width="400" height="25" class="tdbg"><strong>热点文章的点击数：</strong><br>
        文章点击数达到此数时就会被设为热点文章</td>
      <td height="25" class="tdbg"> 
        <input name="HitsOfHot" type="text" id="HitsOfHot" value="<%=HitsOfHot%>" size="6" maxlength="5">
      </td>
    </tr>
    <tr> 
      <td width="400" height="25" class="tdbg"><strong>Session会话的保持时间：</strong><br>
        主要用于后台管理员登录，为了安全，请不要将时间设得太长。建议设为10分钟。</td>
      <td height="25" class="tdbg"> 
        <input name="SessionTimeout" type="text" id="SessionTimeout" value="<%=SessionTimeout%>" size="6" maxlength="5">
        分钟</td>
    </tr>
    <tr> 
      <td width="400" height="25" class="tdbg"><strong>发表评论权限：</strong><br>
        只有具有相应权限的人才能发表评论。</td>
      <td height="25" class="tdbg"> 
        <select name="CommentPurview" id="CommentPurview">
          <option value="9999" <%if CommentPurview=9999 then response.write " selected"%>>游客</option>
          <option value="999" <%if CommentPurview=999 then response.write " selected"%>>注册用户</option>
          <option value="99" <%if CommentPurview=99 then response.write " selected"%>>收费用户</option>
          <option value="9" <%if CommentPurview=9 then response.write " selected"%>>VIP用户</option>
          <option value="5" <%if CommentPurview=5 then response.write " selected"%>>管理员</option>
        </select>
      </td>
    </tr>
    <tr> 
      <td height="25" colspan="2" class="topbg"><a name="Email"></a><strong>用户选项</strong></td>
    </tr>
    <tr> 
      <td width="400" height="25" class="tdbg"><strong>是否允许新用户注册：</strong></td>
      <td height="25" class="tdbg"> 
        <input type="radio" name="EnableUserReg" value="Yes" <%if EnableUserReg="Yes" then response.write "checked"%>>
        是 &nbsp;&nbsp;&nbsp;&nbsp; 
        <input type="radio" name="EnableUserReg" value="No" <%if EnableUserReg="No" then response.write "checked"%>>
        否</td>
    </tr>
    <tr> 
      <td width="400" height="25" class="tdbg"><strong>新用户注册是否需要邮件验证：</strong><br>
        若选择“是”，则用户注册后系统会发一封带有验证码的邮件给此用户，用户必须在通过邮件验证后才能真正成为正式注册用户</td>
      <td height="25" class="tdbg"> 
        <input type="radio" name="EmailCheckReg" value="Yes" <%if EmailCheckReg="Yes" then response.write "checked"%>>
        是 &nbsp;&nbsp;&nbsp;&nbsp; 
        <input type="radio" name="EmailCheckReg" value="No" <%if EmailCheckReg="No" then response.write "checked"%>>
        否</td>
    </tr>
    <tr> 
      <td width="400" height="25" class="tdbg"> 
        <p><strong>新用户注册是否需要管理员认证：</strong><br>
          若选择是，则用户必须在通过管理员认证后才能真正成功正式注册用户。</p>
      </td>
      <td height="25" class="tdbg"> 
        <input type="radio" name="AdminCheckReg" value="Yes" <%if AdminCheckReg="Yes" then response.write "checked"%>>
        是 &nbsp;&nbsp;&nbsp;&nbsp; 
        <input type="radio" name="AdminCheckReg" value="No" <%if AdminCheckReg="No" then response.write "checked"%>>
        否</td>
    </tr>
    <TR bgcolor="#EAEEFB" > 
      <TD width="400"><strong>注册用户的默认计费方式：</strong></TD>
      <TD> 
        <input name="ChargeType_999" type="radio" value="1" <%if ChargeType_999=1 then response.write " checked"%>>
        扣点数<font color="#0000FF">（推荐）</font>：&nbsp;每阅读一篇收费文章，扣除相应点数。&nbsp;<br>
        <input type="radio" name="ChargeType_999" value="2" <%if ChargeType_999=2 then response.write " checked"%>>
        有效期：在有效期内，用户可以任意阅读收费内容</TD>
    </TR>
    <tr bgcolor="#EAEEFB"> 
      <td width="400" height="25"><strong>注册用户的默认可用点数：</strong><br>
        新用户注册成功后默认可以得到的点数，用于注册用户试用一些收费项目</td>
      <td height="25"> 
        <input name="UserPoint_999" type="text" id="UserPoint_999" value="<%=UserPoint_999%>" size="6" maxlength="5">
        点 </td>
    </tr>
    <tr bgcolor="#EAEEFB"> 
      <td width="400" height="25"><strong>注册用户的默认有效期：</strong><br>
        新用户注册成功后默认有效期，用于注册用户试用一些收费项目<br>
        有效期从注册之日开始计算</td>
      <td height="25"> 
        <input name="ValidDays_999" type="text" id="ValidDays_999" value="<%=ValidDays_999%>" size="6" maxlength="5">
        天</td>
    </tr>
    <TR bgcolor="#DFEFFF" > 
      <TD width="400"><strong>收费用户的默认计费方式：</strong></TD>
      <TD> 
        <input name="ChargeType_99" type="radio" value="1" <%if ChargeType_99=1 then response.write " checked"%>>
        扣点数<font color="#0000FF">（推荐）</font>：&nbsp;每阅读一篇收费文章，扣除相应点数。&nbsp;<br>
        <input type="radio" name="ChargeType_99" value="2" <%if ChargeType_99=2 then response.write " checked"%>>
        有效期：在有效期内，用户可以任意阅读收费内容</TD>
    </TR>
    <tr bgcolor="#DFEFFF"> 
      <td width="400" height="25"><strong>收费用户的默认可用点数：</strong><br>
        将用户设为收费用户时，这些用户的默认可用点数</td>
      <td height="25"> 
        <input name="UserPoint_99" type="text" id="UserPoint_99" value="<%=UserPoint_99%>" size="6" maxlength="5">
        点 </td>
    </tr>
    <tr bgcolor="#DFEFFF"> 
      <td width="400" height="25"><strong>收费用户的默认有效期：</strong><br>
        将用户设为收费用户时，这些用户的默认有效期<br>
        有效期从设为收费用户之日开始计算</td>
      <td height="25"> 
        <input name="ValidDays_99" type="text" id="ValidDays_99" value="<%=ValidDays_99%>" size="6" maxlength="5">
        天</td>
    </tr>
    <TR bgcolor="#E6E6FF" > 
      <TD width="400"><strong>VIP用户的默认计费方式：</strong></TD>
      <TD> 
        <input name="ChargeType_9" type="radio" value="1" <%if ChargeType_99=1 then response.write " checked"%>>
        扣点数<font color="#0000FF">（推荐）</font>：&nbsp;每阅读一篇收费文章，扣除相应点数。&nbsp;<br>
        <input type="radio" name="ChargeType_9" value="2" <%if ChargeType_99=2 then response.write " checked"%>>
        有效期：在有效期内，用户可以任意阅读收费内容</TD>
    </TR>
    <tr bgcolor="#E6E6FF"> 
      <td width="400" height="25"><strong>VIP用户的默认可用点数：</strong><br>
        将用户设为VIP用户时，这些用户的默认可用点数</td>
      <td height="25"> 
        <input name="UserPoint_9" type="text" id="UserPoint_9" value="<%=UserPoint_9%>" size="6" maxlength="5">
        点 </td>
    </tr>
    <tr bgcolor="#E6E6FF"> 
      <td width="400" height="25"><strong>VIP用户的默认有效期：</strong><br>
        将用户设为VIP用户时，这些用户的默认有效期<br>
        有效期从设为VIP用户之日开始计算</td>
      <td height="25"> 
        <input name="ValidDays_9" type="text" id="ValidDays_9" value="<%=ValidDays_9%>" size="6" maxlength="5">
        天</td>
    </tr>
    <tr> 
      <td height="25" colspan="2" class="topbg"><a name="Email"></a><strong>邮件服务器选项</strong></td>
    </tr>
    <tr> 
      <td width="400" height="25" class="tdbg"><strong>邮件发送组件：</strong><br>
        请一定要选择服务器上已安装的组件</td>
      <td height="25" class="tdbg"> 
        <select name="MailObject" id="MailObject">
          <option value="Jmail" selected>Jmail</option>
        </select>
      </td>
    </tr>
    <tr> 
      <td width="400" height="25" class="tdbg"><strong>SMTP服务器地址：</strong><br>
        用来发送邮件的SMTP服务器<br>
        如果你不清楚此参数含义，请联系你的空间商 </td>
      <td height="25" class="tdbg"> 
        <input name="MailServer" type="text" id="MailServer" value="<%=MailServer%>" size="40">
      </td>
    </tr>
    <tr> 
      <td width="400" height="25" class="tdbg"><strong>SMTP登录用户名：</strong><br>
        当你的服务器需要SMTP身份验证时还需设置此参数</td>
      <td height="25" class="tdbg"> 
        <input name="MailServerUserName" type="text" id="MailServerUserName" value="<%=MailServerUserName%>" size="40">
      </td>
    </tr>
    <tr> 
      <td width="400" height="25" class="tdbg"><strong>SMTP登录密码：</strong><br>
        当你的服务器需要SMTP身份验证时还需设置此参数 </td>
      <td height="25" class="tdbg"> 
        <input name="MailServerPassWord" type="text" id="MailServerPassWord" value="<%=MailServerPassWord%>" size="40">
      </td>
    </tr>
    <tr> 
      <td width="400" height="25" class="tdbg"><strong>SMTP域名</strong>：<br>
        如果用“name@domain.com”这样的用户名登录时，请指明domain.com</td>
      <td height="25" class="tdbg"> 
        <input name="MailDomain" type="text" id="MailDomain" value="<%=MailDomain%>" size="40">
      </td>
    </tr>
    <tr> 
      <td height="25" colspan="2" class="topbg"><a name="UpFile" id="UpFile"></a><strong>上传文件选项</strong></td>
    </tr>
    <tr> 
      <td width="400" height="25" class="tdbg"><strong>上传文件大小限制：</strong><br>
        建议不要超过1024K，以免影响服务器性能</td>
      <td height="25" class="tdbg"> 
        <input name="MaxFileSize" type="text" id="MaxFileSize" value="<%=MaxFileSize%>" size="6" maxlength="5">
        K</td>
    </tr>
    <tr> 
      <td width="400" height="25" class="tdbg"><strong>存放上传文件的目录：</strong><br>
        请输入相对于首页（Default.asp）的相对路径</td>
      <td height="25" class="tdbg"> 
        <input name="SaveUpFilesPath" type="text" id="SaveUpFilesPath" value="<%=SaveUpFilesPath%>" size="30" maxlength="100">
      </td>
    </tr>
    <tr> 
      <td width="400" height="25" class="tdbg"><strong>允许的上传文件类型：</strong><br>
        只输入扩展名。每种文件类型用“|”号分开。</td>
      <td height="25" class="tdbg"> 
        <input name="UpFileType" type="text" id="UpFileType" value="<%=UpFileType%>" size="50" maxlength="255">
      </td>
    </tr>
    <tr> 
      <td width="400" height="25" class="tdbg"><strong>删除文章时是否同时删除文章中的上传文件：</strong><br>
        此功能需要FSO支持。</td>
      <td height="25" class="tdbg"> 
        <input type="radio" name="DelUpFiles" value="Yes" <%if DelUpFiles="Yes" then response.write "checked"%>>
        是 &nbsp;&nbsp;&nbsp;&nbsp; 
        <input type="radio" name="DelUpFiles" value="No" <%if DelUpFiles="No" then response.write "checked"%>>
        否</td>
    </tr>
    <tr> 
      <td height="40" colspan="2" align="center" class="tdbg"> 
        <input name="Action" type="hidden" id="Action" value="SaveConfig">
        <input name="cmdSave" type="submit" id="cmdSave" value=" 保存设置 " <% If ObjInstalled=false Then response.write "disabled" %>>
      </td>
    </tr>
    <%
If ObjInstalled=false Then
	Response.Write "<tr class='tdbg'><td height='40' colspan='3'><b><font color=red>你的服务器不支持 FSO(Scripting.FileSystemObject)! 不能使用本功能。<br>请直接修改“Inc/config.asp”文件中的内容。</font></b></td></tr>"
End If
%>
  </table>
<%
end sub
%>
</form>
</body>
</html>
<%
sub SaveConfig()
	If ObjInstalled=false Then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>你的服务器不支持 FSO(Scripting.FileSystemObject)! </li>"
		exit sub
	end if
	dim fso,hf
	set fso=Server.CreateObject("Scripting.FileSystemObject")
	set hf=fso.CreateTextFile(Server.mappath("inc/config.asp"),true)
	hf.write "<" & "%" & vbcrlf
	hf.write "Const SiteName=" & chr(34) & trim(request("SiteName")) & chr(34) & "        '网站名称" & vbcrlf
	hf.write "Const SiteTitle=" & chr(34) & trim(request("SiteTitle")) & chr(34) & "        '网站标题" & vbcrlf
	hf.write "Const SiteUrl=" & chr(34) & trim(request("SiteUrl")) & chr(34) & "        '网站地址" & vbcrlf
	hf.write "Const LogoUrl=" & chr(34) & trim(request("LogoUrl")) & chr(34) & "        'Logo地址" & vbcrlf
	hf.write "Const BannerUrl=" & chr(34) & trim(request("BannerUrl")) & chr(34) & "        'Banner地址" & vbcrlf
	hf.write "Const WebmasterName=" & chr(34) & trim(request("WebmasterName")) & chr(34) & "        '站长姓名" & vbcrlf
	hf.write "Const WebmasterEmail=" & chr(34) & trim(request("WebmasterEmail")) & chr(34) & "        '站长信箱" & vbcrlf
	hf.write "Const Copyright=" & chr(34) & trim(request("Copyright")) & chr(34) & "        '版权信息" & vbcrlf
	hf.write "Const ShowPopMenu=" & chr(34) & trim(request("ShowPopMenu")) & chr(34) & "        '后台是否显示右键菜单" & vbcrlf
	hf.write "Const ShowSiteChannel=" & chr(34) & trim(request("ShowSiteChannel")) & chr(34) & "        '是否显示网站频道" & vbcrlf
	hf.write "Const ShowMyStyle=" & chr(34) & trim(request("ShowMyStyle")) & chr(34) & "        '是否显示自选风格" & vbcrlf  
	hf.write "Const ShowClassTreeGuide=" & chr(34) & trim(request("ShowClassTreeGuide")) & chr(34) & "        '是否使用树状导航菜单" & vbcrlf  
	hf.write "Const EnableArticleCheck=" & chr(34) & trim(request("EnableArticleCheck")) & chr(34) & "        '是否启用文章审核功能" & vbcrlf
	hf.write "Const EnableGuestCheck=" & chr(34) & trim(request("EnableGuestCheck")) & chr(34) & "        '是否启用留言审核功能" & vbcrlf
	hf.write "Const EnableUploadFile=" & chr(34) & trim(request("EnableUploadFile")) & chr(34) & "        '是否开放文件上传" & vbcrlf
	hf.write "Const EnableSaveRemote=" & chr(34) & trim(request("EnableSaveRemote")) & chr(34) & "        '是否保存远程图片到本地" & vbcrlf
	hf.write "Const EnableUserReg=" & chr(34) & trim(request("EnableUserReg")) & chr(34) & "        '是否允许新用户注册" & vbcrlf
	hf.write "Const EmailCheckReg=" & chr(34) & trim(request("EmailCheckReg")) & chr(34) & "        '新用户注册是否需要邮件验证" & vbcrlf
	hf.write "Const AdminCheckReg=" & chr(34) & trim(request("AdminCheckReg")) & chr(34) & "        '新用户注册是否需要管理员认证" & vbcrlf
	hf.write "Const EnableLinkReg=" & chr(34) & trim(request("EnableLinkReg")) & chr(34) & "        '是否开放友情链接申请" & vbcrlf
	hf.write "Const PopAnnounce=" & chr(34) & trim(request("PopAnnounce")) & chr(34) & "        '是否弹出公告窗口" & vbcrlf
	hf.write "Const HitsOfHot=" & trim(request("HitsOfHot")) & "        '热门文章点击数" & vbcrlf
	hf.write "Const SessionTimeout=" & trim(request("SessionTimeout")) & "        'Session会话的保持时间" & vbcrlf
	hf.write "Const CommentPurview=" & trim(request("CommentPurview")) & "        '发表评论的权限" & vbcrlf
	hf.write "Const MailObject=" & chr(34) & trim(request("MailObject")) & chr(34) & "        '邮件发送组件" & vbcrlf
	hf.write "Const MailServer=" & chr(34) & trim(request("MailServer")) & chr(34) & "        '用来发送邮件的SMTP服务器" & vbcrlf
	hf.write "Const MailServerUserName=" & chr(34) & trim(request("MailServerUserName")) & chr(34) & "        '登录用户名" & vbcrlf
	hf.write "Const MailServerPassWord=" & chr(34) & trim(request("MailServerPassWord")) & chr(34) & "        '登录密码" & vbcrlf
	hf.write "Const MailDomain=" & chr(34) & trim(request("MailDomain")) & chr(34) & "        '域名" & vbcrlf
	hf.write "Const MaxFileSize=" & trim(request("MaxFileSize")) & "        '上传文件大小限制" & vbcrlf
	hf.write "Const SaveUpFilesPath=" & chr(34) & trim(request("SaveUpFilesPath")) & chr(34) & "        '存放上传文件的目录" & vbcrlf
	hf.write "Const UpFileType=" & chr(34) & trim(request("UpFileType")) & chr(34) & "        '允许的上传文件类型" & vbcrlf
	hf.write "Const DelUpFiles=" & chr(34) & trim(request("DelUpFiles")) & chr(34) & "        '删除文章时是否同时删除文章中的上传文件" & vbcrlf
	hf.write "Const ChargeType_999=" & trim(request("ChargeType_999")) & "        '注册用户的默认计费方式" & vbcrlf
	hf.write "Const UserPoint_999=" & trim(request("UserPoint_999")) & "        '注册用户的默认可用点数" & vbcrlf
	hf.write "Const ValidDays_999=" & trim(request("ValidDays_999")) & "        '注册用户的默认有效期" & vbcrlf
	hf.write "Const ChargeType_99=" & trim(request("ChargeType_99")) & "        '收费用户的默认计费方式" & vbcrlf
	hf.write "Const UserPoint_99=" & trim(request("UserPoint_99")) & "        '收费用户的默认可用点数" & vbcrlf
	hf.write "Const ValidDays_99=" & trim(request("ValidDays_99")) & "        '收费用户的默认有效期" & vbcrlf
	hf.write "Const ChargeType_9=" & trim(request("ChargeType_9")) & "        'VIP用户的默认计费方式" & vbcrlf
	hf.write "Const UserPoint_9=" & trim(request("UserPoint_9")) & "        'VIP用户的默认可用点数" & vbcrlf
	hf.write "Const ValidDays_9=" & trim(request("ValidDays_9")) & "        'VIP用户的默认有效期" & vbcrlf
	hf.write "%" & ">"
	hf.close
	set hf=nothing
	set fso=nothing
	call WriteSuccessMsg("网站配置保存成功！")
end sub
%>