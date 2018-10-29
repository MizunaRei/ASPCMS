<%@language=vbscript codepage=936 %>
<%
option explicit
response.buffer=true	
Const PurviewLevel=2
Const CheckChannelID=5
Const AdminType=True
'response.write "此功能被WEBBOY暂时禁止了！"
'response.end
dim Action,PurviewLevel_Guest
Action=request("Action")
if Action="" then
	Action="Check"
end if
select case Action
	case "adminreply"
		PurviewLevel_Guest="Reply"
	case "del"
		PurviewLevel_Guest="Del"
	case "pass","nopass"
		PurviewLevel_Guest="Check"
	case "edit"
		PurviewLevel_Guest="Modify"
	case else
		PurviewLevel_Guest="Manage"
end select
%>
<!--#include file="inc/conn.asp"-->
<!--#include file="inc/config.asp"-->
<!--#include file="Admin_ChkPurview.asp"-->
<!--#include file="inc/Conn_User.asp"-->
<!--#include file="inc/ubbcode.asp"-->
<!--#include file="inc/function.asp"-->
<!--#include file="inc/admin_code_guest.asp"-->
<!--#include file="inc/md5.asp"-->
<%
dim strChannel,sqlChannel,rsChannel,ChannelUrl,ChannelName
dim strFileName,MaxPerPage,totalPut,CurrentPage,TotalPages
dim BeginTime,EndTime,founderr, errmsg,i
dim rs,sql,rsGuest,sqlGuest
dim PageTitle,strPath,strPageTitle
dim SkinID,ClassID,AnnounceCount
dim UserGuestName,UserType,UserSex,UserEmail,UserHomepage,UserOicq,UserIcq,UserMsn
dim WriteName,WriteType,WriteSex,WriteEmail,WriteOicq,WriteIcq,WriteMsn,WriteHomepage
dim WriteFace,WriteImages,WriteTitle,WriteContent,SaveEdit,SaveEditId
dim GuestType,LoginName,AdminReplyContent
dim SubmitType,GuestPath,TitleName,keyword
Set rsGuest= Server.CreateObject("ADODB.Recordset")

dim Purview_ReplyGuest,Purview_DelGuest,Purview_CheckGuest,Purview_ModifyGuest
Purview_ReplyGuest=False
Purview_DelGuest=False
Purview_CheckGuest=False
Purview_ModifyGuest=False
if AdminPurview=1 or CheckPurview(AdminPurview_Guest,"Reply")=True then Purview_ReplyGuest=True
if AdminPurview=1 or CheckPurview(AdminPurview_Guest,"Del")=True then Purview_DelGuest=True
if AdminPurview=1 or CheckPurview(AdminPurview_Guest,"Check")=True then Purview_CheckGuest=True
if AdminPurview=1 or CheckPurview(AdminPurview_Guest,"Modify")=True then Purview_ModifyGuest=True

strFileName="Admin_Guest.asp"
GuestPath="images/guestbook/"
MaxPerPage=10
SaveEdit=0							
if request("page")<>"" then
    currentPage=cint(request("page"))
else
	currentPage=1
end if
TitleName=ChannelName
select case Action
	case "write"
		PageTitle="签写留言"
	case "savewrite"
		PageTitle="保存留言"
	case "reply"
		PageTitle="回复留言"
	case "edit"
		PageTitle="编辑留言"
	case "adminreply"
		PageTitle="管理员回复留言"
	case "del"
		PageTitle="删除留言"
	case "pass"
		PageTitle="审核留言"
	case "nopass"
		PageTitle="取消审核"
	case else
		PageTitle="网站留言"
end select

SubmitType=request("SubmitType")
select case SubmitType
	case "待审留言"
		Action="shownopassed" 
	case "删除留言"
		Action="del" 
	case "通过审核"
		Action="pass"
	case "取消审核"
		Action="nopass"
end select

GuestType=0
if CheckUserLogined()=true then
	GuestType=1
	LoginName=Trim(Request.Cookies("asp163")("UserName"))
end if

keyword=trim(request("keyword"))
if keyword<>"" then 
	keyword=ReplaceBadChar(keyword)
	keyword=Replace(keyword,"[","")
	keyword=Replace(keyword,"]","")
end if
if keyword<>"" then TitleName="搜索含有 <font color=red>"&keyword&"</font> 的留言"
%>

<html>
<title>网站留言管理</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="Admin_Style.css" rel="stylesheet" type="text/css">
<body leftmargin="2" topmargin="0" marginwidth="0" marginheight="0">
<%
call showtip()
call Guestbook()
%>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
<tr class="tdbg_leftall"> 
  <td> 
	<%
	call ShowGuestPage()
	%>
  </td>
</tr>
</table>
</body>
</html>

<%
'以下内容syscode_guest.asp与Admin_Guest.asp相同

'=================================================
'过程名：GuestBook()
'作  用：留言本功能调用
'参  数：无
'=================================================
sub GuestBook()
	select case Action
		case "write"
			call WriteGuest()
		case "savewrite"
			call SaveWriteGuest()
		case "reply"
			call ReplyGuest()
		case "edit"
			call EditGuest()
		case "adminreply"
			call AdimReplyGuest()
		case "saveadminreply"
			call SaveAdminReplyGuest()
		case "del"
			call DelGuest()
		case "pass"
			call PassGuest()
		case "nopass"
			call PassGuest()
		case "user"
			call ShowAllGuest(3)
		case else
			call GuestMain()
	end select
end sub

'=================================================
'过程名：GuestMain()
'作  用：留言主函数
'参  数：无
'=================================================
sub GuestMain()
	response.write "<form style=""margin:0;padding:0"">"
	if action="shownopassed" then
		call ShowAllGuest(4)
	else
		call ShowAllGuest(0)
	end if
	call ShowGuestBottom()
	response.write "</form>"
end sub

'=================================================
'过程名：ShowAllGuest()
'作  用：分页显示所有留言
'参  数：ShowType-----  0为显示所有
'						1为显示已通过审核及用户自己发表的留言
'						2为显示已通过审核的留言（用于游客显示）
'						3为显示用户自己发表的留言
'=================================================
sub ShowAllGuest(ShowType)
	if ShowType=1 then
		sqlGuest="select * from Guest where (GuestIsPassed=True or GuestName='"&LoginName&"')"
	elseif ShowType=2 then
		sqlGuest="select * from Guest where GuestIsPassed=True"
	elseif ShowType=3 then
		sqlGuest="select * from Guest where GuestName='"&LoginName&"'"
	elseif ShowType=4 then
		sqlGuest="select * from Guest where GuestIsPassed=False"
	else
		if keyword<>"" then
			sqlGuest="select * from Guest where 1"
		else
			sqlGuest="select * from Guest"
		end if
	end if
	if keyword<>"" then
		sqlGuest=sqlGuest & " and (GuestTitle like '%" & keyword & "%' or GuestContent like '%" & keyword & "%' or GuestName like '%" & keyword & "%' or GuestReply like '%" & keyword & "%') "
	end if

	sqlGuest=sqlGuest&" order by GuestMaxId desc"
	set rsGuest=server.createobject("adodb.recordset")
	rsGuest.open sqlGuest,conn,1,1
	if rsGuest.bof and rsGuest.eof then
		totalput=0
		response.write "<br><li>没有任何留言</li>"
	else
		totalput=rsGuest.recordcount
		if currentPage=1 then
			call ShowGuestList()
		else
			if (currentPage-1)*MaxPerPage<totalPut then
         	   	rsGuest.move  (currentPage-1)*MaxPerPage
         		dim bookmark
           		bookmark=rsGuest.bookmark
            	call ShowGuestList()
        	else
	        	currentPage=1
           		call ShowGuestList()
	    	end if
		end if
	end if
	rsGuest.close
	set rsGuest=nothing
end sub

'=================================================
'过程名：ShowGuestList()
'作  用：显示留言
'参  数：无
'=================================================
sub ShowGuestList()
	dim i,GuestTip,TipName,TipSex,TipEmail,TipOicq,TipHomepage,isdelUser
	i=0
	do while not rsGuest.eof
		isdelUser=0
		if rsGuest("GuestType")=1 then
			sql="select * from " & db_User_Table & " where " & db_User_Name & "='" & rsGuest("GuestName")&"'"
			set rs=server.createobject("adodb.recordset")
			rs.open sql,Conn_User,1,1
			if not rs.bof and not rs.eof then
				UserGuestName=rs(db_User_Name)
				UserSex=rs(db_User_Sex)
				UserEmail=rs(db_User_Email)
				UserOicq=rs(db_User_QQ)
				UserIcq=rs(db_User_Icq)
				UserMsn=rs(db_User_Msn)
				UserHomepage=rs(db_User_Homepage)
			else
				isdelUser=1
			end if
		end if
		if rsGuest("GuestType")<>1 or isdelUser=1 then
			UserGuestName=rsGuest("GuestName")
			UserSex=rsGuest("GuestSex")
			UserEmail=rsGuest("GuestEmail")
			UserOicq=rsGuest("GuestOicq")
			UserIcq=rsGuest("GuestIcq")
			UserMsn=rsGuest("GuestMsn")
			UserHomepage=rsGuest("GuestHomepage")
		end if
		TipName=UserGuestName
		if isdelUser=1 then TipName=TipName&"（已删除）"
		if TipEmail="" or isnull(TipEmail) then TipEmail="未填"
		if TipOicq="" or isnull(TipOicq) then TipOicq="未填"
		if TipHomepage="" or isnull(TipHomepage) then TipHomepage="未填"
		if UserIcq="" or isnull(UserIcq) then UserIcq="未填"
		if UserMsn="" or isnull(UserMsn) then UserMsn="未填"
		if UserSex=1 then
			TipSex="（酷哥）"
		elseif UserSex=0 then
			TipSex="(靓妹)"
		else
			TipSex=""
		end if
		GuestTip="&nbsp;姓名："&TipName&" "&TipSex&"<br>&nbsp;邮件："&TipEmail&"<br>&nbsp;OICQ："&TipOicq&"<br>&nbsp;主页："&TipHomepage&"<br>&nbsp;地址："&rsGuest("GuestIP")&"<br>&nbsp;时间："&rsGuest("GuestDatetime")
		%>
		<SCRIPT language=javascript>
		function CheckAll(form)
		{
		  for (var i=0;i<form.elements.length;i++)
			{
			var e = form.elements[i];
			if (e.Name != "chkAll"&&e.disabled!=true)
			   e.checked = form.chkAll.checked;
			}
		}
		</script>
		<table width="99%" border="0" cellspacing="0" cellpadding="0" align="center">
		  <tr>
			<td> 
			  <table width="100%" border="0" cellpadding="0" cellspacing="0" class="border">
				<tr> 
				<tr > 
				  <td align="center" valign="top"> 
					<table width="100%" border="0" cellspacing="2" cellpadding="1" class="title">
					  <tr> 
						<td ><font color=green>&nbsp;&nbsp;主题</font>:&nbsp;<%=KeywordReplace(rsGuest("GuestTitle"))%></td>
						<td width="165"> <img src="<%=GuestPath%>posttime.gif" width="11" height="11" align="absmiddle"> 
						  <font color="#006633">： 
						  <% =rsGuest("GuestDatetime")%>
						  </font> </td>
					  </tr>
					</table>
				  </td>
				</tr>
				<tr class="tdbg_leftall"> 
				  <td align="center" height="153" valign="top"> 
					<table width="100%" border="0" cellpadding="0" cellspacing="0">
					  <%if rsGuest("GuestIsPassed")=True then%>
						<tr class="tdbg_leftall"> 
					  <%else%>
						<tr bgcolor="#f0f0f0"> 
					  <%end if%>
						<td width="100" align="center" height="130" valign="top"> 
						  <table width="100%" border="0" cellspacing="0" cellpadding="3" align="center" >
							<tr> 
							  <td valign="middle" align="center" width="100%">
								<img src="<%=GuestPath%><%=rsGuest("GuestImages")%>.gif" width="80" height="90" onMouseOut=toolTip() onMouseOver="toolTip('<%=GuestTip%>')"><br><br>
								<%
								if rsGuest("GuestType")=1 then
									response.write "<font color=""#006633"">【用户】<br>"&KeywordReplace(UserGuestName)&"</font>"
								else
									response.write "【游客】<br>"&KeywordReplace(UserGuestName)
								end if
								%>
							  </td>
							</tr>
						  </table>
						</td>
						<td align="center" height="153" width="1" bgcolor="#B4C9E7"> 
						</td>
						<td> 
						  <table width="100%" border="0" cellpadding="6" cellspacing="0" class="saytext" height="125" style="TABLE-LAYOUT: fixed">
							<tr> 
							  <td align="left" valign="top"><img src="<%=GuestPath%>face<%=rsGuest("GuestFace")%>.gif" width="19" height="19"> 
								<%
								  if rsGuest("GuestIsPrivate")=true then
									response.write "<font color=green>[隐藏]</font>&nbsp;"
								  end if
								  response.write KeywordReplace(ubbcode(dvHTMLEncode(rsGuest("GuestContent"))))
								  %>
                      </td>
                    </tr>
                    <tr>
                      <td align="left" valign="bottom"> 
                        <%call ShowGuestreply()%>
                      </td>
                    </tr>
                  </table>
						  <table width="100%" height="1" border="0" cellpadding="0" cellspacing="0" bgcolor="#B4C9E7">
							<tr> 
							  <td></td>
							</tr>
						  </table>
						  <%call ShowGuestButton()%>
						</td>
					  </tr>
					</table>
				  </td>
				</tr>
			  </table>
			  <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
				<tr> 
				  <td  height="15" align="center" valign="top"> 
					<table width="98%" border="0" align="center" cellpadding="0" cellspacing="0">
					  <tr> 
						<td height="13" class="tdbg_left2"></td>
					  </tr>
					</table>
				  </td>
				</tr>
			  </table>
			  </td>
		  </tr>
		</table>
		<%
		rsGuest.movenext
		i=i+1
		if i>=MaxPerPage then exit do
	loop
end sub

'=================================================
'过程名：ShowGuestreply()
'作  用：显示回复留言
'参  数：无
'=================================================
sub ShowGuestreply()
	if len(rsGuest("GuestReply")) >0 then
	%>
<table width="100%" border="0" cellpadding="1" cellspacing="0">
  <tr> 
		<td> 
		  
      <table width="100%" border="0" cellspacing="0" cellpadding="2">
        <tr> 
          <td height="1" bgcolor="#B4C9E7"></td>
        </tr>
        <tr>
          <td valign="top"><table width="100%" border="0" cellpadding="0" cellspacing="0" style="TABLE-LAYOUT: fixed">
              <tr> 
                <td><font color="#006633"> 管理员<font color="#FF0000">[<%=rsGuest("GuestReplyAdmin")%>]</font>回复: 
                  <% =rsGuest("GuestReplyDatetime") %>
                  </font> </td>
              </tr>
              <tr> 
                <td valign="bottom"><font color="#006633"><% =KeywordReplace(ubbcode(dvHTMLEncode(rsGuest("GuestReply")))) %></font></td>
              </tr>
            </table></td>
        </tr>
      </table>
		</td>
	  </tr>
	</table>
	<%
	end if
end sub

'**************************************************
'函数名：KeywordReplace
'作  用：标示搜索关键字
'参  数：strChar-----要转换的字符
'返回值：转换后的字符
'**************************************************
function KeywordReplace(strChar)
	if strChar="" then
		KeywordReplace=""
	else
		KeywordReplace=	replace(strChar,""&keyword&"","<font color=red>"&keyword&"</font>")
	end if
end function

'=================================================
'过程名：WriteGuest()
'作  用：签写留言
'参  数：无
'=================================================
sub WriteGuest()
if SaveEdit<>1 then
	WriteType=GuestType
	WriteName=LoginName
	WriteSex="1"
	WriteFace="1"
	WriteImages="01"
	WriteHomepage="http://"
end if
%>
<script language=JavaScript>
function changeimage()
{ 
	document.formwrite.GuestImages.value=document.formwrite.Image.value;
	document.formwrite.showimages.src="<%=GuestPath%>"+document.formwrite.Image.value+".gif";
}
function guestpreview()
{
document.preview.content.value=document.formwrite.GuestContent.value;
var popupWin = window.open('GuestPreview.asp', 'GuestPreview', 'scrollbars=yes,width=620,height=230');
document.preview.submit()
}
function check(thisform)
{
   if(thisform.GuestName.value==""){
		alert("姓名不能为空！")
		 thisform.GuestName.focus()
		  return(false) 
	  }

   if(thisform.GuestTitle.value==""){
		alert("留言主题不能为空！")
		thisform.GuestTitle.focus()
		return(false)
	  }

   if(thisform.GuestContent.value==""){
		alert("留言内容不能为空！")
		thisform.GuestContent.focus()
		  return(false)
	  }

   if(thisform.GuestContent.value.length>800){
		alert("留言内容不能超过800字符！")
		thisform.GuestContent.focus()
		  return(false)
	  }
   if(thisform.reg.checked==true){
	   if(thisform.psw.value=="" || thisform.psw.value.length<6){
			alert("如果注册，注册密码不能为空，且长度至少六位！")
			thisform.psw.focus()
			  return(false)
		  }
	   if(thisform.pswc.value=="" || thisform.pswc.value!=thisform.psw.value){
			alert("确认密码不能为空，且需与注册密码相同！")
			thisform.pswc.focus()
			  return(false)
		  }
	   if(thisform.question.value==""){
			alert("密码问题不能为空！")
			thisform.question.focus()
			  return(false)
		  }
	   if(thisform.answer.value==""){
			alert("问题答案不能为空！")
			thisform.answer.focus()
			  return(false)
		  }
	  }

}
function showreginfo(){
if (document.formwrite.reg.checked == true) {
	reginfo.style.display = "";
//	reginfoshowtext.innerText="暂时不想注册成为贵站会员"
}else{
	reginfo.style.display = "none";
//	reginfoshowtext.innerText="我想同时注册成为贵站会员"
}
}
function gbcount(message,total,used,remain)
{
	var max;
	max = total.value;
	if (event.keyCode==13 && event.ctrlKey)
	formwrite.Submit1.click();
	if (message.value.length > max) {
	message.value = message.value.substring(0,max);
	used.value = max;
	remain.value = 0;
	alert("留言不能超过 1000 个字!");
	}
	else {
	used.value = message.value.length;
	remain.value = max - used.value;
	}
}
</script>
<table width="100%" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr> 
    <td> 
      <table width="100%" cellpadding="1" cellspacing="0" class="border">
        <tr class="title"> 
          <td colspan="5">&nbsp;&nbsp;&nbsp;&nbsp;<font color=green><%=PageTitle%></font></td>
        </tr>
        <tr class="tdbg_leftall"> 
          <td colspan="5" align="center" height="10"></td>
        </tr>
        <form name="formwrite" method="post" action="<%=strFileName%>?action=savewrite" onSubmit="return check(formwrite)">
          <% if WriteType=0 then%>
          <tr class="tdbg_leftall"> 
            <td width="20%" align="center">姓 &nbsp;名:</td>
            <td width="30%"> 
              <input type="text" name="GuestName" maxlength="14" size="20" value="<%=WriteName%>">
			  <font color=red>*</font>
            </td>
            <td width="22%">&nbsp; </td>
            <td colspan="2">&nbsp; </td>
          </tr>
          <tr class="tdbg_leftall"> 
            <td align="center">性&nbsp;&nbsp;别:</td>
            <td> 
              <input type="radio" name="GuestSex" value="1" <% if WriteSex=1 then response.write" checked"%> style="BORDER:0px;">
              男&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
              <input type="radio" name="GuestSex" value="0" <% if WriteSex=0 then response.write" checked"%> style="BORDER:0px;">
              女 </td>
            <td> &nbsp;&nbsp; 
              <select name="Image" size="1" onChange="changeimage();" >
                <%
				for i=1 to 9
					response.write "<option value='0"&i&"'>0"&i&"</option>"
				next
				for i=10 to 23
					response.write "<option value='"&i&"'>"&i&"</option>"
				next
				%>
              </select>
            </td>
            <td colspan="2">&nbsp;</td>
          </tr>
          <tr class="tdbg_leftall"> 
            <td align="center">E-mail</td>
            <td> 
              <input type="text" name="GuestEmail" maxlength="30" size="20" value="<%=WriteEmail%>">
            </td>
            <td rowspan="4"> 
              <input type="hidden" name="GuestImages" value="01">
			  <img name=showimages src="<%=GuestPath%><%=WriteImages%>.gif" width="80" height="90" border="0" onClick=window.open("guestselect.asp?action=guestimages","face","width=480,height=400,resizable=1,scrollbars=1") title=点击选择头像 style="cursor:hand">
			  </td>
            <td colspan="2" rowspan="4">
              <table width="100%" border="0" cellspacing="0" cellpadding="0" id=reginfo style="DISPLAY: none">
                <tr> 
                  <td width="35%" align="center">注册密码：<br>
                  </td>
                  <td> 
                    <input type=password maxlength=16 size=16 name=psw>
                  </td>
                </tr>
                <tr> 
                  <td align="center">密码确认：</td>
                  <td> 
                    <input type=password maxlength=16 size=16 name=pswc>
                  </td>
                </tr>
                <tr> 
                  <td align="center">密码问题：</td>
                  <td> 
                    <input type=text maxlength=30 size=16 name=question>
                  </td>
                </tr>
                <tr> 
                  <td align="center">问题答案：</td>
                  <td> 
                    <input type=text maxlength=30 size=16 name=answer>
                  </td>
                </tr>
              </table>
            </td>
          </tr>
          <tr class="tdbg_leftall"> 
            <td align="center">Oicq:</td>
            <td> 
              <input type="text" name="GuestOicq" maxlength="15" size="20" value="<%=WriteOicq%>">
            </td>
          </tr>
          <tr class="tdbg_leftall"> 
            <td align="center">Icq:</td>
            <td> 
              <input type="text" name="GuestIcq" maxlength="15" size="20" value="<%=WriteIcq%>">
            </td>
          </tr>
          <tr class="tdbg_leftall"> 
            <td align="center">Msn:</td>
            <td> 
              <input type="text" name="GuestMsn" maxlength="40" size="20" value="<%=WriteMsn%>">
            </td>
          </tr>
          <tr class="tdbg_leftall"> 
            <td align="center">个人主页:</td>
			<td colspan="4"> 
              <input type="text" name="GuestHomepage" maxlength="80" size="37" value="<%=WriteHomepage%>">
              &nbsp;&nbsp;&nbsp;&nbsp; 
             <input type="checkbox" name="reg" value="1" onClick=showreginfo() style="BORDER:0px;">
              <span id=reginfoshowtext>我想同时注册成为贵站会员</span>
			</td>
          </tr>
          <tr class="tdbg_leftall"> 
            <td align="center"></td>
            <td colspan="4">&nbsp; </td>
          </tr>
          <%else%>
          <tr class="tdbg_leftall"> 
            <td align="center">选择头像：</td>
            <td> 
              <input type="hidden" name="GuestName"  value="<%=WriteName%>">
              <input type="hidden" name="reg" value="1">
			  <input type="hidden" name="GuestImages" value="<%=WriteImages%>">
			  <img name=showimages src="<%=GuestPath%><%=WriteImages%>.gif" width="80" height="90" border="0" onClick=window.open("guestselect.asp?action=guestimages","face","width=480,height=400,resizable=1,scrollbars=1") title=点击选择头像 style="cursor:hand">
              <select name="Image" size="1" onChange="changeimage();" >
                <%
				for i=1 to 9
					response.write "<option value='0"&i&"'>0"&i&"</option>"
				next
				for i=10 to 23
					response.write "<option value='"&i&"'>"&i&"</option>"
				next
				%>
              </select>
            </td>
            <td>&nbsp;</td>
            <td>&nbsp;</td>
            <td>&nbsp;</td>
          </tr>
          <%end if%>
          <tr class="tdbg_leftall"> 
            <td align="center">留言主题:</td>
            <td colspan="4"> 
              <input type="text" name="GuestTitle" size="37" maxlength="28" value="<%=WriteTitle%>">
			  <font color=red>*</font>
            </td>
          </tr>
          <tr class="tdbg_leftall"> 
            <td align="center">现在心情:</td>
            <td colspan="4"> 
              <%
				for i=1 to 20
					response.write "<input type=""radio"" name=""GuestFace"" value="&i&""
					if i=clng(WriteFace) then response.write " checked"
					response.write " style=""BORDER:0px;width:19;"">"
					response.write "<img src="""&GuestPath&"face"&i&".gif"" width=""19"" height=""19"">"& vbcrlf
					if i mod 10 =0 then response.write "<br>"
				next
				%>
            </td>
          </tr>
          <tr class="tdbg_leftall"> 
            <td align="center">Ubb标签:</td>
            <td colspan="4"> 
              <% call showubb()%>
            </td>
          </tr>
          <tr class="tdbg_leftall"> 
            <td valign="middle" align="center">留言内容: <br>
              (Ctrl+Enter提交)</td>
            <td colspan="4" valign="top"> 
              <textarea name="GuestContent" cols="59" rows="6" title='按 Ctrl+Enter 可直接发送'   onkeydown=gbcount(this.form.GuestContent,this.form.total,this.form.used,this.form.remain); onkeyup=gbcount(this.form.GuestContent,this.form.total,this.form.used,this.form.remain);><%=WriteContent%></textarea>
            </td>
          </tr>
          <tr class="tdbg_leftall"> 
            <td valign="middle" align="center"></td>
            <td colspan="4" valign="top"> 
				最大字数：<INPUT disabled maxLength=4 name=total size=3 value=500>
				已用字数：<INPUT disabled maxLength=4 name=used size=3 value=0>
				剩余字数：<INPUT disabled maxLength=4 name=remain size=3 value=500>
			</td>
          </tr>
          <tr class="tdbg_leftall"> 
            <td valign="middle" align="center">是否隐藏：</td>
            <td colspan="4" valign="top"> 
              <input type="radio" name="GuestIsPrivate" value="no"  checked style="BORDER:0px;">
              正常 
              <input type="radio" name="GuestIsPrivate" value="yes" style="BORDER:0px;">
              隐藏 * 选择隐藏后，此留言只有站长才可以看到。</td>
          </tr>
          <tr class="tdbg_leftall"> 
            <td colspan="5" align="center"  height="40"> 
              <input type="hidden" name="saveedit"  value="<%=SaveEdit%>">
              <input type="hidden" name="saveeditid"  value="<%=SaveEditId%>">
              <input type="submit" name="Submit1" value=" 发 表" >
              &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
              <input type="button" name="Submit2" value=" 预 览 " onclick=guestpreview()>
              &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
              <input type="reset" name="Submit3" value=" 重 填 " >
            </td>
          </tr>
        </form>
		<form name=preview action="GuestPreview.asp" method=post target=GuestPreview>
		<input type=hidden name=title value=><input type=hidden name=content value=>
		</form>
      </table>
      <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
        <tr> 
          <td  height="15" align="center" valign="top"> 
            <table width="98%" border="0" align="center" cellpadding="0" cellspacing="0">
              <tr> 
                <td height="13" class="tdbg_left2"></td>
              </tr>
            </table>
          </td>
        </tr>
      </table>
    </td>
  </tr>
</table>
<%
end sub

'=================================================
'过程名：ReplyGuest()
'作  用：回复留言
'参  数：无
'=================================================
sub ReplyGuest()
	dim ReplyId
	ReplyId=request("guestid")
	if ReplyId="" then
		call Guest_info("<li>请指定要回复的留言ID！</li>")
		exit sub
	else
		ReplyId=clng(ReplyId)
		sqlGuest="select * from Guest where GuestId=" & ReplyId
	end if
	set rsGuest=server.createobject("adodb.recordset")
	rsGuest.open sqlGuest,conn,1,1
	if rsGuest.bof and rsGuest.eof then
		response.write "<br><li>没有任何留言</li>"
		exit sub
	else
		WriteTitle="Re: "&rsGuest("GuestTitle")
		call ShowGuestList()
	end if
	rsGuest.close
	set rsGuest=nothing
	call WriteGuest()
end sub


'=================================================
'过程名：EditGuest()
'作  用：编辑留言
'参  数：无
'=================================================
sub EditGuest()
	dim EditId
	EditId=request("guestid")
	if EditId="" then
		call Guest_info("<li>请指定要编辑的留言ID！</li>")
		exit sub
	else
		EditId=clng(EditId)
		sqlGuest="select * from Guest where GuestId=" & EditId
	end if
	set rsGuest=server.createobject("adodb.recordset")
	rsGuest.open sqlGuest,conn,1,1
	if rsGuest.bof and rsGuest.eof then
		response.write "<br><li>找不到您指定的留言！</li>"
		exit sub
	end if
	
	if Purview_ModifyGuest=True then
		WriteName=rsGuest("GuestName")
		WriteType=rsGuest("GuestType")
		WriteSex=rsGuest("GuestSex")
		WriteEmail=rsGuest("GuestEmail")
		WriteOicq=rsGuest("GuestOicq")
		WriteIcq=rsGuest("GuestIcq")
		WriteMsn=rsGuest("GuestMsn")
		WriteHomepage=rsGuest("GuestHomepage")
		WriteFace=rsGuest("GuestFace")
		WriteImages=rsGuest("GuestImages")
		WriteTitle=rsGuest("GuestTitle")
		WriteContent=rsGuest("GuestContent")
		SaveEdit=1
		SaveEditId=EditId
		call ShowGuestList()
		call WriteGuest()
	else
		call Guest_info("<li>您没有编辑留言的权限！</li>")
	end if    
	rsGuest.close
	set rsGuest=nothing
end sub

'=================================================
'过程名：AdimReplyGuest()
'作  用：站长回复留言
'参  数：无
'=================================================
sub AdimReplyGuest()
	dim AdminReplyId
	if Purview_ReplyGuest=False then
		call Guest_info("<li>您没有回复留言的权限！</li>")
	else
		AdminReplyId=request("guestid")
		if AdminReplyId="" then
			call Guest_info("<li>请指定要回复的留言ID！</li>")
			exit sub
		else
			AdminReplyId=clng(AdminReplyId)
			sqlGuest="select * from Guest where GuestId=" & AdminReplyId
		end if
		set rsGuest=server.createobject("adodb.recordset")
		rsGuest.open sqlGuest,conn,1,1
		if rsGuest.bof and rsGuest.eof then
			response.write "<br><li>找不到您指定的留言！</li>"
			exit sub
		else
			AdminReplyContent=rsGuest("GuestReply")
			call ShowGuestList()
		end if
		rsGuest.close
		set rsGuest=nothing
		call WriteAdimReplyGuest()
	end if    
end sub

'=================================================
'过程名：WriteAdimReplyGuest()
'作  用：填写站长回复留言
'参  数：无
'=================================================
sub WriteAdimReplyGuest()
%>
	<script language=JavaScript>
	function check(thisform)
	{
	   if(thisform.GuestContent.value==""){
			alert("留言内容不能为空！")
			thisform.GuestContent.focus()
			  return(false)
		  }

	   if(thisform.GuestContent.value.length>800){
			alert("留言内容不能超过800字符！")
			thisform.GuestContent.focus()
			  return(false)
		  }
	}
	</script>
	<table width="100%" border="0" cellspacing="0" cellpadding="0" align="center">
	  <tr> 
		<td> 
		  <table width="100%" cellpadding="1" cellspacing="0" class="border">
			<tr class="title"> 
			  <td colspan="3">&nbsp;&nbsp;&nbsp;&nbsp;<font color=green><%=PageTitle%></font></td>
			</tr>
			<tr class="tdbg_leftall"> 
			  <td colspan="3" align="center" height="10"></td>
			</tr>
			<form name="formwrite" method="post" action="<%=strFileName%>?action=saveadminreply" onSubmit="return check(formwrite)">
			  <tr class="tdbg_leftall"> 
				<td align="center">Ubb标签:</td>
				<td colspan="2"> 
					<%call ShowUbb()%>
				</td>
			  </tr>
			  <tr class="tdbg_leftall"> 
				<td width="18%" valign="middle" align="center">留言内容: </td>
				<td colspan="2" valign="top"> 
				  <textarea name="GuestContent" cols="59" rows="6" ><%=AdminReplyContent%></textarea>
				</td>
			  </tr>
			  <tr class="tdbg_leftall"> 
				<td colspan="3" align="center"  height="40"><input name="guestid" type="hidden" value="<%=request("guestid")%>">
				  <input type="submit" name="Submit1" value=" 发 表" >
				  &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
				  <input type="reset" name="Submit2" value=" 重 填 " >
				</td>
			  </tr>
			</form>
		  </table>
		  <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
			<tr> 
			  <td  height="15" align="center" valign="top"> 
				<table width="98%" border="0" align="center" cellpadding="0" cellspacing="0">
				  <tr> 
					<td height="13" class="tdbg_left2"></td>
				  </tr>
				</table>
			  </td>
			</tr>
		  </table>
		</td>
	  </tr>
	</table>
<%
end sub



'=================================================
'过程名：ShowGuestbutton()
'作  用：显示留言功能按钮
'参  数：无
'=================================================
sub ShowGuestButton()
	response.write "<table width=100% border=0 cellpadding=0 cellspacing=3><tr>"
	response.write "<td>"
	if UserHomepage="" or isnull(UserHomepage) then
		response.write "<img src="&GuestPath&"nourl.gif width=45 height=16 alt="&UserGuestName&"没有留下主页地址 border=0>" & vbcrlf
	else
		response.write "<a href="&UserHomepage&" target=""_blank"">"
		response.write "<img src="&GuestPath&"url.gif width=45 height=16 alt="&UserHomepage&" border=0></a>" & vbcrlf
	end if
	if UserOicq="" or isnull(UserOicq) then
		response.write "<img src="&GuestPath&"nooicq.gif width=45 height=16 alt="&UserGuestName&"没有留下QQ号码 border=0>" & vbcrlf
	else
		response.write "<a href=http://search.tencent.com/cgi-bin/friend/user_show_info?ln="&UserOicq&" target=""_blank"">"
		response.write "<img src="&GuestPath&"oicq.gif width=45 height=16 alt="&UserOicq&" border=0 ></a>" & vbcrlf
	end if
	if UserEmail="" or isnull(UserEmail) then
		response.write "<img src="&GuestPath&"noemail.gif width=45 height=16 alt="&UserGuestName&"没有留下Email地址 border=0>" & vbcrlf
	else
		response.write "<a href=mailto:"&UserEmail&">"
		response.write "<img src="&GuestPath&"email.gif width=45 height=16 border=0 alt="&UserEmail&"></a>" & vbcrlf
	end if
	response.write "<img src="&GuestPath&"other.gif width=45 height=16 border=0 onMouseOut=toolTip() onMouseOver=""toolTip('&nbsp;Icq：" & UserIcq & "<br>&nbsp;Msn：" & UserMsn & "<br>&nbsp;I P：" &rsGuest("GuestIP")&"')"">" & vbcrlf

	response.write "&nbsp;</td><td width=1 bgcolor=#B4C9E7></td><td>&nbsp;"
	if Purview_ModifyGuest=True then
		response.write "<a href="&strFileName&"?action=edit&guestid="&rsGuest("GuestId")&">"
		response.write "<img src="&GuestPath&"edit.gif width=45 height=16 border=0 alt=编辑这条留言></a>" & vbcrlf
	end if
	if Purview_ReplyGuest=True then
		response.write "<a href="&strFileName&"?action=adminreply&guestid="&rsGuest("GuestId")&">"
		response.write "<img src="&GuestPath&"adminreply.gif width=45 height=16 alt=管理员回复这条留言 border=0></a>" & vbcrlf
	end if
	if Purview_DelGuest=True then
		response.write "<a href="&strFileName&"?action=del&guestid="&rsGuest("GuestId")&" onClick=""return confirm('确定要删除此留言吗？');"">"
		response.write "<img src="&GuestPath&"del.gif width=45 height=16  alt=删除这条留言 border=0></a>" & vbcrlf
	end if
	if Purview_CheckGuest=True then
		if rsGuest("GuestIsPassed")=False then
			response.write "<a href="&strFileName&"?action=pass&guestid="&rsGuest("GuestId")&">"
			response.write "<img src="&GuestPath&"pass.gif width=45 height=16  alt=审核通过这条留言 border=0 ></a>" & vbcrlf
		else
			response.write "<a href="&strFileName&"?action=nopass&guestid="&rsGuest("GuestId")&">"
			response.write "<img src="&GuestPath&"nopass.gif width=45 height=16  alt=取消这条留言审核 border=0></a>" & vbcrlf
		end if
	end if
	response.write "&nbsp;<input name=guestid type=checkbox id=guestid value="&rsGuest("GuestID")&" style=BORDER:0px;>"
	response.write "</td>" & vbcrlf
	response.write "</tr></table>"
end sub

'=================================================
'过程名：DelGuest()
'作  用：删除留言
'参  数：无
'=================================================
sub DelGuest()
	dim delid
	delid=trim(Request("guestid"))
	if delid="" then
		call Guest_info("<li>请指定要删除的留言ID！</li>")
		exit sub
	end if
	if instr(delid,",")>0 then
		delid=replace(delid," ","")
		sql="Select * from Guest where GuestID in (" & delid & ")"
	else
		delid=clng(delid)
		sql="select * from Guest where GuestID=" & delid
	end if
	Set rs=Server.CreateObject("Adodb.RecordSet")
	rs.Open sql,conn,1,3
	if rs.bof and rs.eof then
		response.write "<br><li>找不到您指定的留言！</li>"
		exit sub
	end if

	if Purview_DelGuest=False then
		call Guest_info("<li>您没有删除留言的权限！</li>")
	else
		do while not rs.eof
			rs.delete
			rs.update
			rs.movenext
		loop
		rs.close
		set rs=nothing
		response.redirect ComeUrl
	end if    
end sub

'=================================================
'过程名：PassGuest()
'作  用：审核留言
'参  数：无
'=================================================
sub PassGuest()
	dim passid
	if Purview_CheckGuest=False then
		call Guest_info("<li>您没有审核留言的权限！</li>")
	else
		passid=trim(Request("guestid"))
		if passid="" then
			call Guest_info("<li>请指定要审核的留言ID！</li>")
			exit sub
		end if
		if instr(passid,",")>0 then
			passid=replace(passid," ","")
			sql="Select * from Guest where GuestID in (" & passid & ")"
		else
			passid=clng(passid)
			sql="select * from Guest where GuestID=" & passid
		end if
		Set rs=Server.CreateObject("Adodb.RecordSet")
		rs.Open sql,conn,1,3
		do while not rs.eof
			if Action="pass" then
				rs("GuestIsPassed")=True
			else
				rs("GuestIsPassed")=False
			end if
			rs.update
			rs.movenext
		loop
		rs.close
		set rs=nothing
		response.redirect ComeUrl
	end if    
end sub

'=================================================
'过程名：SaveAdminReplyGuest()
'作  用：保存站长回复留言
'参  数：无
'=================================================
sub SaveAdminReplyGuest()
	dim GuestReply,SaveAdminReplyId
	dim sqlMaxId,rsMaxId,MaxId
	if Purview_ReplyGuest=False then
		call Guest_info("<li>您没有回复留言的权限！</li>")
	else
		GuestReply=request("GuestContent")
		SaveAdminReplyId=request("guestid")
		if SaveAdminReplyId="" then
			call Guest_info("<li>请指定要回复的留言ID！</li>")
			exit sub
		end if
		sqlMaxId="select max(GuestMaxId) as MaxId from Guest"
		set rsMaxId=conn.execute(sqlMaxId)
		MaxId=rsMaxId("MaxId")
		if MaxId="" or isnull(MaxId) then MaxId=0
		set rsGuest=server.createobject("adodb.recordset")
		sql="select * from Guest where GuestId="&SaveAdminReplyId
		rsGuest.open sql,conn,3,3
		if rsGuest.bof and rsGuest.eof then
			response.write "<br><li>找不到您指定的留言！</li>"
			exit sub
		else
			rsGuest("GuestMaxId")=MaxId+1
			rsGuest("GuestReply")=GuestReply
			rsGuest("GuestReplyAdmin")=session("AdminName")
			rsGuest("GuestReplyDatetime")=now()
			rsGuest.update
		end if
		rsGuest.close
		set rsGuest=nothing
	end if
	call Guest_info("<li>您的回复留言已经发送成功！</li>")
end sub

'=================================================
'过程名：ShowGuestBottom()
'作  用：显示留言底部管理功能
'参  数：无
'=================================================
sub ShowGuestBottom()
	dim strTemp
	if TotalPut>0 then
	 	strTemp= "<table align='center'><tr><td>"
		strTemp= strTemp & "<input  type=""submit"" name=""SubmitType"" value=""待审留言"">&nbsp;&nbsp;"
		strTemp= strTemp & "&nbsp;&nbsp;多项操作："
		if Purview_DelGuest=True then
			strTemp= strTemp & "<input  type=""submit"" name=""SubmitType"" value=""删除留言"" onClick=""return confirm('确定要删除选中的留言吗？');"">&nbsp;&nbsp;"
		end if
		if Purview_CheckGuest=True then
			strTemp= strTemp & "<input  type=""submit"" name=""SubmitType"" value=""通过审核"">&nbsp;&nbsp;"
			strTemp= strTemp & "<input  type=""submit"" name=""SubmitType"" value=""取消审核"">&nbsp;&nbsp;"
		end if
		strTemp= strTemp & "<input name=""chkAll"" type=""checkbox"" id=""chkAll"" onclick=CheckAll(this.form) value=""checkbox"" style=""BORDER:0px;"">选中本页显示的所有留言"
		strTemp= strTemp & "</td></tr></table>"
		response.write strTemp
	end if
end sub
%>
