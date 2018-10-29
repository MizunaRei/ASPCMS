<%
sub Bottom_All()
%>
<div align="center">	<table bgcolor="#FFFFFF" width="989">
	<tr>
		<td colspan="12" width="989" height="100" background="images/首页_slice2_35.jpg">
			<!--<img src="images/首页_slice2_35.jpg" width="989" height="241" alt="">-->
            <P align=center><B>| <SPAN style="CURSOR: hand" 
      onclick="var strHref=window.location.href;this.style.behavior='url(#default#homepage)';this.setHomePage('http://renwen.university.edu.cn');">设为首页</SPAN> | <SPAN title='两课教学网' style="CURSOR: hand" 
      onclick="window.external.addFavorite('http://renwen.university.edu.cn','两课教学网')">收藏本站</SPAN> | <A  href="mailto:86277298@QQ.COM">联系站长</A> | <A  
      href="http://renwen.university.edu.cn/FriendSite/Index.asp" target=_blank>友情链接</A> | <A  href="http://renwen.university.edu.cn/Copyright.asp" 
      target=_blank>版权申明</A> | </B></P>
      <p align="center">本网站由<font color="#3300FF"><a href="http://renwen.university.edu.cn/">university人文社会科学学院</a></font>主办、维护</p>
            </td>
		
	</tr>
	
</table></div>

<%
end sub

sub Top_noIndex()
%>
<div align="center" ><table id="__01" width="989"  border="0" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF" align="left">
	<!--<tr>
		<td colspan="14">
			<img src="images/首页_slice2_01.jpg" width="1024" height="142" alt=""></td>
		<td>
			<img src="images/分隔符.gif" width="1" height="142" alt=""></td>
	</tr>-->
	<tr>
		<!--<td rowspan="17">
			<img src="images/首页_slice2_02.jpg" width="17" height="858" alt=""></td>-->
		<td colspan="12">
			<img   src="images/首页_slice2_03.jpg" width="989" height="188" alt=""></td>
		<!--<td rowspan="17">
			<img src="images/首页_slice2_04.jpg" width="18" height="858" alt=""></td>-->
		<!--<td>
			<img src="images/分隔符.gif" width="1" height="188" alt=""></td>-->
	</tr>
	<tr>
		<!--<td colspan="12">
			<img src="images/首页_slice2_05.jpg" width="989" height="25" alt=""></td>-->
            <td  align="left" colspan="12"  background="images/首页_slice_05.jpg" width="989" height="25"><%call ShowPath()%> </td>
		<td>
			<img src="images/分隔符.gif" width="1" height="25" alt=""></td>
	</tr>
    </table><!--top--></div>
    <%
end sub

'学生改文章列出老师
sub User_ArticleModify_TeacherList()
'		  select case rs("purview")
'		    case 1
'              strPurview="<font color=blue>超级管理员</font>"
'            case 2
'              strpurview="教师管理员"
'			 case 3
'			 	strpurview="学生管理员"
'		  end select
'		  if rs("purview")=3 then
'		   	response.Write("<tr> <td width='40%' class='tdbg'><strong>学    号：</strong></td> <td width='60%' class='tdbg'>  <input name=StudentNumber type=text value=" & rs("StudentNumber")  & "></td></tr>" )
'		  	response.Write("<tr> <td width='40%' class='tdbg'><strong>任课教师：</strong></td> <td width='60%' class='tdbg'>  <select name='TeacherName' ><option value='" & rs("TeacherName") & "'>" & rs("TeacherName") )		  
'			response.Write("<option value='景庆虹'>景庆虹<option value='林震'>林震<option value='路军'>路军<option value='罗美云'>罗美云<option value='宋兵波'>宋兵波<option value='吴守蓉'>吴守蓉<option value='杨志华'>杨志华<option value='于延周'>于延周<option value='张连伟'>张连伟<option value='赵海燕'>赵海燕")
'			response.Write("<option value='朱洪强'>朱洪强 <option value='钟爱军'>钟爱军<option value='陈丽鸿'>陈丽鸿<option value='戴秀丽'>戴秀丽<option value='高兴武'>高兴武<option value='赵亮'>赵亮<option value='周国文'>周国文<option value='张宁'>张宁")
'			response.Write("</select></td></tr>")
'		  end if 
		  
		  'response.write(strPurview)
        if  1<3    then
			
				dim sqlAdmin_Teacher,rsAdmin_Teacher
				sqlAdmin_Teacher = "select TrueName,Purview from Admin where Purview=2"
	
				set rsAdmin_Teacher=server.CreateObject("adodb.recordset")
				rsAdmin_Teacher.open sqlAdmin_Teacher,conn,1,1
	
				
				
'				   	Response.Write("<table  id='StudentAdminPurviewDetail'  style='display:none'   ><tr><td> <strong>学&nbsp;&nbsp;号：</strong> </td> " )
'				  	Response.Write("<td><input name=StudentNumber type=text value=" & rs("StudentNumber")  & "></td></tr>")
					Response.Write("<table><tr> <td> " )		  
					Response.Write("<select name='TeacherName' id='TeacherName' >")
					
					rsAdmin_Teacher.movefirst
						do while not rsAdmin_Teacher.eof
							if  ( rsAdmin_Teacher("Purview")=2 ) then 
								if rsArticle("TeacherName")=rsAdmin_Teacher("TrueName") then
																	Response.Write("<option value='" & rsAdmin_Teacher("TrueName") & "'selected>"   &  rsAdmin_Teacher("TrueName") & "</option>" )

								else
									Response.Write("<option value='" & rsAdmin_Teacher("TrueName") & "'>"   &  rsAdmin_Teacher("TrueName") & "</option>" )
								end if
							end if
						rsAdmin_Teacher.movenext
						loop
					response.Write("</select></td></tr></table>")
				'关闭数据库连接
				rsAdmin_Teacher.close
   				 set rsAdmin_Teacher = nothing
				'关闭数据库连接

		end if


end sub


'学生写文章列出老师
sub User_ArticleTeacherList()
'		  select case rs("purview")
'		    case 1
'              strPurview="<font color=blue>超级管理员</font>"
'            case 2
'              strpurview="教师管理员"
'			 case 3
'			 	strpurview="学生管理员"
'		  end select
'		  if rs("purview")=3 then
'		   	response.Write("<tr> <td width='40%' class='tdbg'><strong>学    号：</strong></td> <td width='60%' class='tdbg'>  <input name=StudentNumber type=text value=" & rs("StudentNumber")  & "></td></tr>" )
'		  	response.Write("<tr> <td width='40%' class='tdbg'><strong>任课教师：</strong></td> <td width='60%' class='tdbg'>  <select name='TeacherName' ><option value='" & rs("TeacherName") & "'>" & rs("TeacherName") )		  
'			response.Write("<option value='景庆虹'>景庆虹<option value='林震'>林震<option value='路军'>路军<option value='罗美云'>罗美云<option value='宋兵波'>宋兵波<option value='吴守蓉'>吴守蓉<option value='杨志华'>杨志华<option value='于延周'>于延周<option value='张连伟'>张连伟<option value='赵海燕'>赵海燕")
'			response.Write("<option value='朱洪强'>朱洪强 <option value='钟爱军'>钟爱军<option value='陈丽鸿'>陈丽鸿<option value='戴秀丽'>戴秀丽<option value='高兴武'>高兴武<option value='赵亮'>赵亮<option value='周国文'>周国文<option value='张宁'>张宁")
'			response.Write("</select></td></tr>")
'		  end if 
		  
		  'response.write(strPurview)
        if  1<3    then
			
				dim sqlAdmin_Teacher,rsAdmin_Teacher
				sqlAdmin_Teacher = "select TrueName,Purview from Admin where Purview=2"
	
				set rsAdmin_Teacher=server.CreateObject("adodb.recordset")
				rsAdmin_Teacher.open sqlAdmin_Teacher,conn,1,1
	
				
				
'				   	Response.Write("<table  id='StudentAdminPurviewDetail'  style='display:none'   ><tr><td> <strong>学&nbsp;&nbsp;号：</strong> </td> " )
'				  	Response.Write("<td><input name=StudentNumber type=text value=" & rs("StudentNumber")  & "></td></tr>")
					Response.Write("<table><tr><td> " )		  
					Response.Write("<select name='TeacherName' id='TeacherName' ><option value=''>请选择任课教师" )
					
					rsAdmin_Teacher.movefirst
						do while not rsAdmin_Teacher.eof
							if  ( rsAdmin_Teacher("Purview")=2 ) then 
								Response.Write("<option value='" & rsAdmin_Teacher("TrueName") & "'>"   &  rsAdmin_Teacher("TrueName") & "</option>" )
							end if
						rsAdmin_Teacher.movenext
						loop
					response.Write("</select><font color='#FF0000'>*请选择任课教师</font></td></tr></table>")
				'关闭数据库连接
				rsAdmin_Teacher.close
   				 set rsAdmin_Teacher = nothing
				'关闭数据库连接

		end if


end sub
'***************************************************
'显示所有栏目文章列表
'****************************************************
sub ShowClassArticleList()



end sub

'***************************************************
'显示某栏目文章名称列表
'****************************************************
dim rsShowClassArticleListName,sqlShowClassArticleListName

function ShowClassArticleListName()

end function 


dim UserLogined,UserName,UserLevel,ChargeType,UserPoint,ValidDays

'**************************************************
'函数名：gotTopic
'作  用：截字符串，汉字一个算两个字符，英文算一个字符
'参  数：str   ----原字符串
'       strlen ----截取长度
'返回值：截取后的字符串
'**************************************************
function gotTopic(str,strlen)
	if str="" then
		gotTopic=""
		exit function
	end if
	dim l,t,c, i
	str=replace(replace(replace(replace(str,"&nbsp;"," "),"&quot;",chr(34)),"&gt;",">"),"&lt;","<")
	l=len(str)
	t=0
	for i=1 to l
		c=Abs(Asc(Mid(str,i,1)))
		if c>255 then
			t=t+2
		else
			t=t+1
		end if
		if t>=strlen then
			gotTopic=left(str,i) & "…"
			exit for
		else
			gotTopic=str
		end if
	next
	gotTopic=replace(replace(replace(replace(gotTopic," ","&nbsp;"),chr(34),"&quot;"),">","&gt;"),"<","&lt;")
end function

'**************************************************
'函数名：JoinChar
'作  用：向地址中加入 ? 或 &
'参  数：strUrl  ----网址
'返回值：加了 ? 或 & 的网址
'**************************************************
function JoinChar(strUrl)
	if strUrl="" then
		JoinChar=""
		exit function
	end if
	if InStr(strUrl,"?")<len(strUrl) then 
		if InStr(strUrl,"?")>1 then
			if InStr(strUrl,"&")<len(strUrl) then 
				JoinChar=strUrl & "&"
			else
				JoinChar=strUrl
			end if
		else
			JoinChar=strUrl & "?"
		end if
	else
		JoinChar=strUrl
	end if
end function

'**************************************************
'过程名：showpage
'作  用：显示“上一页 下一页”等信息
'参  数：sfilename  ----链接地址
'       totalnumber ----总数量
'       maxperpage  ----每页数量
'       ShowTotal   ----是否显示总数量
'       ShowAllPages ---是否用下拉列表显示所有页面以供跳转。有某些页面不能使用，否则会出现JS错误。
'       strUnit     ----计数单位
'**************************************************
sub showpage(sfilename,totalnumber,maxperpage,ShowTotal,ShowAllPages,strUnit)
	dim n, i,strTemp,strUrl
	if totalnumber mod maxperpage=0 then
    	n= totalnumber \ maxperpage
  	else
    	n= totalnumber \ maxperpage+1
  	end if
  	strTemp= "<table align='center'><tr><td>"
	if ShowTotal=true then 
		strTemp=strTemp & "共 <b>" & totalnumber & "</b> " & strUnit & "&nbsp;&nbsp;"
	end if
	strUrl=JoinChar(sfilename)
  	if CurrentPage<2 then
    		strTemp=strTemp & "首页 上一页&nbsp;"
  	else
    		strTemp=strTemp & "<a href='" & strUrl & "page=1'>首页</a>&nbsp;"
    		strTemp=strTemp & "<a href='" & strUrl & "page=" & (CurrentPage-1) & "'>上一页</a>&nbsp;"
  	end if

  	if n-currentpage<1 then
    		strTemp=strTemp & "下一页 尾页"
  	else
    		strTemp=strTemp & "<a href='" & strUrl & "page=" & (CurrentPage+1) & "'>下一页</a>&nbsp;"
    		strTemp=strTemp & "<a href='" & strUrl & "page=" & n & "'>尾页</a>"
  	end if
   	strTemp=strTemp & "&nbsp;页次：<strong><font color=red>" & CurrentPage & "</font>/" & n & "</strong>页 "
    strTemp=strTemp & "&nbsp;<b>" & maxperpage & "</b>" & strUnit & "/页"
	if ShowAllPages=True then
		strTemp=strTemp & "&nbsp;转到：<select name='page' size='1' onchange=""javascript:window.location='" & strUrl & "page=" & "'+this.options[this.selectedIndex].value;"">"   
    	for i = 1 to n   
    		strTemp=strTemp & "<option value='" & i & "'"
			if cint(CurrentPage)=cint(i) then strTemp=strTemp & " selected "
			strTemp=strTemp & ">第" & i & "页</option>"   
	    next
		strTemp=strTemp & "</select>"
	end if
	strTemp=strTemp & "</td></tr></table>"
	response.write strTemp
end sub

'**************************************************
'函数名：IsValidEmail
'作  用：检查Email地址合法性
'参  数：email ----要检查的Email地址
'返回值：True  ----Email地址合法
'       False ----Email地址不合法
'**************************************************
function IsValidEmail(email)
	dim names, name, i, c
	IsValidEmail = true
	names = Split(email, "@")
	if UBound(names) <> 1 then
	   IsValidEmail = false
	   exit function
	end if
	for each name in names
		if Len(name) <= 0 then
			IsValidEmail = false
    		exit function
		end if
		for i = 1 to Len(name)
		    c = Lcase(Mid(name, i, 1))
			if InStr("abcdefghijklmnopqrstuvwxyz_-.", c) <= 0 and not IsNumeric(c) then
		       IsValidEmail = false
		       exit function
		     end if
	   next
	   if Left(name, 1) = "." or Right(name, 1) = "." then
    	  IsValidEmail = false
	      exit function
	   end if
	next
	if InStr(names(1), ".") <= 0 then
		IsValidEmail = false
	   exit function
	end if
	i = Len(names(1)) - InStrRev(names(1), ".")
	if i <> 2 and i <> 3 then
	   IsValidEmail = false
	   exit function
	end if
	if InStr(email, "..") > 0 then
	   IsValidEmail = false
	end if
end function

'**************************************************
'函数名：IsObjInstalled
'作  用：检查组件是否已经安装
'参  数：strClassString ----组件名
'返回值：True  ----已经安装
'       False ----没有安装
'**************************************************
Function IsObjInstalled(strClassString)
	On Error Resume Next
	IsObjInstalled = False
	Err = 0
	Dim xTestObj
	Set xTestObj = Server.CreateObject(strClassString)
	If 0 = Err Then IsObjInstalled = True
	Set xTestObj = Nothing
	Err = 0
End Function

'**************************************************
'函数名：strLength
'作  用：求字符串长度。汉字算两个字符，英文算一个字符。
'参  数：str  ----要求长度的字符串
'返回值：字符串长度
'**************************************************
function strLength(str)
	ON ERROR RESUME NEXT
	dim WINNT_CHINESE
	WINNT_CHINESE    = (len("中国")=2)
	if WINNT_CHINESE then
        dim l,t,c
        dim i
        l=len(str)
        t=l
        for i=1 to l
        	c=asc(mid(str,i,1))
            if c<0 then c=c+65536
            if c>255 then
                t=t+1
            end if
        next
        strLength=t
    else 
        strLength=len(str)
    end if
    if err.number<>0 then err.clear
end function

'------------------检查某一目录是否存在-------------------
Function CheckDir(FolderPath)
	dim fso
	folderpath=Server.MapPath(".")&"\"&folderpath
	Set fso1 = Server.CreateObject("Scripting.FileSystemObject")
	If fso.FolderExists(FolderPath) then
	'存在
		CheckDir = True
	Else
	'不存在
		CheckDir = False
	End if
	Set fso = nothing
End Function

'-------------根据指定名称生成目录---------
Function MakeNewsDir(foldername)
	dim fso,f
	Set fso = Server.CreateObject("Scripting.FileSystemObject")
    Set f = fso.CreateFolder(foldername)
    MakeNewsDir = True
	Set fso = nothing
End Function


'**************************************************
'函数名：SendMail
'作  用：用Jmail组件发送邮件
'参  数：MailtoAddress  ----收信人地址
'        MailtoName    -----收信人姓名
'        Subject       -----主题
'        MailBody      -----信件内容
'        FromName      -----发信人姓名
'        MailFrom      -----发信人地址
'        Priority      -----信件优先级
'**************************************************
function SendMail(MailtoAddress,MailtoName,Subject,MailBody,FromName,MailFrom,Priority)
	on error resume next
	Dim JMail
	Set JMail=Server.CreateObject("JMail.Message")
	if err then
		SendMail= "<br><li>没有安装JMail组件</li>"
		err.clear
		exit function
	end if
	JMail.Charset="gb2312"          '邮件编码
	JMail.silent=true
	JMail.ContentType = "text/html"     '邮件正文格式
	'JMail.ServerAddress=MailServer     '用来发送邮件的SMTP服务器
   	'如果服务器需要SMTP身份验证则还需指定以下参数
	JMail.MailServerUserName = MailServerUserName    '登录用户名
   	JMail.MailServerPassWord = MailServerPassword        '登录密码
  	JMail.MailDomain = MailDomain       '域名（如果用“name@domain.com”这样的用户名登录时，请指明domain.com
	JMail.AddRecipient MailtoAddress,MailtoName     '收信人
	JMail.Subject=Subject         '主题
	JMail.HMTLBody=MailBody       '邮件正文（HTML格式）
	JMail.Body=MailBody          '邮件正文（纯文本格式）
	JMail.FromName=FromName         '发信人姓名
	JMail.From = MailFrom         '发信人Email
	JMail.Priority=Priority              '邮件等级，1为加急，3为普通，5为低级
	JMail.Send(MailServer)
	SendMail =JMail.ErrorMessage
	JMail.Close
	Set JMail=nothing
end function

'**************************************************
'过程名：WriteErrMsg
'作  用：显示错误提示信息
'参  数：无
'**************************************************
sub WriteErrMsg()
	dim strErr
	strErr=strErr & "<html><head><title>错误信息</title><meta http-equiv='Content-Type' content='text/html; charset=gb2312'>" & vbcrlf
	strErr=strErr & "<link href='style.css' rel='stylesheet' type='text/css'></head><body><br><br>" & vbcrlf
	strErr=strErr & "<table cellpadding=2 cellspacing=1 border=0 width=400 class='border' align=center>" & vbcrlf
	strErr=strErr & "  <tr align='center' class='title'><td height='22'><strong>错误信息</strong></td></tr>" & vbcrlf
	strErr=strErr & "  <tr class='tdbg'><td height='100' valign='top'><b>产生错误的可能原因：</b>" & errmsg &"</td></tr>" & vbcrlf
	strErr=strErr & "  <tr align='center' class='tdbg'><td><a href='javascript:history.go(-1)'>&lt;&lt; 返回上一页</a></td></tr>" & vbcrlf
	strErr=strErr & "</table>" & vbcrlf
	strErr=strErr & "</body></html>" & vbcrlf
	response.write strErr
end sub

'**************************************************
'过程名：WriteSuccessMsg
'作  用：显示成功提示信息
'参  数：无
'**************************************************
sub WriteSuccessMsg(SuccessMsg)
	dim strSuccess
	strSuccess=strSuccess & "<html><head><title>成功信息</title><meta http-equiv='Content-Type' content='text/html; charset=gb2312'>" & vbcrlf
	strSuccess=strSuccess & "<link href='style.css' rel='stylesheet' type='text/css'></head><body><br><br>" & vbcrlf
	strSuccess=strSuccess & "<table cellpadding=2 cellspacing=1 border=0 width=400 class='border' align=center>" & vbcrlf
	strSuccess=strSuccess & "  <tr align='center' class='title'><td height='22'><strong>恭喜你！</strong></td></tr>" & vbcrlf
	strSuccess=strSuccess & "  <tr class='tdbg'><td height='100' valign='top'><br>" & SuccessMsg &"</td></tr>" & vbcrlf
	strSuccess=strSuccess & "  <tr align='center' class='tdbg'><td>&nbsp;</td></tr>" & vbcrlf
	strSuccess=strSuccess & "</table>" & vbcrlf
	strSuccess=strSuccess & "</body></html>" & vbcrlf
	response.write strSuccess
end sub

'**************************************************
'函数名：CheckUserLogined
'作  用：检查用户是否登录
'参  数：无
'返回值：True ----已经登录
'        False ---没有登录
'**************************************************
function CheckUserLogined()
	dim Logined,Password,rsLogin,sqlLogin
	Logined=True
	UserName=Request.Cookies("asp163")("UserName")
	Password=Request.Cookies("asp163")("Password")
	UserLevel=Request.Cookies("asp163")("UserLevel")
	if UserName="" then
		Logined=False
	end if
	if Password="" then
		Logined=False
	end if
	if UserLevel="" then
		Logined=False
		UserLevel=9999
	end if
	if Logined=True then
		username=replace(trim(username),"'","")
		password=replace(trim(password),"'","")
		UserLevel=Cint(trim(UserLevel))
		set rsLogin=server.createobject("adodb.recordset")
		sqlLogin="select * from " & db_User_Table & " where " & db_User_LockUser & "=False and " & db_User_Name & "='" & username & "' and " & db_User_Password & "='" & password &"'"
		rsLogin.open sqlLogin,Conn_User,1,1
		if rsLogin.bof and rsLogin.eof then
			Logined=False
		else
			if password<>rsLogin(db_User_Password) or UserLevel<rsLogin(db_User_UserLevel) then
				Logined=False
			end if
			UserName=rsLogin(db_User_Name)
			UserLevel=rsLogin(db_User_UserLevel)
			ChargeType=rsLogin(db_User_ChargeType)
			UserPoint=rsLogin(db_User_UserPoint)
		  	if rsLogin(db_User_Valid_Unit)=1 then
				ValidDays=rsLogin(db_User_Valid_Num)
		  	elseif rsLogin(db_User_Valid_Unit)=2 then
				ValidDays=rsLogin(db_User_Valid_Num)*30
		  	elseif rsLogin(db_User_Valid_Unit)=3 then
				ValidDays=rsLogin(db_User_Valid_Num)*365
		  	end if
		  	ValidDays=ValidDays-DateDiff("D",rsLogin(db_User_BeginDate),now())
		end if
		rsLogin.close
		set rsLogin=nothing
	end if
	CheckUserLogined=Logined
end function

'**************************************************
'函数名：ReplaceBadChar
'作  用：过滤非法的SQL字符
'参  数：strChar-----要过滤的字符
'返回值：过滤后的字符
'**************************************************
function ReplaceBadChar(strChar)
	if strChar="" then
		ReplaceBadChar=""
	else
		ReplaceBadChar=replace(replace(replace(replace(replace(replace(replace(strChar,"'",""),"*",""),"?",""),"(",""),")",""),"<",""),".","")
	end if
end function

'**************************************************
'函数名：CheckLevel
'作  用：检查用户级别
'参  数：LevelNum-----要检查的级别值
'返回值：级别名称
'**************************************************
function CheckLevel(LevelNum)
	select case LevelNum
	case 9999
		CheckLevel="游客"
	case 999
		CheckLevel="注册用户"
	case 99
		CheckLevel="收费用户"
	case 9
		CheckLevel="VIP用户"
	case 5
		CheckLevel="管理员"
	end select
end function

'==================================================
'过程名：ShowLogo
'作  用：显示网站LOGO
'参  数：无
'==================================================
sub ShowLogo()
	if LogoUrl<>"" then
		response.write "<a href='" & SiteUrl & "' title='" & SiteName & "'>"
		if lcase(right(LogoUrl,3))<>"swf" then
			response.write "<img src='" & LogoUrl & "' border='0'>"
		else
			Response.Write "<object classid='clsid:D27CDB6E-AE6D-11cf-96B8-444553540000' codebase='http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=5,0,0,0' width='180' height='60'><param name='movie' value='" & LogoUrl & "'><param name='quality' value='high'><embed src='" & LogoUrl & "' pluginspage='http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash' type='application/x-shockwave-flash' width='480' height='60'></embed></object>"
		end if
		response.write "</a>"
	else
		response.write "<a href='http://www.asp163.net' title='动力空间'><img src='http://www.asp163.net/Photo/images/logo.gif' border='0'></a>"
	end if
end sub

'==================================================
'过程名：ShowBanner
'作  用：显示网站Banner
'参  数：无
'==================================================
sub ShowBanner()
	if BannerUrl<>"" then
		if lcase(right(BannerUrl,3))="swf" then
			Response.Write "<object classid='clsid:D27CDB6E-AE6D-11cf-96B8-444553540000' codebase='http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=5,0,0,0' width='480' height='60'><param name='movie' value='" & BannerUrl & "'><param name='quality' value='high'><embed src='" & BannerUrl & "' pluginspage='http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash' type='application/x-shockwave-flash' width='480' height='60'></embed></object>"
		else
			response.Write "<a href='" & SiteUrl & "' title='" & SiteName & "'><img src='" & BannerUrl & "' width='480' height='60' border='0'></a>"
		end if
	else
		call ShowAD(1)
	end if
end sub

'==================================================
'过程名：ShowVote
'作  用：显示网站调查
'参  数：无
'==================================================
sub ShowVote()
	dim sqlVote,rsVote,i
	sqlVote="select top 1 * from Vote where IsSelected=True"
	sqlVote=sqlVote& " and (ChannelID=0 or ChannelID=" & ChannelID & ") order by ID Desc"
	Set rsVote= Server.CreateObject("ADODB.Recordset")
	rsVote.open sqlVote,conn,1,1
	if rsVote.bof and rsVote.eof then 
		response.Write "&nbsp;没有任何相关调查"
	else
		response.write "<form name='VoteForm' method='post' action='vote.asp' target='_blank'>"
		response.write "<center><font color=ff3300>" & rsVote("Title") & "</font></center>"
		if rsVote("VoteType")="Single" then
			for i=1 to 8
				if trim(rsVote("Select" & i) & "")="" then exit for
				response.Write "<input type='radio' name='VoteOption' value='" & i & "' >" & rsVote("Select" & i) & "<br>"
			next
		else
			for i=1 to 8
				if trim(rsVote("Select" & i) & "")="" then exit for
				response.Write "<input type='checkbox' name='VoteOption' value='" & i & "' style='border:0'>" & rsVote("Select" & i) & "<br>"
			next
		end if
		response.write "<input name='VoteType' type='hidden'value='" & rsVote("VoteType") & "'>"
		response.write "<input name='Action' type='hidden' value='Vote'>"
		response.write "<input name='ID' type='hidden' value='" & rsVote("ID") & "'>"
		response.write "<div align='center'>"
		response.write "<a href='javascript:VoteForm.submit();'><img src='images/voteSubmit.gif' width='52' height='18' border='0'></a>&nbsp;&nbsp;"
        response.write "<a href='Vote.asp?ID=" & rsVote("ID") & "&Action=Show' target='_blank'><img src='images/voteView.gif' width='52' height='18' border='0'></a>"
		response.write "</div></form>"
	end if
	rsVote.close
	set rsVote=nothing
end sub
'==================================================
'过程名：ShowAnnounce
'作  用：显示本站公告信息
'参  数：ShowType ------显示方式，1为纵向，2为横向
'        AnnounceNum  ----最多显示多少条公告
'==================================================
sub ShowAnnounce(ShowType,AnnounceNum)
	dim sqlAnnounce,rsAnnounce,i
	if AnnounceNum>0 and AnnounceNum<=10 then
		sqlAnnounce="select top " & AnnounceNum
	else
		sqlAnnounce="select top 10"
	end if
	sqlAnnounce=sqlAnnounce & " * from Announce where IsSelected=True"
	sqlAnnounce=sqlAnnounce & " and (ChannelID=0 or ChannelID=" & ChannelID & ")"
	sqlAnnounce=sqlAnnounce & " and (ShowType=0 or ShowType=1) order by ID Desc"
	Set rsAnnounce= Server.CreateObject("ADODB.Recordset")
	rsAnnounce.open sqlAnnounce,conn,1,1
	if rsAnnounce.bof and rsAnnounce.eof then 
		AnnounceCount=0
		response.write "<p>&nbsp;&nbsp;没有通告</p>" 
	else 
		AnnounceCount=rsAnnounce.recordcount
		if ShowType=1 then
			do while not rsAnnounce.eof   
				response.Write "&nbsp;&nbsp;&nbsp;&nbsp;<a href='#' onclick=""javascript:window.open('Announce.asp?ChannelID=" & ChannelID & "&ID=" & rsAnnounce("id") &"', 'newwindow', 'height=300, width=400, toolbar=no, menubar=no, scrollbars=auto, resizable=no, location=no, status=no')"" title='" & rsAnnounce("Content") & "'>" & rsAnnounce("title") & "</div><br><div align='right'>" & rsAnnounce("Author") & "&nbsp;&nbsp;<br>" & FormatDateTime(rsAnnounce("DateAndTime"),1) & "</a>&nbsp;&nbsp;"
				rsAnnounce.movenext
				i=i+1
				if i<AnnounceCount then response.write "<hr>"   
			loop
		else
			do while not rsAnnounce.eof   
				response.Write "&nbsp;&nbsp;&nbsp;&nbsp;<a href='#' onclick=""javascript:window.open('Announce.asp?ChannelID=" & ChannelID & "&ID=" & rsAnnounce("id") &"', 'newwindow', 'height=300, width=400, toolbar=no, menubar=no, scrollbars=auto, resizable=no, location=no, status=no')"" title='" & rsAnnounce("Content") & "' >" & rsAnnounce("title") & "&nbsp;&nbsp;[" & rsAnnounce("Author") & "&nbsp;&nbsp;" & FormatDateTime(rsAnnounce("DateAndTime"),1) & "]</a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
				rsAnnounce.movenext
			loop
       	end if	
	end if  
	rsAnnounce.close
	set rsAnnounce=nothing
end sub

'==================================================
'过程名：ShowAnnounce_Index
'作  用：显示本站公告信息
'参  数：ShowType ------显示方式，1为纵向，2为横向
'        AnnounceNum  ----最多显示多少条公告
'==================================================
sub ShowAnnounce_Index(ShowType,AnnounceNum,ChannelID_Index)
	dim sqlAnnounce,rsAnnounce,i
	if AnnounceNum>0 and AnnounceNum<=10 then
		sqlAnnounce="select top " & AnnounceNum
	else
		sqlAnnounce="select top 10"
	end if
	sqlAnnounce=sqlAnnounce & " * from Announce where IsSelected=True"
	sqlAnnounce=sqlAnnounce & " and (ChannelID=0 or ChannelID=" & ChannelID_Index & ")"
	sqlAnnounce=sqlAnnounce & " and (ShowType=0 or ShowType=1) order by ID Desc"
	Set rsAnnounce= Server.CreateObject("ADODB.Recordset")
	rsAnnounce.open sqlAnnounce,conn,1,1
	if rsAnnounce.bof and rsAnnounce.eof then 
		AnnounceCount=0
		response.write "<p>&nbsp;&nbsp;没有通告</p>" 
	else 
		AnnounceCount=rsAnnounce.recordcount
		if ShowType=1 then
			do while not rsAnnounce.eof   
				response.Write "&nbsp;&nbsp;&nbsp;&nbsp;<a href='#' onclick=""javascript:window.open('Announce.asp?ChannelID=" & ChannelID & "&ID=" & rsAnnounce("id") &"', 'newwindow', 'height=300, width=400, toolbar=no, menubar=no, scrollbars=auto, resizable=no, location=no, status=no')"" title='" & rsAnnounce("Content") & "'>" & rsAnnounce("title") & "</div><br><div align='right'>" & rsAnnounce("Author") & "&nbsp;&nbsp;<br>" & FormatDateTime(rsAnnounce("DateAndTime"),1) & "</a>&nbsp;&nbsp;"
				rsAnnounce.movenext
				i=i+1
				if i<AnnounceCount then response.write "<hr>"   
			loop
		else
			do while not rsAnnounce.eof   
				response.Write "&nbsp;&nbsp;&nbsp;&nbsp;<a href='#' onclick=""javascript:window.open('Announce.asp?ChannelID=" & ChannelID & "&ID=" & rsAnnounce("id") &"', 'newwindow', 'height=300, width=400, toolbar=no, menubar=no, scrollbars=auto, resizable=no, location=no, status=no')"" title='" & rsAnnounce("Content") & "' >" & rsAnnounce("title") & "&nbsp;&nbsp;[" & rsAnnounce("Author") & "&nbsp;&nbsp;" & FormatDateTime(rsAnnounce("DateAndTime"),1) & "]</a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
				rsAnnounce.movenext
			loop
       	end if	
	end if  
	rsAnnounce.close
	set rsAnnounce=nothing
end sub


'end sub showannounce_Index()
'==================================================
'过程名：ShowFriendSite
'作  用：显示友情链接站点
'参  数：LinkType  ----链接方式，1为LOGO链接，2为文字链接
'       SiteNum   ----最多显示多少个站点
'       Cols      ----分几列显示
'       ShowType  ----显示方式。1为向上滚动，2为横向列表，3为下拉列表框
'==================================================
sub ShowFriendSite(LinkType,SiteNum,Cols,ShowType)
	dim sqlLink,rsLink,SiteCount,i,strLink
	if LinkType<>1 and LinkType<>2 then
		LinkType=1
	else
		LinkType=Cint(LinkType)
	end if
	if SiteNum<=0 or SiteNum>100 then
		SiteNum=10
	end if
	if Cols<=0 or Cols>20 then
		Cols=10
	end if
	if ShowType=1 then
'		strLink=strLink & "<marquee id='LinkScrollArea' direction='up' scrolldelay='50' scrollamount='1' width='100' height='100' onmouseover='this.stop();' onmouseout='this.start();'>"
        strLink=strLink & "<div id=rolllink style=overflow:hidden;height:100;width:100><div id=rolllink1>"    '新增加的代码
	elseif ShowType=3 then
		strLink=strLink & "<select name='FriendSite' onchange=""if(this.options[this.selectedIndex].value!=''){window.open(this.options[this.selectedIndex].value,'_blank');}""><option value=''>友情文字链接站点</option>"
	end if
	if ShowType=1 or ShowType=2 then
		strLink=strLink & "<table width='100%' cellSpacing='5'><tr align='center' class='tdbg'>"
	end if
	
	sqlLink="select top " & SiteNum & " * from FriendSite where IsOK=True and LinkType=" & LinkType & " order by IsGood,id desc"
	set rsLink=server.createobject("adodb.recordset")
	rsLink.open sqlLink,conn,1,1
	if rsLink.bof and rsLink.eof then
		if ShowType=1 or ShowType=2 then
	  		for i=1 to SiteNum
				strLink=strLink & "<td><a href='FriendSiteReg.asp' target='_blank'>"
				if LinkType=1 then
					strLink=strLink & "<img src='images/nologo.jpg' width='88' height='31' border='0' alt='点击申请'>"
				else
					strLink=strLink & "点击申请"
				end if
				strLink=strLink & "</a></td>"
				if i mod Cols=0 and i<SiteNum then
					strLink=strLink & "</tr><tr align='center' class='tdbg'>"
				end if
			next
		end if
	else
		SiteCount=rsLink.recordcount
		for i=1 to SiteCount
			if ShowType=1 or ShowType=2 then
			  if LinkType=1 then
				strLink=strLink & "<td width='88'><a href='" & rsLink("SiteUrl") & "' target='_blank' title='网站名称：" & rsLink("SiteName") & vbcrlf & "网站地址：" & rsLink("SiteUrl") & vbcrlf & "网站简介：" & rsLink("SiteIntro") & "'>"
				if rsLink("LogoUrl")="" or rsLink("LogoUrl")="http://" then
					strLink=strLink & "<img src='images/nologo.gif' width='88' height='31' border='0'>"
				else
					strLink=strLink & "<img src='" & rsLink("LogoUrl") & "' width='88' height='31' border='0'>"
				end if
				strLink=strLink & "</a></td>"
			  else
				strLink=strLink & "<td width='88'><a href='" & rsLink("SiteUrl") & "' target='_blank' title='网站名称：" & rsLink("SiteName") & vbcrlf & "网站地址：" & rsLink("SiteUrl") & vbcrlf & "网站简介：" & rsLink("SiteIntro") & "'>" & rsLink("SiteName") & "</a></td>"
			  end if
			  if i mod Cols=0 and i<SiteNum then
				strLink=strLink & "</tr><tr align='center' class='tdbg'>"
			  end if
			else
				strLink=strLink & "<option value='" & rsLink("SiteUrl") & "'>" & rsLink("SiteName") & "</option>"
			end if
			rsLink.moveNext
		next
		if SiteCount<SiteNum and (ShowType=1 or ShowType=2) then
			for i=SiteCount+1 to SiteNum
				if LinkType=1 then
					strLink=strLink & "<td width='88'><a href='FriendSiteReg.asp' target='_blank'><img src='images/nologo.jpg' width='88' height='31' border='0' alt='点击申请'></a></td>"
				else
					strLink=strLink & "<td width='88'><a href='FriendSiteReg.asp' target='_blank'>点击申请</a></td>"
				end if
				if i mod Cols=0 and i<SiteNum then
					strLink=strLink & "</tr><tr align='center' class='tdbg'>"
				end if
			next
		end if
	end if
	if ShowType=1 or ShowType=2 then
		strLink=strLink & "</tr></table>"
	end if
	if ShowType=1 then
'		strLink=strLink & "</marquee>"
        strLink=strLink & "</div><div id=rolllink2></div></div>"   '新增代码
	elseif ShowType=3 then
		strLink=strLink & "</select>"
	end if
	response.write strLink
	if ShowType=1 then call RollFriendSite()    '新增代码
	rsLink.close
	set rsLink=nothing
end sub

'==================================================
'过程名：RollFriendSite
'作  用：滚动显示友情链接站点
'参  数：无
'==================================================
sub RollFriendSite()
%>
<script>
   var rollspeed=30
   rolllink2.innerHTML=rolllink1.innerHTML //克隆rolllink1为rolllink2
   function Marquee(){
   if(rolllink2.offsetTop-rolllink.scrollTop<=0) //当滚动至rolllink1与rolllink2交界时
   rolllink.scrollTop-=rolllink1.offsetHeight  //rolllink跳到最顶端
   else{
   rolllink.scrollTop++
   }
   }
   var MyMar=setInterval(Marquee,rollspeed) //设置定时器
   rolllink.onmouseover=function() {clearInterval(MyMar)}//鼠标移上时清除定时器达到滚动停止的目的
   rolllink.onmouseout=function() {MyMar=setInterval(Marquee,rollspeed)}//鼠标移开时重设定时器
</script>
<%
end sub

sub ShowGoodSite(SiteNum)
	dim sqlLink,rsLink,SiteCount,i,strLink
	if SiteNum<=0 or SiteNum>100 then
		SiteNum=10
	end if
	strLink=strLink & "<table width='100%' cellSpacing='5'>"
	
	sqlLink="select top " & SiteNum & " * from FriendSite where IsOK=True and LinkType=1 and IsGood=True order by id desc"
	set rsLink=server.createobject("adodb.recordset")
	rsLink.open sqlLink,conn,1,1
	if rsLink.bof and rsLink.eof then
	 	for i=1 to SiteNum
			strLink=strLink & "<tr align='center'><td><a href='FriendSiteReg.asp' target='_blank'><img src='images/nologo.jpg' width='88' height='31' border='0' alt='点击申请'></a></td></tr>"
		next
	else
		SiteCount=rsLink.recordcount
		for i=1 to SiteCount
			strLink=strLink & "<tr align='center'><td><a href='" & rsLink("SiteUrl") & "' target='_blank' title='网站名称：" & rsLink("SiteName") & vbcrlf & "网站地址：" & rsLink("SiteUrl") & vbcrlf & "网站简介：" & rsLink("SiteIntro") & "'>"
			if rsLink("LogoUrl")="" or rsLink("LogoUrl")="http://" then
				strLink=strLink & "<img src='images/nologo.gif' width='88' height='31' border='0'>"
			else
				strLink=strLink & "<img src='" & rsLink("LogoUrl") & "' width='88' height='31' border='0'>"
			end if
			strLink=strLink & "</a></td></tr>"
			rsLink.moveNext
		next
		for i=SiteCount+1 to SiteNum
			strLink=strLink & "<tr align='center'><td><a href='FriendSiteReg.asp' target='_blank'><img src='images/nologo.jpg' width='88' height='31' border='0' alt='点击申请'></a></td></tr>"
		next
	end if
	strLink=strLink & "</table>"
	response.write strLink
	rsLink.close
	set rsLink=nothing

end sub

sub Bottom()
	dim strTemp
	strTemp="<table width='770' align='center' border='0'  bgcolor=#949694 cellpadding='0' cellspacing='0' height=50><tr align='center'><td >"
	strTemp= strTemp & "copyRight 2006-2008,26265.cn Inc. All Rights Reserved"
	strTemp= strTemp & "</td></tr></table>"
	response.write strTemp
end sub


'==================================================
'过程名：ShowUserLogin
'作  用：显示用户登录表单
'参  数：无
'==================================================
sub ShowUserLogin()
	dim strLogin
	if CheckUserLogined()=False or session("AdminName")<>"" then
    	if session("AdminName")="" then
	strLogin="<table align='center' width='100%' border='0' cellspacing='0' cellpadding='0'>" & vbcrlf
		strLogin=strLogin &  "<form action='User_ChkLogin.asp' method='post' name='UserLogin' onSubmit='return CheckForm();'>" & vbcrlf
        strLogin=strLogin & "<tr><td height='25' align='right'>用户名：</td><td height='25'><input name='UserName' type='text' id='UserName' size='10' maxlength='20'></td></tr>" & vbcrlf
        strLogin=strLogin & "<tr><td height='25' align='right'>密&nbsp;&nbsp;码：</td><td height='25'><input name='Password' type='password' id='Password' size='10' maxlength='20'></td></tr>" & vbcrlf
'用户管理员统一登录界面插件修改代码开始
	'验证码不兼容XP SP3及以后操作系统
	'strLogin=strLogin & "<tr><td height='25' align='right'>验证码：</td><td height='25'><input name='CheckCode' size='6' maxlength='4'><img src='inc/checkcode.asp'></td></tr>" & vbcrlf
'插入代码结束
        strLogin=strLogin & "<tr><td height='25' align='right'>Cookie：</td><td height='25'><select name=CookieDate><option selected value=0>不保存</option><option value=1>保存一天</option>" & vbcrlf
		strLogin=strLogin & "<option value=2>保存一月</option><option value=3>保存一年</option></select></td></tr>" & vbcrlf
		strLogin=strLogin & "<tr align='center'><td height='30' colspan='2'><input name='Login' type='submit' id='Login' value=' 登录 '> <input name='Reset' type='reset' id='Reset' value=' 清除 '><br>" & vbcrlf
      else
	strLogin="<table align='center' width='100%' border='0' cellspacing='0' cellpadding='0'><tr align='center'><td height='30' colspan='2'>" & vbcrlf
		strLogin=strLogin &  "欢迎！<font color=green><b>" & session("AdminName") & vbcrlf
		strLogin=strLogin &  "</b></font><br>" & vbcrlf
		strLogin=strLogin & "您的身份：后台管理员<br>谢谢您的无私贡献！" & vbcrlf
		
      end if
        if session("AdminName")="" then
	  strLogin=strLogin & "<a href='User_Reg.asp' target='_blank'>用户注册</a>&nbsp;&nbsp;" & vbcrlf
	  strLogin=strLogin & "<a href='User_GetPassword.asp'>忘记密码</a>" & vbcrlf
	  strLogin=strLogin & "<br></td></tr></form></table>" & vbcrlf
	  response.write strLogin
	else
	  strLogin=strLogin & "<br><center><a href=""JavaScript:openScript('Admin_Index.asp')"">【进入后台管理中心】</a>" & vbcrlf
	  strLogin=strLogin & "<div align='center'><a href='Admin_Logout.asp'>【退出时记得注销登录】</a></div>" & vbcrlf	
	  strLogin=strLogin & "</td></tr></form></table>" & vbcrlf
	  response.write strLogin

	end if
%>
<script language=javascript>
	function CheckForm()
	{
		if(document.UserLogin.UserName.value=="")
		{
			alert("请输入用户名！");
			document.UserLogin.UserName.focus();
			return false;
		}
		if(document.UserLogin.Password.value == "")
		{
			alert("请输入密码！");
			document.UserLogin.Password.focus();
			return false;
		}
		if (document.UserLogin.CheckCode.value=="")
		{
       			alert ("请输入您的验证码！");
       			document.UserLogin.CheckCode.focus();
       			return(false);
    		}
	}
	function openScript(url, width, height)
	{
		var Win = window.open(url,"UserControlPad",'width=' + width + ',height=' + height + ',resizable=1,scrollbars=yes,menubar=yes,status=yes' );
	}
</script>
<%
	Else 
		response.write "欢迎<font color=green><b>" & UserName & "</b></font>！"
		response.write "<br>您在的身份："
		if UserLevel=999 then
			response.write "注册用户"
		elseif UserLevel=99 then
			response.write "收费用户"
		elseif UserLevel=9 then
			response.write "高级用户"
		end if
		response.write "<br><center><a href=""JavaScript:openScript('User_ControlPad.asp?Action=user_main')"">【个人管理中心】</a></center>" & vbcrlf
		response.write "<div align='center'><a href='User_Logout.asp'>【退出时记得注销登录】</a></div>" & vbcrlf
	end if
%>
<script language=javascript>
	function openScript(url)
	{
		var Win = window.open(url,"UserControlPad");
	}
	function openScript2(url, width, height)
	{
		var Win = window.open(url,"UserControlPad",'width=' + width + ',height=' + height + ',resizable=1,scrollbars=yes,menubar=yes,status=yes' );
	}
</script>
<%
end sub

'==================================================
'过程名：ShowTopUser
'作  用：显示用户排行，按已发表的文章数排序，若相等，再按注册先后顺序排序
'参  数：UserNum-------显示的用户个数
'==================================================
sub ShowTopUser(UserNum)
	if UserNum<=0 or UserNum>100 then UserNum=10
	dim sqlTopUser,rsTopUser,i
	sqlTopUser="select top " & UserNum & " * from " & db_User_Table & " order by " & db_User_ArticleChecked & " desc," & db_User_ID & " asc"
	set rsTopUser=server.createobject("adodb.recordset")
	rsTopUser.open sqlTopUser,Conn_User,1,1
	if rsTopUser.bof and rsTopUser.eof then
		response.write "没有任何用户"
	else
		response.write "<table width='100%' border='0' cellspacing='0' cellpadding='0'><tr><td align=center>用户名</td><td >文章数</td></tr>"
		for i=1 to rsTopUser.recordcount
			response.write "<tr><td align=center><a href='UserInfo.asp?UserID=" & rsTopUser(db_User_ID) & "'>" & rsTopUser(db_User_Name) &"文集"& "</a></td><td>" & rsTopUser(db_User_ArticleChecked) & "</td></tr>"
			rsTopUser.movenext
		next
		response.write "</table>"
	end if
	set rsTopUser=nothing
end sub
sub ShownewUser(UserNum)
	if UserNum<=0 or UserNum>100 then UserNum=10
	dim sqlnewUser,rsnewUser,i
	sqlnewuser="select top " & UserNum & " * from " & db_User_Table & " order by " & db_User_begindate & " desc," & db_User_ID & " asc"
	set rsnewUser=server.createobject("adodb.recordset")
	rsnewUser.open sqlnewUser,Conn_User,1,1
	if rsnewUser.bof and rsnewUser.eof then
		response.write "没有任何用户"
	else
		response.write "<table width='100%' border='0' cellspacing='0' cellpadding='0'><tr><td align=center>用户名</td><td >加入日期</td></tr>"
		for i=1 to rsnewUser.recordcount
			response.write "<tr><td align=center>" & rsnewUser(db_User_Name) & "</td><td>" & rsnewUser(db_User_begindate) & "</td></tr>"
			rsnewUser.movenext
		next
		response.write "</table>"
	end if
	set rsnewUser=nothing
end sub

'==================================================
'过程名：ShowAllUser
'作  用：分页显示所有用户
'参  数：无
'==================================================
sub ShowAllUser()
	select case OrderType
	case 1
		sqlUser="select * from " & db_User_Table & " order by " & db_User_ArticleChecked & " desc"
	case 2
		sqlUser="select * from " & db_User_Table & " order by " & db_User_RegDate & " desc"
	case 3
		sqlUser="select * from " & db_User_Table & " order by " & db_User_ID & " desc"
	case 4
		sqlUser="select * from " & db_User_Table & " order by diaryNum desc"
	case 5
		sqlUser="select * from " & db_User_Table & " order by diaryVisit desc"
 	case 6
		sqlUser="select * from " & db_User_Table & " order by diaryVisit desc"
'	case 7
'		sqlUser="select * from " & db_User_Table & " order by sum(ArticleComment.Score) desc"
       
        end select
	set rsUser=server.createobject("adodb.recordset")
	rsUser.open sqlUser,Conn_User,1,1
'	if OrderType=7 then 
'	rsUser.open sqlUser,Conn,1,1
'	else
'	rsUser.open sqlUser,Conn_User,1,1
'	end if 
	if rsUser.bof and rsUser.eof then
		totalput=0
		response.write "<br><li>没有任何用户</li>"
	else
		totalput=rsUser.recordcount
		if currentPage=1 then
			call ShowUserList()
		else
			if (currentPage-1)*MaxPerPage<totalPut then
         	   	rsUser.move  (currentPage-1)*MaxPerPage
         		dim bookmark
           		bookmark=rsUser.bookmark
            	call ShowUserList()
        	else
	        	currentPage=1
           		call ShowUserList()
	    	end if
		end if
	end if
	rsUser.close
	set rsUser=nothing
end sub

sub ShowUserList()
	dim i,rsArticleCommentScore,sqlArticleCommentScore
	i=0

	'response.write "<div align='center'><a href='UserList.asp?OrderType=1'>按发表文章数排序</a>&nbsp;&nbsp;&nbsp;&nbsp;<a href='UserList.asp?OrderType=2'>按注册日期排序</a>&nbsp;&nbsp;&nbsp;&nbsp;<a href='UserList.asp?OrderType=3'>按用户ID排序</a>&nbsp;&nbsp;&nbsp;&nbsp;<a href='UserList.asp?OrderType=7'>按文章得分排序</a><br></div>"
		response.write "<div align='center'><a href='UserList.asp?OrderType=1'>按发表文章数排序</a>&nbsp;&nbsp;&nbsp;&nbsp;<a href='UserList.asp?OrderType=2'>按注册日期排序</a>&nbsp;&nbsp;&nbsp;&nbsp;<a href='UserList.asp?OrderType=3'>按用户ID排序</a><br></div>"

	response.write "<table width='100%' border='1' cellpadding='0' cellspacing='0' bordercolor='#808000' style='border-collapse: collapse'><tr align='center'><td width='15%'><font color=red>用户名</font></td><td width='30%'><font color=red>个人简介</font></td><td width='10%'><font color=red>注册日期</font></td><td width='5%'><font color=red>文章</font></td><td width='8%'><font color=red>文章得分</font></td><tr>"
	do while not rsUser.eof
		response.write "<tr onmouseout=""this.style.backgroundColor=''"" onmouseover=""this.style.backgroundColor='#BFDFFF'"">"
		response.write "<td><a href='User_UserInfo.asp?UserID=" & rsUser(db_User_ID) & "'>" & rsUser(db_User_Name) & "</a></td>"

		
		response.write "<td align='top'>&nbsp;&nbsp;"
		if rsUser(db_User_Msn)<>"" then
			response.write rsUser(db_User_Msn)
		else
			response.write "尚无任何简介，请补充……"
		end if
		
		response.write "</td><td align='center'>" & FormatDateTime(rsUser(db_User_RegDate),2) & "</td><td align='right'>" & rsUser(db_User_ArticleChecked) & "</td>"
		'response.write "<td align='right'><a href=diary_index.asp?diaryOwner=" & rsUser("username") & " title=查看该用户的公开日记>" & rsUser("diaryNum") & "</a></td><td align='right'>" & rsUser("diaryVisit") & "</td>"
		set rsArticleCommentScore=server.createobject("adodb.recordset")
		sqlArticleCommentScore="select sum(ArticleComment.Score) from Article,ArticleComment where Article.ArticleID=ArticleComment.ArticleID and Article.Author='" & rsUser(db_User_Name)  & "'"
		rsArticleCommentScore.open sqlArticleCommentScore,Conn,3,1
		Response.Write("<td  align='right'> " & rsArticleCommentScore(0) & "</td>" )
		response.Write("</tr>")
		rsArticleCommentScore.close
		set rsArticleCommentScore=nothing
		rsUser.movenext
		i=i+1
		if i>=MaxPerPage then exit do
	loop
	response.write "</table>"
end sub

'==================================================
'过程名：Admin_ShowAllUser
'作  用：分页显示所有用户
'参  数：无
'==================================================
sub Admin_ShowAllUser()
dim rsOrderBy,sqlOrderBy
if  OrderType = 10 then
  set rsOrderBy=server.createobject("adodb.recordset")
  rsOrderBy.open sqlOrderBy,Conn_User,1,1
  rsOrderBy.open sqlOrderBy,Conn,1,1
  else
	select case OrderType
	case 1
		sqlUser="select * from " & db_User_Table & " order by " & db_User_ArticleChecked & " desc"
	case 2
		sqlUser="select * from " & db_User_Table & " order by " & db_User_RegDate & " desc"
	case 3
		sqlUser="select * from " & db_User_Table & " order by " & db_User_ID & " desc"
	case 4
		sqlUser="select * from " & db_User_Table & " order by diaryNum desc"
	case 5
		sqlUser="select * from " & db_User_Table & " order by diaryVisit desc"
'两课管理员用
 	case 6
		sqlUser="select * from " & db_User_Table & " order by " & db_User_TrueName &  " desc"
 	case 7
		sqlUser="select * from " & db_User_Table & " order by " & db_User_StudentNumber &  " desc"
	case 8
		sqlUser="select * from " & db_User_Table & " order by " & db_User_StudentClass &  " desc"
	case 9
		sqlUser="select * from " & db_User_Table & " order by " & db_User_College &  " desc"

'case 10
		'sqlUser="select * from " & db_User_Table & " order by " & db_User_ArticleChecked &  " desc"
	'case 10
	'sqlUser="select * from [user].[dbo].[user] union select erpid,erpname from [adsfkldfogowerjnokfdslwejhdfsjhk].[dbo].[article]"
       
    end select
	set rsUser=server.createobject("adodb.recordset")
	rsUser.open sqlUser,Conn_User,1,1
	if rsUser.bof and rsUser.eof then
		totalput=0
		response.write "<br><li>没有任何用户</li>"
	else
		totalput=rsUser.recordcount
		if currentPage=1 then
			call Admin_ShowUserList()
		else
			if (currentPage-1)*MaxPerPage<totalPut then
         	   	rsUser.move  (currentPage-1)*MaxPerPage
         		dim bookmark
           		bookmark=rsUser.bookmark
            	call Admin_ShowUserList()
        	else
	        	currentPage=1
           		call Admin_ShowUserList()
	    	end if
		end if
	end if
	rsUser.close
	set rsUser=nothing




end if
end sub

sub Admin_ShowUserList()
	dim i
	i=0

	response.write "<div align='center'><a href='Admin_UserList.asp?OrderType=3'>按用户ID排序</a>&nbsp;<a href='Admin_UserList.asp?OrderType=6'>按学生姓名排序</a>&nbsp;<a href='Admin_UserList.asp?OrderType=7'>按学生学号排序</a>&nbsp;<a href='Admin_UserList.asp?OrderType=8'>按学生所属班级排序</a>&nbsp;<a href='Admin_UserList.asp?OrderType=9'>按学生所属学院排序</a>&nbsp;<a href='Admin_UserList.asp?OrderType=1'>按发表文章数排序</a>&nbsp;<a href='Admin_UserList.asp?OrderType=2'>按注册日期排序</a><br>"
	response.Write()
	response.Write("</div>")
	response.write "<table width='100%' border='1' cellpadding='0' cellspacing='0' bordercolor='#808000' style='border-collapse: collapse'><tr align='center'><td width='15%'><font color=red>用户名</font></td><td width='5%'><font color=red>用户ID</font></td><td width='10%'><font color=red>真实姓名</font></td><td width='6%'><font color=red>学号</font></td><td width='9%'><font color=red>班级</font></td><td width='8%'><font color=red>学院</font></td><td width='8%'><font color=red>注册日期</font></td><td width='7%'><font color=red>已审核文章</font></td><td width='5%'><font color=red>文章得分</font></td><tr>"
	do while not rsUser.eof
		response.write "<tr onmouseout=""this.style.backgroundColor=''"" onmouseover=""this.style.backgroundColor='#BFDFFF'"">"
		response.write "<td><a href='UserInfo.asp?UserID=" & rsUser(db_User_ID) & "'>" & rsUser(db_User_Name) & "</a></td>"
						response.write "<td><a href='UserInfo.asp?UserID=" & rsUser(db_User_ID) & "'>" & rsUser(db_User_ID) & "</a></td>"
		response.Write("<td>" & rsUser(db_User_TrueName) & "</td>" )
		response.Write("<td>" & rsUser(db_User_StudentNumber) & "</td>" )
				response.Write("<td>" & rsUser(db_User_StudentClass) & "</td>" )
						response.Write("<td>" & rsUser(db_User_College) & "</td>" )
'		response.write "<td align='top'>&nbsp;&nbsp;&nbsp;&nbsp;"
'		if rsUser(db_User_Msn)<>"" then
'			response.write rsUser(db_User_Msn)
'		else
'			response.write "尚无任何简介，请补充……"
'		end if
		
		response.write "</td><td align='center'>" & FormatDateTime(rsUser(db_User_RegDate),2) & "</td><td align='right'>" & rsUser(db_User_ArticleChecked) & "</td>"
		'response.write "<td align='right'><a href=diary_index.asp?diaryOwner=" & rsUser("username") & " title=查看该用户的公开日记>" & rsUser("diaryNum") & "</a></td><td align='right'>" & rsUser("diaryVisit") & "</td>"
		set rsArticleCommentScore=server.createobject("adodb.recordset")
		sqlArticleCommentScore="select sum(ArticleComment.Score) from Article,ArticleComment where Article.ArticleID=ArticleComment.ArticleID and Article.Author='" & rsUser(db_User_Name)  & "'"
		rsArticleCommentScore.open sqlArticleCommentScore,Conn,3,1
		Response.Write("<td  align='right'> " & rsArticleCommentScore(0) & "</td>" )
		response.Write("</tr>")
		rsUser.movenext
		i=i+1
		if i>=MaxPerPage then exit do
	loop
	response.write "</table>"
end sub
'end sub Admin_ShowAllUser

'==================================================
'过程名：PopAnnouceWindow
'作  用：弹出公告窗口
'参  数：Width-------弹出窗口宽度
'		 Height------弹出窗口高度
'==================================================
sub PopAnnouceWindow(Width,Height)
	dim popCount,rsAnnounce
	set rsAnnounce=conn.execute("select count(*) from Announce where IsSelected=True and (ChannelID=0 or ChannelID=" & ChannelID & ") and (ShowType=0 or ShowType=2)")
	popCount=rsAnnounce(0)
	if popCount>0 then
		if  PopAnnounce="Yes" and session("Poped")<>ChannelID then
			response.write "<script LANGUAGE='JavaScript'>"
			response.write "window.open ('Announce.asp?ChannelID=" & ChannelID & "', 'newwindow', 'height=" & Height & ", width=" & Width & ", toolbar=no, menubar=no, scrollbars=auto, resizable=no, location=no, status=no')"
			response.write "</script>"
			session("Poped")=ChannelID
		end if
	end if
end sub

'==================================================
'过程名：ShowPath
'作  用：显示“你现在所有位置”导航信息
'参  数：无
'==================================================
sub ShowPath()
	if PageTitle<>"" and ChannelID<>1 then
		strPath=strPath & "&nbsp;&gt;&gt;&nbsp;" & PageTitle
	end if
	response.write strPath
end sub

'==================================================
'过程名：MenuJS
'作  用：生成下拉菜单相关的JS代码
'参  数：无
'==================================================
sub MenuJS()
	dim strMenu
	if ShowMyStyle="Yes" then
%>
<script language="JavaScript" type="text/JavaScript">
//下拉菜单相关代码
 var h;
 var w;
 var l;
 var t;
 var topMar = 1;
 var leftMar = -2;
 var space = 1;
 var isvisible;
 var MENU_SHADOW_COLOR='#999999';//定义下拉菜单阴影色
 var global = window.document
 global.fo_currentMenu = null
 global.fo_shadows = new Array

function HideMenu() 
{
 var mX;
 var mY;
 var vDiv;
 var mDiv;
	if (isvisible == true)
{
		vDiv = document.all("menuDiv");
		mX = window.event.clientX + document.body.scrollLeft;
		mY = window.event.clientY + document.body.scrollTop;
		if ((mX < parseInt(vDiv.style.left)) || (mX > parseInt(vDiv.style.left)+vDiv.offsetWidth) || (mY < parseInt(vDiv.style.top)-h) || (mY > parseInt(vDiv.style.top)+vDiv.offsetHeight)){
			vDiv.style.visibility = "hidden";
			isvisible = false;
		}
}
}

function ShowMenu(vMnuCode,tWidth) {
	vSrc = window.event.srcElement;
	vMnuCode = "<table id='submenu' cellspacing=1 cellpadding=3 style='width:"+tWidth+"' class=menu onmouseout='HideMenu()'><tr height=23><td nowrap align=left class=MenuBody>" + vMnuCode + "</td></tr></table>";

	h = vSrc.offsetHeight;
	w = vSrc.offsetWidth;
	l = vSrc.offsetLeft + leftMar+4;
	t = vSrc.offsetTop + topMar + h + space-2;
	vParent = vSrc.offsetParent;
	while (vParent.tagName.toUpperCase() != "BODY")
	{
		l += vParent.offsetLeft;
		t += vParent.offsetTop;
		vParent = vParent.offsetParent;
	}

	menuDiv.innerHTML = vMnuCode;
	menuDiv.style.top = t;
	menuDiv.style.left = l;
	menuDiv.style.visibility = "visible";
	isvisible = true;
    makeRectangularDropShadow(submenu, MENU_SHADOW_COLOR, 4)
}

function makeRectangularDropShadow(el, color, size)
{
	var i;
	for (i=size; i>0; i--)
	{
		var rect = document.createElement('div');
		var rs = rect.style
		rs.position = 'absolute';
		rs.left = (el.style.posLeft + i) + 'px';
		rs.top = (el.style.posTop + i) + 'px';
		rs.width = el.offsetWidth + 'px';
		rs.height = el.offsetHeight + 'px';
		rs.zIndex = el.style.zIndex - i;
		rs.backgroundColor = color;
		var opacity = 1 - i / (i + 1);
		rs.filter = 'alpha(opacity=' + (100 * opacity) + ')';
		el.insertAdjacentElement('afterEnd', rect);
		global.fo_shadows[global.fo_shadows.length] = rect;
	}
}
</script>
<%
		response.write "<script language='JavaScript' type='text/JavaScript'>" & vbcrlf
		response.write "//菜单列表" & vbcrlf
	
		'自选风格的菜单定义
		strMenu="var menu_skin=" & chr(34)
		dim rsSkin
		set rsSkin=conn.execute("select SkinID,SkinName from Skin")
		do while not rsSkin.eof
			strMenu=strMenu & "&nbsp;<a style=font-size:9pt;line-height:14pt; href='SetCookie.asp?Action=SetSkin&ClassID=" & ClassID & "&SkinID=" & rsSkin(0) & "'>" & rsSkin(1) & "</a><br>"
			rsSkin.movenext
		loop
		rsSkin.close
		set rsSkin=nothing
		response.write strMenu & chr(34) & ";" & vbcrlf
		response.write "</script>" & vbcrlf
	else
	%>
	<script language="JavaScript" type="text/JavaScript">
	function HideMenu() 
	{
	}
	</script>
	<%
	end if
	
	if ChannelID>=2 and ChannelID<=4 then
		'无限级下拉菜单的JS代码文件
		response.write "<script type='text/javascript' language='JavaScript1.2' src='stm31.js'></script>"
		if ShowClassTreeGuide="Yes" then
%>
<script language="JavaScript" type="text/JavaScript">
//树形导航的JS代码
document.write("<style type=text/css>#master {LEFT: -200px; POSITION: absolute; TOP: 25px; VISIBILITY: visible; Z-INDEX: 999}</style>")
document.write("<table id=master width='218' border='0' cellspacing='0' cellpadding='0'><tr><td><img border=0 height=6 src=images/menutop.gif  width=200></td><td rowspan='2' valign='top'><img id=menu onMouseOver=javascript:expand() border=0 height=70 name=menutop src=images/menuo.gif width=18></td></tr>");
document.write("<tr><td valign='top'><table width='100%' border='0' cellspacing='5' cellpadding='0'><tr><td height='400' valign='top'><table width=100% height='100%' border=1 cellpadding=0 cellspacing=5 bordercolor='#666666' bgcolor=#ecf6f5 style=FILTER: alpha(opacity=90)><tr>");
document.write("<td height='10' align='center' bordercolor='#ecf6f5'><font color=999900><strong>栏 目 树 形 导 航</strong></font></td></tr><tr><td valign='top' bordercolor='#ecf6f5'>");
document.write("<iframe width=100% height=100% src='classtree.asp?ChannelID=<%=ChannelID%>' frameborder=0></iframe></td></tr></table></td></tr></table></td></tr></table>");

var ie = document.all ? 1 : 0
var ns = document.layers ? 1 : 0
var master = new Object("element")
master.curLeft = -200;	master.curTop = 10;
master.gapLeft = 0;		master.gapTop = 0;
master.timer = null;

if(ie){var sidemenu = document.all.master;}
if(ns){var sidemenu = document.master;}
setInterval("FixY()",100);

function moveAlong(layerName, paceLeft, paceTop, fromLeft, fromTop){
	clearTimeout(eval(layerName).timer)
	if(eval(layerName).curLeft != fromLeft){
		if((Math.max(eval(layerName).curLeft, fromLeft) - Math.min(eval(layerName).curLeft, fromLeft)) < paceLeft){eval(layerName).curLeft = fromLeft}
		else if(eval(layerName).curLeft < fromLeft){eval(layerName).curLeft = eval(layerName).curLeft + paceLeft}
			else if(eval(layerName).curLeft > fromLeft){eval(layerName).curLeft = eval(layerName).curLeft - paceLeft}
		if(ie){document.all[layerName].style.left = eval(layerName).curLeft}
		if(ns){document[layerName].left = eval(layerName).curLeft}
	}
	if(eval(layerName).curTop != fromTop){
   if((Math.max(eval(layerName).curTop, fromTop) - Math.min(eval(layerName).curTop, fromTop)) < paceTop){eval(layerName).curTop = fromTop}
		else if(eval(layerName).curTop < fromTop){eval(layerName).curTop = eval(layerName).curTop + paceTop}
			else if(eval(layerName).curTop > fromTop){eval(layerName).curTop = eval(layerName).curTop - paceTop}
		if(ie){document.all[layerName].style.top = eval(layerName).curTop}
		if(ns){document[layerName].top = eval(layerName).curTop}
	}
	eval(layerName).timer=setTimeout('moveAlong("'+layerName+'",'+paceLeft+','+paceTop+','+fromLeft+','+fromTop+')',30)
}

function setPace(layerName, fromLeft, fromTop, motionSpeed){
	eval(layerName).gapLeft = (Math.max(eval(layerName).curLeft, fromLeft) - Math.min(eval(layerName).curLeft, fromLeft))/motionSpeed
	eval(layerName).gapTop = (Math.max(eval(layerName).curTop, fromTop) - Math.min(eval(layerName).curTop, fromTop))/motionSpeed
	moveAlong(layerName, eval(layerName).gapLeft, eval(layerName).gapTop, fromLeft, fromTop)
}

var expandState = 0

function expand(){
	if(expandState == 0){setPace("master", 0, 10, 10); if(ie){document.menutop.src = "images/menui.gif"}; expandState = 1;}
	else{setPace("master", -200, 10, 10); if(ie){document.menutop.src = "images/menuo.gif"}; expandState = 0;}
}

function FixY(){
	if(ie){sidemenu.style.top = document.body.scrollTop+10}
	if(ns){sidemenu.top = window.pageYOffset+10}
}
</script>
<%
		end if
	end if
end sub

'==================================================
'过程名：ShowSearchForm
'作  用：显示文章搜索表单
'参  数：ShowType ----显示方式。1为简洁模式，2为标准模式，3为高级模式
'==================================================
sub ShowSearchForm(Action,ShowType)
	if ShowType<>1 and ShowType<>2 and ShowType<>3 then
		ShowType=1
	end if
	response.write "<table border='0' cellpadding='0' cellspacing='0'>"
	response.write "<form method='Get' name='SearchForm' action='" & Action & "'>"
	response.write "<tr><td height='28' align='center'>"
	if ShowType=1 then
		response.write "<input type='text' name='keyword'  size='15' value='关键字' maxlength='50' onFocus='this.select();'>&nbsp;"
		response.write "<input type='hidden' name='field' value='Title'>"
		response.write "<input type='submit' name='Submit'  value='搜索'>"
		'response.write "<br><br>高级搜索"
	elseif Showtype=2 then
		response.write "<select name='Field' size='1'>"
    	if ChannelID=2 then
			response.write "<option value='Title' selected>文章标题</option>"
			response.write "<option value='Content'>文章内容</option>"
			response.write "<option value='Author'>文章作者</option>"
			response.write "<option value='Editor'>编辑姓名</option>"
		elseif ChannelID=3 then	
			response.write "<option value='SoftName' selected>软件名称</option>"
			response.write "<option value='SoftIntro'>软件简介</option>"
			response.write "<option value='Author'>软件作者</option>"
			response.write "<option value='Editor'>编辑姓名</option>"
		elseif ChannelID=4 then	
			response.write "<option value='PhotoName' selected>图片名称</option>"
			response.write "<option value='PhotoIntro'>图片简介</option>"
			response.write "<option value='Author'>图片作者</option>"
			response.write "<option value='Editor'>编辑姓名</option>"
		else
			response.write "<option value='Title' selected>文章标题</option>"
			response.write "<option value='Content'>文章内容</option>"
			response.write "<option value='Author'>文章作者</option>"
			response.write "<option value='Editor'>编辑姓名</option>"
		end if
		response.write "</select>&nbsp;"
		response.write "<select name='ClassID'><option value=''>所有栏目</option>"
		call Admin_ShowClass_Option(5,0)
		response.write "</select>&nbsp;<input type='text' name='keyword'  size='20' value='关键字' maxlength='50' onFocus='this.select();'>&nbsp;"
		response.write "<input type='submit' name='Submit'  value=' 搜索 '>"
	elseif Showtype=3 then
	
	end if
	response.write "</td></tr></form></table>"
end sub

'==================================================
'过程名：ShowGuest
'作  用：显示网站留言
'参  数：GuestTitleLen ---显示留言标题长度
'		 GuestItemNum  ---显示留言条数
'==================================================
sub ShowGuest(GuestTitleLen,GuestItemNum)
 	dim sqlGuest,rsGuest
 	if GuestItemNum<=0 or GuestItemNum>50 then
 		GuestItemNum=10
	end if
 	sqlGuest="select top " & GuestItemNum & " * from Guest where GuestIsPassed=True order by GuestMaxId desc"
 	Set rsGuest= Server.CreateObject("ADODB.Recordset")
 	rsGuest.open sqlGuest,conn,1,1
 	if rsGuest.bof and rsGuest.eof then 
  		response.Write " 没有任何留言"
 	else
		do while Not rsGuest.eof
			response.write "<font color=#b70000><b>&nbsp;&nbsp;&nbsp;·</b></font><a href='guestbook.asp' "
			response.write " title='主题：" & rsGuest("GuestTitle") & vbcrlf & "姓名：" & rsGuest("GuestName") & vbcrlf & "时间：" & rsGuest("GuestDatetime") &"'"
			response.write " target='_blank'>"
			response.write gotTopic(rsGuest("GuestTitle"),GuestTitleLen)
			response.write "</a><br>"
			rsGuest.movenext
  		Loop
 	end if
 	rsGuest.close
 	set rsGuest=nothing
end sub

'==================================================
'过程名：ShowAD
'作  用：显示广告
'参  数：ADType ---广告类型
'==================================================
sub ShowAD(ADType)
	dim sqlAD,rsAD,AD,arrSetting,popleft,poptop,floatleft,floattop,fixedleft,fixedtop
	sqlAD="select * from Advertisement where IsSelected=True"
	sqlAD=sqlAD & " and (ChannelID=0 or ChannelID=" & ChannelID & ")"
	sqlAD=sqlAD & " and ADType=" & ADtype & " order by ID Desc"
	set rsAD=server.createobject("adodb.recordset")
	rsAD.open sqlAD,conn,1,1
	if not rsAd.bof and not rsAD.eof then
		do while not rsAD.eof
			if rsAD("isflash")=true then
				AD= "<object classid='clsid:D27CDB6E-AE6D-11cf-96B8-444553540000' codebase='http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=5,0,0,0'"
				if rsAD("ImgWidth")>0 then AD = AD & " width='" & rsAD("ImgWidth") & "'"
				if rsAD("ImgHeight")>0 then AD = AD & " height='" & rsAD("ImgHeight") & "'"
				AD = AD & "><param name='movie' value='" & rsAD("ImgUrl") & "'><param name='quality' value='high'><embed src='" & rsAD("ImgUrl") & "' pluginspage='http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash' type='application/x-shockwave-flash'"
				if rsAD("ImgWidth")>0 then AD = AD & " width='" & rsAD("ImgWidth") & "'"
				if rsAD("ImgHeight")>0 then AD = AD & " height='" & rsAD("ImgHeight") & "'"
				AD = AD & "></embed></object>"
			else
				AD ="<a href='" & rsAD("SiteUrl") & "' target='_blank' title='" & rsAD("SiteName") & "：" & rsAD("SiteUrl") & "'><img src='" & rsAD("ImgUrl") & "'"
				if rsAD("ImgWidth")>0 then AD = AD & " width='" & rsAD("ImgWidth") & "'"
				if rsAD("ImgHeight")>0 then AD = AD & " height='" & rsAD("ImgHeight") & "'"
				AD = AD & " border='0'></a>"
			end if
			if ADtype=0 then
				if  session("PopAD"&rsAD("ID")&ChannelID)<>True then
					if instr(rsAD("ADSetting"),"|")>0 then
						arrSetting=split(rsAD("ADSetting"),"|")
						popleft=arrsetting(0)
						poptop=arrsetting(1)
					end if
					response.write "<SCRIPT language=javascript>"
					response.write "window.open(""PopAD.asp?Id="& rsAD("ID")&""",""popad"&rsAD("ID")&""",""toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=no,width="&rsAD("ImgWidth")&",height="&rsAD("ImgHeight")&",top="&poptop&",left="&popleft&""");"
					response.write "</SCRIPT>"
					session("PopAD"&rsAD("ID")&ChannelID)=True
				end if
			elseif ADtype=1 then
				response.write AD
				exit do
			elseif ADtype=2 then
				response.write AD
				exit do
			elseif ADtype=3 then
				response.write AD
				exit do
			elseif ADtype=4 then
				if instr(rsAD("ADSetting"),"|")>0 then
					arrSetting=split(rsAD("ADSetting"),"|")
					floatleft=arrsetting(0)
					floattop=arrsetting(1)
				end if
				response.write "<div id='FlAD' style='position:absolute; z-index:10;left: "&floatleft&"; top: "&floattop&"'>" & AD & "</div>"
				call FloatAD()
				exit do
			elseif ADtype=5 then
				if instr(rsAD("ADSetting"),"|")>0 then
					arrSetting=split(rsAD("ADSetting"),"|")
					fixedleft=arrsetting(0)
					fixedtop=arrsetting(1)
				end if
				response.write "<div id='FixAD' style='position:absolute; z-index:10;left: "&fixedleft&"; top: "&fixedtop&"'>" & AD & "</div>"
				call FixedAD()
				exit do
			elseif ADtype=6 then
				response.write rsAD("ImgUrl")
				exit do
			end if
			rsAD.movenext
		loop
	end if
	rsAD.close
	set rsAD=nothing
end sub

'==================================================
'过程名：FloatAD
'作  用：浮动广告
'参  数：无
'==================================================
sub FloatAD()
%>
<SCRIPT language=javascript>
<!--moving logo-->
window.onload=FlAD;
var brOK=false;
var mie=false;
var aver=parseInt(navigator.appVersion.substring(0,1));
var aname=navigator.appName;
var mystop=0;

function checkbrOK()
{if(aname.indexOf("Internet Explorer")!=-1)
{if(aver>=4) brOK=navigator.javaEnabled();
mie=true;
}
if(aname.indexOf("Netscape")!=-1)  
{if(aver>=4) brOK=navigator.javaEnabled();}
}
var vmin=2;
var vmax=5;
var vr=2;
var timer1;

function Chip(chipname,width,height)
{this.named=chipname;
this.vx=vmin+vmax*Math.random();
this.vy=vmin+vmax*Math.random();
this.w=width;
this.h=height;
this.xx=0;
this.yy=0;
this.timer1=null;
}

function movechip(chipname)
{
if(brOK && mystop==0)
{eval("chip="+chipname);
if(!mie)
{pageX=window.pageXOffset;
pageW=window.innerWidth;
pageY=window.pageYOffset;
pageH=window.innerHeight;
}
else
{pageX=window.document.body.scrollLeft;
pageW=window.document.body.offsetWidth-8;
pageY=window.document.body.scrollTop;
pageH=window.document.body.offsetHeight;
} 
chip.xx=chip.xx+chip.vx;
chip.yy=chip.yy+chip.vy;
chip.vx+=vr*(Math.random()-0.5);
chip.vy+=vr*(Math.random()-0.5);
if(chip.vx>(vmax+vmin))  chip.vx=(vmax+vmin)*2-chip.vx;
if(chip.vx<(-vmax-vmin)) chip.vx=(-vmax-vmin)*2-chip.vx;
if(chip.vy>(vmax+vmin))  chip.vy=(vmax+vmin)*2-chip.vy;
if(chip.vy<(-vmax-vmin)) chip.vy=(-vmax-vmin)*2-chip.vy;
if(chip.xx<=pageX)
{chip.xx=pageX;
chip.vx=vmin+vmax*Math.random();
}
if(chip.xx>=pageX+pageW-chip.w)
{chip.xx=pageX+pageW-chip.w;
chip.vx=-vmin-vmax*Math.random();
}
if(chip.yy<=pageY)
{chip.yy=pageY;
chip.vy=vmin+vmax*Math.random();
}
if(chip.yy>=pageY+pageH-chip.h)
{chip.yy=pageY+pageH-chip.h;
chip.vy=-vmin-vmax*Math.random();
}
if(!mie)
{eval('document.'+chip.named+'.top ='+chip.yy);
eval('document.'+chip.named+'.left='+chip.xx);
} 
else
{eval('document.all.'+chip.named+'.style.pixelLeft='+chip.xx);
eval('document.all.'+chip.named+'.style.pixelTop ='+chip.yy); 
}
	chip.timer1=setTimeout("movechip('"+chip.named+"')",100);
}
}
function stopme(x)
{
brOk=true;
mystop=x;
movechip("FlAD");
}
var FlAD;
var chip;
function FlAD()
{checkbrOK(); 
FlAD=new Chip("FlAD",80,80);
if(brOK) 
{ movechip("FlAD");
}
}
ns4=(document.layers)?true:false;
ie4=(document.all)?true:false;

function cncover()
{
if(ns4){
	//document.cnc.left=window.innerWidth/2-400;
	document.FlAD.visibility="hide";
	eval('document.cnc.left=document.'+chip.named+'.left');
	eval('document.cnc.top=document.'+chip.named+'.top-15');	
	document.cnc.visibility="show";
	}else if(ie4) 
	{
	document.all.FlAD.style.visibility="hidden";
	//document.all.cnc.style.left=window.document.body.offsetWidth/2-400;
	document.all.cnc.style.left=parseInt(document.all.FlAD.style.left)-0;
	document.all.cnc.style.top=parseInt(document.all.FlAD.style.top)-0;	
	document.all.cnc.style.visibility="visible";
	stopme(1);
	}
}

function cncout()
{
if(ns4){
	document.cnc.visibility="hide";
	document.FlAD.visibility="show";
	}else if(ie4) 
	{
	document.all.cnc.style.visibility="hidden";
	document.all.FlAD.style.visibility="visible";
	stopme(0);
	}
}
</script>
<%
end sub


'==================================================
'过程名：FixedAD
'作  用：固定位置广告
'参  数：无
'==================================================
sub FixedAD()
%>
<script LANGUAGE="JavaScript">
<!-- Begin
var imgheight
var imgleft
document.ns = navigator.appName == "Netscape"
if (navigator.appName == "Netscape")
{
imgheight=document.FixAD.pageY
imgleft=document.FixAD.pageX
}
else
{
imgheight=600-parseInt(FixAD.style.top)
imgleft=parseInt(FixAD.style.left)
}
myload()
function myload()
{
if (navigator.appName == "Netscape")
{document.FixAD.pageY=pageYOffset+window.innerHeight-imgheight;
document.FixAD.pageX=imgleft;
leftmove();
}
else
{
FixAD.style.top=document.body.scrollTop+document.body.offsetHeight-imgheight;
FixAD.style.left=imgleft;
leftmove();
}
}
function leftmove()
 {
 if(document.ns)
 {
 document.FixAD.top=pageYOffset+window.innerHeight-imgheight
 document.FixAD.left=imgleft;
 setTimeout("leftmove();",50)
 }
 else
 {
 FixAD.style.top=document.body.scrollTop+document.body.offsetHeight-imgheight;
 FixAD.style.left=imgleft;
 setTimeout("leftmove();",50)
 }
 }

function MM_reloadPage(init) {  //reloads the window if Nav4 resized
  if (init==true) with (navigator) {if ((appName=="Netscape")&&(parseInt(appVersion)==4)) {
    document.MM_pgW=innerWidth; document.MM_pgH=innerHeight; onresize=MM_reloadPage; }}
  else if (innerWidth!=document.MM_pgW || innerHeight!=document.MM_pgH) location.reload();
}
MM_reloadPage(true)
//  End -->
</script>
<%
end sub

'==================================================
'过程名：FixedAD1
'作  用：固定位置广告（图片位置超过窗口时卷动时有问题）
'参  数：无
'==================================================
sub FixedAD1()
%>
<script LANGUAGE="JavaScript">
<!-- Begin
        self.onError=null;
        currentX = currentY = 0;
        whichIt = null;
        lastScrollX = 0; lastScrollY = 0;
        NS = (document.layers) ? 1 : 0;
        IE = (document.all) ? 1: 0;
        function heartBeat() {
                if(IE) { diffY = document.body.scrollTop; diffX = document.body.scrollLeft; }
            if(NS) { diffY = self.pageYOffset; diffX = self.pageXOffset; }
                if(diffY != lastScrollY) {
                        percent = .1 * (diffY - lastScrollY);
                        if(percent > 0) percent = Math.ceil(percent);
                        else percent = Math.floor(percent);
                                        if(IE) document.all.FixAD.style.pixelTop += percent;
                                        if(NS) document.FixAD.top += percent;
                        lastScrollY = lastScrollY + percent;
            }
                if(diffX != lastScrollX) {
                        percent = .1 * (diffX - lastScrollX);
                        if(percent > 0) percent = Math.ceil(percent);
                        else percent = Math.floor(percent);
                        if(IE) document.all.FixAD.style.pixelLeft += percent;
                        if(NS) document.FixAD.left += percent;
                        lastScrollX = lastScrollX + percent;
                }
        }
        function checkFocus(x,y) {
                stalkerx = document.FixAD.pageX;
                stalkery = document.FixAD.pageY;
                stalkerwidth = document.FixAD.clip.width;
                stalkerheight = document.FixAD.clip.height;
                if( (x > stalkerx && x < (stalkerx+stalkerwidth)) && (y > stalkery && y < (stalkery+stalkerheight))) return true;
                else return false;
        }
        function grabIt(e) {
                if(IE) {
                        whichIt = event.srcElement;
                        while (whichIt.id.indexOf("FixAD") == -1) {
                                whichIt = whichIt.parentElement;
                                if (whichIt == null) { return true; }
                    }
                        whichIt.style.pixelLeft = whichIt.offsetLeft;
                    whichIt.style.pixelTop = whichIt.offsetTop;
                        currentX = (event.clientX + document.body.scrollLeft);
                           currentY = (event.clientY + document.body.scrollTop);
                } else {
                window.captureEvents(Event.MOUSEMOVE);
                if(checkFocus (e.pageX,e.pageY)) {
                        whichIt = document.FixAD;
                        StalkerTouchedX = e.pageX-document.FixAD.pageX;
                        StalkerTouchedY = e.pageY-document.FixAD.pageY;
                }
                }
            return true;
        }
        function moveIt(e) {
                if (whichIt == null) { return false; }
                if(IE) {
                    newX = (event.clientX + document.body.scrollLeft);
                    newY = (event.clientY + document.body.scrollTop);
                    distanceX = (newX - currentX);    distanceY = (newY - currentY);
                    currentX = newX;    currentY = newY;
                    whichIt.style.pixelLeft += distanceX;
                    whichIt.style.pixelTop += distanceY;
                        if(whichIt.style.pixelTop < document.body.scrollTop) whichIt.style.pixelTop = document.body.scrollTop;
                        if(whichIt.style.pixelLeft < document.body.scrollLeft) whichIt.style.pixelLeft = document.body.scrollLeft;
                        if(whichIt.style.pixelLeft > document.body.offsetWidth - document.body.scrollLeft - whichIt.style.pixelWidth - 20) whichIt.style.pixelLeft = document.body.offsetWidth - whichIt.style.pixelWidth - 20;
                        if(whichIt.style.pixelTop > document.body.offsetHeight + document.body.scrollTop - whichIt.style.pixelHeight - 5) whichIt.style.pixelTop = document.body.offsetHeight + document.body.scrollTop - whichIt.style.pixelHeight - 5;
                        event.returnValue = false;
                } else {
                        whichIt.moveTo(e.pageX-StalkerTouchedX,e.pageY-StalkerTouchedY);
                if(whichIt.left < 0+self.pageXOffset) whichIt.left = 0+self.pageXOffset;
                if(whichIt.top < 0+self.pageYOffset) whichIt.top = 0+self.pageYOffset;
                if( (whichIt.left + whichIt.clip.width) >= (window.innerWidth+self.pageXOffset-17)) whichIt.left = ((window.innerWidth+self.pageXOffset)-whichIt.clip.width)-17;
                if( (whichIt.top + whichIt.clip.height) >= (window.innerHeight+self.pageYOffset-17)) whichIt.top = ((window.innerHeight+self.pageYOffset)-whichIt.clip.height)-17;
                return false;
                }
            return false;
        }
        function dropIt() {
                whichIt = null;
            if(NS) window.releaseEvents (Event.MOUSEMOVE);
            return true;
        }
        if(NS) {
                window.captureEvents(Event.MOUSEUP|Event.MOUSEDOWN);
                window.onmousedown = grabIt;
                 window.onmousemove = moveIt;
                window.onmouseup = dropIt;
        }
        if(IE) {
                document.onmousedown = grabIt;
                 document.onmousemove = moveIt;
                document.onmouseup = dropIt;
        }
        if(NS || IE) action = window.setInterval("heartBeat()",1);
//  End -->
</script>
<%
end sub
%>