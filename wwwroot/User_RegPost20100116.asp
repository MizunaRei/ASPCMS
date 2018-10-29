<!--#include file="Inc/syscode_article.asp"-->
<!--#include file="Inc/md5.asp"-->
<!--#include file="Inc/RegBBS.asp"-->
<%
if EnableUserReg<>"Yes" then
	FoundErr=true
	ErrMsg=ErrMsg & "<br><li>对不起，本站暂停新用户注册服务！</li>"
	call WriteErrMsg()
	response.end
end if

const ChannelID=2
Const ShowRunTime="Yes"
SkinID=0

dim RegUserName,Password,PwdConfirm,Question,Answer,Sex,Email,Homepage,QQ,MSN,TrueName,StudentNumber,College,StudentClass',StudentClassName,StudentClassYear,StudentClassNumber
RegUserName=trim(request("UserName"))

'人文学院两课教改网站新增的数据变量名
TrueName=Trim(Request("TrueName"))
StudentNumber=Trim(Request("StudentNumber"))
College=Trim(Request("College"))
StudentClass=Trim(Request("StudentClassName")) & Trim(Request("StudentClassYear"))  &  Trim(Request("StudentClassNumber"))
StudentClassName=Trim(Request("StudentClassName"))
StudentClassYear=Trim(Request("StudentClassYear"))
StudentClassNumber=Trim(Request("StudentClassNumber"))
'以上是人文学院两课教改网站新增的数据变量名

Password=trim(request("Password"))
PwdConfirm=trim(request("PwdConfirm"))
Question=trim(request("Question"))
Answer=trim(request("Answer"))
Sex=trim(Request("Sex"))
Email=trim(request("Email"))
Homepage=trim(request("Homepage"))
if Homepage="http://" or isnull(Homepage) then Homepage=""
QQ=trim(request("QQ"))
MSN=trim(request("MSN"))
dim CheckNum,CheckUrl
randomize
CheckNum = int(7999*rnd+2000) '随机验证码
CheckUrl=Request.ServerVariables("HTTP_REFERER")
CheckUrl=left(CheckUrl,instrrev(CheckUrl,"/")) & "User_RegCheck.asp?Action=Check&UserName=" & RegUserName & "&Password=" & Password & "&CheckNum=" & CheckNum
if RegUserName="" or strLength(RegUserName)>14 or strLength(RegUserName)<4 then
	founderr=true
	errmsg=errmsg & "<br><li>请输入用户名(不能大于14小于4)</li>"
else
  	if Instr(RegUserName,"=")>0 or Instr(RegUserName,"%")>0 or Instr(RegUserName,chr(32))>0 or Instr(RegUserName,"?")>0 or Instr(RegUserName,"&")>0 or Instr(RegUserName,";")>0 or Instr(RegUserName,",")>0 or Instr(RegUserName,"'")>0 or Instr(RegUserName,",")>0 or Instr(RegUserName,chr(34))>0 or Instr(RegUserName,chr(9))>0 or Instr(RegUserName,"")>0 or Instr(RegUserName,"$")>0 then
		errmsg=errmsg+"<br><li>用户名中含有非法字符</li>"
		founderr=true
	end if
end if
if Password="" or strLength(Password)>12 or strLength(Password)<6 then
	founderr=true
	errmsg=errmsg & "<br><li>请输入密码(不能大于12小于6)</li>"
else
	if Instr(Password,"=")>0 or Instr(Password,"%")>0 or Instr(Password,chr(32))>0 or Instr(Password,"?")>0 or Instr(Password,"&")>0 or Instr(Password,";")>0 or Instr(Password,",")>0 or Instr(Password,"'")>0 or Instr(Password,",")>0 or Instr(Password,chr(34))>0 or Instr(Password,chr(9))>0 or Instr(Password,"")>0 or Instr(Password,"$")>0 then
		errmsg=errmsg+"<br><li>密码中含有非法字符</li>"
		founderr=true
	end if
end if
if PwdConfirm="" then
	founderr=true
	errmsg=errmsg & "<br><li>请输入确认密码(不能长于12短于6)</li>"
else
	if Password<>PwdConfirm then
		founderr=true
		errmsg=errmsg & "<br><li>密码和确认密码不一致</li>"
	end if
end if
if Question="" then
	founderr=true
	errmsg=errmsg & "<br><li>密码提示问题不能为空</li>"
end if
if Answer="" then
	founderr=true
	errmsg=errmsg & "<br><li>密码答案不能为空</li>"
end if
'校验学生真实姓名
if TrueName="" or strLength(TrueName)>14 or strLength(TrueName)<4 then
	founderr=true
	errmsg=errmsg & "<br><li>请输入真实姓名(不能大于14或小于4个半角字符)</li>"
else
  	if Instr(TrueName,"=")>0 or Instr(TrueName,"%")>0 or Instr(TrueName,chr(32))>0 or Instr(TrueName,"?")>0 or Instr(TrueName,"&")>0 or Instr(TrueName,";")>0 or Instr(TrueName,",")>0 or Instr(TrueName,"'")>0 or Instr(TrueName,",")>0 or Instr(TrueName,chr(34))>0 or Instr(TrueName,chr(9))>0 or Instr(TrueName,"")>0 or Instr(TrueName,"$")>0 then
		errmsg=errmsg+"<br><li>真实姓名中含有非法字符</li>"
		founderr=true
	end if
end if
'结束校验学生真实姓名

'校验学号
if StudentNumber<>"" then
	if not isnumeric(StudentNumber) or len(cstr(StudentNumber))>12 then
		errmsg=errmsg & "<br><li>学号只能是4-12位数字，请您输入。</li>"
		founderr=true
	end if
end if
'结束校验学号

'校验学院
if College="" then 
	founderr=true
	errmsg=errmsg & "<br><li>必须选择所属学院</li>"
end if
'结束校验学院

'校验班级
if StudentClassName=""  or StudentClassYear="" or StudentClassNumber=""  then 
	founderr=true
	errmsg=errmsg & "<br><li>必须选择所属班级</li>"
end if
'结束校验班级



if Sex="" then
	founderr=true
	errmsg=errmsg & "<br><li>性别不能为空</li>"
else
	sex=cint(sex)
	if Sex<>0 and Sex<>1 then
		Sex=1
	end if
end if


if Email="" then
	founderr=true
	errmsg=errmsg & "<br><li>Email不能为空</li>"
else
	if IsValidEmail(Email)=false then
		errmsg=errmsg & "<br><li>您的Email地址格式有错误</li>"
   		founderr=true
	end if
end if
if QQ<>"" then
	if not isnumeric(QQ) or len(cstr(QQ))>10 then
		errmsg=errmsg & "<br><li>QQ号码只能是4-10位数字，您可以选择不输入。</li>"
		founderr=true
	end if
end if


if founderr=false then
	dim sqlReg,rsReg,sqladminname,rsadminname

	sqladminname="select * from admin where UserName='" & RegUserName & "'"
	set rsadminname=server.createobject("adodb.recordset")
	rsadminname.open sqladminname,Conn,1,3
	if not(rsadminname.bof and rsadminname.eof) then
		founderr=true
		errmsg=errmsg & "<br><li>这是管理员的用户名！请换一个用户名再尝试注册！</li>"
	else

	sqlReg="select * from " & db_User_Table & " where " & db_User_Name & "='" & RegUserName & "'"
	set rsReg=server.createobject("adodb.recordset")
	rsReg.open sqlReg,Conn_User,1,3
	if not(rsReg.bof and rsReg.eof) then
		founderr=true
		errmsg=errmsg & "<br><li>你注册的用户名已经存在！请换一个用户名再尝试注册！</li>"
	else
		rsReg.addnew
		rsReg(db_User_Name)=RegUserName
		rsReg(db_User_Password)=md5(Password)
		rsReg(db_User_Question)=Question
		rsReg(db_User_Answer)=md5(Answer)
		
		
		'人文学院两课教改网站新增的字段名
		rsReg(db_User_StudentNumber)=StudentNumber
		rsReg(db_User_TrueName)=TrueName
		rsReg(db_User_StudentClass)=StudentClass
		rsReg(db_User_College)=College
		rsReg("StudentClassName")=Trim(Request("StudentClassName"))
		rsReg("StudentClassYear")=Trim(Request("StudentClassYear"))
		rsReg("StudentClassNumber")=Trim(Request("StudentClassNumber"))
'		StudentClass=Trim(Request("StudentClassName")) & Trim(Request("StudentClassYear"))  &  Trim(Request("StudentClassNumber"))
		
		'以上是人文学院两课教改网站新增的字段名
		
		
		rsReg(db_User_Sex)=Sex
		rsReg(db_User_Email)=Email
		rsReg(db_User_Homepage)=Homepage
		rsReg(db_User_QQ)=QQ
		rsReg(db_User_Msn)=MSN
		rsReg(db_User_RegDate)=Now()
		rsReg(db_User_ArticleCount)=0
		rsReg(db_User_ArticleChecked)=0
		rsReg(db_User_LoginTimes)=1
		rsReg(db_User_LastLoginTime)=NOW()
		rsReg(db_User_ChargeType)=ChargeType_999
		rsReg(db_User_UserPoint)=UserPoint_999
		rsReg(db_User_BeginDate)=formatdatetime(now(),2)
		rsReg(db_User_Valid_Num)=ValidDays_999
		rsReg(db_User_Valid_Unit)=1
		if EmailCheckReg="Yes" then
			rsReg(db_User_UserLevel)=3000
			call SendRegEmail()
		else
			if AdminCheckReg="Yes" then
				rsReg(db_User_UserLevel)=2000
			else			
				rsReg(db_User_UserLevel)=999
				Response.Cookies("asp163")("UserName")=RegUserName
				Response.Cookies("asp163")("Password") =md5(Password)
				Response.Cookies("asp163")("UserLevel")=999
			end if
		end if		

		if UserTableType="Dvbbs6.0" or UserTableType="Dvbbs6.1" then
			rsReg(db_User_UserClass) = FU_UserClass
			rsReg(db_User_TitlePic) = FU_TitlePic
			rsReg(db_User_UserGroupID) = FU_UserGroupID
			rsReg(db_User_Face) = FU_Face
			rsReg(db_User_FaceWidth) = FU_FaceWidth
			rsReg(db_User_FaceHeight) = FU_FaceHeight
			rsReg(db_User_UserWealth) = FU_UserWealth
			rsReg(db_User_UserEP) = FU_UserEP
			rsReg(db_User_UserCP) = FU_UserCP
			rsReg(db_User_UserGroup) = FU_UserGroup
			rsReg(db_User_Showre) = FU_Showre
		end if 

		rsReg.update
		call UpdateUserNum(RegUserName)
	end if
	rsReg.close
	set rsReg=nothing
	end if
	rsadminname.close
	set rsadminname=nothing
end if		
PageTitle="注册成功"
%>
<html>
<head>
<title><%=strPageTitle & " >> " & PageTitle%></title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<!--#include file="inc/Skin_CSS.asp"-->
<% call MenuJS() %>
</head>

<body leftmargin="2" topmargin="0" marginwidth="0" marginheight="0"    bgcolor=#ffffff style="BACKGROUND-COLOR: #ffffff" >
<table width="989"><!--the great talbe--><tr><td>
<%	call Top_noIndex() %></td></tr>
<tr><td>
<table width="989" height="300" border="0" align="center" cellpadding="0" cellspacing="0" class="border2">
  <tr>
    <td width="180" valign="top" class="tdbg_leftall"><TABLE cellSpacing=0 cellPadding=0 width="100%" border="0" style="word-break:break-all">
        <TR class="title_left"> 
          <TD align="center"><table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr> 
                <td class="title_lefttxt"><div align="center"><b>·注册<%=SiteName%></b></div></td>
              </tr>
            </table></TD>
        </TR>
        <TR> 
          <TD height="80" valign="top" class="tdbg_left"> <table width="100%" height="100%" border="0" cellpadding="0" cellspacing="5">
              <tr> 
                <td valign="top"><br> <b>&nbsp;&nbsp;注册步骤</b><br> &nbsp;&nbsp;一、阅读并同意协议<font color="#FF0000">√</font><br> 
                  &nbsp;&nbsp;二、填写注册资料<font color="#FF0000">√</font><br> &nbsp;&nbsp;三、完成注册<font color="#FF0000">→</font></td>
              </tr>
            </table></TD>
        </TR>
      </table>
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td class="tdbg_left"><table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
              <tr> 
                <td height="11" Class="title_left2"></td>
              </tr>
            </table></td>
        </tr>
      </table>
      
    </td>
    <td width=5></td>
    <td width="800" align="center" valign=top><table width="100%" height="280" border="0" cellpadding="0" cellspacing="0" class="border">
        <tr>
          <td> <div align="center">
              <%
			if founderr=false then
				call RegSuccess()
			else
				call WriteErrmsg()
			end if
			%>
              <br>
              <br>
            </div></td>
        </tr>
      </table> 
      
      <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
        <tr> 
          <td  height="15" align="center" valign="top"> <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
              <tr> 
                <td height="13" Class="tdbg_left2"></td>
              </tr>
            </table></td>
        </tr>
      </table></td>
  </tr>
</table></td></tr>


<tr><td>    <!--Bottom--><%  call Bottom_All()  %></td></tr>
    </table><!--the great talbe-->
</body>
</html>
<%
call CloseConn
call CloseConn_User

sub WriteErrMsg()
    response.write "<br><br><table align='center' width='300' border='0' cellpadding='2' cellspacing='0' class='border'>"
    response.write "<tr class='title'><td align='center' height='22'>由于以下的原因不能注册用户！</td></tr>"
    response.write "<tr class='tdbg'><td align='left' height='100'><br>" & errmsg & "<p align='center'>【<a href='javascript:onclick=history.go(-1)'>返 回</a>】<br></p></td></tr>"
	response.write "</table>" 
end sub

sub RegSuccess()
    response.write "<br><br><table align='center' width='300' border='0' cellpadding='2' cellspacing='0' class='border'>"
    response.write "<tr class='title'><td align='center' height='22'>成功注册用户！</td></tr>"
    response.write "<tr class='tdbg'><td align='left' height='100'><br>你注册的用户名：" & RegUserName & "<br>"
	if EmailCheckReg="Yes" then
		response.write "系统已经发送了一封确认信到你注册时填写的信箱中，你必须在收到确认信并通过确认信中链接进行确认后，你才能正式成为本站的注册用户。"
	else
		if AdminCheckReg="Yes" then
			response.write "请等待管理通过你的注册申请后，你就可以正式成为本站的注册用户了。"
		else			
			response.write "欢迎您的加入！！！"
		end if
	end if		
	response.write "<p align='center'>【<a href='javascript:onclick=window.close()'>关 闭</a>】<br></p></td></tr>"
	response.write "</table>" 
end sub

sub SendRegEmail()
	dim MailtoAddress,MailtoName,Subject,MailBody,FromName,MailFrom,Priority
	MailtoAddress=Email
	MailtoName=RegUserName
	Subject="注册确认信"
	MailBody="这是一封注册确认信。你的验证码是：" & CheckNum & vbcrlf & "<br>请点此进行确认：<a href='" & CheckUrl & "'>" & CheckUrl & "</a>"
	FromName=SiteName
	MailFrom=WebmasterEmail
	Priority=3
	ErrMsg=SendMail(MailtoAddress,MailtoName,Subject,MailBody,FromName,MailFrom,Priority)
	if ErrMsg<>"" then FoundErr=True
end sub
%>