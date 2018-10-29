<!--#include file="inc/Conn_User.asp"-->
<!--#include file="inc/md5.asp"-->
<!--#include file="inc/function.asp"-->
<%
if CheckUserLogined()=False then
	response.Redirect "User_Login.asp"
end if

dim Action,FoundErr,ErrMsg
dim rsUser,sqlUser
Action=trim(request("Action"))
if Action="Modify" then
	UserName=trim(request("UserName"))
else
	UserName=Trim(Request.Cookies("asp163")("UserName"))
end if
if  UserName="" then
	FoundErr=True
	ErrMsg=ErrMsg & "<br><li>参数不足！</li>"
end if
if FoundErr=true then
	call WriteErrMsg()
else
	Set rsUser=Server.CreateObject("Adodb.RecordSet")
	sqlUser="select * from " & db_User_Table & " where " & db_User_Name & "='" & UserName & "'"
	rsUser.Open sqlUser,Conn_User,1,3
	if rsUser.bof and rsUser.eof then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>找不到指定的用户！</li>"
		call writeErrMsg()
	else
		if Action="Modify" then
			dim Sex,Email,Homepage,Company,Department,jszc
			Sex=trim(Request("Sex"))
			Email=trim(request("Email"))
			Homepage=trim(request("Homepage"))
			OICQ=trim(request("OICQ"))
			MSN=trim(request("MSN"))
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
					errmsg=errmsg & "<br><li>您的Email有错误</li>"
			   		founderr=true
				end if
			end if
			if FoundErr<>true then
				rsUser(db_User_Sex)=Sex
				rsUser(db_User_Email)=Email
				rsUser(db_User_Homepage)=HomePage
				rsUser(db_User_QQ)=OICQ
				rsUser(db_User_Msn)=MSN
				'两课网站代码
				
				rsUser("TrueName")=Trim(Request("TrueName"))
				rsUser("StudentClass")=Trim(Request("StudentClassName")) & Trim(Request("StudentClassYear")) & Trim(Request("StudentClassNumber"))
				rsUser("StudentNumber")=Trim(Request("StudentNumber"))
				rsUser("College")=Trim(Request("College"))

				'结束两课网站代码
				rsUser.update
				call WriteSuccessMsg("成功修改用户信息！")
			else
				call WriteErrMsg()
			end if
		else

%>
<html>
<head>
<title>修改注册用户信息</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="Admin_STYLE.CSS" rel="stylesheet" type="text/css">
</head>
<body>
<br>
<FORM name="Form1" action="User_ModifyInfo.asp" method="post">
  <table width=400 border=0 align="center" cellpadding=2 cellspacing=1 class='border'>
    <TR align=center class='title'>
      <TD height=22 colSpan=2><font class=en><b>修改注册用户信息</b></font></TD>
    </TR>
    <TR class="tdbg" >
      <TD width="120" align="right"><b>用 户 名：</b></TD>
      <TD><%=Trim(Request.Cookies("asp163")("UserName"))%>
        <input name="UserName" type="hidden" value="<%=Trim(Request.Cookies("asp163")("UserName"))%>"></TD>
    </TR>
    <!--两课网站代码-->
    <TR class="tdbg" >
      <TD width="120" align="right"><b>真实姓名：</b></TD>
      <TD><INPUT name=TrueName id="TrueName" value="<%=rsUser("TrueName")%>" size=30 maxLength=20></TD>
    </TR>
    <TR class="tdbg" >
      <TD width="120" align="right"><b>学号：</b></TD>
      <TD><INPUT name=StudentNumber id="StudentNumber" value="<%=rsUser("StudentNumber")%>" size=30 maxLength=20></TD>
    </TR>
    <TR class="tdbg" >
      <TD width="120" align="right"><b>学院：</b></TD>
      <TD><select  maxLength=50 size="1" name="College">
          <option  value="" >请选择所属学院</option>
          <option  value="林学院" >林学院</option>
          <option  value="水保学院" >水保学院</option>
          <option  value="园林学院">园林学院</option>
          <option  value="生物学院">生物学院</option>
          <option  value="材料学院" >材料学院</option>
          <option  value="经管学院" >经管学院</option>
          <option  value="人文学院" >人文学院</option>
          <option  value="信息学院" >信息学院</option>
          <option  value="外语学院" >外语学院</option>
          <option  value="工学院" >工学院</option>
          <option  value="理学院" >理学院</option>
          <option  value="保护区学院" >保护区学院</option>
          <option  value="环境学院" >环境学院</option>
        </select></TD>
    </TR>
    <TR class="tdbg" >
      <TD width="120" align="right"><b>班级：</b></TD>
      <TD><!-- 专业名称-->
        <select  maxLength=50 size="1" name="StudentClassName">
          <option  value="" >请选择所属专业名称</option>
          <option  value="林学" >林学</option>
          <option  value="游憩" >游憩</option>
          <option  value="草业">草业</option>
          <option  value="地信">地信</option>
          <option  value="草坪">草坪</option>
          <option  value="水保" >水保</option>
          <option  value="环规" >环规</option>
          <option  value="土木" >土木</option>
          <option  value="园林" >园林</option>
          <option  value="游旅" >游旅</option>
          <option  value="风园" >风园</option>
          <option  value="园艺" >园艺</option>
          <option  value="城规" >城规</option>
          <option  value="生科">生科</option>
          <option  value="生物" >生物</option>
          <option  value="食品" >食品</option>
          <option  value="木工" >木工</option>
          <option  value="包装">包装</option>
          <option  value="林化">林化</option>
          <option  value="艺设" >艺设</option>
          <option  value="林经" >林经</option>
          <option  value="工商" >工商</option>
          <option  value="会计" >会计</option>
          <option  value="统计" >统计</option>
          <option  value="金融" >金融</option>
          <option  value="国贸" >国贸</option>
          <option  value="营销" >营销</option>
          <option  value="人资" >人资</option>
          <option  value="法学" >法学</option>
          <option  value="心理" >心理</option>
          <option  value="信息" >信息</option>
          <option  value="计算机" >计算机</option>
          <option  value="数媒" >数媒</option>
          <option  value="动画" >动画</option>
          <option  value="英语" >英语</option>
          <option  value="日语" >日语</option>
          <option  value="机械" >机械</option>
          <option  value="车辆" >车辆</option>
          <option  value="工设" >工设</option>
          <option  value="自动化" >自动化</option>
          <option  value="电气" >电气</option>
          <option  value="电子" >电子</option>
          <option  value="数学" >数学</option>
          <option  value="保护区" >保护区</option>
          <option  value="环境" >环境</option>
          <option  value="环工" >环工</option>
          <option  value="梁希" >梁希</option>
        </select>
        <!-- 结束专业名称下拉框-->
        <select  maxLength=50 size="1" name="StudentClassYear">
          <option  value="" >请选择所属年级</option>
          <option  value="03" >03</option>
          <option  value="04" >04</option>
          <option  value="05">05</option>
          <option  value="06">06</option>
          <option  value="07" >07</option>
          <option  value="08" >08</option>
          <option  value="09" >09</option>
          <option  value="10" >10</option>
          <option  value="11" >11</option>
          <option  value="12" >12</option>
          <option  value="13" >13</option>
          <option  value="14" >14</option>
          <option  value="15" >15</option>
        </select>
        <select  maxLength=50 size="1" name="StudentClassNumber">
          <option  value="" >请选择所属班级(本专业只有一个班的选"01")</option>
          <option  value="01" >01</option>
          <option  value="02" >02</option>
          <option  value="03" >03</option>
          <option  value="04" >04</option>
          <option  value="05" >05</option>
          <option  value="06" >06</option>
          <option  value="07" >07</option>
          <option  value="08" >08</option>
          <option  value="09" >09</option>
          <option  value="10" >10</option>
        </select>
      </TD>
    </TR>
    <!--    结束两课网站代码 -->
    <TR class="tdbg" >
      <TD width="120" align="right"><strong>性别：</strong></TD>
      <TD><INPUT type=radio value="1" name=sex <%if rsUser(db_User_Sex)=1 then response.write "CHECKED"%>>
        男 &nbsp;&nbsp;
        <INPUT type=radio value="0" name=sex <%if rsUser(db_User_Sex)=0 then response.write "CHECKED"%>>
        女</TD>
    </TR>
    <TR class="tdbg" >
      <TD width="120" align="right"><strong>Email地址：</strong></TD>
      <TD><INPUT name=Email value="<%=rsUser(db_User_Email)%>" size=30   maxLength=50>
      </TD>
    </TR>
    <!--    <TR class="tdbg" >
      <TD align="right"><strong>主页：</strong></TD>
      <TD><INPUT   maxLength=100 size=30 name=homepage value="<%=rsUser(db_User_Homepage)%>">
      </TD>
    </TR>
-->
    <TR class="tdbg" >
      <TD align="right"><strong>OICQ号码：</strong></TD>
      <TD><INPUT name=OICQ id="OICQ" value="<%=rsUser(db_User_QQ)%>" size=30 maxLength=20>
      </TD>
    </TR>
    <TR class="tdbg" >
      <TD align="right"><strong>自我简介：</strong></TD>
      <TD><textarea name="msn" cols="30" rows="5"><%=rsUser(db_User_Msn)%></textarea>
      </TD>
    </TR>
    <TR align="center" class="tdbg" >
      <TD height="40" colspan="2"><input name="Action" type="hidden" id="Action" value="Modify">
        <input name=Submit   type=submit id="Submit" value="保存修改结果">
      </TD>
    </TR>
  </TABLE>
</form>
</body>
</html>
<%
		end if
	end if
	rsUser.close
	set rsUser=nothing
end if
call CloseConn_User()
%>
