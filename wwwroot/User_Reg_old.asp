<!--#include file="Inc/syscode_article.asp"-->
<%
if EnableUserReg<>"Yes" then
	FoundErr=true
	ErrMsg=ErrMsg & "<br><li>对不起，本站暂停新用户注册服务！</li>"
	call WriteErrMsg()
	response.end
end if
const ChannelID=0
Const ShowRunTime="Yes"

dim action
action=trim(request("action"))
SkinID=0
if action="apply" then
	PageTitle="新用户注册"
else
	PageTitle="服务条款和声明"
end if
%>
<html>
<head>
<title><%=strPageTitle & " >> " & PageTitle%></title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<!--#include file="inc/Skin_CSS.asp"-->
<% call MenuJS() %>
</head>

<body leftmargin="2" topmargin="0" marginwidth="0" marginheight="0">
<!--#include file="Top.asp"-->
<table width="760" height="300" border="0" align="center" cellpadding="0" cellspacing="0" class="border2">
  <tr>
      
    <td width="180" height="300" valign="top" class="tdbg_leftall"> 
      <TABLE cellSpacing=0 cellPadding=0 width="100%" border="0" style="word-break:break-all">
        <TR class="title_left"> 
          <TD align="center"><table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr> 
                <td class="title_lefttxt"><div align="center"><strong><b>·注册<%=SiteName%></b></strong></div></td>
              </tr>
            </table></TD>
        </TR>
        <TR> 
          <TD height="80" valign="top" class="tdbg_left"> <table width="100%" height="100%" border="0" cellpadding="0" cellspacing="5">
              <tr> 
                <td valign="top"><br> <b>&nbsp;&nbsp;注册步骤</b><br> &nbsp;&nbsp;一、阅读并同意协议<font color="#FF0000">
                  <%if action<>"apply" then %>
                  →
                  <%else%>
                  √
                  <%end if%>
                  </font><br> &nbsp;&nbsp;二、填写注册资料<font color="#FF0000">
                  <%if action="apply" then %>
                  →
                  <%end if%>
                  </font><br> &nbsp;&nbsp;三、完成注册 </td>
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
      </table></td>
    <td width=5></td>
    <td width="575" align="center" valign=top> 
<%
if action<>"apply" then
%>
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr class='title_main'>
          <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr> 
                <td class="title_maintxt"><div align="center"><b><%=SiteName%>服务条款和声明</b></div></td>
              </tr>
            </table></td>
        </tr>
      </table>
      <table width=100% border=0 align=center cellpadding=0 cellspacing=0 class="border">
        <tr class='tdbg'> 
          <td height="382" align=left><br> 
            <div align=center> 
              <textarea readonly style="width:98%;height:310;font size:9pt">
 欢迎您注册成为<%=SiteName%>用户！
 请仔细阅读下面的协议，只有接受协议才能继续进行注册。 

 1．服务条款的确认和接纳

　　<%=SiteName%>用户服务的所有权和运作权归<%=SiteName%>拥有。<%=SiteName%>所提供的服务将按照有关章程、服务条款和操作规则严格执行。用户通过注册程序点击“我同意” 按钮，即表示用户与<%=SiteName%>达成协议并接受所有的服务条款。

 2． <%=SiteName%>服务简介

　　<%=SiteName%>通过国际互联网为用户提供邮件收发、BBS论坛、网上聊天和软件下载等服务。

　　用户必须： 
　　1)购置设备，包括个人电脑一台、调制解调器一个及配备上网装置。 
　　2)个人上网和支付与此服务有关的电话费用、网络费用。

　　用户同意： 
　　1)提供及时、详尽及准确的个人资料。 
　　2)不断更新注册资料，符合及时、详尽、准确的要求。所有原始键入的资料将引用为注册资料。 
　　3)用户同意遵守《中华人民共和国保守国家秘密法》、《中华人民共和国计算机信息系统安全保护条例》、《计算机软件保护条例》等有关计算机及互联网规定的法律和法规、实施办法。在任何情况下，中国站长站合理地认为用户的行为可能违反上述法律、法规，中国站长站可以在任何时候，不经事先通知终止向该用户提供服务。用户应了解国际互联网的无国界性，应特别注意遵守当地所有有关的法律和法规。
　　 
 3． 服务条款的修改

　　<%=SiteName%>会不定时地修改服务条款，服务条款一旦发生变动，将会在相关页面上提示修改内容。如果您同意改动，则再一次点击“我同意”按钮。 如果您不接受，则及时取消您的用户使用服务资格。

 4． 服务修订

　　<%=SiteName%>保留随时修改或中断服务而不需知照用户的权利。中国站长站行使修改或中断服务的权利，不需对用户或第三方负责。*邮件系统保留收费的可能。

 5． 用户隐私制度

　　尊重用户个人隐私是<%=SiteName%>的 基本政策。<%=SiteName%>不会公开、编辑或透露用户的邮件内容，除非有法律许可要求，或<%=SiteName%>在诚信的基础上认为透露这些信件在以下三种情况是必要的： 
　　1)遵守有关法律规定，遵从合法服务程序。 
　　2)保持维护<%=SiteName%>的商标所有权。 
　　3)在紧急情况下竭力维护用户个人和社会大众的隐私安全。 
　　4)符合其他相关的要求。 

 6．用户的帐号，密码和安全性

　　一旦注册成功成为<%=SiteName%>用户，您将得到一个密码和帐号。如果您不保管好自己的帐号和密码安全，将对因此产生的后果负全部责任。另外，每个用户都要对其帐户中的所有活动和事件负全责。您可随时根据指示改变您的密码，也可以结束旧的帐户重开一个新帐户。用户同意若发现任何非法使用用户帐号或安全漏洞的情况，立即通知中国站长站。

 7． 免责条款

　　用户明确同意邮件服务的使用由用户个人承担风险。 　　 
　　<%=SiteName%>不作任何类型的担保，不担保服务一定能满足用户的要求，也不担保服务不会受中断，对服务的及时性，安全性，出错发生都不作担保。用户理解并接受：任何通过中国站长站服务取得的信息资料的可靠性取决于用户自己，用户自己承担所有风险和责任。 
 
 8．有限责任

　　<%=SiteName%>对任何直接、间接、偶然、特殊及继起的损害不负责任，这些损害来自：不正当使用邮件服务，在网上购买商品或服务，在网上进行交易，非法使用邮件服务或用户传送的信息所变动。
 
 9． 不提供零售和商业性服务 
　　 
　　用户使用邮件服务的权利是个人的。用户只能是一个单独的个体而不能是一个公司或实体商业性组织。用户承诺不经<%=SiteName%>同意，不能利用邮件服务进行销售或其他商业用途。

 10．用户责任 

　　用户单独承担传输内容的责任。用户必须遵循： 
　　1)从中国境内向外传输技术性资料时必须符合中国有关法规。 
　　2)使用邮件服务不作非法用途。 
　　3)不干扰或混乱网络服务。 
　　4)不在论坛BBS或留言簿发表任何与政治相关的信息。 
　　5)遵守所有使用邮件服务的网络协议、规定、程序和惯例。
　　6)不得利用本站危害国家安全、泄露国家秘密，不得侵犯国家社会集体的和公民的合法权益。
　　7)不得利用本站制作、复制和传播下列信息： 
　　　1、煽动抗拒、破坏宪法和法律、行政法规实施的；
　　　2、煽动颠覆国家政权，推翻社会主义制度的；
　　　3、煽动分裂国家、破坏国家统一的；
　　　4、煽动民族仇恨、民族歧视，破坏民族团结的；
　　　5、捏造或者歪曲事实，散布谣言，扰乱社会秩序的；
　　　6、宣扬封建迷信、淫秽、色情、赌博、暴力、凶杀、恐怖、教唆犯罪的；
　　　7、公然侮辱他人或者捏造事实诽谤他人的，或者进行其他恶意攻击的；
　　　8、损害国家机关信誉的；
　　　9、其他违反宪法和法律行政法规的；
　　　10、进行商业广告行为的。
　　 
　　用户不能利用邮件服务作连锁邮件，垃圾邮件或分发给任何未经允许接收信件的人。用户须承诺不传输任何非法的、骚扰性的、中伤他人的、辱骂性的、恐吓性的、伤害性的、庸俗的和淫秽的信息资料。另外，用户也不能传输任何教唆他人构成犯罪行为的资料；不能传输长国内不利条件和涉及国家安全的资料；不能传输任何不符合当地法规、国家法律和国际法 律的资料。未经许可而非法进入其它电脑系统是禁止的。若用户的行为不符合以上的条款，<%=SiteName%>将取消用户服务帐号。

 11．发送通告

　　所有发给用户的通告都可通过电子邮件或常规的信件传送。<%=SiteName%>会通过邮件服务发报消息给用户，告诉他们服务条款的修改、服务变更、或其它重要事情。

 12．网站内容的所有权

　　<%=SiteName%>定义的内容包括：文字、软件、声音、相片、录象、图表；在广告中全部内容；电子邮件的全部内容；中国站长站为用户提供的商业信息。所有这些内容受版权、商标、标签和其它财产所有权法律的保护。所以，用户只能在中国站长站和广告商授权下才能使用这些内容，而不能擅自复制、篡改这些内容、或创造与内容有关的派生产品。

 13．附加信息服务

　　用户在享用<%=SiteName%>提供的免费服务的同时，同意接受<%=SiteName%>提供的各类附加信息服务。

 14．解释权

　　本注册协议的解释权归<%=SiteName%>所有。如果其中有任何条款与国家的有关法律相抵触，则以国家法律的明文规定为准。
</textarea>
            </div>
            <div align="center"> 
              <form action="User_Reg.asp" method="get">
                <input name="Action" type="hidden" id="Action" value="apply">
                <input name="Submit" type="submit" value=" 我同意 " style="cursor:hand;">
                &nbsp; 
                <input type="button" value=" 不同意 " onClick="window.location.href='index.asp'"  style="cursor:hand;">
              </form>
            </div></td>
        </tr>
      </table>
      <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
        <tr> 
          <td  height="15" align="center" valign="top"> <table width="98%" border="0" align="center" cellpadding="0" cellspacing="0">
              <tr> 
                <td height="13" Class="tdbg_left2"></td>
              </tr>
            </table></td>
        </tr>
      </table>
      <%
else
%>
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr class='title_main'> 
          <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr> 
                <td class="title_maintxt"><div align="center"><font class=en><b>新用户注册</b></font></div></td>
              </tr>
            </table></td>
        </tr>
      </table> 
      <table width=100% border=0 cellpadding=2 cellspacing=4 bordercolor="#FFFFFF" class="border" style="border-collapse: collapse">
        <FORM name='UserReg' action='User_RegPost.asp' method='post'>
          <TR class="tdbg" > 
            <TD width="43%"><b>用户名：</b><BR>
              不能超过14个字符（7个汉字）</TD>
            <TD width="57%"> <INPUT   maxLength=14 size=30 name="UserName"> <font color="#FF0000">*</font> 
              <input name="Check" type="button" id="Check" value="检查用户名" onClick="checkreg();"></TD>
          </TR>
          <TR class="tdbg" > 
            <TD width="43%"><B>密码(至少6位)：</B><BR>
              请输入密码，区分大小写。 请不要使用任何类似 '*'、' ' 或 HTML 字符 </TD>
            <TD width="57%"> <INPUT   type=password maxLength=12 size=33.5 name="Password"> 
              <font color="#FF0000">*</font> </TD>
          </TR>
          <TR class="tdbg" > 
            <TD width="43%"><strong>确认密码(至少6位)：</strong><BR>
              请再输一遍确认</TD>
            <TD width="57%"> <INPUT   type=password maxLength=12 size=33.5 name="PwdConfirm"> 
              <font color="#FF0000">*</font> </TD>
          </TR>
          <TR class="tdbg" > 
            <TD width="43%"><strong>密码问题：</strong><BR>
              忘记密码的提示问题</TD>
            <TD width="57%"> <input   type=text maxlength=50 size=30 name="Question"> 
              <font color="#FF0000">*</font> </TD>
          </TR>
          <TR class="tdbg" > 
            <TD width="43%"><strong>问题答案：</strong><BR>
              忘记密码的提示问题答案，用于取回密码</TD>
            <TD width="57%"> <INPUT   type=text maxLength=20 size=30 name="Answer"> 
              <font color="#FF0000">*</font> </TD>
          </TR>
		  
		  <!--真实姓名输入-->
		            <TR class="tdbg" > 
            <TD width="43%"><strong>姓名：</strong><BR>
              请输入您的真实姓名</TD>
            <TD width="57%"><input   type=text maxlength=50 size=30 name="TrueName"> 
              <font color="#FF0000">*</font> </TD>
          </TR>

		  
		  <!--结束真实姓名输入-->
		  
		  
		  <!--真实学号输入-->
		  <TR class="tdbg" > 
            <TD width="43%"><strong>学号：</strong><BR>
              请输入您的学号</TD>
            <TD width="57%"><input   type=text maxlength=50 size=30 name="StudentNumber"> 
              <font color="#FF0000">*</font> </TD>
          </TR>

		  
		  <!--结束真实学号输入-->

		  
          <TR class="tdbg" > 
            <TD width="43%"><strong>性别：</strong><BR>
              请选择您的性别</TD>
            <TD width="57%"> <INPUT type=radio CHECKED value="1" name=sex>
              男 &nbsp;&nbsp; <INPUT type=radio value="0" name=sex>
              女</TD>
          </TR>
          
		  <!--学院列表下拉框-->
		  <TR class="tdbg" > 
            <TD width="43%"><strong>学院：</strong></TD>
            <TD width="57%"> 
				<select  maxLength=50 size="1" name="College">
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

				</select> <font color="#FF0000">*</font></TD>
          </TR>
		  <!--结束学院列表下拉框-->
		  
		  	  <!--班级列表下拉框-->
		  <TR class="tdbg" > 
            <TD width="43%"><strong>班级：</strong></TD>
            <TD width="57%"> <!-- 专业名称-->
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
					
					
					
					

				</select><!-- 结束专业名称下拉框-->
				
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
					<option  value="05">05</option>
					<option  value="06">06</option>
					<option  value="07" >07</option>
					<option  value="08" >08</option>
					<option  value="09" >09</option>
					<option  value="10" >10</option>
					

				</select>

				
				 <font color="#FF0000">*</font>
		    </TD>
          </TR>
		  <!--结束班级列表下拉框-->
		  
		  
		  <TR class="tdbg" > 
            <TD width="43%"><strong>Email地址：</strong><BR>
              请输入有效的邮件地址，这将使您能用到网站中的所有功能</TD>
            <TD width="57%"> <INPUT   maxLength=50 size=30 name=Email> <font color="#FF0000">*</font></TD>
          </TR>
          <TR align="center" class="tdbg" > 
            <TD height="30" colspan="2"> <input   type=submit value=" 注 册 " name=Submit2> 
              &nbsp; <input name=Reset   type=reset id="Reset" value=" 清 除 "> 
            </TD>
          </TR>
        </form>
        <form name='reg' action='User_checkreg.asp' method='post' target='CheckReg'>
          <input type='hidden' name='username' value=''>
        </form>
      </TABLE>
      <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
        <tr> 
          <td  height="15" align="center" valign="top"> <table width="98%" border="0" align="center" cellpadding="0" cellspacing="0">
              <tr> 
                <td height="13" Class="tdbg_left2"></td>
              </tr>
            </table></td>
        </tr>
      </table>
      <script language=javascript>
function checkreg()
{
  if (document.UserReg.UserName.value=="")
	{
	alert("请输入用户名！");
	document.UserReg.UserName.focus();
	return false;
	}
  else
    {
	document.reg.username.value=document.UserReg.UserName.value;
	var popupWin = window.open('User_CheckReg.asp', 'CheckReg', 'scrollbars=no,width=340,height=200');
	document.reg.submit();
	}
}
</script>
<%end if%>
    </td>
    </tr>
</table>

<table width="760" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
  <tr> 
    <td  height="13" align="center" valign="top"><table width="755" border="0" align="center" cellpadding="0" cellspacing="0">
        <tr> 
          <td height="13" Class="tdbg_left2"></td>
        </tr>
      </table></td>
  </tr>
</table>
<%
call Bottom()
call CloseConn
%>
</body>
</html>
