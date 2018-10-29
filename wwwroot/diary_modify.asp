<!--#include file="Inc/syscode_diary.asp"-->
<%
if CurrentLoginUser=empty then
	response.redirect sysLoginFile
	response.end
end if

PageTitle="修改日记"
set rs=server.createobject("adodb.recordset")

dim ismyself,diaryID
diaryID=trim(request("diaryID"))
diaryID=cint(diaryID)
if diaryID=0 then
	founderr=true
	errmsg="<br><br><li>参数非法！</li>"
end if
ismyself=true		'初始化为自己的日记本

sql="SELECT top 1 * from diary where ID="&diaryID&""
rs.open sql,conn_User,1,3
if rs.eof and rs.bof then
	founderr=true
	errmsg="<br><br><li>你要查看的日记不存在！</li>"
else
	if rs("diaryBg")<>"0" then strRndBg="diary_images/back/"&rs("diaryBg")
	if rs("diaryOwner")<>CurrentLoginUser then
		founderr=true
		errmsg="<br><br><li>这则日记不是你写的，你无权修改！</li>"
	end if
end if

if request.form("modify")="modify" and trim(request.form("message"))<>"" then
	dim pic,memo,readers
	pic=request.form("xc")
	pic="face/" & pic & ".gif"
	readers=request("readers")&"|"
	readers=replace(readers,"||","|")
	if len(readers)>255 then
		dim tempA
		for i=1 to len(readers)
			tempA=mid(Readers,len(readers)-i,1)
			if tempA="|" and len(left(readers,len(readers)-i))<=255 then
				readers=left(readers,len(readers)-i)
				exit for
			end if
		next
	end if

	rs("diaryOwner")=CurrentLoginUser
	rs("diaryDate")=request.form("year")&"-"&request.form("month")&"-"&request.form("day")
	rs("weather")=request.form("tq")
	rs("mood")=pic
	rs("diarybg")=request.form("diarybg")
	memo=trim(request.form("message"))
	Memo=server.HTMLEncode(Memo)
	Memo=Replace( Memo, vbCrLf, "<BR>" )
	Memo=Replace( Memo, " ", "&nbsp;" )
	rs("diaryContent")=memo
	rs("secret")=request.form("baomi")
	if cint(request.form("baomi"))=99 then rs("readers")=readers	
	rs("addtime")=now()
	rs.update
	%>
	<script language=javascript>
		window.alert("修改日记成功！");
		window.location.href("diary_show.asp?diaryID=<%=diaryID%>");
	</script>
	<%
end if
%>

<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>修改日记</title>
<style>
BODY {scrollbar-track-color:#ffffff; SCROLLBAR-FACE-COLOR: #ffffff; FONT-SIZE: 9pt; SCROLLBAR-HIGHLIGHT-COLOR: #ffffff; SCROLLBAR-SHADOW-COLOR: #eeeeee;  SCROLLBAR-3DLIGHT-COLOR: #eeeeee; SCROLLBAR-ARROW-COLOR: #dddddd; FONT-FAMILY: "Verdana"; SCROLLBAR-DARKSHADOW-COLOR: #ffffff
}
select{font-size:8pt;font-family:verdana;background-color:#ffffff;border:1px dotted #cccccc; color:#333333;}
input{font-size:8pt;font-family:verdana;background-color:#ffffff;border-bottom:1px solid #cccccc;border-left:1px solid #cccccc; border-top:0px solid #cccccc;border-right:0px dotted #cccccc;color:#333333;}
textarea{font-size:8pt; font-family:verdana;background-color:#ffffff;border:1px dotted #cccccc;color:#333333;letter-spacing : 1pt ;line-height : 150%}
A {
	COLOR: #333333; TEXT-DECORATION: none ;border-bottom:1px dotted
}
A:hover {
	COLOR: #333333; background-color:#C0FFFF;
}
td {FONT-SIZE: 9pt;  FONT-FAMILY: "Verdana"; color:#3333333;letter-spacing : 1pt ;line-height : 150%}
.td{border:1px dotted #999999}
</style>
<!--#include file="inc/Skin_CSS.asp"-->
<%call MenuJS()%>
<script language="JavaScript">
function checkbaomi()
{
	if(myform.baomi.value=="99")
	{
		showreaders.style.display="";
		myform.selectuser.disabled=false;
	}
	else
	{
		showreaders.style.display="none";
		myform.selectuser.disabled=true;
	}
}
</script>
</head>

<body  <%=Body_Label%> onmousemove='HideMenu()'>
<!--#include file="top.asp"-->
<table width="760" border="0" align="center" cellpadding="0" cellspacing="0" class="tdbg">
  <tr>
    <td><table width="106" align="center" style="FILTER: glow(color=#333333,direction=2)">
        <tr>
          <td><font face=幼圆 color=#ffffff size=2>&nbsp;修 改 日 记&nbsp;</font></td>
        </tr>
      </table>
	  <% if founderr=true then
	  		call writeerrmsg()
		else
			call modifyform()
		end if
		rs.close
%>
</td>
  </tr>
</table>
<%

call bottom()
set rs=nothing
conn_user.close
set conn_user=nothing
%>

</body>

</html>
<% sub modifyform()%>
      <form method="post" name="myform" action="diary_modify.asp">
	  <input type="hidden" name="diaryID" value="<%=diaryID%>">
	  <input type="hidden" name="modify" value="modify">
        <table width="90%" border="0" align="center" cellpadding="0" cellspacing="10" style="border:1px dotted #999999">
          <tr>
            <td align="center"> <table border="0" cellpadding="2" cellspacing="4" width="100%">
                <tr>
                  <td>日记日期
                    <select name="Year">
                      <option value="<%=year(rs("diaryDate"))%>" selected><%=year(rs("diaryDate"))%></option>
                      <option value="<%=year(now())-1%>"><%=year(now())-1%></option>
                      <option value="<%=year(now())%>"><%=year(now())%></option>
                    </select>
                    年
                    <select name="Month">
                      <option value="<%=month(rs("diaryDate"))%>" selected><%=month(rs("diaryDate"))%></option>
                      <option value="1">1</option>
                      <option value="2">2</option>
                      <option value="3">3</option>
                      <option value="4">4</option>
                      <option value="5">5</option>
                      <option value="6">6</option>
                      <option value="7">7</option>
                      <option value="8">8</option>
                      <option value="9">9</option>
                      <option value="10">10</option>
                      <option value="11">11</option>
                      <option value="12">12</option>
                    </select>
                    月
                    <select name="Day">
                      <option value="<%=day(rs("diaryDate"))%>" selected><%=day(rs("diaryDate"))%></option>
                      <option value="1">1</option>
                      <option value="2">2</option>
                      <option value="3">3</option>
                      <option value="4">4</option>
                      <option value="5">5</option>
                      <option value="6">6</option>
                      <option value="7">7</option>
                      <option value="8">8</option>
                      <option value="9">9</option>
                      <option value="10">10</option>
                      <option value="11">11</option>
                      <option value="12">12</option>
                      <option value="13">13</option>
                      <option value="14">14</option>
                      <option value="15">15</option>
                      <option value="16">16</option>
                      <option value="17">17</option>
                      <option value="18">18</option>
                      <option value="19">19</option>
                      <option value="20">20</option>
                      <option value="21">21</option>
                      <option value="22">22</option>
                      <option value="23">23</option>
                      <option value="24">24</option>
                      <option value="25">25</option>
                      <option value="26">26</option>
                      <option value="27">27</option>
                      <option value="28">28</option>
                      <option value="29">29</option>
                      <option value="30">30</option>
                      <option value="31">31</option>
                    </select>
                    日 </td>
                </tr>
				<tr><td>
					天气:
                    <select size="1" name="tq">
                      <option value="<%=rs("weather")%>" selected><%=rs("weather")%></option>
                      <option value="晴" selected>晴</option>
                      <option value="雨">雨</option>
                      <option value="雪">雪</option>
                      <option value="多云">多云</option>
                      <option value="风">风</option>
                      <option value="冰雹">冰雹</option>
                      <option value="晴转多云">晴转多云</option>
                      <option value="多云转雪">多云转雪</option>
                      <option value="多云转雨">多云转雨</option>
                      <option value="雾">雾</option>
                    </select>
					表情:
                    <select size="1" name="xc">
                      <option value="01" selected>普通</option>
                      <option value="02">高兴</option>
                      <option value="03">扮酷</option>
                      <option value="04">困了</option>
                      <option value="05">遗憾</option>
                      <option value="06">伤心</option>
                      <option value="07">脸红</option>
                    </select>
                    背景:
                    <select name="diarybg">
                      <option value="0" selected>- 随机背景 -</option>
                      <option value="01.gif">-- 01 --</option>
                      <option value="02.gif">-- 02 --</option>
                      <option value="03.gif">-- 03 --</option>
                      <option value="04.gif">-- 04 --</option>
                      <option value="05.gif">-- 05 --</option>
                      <option value="06.gif">-- 06 --</option>
                      <option value="07.gif">-- 07 --</option>
                      <option value="08.gif">-- 08 --</option>
                      <option value="09.gif">-- 09 --</option>
                      <option value="10.gif">-- 10--</option>
                      <option value="11.gif">-- 11 --</option>
                      <option value="12.gif">-- 12 --</option>
                      <option value="13.gif">-- 13 --</option>
                      <option value="14.gif">-- 14 --</option>
                      <option value="15.gif">-- 15 --</option>
                      <option value="16.gif">-- 16 --</option>
                      <option value="17.gif">-- 17 --</option>
                      <option value="18.gif">-- 18 --</option>
                      <option value="19.gif">-- 19 --</option>
                      <option value="20.gif">-- 20--</option>
                      <option value="21.gif">-- 21 --</option>
                      <option value="22.gif">-- 22 --</option>
                      <option value="23.gif">-- 23 --</option>
                      <option value="24.gif">-- 24 --</option>
                      <option value="25.gif">-- 25 --</option>
                      <option value="26.gif">-- 26 --</option>
                      <option value="27.gif">-- 27 --</option>
                      <option value="28.gif">-- 28 --</option>
                      <option value="29.gif">-- 29 --</option>
                      <option value="30.gif">-- 30--</option>
                    </select> <a href="diary_bgpics.htm" target="_blank">{背景预览}</a>
				</td></tr>
                <tr>
                  <td>保密: 
                    <select name="baomi" onChange="checkbaomi();">
                      <option value="0" <%if rs("secret")=0 then response.write"selected"%>>=完全公开=</option>
                      <option value="9" <%if rs("secret")=9 then response.write"selected"%>>对用户公开</option>
                      <option value="99" <%if rs("secret")=99 then response.write"selected"%>>对朋友公开</option>
                      <option value="999" <%if rs("secret")=999 then response.write"selected"%>>*完全保密*</option>
                    </select>
                    <span name="showreaders" id="showreaders" style="display:none;">&nbsp;朋友: 
                    <input name="readers" type="text" id="readers" size="56" readonly=true value="<%=rs("readers")%>">
					<input type="button" name="selectuser" id="selectuser" disabled="true" value="选择朋友" onClick="window.open('diary_selectusers.asp', 'winUserList', 'width=760,height=500,resizable=0,scrollbars=yes');">
                  </span>
				  <script language="javascript">checkbaomi();</script></td>
                </tr>
                <tr>
                  <td align="center"><img border="0" src="diary_images/dia-b-icon.gif" width="21" height="21" align="absmiddle">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<img border="0" src="diary_images/dia-b-icon.gif" width="21" height="21" align="absmiddle">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<img border="0" src="diary_images/dia-b-icon.gif" width="21" height="21" align="absmiddle"><br>
                    <textarea name="Message" cols="100" rows="14"><%=replace(rs("diarycontent"),"<BR>", vbCrLf)%></textarea>
                  </td>
                </tr>
              </table></td>
          </tr>
        </table>
        <p align=center>
          <input type="submit" name="btnAdd" value="确定修改">
          &nbsp;&nbsp;
          <input type="Reset" name="btnReset" value="我要重来">
      		&nbsp; &nbsp;&nbsp;<a href="diary_index.asp?diaryOwner=<%=CurrentLoginUser%>">不改了，我要返回</a>
      </form>
<% end sub %>