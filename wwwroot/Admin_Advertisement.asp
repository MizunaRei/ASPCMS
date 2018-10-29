<%@language=vbscript codepage=936 %>
<%
option explicit
response.buffer=true	
Const PurviewLevel=2
Const CheckChannelID=0
Const PurviewLevel_Others="AD"
%>
<!--#include file="inc/conn.asp"-->
<!--#include file="inc/config.asp"-->
<!--#include file="Admin_ChkPurview.asp"-->
<!--#include file="inc/ubbcode.asp"-->
<!--#include file="inc/function.asp"-->
<%
dim sql,rs,strFileName
dim Action,Channel,FoundErr,ErrMsg
Action=Trim(Request("Action"))
Channel=Trim(Request("Channel"))
if Channel="" then
	Channel=0
else
	Channel=CLng(Channel)
end if
strFileName="Admin_Advertisement.asp?Channel="&Channel
%>
<html>
<head>
<title>广告管理</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" type="text/css" href="Admin_Style.css">
<SCRIPT language=javascript>
function unselectall(thisform)
{
    if(thisform.chkAll.checked)
	{
		thisform.chkAll.checked = thisform.chkAll.checked&0;
    } 	
}

function CheckAll(thisform)
{
	for (var i=0;i<thisform.elements.length;i++)
    {
	var e = thisform.elements[i];
	if (e.Name != "chkAll"&&e.disabled!=true)
		e.checked = thisform.chkAll.checked;
    }
}
function ConfirmDel(thisform)
{
	if(thisform.Action.value=="Del")
	{
		if(confirm("确定要删除选中的广告吗？"))
		    return true;
		else
			return false;
	}
}
function showsetting(thisform)
{
	for (var j=0;j<6;j++)
	{
		var tab = eval("document.all.settable"+j);
		if(thisform.ADType.selectedIndex==j)
			tab.style.display = "";
		else
			tab.style.display = "none";
	}
	if(thisform.ADType.selectedIndex==6)
	{
		nocodead.style.display = "none";
		iscodead.style.display = "";
	}
	else
	{
		nocodead.style.display = "";
		iscodead.style.display = "none";
	}
}
function check(thisform)
{
	if(thisform.ADType.selectedIndex==6)
	{
		if(thisform.ADCode.value=="")
		{
			alert("广告代码不能为空！")
			thisform.ADCode.focus()
			return false;
		}
		if(thisform.ADCode.value.length>250)
		{
			alert("广告代码长度不能超过250字符！")
			thisform.ADCode.focus()
			return false;
		}
	}
	else
	{
		if(thisform.SiteName.value=="")
		{
		  alert("网站名称不能为空！");
		  thisform.SiteName.focus();
		  return false;
		}
		if(thisform.ImgUrl.value=="")
		{
		  alert("图片地址不能为空！");
		  thisform.ImgUrl.focus();
		  return false;
		}
		if(thisform.ADType.selectedIndex==0)
		{
			if(thisform.ImgWidth.value=="")
			{
			  alert("弹出广告的图片宽度不能留空！");
			  thisform.ImgWidth.focus();
			  return false;
			}
			if(thisform.ImgHeight.value=="")
			{
			  alert("弹出广告的图片高度不能留空！");
			  thisform.ImgHeight.focus();
			  return false;
			}
		}
	}
}
</script>
</head>
<body leftmargin="2" topmargin="0" marginwidth="0" marginheight="0">
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" Class="border">
  <tr class="topbg"> 
    <td height="22" colspan=2 align=center><strong>广 告 管 理</strong></td>
  </tr>
  <tr class="tdbg"> 
    <td width="70" height="30"><strong>管理导航：</strong></td>
    <td height="30"><a href="Admin_Advertisement.asp?Action=Add">添加新广告</a> | <a href="Admin_Advertisement.asp">所有频道广告</a> 
      | <a href="Admin_Advertisement.asp?Channel=1">网站首页广告</a> | <a href="Admin_Advertisement.asp?Channel=2">文章频道广告</a> 
      \<a href="Admin_Advertisement.asp?Channel=4"></a> | <a href="Admin_Advertisement.asp?Channel=5">留言频道广告</a> 
      | </td>
  </tr>
</table>
<%
if Action="Add" then
	call Add()
elseif Action="SaveAdd" then
	call SaveAdd()
elseif Action="Modify" then
	call Modify()
elseif Action="SaveModify" then
	call SaveModify()
elseif Action="SetNew" then
	call SetNew()
elseif Action="CancelNew" then
	call CancelNew()
elseif Action="Move" then
	call MoveAdvertisement()
elseif Action="Del" then
	call DelAD()
else
	call main()
end if
if FoundErr=True then
	call WriteErrMsg()
end if
call CloseConn()


sub main()
	sql="select * from Advertisement"
	sql=sql & " where ChannelID=" & Channel
	sql=sql & " order by IsSelected,id desc"
	set rs=server.createobject("adodb.recordset")
	rs.open sql,conn,1,1
%>
<form name="form1" method="POST" action=<%=strFileName%> >
<%
response.write "您现在的位置：网站广告管理&nbsp;&gt;&gt;&nbsp;<font color=red>"
select case Channel
	case 0
		response.write "所有频道广告"
	case 1
		response.write "网站首页广告"
	case 2
		response.write "文章频道广告"
	case 3
		response.write "软件频道广告"
	case 4
		response.write "图片频道广告"
	case 5
		response.write "留言频道广告"
	case else
		response.write "错误的参数"
end select
response.write "</font><br>"
%>
  <table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" Class="border">
    <tr class="title"> 
      <td width="30" height="22" align="center"><strong>选择</strong></td>
      <td width="30" height="22" align="center"><strong>ID</strong></td>
      <td width="20" align="center"><strong>新</strong></td>
      <td width="80" align="center"><strong>广告类型</strong></td>
      <td width="100" height="22" align="center"><strong>网站名称</strong></td>
      <td height="22" align="center"><strong>广告图片</strong></td>
      <td width="80" height="22" align="center"><strong>操作</strong></td>
    </tr>
    <%
if not(rs.bof and rs.eof) then
	do while not rs.eof
%>
    <tr class="tdbg" onmouseout="this.style.backgroundColor=''" onmouseover="this.style.backgroundColor='#BFDFFF'"> 
      <td width="30" align="center"> 
        <input type="checkbox" value=<%=rs("ID")%> name="ID" onclick="unselectall(this.form)">
      </td>
      <td width="30" align="center"><%=rs("ID")%></td>
      <td width="20" align="center">
		  <%if rs("IsSelected")=true then response.write "<font color=#009900>新</font>" end if%>
	  </td>
      <td width="80" align="center"> 
        <%
		if rs("ADType")=0 then
			response.write "弹出广告" 
		elseif rs("ADType")=1 then
			response.write "Banner广告" 
		elseif rs("ADType")=2 then
			response.write "栏目广告" 
		elseif rs("ADType")=3 then
			response.write "文章内容页广告" 
		elseif rs("ADType")=4 then
			response.write "浮动广告" 
		elseif rs("ADType")=5 then
			response.write "页面固定广告" 
		elseif rs("ADType")=6 then
			response.write "代码广告" 
		else
			response.write "其它广告" 
		end if
		%>
      </td>
      <td width="100"><a href="<%=rs("SiteUrl")%>" target='blank' title="网站地址：<%=rs("SiteUrl") & vbcrlf %>网站简介：<%=vbcrlf & rs("SiteIntro")%>"><%=rs("SiteName")%></a></td>
      <td align="center"> 
        <%
		if rs("ADType")=6 then
			Response.Write rs("ImgUrl")
		else
			if lcase(right(rs("ImgUrl"),4))=".swf" then
				Response.Write "<object classid='clsid:D27CDB6E-AE6D-11cf-96B8-444553540000' codebase='http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=5,0,0,0'"
				if rs("ImgWidth")>0 then 
					response.write " width='" & rs("ImgWidth") & "'"
					if rs("ImgHeight")>0 then response.write " height='" & rs("ImgHeight") & "'"
				end if
				response.write "><param name='movie' value='" & rs("ImgUrl") & "'><param name='quality' value='high'><embed src='" & rs("ImgUrl") & "' pluginspage='http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash' type='application/x-shockwave-flash'"
				if rs("ImgWidth")>0 then 
					response.write " width='" & rs("ImgWidth") & "'"
					if rs("ImgHeight")>0 then response.write " height='" & rs("ImgHeight") & "'"
				end if
				response.write "></embed></object>"
			else
				response.write "<a href='" & rs("SiteUrl") & "' target='_blank' title='图片地址：" & rs("ImgUrl") & vbcrlf & "图片宽度：" & rs("ImgWidth") & "像素" & vbcrlf & "图片高度：" & rs("ImgHeight") & "像素'><img src='" & rs("ImgUrl") & "'"
				if rs("ImgWidth")>0 then 
					response.write " width='" & rs("ImgWidth") & "'"
					if rs("ImgHeight")>0 then response.write " height='" & rs("ImgHeight") & "'"
				end if
				response.write " border='0'></a>"
			end if
		end if
		%>
      </td>
      <td width="80" align="center"> 
        <%
	  response.write "<a href='" & strFileName & "&Action=Modify&ID=" & rs("ID") & "'>修改</a>&nbsp;&nbsp;"
		response.write "<a href='" & strFileName & "&Action=Del&ID=" & rs("ID") & "' onclick=""return confirm('确定要删除此广告吗？');"">删除</a>"
	  %>
      </td>
    </tr>
    <%
rs.movenext
loop
%>
    <tr class="tdbg"> 
      <td colspan=7 height="30"> 
        <input name="chkAll" type="checkbox" id="chkAll" onclick=CheckAll(this.form) value="checkbox">
        选中所有广告 
        <input name="Action" type="hidden" id="Action" value="Del">
		&nbsp;&nbsp;&nbsp;&nbsp;将选定的广告： 
        <input type="submit" value=" 删除 " name="submit" onClick="this.form.Action.value='Del';return ConfirmDel(this.form);"">
        &nbsp;&nbsp;
        <input type="submit" value="设为最新广告" name="submit" onClick="this.form.Action.value='SetNew'">
        &nbsp;&nbsp;
        <input type="submit" value="取消最新广告" name="submit" onClick="this.form.Action.value='CancelNew'">
        &nbsp;&nbsp;
        <input type="submit" value="移动至" name="submit" onClick="this.form.Action.value='Move'">
		<select name='ChannelID'>
			<option value='0'>全部</option>
			<option value='1'>首页</option>
			<option value='2'>文章</option>
			<option value='3'>软件</option>
			<option value='4'>图片</option>
			<option value='5'>留言</option>
		</select>
      </td>
    </tr>
    <% end if%>
  </table>
</form>
<%
	rs.close
	set rs=nothing
end sub

sub Add()
%>
<form name="myform" method="post" action="<%=strFileName%>" onSubmit="return check(myform)">
  <table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="border">
    <tr class="title"> 
      <td height="22" colspan="2" align="center"><strong>添 加 广 告</strong></td>
    </tr>
    <tr class="tdbg"> 
      <td width="250"><strong>所属频道：</strong></td>
      <td width="550"> 
        <input type='radio' name='ChannelID' value='0' checked>
        全部&nbsp; 
        <input type='radio' name='ChannelID' value='1'>
        首页&nbsp; 
        <input type='radio' name='ChannelID' value='2'>
        文章&nbsp; 
        <input type='radio' name='ChannelID' value='3'>
        软件&nbsp; 
        <input type='radio' name='ChannelID' value='4'>
        图片&nbsp; 
        <input type='radio' name='ChannelID' value='5'>
        留言&nbsp; </td>
    </tr>
    <tr class="tdbg"> 
      <td width="250"><strong>广告类型：</strong></td>
      <td width="550"> 
        <select name="ADType" id="ADType" onchange=showsetting(this.form)>
          <option value="0" selected >弹出广告</option>
          <option value="1" >Banner广告</option>
          <option value="2" >栏目广告</option>
          <option value="3" >文章内容页广告</option>
          <option value="4" >浮动广告</option>
          <option value="5" >页面固定广告</option>
          <option value="6" >代码广告</option>
        </select>
		&nbsp;&nbsp;&nbsp;&nbsp;<input name="IsSelected" type="checkbox" id="IsSelected" value="True" checked>
        设为最新广告 
      </td>
    </tr>
  </table>
  <table id="nocodead" width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="border">
    <tr class="tdbg"> 
      <td width="250"><strong>广告设置：</strong></td>
      <td height="26" width="550"> 
        <table width="100%" border="0" cellspacing="0" cellpadding="0" id="settable0">
          <tr> 
            <td>左： 
              <input name="popleft" type="text" id="popleft" value="100" size="6" maxlength="5">
              上： 
              <input name="poptop" type="text" id="poptop" value="100" size="6" maxlength="5"> 
            </td>
          </tr>
        </table>
        <table width="100%" border="0" cellspacing="0" cellpadding="0" id="settable1" style="DISPLAY: none">
          <tr> 
            <td> </td>
          </tr>
        </table>
        <table width="100%" border="0" cellspacing="0" cellpadding="0" id="settable2" style="DISPLAY: none">
          <tr> 
            <td> </td>
          </tr>
        </table>
        <table width="100%" border="0" cellspacing="0" cellpadding="0" id="settable3" style="DISPLAY: none">
          <tr> 
            <td> </td>
          </tr>
        </table>
        <table width="100%" border="0" cellspacing="0" cellpadding="0" id="settable4" style="DISPLAY: none">
          <tr> 
            <td>左： 
              <input name="floatleft" type="text" id="floatleft" value="100" size="6" maxlength="5">
              上： 
              <input name="floattop" type="text" id="floattop" value="100" size="6" maxlength="5"> 
            </td>
          </tr>
        </table>
        <table width="100%" border="0" cellspacing="0" cellpadding="0" id="settable5" style="DISPLAY: none">
          <tr> 
            <td>左： 
              <input name="fixedleft" type="text" id="fixedleft" value="100" size="6" maxlength="5">
              上： 
              <input name="fixedtop" type="text" id="fixedtop" value="100" size="6" maxlength="5"> 
            </td>
          </tr>
        </table></td>
    </tr>
    <tr class="tdbg"> 
      <td width="250"><strong>网站名称：</strong></td>
      <td width="550"> 
        <input name="SiteName" type="text" id="SiteName" value="" size="58" maxlength="255"> 
      </td>
    </tr>
    <tr class="tdbg"> 
      <td width="250"><strong>网站地址：</strong></td>
      <td width="550"> 
        <input name="SiteUrl" type="text" id="SiteUrl" value="http://" size="58" maxlength="255"> 
      </td>
    </tr>
    <tr class="tdbg"> 
      <td width="250"><strong>网站简介：</strong></td>
      <td width="550"> 
        <input name="SiteIntro" type="text" id="SiteIntro" size="58" maxlength="255"> 
      </td>
    </tr>
    <tr class="tdbg"> 
      <td width="250"><strong>图片地址：</strong><br>
        图片格式为：jpg,gif,bmp,png,swf</td>
      <td width="550"> 
        <input name="ImgUrl" type="text" id="ImgUrl" size="58" maxlength="255"> 
      </td>
    </tr>
    <tr class="tdbg"> 
      <td width="250"><strong>图片上传：</strong></td>
      <td width="550"> <iframe style="top:2px" ID="UploadFiles" src="Upload_AdPic.asp" frameborder=0 scrolling=no width="450" height="25"></iframe> 
      </td>
    </tr>
    <tr class="tdbg"> 
      <td width="250"><strong>图片大小：</strong>（留空则取原始大小）</td>
      <td width="550">宽： 
        <input name="ImgWidth" type="text" id="ImgWidth"  size="6" maxlength="5">
        像素&nbsp;&nbsp;&nbsp;&nbsp;高： 
        <input name="ImgHeight" type="text" id="ImgHeight"  size="6" maxlength="5">
        像素&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<font color="#FF0000">＊&nbsp;</font><font color="#0000FF">弹出广告图片大小不能留空</font></td>
    </tr>
    <tr class="tdbg"> 
      <td width="250"><strong>是否FLASH：</strong></td>
      <td width="550"> 
        <input type="radio" name="IsFlash" value="True">
        是&nbsp;&nbsp;&nbsp;&nbsp; <input name="IsFlash" type="radio" value="False" checked>
        否</td>
    </tr>
  </table>
  <table id="iscodead" width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="border"  style="DISPLAY: none">
    <tr class="tdbg"> 
      <td width="250"><strong>广告代码：</strong></td>
      <td width="550"> <br>
        <textarea name="ADCode" id="ADCode" cols="50" rows="10"></textarea>
        <br>
        <br>
      </td>
    </tr>
  </table>
  <table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" >
    <tr> 
      <td height="40" colspan="2" align="center"> 
        <input name="Action" type="hidden" id="Action" value="SaveAdd">
        <input type="submit" name="Submit" value=" 添 加 ">
      </td>
    </tr>
  </table>
</form>
<%
end sub

sub Modify()
	dim ID,arrSetting,popleft,poptop,floatleft,floattop,fixedleft,fixedtop
	ID=trim(request("ID"))
	if ID="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>请指定广告ID</li>"
		exit sub
	else
		ID=Clng(ID)
	end if
	sql="select * from Advertisement where ID=" & ID
	set rs=server.createobject("adodb.recordset")
	rs.open sql,conn,1,1
	if rs.bof and rs.eof then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>找不到指定的广告！</li>"
		rs.close
		set rs=nothing
		exit sub
	end if

	popleft="100"
	poptop="100"
	floatleft="100"
	floattop="100"
	fixedleft="100"
	fixedtop="100"
	if rs("ADType")=0 then
		if instr(rs("ADSetting"),"|")>0 then
			arrSetting=split(rs("ADSetting"),"|")
			popleft=arrsetting(0)
			poptop=arrsetting(1)
		end if
	elseif rs("ADType")=4 then
		if instr(rs("ADSetting"),"|")>0 then
			arrSetting=split(rs("ADSetting"),"|")
			floatleft=arrsetting(0)
			floattop=arrsetting(1)
		end if
	elseif rs("ADType")=5 then
		if instr(rs("ADSetting"),"|")>0 then
			arrSetting=split(rs("ADSetting"),"|")
			fixedleft=arrsetting(0)
			fixedtop=arrsetting(1)
		end if
	end if


%>
<form name="myform" method="post" action="<%=strFileName%>" onSubmit="return check(myform)">
  <table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="border">
    <tr class="title"> 
      <td height="22" colspan="2" align="center"><strong>修 改 广 告</strong></td>
    </tr>
    <tr class="tdbg"> 
      <td width="250"><strong>所属频道：</strong></td>
      <td width="550"> 
        <input type='radio' name='ChannelID' value='0' <%if rs("ChannelID")=0 then response.write "checked"%>>
        全部&nbsp; 
        <input type='radio' name='ChannelID' value='1' <%if rs("ChannelID")=1 then response.write "checked"%>>
        首页&nbsp; 
        <input type='radio' name='ChannelID' value='2' <%if rs("ChannelID")=2 then response.write "checked"%>>
        文章&nbsp; 
        <input type='radio' name='ChannelID' value='3' <%if rs("ChannelID")=3 then response.write "checked"%>>
        软件&nbsp; 
        <input type='radio' name='ChannelID' value='4' <%if rs("ChannelID")=4 then response.write "checked"%>>
        图片&nbsp; 
        <input type='radio' name='ChannelID' value='5' <%if rs("ChannelID")=5 then response.write "checked"%>>
        留言&nbsp; </td>
    </tr>
    <tr class="tdbg"> 
      <td width="250"><strong>广告类型：</strong></td>
      <td width="550"> 
        <select name="ADType" id="ADType" onchange=showsetting(this.form)>
          <option value="0" <%if rs("ADType")=0 then response.write "selected"%>>弹出广告</option>
          <option value="1" <%if rs("ADType")=1 then response.write "selected"%>>Banner广告</option>
          <option value="2" <%if rs("ADType")=2 then response.write "selected"%>>栏目广告</option>
          <option value="3" <%if rs("ADType")=3 then response.write "selected"%>>文章内容页广告</option>
          <option value="4" <%if rs("ADType")=4 then response.write "selected"%>>浮动广告</option>
          <option value="5" <%if rs("ADType")=5 then response.write "selected"%>>页面固定广告</option>
          <option value="6" <%if rs("ADType")=6 then response.write "selected"%>>代码广告</option>
        </select>
		&nbsp;&nbsp;&nbsp;&nbsp;<input name="IsSelected" type="checkbox" id="IsSelected" value="True" <% if rs("IsSelected")=true then response.write "checked"%>>
        设为最新广告
      </td>
    </tr>
  </table>
  <table id="nocodead" width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="border">
    <tr class="tdbg"> 
      <td width="250"><strong>广告设置：</strong></td>
      <td width="550" height="26"> 
        <table width="100%" border="0" cellspacing="0" cellpadding="0" id="settable0">
          <tr> 
            <td>左： 
              <input name="popleft" type="text" id="popleft" value='<%=popleft%>' size="6" maxlength="5">
              上： 
              <input name="poptop" type="text" id="poptop" value='<%=poptop%>' size="6" maxlength="5">
            </td>
          </tr>
        </table>
        <table width="100%" border="0" cellspacing="0" cellpadding="0" id="settable1" style="DISPLAY: none">
          <tr> 
            <td> </td>
          </tr>
        </table>
        <table width="100%" border="0" cellspacing="0" cellpadding="0" id="settable2" style="DISPLAY: none">
          <tr> 
            <td> </td>
          </tr>
        </table>
        <table width="100%" border="0" cellspacing="0" cellpadding="0" id="settable3" style="DISPLAY: none">
          <tr> 
            <td> </td>
          </tr>
        </table>
        <table width="100%" border="0" cellspacing="0" cellpadding="0" id="settable4" style="DISPLAY: none">
          <tr> 
            <td>左： 
              <input name="floatleft" type="text" id="floatleft" value='<%=floatleft%>' size="6" maxlength="5">
              上： 
              <input name="floattop" type="text" id="floattop" value='<%=floattop%>' size="6" maxlength="5">
            </td>
          </tr>
        </table>
        <table width="100%" border="0" cellspacing="0" cellpadding="0" id="settable5" style="DISPLAY: none">
          <tr> 
            <td>左： 
              <input name="fixedleft" type="text" id="fixedleft" value='<%=fixedleft%>' size="6" maxlength="5">
              上： 
              <input name="fixedtop" type="text" id="fixedtop" value='<%=fixedtop%>' size="6" maxlength="5">
            </td>
          </tr>
        </table>
      </td>
    </tr>
    <tr class="tdbg"> 
      <td width="250"><strong>网站名称：</strong></td>
      <td width="550"> 
        <input name="SiteName" type="text" id="SiteName" value="<%=rs("SiteName")%>" size="50" maxlength="255">
      </td>
    </tr>
    <tr class="tdbg"> 
      <td width="250"><strong>网站地址：</strong></td>
      <td width="550"> 
        <input name="SiteUrl" type="text" id="SiteUrl" value="<%=rs("SiteUrl")%>" size="50" maxlength="255">
      </td>
    </tr>
    <tr class="tdbg"> 
      <td width="250"><strong>网站简介：</strong></td>
      <td width="550"> 
        <input name="SiteIntro" type="text" id="SiteIntro" value="<%=rs("SiteIntro")%>" size="50" maxlength="255">
      </td>
    </tr>
    <tr class="tdbg"> 
      <td width="250"><strong>图片地址：</strong></td>
      <td width="550"> 
        <input name="ImgUrl" type="text" id="ImgUrl" value="<%if rs("ADType")<>6 then response.write rs("ImgUrl")%>" size="50" maxlength="255">
      </td>
    </tr>
    <tr class="tdbg">
      <td width="250"><strong>图片上传：</strong></td>
      <td width="550">
	  <iframe style="top:2px" ID="UploadFiles" src="Upload_AdPic.asp" frameborder=0 scrolling=no width="450" height="25"></iframe>
	  </td>
    </tr>
    <tr class="tdbg"> 
      <td width="250"><strong>图片大小：</strong></td>
      <td width="550">宽： 
        <input name="ImgWidth" type="text" id="ImgWidth" value="<%=rs("ImgWidth")%>" size="6" maxlength="5">
        像素&nbsp;&nbsp;&nbsp;&nbsp;高： 
        <input name="ImgHeight" type="text" id="ImgHeight" value="<%=rs("ImgHeight")%>" size="6" maxlength="5">
        像素&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<font color="#FF0000">＊&nbsp;</font><font color="#0000FF">弹出广告图片大小不能留空</font></td>
    </tr>
    <tr class="tdbg"> 
      <td width="250"><strong>是否FLASH：</strong></td>
      <td width="550"> 
        <input type="radio" name="IsFlash" value="True" <% if rs("IsFlash")=true then response.write "checked"%>>
        是&nbsp;&nbsp;&nbsp;&nbsp; 
        <input name="IsFlash" type="radio" value="False" <% if rs("IsFlash")=false then response.write "checked"%>>
        否</td>
    </tr>
  </table>
  <table id="iscodead" width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="border" style="DISPLAY: none">
    <tr class="tdbg"> 
      <td width="250"><strong>广告代码：</strong></td>
      <td width="550"> <br>
        <textarea name="ADCode" id="ADCode" cols="50" rows="10"><%if rs("ADType")=6 then response.write  rs("ImgUrl")%></textarea>
        <br>
        <br>
      </td>
    </tr>
  </table>
  <table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" >
    <tr> 
      <td height="40" colspan="2" align="center"> 
        <input name="Action" type="hidden" id="Action" value="SaveModify">
        <input name="ID" type="hidden" id="ID" value="<%=rs("ID")%>">
        <input type="submit" name="Submit" value=" 保 存 ">
      </td>
    </tr>
  </table>
<script language=JavaScript>
showsetting(myform)
</script>
</form>
<%
	rs.close
	set rs=nothing
end sub
%>
<table width="100%" border="0" cellspacing="1" cellpadding="2" class="border">
  <tr class="title"> 
    <td height="22" colspan="2"><strong>广告类型说明</strong></td>
  </tr>
  <tr class="tdbg"> 
    <td>弹出广告：</td>
    <td>指采用弹出窗口形式的广告</td>
  </tr>
  <tr class="tdbg"> 
    <td>Banner广告：</td>
    <td>指页面顶部中间Banner处的广告，其大小一般为480*60</td>
  </tr>
  <tr class="tdbg"> 
    <td>栏目广告：</td>
    <td>指穿插在各栏目间的广告，其大小一般为480*60</td>
  </tr>
  <tr class="tdbg"> 
    <td>文章内容页广告：</td>
    <td>指显示在文章内容中间的广告，其大小一般为300*300</td>
  </tr>
  <tr class="tdbg"> 
    <td>浮动广告：</td>
    <td>指漂浮在页面上不断移动的广告，其大小一般为80*80 </td>
  </tr>
  <tr class="tdbg"> 
    <td>页面固定广告：</td>
    <td>指固定显示在页面某一位置的广告</td>
  </tr>
  <tr class="tdbg"> 
    <td>代码广告：</td>
    <td>指包含html内容的网站推广代码</td>
  </tr>
</table>
</body>
</html>
<%
sub SaveAdd()
	dim ChannelID,ADType,SiteName,SiteUrl,SiteIntro,ImgUrl,ImgWidth,ImgHeight,IsFlash,IsSelected,ADSetting
	dim popleft,poptop,floatleft,floattop,fixedleft,fixedtop
	ChannelID=Clng(request("ChannelID"))
	ADType=Clng(request("ADType"))
	SiteUrl=trim(request("SiteUrl"))
	SiteName=trim(request("SiteName"))
	SiteIntro=trim(request("SiteIntro"))
	ImgWidth=trim(request("ImgWidth"))
	ImgHeight=Trim(request("ImgHeight"))
	IsFlash=trim(request("IsFlash"))
	IsSelected=trim(request("IsSelected"))
	if ADType<>6 then
		ImgUrl=trim(request("ImgUrl"))
		if SiteName="" then
			FoundErr=True
			ErrMsg=ErrMsg & "<br><li>网站名称不能为空！</li>"
		end if
		if ImgUrl="" then
			FoundErr=True
			ErrMsg=ErrMsg & "<br><li>广告图片不能为空！</li>"
		end if
	else
		ImgUrl=trim(request("ADCode"))
		if ImgUrl="" then
			FoundErr=True
			ErrMsg=ErrMsg & "<br><li>广告代码不能为空！</li>"
		end if
	end if

	if FoundErr=True then
		exit sub
	end if
	
	if SiteUrl="http://" then SiteUrl="http://www.asp163.net"
	if ImgWidth="" then 
		ImgWidth=0
	else
		ImgWidth=Cint(ImgWidth)
	end if
	if ImgHeight="" then
		ImgHeight=0
	else
		ImgHeight=Cint(ImgHeight)
	end if
	if IsFlash="" then IsFlash=false
	if IsSelected="" then IsSelected=false

	
	if ADType=0 then
		if trim(request("popleft"))="" then popleft=0 else popleft=trim(request("popleft"))
		if trim(request("poptop"))="" then poptop=0 else poptop=trim(request("poptop"))
		ADSetting=popleft & "|" & poptop
	elseif ADType=4 then
		if trim(request("floatleft"))="" then floatleft=0 else floatleft=trim(request("floatleft"))
		if trim(request("floattop"))="" then floattop=0 else floattop=trim(request("floattop"))
		ADSetting=floatleft & "|" & floattop
	elseif ADType=5 then
		if trim(request("fixedleft"))="" then fixedleft=0 else fixedleft=trim(request("fixedleft"))
		if trim(request("fixedtop"))="" then fixedtop=0 else fixedtop=trim(request("fixedtop"))
		ADSetting=fixedleft & "|" & fixedtop
	end if

	sql="select * from Advertisement"
	set rs=server.createobject("adodb.recordset")
	rs.open sql,conn,1,3
	rs.addnew
	rs("ChannelID")=ChannelID
	rs("ADType")=ADType
	rs("ADSetting")=ADSetting
	rs("SiteName")=SiteName
	rs("SiteUrl")=SiteUrl
	rs("SiteIntro")=SiteIntro
	rs("ImgUrl")=ImgUrl
	rs("ImgWidth")=ImgWidth
	rs("ImgHeight")=ImgHeight
	rs("IsFlash")=IsFlash
	rs("IsSelected")=IsSelected
	rs.update
	rs.close
	set rs=nothing
	call CloseConn()
	response.redirect "Admin_Advertisement.asp?Channel="&ChannelID
end sub

sub SaveModify()
	dim sql,rs
	dim ID,ChannelID,ADType,SiteName,SiteUrl,SiteIntro,ImgUrl,ImgWidth,ImgHeight,IsFlash,IsSelected,ADSetting
	dim popleft,poptop,floatleft,floattop,fixedleft,fixedtop
	ID=trim(request("ID"))
	ChannelID=Clng(request("ChannelID"))
	ADType=Clng(request("ADType"))
	SiteName=trim(request("SiteName"))
	SiteUrl=trim(request("SiteUrl"))
	SiteIntro=trim(request("SiteIntro"))
	ImgWidth=trim(request("ImgWidth"))
	ImgHeight=Trim(request("ImgHeight"))
	IsFlash=trim(request("IsFlash"))
	IsSelected=trim(request("IsSelected"))
	if ID="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>请指定广告ID</li>"
	else
		ID=Clng(ID)
	end if
	if ADType<>6 then
		ImgUrl=trim(request("ImgUrl"))
		if SiteName="" then
			FoundErr=True
			ErrMsg=ErrMsg & "<br><li>网站名称不能为空！</li>"
		end if
		if ImgUrl="" then
			FoundErr=True
			ErrMsg=ErrMsg & "<br><li>广告图片不能为空！</li>"
		end if
	else
		ImgUrl=trim(request("ADCode"))
		if ImgUrl="" then
			FoundErr=True
			ErrMsg=ErrMsg & "<br><li>广告代码不能为空！</li>"
		end if
	end if
	if FoundErr=True then
		exit sub
	end if
	if SiteUrl="http://" then SiteUrl="http://www.asp163.net"
	if ImgWidth="" then 
		ImgWidth=0
	else
		ImgWidth=Cint(ImgWidth)
	end if
	if ImgHeight="" then
		ImgHeight=0
	else
		ImgHeight=Cint(ImgHeight)
	end if
	if IsFlash="" then IsFlash=false
	if IsSelected="" then IsSelected=false

	if ADType=0 then
		if trim(request("popleft"))="" then popleft=0 else popleft=trim(request("popleft"))
		if trim(request("poptop"))="" then poptop=0 else poptop=trim(request("poptop"))
		ADSetting=popleft & "|" & poptop
	elseif ADType=4 then
		if trim(request("floatleft"))="" then floatleft=0 else floatleft=trim(request("floatleft"))
		if trim(request("floattop"))="" then floattop=0 else floattop=trim(request("floattop"))
		ADSetting=floatleft & "|" & floattop
	elseif ADType=5 then
		if trim(request("fixedleft"))="" then fixedleft=0 else fixedleft=trim(request("fixedleft"))
		if trim(request("fixedtop"))="" then fixedtop=0 else fixedtop=trim(request("fixedtop"))
		ADSetting=fixedleft & "|" & fixedtop
	end if

	sql="select * from Advertisement where ID=" & ID
	set rs=server.createobject("adodb.recordset")
	rs.open sql,conn,1,3
	if rs.bof and rs.eof then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>找不到指定的广告！</li>"
		rs.close
		set rs=nothing
		exit sub
	end if
	rs("ChannelID")=ChannelID
	rs("ADType")=ADType
	rs("ADSetting")=ADSetting
	rs("SiteName")=SiteName
	rs("SiteUrl")=SiteUrl
	rs("SiteIntro")=SiteIntro
	rs("ImgUrl")=ImgUrl
	rs("ImgWidth")=ImgWidth
	rs("ImgHeight")=ImgHeight
	rs("IsFlash")=IsFlash
	rs("IsSelected")=IsSelected
	rs.update
	rs.close
	set rs=nothing
	call CloseConn()
	response.redirect strFileName
end sub

sub SetNew()
	dim ID
	ID=Trim(Request("ID"))
	if ID="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>请指定广告ID</li>"
		exit sub
	end if
	if Instr(ID,",")>0 then
		dim arrID,i
		arrID=split(ID,",")
		for i=0 to Ubound(arrID)
			conn.execute "Update Advertisement set IsSelected=True Where ID=" & CLng(arrID(i))
		next
	else
		conn.execute "Update Advertisement set IsSelected=True Where ID=" & CLng(ID)
	end if
	response.redirect strFileName
end sub

sub CancelNew()
	dim ID
	ID=Trim(Request("ID"))
	if ID="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>请指定广告ID</li>"
		exit sub
	end if
	if Instr(ID,",")>0 then
		dim arrID,i
		arrID=split(ID,",")
		for i=0 to Ubound(arrID)
			conn.execute "Update Advertisement set IsSelected=False Where ID=" & CLng(arrID(i))
		next
	else
		conn.execute "Update Advertisement set IsSelected=False Where ID=" & CLng(ID)
	end if
	response.redirect strFileName
end sub

sub MoveAdvertisement()
	dim ID,MoveChannelID
	ID=Trim(Request("ID"))
	if ID="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>请指定广告ID</li>"
		exit sub
	end if
	MoveChannelID=Trim(Request("ChannelID"))
	if Instr(ID,",")>0 then
		dim arrID,i
		arrID=split(ID,",")
		for i=0 to Ubound(arrID)
			conn.execute "Update Advertisement set ChannelID = "& MoveChannelID & " where ID=" & CLng(arrID(i))
		next
	else
		conn.execute "Update Advertisement set ChannelID = "& MoveChannelID & " where ID=" & CLng(ID)
	end if
	response.redirect strFileName
end sub

sub DelAD()
	dim ID
	ID=Trim(Request("ID"))
	if ID="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>请指定广告ID</li>"
		exit sub
	end if
	if Instr(ID,",")>0 then
		dim arrID,i
		arrID=split(ID,",")
		for i=0 to Ubound(arrID)
		conn.execute "delete from Advertisement where ID=" & CLng(arrID(i))
		next
	else
		conn.execute "delete from Advertisement where ID=" & CLng(ID)
	end if
	response.redirect strFileName
end sub
%>
