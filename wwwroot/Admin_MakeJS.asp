<%@language=vbscript codepage=936 %>
<%
option explicit
response.buffer=true	
Const PurviewLevel=2    '操作权限
Const CheckChannelID=0    '所属频道，0为不检测所属频道
Const PurviewLevel_Others="JS"
%>
<!--#include file="inc/conn.asp"-->
<!--#include file="inc/config.asp"-->
<!--#include file="Admin_ChkPurview.asp"-->
<!--#include file="inc/admin_code_article.asp"-->
<%
dim ClassID,SpecialID
%>
<html>
<head>
<title>JS代码管理</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="Admin_Style.css" rel="stylesheet" type="text/css">
<script language="javascript">
<!--
function copy() {
document.myform.JsCode.focus();
document.myform.JsCode.select();
textRange = document.myform.JsCode.createTextRange();
textRange.execCommand("Copy");
}
// -->
</script>
</head>

<body leftmargin="2" topmargin="0" marginwidth="0" marginheight="0">
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="border">
  <tr class="topbg"> 
    <td height="22" colspan="2" align="center"><strong>JS 代 码 管 理</strong></td>
  </tr>
  <tr class="tdbg"> 
    <td width="70" height="30"><strong>管理导航：</strong></td>
    <td height="30"><a href="Admin_MakeJS.asp?Action=JS_Common">普通文章的JS代码</a> 
      | <a href="Admin_MakeJS.asp?Action=JS_Pic">首页图文的JS代码</a></td>
  </tr>
</table>
<%
dim Action
Action=trim(request("Action"))
if Action="JS_Common" then
	call JS_Common()
elseif Action="JS_Pic" then
	call JS_Pic
end if
call CloseConn()


sub JS_Common()
%>
<form action="" method="post" name="myform" id="myform">
  <table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="border">
    <tr class="title"> 
      <td height="22" colspan="2" align="center"><strong>普通文章的JS代码</strong></td>
    </tr>
    <tr class="tdbg"> 
      <td height="25" align="right">所属栏目：</td>
      <td height="25"><select name="ClassID" id="ClassID">
          <option value="0">不指定栏目</option>
          <%call Admin_ShowClass_Option(2,0)%>
        </select> <font color="#0000FF">&nbsp;&nbsp;不能指定为外部栏目</font>&nbsp;&nbsp;&nbsp;&nbsp; 
        <input name="IncludeChild" type="checkbox" id="IncludeChild" value="True" checked>
        包含子栏目</td>
    </tr>
    <tr class="tdbg">
      <td height="25" align="right">所属专题：</td>
      <td height="25">
        <%call Admin_ShowSpecial_Option(1,SpecialID)%>
      </td>
    </tr>
    <tr class="tdbg"> 
      <td height="25" align="right">文章数目：</td>
      <td height="25"><input name="ArticleNum" type="text" value="10" size="5" maxlength="3"> 
        &nbsp;&nbsp;&nbsp;&nbsp;<font color="#FF0000">如果为空，则显示所有文章</font></td>
    </tr>
    <tr class="tdbg"> 
      <td height="25" align="right">显示类型：</td>
      <td height="25"><select name="ShowType" id="select">
          <option value="1" selected>文章标题列表</option>
          <option value="2">文章标题+部分内容</option>
        </select> &nbsp;&nbsp;&nbsp;&nbsp;分栏： 
        <select name="ShowCols" id="ShowCols">
          <option value="1" selected>一栏</option>
          <option value="2">两栏</option>
          <option value="3">三栏</option>
        </select> </td>
    </tr>
    <tr class="tdbg"> 
      <td height="50" align="right">显示内容：</td>
      <td height="50"><input name="ShowProperty" type="checkbox" id="ShowProperty" value="True" checked>
        文章属性&nbsp;&nbsp;&nbsp;&nbsp; <input name="ShowClassName" type="checkbox" id="ShowClassName" value="True">
        所属栏目&nbsp;&nbsp;&nbsp;&nbsp; <input name="ShowIncludePic" type="checkbox" id="ShowIncludePic" value="True" checked>
        图文标志&nbsp;&nbsp;&nbsp;&nbsp; <input name="ShowTitle" type="checkbox" id="ShowTitle" value="True" checked disabled>
        文章标题 
        <input type="checkbox" name="checkbox2" value="checkbox">
        文章内容<br> <input name="ShowUpdateTime" type="checkbox" id="ShowUpdateTime" value="True">
        更新时间&nbsp;&nbsp;&nbsp;&nbsp; <input name="ShowHits" type="checkbox" id="ShowHits" value="True">
        点击次数&nbsp;&nbsp;&nbsp;&nbsp; <input name="ShowAuthor" type="checkbox" id="ShowAuthor" value="True">
        作者&nbsp;&nbsp;&nbsp;&nbsp; <input name="ShowHot" type="checkbox" id="ShowHot" value="True" checked>
        热点文章标志 
        <input name="ShowMore" type="checkbox" id="ShowMore" value="True">
        “更多……”字样</td>
    </tr>
    <tr class="tdbg"> 
      <td height="25" align="right">标题最多字符数：</td>
      <td height="25"><input name="TitleMaxLen" type="text" id="TitleMaxLen" value="30" size="5" maxlength="3"> 
        &nbsp;&nbsp;&nbsp;&nbsp;<font color="#FF0000">如果为空，则显示完整标题。字母算一个字符汉字算两个字符。</font></td>
    </tr>
    <tr class="tdbg"> 
      <td height="25" align="right">文章内容字符数：</td>
      <td height="25"><input name="ContentMaxLen" type="text" id="ContentMaxLen" value="200" size="5" maxlength="3"> 
        &nbsp;&nbsp;&nbsp;&nbsp;<font color="#FF0000">只有当显示类型设为“标题+内容”时才有效</font></td>
    </tr>
    <tr class="tdbg"> 
      <td height="25" align="right">文章属性：</td>
      <td height="25"> <input name="Hot" type="checkbox" id="Hot" value="True">
        热门文章&nbsp;&nbsp;&nbsp;&nbsp; <input name="Elite" type="checkbox" id="Elite" value="True">
        推荐文章 &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<font color="#FF0000">如果都不选，将显示所有文章</font></td>
    </tr>
    <tr class="tdbg"> 
      <td height="25" align="right">日期范围：</td>
      <td height="25">只显示最近 
        <input name="DateNum" type="text" id="DateNum" value="10" size="5" maxlength="3">
        天内的文章&nbsp;&nbsp;&nbsp;&nbsp;<font color="#FF0000">如果为空，则显示所有天数的文章</font></td>
    </tr>
    <tr class="tdbg"> 
      <td height="25" align="right">排序字段：</td>
      <td height="25"><select name="OrderField" id="OrderField">
          <option value="ArticleID" selected>文章ID</option>
          <option value="UpdateTime">更新时间</option>
          <option value="Hits">点击次数</option>
        </select> &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;排序方法： 
        <select name="OrderType" id="OrderType">
          <option value="asc">升序</option>
          <option value="desc" selected>降序</option>
        </select></td>
    </tr>
    <tr class="tdbg"> 
      <td height="40" colspan="2" align="center"><input name="SystemPath" type="hidden" id="SystemPath" value="<%= "http://"&request.servervariables("server_name")&replace(lcase(request.servervariables("url")),"admin_makejs.asp","") %>"> 
        <input name="MakeJS" type="button" id="MakeJS" onclick="makejs();" value="生成JS代码"> 
        &nbsp;&nbsp; <input name="Copy" type="button" id="Copy" value="复制到剪贴板 " onclick="copy();"> 
        &nbsp;&nbsp; <input type="reset" name="Reset" value="恢复默认值"></td>
    </tr>
    <tr class="tdbg"> 
      <td height="40" colspan="2" align="center"><textarea name="JsCode" cols="80" rows="10" id="JsCode"></textarea></td>
    </tr>
    <tr class="tdbg"> 
      <td height="60" colspan="2">使用方法：<br> &nbsp;&nbsp;&nbsp;&nbsp;首先设定各选项，然后点“生成JS代码”，最后将生成的JS代码复制到网页代码的相应位置处即可。可参看js.htm文件中范例。<font color="#0000FF">有关字体、字号、颜色及行距等属性的设置，请自行在调用代码的页面用CSS进行设置。</font></td>
    </tr>
  </table>
</form>
<script language="JavaScript" type="text/JavaScript">
function makejs()
{
if(document.myform.ClassID.value=="")
{
	alert("文章栏目不能指定外部栏目！");
	document.myform.ClassID.focus();
	return false;
}
var strJS;
strJS="<!--代码开始-->\n<";
strJS+="script language='JavaScript' type='text/JavaScript' src='";
strJS+=document.myform.SystemPath.value+"article_js.asp"
strJS+="?ClassID="+document.myform.ClassID.value;
strJS+="&IncludeChild="+document.myform.IncludeChild.checked;
strJS+="&SpecialID="+document.myform.SpecialID.value;
strJS+="&ArticleNum="+document.myform.ArticleNum.value;
strJS+="&ShowType="+document.myform.ShowType.value;
strJS+="&ShowCols="+document.myform.ShowCols.value;
strJS+="&ShowProperty="+document.myform.ShowProperty.checked;
strJS+="&ShowClassName="+document.myform.ShowClassName.checked;
strJS+="&ShowIncludePic="+document.myform.ShowIncludePic.checked;
strJS+="&ShowTitle="+document.myform.ShowTitle.checked;
strJS+="&ShowUpdateTime="+document.myform.ShowUpdateTime.checked;
strJS+="&ShowHits="+document.myform.ShowHits.checked;
strJS+="&ShowAuthor="+document.myform.ShowAuthor.checked;
strJS+="&ShowHot="+document.myform.ShowHot.checked;
strJS+="&ShowMore="+document.myform.ShowMore.checked;
strJS+="&TitleMaxLen="+document.myform.TitleMaxLen.value;
strJS+="&ContentMaxLen="+document.myform.ContentMaxLen.value;
strJS+="&Hot="+document.myform.Hot.checked;
strJS+="&Elite="+document.myform.Elite.checked;
strJS+="&DateNum="+document.myform.DateNum.value;
strJS+="&OrderField="+document.myform.OrderField.value;
strJS+="&OrderType="+document.myform.OrderType.value;
strJS+="'></";
strJS+="script";
strJS+=">\n<!--代码结束-->";
document.myform.JsCode.value=strJS;
}
</script>
<%
end sub

sub JS_Pic()
%>
<form action="" method="post" name="myform" id="myform">
  <table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="border">
    <tr class="title"> 
      <td height="22" colspan="2" align="center"><strong>首页图文的JS代码</strong></td>
    </tr>
    <tr class="tdbg"> 
      <td height="25" align="right">所属栏目：</td>
      <td height="25"><select name="ClassID" id="ClassID"><option value="0">不指定栏目</option>
          <%call Admin_ShowClass_Option(2,0)%>
        </select>
        <font color="#0000FF">&nbsp;&nbsp;不能指定为外部栏目</font>&nbsp;&nbsp;&nbsp;&nbsp;
        <input name="IncludeChild" type="checkbox" id="IncludeChild" value="True" checked>
        包含子栏目</td>
    </tr>
    <tr class="tdbg">
      <td height="25" align="right">所属专题：</td>
      <td height="25">
        <%call Admin_ShowSpecial_Option(1,SpecialID)%>
      </td>
    </tr>
    <tr class="tdbg"> 
      <td height="25" align="right">文章数目：</td>
      <td height="25"><input name="ArticleNum" type="text" value="10" size="5" maxlength="3"> 
        &nbsp;&nbsp;&nbsp;&nbsp;<font color="#FF0000">如果为空，则显示所有文章</font></td>
    </tr>
    <tr class="tdbg"> 
      <td height="25" align="right">显示类型：</td>
      <td height="25"><select name="ShowType" id="select">
          <option value="3" selected>图片+标题</option>
          <option value="4">图片+标题+内容</option>
        </select>
      </td>
    </tr>
    <tr class="tdbg"> 
      <td height="50" align="right">显示内容：</td>
      <td height="50"> 
        <input name="ShowClassName" type="checkbox" id="ShowClassName" value="True">
        文章栏目&nbsp;&nbsp;&nbsp;&nbsp; 
        <input name="ShowTitle" type="checkbox" id="ShowTitle" value="True" checked disabled>
        文章标题
        <input type="checkbox" name="checkbox2" value="checkbox">
        文章内容&nbsp;&nbsp;&nbsp;&nbsp; 
        <input name="ShowUpdateTime" type="checkbox" id="ShowUpdateTime" value="True">
        更新时间<br>
        <input name="ShowHits" type="checkbox" id="ShowHits" value="True">
        点击次数&nbsp;&nbsp;&nbsp;&nbsp; <input name="ShowAuthor" type="checkbox" id="ShowAuthor" value="True">
        作者&nbsp;&nbsp;&nbsp;&nbsp; <input name="ShowHot" type="checkbox" id="ShowHot" value="True" checked>
        热点文章标志
        <input name="ShowMore" type="checkbox" id="ShowMore" value="True">
        “更多……”字样</td>
    </tr>
    <tr class="tdbg"> 
      <td height="25" align="right">标题最多字符数：</td>
      <td height="25"><input name="TitleMaxLen" type="text" id="TitleMaxLen" value="30" size="5" maxlength="3"> 
        &nbsp;&nbsp;&nbsp;&nbsp;<font color="#FF0000">如果为空，则显示完整标题。字母算一个字符汉字算两个字符。</font></td>
    </tr>
    <tr class="tdbg">
      <td height="25" align="right">文章内容字符数：</td>
      <td height="25"><input name="ContentMaxLen" type="text" id="ContentMaxLen" value="200" size="5" maxlength="3">
        &nbsp;&nbsp;&nbsp;&nbsp;<font color="#FF0000">只有当显示类型设为“图片+标题+内容”时才有效</font></td>
    </tr>
    <tr class="tdbg"> 
      <td height="25" align="right">文章属性：</td>
      <td height="25"> <input name="Hot" type="checkbox" id="Hot" value="True">
        热门文章&nbsp;&nbsp;&nbsp;&nbsp; <input name="Elite" type="checkbox" id="Elite" value="True">
        推荐文章 &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<font color="#FF0000">如果都不选，将显示所有文章</font></td>
    </tr>
    <tr class="tdbg"> 
      <td height="25" align="right">日期范围：</td>
      <td height="25">只显示最近 
        <input name="DateNum" type="text" id="DateNum" value="10" size="5" maxlength="3">
        天内的文章&nbsp;&nbsp;&nbsp;&nbsp;<font color="#FF0000">如果为空，则显示所有天数的文章</font></td>
    </tr>
    <tr class="tdbg"> 
      <td height="25" align="right">排序字段：</td>
      <td height="25"><select name="OrderField" id="OrderField">
          <option value="ArticleID" selected>文章ID</option>
          <option value="UpdateTime">更新时间</option>
          <option value="Hits">点击次数</option>
        </select> &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;排序方法： 
        <select name="OrderType" id="OrderType">
          <option value="asc">升序</option>
          <option value="desc" selected>降序</option>
        </select></td>
    </tr>
    <tr class="tdbg"> 
      <td height="25" align="right">首页图片设置：</td>
      <td height="25">&nbsp;宽度： 
        <input name="ImgWidth" type="text" id="ImgWidth" value="180" size="5" maxlength="3">
        像素&nbsp;&nbsp;&nbsp;&nbsp;高度： <input name="ImgHeight" type="text" id="ImgHeight" value="120" size="5" maxlength="3">
        像素</td>
    </tr>
    <tr class="tdbg"> 
      <td height="40" colspan="2" align="center"><input name="SystemPath" type="hidden" id="SystemPath" value="<%= "http://"&request.servervariables("server_name")&replace(lcase(request.servervariables("url")),"admin_makejs.asp","") %>"> 
        <input name="MakeJS" type="button" id="MakeJS" onclick="makejs();" value="生成JS代码"> 
        &nbsp;&nbsp; <input name="Copy" type="button" id="Copy" value="复制到剪贴板 " onclick="copy();"> 
        &nbsp;&nbsp; <input type="reset" name="Reset" value="恢复默认值"></td>
    </tr>
    <tr class="tdbg"> 
      <td height="40" colspan="2" align="center"><textarea name="JsCode" cols="80" rows="10" id="JsCode"></textarea></td>
    </tr>
    <tr class="tdbg"> 
      <td height="60" colspan="2">使用方法：<br>
        &nbsp;&nbsp;&nbsp;&nbsp;首先设定各选项，然后点“生成JS代码”，最后将生成的JS代码复制到网页代码的相应位置处即可。可参看js.htm文件中范例。<font color="#0000FF">有关字体、字号、颜色及行距等属性的设置，请自行在调用代码的页面用CSS进行设置。</font></td>
    </tr>
  </table>
</form>
<script language="JavaScript" type="text/JavaScript">
function makejs()
{
if(document.myform.ClassID.value=="")
{
	alert("文章栏目不能指定外部栏目！");
	document.myform.ClassID.focus();
	return false;
}
var strJS;
strJS="<!--代码开始-->\n<";
strJS+="script language='JavaScript' type='text/JavaScript' src='";
strJS+=document.myform.SystemPath.value+"article_js.asp"
strJS+="?ClassID="+document.myform.ClassID.value;
strJS+="&IncludeChild="+document.myform.IncludeChild.checked;
strJS+="&SpecialID="+document.myform.SpecialID.value;
strJS+="&ArticleNum="+document.myform.ArticleNum.value;
strJS+="&ShowType="+document.myform.ShowType.value;
strJS+="&ShowClassName="+document.myform.ShowClassName.checked;
strJS+="&ShowTitle="+document.myform.ShowTitle.checked;
strJS+="&ShowUpdateTime="+document.myform.ShowUpdateTime.checked;
strJS+="&ShowHits="+document.myform.ShowHits.checked;
strJS+="&ShowAuthor="+document.myform.ShowAuthor.checked;
strJS+="&ShowHot="+document.myform.ShowHot.checked;
strJS+="&ShowMore="+document.myform.ShowMore.checked;
strJS+="&TitleMaxLen="+document.myform.TitleMaxLen.value;
strJS+="&ContentMaxLen="+document.myform.ContentMaxLen.value;
strJS+="&Hot="+document.myform.Hot.checked;
strJS+="&Elite="+document.myform.Elite.checked;
strJS+="&DateNum="+document.myform.DateNum.value;
strJS+="&OrderField="+document.myform.OrderField.value;
strJS+="&OrderType="+document.myform.OrderType.value;
strJS+="&ImgWidth="+document.myform.ImgWidth.value;
strJS+="&ImgHeight="+document.myform.ImgHeight.value;
strJS+="'></";
strJS+="script";
strJS+=">\n<!--代码结束-->";
document.myform.JsCode.value=strJS;
}
</script>
<%
end sub
%>
</body>
</html>
