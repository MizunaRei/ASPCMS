<%@language=vbscript codepage=936 %>
<%
option explicit
response.buffer=true	
Const PurviewLevel=2
Const CheckChannelID=2
Const PurviewLevel_Article=3
%>
<!--#include file="inc/conn.asp"-->
<!--#include file="inc/config.asp"-->
<!--#include file="Admin_ChkPurview.asp"-->
<!--#include file="inc/function.asp"-->
<!--#include file="inc/admin_code_article.asp"-->
<%
dim ClassID,SpecialID
dim SkinID,LayoutID,SkinCount,LayoutCount,ClassMaster,BrowsePurview,AddPurview
ClassID=session("ClassID_Article")
SpecialID=session("SpecialID_Article")
if ClassID="" then
	ClassID=0
else
	ClassID=Clng(ClassID)
end if
if SpecialID="" then
	SpecialID=0
else
	SpecialID=Clng(SpecialID)
end if
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="GENERATOR" content="Microsoft FrontPage 3.0">
<title>添加文章（高级模式）</title>
<link rel="stylesheet" type="text/css" href="Admin_Style.css">
<script language = "JavaScript">
function AddItem(strFileName){
  document.myform.IncludePic.checked=true;
  document.myform.DefaultPicUrl.value=strFileName;
  document.myform.DefaultPicList.options[document.myform.DefaultPicList.length]=new Option(strFileName,strFileName);
  document.myform.DefaultPicList.selectedIndex+=1;
  if(document.myform.UploadFiles.value==''){
	document.myform.UploadFiles.value=strFileName;
  }
  else{
    document.myform.UploadFiles.value=document.myform.UploadFiles.value+"|"+strFileName;
  }
}
function selectPaginationType()
{
  document.myform.PaginationType.selectedIndex=2;
}
function CheckForm()
{
  if (editor.EditMode.checked==true)
	  document.myform.Content.value=editor.HtmlEdit.document.body.innerText;
  else
	  document.myform.Content.value=editor.HtmlEdit.document.body.innerHTML; 

  if (document.myform.ClassID.value=="")
  {
    alert("文章所属栏目不能指定为含有子栏目的栏目！");
	document.myform.ClassID.focus();
	return false;
  }
  if (document.myform.ClassID.value==0)
  {
    alert("文章所属栏目不能指定为外部栏目！");
	document.myform.ClassID.focus();
	return false;
  }

  if (document.myform.Title.value=="")
  {
    alert("文章标题不能为空！");
	document.myform.Title.focus();
	return false;
  }
  
    if (document.myform.SpecialID.value==0)
  {
    alert("请指定文章所属课程！");
	document.myform.SpecialID.focus();
	return false;
  }

  
  if (document.myform.Key.value=="")
  {
    alert("关键字不能为空！");
	document.myform.Key.focus();
	return false;
  }
  if (document.myform.Content.value=="")
  {
    alert("文章内容不能为空！");
	editor.HtmlEdit.focus();
	return false;
  }
  if (document.myform.Content.value.length>65536)
  {
    alert("文章内容太长，超出了ACCESS数据库的限制（64KB）！建议将文章分成几部分录入。");
	editor.HtmlEdit.focus();
	return false;
  }
  return true;  
}
</script>

</head>

<body>
<form method="POST" name="myform" onSubmit="return CheckForm();" action="Admin_ArticleSave.asp" target="_self">
  <table border="0" align="center" cellpadding="0" cellspacing="0" class="border">
    <tr class="title">
      <td height="22" align="center"><b>添 加 文 章（高级模式）</b></td>
    </tr>
    <tr align="center">
      <td>
	<table width="100%" border="0" cellpadding="2" cellspacing="1">
          <tr class="tdbg"> 
            <td width="100" align="right"><strong>所属栏目：</strong></td>
            <td><select name='ClassID'>
                <%call Admin_ShowClass_Option(3,ClassID)%>
              </select>
              </td>
            <td><%response.write "<font color='#FF0000'><strong>注意：</strong></font><font color='#0000FF'>1、不能指定为含有子栏目的栏目，或者外部栏目"
			if AdminPurview=2 and AdminPurview_Article=3 then
              response.write "<br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;2、你只能在<font color='#FF0000'>红色栏目</font>及其子栏目中发表文章</font>"
			end if%>
            </td>
          </tr>
          <tr class="tdbg"> 
            <td width="100" align="right"><strong>所属课程：</strong></td>
            <td colspan="2"><%call Admin_ShowSpecial_Option(1,SpecialID)%>  <font color="#FF0000">*</font>  </td>
          </tr>
          <tr class="tdbg"> 
            <td width="100" align="right"><strong>文章标题：</strong></td>
            <td colspan="2"><input name="Title" type="text" id="Title" value="" size="50" maxlength="255"> 
              <font color="#FF0000">*</font> <select name="TitleFontColor" id="TitleFontColor">
                <option value="" selected>颜色</option>
                <OPTION value="">默认</OPTION>
                <OPTION value="#000000" style="background-color:#000000"></OPTION>
                <OPTION value="#FFFFFF" style="background-color:#FFFFFF"></OPTION>
                <OPTION value="#008000" style="background-color:#008000"></OPTION>
                <OPTION value="#800000" style="background-color:#800000"></OPTION>
                <OPTION value="#808000" style="background-color:#808000"></OPTION>
                <OPTION value="#000080" style="background-color:#000080"></OPTION>
                <OPTION value="#800080" style="background-color:#800080"></OPTION>
                <OPTION value="#808080" style="background-color:#808080"></OPTION>
                <OPTION value="#FFFF00" style="background-color:#FFFF00"></OPTION>
                <OPTION value="#00FF00" style="background-color:#00FF00"></OPTION>
                <OPTION value="#00FFFF" style="background-color:#00FFFF"></OPTION>
                <OPTION value="#FF00FF" style="background-color:#FF00FF"></OPTION>
                <OPTION value="#FF0000" style="background-color:#FF0000"></OPTION>
                <OPTION value="#0000FF" style="background-color:#0000FF"></OPTION>
                <OPTION value="#008080" style="background-color:#008080"></OPTION>
              </select> <select name="TitleFontType" id="TitleFontType">
                <option value="0" selected>字形</option>
                <option value="1">粗体</option>
                <option value="2">斜体</option>
                <option value="3">粗+斜</option>
                <option value="0">规则</option>
              </select> </td>
          </tr>
          <tr class="tdbg"> 
            <td width="100" align="right"><strong>关键字：</strong></td>
            <td colspan="2"><input name="Key" type="text"
           id="Key" value="<%=session("Key")%>" size="50" maxlength="255"> <font color="#FF0000">*</font><br>
              <font color="#0000FF">用来查找相关文章，可输入多个关键字，中间用<font color="#FF0000">“|”</font>分开。不能出现&quot;&quot;'*?,.()等字符。</font></td>
          </tr>
          <tr class="tdbg"> 
            <td width="100" align="right"><strong>作者：</strong></td>
            <td colspan="2"> 姓名： 
              <input name="AuthorName" type="text"
           id="AuthorName" value="<%=session("AuthorName")%>" size="20" maxlength="30"> 
              &nbsp;&nbsp;&nbsp;&nbsp;Email： 
              <input name="AuthorEmail" type="text" id="AuthorEmail" value="<%=session("AuthorEmail")%>" size="40" maxlength="100"> 
            </td>
          </tr>
          <tr class="tdbg"> 
            <td width="100" align="right"><strong>转贴自：</strong></td>
            <td colspan="2">名称： 
              <input name="CopyFromName" type="text"
           id="CopyFromName" value="<%=session("CopyFromName")%>" size="20" maxlength="50"> 
              &nbsp;&nbsp;&nbsp;&nbsp;地 址： 
              <input name="CopyFromUrl" type="text" id="CopyFromUrl2" value="<%=session("CopyFromUrl")%>" size="40" maxlength="200"> 
            </td>
          </tr>
          <tr class="tdbg"> 
            <td width="100" align="right" valign="middle"><p><strong>文章内容：</strong></p>
              <p align="left"><font color="#006600"><% if EnableSaveRemote="Yes" then%>&middot;　如果是从其它网站上复制内容，并且内容中包含有图片，本系统将会把图片复制到本站服务器上，系统会因复制图片的大小而影响速度，请稍候（此功能需要服务器安装了IE5.5以上版本才有效）。<%end if%><br>
                <br>
                &middot;　换行请按Shift+Enter</font><br>
                <font color="#006600">&middot;　另起一段请按Enter</font></p></td>
            <td colspan="2"><textarea name="Content" style="display:none"></textarea> 
              <iframe ID="editor" src="editor.asp" frameborder=1 scrolling=no width="600" height="405"></iframe> 
            </td>
          </tr>
          <tr class="tdbg"> 
            <td width="100" align="right"><strong>内容分页方式：</strong></td>
            <td colspan="2"><select name="PaginationType" id="PaginationType">
                <option value="0" <%if session("PaginationType")=0 then response.write " selected"%>>不分页</option>
                <option value="1" <%if session("PaginationType")=1 then response.write " selected"%>>自动分页</option>
                <option value="2" <%if session("PaginationType")=2 then response.write " selected"%>>手动分页</option>
              </select> &nbsp;&nbsp;&nbsp;&nbsp;<strong><font color="#0000FF">注：</font></strong><font color="#0000FF">手动分页符标记为“</font><font color="#FF0000">[NextPage]</font><font color="#0000FF">”，注意大小写</font></td>
          </tr>
          <tr class="tdbg">
            <td align="right">&nbsp;</td>
            <td colspan="2">自动分页时的每页大约字符数（包含HTML标记）：<strong> 
              <input name="MaxCharPerPage" type="text" id="MaxCharPerPage" value="10000" size="8" maxlength="8">
              </strong></td>
          </tr>
          <tr class="tdbg"> 
            <td width="100" align="right"><strong>文章阅读等级：</strong></td>
            <td colspan="2"><select name="ReadLevel" id="ReadLevel">
                <option value="9999" <%if session("ReadLevel")=9999 then response.write " selected"%>>游客</option>
                <option value="999" <%if session("ReadLevel")=999 then response.write " selected"%>>注册用户</option>
                <option value="99" <%if session("ReadLevel")=99 then response.write " selected"%>>收费用户</option>
                <option value="9" <%if session("ReadLevel")=9 then response.write " selected"%>>VIP用户</option>
                <option value="5" <%if session("ReadLevel")=5 then response.write " selected"%>>管理员</option>
              </select>
              &nbsp;&nbsp;&nbsp;<font color="#0000FF">只有具有相应权限的人才能阅读此文章。</font></td>
          </tr>
          <tr class="tdbg"> 
            <td width="100" align="right"><strong>文章阅读点数：</strong></td>
            <td colspan="2"><input name="ReadPoint" type="text" id="ReadPoint" value="<%=session("ReadPoint")%>" size="5" maxlength="3"> 
              &nbsp;&nbsp;&nbsp;&nbsp; <font color="#0000FF">如果大于0，则用户阅读此文章时将消耗相应点数。（对游客和管理员无效）</font></td>
          </tr>
          <tr class="tdbg"> 
            <td width="100" align="right"><strong>包含图片：</strong></td>
            <td colspan="2"><input name="IncludePic" type="checkbox" id="IncludePic" value="yes">
              是<font color="#0000FF">（如果选中的话会在标题前面显示[图文]）</font></td>
          </tr>
          <tr class="tdbg"> 
            <td width="100" align="right"><strong>首页图片：</strong></td>
            <td colspan="2"><input name="DefaultPicUrl" type="text" id="DefaultPicUrl" size="56" maxlength="200">
              用于在首页的图片文章处显示 <br>
              直接从上传图片中选择： 
              <select name="DefaultPicList" id="DefaultPicList" onChange="DefaultPicUrl.value=this.value;">
                <option selected>不指定首页图片</option>
              </select> <input name="UploadFiles" type="hidden" id="UploadFiles"> 
            </td>
          </tr>
<%if AdminPurview=1 or AdminPurview_Article<=2 then%>
          <tr class="tdbg"> 
            <td width="100" align="right"><strong>文章性质：</strong></td>
            <td colspan="2"><input name="OnTop" type="checkbox" id="OnTop" value="yes">
              固顶文章&nbsp;&nbsp;&nbsp;&nbsp; <input name="Hot" type="checkbox" id="Hot" value="yes" onClick="javascript:document.myform.Hits.value=<%=HitsOfHot%>">
              热点文章&nbsp;&nbsp;&nbsp;&nbsp; <input name="Elite" type=checkbox id="Elite" value="yes">
              推荐文章&nbsp;&nbsp;&nbsp;&nbsp;文章评分等级： 
              <select name="Stars" id="Stars">
                <option value="5">★★★★★</option>
                <option value="4">★★★★</option>
                <option value="3" selected>★★★</option>
                <option value="2">★★</option>
                <option value="1">★</option>
                <option value="0">无</option>
              </select></td>
          </tr>
          <tr class="tdbg"> 
            <td width="100" align="right"><strong>点击数初始值：</strong></td>
            <td colspan="2"><input name="Hits" type="text" id="Hits" value="0" size="10" maxlength="10"> 
              &nbsp;&nbsp; <font color="#0000FF">这功能是提供给管理员作弊用的。不过尽量不要用呀！^_^</font></td>
          </tr>
          <tr class="tdbg"> 
            <td width="100" align="right"><strong>录入时间：</strong></td>
            <td colspan="2"><input name="UpdateTime" type="text" id="UpdateTime" value="<%=now()%>" maxlength="50">
              时间格式为“年-月-日 时:分:秒”，如：<font color="#0000FF">2003-5-12 12:32:47</font></td>
          </tr>
<%end if%>
          <tr class="tdbg"> 
            <td width="100" align="right"><strong>配色模板：</strong></td>
            <td colspan="2"><%call Admin_ShowSkin_Option(session("SkinID"))%>&nbsp;相关模板中包含CSS、颜色、图片等信息</td>
          </tr>
          <tr class="tdbg"> 
            <td width="100" align="right"><strong>版面设计模板：</strong></td>
            <td colspan="2"><%call Admin_ShowLayout_Option(3,session("LayoutID"))%>&nbsp;相关模板中包含了版面设计的版式等信息</td>
          </tr>
<%if AdminPurview=1 or AdminPurview_Article<=2 then%>
          <tr class="tdbg"> 
            <td width="100" align="right"><strong>立即发布：</strong></td>
            <td colspan="2"><input name="Passed" type="checkbox" id="Passed" value="yes" checked>
              是<font color="#0000FF">（如果选中的话将直接发布，否则审核后才能发布。）</font></td>
          </tr>
<%end if%>
        </table>
      </td>
    </tr>
  </table>
  <div align="center"> 
    <p> 
      <input name="Action" type="hidden" id="Action" value="Add2">
      <input
  name="Add" type="submit"  id="Add" value=" 添 加 " onClick="document.myform.action='Admin_ArticleSave.asp';document.myform.target='_self';" style="cursor:hand;">
      &nbsp; 
      <input
  name="Preview" type="submit"  id="Preview" value=" 预 览 " onClick="document.myform.action='Admin_ArticlePreview.asp';document.myform.target='_blank';" style="cursor:hand;">
      &nbsp;
      <input name="Cancel" type="button" id="Cancel" value=" 取 消 " onClick="window.location.href='Admin_ArticleManage.asp'" style="cursor:hand;">
    </p>
  </div>
</form>
</body>
</html>
<%
call CloseConn()
%>
