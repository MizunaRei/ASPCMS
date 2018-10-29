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
<title>添加文章（简洁模式）</title>
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
  document.myform.PaginationType.value=2;
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
    alert("文章内容太长，超出了ACCESS数据库的限制（64K）！建议将文章分成几部分录入。");
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
      <td height="22" align="center"><b>添 加 文 章（简洁模式）</b></td>
    </tr>
    <tr align="center">
      <td>
	<table width="100%" border="0" cellpadding="2" cellspacing="1">
          <tr class="tdbg"> 
            <td width="100" align="right"><strong>所属栏目：</strong></td>
            <td><select name='ClassID'><%call Admin_ShowClass_Option(3,ClassID)%></select>              
            </td>
            <td><%response.write "<font color='#FF0000'><strong>注意：</strong></font><font color='#0000FF'>1、不能指定为含有子栏目的栏目，或者外部栏目"
			if AdminPurview=2 and AdminPurview_Article=3 then
              response.write "<br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;2、你只能在<font color='#FF0000'>红色栏目</font>及其子栏目中发表文章</font>"
			end if%></td>
          </tr>
          <tr class="tdbg"> 
            <td width="100" align="right"><strong>所属课程：</strong></td>
            <td colspan="2"><%call Admin_ShowSpecial_Option(1,SpecialID)%> <font color="#FF0000">*</font>  </td>
          </tr>
          <tr class="tdbg"> 
            <td width="100" align="right"><strong>文章标题：</strong></td>
            <td colspan="2"><input name="Title" type="text" id="Title" value="" size="50" maxlength="255"> 
              <font color="#FF0000">*</font></td>
          </tr>
          <tr class="tdbg"> 
            <td width="100" align="right"><strong>关键字：</strong></td>
            <td colspan="2"><input name="Key" type="text"
           id="Key" value="<%=session("Key")%>" size="50" maxlength="255"> <font color="#FF0000">*</font><br>
              <font color="#0000FF">用来查找相关文章，可输入多个关键字，中间用<font color="#FF0000">“|”</font>隔开。不能出现&quot;&quot;'*?,.()等字符。</font></td>
          </tr>
          <tr class="tdbg"> 
            <td width="100" align="right"><strong>作者：</strong></td>
            <td colspan="2"> <input name="Author" type="text"
           id="Author" value="<%=session("Author")%>" size="50" maxlength="100"> 
            </td>
          </tr>
          <tr class="tdbg"> 
            <td width="100" align="right"><strong>转贴自：</strong></td>
            <td colspan="2"> <input name="CopyFrom" type="text"
           id="CopyFrom" value="<%=session("CopyFrom")%>" size="50" maxlength="100"> 
            </td>
          </tr>
          <tr class="tdbg"> 
            <td width="100" align="right" valign="middle"><p><strong>文章内容：</strong></p>
              <p align="left"><font color="#006600"><% if EnableSaveRemote="Yes" then%>&middot;　如果是从其它网站上复制内容，并且内容中包含有图片，本系统将会把图片复制到本站服务器上，系统会因复制图片的大小而影响速度，请稍候（此功能需要服务器安装了IE5.5以上版本才有效）。<%end if%><br>
                <br>
                &middot;　换行请按Shift+Enter</font><br>
                <font color="#006600">&middot;　另起一段请按Enter</font><br>
              </p></td>
            <td colspan="2"><textarea name="Content" style="display:none"></textarea> 
              <iframe ID="editor" src="editor.asp" frameborder=1 scrolling=no width="600" height="405"></iframe> 
            </td>
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
      <input name="PaginationType" type="hidden" id="PaginationType" value="0"> 
      <%dim trs
	  set trs=conn.execute("select SkinID from Skin where IsDefault=True")
	  %>
	  <input name="SkinID" type="hidden" id="SkinID" value="<%=trs(0)%>">
      <%
	  set trs=conn.execute("select LayoutID from Layout where IsDefault=True and LayoutType=3")
	  %>
      <input name="LayoutID" type="hidden" id="LayoutID" value="<%=trs(0)%>">
      <input name="Action" type="hidden" id="Action" value="Add1">
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
