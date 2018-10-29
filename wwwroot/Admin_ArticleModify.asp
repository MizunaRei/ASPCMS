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
dim ArticleID,sql,rsArticle,FoundErr,ErrMsg,PurviewChecked
dim Author,AuthorName,AuthorEmail,CopyFrom,CopyFromName,CopyFromUrl
dim ClassID,tClass,ClassName,RootID,ParentID,Depth,ParentPath,ClassMaster
dim SkinID,LayoutID,SkinCount,LayoutCount,BrowsePurview,AddPurview
ArticleID=trim(request("ArticleID"))
FoundErr=False
PurviewChecked=False
if ArticleID="" then 
	FoundErr=True
	ErrMsg=ErrMsg & "<br><li>请指定要修改的文章ID</li>"
	call WriteErrMsg()
	call CloseConn()
	response.end
else
	ArticleID=Clng(ArticleID)
end if
sql="select * from article where ArticleID=" & ArticleID & ""
Set rsArticle= Server.CreateObject("ADODB.Recordset")
rsArticle.open sql,conn,1,1
if rsArticle.bof and rsArticle.eof then
	FoundErr=True
	ErrMsg=ErrMsg & "<br><li>找不到文章</li>"
else
	ClassID=rsArticle("ClassID")
	set tClass=conn.execute("select ClassName,RootID,ParentID,Depth,ParentPath,ClassMaster From ArticleClass where ClassID=" & ClassID)
	if tClass.bof and tClass.eof then
		founderr=True
		ErrMsg=ErrMsg & "<br><li>找不到指定的栏目</li>"
	else
		ClassName=tClass(0)
		RootID=tClass(1)
		ParentID=tClass(2)
		Depth=tClass(3)
		ParentPath=tClass(4)
		ClassMaster=tClass(5)
	end if
	if rsArticle("Editor")=AdminName and rsArticle("Passed")=False then
		PurviewChecked=True
	else
		if AdminPurview=1 or AdminPurview_Article<=2 then
			PurviewChecked=True
		else
			PurviewChecked=CheckClassMaster(ClassMaster,AdminName)
			if PurviewChecked=False and ParentID>0 then
				set tClass=conn.execute("select ClassMaster from ArticleClass where ClassID in (" & ParentPath & ")")
				do while not tClass.eof
					PurviewChecked=CheckClassMaster(tClass(0),AdminName)
					if PurviewChecked=True then exit do
					tClass.movenext
				loop
			end if
		end if
		if PurviewChecked=False then
			FoundErr=True
			ErrMsg=ErrMsg & "<br><li>对不起，您的权限不够，不能修改此文！</li>"
		end if
	end if
end if
if FoundErr=True then
	call WriteErrMsg()
else
	Author=rsArticle("Author")
	CopyFrom=rsarticle("CopyFrom")
	if instr(Author,"|")>0 then
		AuthorName=left(Author,instr(Author,"|")-1)
		AuthorEmail=right(Author,len(Author)-instr(Author,"|"))
	else
		AuthorName=Author
		AuthorEmail=""
	end if
	if instr(CopyFrom,"|")>0 then
		CopyFromName=left(CopyFrom,instr(CopyFrom,"|")-1)
		CopyFromUrl=right(CopyFrom,len(CopyFrom)-instr(CopyFrom,"|"))
	else
		CopyFromName=CopyFrom
		CopyFromUrl=""
	end if
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>修改文章</title>
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
<body leftmargin="5" topmargin="10">
<form method="POST" name="myform" onSubmit="return CheckForm();" action="Admin_ArticleSave.asp?action=Modify">
  <table border="0" align="center" cellpadding="0" cellspacing="0" class="border">
    <tr class="title">
      <td height="22" align="center"><b>修 改 文 章</b></td>
    </tr>
    <tr align="center">
      <td>
	<table width="100%" border="0" cellpadding="2" cellspacing="1">
          <tr class="tdbg"> 
            <td width="100" align="right"><strong>所属栏目：</strong></td>
            <td>
             <table width="100%" border="0" cellspacing="0" cellpadding="0">
               <tr>
                 <td><%
			if AdminPurview=1 or AdminPurview_Article<=2 then
			 	response.write "<select name='ClassID'>"	
				call Admin_ShowClass_Option(3,rsArticle("ClassID"))
				response.write "</select></td><td>"
				response.write "<font color='#FF0000'><strong>注意：</strong></font><font color='#0000FF'>1、不能指定为含有子栏目的栏目，或者外部栏目"
			 else
			 	call Admin_ShowPath2(ParentPath,ClassName,Depth)
				response.write "<input type='hidden' name='ClassID' value='" & rsArticle("ClassID") & "'>"
			 end if
			 %>
                 </td>
               </tr>
             </table>
</td>
          </tr>
          <tr class="tdbg"> 
            <td width="100" align="right"><strong>所属专题：</strong></td>
            <td> 
              <%
			  if AdminPurview=1 or AdminPurview_Article<=2 then
			  	call Admin_ShowSpecial_Option(1,rsArticle("SpecialID"))
			  else
				if rsArticle("SpecialID")>0 then
					dim rsSpecial
					set rsSpecial=conn.execute("select * from Special where SpecialID=" & rsArticle("SpecialID"))
					if rsSpecial.bof and rsSpecial.eof then
						response.write "找不到所属专题！可能所属专题已经被删除！"
					else
						response.write rsSpecial("SpecialName")
					end if
					set rsSpecial=nothing
				end if
				response.write "<input type='hidden' name='SpecialID' value='" & rsArticle("SpecialID") & "'>"
			  end if%>
            </td>
          </tr>
          <tr class="tdbg"> 
            <td width="100" align="right"><strong>文章标题：</strong></td>
            <td><input name="Title" type="text"
           id="Title" value="<%=rsArticle("Title")%>" size="50" maxlength="255">
              <font color="#FF0000">*</font>
              <select name="TitleFontColor" id="TitleFontColor">
                <option value="" <%if rsArticle("TitleFontColor")="" then response.write " selected"%>>颜色</option>
                <OPTION value="">默认</OPTION>
                <OPTION value="#000000" style="background-color:#000000" <%if rsArticle("TitleFontColor")="#000000" then response.write " selected"%>></OPTION>
                <OPTION value="#FFFFFF" style="background-color:#FFFFFF" <%if rsArticle("TitleFontColor")="#FFFFFF" then response.write " selected"%>></OPTION>
                <OPTION value="#008000" style="background-color:#008000" <%if rsArticle("TitleFontColor")="#008000" then response.write " selected"%>></OPTION>
                <OPTION value="#800000" style="background-color:#800000" <%if rsArticle("TitleFontColor")="#800000" then response.write " selected"%>></OPTION>
                <OPTION value="#808000" style="background-color:#808000" <%if rsArticle("TitleFontColor")="#808000" then response.write " selected"%>></OPTION>
                <OPTION value="#000080" style="background-color:#000080" <%if rsArticle("TitleFontColor")="#000080" then response.write " selected"%>></OPTION>
                <OPTION value="#800080" style="background-color:#800080" <%if rsArticle("TitleFontColor")="#800080" then response.write " selected"%>></OPTION>
                <OPTION value="#808080" style="background-color:#808080" <%if rsArticle("TitleFontColor")="#808080" then response.write " selected"%>></OPTION>
                <OPTION value="#FFFF00" style="background-color:#FFFF00" <%if rsArticle("TitleFontColor")="#FFFF00" then response.write " selected"%>></OPTION>
                <OPTION value="#00FF00" style="background-color:#00FF00" <%if rsArticle("TitleFontColor")="#00FF00" then response.write " selected"%>></OPTION>
                <OPTION value="#00FFFF" style="background-color:#00FFFF" <%if rsArticle("TitleFontColor")="#00FFFF" then response.write " selected"%>></OPTION>
                <OPTION value="#FF00FF" style="background-color:#FF00FF" <%if rsArticle("TitleFontColor")="#FF00FF" then response.write " selected"%>></OPTION>
                <OPTION value="#FF0000" style="background-color:#FF0000" <%if rsArticle("TitleFontColor")="#FF0000" then response.write " selected"%>></OPTION>
                <OPTION value="#0000FF" style="background-color:#0000FF" <%if rsArticle("TitleFontColor")="#0000FF" then response.write " selected"%>></OPTION>
                <OPTION value="#008080" style="background-color:#008080" <%if rsArticle("TitleFontColor")="#008080" then response.write " selected"%>></OPTION>
              </select>
              <select name="TitleFontType" id="TitleFontType">
                <option value="0" <%if rsArticle("TitleFontType")="0" then response.write " selected"%>>字形</option>
                <option value="1" <%if rsArticle("TitleFontType")="1" then response.write " selected"%>>粗体</option>
                <option value="2" <%if rsArticle("TitleFontType")="2" then response.write " selected"%>>斜体</option>
                <option value="3" <%if rsArticle("TitleFontType")="3" then response.write " selected"%>>粗+斜</option>
                <option value="0" <%if rsArticle("TitleFontType")="4" then response.write " selected"%>>规则</option>
              </select>
            </td>
          </tr>
          <tr class="tdbg"> 
            <td width="100" align="right"><strong>关键字：</strong></td>
            <td><input name="Key" type="text"
           id="Key" value="<%=mid(rsArticle("Key"),2,len(rsArticle("Key"))-2)%>" size="50" maxlength="255"> 
              <font color="#FF0000">*</font><br>
              <font color="#0000FF">用来查找相关文章，可输入多个关键字，中间用<font color="#FF0000">“|”</font>分开。不能出现&quot;'*?()等字符。</font></td>
          </tr>
          <tr class="tdbg"> 
            <td width="100" align="right"><strong>作者：</strong></td>
            <td>姓名： 
              <input name="AuthorName" type="text"
           id="AuthorName" value="<%=AuthorName%>" size="20" maxlength="30"> 
              &nbsp;&nbsp;&nbsp;&nbsp;Email： 
              <input name="AuthorEmail" type="text" id="AuthorEmail" value="<%=AuthorEmail%>" size="40" maxlength="100"> 
            </td>
          </tr>
          <tr class="tdbg"> 
            <td width="100" align="right"><strong>原出处：</strong></td>
            <td>名称： 
              <input name="CopyFromName" type="text"
           id="CopyFromName" value="<%=CopyFromName%>" size="20" maxlength="50"> 
              &nbsp;&nbsp;&nbsp;&nbsp;地 址： 
              <input name="CopyFromUrl" type="text" id="CopyFromUrl" value="<%=CopyFromUrl%>" size="40" maxlength="200"></td>
          </tr>
          <tr class="tdbg"> 
            <td width="100" align="right" valign="middle"><p><strong>文章内容：</strong></p>
              <p align="left"><font color="#006600">&middot;　如果是从其它网站上复制内容，并且内容中包含有图片，本系统保存修改结果时将把非本站图片复制到本站服务器上，系统会因复制图片的大小而影响速度，请稍候（此功能需要服务器安装了IE5.5以上版本才有效）。<br>
                <br>
                &middot;　换行请按Shift+Enter</font><br>
                <font color="#006600">&middot;　另起一段请按Enter</font></p></td>
            <td><textarea name="Content" style="display:none"></textarea>
              <iframe ID="editor" src="editor.asp?Action=Modify&ArticleID=<%=ArticleID%>" frameborder=1 scrolling=no width="600" height="405"></iframe> 
            </td>
          </tr>
          <tr class="tdbg"> 
            <td width="100" align="right"><strong>内容分页方式：</strong></td>
            <td><select name="PaginationType" id="PaginationType">
                <option value="0" <%if rsArticle("PaginationType")=0 then response.write " selected"%>>不分页</option>
                <option value="1" <%if rsArticle("PaginationType")=1 then response.write " selected"%>>自动分页</option>
                <option value="2" <%if rsArticle("PaginationType")=2 then response.write " selected"%>>手动分页</option>
              </select> &nbsp;&nbsp;&nbsp;&nbsp;<strong><font color="#0000FF">注：</font></strong><font color="#0000FF">手动分页符标记为“</font><font color="#FF0000">[NextPage]</font><font color="#0000FF">”，注意大小写</font></td>
          </tr>
          <tr class="tdbg">
            <td align="right">&nbsp;</td>
            <td>自动分页时的每页大约字符数（包含HTML标记）：<strong> 
              <input name="MaxCharPerPage" type="text" id="MaxCharPerPage" value="<%=rsArticle("MaxCharPerPage")%>" size="8" maxlength="8">
              </strong></td>
          </tr>
          <tr class="tdbg"> 
            <td width="100" align="right"><strong>文章阅读等级：</strong></td>
            <td><select name="ReadLevel" id="select">
                <option value="9999" <%if rsArticle("ReadLevel")=9999 then response.write " selected"%>>游客</option>
                <option value="999" <%if rsArticle("ReadLevel")=999 then response.write " selected"%>>注册用户</option>
                <option value="99" <%if rsArticle("ReadLevel")=99 then response.write " selected"%>>收费用户</option>
                <option value="9" <%if rsArticle("ReadLevel")=9 then response.write " selected"%>>VIP用户</option>
                <option value="5" <%if rsArticle("ReadLevel")=5 then response.write " selected"%>>管理员</option>
              </select>
              &nbsp;&nbsp;&nbsp;<font color="#0000FF">只有具有相应权限的人才能阅读此文章。</font></td>
          </tr>
          <tr class="tdbg">
            <td width="100" align="right"><strong>文章阅读点数：</strong></td>
            <td><input name="ReadPoint" type="text" id="ReadPoint" value="<%=rsArticle("ReadPoint")%>" size="5" maxlength="3">
              &nbsp;&nbsp;&nbsp;&nbsp; <font color="#0000FF">如果大于0，则用户阅读此文章时将消耗相应点数。（对游客和管理员无效）</font></td>
          </tr>
          <tr class="tdbg"> 
            <td width="100" align="right"><strong>包含图片：</strong></td>
            <td><input name="IncludePic" type="checkbox" id="IncludePic" value="yes" <% if rsArticle("IncludePic")=true then response.Write("checked") end if%>>
              是<font color="#0000FF">（如果选中的话会在标题前面显示[图文]）</font></td>
          </tr>
          <tr class="tdbg"> 
            <td width="100" align="right"><strong>首页图片：</strong></td>
            <td><input name="DefaultPicUrl" type="text" id="DefaultPicUrl" value="<%=rsArticle("DefaultPicUrl")%>" size="56" maxlength="200">
              用于在首页的图片文章处显示 <br>
              直接从上传图片中选择： 
              <select name="DefaultPicList" id="DefaultPicList" onChange="DefaultPicUrl.value=this.value;">
                <option value=""<% if rsArticle("DefaultPicUrl")="" then response.write "selected" %>>不指定首页图片</option>
                <%
				if rsArticle("UploadFiles")<>"" then
					dim IsOtherUrl
					IsOtherUrl=True
					if instr(rsArticle("UploadFiles"),"|")>1 then
						dim arrUploadFiles,intTemp
						arrUploadFiles=split(rsArticle("UploadFiles"),"|")						
						for intTemp=0 to ubound(arrUploadFiles)
							if rsArticle("DefaultPicUrl")=arrUploadFiles(intTemp) then
								response.write "<option value='" & arrUploadFiles(intTemp) & "' selected>" & arrUploadFiles(intTemp) & "</option>"
								IsOtherUrl=False
							else
								response.write "<option value='" & arrUploadFiles(intTemp) & "'>" & arrUploadFiles(intTemp) & "</option>"
							end if
						next
					else
						if rsArticle("UploadFiles")=rsArticle("DefaultPicUrl") then
							response.write "<option value='" & rsArticle("UploadFiles") & "' selected>" & rsArticle("UploadFiles") & "</option>"
							IsOtherUrl=False
						else
							response.write "<option value='" & rsArticle("UploadFiles") & "'>" & rsArticle("UploadFiles") & "</option>"		
						end if
					end If
					if IsOtherUrl=True then
						response.write "<option value='" & rsArticle("DefaultPicUrl") & "' selected>" & rsArticle("DefaultPicUrl") & "</option>"
					end if
				end if
				 %>
              </select> <input name="UploadFiles" type="hidden" id="UploadFiles" value="<%=rsArticle("UploadFiles")%>"> 
            </td>
          </tr>
<%if AdminPurview=1 or AdminPurview_Article<=2 then%>
          <tr class="tdbg"> 
            <td width="100" align="right"><strong>文章性质：</strong></td>
            <td><input name="OnTop" type="checkbox" id="OnTop" value="yes" <% if rsArticle("OnTop")=true then response.Write("checked") end if%>>
              固顶文章&nbsp;&nbsp;&nbsp;&nbsp; <input name="Hot" type="checkbox" id="Hot" value="yes" onclick="javascript:document.myform.Hits.value=<%=HitsOfHot%>" disabled>
              热点文章&nbsp;&nbsp;&nbsp;&nbsp; <input name="Elite" type=checkbox id="Elite" value="yes" <% if rsArticle("Elite")=true then response.Write("checked") end if%>>
              推荐文章&nbsp;&nbsp;&nbsp;&nbsp;文章评分等级： 
              <select name="Stars" id="Stars">
                <option value="5" <%if rsArticle("Stars")=5 then response.write " selected"%>>★★★★★</option>
                <option value="4" <%if rsArticle("Stars")=4 then response.write " selected"%>>★★★★</option>
                <option value="3" <%if rsArticle("Stars")=3 then response.write " selected"%>>★★★</option>
                <option value="2" <%if rsArticle("Stars")=2 then response.write " selected"%>>★★</option>
                <option value="1" <%if rsArticle("Stars")=1 then response.write " selected"%>>★</option>
                <option value="0" <%if rsArticle("Stars")=0 then response.write " selected"%>>无</option>
              </select>
            </td>
          </tr>
          <tr class="tdbg"> 
            <td width="100" align="right"><strong>点击数：</strong></td>
            <td><input name="Hits" type="text" id="Hits" value="<%=rsArticle("Hits")%>" size="10" maxlength="10"> 
              &nbsp;&nbsp;<font color="#0000FF">这功能是提供给管理员作弊用的。不过尽量不要用呀！^_^</font></td>
          </tr>
          <tr class="tdbg"> 
            <td width="100" align="right"><strong>录入时间：</strong></td>
            <td><input name="UpdateTime" type="text" id="UpdateTime" value="<%=rsArticle("UpdateTime")%>" maxlength="50">
              时间格式为“年-月-日 时:分:秒”，如：<font color="#0000FF">2003-5-12 12:32:47</font> 
            </td>
          </tr>
<%end if%>
          <tr class="tdbg"> 
            <td width="100" align="right"><strong>配色模板：</strong></td>
            <td><%call Admin_ShowSkin_Option(rsArticle("SkinID"))%>&nbsp;相关模板中包含CSS、颜色、图片等信息</td>
          </tr>
          <tr class="tdbg"> 
            <td width="100" align="right"><strong>版面设计模板：</strong></td>
            <td><%call Admin_ShowLayout_Option(3,rsArticle("LayoutID"))%>&nbsp;相关模板中包含了版面设计的版式等信息</td>
          </tr>
<%if AdminPurview=1 or AdminPurview_Article<=2 then%>
          <tr class="tdbg"> 
            <td width="100" align="right"><strong>已通过审核：</strong></td>
            <td><input name="Passed" type="checkbox" id="Passed" value="yes" <% if rsArticle("Passed")=true then response.Write("checked") end if%>>
              是<font color="#0000FF">（如果选中的话将直接发布，否则审核后才能发布。）</font></td>
          </tr>
<%end if%>
        </table>
      </td>
    </tr>
  </table>
  <div align="center"> 
    <p> 
      <input name="ArticleID" type="hidden" id="ArticleID" value="<%=rsArticle("ArticleID")%>">
      <input
  name="Save" type="submit"  id="Save" value="保存修改结果" style="cursor:hand;">
      &nbsp; 
      <input name="Cancel" type="button" id="Cancel" value=" 取 消 " onClick="window.location.href='Admin_ArticleManage.asp'" style="cursor:hand;">
    </p>
  </div>
</form>
</body>
</html>
<%
end if
rsArticle.close
set rsArticle=nothing
call CloseConn()
%>