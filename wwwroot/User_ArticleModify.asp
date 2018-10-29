<%@language=vbscript codepage=936 %>
<%
option explicit
response.buffer=true	
%>
<!--#include file="inc/conn.asp"-->
<!--#include file="inc/Conn_User.asp"-->
<!--#include file="inc/function.asp"-->
<!--#include file="inc/config.asp"-->
<!--#include file="inc/admin_code_article.asp"-->
<%
if CheckUserLogined()=False then
	response.Redirect "User_Login.asp"
end if

dim ArticleID,rs,sql,rsArticle,FoundErr,ErrMsg,PurviewChecked
dim Author,AuthorName,AuthorEmail,CopyFrom,CopyFromName,CopyFromUrl
ArticleID=trim(request("ArticleID"))
FoundErr=False
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
	if rsArticle("Editor")=request.Cookies("asp163")("UserName") then
	 	if rsArticle("Passed")=True then
			FoundErr=True
			ErrMsg=ErrMsg & "<br><li>文章已经被审核通过，你不能再进行修改！</li>"
		end if
	else
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>你不能修改其他人的文章！</li>"
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

	dim ClassID,tClass,ClassName,RootID,ParentID,Depth,ParentPath
	ClassID=rsArticle("ClassID")
	set tClass=conn.execute("select ClassName,RootID,ParentID,Depth,ParentPath From ArticleClass where ClassID=" & ClassID)
	ClassName=tClass(0)
	RootID=tClass(1)
	ParentID=tClass(2)
	Depth=tClass(3)
	ParentPath=tClass(4)
	set tClass=nothing
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>修改文章</title>
<link rel="stylesheet" type="text/css" href="Admin_style.css">
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
<form method="POST" name="myform" onSubmit="return CheckForm();" action="User_ArticleSave.asp?action=Modify">
  <table border="0" align="center" cellpadding="0" cellspacing="0" class="border">
    <tr class="title"> 
      <td height="22" align="center"><b>修 改 文 章</b></td>
    </tr>
    <tr align="center"> 
      <td class="tdbg"> <table width="100%" border="0" cellpadding="2" cellspacing="1">
          <tr class="tdbg"> 
            <td width="100" align="right"><strong>所属栏目：</strong></td>
            <td><%call Admin_ShowPath2(ParentPath,ClassName,Depth)
			response.write "<input type='hidden' name='ClassID' value='" & rsArticle("ClassID") & "'>"%></td>
          </tr>
          <tr class="tdbg"> 
            <td width="100" align="right"><strong>所属课程：</strong></td>
            <td><%
			if rsArticle("SpecialID")=0 then
				response.write "&nbsp;不属于任何课程"
			else
			  	dim rsSpecial
				set rsSpecial=conn.execute("select Specialname from Special where SpecialID=" & rsArticle("SpecialID"))
							  	dim rsSpecialAll
				set rsSpecialAll=conn.execute("select Specialname,SpecialID from Special")

				if rsSpecial.bof and rsSpecial.eof then
					response.write "找不到所属课程！可能所属课程已经被删除！"
				else
					Response.Write("<select name='SpecialID' id='SpecialID'>")
      
					if rsSpecialAll.bof and rsSpecialAll.eof then
						response.write "<option value=''>没有任何课程</option>"
					else
						
'						response.write "<option value='" & rsArticle("SpecialID") & "'>" & rsSpecial("SpecialName") & "</option>"

							rsSpecialAll.movefirst
							do while not rsSpecialAll.eof
								if rsSpecialAll("SpecialID")=rsArticle("SpecialID") then
									response.write "<option value='" & rsSpecialAll("SpecialID") & "'selected>" & rsSpecialAll("SpecialName") & "</option>"
								else
									response.write "<option value='" & rsSpecialAll("SpecialID") & "'>" & rsSpecialAll("SpecialName") & "</option>"
								end if
								rsSpecialAll.movenext
							loop
						end if

			 Response.Write(" </select>")
'					response.write rsSpecial(0)
			  	end if
				set rsSpecial=nothing
				set rsSpecialAll=nothing
			end if
			
'			response.write "<input type='hidden' name='SpecialID' value='" & rsArticle("SpecialID") & "'>"
			%>
            
            </td></tr>
			<tr class="tdbg"><td width="100" align="right"><strong>任课教师：</strong></td>
			<td><%call User_ArticleModify_TeacherList()
			%></td>
          </tr>
          <tr> 
            <td width="100" align="right"><strong>文章标题：</strong></td>
            <td><input name="Title" type="text"
           id="Title" value="<%=rsArticle("Title")%>" size="50" maxlength="255"> 
              <font color="#FF0000">*</font> <select name="TitleFontColor" id="TitleFontColor">
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
              </select> <select name="TitleFontType" id="TitleFontType">
                <option value="0" <%if rsArticle("TitleFontType")="0" then response.write " selected"%>>字形</option>
                <option value="1" <%if rsArticle("TitleFontType")="1" then response.write " selected"%>>粗体</option>
                <option value="2" <%if rsArticle("TitleFontType")="2" then response.write " selected"%>>斜体</option>
                <option value="3" <%if rsArticle("TitleFontType")="3" then response.write " selected"%>>粗+斜</option>
                <option value="0" <%if rsArticle("TitleFontType")="4" then response.write " selected"%>>规则</option>
              </select> </td>
          </tr>
          <tr> 
            <td width="100" align="right"><strong>关键字：</strong></td>
            <td><input name="Key" type="text"
           id="Key" value="<%=mid(rsArticle("Key"),2,len(rsArticle("Key"))-2)%>" size="50" maxlength="255"> 
              <font color="#FF0000">*</font><br> <font color="#0000FF">用来查找相关文章，可输入多个关键字，中间用<font color="#FF0000">“|”</font>分开。不能出现&quot;'*?()等字符。</font></td>
          </tr>
          <tr> 
            <td width="100" align="right"><strong>作者：</strong></td>
            <td>姓名： 
              <input name="AuthorName" type="text"
           id="AuthorName" value="<%=AuthorName%>" size="20" maxlength="30"> &nbsp;&nbsp;&nbsp;&nbsp;Email： 
              <input name="AuthorEmail" type="text" id="AuthorEmail" value="<%=AuthorEmail%>" size="40" maxlength="100"> 
            </td>
          </tr>
          <tr> 
            <td width="100" align="right"><strong>原出处：</strong></td>
            <td>名称： 
              <input name="CopyFromName" type="text"
           id="CopyFromName" value="<%=CopyFromName%>" size="20" maxlength="50"> 
              &nbsp;&nbsp;&nbsp;&nbsp;地 址： 
              <input name="CopyFromUrl" type="text" id="CopyFromUrl2" value="<%=CopyFromUrl%>" size="40" maxlength="200"></td>
          </tr>
          <tr> 
            <td width="100" align="right" valign="middle"><strong>文章内容：</strong></td>
            <td><textarea name="Content" style="display:none"></textarea> 
              <iframe ID="editor" src="editor.asp?UserType=User&Action=Modify&ArticleID=<%=ArticleID%>" frameborder=1 scrolling=no width="600" height="405"></iframe> 
            </td>
          </tr>
          <tr> 
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
          <tr> 
            <td width="100" align="right"><strong>包含图片：</strong></td>
            <td><input name="IncludePic" type="checkbox" id="IncludePic" value="yes" <% if rsArticle("IncludePic")=true then response.Write("checked") end if%>>
              是<font color="#0000FF">（如果选中的话会在标题前面显示[图文]）</font></td>
          </tr>
          <tr> 
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
          <%if PurviewChecked=True then%>
          <%end if%>
        </table></td>
    </tr>
  </table>
  <div align="center"> 
    <p> 
      <input name="SkinID" type="hidden" id="SkinID" value="<%=rsArticle("SkinID")%>">
      <input name="LayoutID" type="hidden" id="LayoutID" value="<%=rsArticle("LayoutID")%>">
      <input name="ArticleID" type="hidden" id="ArticleID" value="<%=rsArticle("ArticleID")%>">
      <input
  name="Save" type="submit"  id="Save" value="保存修改结果" style="cursor:hand;">
      &nbsp; 
      <input name="Cancel" type="button" id="Cancel" value=" 取 消 " onClick="window.location.href='User_ArticleManage.asp'" style="cursor:hand;">
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