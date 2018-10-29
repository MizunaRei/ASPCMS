<%@language=vbscript codepage=936 %>
<%
option explicit
response.buffer=true	
Const PurviewLevel=2    '操作权限
Const CheckChannelID=0    '所属频道，0为不检测所属频道
Const PurviewLevel_Others="UpFile"
%>
<!--#include file="inc/conn.asp"-->
<!--#include file="inc/config.asp"-->
<!--#include file="Admin_ChkPurview.asp"-->
<!--#include file="inc/function.asp"-->
<%
Const MaxPerPage=20
dim strFileName
dim Action
dim totalPut,CurrentPage,TotalPages
dim UploadDir,TruePath,fso,theFolder,theFile,thisfile,FileCount,TotalSize,TotalSize_Page
dim strFileType
dim sql,rs,strFiles,i
dim strDirName
Action=trim(Request("Action"))

if request("page")<>"" then
    currentPage=cint(request("page"))
else
	currentPage=1
end if

set rs=server.CreateObject("adodb.recordset")
select case trim(request("UploadDir"))
case "UploadFiles"
	UploadDir="UploadFiles"
	strDirName="文章频道的上传文件"
	sql="select UploadFiles from Article"
	rs.open sql,conn,1,1
	do while not rs.eof
		if rs(0)<>"" then
			strFiles=strFiles & "|" & rs(0)
		end if
		rs.movenext
	loop
case "UploadThumbs"
	UploadDir="UploadThumbs"
	strDirName="图片频道的缩略图"
	sql="select PhotoUrl_Thumb,PhotoUrl,PhotoUrl2,PhotoUrl3,PhotoUrl4 from Photo"
	rs.open sql,conn,1,1
	do while not rs.eof
		if rs(0)<>"" then
			strFiles=strFiles & "|" & rs(0)
		end if
		if rs(1)<>"" then
			strFiles=strFiles & "|" & rs(0)
		end if
		if rs(2)<>"" then
			strFiles=strFiles & "|" & rs(0)
		end if
		if rs(3)<>"" then
			strFiles=strFiles & "|" & rs(0)
		end if
		if rs(4)<>"" then
			strFiles=strFiles & "|" & rs(0)
		end if
		rs.movenext
	loop
case "UploadPhotos"
	UploadDir="UploadPhotos"
	strDirName="图片频道的上传图片"
	sql="select PhotoUrl_Thumb,PhotoUrl,PhotoUrl2,PhotoUrl3,PhotoUrl4 from Photo"
	rs.open sql,conn,1,1
	do while not rs.eof
		if rs(0)<>"" then
			strFiles=strFiles & "|" & rs(0)
		end if
		if rs(1)<>"" then
			strFiles=strFiles & "|" & rs(0)
		end if
		if rs(2)<>"" then
			strFiles=strFiles & "|" & rs(0)
		end if
		if rs(3)<>"" then
			strFiles=strFiles & "|" & rs(0)
		end if
		if rs(4)<>"" then
			strFiles=strFiles & "|" & rs(0)
		end if
		rs.movenext
	loop
case "UploadSoftPic"
	UploadDir="UploadSoftPic"
	strDirName="下载频道的软件图片"
	sql="select SoftPicUrl from Soft"
	rs.open sql,conn,1,1
	do while not rs.eof
		if rs(0)<>"" then
			strFiles=strFiles & "|" & rs(0)
		end if
		rs.movenext
	loop
case "UploadSoft"
	UploadDir="UploadSoft"
	strDirName="下载频道的上传软件"
	sql="select DownloadUrl1,DownloadUrl2,DownloadUrl3,DownloadUrl4 from Soft"
	rs.open sql,conn,1,1
	do while not rs.eof
		if rs(0)<>"" then
			strFiles=strFiles & "|" & rs(0)
		end if
		if rs(1)<>"" then
			strFiles=strFiles & "|" & rs(0)
		end if
		if rs(2)<>"" then
			strFiles=strFiles & "|" & rs(0)
		end if
		if rs(3)<>"" then
			strFiles=strFiles & "|" & rs(0)
		end if
		rs.movenext
	loop
case "UploadAdPic"
	UploadDir="UploadAdPic"
	strDirName="网站广告的上传图片"
	sql="select ImgUrl from Advertisement"
	rs.open sql,conn,1,1
	do while not rs.eof
		if rs(0)<>"" then
			strFiles=strFiles & "|" & rs(0)
		end if
	rs.movenext
	loop
case else
	UploadDir="UploadFiles"
	strDirName="文章频道的上传文件"
	sql="select UploadFiles from Article"
	rs.open sql,conn,1,1
	do while not rs.eof
		if rs(0)<>"" then
			strFiles=strFiles & "|" & rs(0)
		end if
		rs.movenext
	loop
end select
rs.close
set rs=nothing

strFileName="Admin_UploadFile.asp?UploadDir=" & UploadDir
if right(UploadDir,1)<>"/" then
	UploadDir=UploadDir & "/"
end if
TruePath=Server.MapPath(UploadDir)
%>

<html>
<head>
<title>上传文件管理</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="Admin_Style.css" rel="stylesheet" type="text/css">
</head>
<SCRIPT language=javascript>
function unselectall()
{
    if(document.myform.chkAll.checked){
	document.myform.chkAll.checked = document.myform.chkAll.checked&0;
    } 	
}

function CheckAll(form)
{
  for (var i=0;i<form.elements.length;i++)
    {
    var e = form.elements[i];
    if (e.Name != "chkAll")
       e.checked = form.chkAll.checked;
    }
}

</script>
<body leftmargin="2" topmargin="0" marginwidth="0" marginheight="0">
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" Class="border">
  <tr class="topbg"> 
    <td height="22" colspan=2 align=center><b>上 传 文 件 管 理</b></td>
  </tr>
  <tr class="tdbg"> 
    <td width="70" height="30"><strong>管理导航：</strong></td>
    <td height="30"><a href="Admin_UploadFile.asp?UploadDir=UploadFiles">文章频道的上传文件</a> 
      |<a href="Admin_UploadFile.asp?UploadDir=UploadAdPic"> 网站广告的上传图片</a> | <a href="Admin_UploadFile.asp?Action=Clear">清除无用文件</a> 
      | </td>
  </tr>
</table>
<%
If not IsObjInstalled("Scripting.FileSystemObject") Then
	Response.Write "<b><font color=red>你的服务器不支持 FSO(Scripting.FileSystemObject)! 不能使用本功能</font></b>"
Else
	set fso=CreateObject("Scripting.FileSystemObject")
	if Action="Del" then
		call DelFiles()
	elseif Action="DelAll" then
		call DelAll()
	elseif Action="Clear" or Action="DoClear" then
		call ClearFile()
	else
		call main()
	end if
end if

sub main()
	if fso.FolderExists(TruePath)=False then
		response.write "找不到文件夹！可能是配置有误！"
		exit sub
	end if
	
	response.write "<br>您现在的位置：<a href='Admin_UploadFile.asp'>上传文件管理</a>&nbsp;&gt;&gt;&nbsp;<a href='Admin_UploadFile.asp?UploadDir=" & UploadDir & "'><font color=red>" & strDirName & "</font></a>" 
	FileCount=0
	TotalSize=0
	Set theFolder=fso.GetFolder(TruePath)
	For Each theFile In theFolder.Files
		FileCount=FileCount+1
		TotalSize=TotalSize+theFile.Size
	next
	totalPut=FileCount
	if currentpage<1 then
		currentpage=1
	end if
	if (currentpage-1)*MaxPerPage>totalput then
		if (totalPut mod MaxPerPage)=0 then
			currentpage= totalPut \ MaxPerPage
		else
			currentpage= totalPut \ MaxPerPage + 1
		end if
			end if
	if currentPage=1 then
		showContent     	
		showpage2 strFileName,totalput,MaxPerPage
		response.write "<br><div align='center'>本页共显示 <b>" & FileCount & "</b> 个文件，占用 <b>" & TotalSize_Page\1024 & "</b> K</div>"
	else
		if (currentPage-1)*MaxPerPage<totalPut then
			showContent     	
			showpage2 strFileName,totalput,MaxPerPage
			response.write "<br><div align='center'>本页共显示 <b>" & FileCount & "</b> 个文件，占用 <b>" & TotalSize_Page\1024 & "</b> K</div>"
		else
			currentPage=1
			showContent     	
			showpage2 strFileName,totalput,MaxPerPage
			response.write "<br><div align='center'>本页共显示 <b>" & FileCount & "</b> 个文件，占用 <b>" & TotalSize_Page\1024 & "</b> K</div>"
		end if
	end if
end sub

sub showContent()
   	dim c
	FileCount=0
	TotalSize_Page=0
%>
<br>
<table width='100%' border="0" cellpadding="0" cellspacing="0">
  <tr>
  <form name="myform" method="Post" action="Admin_UploadFile.asp" onsubmit="return confirm('确定要删除选中的文件吗？');">
     <td>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="3" class="border">
  <tr class="tdbg">
    <%

For Each theFile In theFolder.Files
	c=c+1
	if FileCount>=MaxPerPage then
		exit for
	elseif c>MaxPerPage*(CurrentPage-1) then
%>
    <td>
      <table width="100%" height="100%" border="0" cellpadding="0" cellspacing="2">
        <tr>
          <td colspan="2" align="center">
            <%
		  strFileType=lcase(mid(theFile.Name,instrrev(theFile.Name,".")+1))
		  response.write "<a href='" & UploadDir & theFile.Name & "'>"
		  select case strFileType
		  case "jpg","gif","bmp","png"
			  if instr(strFiles,theFile.Name)>0 then
			  	response.write "<img src='" & UploadDir & theFile.Name & "' width='140' height='100' border='0'></a>"
			  else
			  	response.write "<img src='" & UploadDir & theFile.Name & "' width='140' height='100' border='2' Title='无用的上传文件'></a>"
			  end if
		  case "swf"
			  if instr(strFiles,theFile.Name)>0 then
			  	response.write "<img src='images/filetype_flash.gif' width='140' height='100' border='0'></a>"
			  else
				response.write "<img src='images/filetype_flash.gif' width='140' height='100' border='2' Title='无用的上传文件'></a>"
			  end if
		  case "wmv","avi","asf","mpg"
			  if instr(strFiles,theFile.Name)>0 then
			  	response.write "<img src='images/filetype_media.gif' width='140' height='100' border='0'></a>"
			  else	
				response.write "<img src='images/filetype_media.gif' width='140' height='100' border='2' Title='无用的上传文件'></a>"
			  end if
		  case "rm","ra","ram"
			  if instr(strFiles,theFile.Name)>0 then
		  		response.write "<img src='images/filetype_rm.gif' width='140' height='100' border='0'></a>"
		  	  else		
				response.write "<img src='images/filetype_rm.gif' width='140' height='100' border='2' Title='无用的上传文件'></a>"
			  end if
		  case "rar"
		    response.write "<img src='images/filetype_rar.gif' width='140' height='100' border='0'></a>"
		  case "zip"
		    response.write "<img src='images/filetype_zip.gif' width='140' height='100' border='0'></a>"
		  case "exe"
		    response.write "<img src='images/filetype_exe.gif' width='140' height='100' border='0'></a>"
		  case else
			  if instr(strFiles,theFile.Name)>0 then
		  		response.write "<img src='images/filetype_other.gif' width='140' height='100' border='0'></a>"
		  	  else		
				response.write "<img src='images/filetype_other.gif' width='140' height='100' border='2' Title='无用的上传文件'></a>"
			  end if
		  end select
		  %>
          </td>
        </tr>
        <tr>
          <td align="right">文 件 名：</td>
          <td><%
		  if instr(strFiles,theFile.Name)>0 then
		  	response.write "<a href='" & UploadDir & theFile.Name & "' target='_blank'>" & theFile.Name & "</a>"
		  else
		  	response.write "<a href='" & UploadDir & theFile.Name & "' target='_blank' title='无用的上传文件'><font color=red>" & theFile.Name & "</font></a>"
		  end if%>
		  </td>
        </tr>
        <tr>
          <td align="right">文件大小：</td>
          <td><%=round(theFile.size/1024) & " K"%></td>
        </tr>
        <tr>
          <td align="right">文件类型：</td>
          <td><%=theFile.type%></td>
        </tr>
        <tr>
          <td align="right">修改时间：</td>
          <td><%=theFile.DateLastModified%></td>
        </tr>
        <tr>
          <td align="right">操作选项：</td>
          <td><input name="FileName" type="checkbox" id="FileName" value="<%=theFile.Name%>" onclick="unselectall()" <%if instr(strFiles,theFile.Name)<=0 then response.write "checked"%>>
            选中&nbsp;&nbsp;&nbsp;&nbsp;<a href="Admin_UploadFile.asp?Action=Del&FileName=<%=theFile.Name%>&UploadDir=<%=left(UploadDir,len(UploadDir)-1)%>" onclick="return confirm('你真的要删除此文件吗!')">删除</a></td>
        </tr>
      </table>
    </td>
    <%
		FileCount=FileCount+1
		if FileCount mod 4=0 then response.write "</td><tr class='tdbg'>"
		TotalSize_Page=TotalSize_Page+theFile.Size
	end if
Next
%>
  </tr>
</table>
<table width="100%" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td width="200" height="30"><input name="chkAll" type="checkbox" id="chkAll" onclick=CheckAll(this.form) value="checkbox">
      选中本页显示的所有文件</td>
    <td><input name="Action" type="hidden" id="Action" value="Del">
      <input name="UploadDir" type="hidden" id="UploadDir" value="<%=left(UploadDir,len(UploadDir)-1)%>">
              <input type="submit" name="Submit" value="删除选中的文件">
              &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
              <input type="submit" name="Submit2" value="删除所有文件" onClick="document.myform.Action.value='DelAll';">
              </td>
  </tr>
</table>
</td></form></tr></table>
<%
end sub

sub ClearFile()
%>
<br>
<table width="100%" border="0" cellspacing="1" cellpadding="2" class="border">
  <tr class="title">
    <td height="22" align="center"><strong>清理无用的上传文件</strong></td>
  </tr>
  <tr class="tdbg">
    <td height="150">
<%
if Action="Clear" then
%>
<form name="form1" method="post" action="Admin_UploadFile.asp" onSubmit="javascript:if(document.form1.UploadFiles.checked==false&&document.form1.UploadThumbs.checked==false&&document.form1.UploadPhotos.checked==false&&document.form1.UploadSoftPic.checked==false&&document.form1.UploadSoft.checked==false&&document.form1.UploadAdPic.checked==false){alert('请先至少选择一个要清空的目录！');return false;}">
&nbsp;&nbsp;&nbsp;&nbsp;在添加文章时，经常会出现上传了图片后但却最终没有发布这篇文章的情况，时间一久，就会产生大量无用垃圾文件。所以需要定期使用本功能进行清理。      
<p>&nbsp;&nbsp;&nbsp;&nbsp;如果上传文件很多，或者文章数量较多，执行本操作需要耗费相当长的时间，请在访问量少时执行本操作。</p>
<p>&nbsp;&nbsp;&nbsp;&nbsp;请选择需要清理的上传目录：</p>
        <table width="150" border="0" align="center" cellpadding="0" cellspacing="2">
          <tr> 
            <td><input name="UploadFiles" type="checkbox" id="UploadFiles" value="Yes">
              文章频道的上传文件</td>
          </tr>
  
          <tr> 
            <td><input name="UploadAdPic" type="checkbox" id="UploadAdPic" value="Yes">
              网站广告的上传图片</td>
          </tr>
        </table>
<p align="center"><input name="Action" type="hidden" id="Action" value="DoClear">
      <input type="submit" name="Submit3" value=" 开始清理 ">
</p>
</form>
<%
else
	call DoClear()
end if
%>
    </td>
  </tr>
</table>
<%
end sub
%>
</body> 
</html>
<%
sub showpage2(sfilename,totalnumber,maxperpage)
	dim n, i,strTemp
	if totalnumber mod maxperpage=0 then
    	n= totalnumber \ maxperpage
  	else
    	n= totalnumber \ maxperpage+1
  	end if
  	strTemp= "<table align='center'><form name='showpages' method='Post' action='" & sfilename & "'><tr><td>"
	strTemp=strTemp & "共 <b>" & totalnumber & "</b> 个文件，占用 <b>" & TotalSize\1024 & "</b> K&nbsp;&nbsp;&nbsp;"
	sfilename=JoinChar(sfilename)
  	if CurrentPage<2 then
    		strTemp=strTemp & "首页 上一页&nbsp;"
  	else
    		strTemp=strTemp & "<a href='" & sfilename & "page=1'>首页</a>&nbsp;"
    		strTemp=strTemp & "<a href='" & sfilename & "page=" & (CurrentPage-1) & "'>上一页</a>&nbsp;"
  	end if

  	if n-currentpage<1 then
    		strTemp=strTemp & "下一页 尾页"
  	else
    		strTemp=strTemp & "<a href='" & sfilename & "page=" & (CurrentPage+1) & "'>下一页</a>&nbsp;"
    		strTemp=strTemp & "<a href='" & sfilename & "page=" & n & "'>尾页</a>"
  	end if
   	strTemp=strTemp & "&nbsp;页次：<strong><font color=red>" & CurrentPage & "</font>/" & n & "</strong>页 "
    strTemp=strTemp & "&nbsp;<b>" & maxperpage & "</b>" & "个文件/页"
	strTemp=strTemp & "&nbsp;转到：<select name='page' size='1' onchange='javascript:submit()'>"   
    for i = 1 to n   
   		strTemp=strTemp & "<option value='" & i & "'"
		if cint(CurrentPage)=cint(i) then strTemp=strTemp & " selected "
		strTemp=strTemp & ">第" & i & "页</option>"   
	next
	strTemp=strTemp & "</select>"
	strTemp=strTemp & "</td></tr></form></table>"
	response.write strTemp
end sub

sub DelFiles()
	dim whichfile,arrFileName,i
	whichfile=trim(Request("FileName"))
	if whichfile="" then exit sub
	if instr(whichfile,",")>0 then
		arrFileName=split(whichfile,",")
		for i=0 to ubound(arrFileName)
			if left(trim(arrFileName(i)),3)<>"../" and left(trim(arrFileName(i)),1)<>"/" then
				whichfile=server.MapPath(UploadDir & trim(arrFileName(i)))
				set thisfile=fso.GetFile(whichfile)
				thisfile.Delete True
			end if
		next
	else
		if left(whichfile,3)<>"../" and left(whichfile,1)<>"/" then
			Set thisfile = fso.GetFile(server.MapPath(UploadDir & whichfile))
			thisfile.Delete True
		end if
	end if
	call main()
end sub

sub DelAll()
	Set theFolder=fso.GetFolder(TruePath)
	For Each theFile In theFolder.Files
		theFile.Delete True
	next
	call main()
end sub

sub DoClear()
	set rs=server.CreateObject("adodb.recordset")
	if trim(request("UploadFiles"))="Yes" then
		strFiles=""
		sql="select UploadFiles from Article"
		rs.open sql,conn,1,1
		do while not rs.eof
			if rs(0)<>"" then
				strFiles=strFiles & "|" & rs(0)
			end if
			rs.movenext
		loop
		rs.close
		call DelFile_Useless("UploadFiles","文章频道的上传文件")
	end if
	
	if trim(request("UploadThumbs"))="Yes" then
		strFiles=""
		sql="select PhotoUrl_Thumb,PhotoUrl,PhotoUrl2,PhotoUrl3,PhotoUrl4 from Photo"
		rs.open sql,conn,1,1
		do while not rs.eof
			if rs(0)<>"" then
				strFiles=strFiles & "|" & rs(0)
			end if
			if rs(1)<>"" then
				strFiles=strFiles & "|" & rs(0)
			end if
			if rs(2)<>"" then
				strFiles=strFiles & "|" & rs(0)
			end if
			if rs(3)<>"" then
				strFiles=strFiles & "|" & rs(0)
			end if
			if rs(4)<>"" then
				strFiles=strFiles & "|" & rs(0)
			end if
			rs.movenext
		loop
		rs.close
		call DelFile_Useless("UploadThumbs","图片频道的缩略图")
	end if
	
	if trim(request("UploadPhotos"))="Yes" then
		strFiles=""
		sql="select PhotoUrl_Thumb,PhotoUrl,PhotoUrl2,PhotoUrl3,PhotoUrl4 from Photo"
		rs.open sql,conn,1,1
		do while not rs.eof
			if rs(0)<>"" then
				strFiles=strFiles & "|" & rs(0)
			end if
			if rs(1)<>"" then
				strFiles=strFiles & "|" & rs(0)
			end if
			if rs(2)<>"" then
				strFiles=strFiles & "|" & rs(0)
			end if
			if rs(3)<>"" then
				strFiles=strFiles & "|" & rs(0)
			end if
			if rs(4)<>"" then
				strFiles=strFiles & "|" & rs(0)
			end if
			rs.movenext
		loop
		rs.close
		call DelFile_Useless("UploadPhotos","图片频道的上传图片")
	end if

	if trim(request("UploadSoftPic"))="Yes" then
		strFiles=""
		sql="select SoftPicUrl from Soft"
		rs.open sql,conn,1,1
		do while not rs.eof
			if rs(0)<>"" then
				strFiles=strFiles & "|" & rs(0)
			end if
			rs.movenext
		loop
		rs.close
		call DelFile_Useless("UploadSoftPic","下载频道的软件图片")
	end if
	
	if trim(request("UploadSoft"))="Yes" then
		strFiles=""
		sql="select DownloadUrl1,DownloadUrl2,DownloadUrl3,DownloadUrl4 from Soft"
		rs.open sql,conn,1,1
		do while not rs.eof
			if rs(0)<>"" then
				strFiles=strFiles & "|" & rs(0)
			end if
			if rs(1)<>"" then
				strFiles=strFiles & "|" & rs(0)
			end if
			if rs(2)<>"" then
				strFiles=strFiles & "|" & rs(0)
			end if
			if rs(3)<>"" then
				strFiles=strFiles & "|" & rs(0)
			end if
			rs.movenext
		loop
		rs.close
		call DelFile_Useless("UploadSoft","下载频道的上传软件")
	end if

	if trim(request("UploadAdPic"))="Yes" then
		strFiles=""
		sql="select ImgUrl from Advertisement"
		rs.open sql,conn,1,1
		do while not rs.eof
			if rs(0)<>"" then
				strFiles=strFiles & "|" & rs(0)
			end if
			rs.movenext
		loop
		rs.close
		call DelFile_Useless("UploadAdPic","网站广告的上传图片")
	end if

	set rs=nothing
end sub

sub DelFile_Useless(strDir,strDirName)
	dim i
	i=0
	Set theFolder=fso.GetFolder(server.MapPath(strDir))
	For Each theFile In theFolder.Files
		if instr(strFiles,theFile.Name)<=0 then
			theFile.Delete True
			i=i+1
		end if
	next
	response.write "操作执行成功！在 <font color=blue>" & strDirName & "</font> 目录中共删除了 <font color=red><b>" & i & "</b></font> 个无用的文件。<br><br>"
end sub
%>