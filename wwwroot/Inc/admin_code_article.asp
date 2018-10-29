<%

sub Admin_ShowRootClass()
	dim sqlRoot,rsRoot
	sqlRoot="select ClassID,ClassName,RootID,Child From ArticleClass where ParentID=0 and LinkUrl='' order by RootID"
	Set rsRoot= Server.CreateObject("ADODB.Recordset")
	rsRoot.open sqlRoot,conn,1,1
	if rsRoot.bof and rsRoot.eof then 
		response.Write("还没有任何栏目，请首先添加栏目。")
	else
		response.write "|&nbsp;"
		do while not rsRoot.eof
			if rsRoot(2)=RootID then
				response.Write("<a href='" & FileName & "?ClassID=" & rsRoot(0) & "'><font color=red>" & rsRoot(1) & "</font></a> | ")
				tID=rsRoot(0)
				tChild=rsRoot(3)
			else
				response.Write("<a href='" & FileName & "?ClassID=" & rsRoot(0) & "'>" & rsRoot(1) & "</a> | ")
			end if
			rsRoot.movenext
		loop
	end if
	rsRoot.close
	set rsRoot=nothing
end sub

sub Admin_ShowClass_Option(ShowType,CurrentID)
	if ShowType=0 then
	    response.write "<option value='0'"
		if CurrentID=0 then response.write " selected"
		response.write ">无（作为一级栏目）</option>"
	end if
	dim rsClass,sqlClass,strTemp,tmpDepth,i
	dim arrShowLine(20)
	for i=0 to ubound(arrShowLine)
		arrShowLine(i)=False
	next
	sqlClass="Select * From ArticleClass order by RootID,OrderID"
	set rsClass=server.CreateObject("adodb.recordset")
	rsClass.open sqlClass,conn,1,1
	if rsClass.bof and rsClass.bof then
		response.write "<option value=''>请先添加栏目</option>"
	else
		dim UserLevel
		UserLevel=request.Cookies("asp163")("UserLevel")
		if UserLevel="" then
			UserLevel=5000
		else
			UserLevel=Cint(UserLevel)
		end if
		do while not rsClass.eof
			tmpDepth=rsClass("Depth")
			if rsClass("NextID")>0 then
				arrShowLine(tmpDepth)=True
			else
				arrShowLine(tmpDepth)=False
			end if
			if ShowType=1 then
				if rsClass("LinkUrl")<>"" then
					strTemp="<option value=''"
				else
					strTemp="<option value='" & rsClass("ClassID") & "'"
				end if
				if AdminPurview=2 and AdminPurview_Article=3 then
					if CheckClassMaster(rsClass("ClassChecker"),AdminName)=True then
						strTemp=strTemp & "style='background-color:#ff0000'"
					end if
				end if
			elseif ShowType=2 then
				if rsClass("LinkUrl")<>"" then
					strTemp="<option value=''"
				else
					strTemp="<option value='" & rsClass("ClassID") & "'"
				end if
				if AdminPurview=2 and AdminPurview_Article=3 then
					if CheckClassMaster(rsClass("ClassMaster"),AdminName)=True then
						strTemp=strTemp & "style='background-color:#ff0000'"
					end if
				end if
			elseif ShowType=3 then
				if rsClass("Child")>0 then
					strTemp="<option value=''"
				elseif rsClass("LinkUrl")<>"" then
					strTemp="<option value='0'"
				else
					strTemp="<option value='" & rsClass("ClassID") & "'"
				end if
				if AdminPurview=2 and AdminPurview_Article=3 then
					if CheckClassMaster(rsClass("ClassInputer"),AdminName)=True then
						strTemp=strTemp & "style='background-color:#ff0000'"
					end if
				end if
			elseif ShowType=4 then
				if rsClass("Child")>0 then
					strTemp="<option value=''"
				elseif rsClass("LinkUrl")<>"" then
					strTemp="<option value='0'"
				elseif rsClass("AddPurview")<UserLevel then
					strTemp="<option value='-1'"
				else
					strTemp="<option value='" & rsClass("ClassID") & "'"
				end if
			else
				strTemp="<option value='" & rsClass("ClassID") & "'"
			end if
			if CurrentID>0 and rsClass("ClassID")=CurrentID then
				 strTemp=strTemp & " selected"
				 SkinID=rsClass("SkinID")
				 LayoutID=rsClass("LayoutID")
				 BrowsePurview=rsClass("BrowsePurview")
				 AddPurview=rsClass("AddPurview")
			end if
			strTemp=strTemp & ">"
			
			if tmpDepth>0 then
				for i=1 to tmpDepth
					strTemp=strTemp & "&nbsp;&nbsp;"
					if i=tmpDepth then
						if rsClass("NextID")>0 then
							strTemp=strTemp & "├&nbsp;"
						else
							strTemp=strTemp & "└&nbsp;"
						end if
					else
						if arrShowLine(i)=True then
							strTemp=strTemp & "│"
						else
							strTemp=strTemp & "&nbsp;"
						end if
					end if
				next
			end if
			strTemp=strTemp & rsClass("ClassName")
			if rsClass("LinkUrl")<>"" then
				strTemp=strTemp & "(外)"
			end if
			if ShowType=4 and rsClass("AddPurview")<UserLevel then
				strTemp=strTemp & " *"
			end if
			strTemp=strTemp & "</option>"
			response.write strTemp
			rsClass.movenext
		loop
	end if
	rsClass.close
	set rsClass=nothing
end sub


sub Admin_ShowPath(RootName)
	response.write "您现在的位置：&nbsp;<a href='" & FileName & "'>" & RootName & "</a>&nbsp;&gt;&gt;&nbsp;"
	if ClassID>0 then
		if ParentID>0 then
			dim sqlPath,rsPath
			sqlPath="select ClassID,ClassName From ArticleClass where ClassID in (" & ParentPath & ") order by Depth"
			set rsPath=server.createobject("adodb.recordset")
			rsPath.open sqlPath,conn,1,1
			do while not rsPath.eof
				response.Write "<a href='" & FileName & "?ClassID=" & rsPath(0) & "'>" & rsPath(1) & "</a>&nbsp;&gt;&gt;&nbsp;"
				rsPath.movenext
			loop
			rsPath.close
			set rsPath=nothing
		end if
		response.write "<a href='" & FileName & "?ClassID=" & ClassID & "'>" & ClassName & "</a>&nbsp;&gt;&gt;&nbsp;"
	end if
	if ManageType="MyArticle" then
		response.write "<font color=red>" & AdminName & "</font> 的文章"
	else
		if keyword="" then
			response.write "所有文章"
		else
			select case strField
				case "Title"
					response.Write "标题中含有 <font color=red>"&keyword&"</font> 的文章"
				case "Content"
					response.Write "内容含有 <font color=red>"&keyword&"</font> 的文章"
				case "Author"
					response.Write "作者姓名中含有 <font color=red>"&keyword&"</font> 的文章"
				case "Editor"
					response.write "编辑姓名中含有 <font color=red>" & keyword & "</font> 的文章"
				case else
					response.Write "标题中含有 <font color=red>"&keyword&"</font> 的文章"
			end select
		end if
	end if
end sub


sub Admin_Add_ShowSpecial_Checkbox()
'数据库连接
	dim sqlSpecial,rsSpecial

		sqlSpecial = "select SpecialID,SpecialName from Special"

	set rsSpecial=server.CreateObject("adodb.recordset")
	rsSpecial.open sqlSpecial,conn,1,1
'已开启数据库连接
	rsSpecial.movefirst

	do while not rsSpecial.eof

		if instr(session("AdminPurview_SpecialID"),rsSpecial("SpecialID"))>0 then 
				Response.Write("<input type='checkbox'  name='AdminPurview_SpecialID' value='" & rsSpecial("SpecialID") & "' checked >" & rsSpecial("SpecialName") & "<br/>"  )
			else
				Response.Write("<input type='checkbox'  name='AdminPurview_SpecialID' value='" & rsSpecial("SpecialID") & "'")
				 if AdminPurview=1 then 
					 	Response.Write(" >")
					 end if
					 if AdminPurview=2 then 
					    Response.Write(" disabled >")
					 end if
				 Response.Write( rsSpecial("SpecialName") & "<br/>"  )
			end if
		rsSpecial.movenext
	loop

'response.write rsAdminPurview_Special("AdminPurview_Special")
'response.end

'关闭数据库连接
	rsSpecial.close
    set rsSpecial = nothing
'关闭数据库连接

end sub

sub Admin_ShowSpecial_Checkbox(AdminID)
'数据库连接
	dim sqlSpecial,rsSpecial

		sqlSpecial = "select SpecialID,SpecialName from Special"

	set rsSpecial=server.CreateObject("adodb.recordset")
	rsSpecial.open sqlSpecial,conn,1,1
'已开启数据库连接
	
	
	'数据库连接
	dim sqlAdminPurview_Special,rsAdminPurview_Special
		sqlAdminPurview_Special = "select AdminPurview_Special,AdminPurview_SpecialID from Admin where Id=" & AdminID
	
	set rsAdminPurview_Special=server.CreateObject("adodb.recordset")
	rsAdminPurview_Special.open sqlAdminPurview_Special,conn,1,1
'已开启数据库连接

	
		
'	dim SpecialNameTotal
	
'	do while not rsSpecial.eof
'		SpecialNameTotal= rsSpecial("SpecialName") & ","
'		rsSpecial.movenext
'	loop
	rsSpecial.movefirst
	rsAdminPurview_Special.movefirst
	do while not rsSpecial.eof


		if (instr(rsAdminPurview_Special("AdminPurview_SpecialID"),rsSpecial("SpecialID")) >0) then
		
			Response.Write("<input type='checkbox' name='AdminPurview_SpecialID' value='" & rsSpecial("SpecialID") & "' checked >" & rsSpecial("SpecialName") & "<br/>"  )
		else
			Response.Write("<input type='checkbox'  name='AdminPurview_SpecialID' value='" & rsSpecial("SpecialID") & "' "  ) 
			if AdminPurview=1 then 
					 	Response.Write(" >")
					 end if
					 if AdminPurview=2 then 
					    Response.Write(" disabled >")
					 end if
			Response.Write( rsSpecial("SpecialName") & "<br/>"  )
		end if	
'				if (instr(rsAdminPurview_Special("AdminPurview_Special"),rsSpecial("SpecialName")) >0) then
'		
'			Response.Write("<input type='checkbox' name='AdminPurview_Special' value='" & rsSpecial("SpecialName") & "' checked >" & rsSpecial("SpecialName") & "<br/>"  )
'		else
'			Response.Write("<input type='checkbox'  name='AdminPurview_Special' value='" & rsSpecial("SpecialName") & "'>" & rsSpecial("SpecialName") & "<br/>"  )
'		end if	

		rsSpecial.movenext
	loop

'response.write rsAdminPurview_Special("AdminPurview_Special")
'response.end

'关闭数据库连接
	rsSpecial.close
    set rsSpecial = nothing
'关闭数据库连接

'关闭数据库连接
	rsAdminPurview_Special.close
    set rsAdminPurview_Special = nothing
'关闭数据库连接


end sub


sub Admin_ShowSpecial_Option(ShowType,SpecialID)
	dim UserLevel
	UserLevel=request.Cookies("asp163")("UserLevel")
	if UserLevel="" then
		UserLevel=5000
	else
		UserLevel=Cint(UserLevel)
	end if
	response.write "<select name='SpecialID' id='SpecialID'><option value=''"
	if SpecialID=0 then
		response.write " selected"
	end if
	response.write "><font color=#0000ff>请选择课程</font></option>"
	                
	dim sqlSpecial,rsSpecial
    if ShowType=1 then
		sqlSpecial = "select * from Special"
	else
		sqlSpecial="select * from Special where AddPurview>=" & UserLevel
	end if	
	set rsSpecial=server.CreateObject("adodb.recordset")
	rsSpecial.open sqlSpecial,conn,1,1
	rsSpecial.movefirst
	do while not rsSpecial.eof
		if rsSpecial("SpecialID")=SpecialID then
			response.write "<option value='" & rsSpecial("SpecialID") & "' selected>" & rsSpecial("SpecialName") & "</option>"
		else
			response.write "<option value='" & rsSpecial("SpecialID") & "'>" & rsSpecial("SpecialName") & "</option>"
		end if
		rsSpecial.movenext
	loop
	rsSpecial.close
    set rsSpecial = nothing
end sub

sub Admin_ShowChild()
	dim sqlChild,rsChild
	sqlChild="select ClassID,ClassName,Child From ArticleClass where ParentID=" & ClassID & " order by OrderID"
	Set rsChild= Server.CreateObject("ADODB.Recordset")
	rsChild.open sqlChild,conn,1,1
	i=0
	do while not rsChild.eof
		response.Write "&nbsp;&nbsp;<a href='" & FileName & "?ClassID=" & rsChild(0) & "'>" & rsChild(1) & "</a>"
		if rsChild(2)>0 then
			response.write "(" & rsChild(2) & ")"
		else
			if ChildID="" then
				ChildID=Cstr(rsChild(0))
			else
				ChildID=ChildID & "," & Cstr(rsChild(0))
			end if
		end if		
		rsChild.movenext
		i=i+1
		if i mod 8=0 then
			response.write "<br>"
		else
			response.write "&nbsp;&nbsp;"
		end if
	loop
	rsChild.close
	set rsChild=nothing
end sub


sub Admin_ShowChild2()
	response.write "<table width='100%' border='0' align='center' cellpadding='5' cellspacing='1'>"
	response.write "  <tr class='tdbg'>"
	dim sqlChild,rsChild,rsChild2
	sqlChild="select ClassID,ClassName,Child From ArticleClass where ParentID=" & ClassID
	Set rsChild= Server.CreateObject("ADODB.Recordset")
	set rsChild2= Server.CreateObject("ADODB.Recordset")
	rsChild.open sqlChild,conn,1,1
	i=0
	do while not rsChild.eof
		response.Write "<td width='50%' valign='top'><a href='" & FileName & "?ClassID=" & rsChild(0) & "'><b><font color=red>" & rsChild(1) & "</b></font></a>"
		if rsChild(2)>0 then
			response.write "(" & rsChild(2) & ")<br>"
			sqlChild="select ClassID,ClassName,Child From ArticleClass where ParentID=" & rsChild(0)
			rsChild2.open sqlChild,conn,1,1
			j=0
			do while not rsChild2.eof
				response.Write "<a href='" & FileName & "?ClassID=" & rsChild2(0) & "'>" & rsChild2(1) & "</a>"
				if rsChild2(2)>0 then
					response.write "(" & rsChild2(2) & ")"
				end if
				response.write "&nbsp;&nbsp;&nbsp;&nbsp;"
				rsChild2.movenext
				j=j+1
				if j mod 5=0 then response.write "<br>"
			loop
			rsChild2.close
		else
			if ChildID="" then
				ChildID=Cstr(rsChild(0))
			else
				ChildID=ChildID & "," & Cstr(rsChild(0))
			end if
		end if		
		rsChild.movenext
		i=i+1
		response.write "</td>"
		if i mod 2=0 then
			response.write "</tr><tr class='tdbg'>"
		end if
	loop
	if i mod 2<>0 then response.write "<td>&nbsp;</td>"
	
	rsChild.close
	set rsChild=nothing
	set rsChild2=nothing
  	response.write "</tr></table>"	
end sub

sub Admin_ShowPath2(strParentPath,strClassName,iDepth)
	if iDepth<=0 then
		response.write strClassName
		exit sub
	end if
	dim sqlPath,rsPath,i
	sqlPath="select * From ArticleClass where ClassID in (" & strParentPath & ") order by Depth"
	set rsPath=server.createobject("adodb.recordset")
	rsPath.open sqlPath,conn,1,1
	do while not rsPath.eof
		for i=1 to rsPath("Depth")
			response.write "&nbsp;&nbsp;&nbsp;"
		next
		if rsPath("Depth")>0 then
			response.write "└"
		end if
		response.Write rsPath("ClassName") & "<br>"
		rsPath.movenext
	loop
	rsPath.close
	set rsPath=nothing
	if iDepth>0 and strClassName<>"" then
		for i=1 to iDepth
			response.write "&nbsp;&nbsp;&nbsp;"
		next
		response.write "└" & strClassName
	end if
end sub

sub AddMaster(ClassMaster)
	dim arrClassMaster,rsAdmin
	if instr(ClassMaster,"|")>0 then
		arrClassMaster=split(ClassMaster,"|")
		ClassMaster=""
		for i=0 to ubound(arrClassMaster)
			set rsAdmin=conn.execute("select * from Admin where UserName='" & arrClassMaster(i) & "'")
			if rsAdmin.bof and rsAdmin.eof then
				FoundErr=True
				ErrMsg=ErrMsg & "<br><li>管理员“" & arrClassMaster(i) & "”不存在！是否输入错了？</li>"
			else
				if rsAdmin("Purview")>4 then
					FoundErr=True
					ErrMsg=ErrMsg & "<br><li>“" & arrClassMaster(i) & "”权限不够！不能设为栏目编辑。"
				else
					if rsAdmin("Purview")=4 then
						if ClassMaster="" then
							ClassMaster=arrClassMaster(i)
						else
							ClassMaster=ClassMaster & "|" & arrClassMaster(i)
						end if
					end if
				end if
			end if
		next
	else
		set rsAdmin=conn.execute("select * from Admin where UserName='" & ClassMaster & "'")
		if rsAdmin.bof and rsAdmin.eof then
			FoundErr=True
			ErrMsg=ErrMsg & "<br><li>管理员“" & ClassMaster & "”不存在！是否输入错了？</li>"
		else
			if rsAdmin("Purview")>4 then
				FoundErr=True
				ErrMsg=ErrMsg & "<br><li>“" & ClassMaster & "”权限不够！不能设为栏目编辑。"
			else
				if rsAdmin("Purview")<4 then
					ClassMaster=""
				end if
			end if
		end if
	end if
end sub


sub Admin_ShowSearchForm(Action,ShowType)
	if ShowType<>1 and ShowType<>2 and ShowType<>3 then
		ShowType=1
	end if
	response.write "<table border='0' cellpadding='0' cellspacing='0'>"
	response.write "<form method='Get' name='SearchForm' action='" & Action & "'>"
	response.write "<tr><td height='28' align='center'>"
	if ShowType=1 then
		response.write "<input type='hidden' name='field' value='Title'>"
		response.write "<input type='text' name='keyword'  size='20' value='关键字' maxlength='50' onFocus='this.select();'>"
		response.write "<input type='submit' name='Submit'  value='搜索'>"
		response.write "<br>高级搜索"
	elseif Showtype=2 then
		response.write "<select name='Field' size='1'>"
    	response.write "<option value='Title' selected>文章标题</option>"
	    response.write "<option value='Content'>文章内容</option>"
    	response.write "<option value='Author'>文章作者</option>"
    	if AdminPurview<=3 or AdminPurview_Article<=2 then	
			response.write "<option value='Editor'>编辑姓名</option>"
		end if
		response.write "</select>"
		response.write "<select name='ClassID'><option value=''>所有栏目</option>"
		call Admin_ShowClass_Option(1,0)
		response.write "<input type='text' name='keyword'  size='20' value='关键字' maxlength='50' onFocus='this.select();'>"
		response.write "<input type='submit' name='Submit'  value='搜索'>"
	end if
	response.write "</td></tr></form></table>"
end sub

sub DelFiles(strUploadFiles)
	if strUploadFiles="" then exit sub
	if DelUpFiles="Yes" and ObjInstalled=True then
		dim fso,arrUploadFiles,i
		Set fso = CreateObject("Scripting.FileSystemObject")
		if instr(strUploadFiles,"|")>1 then
			arrUploadFiles=split(strUploadFiles,"|")
			for i=0 to ubound(arrUploadFiles)
				if fso.FileExists(server.MapPath("" & arrUploadfiles(i))) then
					fso.DeleteFile(server.MapPath("" & arrUploadfiles(i)))
				end if
			next
		else
			if fso.FileExists(server.MapPath("" & strUploadfiles)) then
				fso.DeleteFile(server.MapPath("" & strUploadfiles))
			end if
		end if
		Set fso = nothing
	end if
end sub

sub Admin_ShowSkin_Option(SkinID)
	response.write "<select name='SkinID' id='SkinID'>"
	dim sqlSkin,rsSkin
	sqlSkin="select * from Skin"
	set rsSkin=server.CreateObject("adodb.recordset")
	rsSkin.open sqlSkin,conn,1,1
	if rsSkin.bof and rsSkin.eof then
	 	SkinCount=0
	  	response.write "<option value=''>请先添加栏目配色模板</option>"
	else
	  	SkinCount=rsSkin.recordcount
	  	do while not rsSkin.eof
	  		if rsSkin("SkinID")=SkinID then
				response.write "<option value='" & rsSkin("SkinID") & "' selected>" & rsSkin("SkinName") & "</option>"
			else		
				response.write "<option value='" & rsSkin("SkinID") & "'>" & rsSkin("SkinName") & "</option>"
	  		end if		
			rsSkin.movenext
	  	loop
	end if
	rsSkin.close
	set rsSkin=nothing
    response.write "</select> <input name='PreviewSkin' type='button' id='PreviewSkin' value='查看效果图' onclick=""alert('此功能尚在设计中，敬请期待！');"">"
end sub

sub Admin_ShowLayout_Option(LayoutType,LayoutID)
	response.write "<select name='LayoutID' id='LayoutID'>"
	dim sqlLayout,rsLayout
	sqlLayout="select * from Layout where LayoutType=" & LayoutType
	set rsLayout=server.CreateObject("adodb.recordset")
	rsLayout.open sqlLayout,conn,1,1
	if rsLayout.bof and rsLayout.eof then
	  	LayoutCount=0
	  	response.write "<option value=''>请先添加栏目版面设计模板</option>"
	else
	  	LayoutCount=rsLayout.recordcount
	  	do while not rsLayout.eof
	  		if rsLayout("LayoutID")=LayoutID then
				response.write "<option value='" & rsLayout("LayoutID") & "' selected>" & rsLayout("LayoutName") & "</option>"
			else		
				response.write "<option value='" & rsLayout("LayoutID") & "'>" & rsLayout("LayoutName") & "</option>"
	  		end if		
			rsLayout.movenext
	  	loop
	end if
	rsLayout.close
	set rsLayout=nothing
    response.write "</select> <input name='PreviewLayout' type='button' id='PreviewLayout' value='查看效果图' onclick=""alert('此功能尚在设计中，敬请期待！');"">"
end sub
%>
