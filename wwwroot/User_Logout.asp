<%
	Response.Cookies("asp163")("UserName")=""
	Response.Cookies("asp163")("Password")=""
	Response.Cookies("asp163")("UserLevel")=""
	
	Response.Cookies("aspsky")("username")=""
	Response.Cookies("aspsky")("password")=""
	Response.Cookies("aspsky")("userclass")=""
	Response.Cookies("aspsky")("userid")=""
	Response.Cookies("aspsky")("userhidden")=""
	Response.Cookies("aspsky")("usercookies")=""
	session("userid")=""
	
	dim ComeUrl
	ComeUrl=trim(request("ComeUrl"))
	if ComeUrl="" then
		ComeUrl=Request.ServerVariables("HTTP_REFERER")
		if ComeUrl="" then ComeUrl="./" end if
	end if
	Response.Redirect ComeUrl
%>