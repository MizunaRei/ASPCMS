<%@ LANGUAGE="VBScript" %>
<%
Option Explicit

dim DiaryOwner
DiaryOwner=request("DiaryOwner")

Dim dtToday
dtToday = Date()

Dim dtCurViewMonth		' First day of the currently viewed month
Dim dtCurViewDay		' Current day of the currently viewed month
%>


<% REM This section defines functions to be used later on. %>
<% REM This sets the Previous Sunday and the Current Month %>
<%

'--------------------------------------------------
   Function DtPrevSunday(ByVal dt)
      Do While WeekDay(dt) > vbSunday
         dt = DateAdd("d", -1, dt)
      Loop
   DtPrevSunday = dt
   End Function
'--------------------------------------------------

%>

<%REM Set current view month from posted CURDATE,  or
' the current date as appropriate.

' if posted from the form
' if prev button was hit on the form
   If InStr(1, Request.Form, "subPrev", 1) > 0 Then
      dtCurViewMonth = DateAdd("m", -1, Request.Form("CURDATE"))
' if next button was hit on the form
   ElseIf InStr(1, Request.Form, "subNext", 1) > 0 Then
      dtCurViewMonth = DateAdd("m", 1, Request.Form("CURDATE"))
' anyother time
      Else
         dtCurViewMonth = DateSerial(Year(dtToday), Month(dtToday), 1)
   End If
%>


<% REM --------BEGINNING OF DRAW CALENDAR SECTION-------- %>
<% REM This section executes the event query and draws a matching calendar. %>
<%
   Dim iDay, iWeek, sFontColor
%>

<HTML>
<HEAD>
<title>·<%=DiaryOwner%>的日历·</title>
<style>
BODY {scrollbar-track-color:#ffffff; SCROLLBAR-FACE-COLOR: #ffffff; FONT-SIZE: 9pt; SCROLLBAR-HIGHLIGHT-COLOR: #ffffff; SCROLLBAR-SHADOW-COLOR: #eeeeee;  SCROLLBAR-3DLIGHT-COLOR: #eeeeee; SCROLLBAR-ARROW-COLOR: #dddddd; FONT-FAMILY: "Verdana"; SCROLLBAR-DARKSHADOW-COLOR: #ffffff
}
select{font-size:8pt;font-family:verdana;background-color:#ffffff;border:1px dotted #cccccc; color:#333333;}
input{font-size:8pt;font-family:verdana;background-color:#ffffff;border-bottom:1px solid #51bfe0;border-left:1px solid #51bfe0; border-top:0px solid #cccccc;border-right:0px dotted #cccccc;color:#333333;}
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
<script language=javascript>
	function openScript(url)
	{
		opener.window.location.href(url);
		window.close();
	}
</script>
</HEAD>
<BODY leftmargin=10 topmargin=10 bgcolor="#FFFFFF">
<CENTER>
  <font face="Verdana" size="2">· <b><%=DiaryOwner%>&nbsp; 的 日 历 </b>·</font>
  <FORM NAME="fmNextPrev" ACTION="diary_calendar.asp?DiaryOwner=<%=DiaryOwner%>" METHOD=POST>
    <TABLE CELLPADDING=3 CELLSPACING=0 WIDTH="450" BORDER=2 BGCOLOR="#51bfe0" BORDERCOLORDARK="#51bfe0" BORDERCOLORLIGHT="#FFFFFF">
      <TR VALIGN=MIDDLE ALIGN=CENTER>
             <TD COLSPAN=7>
             <TABLE CELLPADDING=0 CELLSPACING=0 WIDTH="100%" BORDER=0>
                <TR VALIGN=MIDDLE ALIGN=CENTER>
                   <TD WIDTH="30%" ALIGN=RIGHT>
                <INPUT TYPE=submit NAME="subPrev" value=" 上 月 ">
                   </TD>
                   <TD WIDTH="40%">
                      <FONT FACE="verdana" COLOR="#333333" size=2>
                      <B><%= Year(dtCurViewMonth)& "年" & MonthName(Month(dtCurViewMonth))%></B>
                     </FONT>
                   </TD>
                   <TD WIDTH="26%" ALIGN=LEFT>

                <INPUT TYPE=submit NAME="subNext" value=" 下 月 ">
                   </TD>

              <TD WIDTH="4%" ALIGN=LEFT> <a href="" onclick="openScript('diary_index.asp?DiaryOwner=<%=DiaryOwner%>')"><img border="0" src="diary_images/home.gif" alt="返回日记首页" width="18" height="18"></a></TD>
                </TR>
             </TABLE>
             </TD>
          </TR>

          <TR VALIGN=TOP ALIGN=CENTER BGCOLOR="#003366">

          <% For iDay = vbSunday To vbSaturday %>

        <TH WIDTH="14%" bgcolor="#FFFFFF"><FONT FACE="Arial" SIZE="-2" COLOR="#333333"><%=WeekDayName(iDay)%></FONT></TH>
          <%Next %>

         </TR>

<%
   dtCurViewDay = DtPrevSunday(dtCurViewMonth)

   For iWeek = 0 To 5
      Response.Write "<TR>" & vbCrLf

      For iDay = 0 To 6
         Response.Write "<TD HEIGHT=35 align=center>"

         If Month(dtCurViewDay) = Month(dtCurViewMonth) Then
            If dtCurViewDay = dtToday Then
               sFontColor = "#FF3300"
            Else
               sFontColor = "#000000"
            End If
         	'---- Write day of month
            Response.Write "<FONT FACE=""verdana"" SIZE=""2"" COLOR=""" & sFontColor & """><B>"
            Response.Write Day(dtCurViewDay) & "</B></FONT>"

			If dtCurViewDay <= dtToday Then
				dim strUrl
				'response.write dtCurViewMonth
				strUrl = "diary_index.asp?DiaryOwner="&DiaryOwner
				strUrl = strUrl & "&diaryDate=" & (dtCurViewDay)
				Response.Write "<a href='' onclick='openScript("""&strUrl&""")' title='查看本日的日记'>"
				Response.Write "<img src=diary_images/mess.gif border=0 align=absmiddle></a>"
			Else
				Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;"
			end if
         Else
            Response.Write "&nbsp;"
         End If

         Response.Write "</TD>" & vbCrLf
         dtCurViewDay = DateAdd("d", 1, dtCurViewDay)
      Next
      Response.Write "</TR>" & vbCrLf
   Next
%>
<%REM --------END OF DRAW CALENDAR SECTION-------- %>
</TABLE>
<INPUT TYPE=HIDDEN NAME="CURDATE" VALUE="<%=dtCurViewMonth%>">
</FORM>
</CENTER>
<NOSCRIPT><IFRAME SRC=diary_calendar.asp></IFRAME></NOSCRIPT>
</BODY>
</HTML>
