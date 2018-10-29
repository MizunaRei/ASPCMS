
<%
sessionID = session.SessionID
timeout = 5
' set how long to keep this session in minute you can increase this number

Conn_String = "DRIVER={Microsoft Access Driver (*.mdb)};DBQ=" & Server.MapPath("count.mdb")
'Conn_String = "activeUser"
'set your DSN = "activeuser" is a better way because you need include this file to all your asp scripts.


Set ConnCount =Server.CreateObject("ADODB.Connection")
ConnCount.Open Conn_String

' delete session after timeout
aaa = dateadd("n", -timeout, now())
connCount.Execute ("delete * from count where postdate < #" & aaa & "#")


' keep sessionID
sql0 = "select sess from count where sess='" & sessionID & "'"
set rscheck = connCount.Execute (sql0)
if rscheck.eof then
sql = "insert into count (sess,postdate) values('" & sessionID & "', '" & now() & "')"
connCount.Execute (sql)
end if
rscheck.close
set rscheck = nothing

'count sessionID
sql2 = "select count(sess) from count"
set rs = connCount.Execute (sql2)
count = rs(0)
rs.close
set rs = nothing


sql3 = "select * from count"
set rspredel = connCount.Execute (sql3)
do until rspredel.eof
xxx=DateDiff("n", rspredel("postdate"), Now())
if xxx > timeout then
count = count-1
end if
rspredel.movenext
loop
rspredel.close
set rspredel = nothing

connCount.Close
set connCount = nothing

if count = 0 then
count = 1
end if
%>	
当前<%=count%>人在线
