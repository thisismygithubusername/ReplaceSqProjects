<%@ CodePage=65001 %>
<!-- #include file="inc_simple_logging.asp" -->
<!-- #include file="inc_dbconn_str.asp" -->
<%
'Dim SessionFarm
'set SessionFarm = Server.CreateObject("SessionFarm.SFSession")
response.charset="utf-8"
	
wsurl = "/ASP/ws.asp?studioid=" & request.querystring("studioid") 
'session.abandon
'sessionFarm.Abandon()
if NOT DISABLE_3153 then
	dim cookieName
	for each cookieName in request.Cookies
		if InStr(cookieName, "S" & request.querystring("studioid") & "_")=1 then
			response.AddHeader "Set-Cookie", cookieName & "=;path=/;expires=Thu, 01-Jan-70 00:00:01 GMT"
		end if
		if InStr(cookieName, "H" & request.querystring("studioid") & "_")=1 then
			response.AddHeader "Set-Cookie", cookieName & "=;path=/;expires=Thu, 01-Jan-70 00:00:01 GMT"
		end if
	next
end if

response.AddHeader "Set-Cookie", "SessionFarm%5FGUID=;path=/;expires=Thu, 01-Jan-70 00:00:01 GMT"
response.AddHeader "Cache-Control", "no-cache, must-revalidate"
response.Status = 302
response.AddHeader "Location", wsurl

response.Flush
response.End
%>
