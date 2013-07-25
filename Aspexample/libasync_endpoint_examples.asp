<!-- #include file="json2.asp" -->
<!-- #include file="inc_post.asp" -->
<!-- #include file="inc_i18n.asp" -->
<!-- #include file="inc_dbconn.asp" -->
<%
if NOT isNull(request.QueryString("test1")) then
	' do a whole bunch of work
	' return json data that libasync can work worth
	' at the very least we need to return an object with it's 
	response.Write("{")
	response.Write("'success': true"
	response.Write("}")
	response.End
end if
if NOT isNull(request.QueryString("test2")) then
end if
if NOT isNull(request.QueryString("test3")) then
end if

%>
