<!-- #include file="inc_dbconn.asp" -->
<!-- #include file="inc_has_arrivals.asp" -->
<!-- #include file="inc_build_session.asp" -->
<%
buildSession(Request.QueryString("StudioID"))

Response.ContentType = "text/css"
%>
