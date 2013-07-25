<% 
On Error Resume Next
if VarType(pageType) = 0 then
	Execute "Dim pageType"
end if
if VarType(includeContentType) = 0 then
	Execute "dim includeContentType"
end if
On Error Goto 0

if VarType(pageType) = 1 OR VarType(pageType) = 0 then
	pageType = "default"
end if


if request.QueryString("isLibAsync") = "true" then
	pageType = "async"
end if

if request.QueryString("isLibAsync") = "true" AND request.QueryString("isJson") = "true" then
	if Session("StudioID") = "" then
		response.ContentType = "application/json"
		response.Write("{")
		response.Write("""success"": false,")
		response.Write("""sessionExpired"": true")
		response.Write("}")
		response.End
	elseif Session("StudioID") <> request.QueryString("studioId") then
		response.ContentType = "application/json"
		response.Write("{")
		response.Write("""success"": false,")
		response.Write("""sessionWrong"": true")
		response.Write("}")
		response.End
	end if
end if



if pageType <> "async" then
if includeContentType="" then
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<%
	dim init_id_network : Set init_id_network = CreateObject("Wscript.Network") 
	response.write "<!-- " & init_id_network.ComputerName & " -->"
end if

' Add some login protection to adm/main_ pages
' this should be done to all adm pages, but some are shared, and due to the short period between this fix
' and it's release into production, will limited it just to main_ pages
if Session("StudioID") = "" OR (InStr(Lcase(request.ServerVariables("SCRIPT_NAME")),"adm/main_") > 0 AND (NOT session("Pass") OR session("Admin") = "false")) then
'if Session("StudioID") = "" then
%>
<script type="text/javascript">
	parent.resetSession();
</script>
<%
	response.End
end if
dim subTabLoc

if isNum(request.QueryString("tabID")) then
    session("tabID") = request.QueryString("tabID")
end if

' only set currentTabHref if we have a tabID, otherwise we're on some other processing page that doesn't need it
' don't set currentTabHref if being rendered through preContents.asp (mvc page) as those controller actions should be setting it
if isNum(request.QueryString("tabID")) and instr(request.ServerVariables("URL"), "preContents") = 0 then
	dim currentTabHref : currentTabHref = request.ServerVariables("URL")
	if request.ServerVariables("QUERY_STRING")<>"" then
		currentTabHref = currentTabHref & "?" & request.ServerVariables("QUERY_STRING")
	end if
	session("currentTabHref") = currentTabHref
end if


end if
%>

<!-- #include file="inc_init_functions.asp" -->

<%
dim allHotWords : allHotWords = getAllHotWords()

%>
