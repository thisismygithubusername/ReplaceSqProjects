<%if isPdf then%>
<base href="http://<%=getLocalhostString()%>"  />
<%end if %>

<!-- #include file="inc_css.asp" -->
<!-- #include file="inc_jquery.asp" -->
<!-- #include file="inc_tabs.asp" -->
<!-- #include file="inc_implementation_switch_js.asp" -->


<style type="text/css">
table.transForm { border: 0 !important; }
/* * { margin: 0; padding: 0; } */
#memoryticker{
font: bold 11px Arial;
color: <%=session("pageColor")%>;
border:;
padding: 3px;
}
#memoryticker a:link {color: <%=session("pageColor")%>; text-decoration: none;}
#memoryticker a:active {color: <%=session("pageColor")%>; text-decoration: none;}
#memoryticker a:visited {color: <%=session("pageColor")%>; text-decoration: none;}
#memoryticker a:hover {color: <%=session("pageColor3")%>; text-decoration: none;}

.tinytableopacity22    {
background: transparent;
background:url("http<%=addS%>://<%=request.servervariables("SERVER_NAME")%>/asp/images/columnbg22.gif");
}
.tinytableopacity22 TD {
filter:alpha(opacity=85);
-moz-opacity:.85;
opacity:.85;
}


<% if session("MBOBranding")="False" then %>
#footer 
{
	display: none;
	height: 0;
	margin: 0;
}
#wrapper-bottompad
{
	padding-bottom: 0;
}
#wrapper-minheight
{
	min-height: 0;
}

<% elseif isConsumerMode OR isMobileBrowser then %>

#pageWrapper #wrapper-minhieght 
{
	min-height: 0;
}
BODY #footer {
	visibility: visible;
}

<% end if %>

</style>
<script type="text/javascript">
function subForm() {
	document.search2.submit();
<% if Session("Pass") then %>
	parent.mainFrame.focus();
<% else %>
	parent.mainFrame.document.frmLogon.requiredtxtUserName.focus();
<% end if %>
}
function addToFavorites(sType) {
	var favoritesURL = "http<%=addS%>://<%=request.servervariables("SERVER_NAME")%>/ws.asp?studio=<%=session("studioShort")%>";
	if (sType!=0) {
		favoritesURL += "&sType=" + sType;
	}
	var favoritesName = "<%=Replace(session("StudioName"), """", "''")%> Online";
	if (window.sidebar) { // firefox
		window.sidebar.addPanel(favoritesName, favoritesURL, "");
	} else if(window.opera && window.print) { // opera
		var elem = document.createElement('a');
		elem.setAttribute('href',favoritesURL);
		elem.setAttribute('title',favoritesName);
		elem.setAttribute('rel','sidebar');
		elem.click();
	} else if(document.all) {	// ie
		window.external.AddFavorite(favoritesURL, favoritesName);
	}
}
function goTo(top_loc,tabID,qstrParam, botPage) {
	var clientId = "";
	if (document.getElementById("optSelectedClients") != null) {
		if (document.getElementById("optSelectedClients").selectedIndex >= 0) {
			clientId = "&fcltid=" + document.getElementById("optSelectedClients").options[document.getElementById("optSelectedClients").selectedIndex].value;
		}
	} 
	if (botPage!="" && botPage!=null) {
		parent.mainFrame.location = "/asp/adm/" + botPage + "?fl=true&tabID=" + tabID + "&qParam=" + qstrParam + clientId;
	} else {
		if (top_loc == "appts")
		{
			<% if session("optVersion9") = "v2.0" then %>
				<% if Session("UseBookMultipleAppt") = "true" then %>
					parent.mainFrame.location = "/MainAppointments/Index?fl=true&tabID=" + tabID + "&qParam=" + qstrParam + clientId;
				<%else %>
					parent.mainFrame.location = "/asp/adm/main_" + top_loc + "_new_alt0.asp?fl=true&tabID=" + tabID + "&qParam=" + qstrParam + clientId;
				<%end if %>
			<% elseif session("optVersion9") = "v1.0" then %>
				parent.mainFrame.location = "/asp/adm/main_" + top_loc + ".asp?fl=true&tabID=" + tabID + "&qParam=" + qstrParam + clientId;
			<% elseif session("DefaultApptVersion") = "2" then %>
				<% if Session("UseBookMultipleAppt") = "true" then %>
					parent.mainFrame.location = "/MainAppointments/Index?fl=true&tabID=" + tabID + "&qParam=" + qstrParam + clientId;
				<%else %>
					parent.mainFrame.location = "/asp/adm/main_" + top_loc + "_new_alt0.asp?fl=true&tabID=" + tabID + "&qParam=" + qstrParam + clientId;
				<%end if %>
			<% else %>
				parent.mainFrame.location = "/asp/adm/main_" + top_loc + ".asp?fl=true&tabID=" + tabID + "&qParam=" + qstrParam + clientId;
			<% end if %>
		}
		else
		{
			parent.mainFrame.location = "/asp/adm/main_" + top_loc + ".asp?fl=true&tabID=" + tabID + "&qParam=" + qstrParam + clientId;
		}
		
	}
}
function isPdf() {
	return <%= LCase(isPdf()) %>;
}

	// To handle client-side dates, we need to include some JS. I've picked Moment.js to do the client-side date manipulation,
	// but it needs some information like our time offset. All client-side date and time manipulation should be in UTC.
	window.MomentJSDateFormatString = '<%	select case session("dateFormatCode")
			case 1	
				RW("DD/MM/YYYY")
			case 2
				RW("MM/DD/YYYY")
			case 3
				RW("YYYY-MM-DD")
			case 4
				RW("YYYY/MM/DD")
			case 5
				RW("DD.MM.YYYY")
			case 6
				RW("DD-MM-YYYY")
		end select
	%>';

	// UTCOffset is PST (-8) + Site Offset Time
		window.UTCOffsetMinutes = <%=(-7 * 60) + session("tzOffset") %>;
</script>


<%= js(array("site","inc_login_content")) %>
<%
if session("Admin")<>"false" then
%>
<%= js(array("adm/adm_site")) %>

<%
else
%>
<%= js(array("cm_site")) %>
<%
end if
%>
