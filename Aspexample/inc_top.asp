<%
dim tableStyle
if session("Admin")<>"false" then 
	tableStyle ="width: 100%;"
else
	tableStyle = "max-width: 960px; margin: 0 auto;"
end if
%>
<style type="text/css">
	.logo-table{color: <%=session("pageColor")%>; }
	.logo-table a:hover{text-decoration:underline;}
	.siteDeactivatedMsg
	{
		color:red;
		font-size: 20px;
		font-weight: bold;
	}
</style>
<div id="topSectionBG">
	<div id="topWrap">
	<table id="top-section-table" cellspacing="0" style='<%= tableStyle %>'>
		<tr>
			<td class="top-section-table-ends">
				<div id="top-logo-container">
					<!-- #include file="inc_logo.asp" -->
				</div>
			</td>
			<td>
			<% ' This is a terrible, terrible hack to make noframes work the same way as clients. - MB %>
			<% if NOT isConsumerMode then %>
				<table style="margin: 0 auto;">
					<tr>
						<td>
			<% else %>
				<div id="top-bb-container">
			<% end if %>
			<!-- #include file="adm/inc_bb.asp" -->
			<% if NOT isConsumerMode then %>
						</td>
					</tr>
				</table>
			<% else %>
				</div>
			<% end if %>

			</td>
			<td id="top-login-container-td" class="top-section-table-ends">
				<div id="top-login-container">
					<!-- #include file="inc_login.asp" -->
				</div>
				<div style="clear:both;"></div>
			</td>
		</tr>
	</table>
	</div>
	<!-- #include file="inc_topnav_div.asp" -->
</div>
<%


%>
<%if session("Admin")="false" then %>
	<div id="tabBottomBorder"></div>

	<div class="menu">
		<%if NOT session("freeScheduler") then %>
		<div class="menu-right">
			<a href="javascript:printWindow();" style="float: right; margin-top: 2px; margin-right: 17px;"><img src="<%= contentUrl("/asp/images/printer-20px.png") %>" title="<%=xssStr(allHotWords(809))%>"></a>
		</div>
		<%end if %>
	</div>
<%else %>
	<table class="sub-tab-table-biz" id="sub-tab-table-biz">
	<tr>
		<td id="printerTD">
			<a href="javascript:printWindow();" class="textSmall"><img src="<%= contentUrl("/asp/images/printer-20px.png") %>" title="<%=xssStr(allHotWords(809))%>"></a>
		</td>
		<td id="curClientTD">
<%
	if (subTabLoc="" OR subTabLoc="noload") AND NOT ss_HomeTabDefaultToProfile then
			frmActionStr = "main_info.asp"
	elseif subTabLoc="sch" then
			frmActionStr = "adm_clt_sch.asp"
	elseif subTabLoc="vh" then 
			frmActionStr = "adm_clt_vh.asp"
	elseif subTabLoc="ph" then
			frmActionStr = "adm_clt_ph.asp"
	elseif subTabLoc="pappt" then
			frmActionStr = "adm_clt_past_appt.asp"
	elseif subTabLoc="canc" then
			frmActionStr = "adm_clt_canc.asp"
	elseif subTabLoc="purch" then
			frmActionStr = "adm_clt_purch.asp"
	elseif subTabLoc="cl" then
			frmActionStr = "adm_clt_conlog.asp"
	elseif subTabLoc="files" then
			frmActionStr = "adm_clt_files.asp"
	else
			frmActionStr = "adm_clt_profile.asp"
	end if
%>
			<form name="topsearch2" method="post" target="mainFrame" action="<%=frmActionStr%>">
				<input type="hidden" name="pageNum" value="1" />
				<input type="hidden" id="tabID" name="tabID" value="<%=session("tabID")%>" />
				<input type="hidden" name="scrLeft" id="scrLeft" value="<%= xssStr(request.form("scrLeft")) %>" />
				<!-- #include file="adm/inc_cur_clt.asp" -->
			</form>
		</td>
	</tr>
	</table>
	<div id="admMenu"></div>
<% end if 'session("Admin")="false" %>

