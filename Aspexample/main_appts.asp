<%@ CodePage=65001 %>
<% Option Explicit %>
<!-- #include file="init.asp" -->
<!-- #include file="json/JSON.asp" -->
<%
'Dim SessionFarm
'set SessionFarm = Server.CreateObject("SessionFarm.SFSession")
response.charset="utf-8"
%>

<%
session("pageID")="_appts"
if isNum(request.form("tabID"))then
	session("tabID") = request.form("tabID")
elseif isNum(request.QueryString("tabID")) then
	session("tabID") = request.querystring("tabID")
end if
%>

		<!-- #include file="inc_internet_guest.asp" -->
		<!-- #include file="inc_i18n.asp" -->
		<!-- #include file="adm/inc_chk_holiday.asp" -->
        <% if session("CR_Memberships") <> 0 then %>
            <!-- #include file="inc_dbconn_regions.asp" -->
            <!-- #include file="inc_dbconn_wsMaster.asp" -->
            <!-- #include file="adm/inc_masterclients_util.asp" -->
        <% end if %>
		<!-- #include file="adm/inc_acct_balance.asp" -->
		<!-- #include file="adm/inc_crypt.asp" -->
		<!-- #include file="inc_loading.asp" -->
		<!-- #include file="inc_tinymcesetup.asp" -->
<%
    ' Load the phrases for this page and set them to local variables
    dim phraseDictionary
    set phraseDictionary = LoadPhrases("ConsumermodeappointmentschedulePage", 16)

	Dim loadSearch
	'check for load option
	if request.querystring("load")="sched" then
		loadSearch = false
	elseif checkStudioSetting("tblApptOpts", "CltModeDisableApptSch") then
		loadSearch = true
	else
		if checkStudioSetting("tblApptOpts", "CltModeApptSchedLoad") then
			loadSearch = false
		else
			if request.form("PageNum")<>"" then
				loadSearch = false
			else
				loadSearch = true
			end if
		end if
	end if

	'forward to appointment search if necessary
	if loadSearch then
		dim queryString , n, varItem, contents
		n=0
		For Each varItem in Request.QueryString
		
			On Error Resume Next
				contents = Request.QueryString(varItem)
			On Error Goto 0	
			queryString = queryString & Server.URLEncode(varItem) & "=" & Server.URLEncode(contents) 
			
			n=n+1
			
			if n <> Request.QueryString.Count then queryString = queryString +"&"
		Next				
%>
		<script type="text/javascript">
			document.location.replace("adm/adm_appt_search.asp?<%=queryString%>");
		</script>
<%
		response.end
	end if

%>
<!-- #include file="pre.asp" -->
<!-- #include file="frame_bottom.asp" -->
<!-- #include file="inc_date_ctrl.asp" -->
<!-- #include file="adm/inc_hotword.asp" -->
<!-- begin client alerts -->
<%
	'client alert context vars
	focusFrmElement = ""
	cltAlertList = setClientAlertsList(session("mvarUserID"))
%>
	<!-- #include file="inc_ajax.asp" -->
	<!-- #include file="adm/inc_alert_js.asp" -->
	<!-- end client alerts  -->
<%
	Dim curTrn, strUserAgent, rsApptClr, AvailColor, numRows, count, tmpcount, maskCount, icount, scheduleOffsetMins, i, j, k, u, rsCTrainers
	Dim numLeft, tmpCurTimeBlock, tmprcount, contUp, pass3, pass4, expandAvail, tmpSpan, cont, tmpIndex, firstTime, tmpcPage, icounter, pageStr
	Dim tmpSPage, tmpEPage, spacer, strWidth, countDays, loopCount, strColSpan, first, str1, str2, resClientID, curTrnName, ss_SchedShowDay, apptDefToWeek, rsEntry
	Dim sysDate, curDate, strCurWeekDay, strShortDate, curTG, curTGname, curTGOffset, curTGOffsetEnd, curTGBlockLength, tmpDate, minTime, constMinTime, minHour, maxHour, numTGs, tgIndex, tmpHour,calcEtime, availLength, maxTGBlockLength, minTGBlockLength, numTrainers
	Dim curTGCapacity, totalNoTrn, cPage, curMasked, confirmedColor, bookedColor, instructHW, testStr, tempCounter, blockSize, todaysDate, curTime, trnsPerPage, launchTG, cView, ss_hideAppts, SchHrsClosed, SchHrsStartTime, SchHrsEndTime, SameDaySchCutoff, trnCount
	Dim tgCapacity()
	Dim tgIDs()
	Dim tgNames()
	Dim tgOffset()
	Dim tgOffsetEnd()
	Dim tgBlockLength()

	set rsEntry = Server.CreateObject("ADODB.Recordset")
	strSQL = "SELECT HotWordText, HotWordID FROM tblHotWords WHERE (HotWordID = 18)"
	rsEntry.CursorLocation = 3
	rsEntry.open strSQL, cnWS
	Set rsEntry.ActiveConnection = Nothing
	if not rsEntry.EOF then
		instructHW = rsEntry("HotWordText")
	else
		instructHW = "Instructor"
	end if
	rsEntry.close
	strSQL = "SELECT tblApptOpts.SchedShowDay, tblApptOpts.apptSchDiscreteBlocks, tblApptOpts.apptDefaultToWeekView, tblApptOpts.apptBlockLength, tblApptOpts.apptSchTrnsPerPage, tblApptOpts.CltModeHideAppt, tblAppearance.confirmedColor, tblAppearance.bookedColor, tblAppearance.AvailColor, tblGenOpts.CustSchHours, tblGenOpts.CustSchHrsStart, tblGenOpts.CustSchHrsEnd, tblGenOpts.SameDaySchCutoff FROM tblApptOpts INNER JOIN tblAppearance ON tblApptOpts.StudioID = tblAppearance.StudioID INNER JOIN tblGenOpts ON tblApptOpts.StudioID = tblGenOpts.StudioID WHERE tblApptOpts.StudioID=" & session("studioID") & " AND tblApptOpts.StudioID=tblAppearance.StudioID"
	rsEntry.CursorLocation = 3
	rsEntry.open strSQL, cnWS
	Set rsEntry.ActiveConnection = Nothing
	if NOT rsEntry.EOF then
		ss_SchedShowDay = rsEntry("SchedShowDay")
		expandAvail = rsEntry("apptSchDiscreteBlocks")
		apptDefToWeek = rsEntry("apptDefaultToWeekView")
		confirmedColor = rsEntry("confirmedColor")
		bookedColor = rsEntry("bookedColor")
		AvailColor = rsEntry("AvailColor")
		blockSize = rsEntry("apptBlockLength")
		trnsPerPage = rsEntry("apptSchTrnsPerPage")

		'fix for consumer mode.  the setting "apptSchTrnsPerPage" can go up to 20.  This is great for biz mode,
		'but for the consumer mode resdesign, the table will overflow past the edge of the new, thinner layout.
		'15 is the highest number that still looks good when rendered.
		if trnsPerPage > 15 then
			trnsPerPage = 15
		end if

		ss_HideAppts = rsEntry("CltModeHideAppt")
		SchHrsClosed = rsEntry("CustSchHours")
		if SchHrsClosed then	'If NOT 24 Hr scheduling
			''Validate not null and stime < end time
			if NOT isNULL(rsEntry("CustSchHrsStart")) AND NOT isNULL(rsEntry("CustSchHrsEnd")) then
				if TimeValue(rsEntry("CustSchHrsStart")) < TimeValue(rsEntry("CustSchHrsEnd")) then
					'''Check if cur time is closed
					if TimeValue(DateAdd("n", Session("tzOffset"),Time)) < TimeValue(rsEntry("CustSchHrsStart")) OR TimeValue(DateAdd("n", Session("tzOffset"),Time)) > TimeValue(rsEntry("CustSchHrsEnd")) then
						'' Schedule is currently closed
						SchHrsStartTime = rsEntry("CustSchHrsStart")
						SchHrsEndTime = rsEntry("CustSchHrsEnd")
					else
						SchHrsClosed = false
					end if
				else
					SchHrsClosed = false
				end if
			else 'Null Time Val
				SchHrsClosed = false
			end if
		end if

		Dim SchOpenDOMCloseDate : SchOpenDOMCloseDate = "1/1/2200"
		if Session("SchOpenDOMCloseDate")<>"" then
			SchOpenDOMCloseDate = CDATE(Session("SchOpenDOMCloseDate"))
		end if

		if isNULL(rsEntry("SameDaySchCutoff")) then
			SameDaySchCutoff = false
		else
			if TimeValue(DateAdd("n", Session("tzOffset"),Time)) > TimeValue(rsEntry("SameDaySchCutoff")) then
				SameDaySchCutoff = true
			else
				SameDaySchCutoff = false
			end if
		end if
	else 'rsEntry.EOF
		confirmedColor = "#AC2F1E"
		bookedColor = "#000066"
		AvailColor = "CCAACC"
		blockSize = 60
		trnsPerPage = 4
		ss_HideAppts = false
		SchHrsClosed = false
		SameDaySchCutoff = false
	end if 'if NOT rsEntry.EOF
	rsEntry.close

	if request.querystring("tg")<>"" then
		launchTG = request.querystring("tg")
	else
		launchTG = ""
	end if

''''''''''''''''Setup input params''''''''''''''''''''
	if request.querystring("page") <> "" then
		cPage = request.querystring("page")
	elseif request.form.item("pageNum") <> "" then
		cPage = request.form.item("pageNum")
	else
		cPage = 1
	end if
	
	if request.querystring("view") <> "" then
		cView = request.querystring("view")
	elseif request.Form.item("optView") <> "" then
		cView = request.Form.item("optView")
	elseif apptDefToWeek then
		cView = "week"
	else
		if ss_SchedShowDay then
			cView = "day"
		else
			cView = "week"
		end if
	end if
	
	'MB setting session("curLocation") is moved to inc_loc_list.asp
%>
	<!-- #include file="inc_loc_list.asp" -->
<%
	setLocationSessionVar(false)

	' Pull out the current TG
	if request.querystring("tg")<>"" then
		curTG = CINT(request.querystring("TG"))
	elseif request.form("optTG")<>"" then
		curTG = CINT(request.form("optTG"))
	else
		curTG = 0
	end if
	if curTG>21 then
		curTG = 0
	end if

	' Pull out the current trainer
	if isNum(request.querystring("tid")) then
		curTrn = CLNG(request.querystring("tid"))
	elseif isNum(request.querystring("trn")) then
		curTrn = CLNG(request.querystring("trn"))
	elseif isNum(request.form("optInstructor")) then
		if inStr(request.form("optInstructor"),"m")>0 then
			curTrn = CLNG(Replace(request.form("optInstructor"),"m",""))
			curMasked = true
		else
			curTrn = CLNG(request.form("optInstructor"))
		end if
	else
		curTrn = 0
	end if
	strUserAgent = cstr(request.ServerVariables("HTTP_USER_AGENT"))

	curDate = DateValue(cFrmDate)
%>
	<%= js(array("mb", "MBS", "calendar" & dateFormatCode, "formchek", "VCC2")) %>
	<%= css(array("day_week_navigator","jquery.selectBox")) %>
	<script type="text/javascript">
		function autoSubmit2() {
			document.searchFrm.submit();
		}
	</script>
	<style type="text/css">
		.main-appts-table-grid .instructorLink
		{
			color: White;
			text-decoration: none;
		}
		.main-appts-table-grid a 
		{
			text-decoration: none;
		}
		.main-appts-table-grid a:hover
		{
			text-decoration: underline;
		}
		.topFilters
		{
			margin: 0 auto;
			position: relative;
			width: 960px;
		}
		.topFiltersInner
		{
			float: right;
			padding: 10px 0 32px 10px;
		}

		.dateControls
		{
			background-clip: padding-box;
			border-radius: 10px 10px 10px 10px;
			margin: 0 auto;
			padding: 5px 0;
			position: absolute;
			right: 10px;
			top: 25px;
			/*width: 335px;*/
			z-index: 10500;
		}

		.dateControls .leftSide
		{
			float: left;
			padding-top: 1px;
		}
		.dateControls .rightSide
		{
			float: right;
			margin-left: 4px;
			padding-top: 1px;
		}
		.cur-date
		{
			display: none;
		}
		#day-tog-c, #week-tog-c
		{
			color: #555;
		}
		h1
		{
			background: #FFFFFF;
			font-size: 24px;
			line-height: 32px;
			margin-top: 10px;
			padding: 21px 0 0 5px;
		}
		#main-content
		{
			padding-top: 45px;
		}
		.section
		{
			margin-bottom: 10px;
		}
		.avail
		{
			background-color:<%= AvailColor%>;
		}
		.empty
		{
			background-color:#999999;
		}
	</style>
	<!-- #include file="inc_desc_lightbox.asp" -->
	<!-- #include file="adm/inc_alert_content.asp" -->

<%

	dim lockLoc, ss_hideApptTGs, ss_apptDefaultToWeekView, ss_UpperCaseTabs, ss_ApptSchedGenderFilter, curTmpDate, ss_ShowDayViewInCM
	set rsEntry = Server.CreateObject("ADODB.Recordset")
	strSQL = "SELECT Studios.StudioURL, tblGenOpts.StudioLinkTab, tblGenOpts.HideCltHelp, tblGenOpts.HideCltForgotPwd, tblGenOpts.UpperCaseTabs, tblAppearance.LogoHeight, tblAppearance.LogoWidth, tblAppearance.topBGColor, tblGenOpts.ClientModeLockTGRemoveBuy, tblGenOpts.CltModeSigup, tblGenOpts.CltModeSigupExisting, tblGenOpts.BBenabled, tblGenOpts.ClientModeLockLoc, tblApptOpts.TrnOrderBy, tblApptOpts.hideApptTGs, tblApptOpts.SchedShowDay, tblApptOpts.apptDefaultToWeekView, tblApptOpts.ApptSchedGenderFilter FROM Studios INNER JOIN tblGenOpts ON Studios.StudioID = tblGenOpts.StudioID INNER JOIN tblAppearance ON Studios.StudioID = tblAppearance.StudioID INNER JOIN tblApptOpts ON Studios.StudioID = tblApptOpts.StudioID WHERE (Studios.StudioID = " & session("StudioID") & ")"
	rsEntry.CursorLocation = 3
	rsEntry.open strSQL, cnWS
	Set rsEntry.ActiveConnection = Nothing

	''Appt Specific
	ss_hideApptTGs = rsEntry("hideApptTGs")
	ss_ApptSchedGenderFilter = rsEntry("ApptSchedGenderFilter")
	ss_apptDefaultToWeekView = rsEntry("apptDefaultToWeekView")
	ss_ShowDayViewInCM = rsEntry("SchedShowDay")
	rsEntry.close
%>
	<!-- #include file="inc_fixed_bar.asp" -->
	<%= js(array("main_appts", "plugins/jquery.selectBox")) %>
	<script type="text/javascript">

		function updateView() {
			document.search2.action = "main_class.asp";
			document.search2.submit();
<%
			if Session("Pass") then
%>
				parent.mainFrame.focus();
<%
			else
%>
				loginFocus();
<%
			end if
%>
		}
		$(document).ready(function () {

			//
			// Enable selectBox control and bind events
			//

			$(".filterList SELECT").selectBox({ fixed: true });

		});
	</script>
	<% fixedSpacer %>
	<% pageStart %>
	<table width="<%=strPageWidth%>" cellspacing="0" height="100%" class="appointment-schedule">
		<tr>
			<td class="center-ch" valign="top" height="100%" width="100%" style="background-color: #FFFFFF;	padding-top: 15px;">
				<!-- #include file="inc_appt_sch.asp" -->								
			</td>
		</tr>
	</table>
	
	<form name="search2" method="post" action="main_appts.asp">
	<input type="hidden" name="pageNum" value="1" />
	<input type="hidden" name="requiredtxtUserName" value="" />
	<input type="hidden" name="requiredtxtPassword" value="" />
	<input type="hidden" name="optForwardingLink" value="" />
	<input type="hidden" name="optRememberMe" value="" />
	<input type="hidden" name="tabID" value="<%=session("tabID")%>" />
	<div class="wrapperTop">
		<div class="wrapperTopInner">
			<div class="pageTop">
				<div class="pageTopLeft">
					&nbsp;
				</div>
				<div class="pageTopRight">
					&nbsp;
				</div>
				<h1><%=DisplayPhrase(phraseDictionary,"Browseappointments")%></h1>
				<div id="dateControls" class="dateControls">
					<div class="leftSide">
						<% dayAndWeekControls %>
					</div>
					<div class="rightSide">
						<% dayInfo %>
					</div>
				</div>
			</div>
			<!-- pageTop -->
		</div>
		<!-- wrapperTopInner -->
	</div>
	<!-- wrapperTop -->
	<div class="fixedHeader">
		<div class="topFilters">
			<div class="topFiltersInner filterList">
				<% setLocationSessionVar(true) %>
<%
				Dim boolShowTG
				if ss_hideApptTGs then
				else
					boolShowTG = false
					''Check for > 1 TG otherwise don't display TG selector
					strSQL = "SELECT tblTypeGroup.TypeGroupID, tblTypeGroup.TypeGroup FROM tblTypeGroup "
					if session("tabID")<>"" AND isNumeric(session("tabID")) then
						strSQL = strSQL & " INNER JOIN tblTypeGroupTab ON tblTypeGroupTab.TypeGroupID = tblTypeGroup.TypeGroupID AND tblTypeGroupTab.TabID = " & session("tabID")
					end if
					strSQL = strSQL & "  WHERE active=1 AND wsAppointment=1 AND wsDisable<>1 ORDER BY TypeGroup"
					rsEntry.CursorLocation = 3
					rsEntry.open strSQL, cnWS
					Set rsEntry.ActiveConnection = Nothing

					if NOT rsEntry.EOF then
						rsEntry.MoveNext
						if NOT rsEntry.EOF then
							boolShowTG = true
							rsEntry.MoveFirst
						end if
					end if


					if boolShowTG then
%>
						<select name="optTG" id="optTG" class="bf-filter" onchange="subFormClrPNum();">
							<option value="0">
								<%= allHotWords(149) %>
							</option>
<%
							Do While NOT rsEntry.EOF
%>
								<option value="<%=rsEntry("TypeGroupID")%>" title="<%=stringIfLengthOverLimit(rsEntry("TypeGroup"),18)%>" <%if curTG=rsEntry("TypeGroupID") then response.write "selected" end if%>>
									<%=truncateString(rsEntry("TypeGroup"),18)%>
								</option>
<%
								rsEntry.MoveNext
							Loop
%>
						</select>
						<script type="text/javascript">
							document.search2.optTG.options[0].text = "<%=jsEsc(allHotWords(516))%>";
						</script>
<%
					end if '''1 or less hide
					rsEntry.close
				end if	'''hideTG
%>
				<select name="optInstructor" id="optInstructor" class="bf-filter" onchange="subForm();">
					<option value="0">
						<%= allHotWords(149) %>
					</option>
<%
					if session("tabID")<>"" AND isNumeric(session("tabID")) then
						dim TGStrTop
						TGStrTop = ""
						'strSQL = " SELECT DISTINCT TypeGroupID FROM tblTypeGroupTab WHERE TabID = " & session("tabID")
						'CB 7/9/2008 Updatdd to prevent error where session("tabID") is pointing to a non appointment tab which then creates bad references to tblTrainerSchedules.TGX
						strSQL = "SELECT DISTINCT tblTypeGroupTab.TypeGroupID FROM tblTypeGroupTab INNER JOIN tblTabs ON tblTypeGroupTab.TabID = tblTabs.TabID WHERE (tblTypeGroupTab.TabID = " & session("tabID") & ") AND (tblTabs.wsAppointment = 1)"
						rsEntry.CursorLocation = 3
						rsEntry.open strSQL, cnWS
						Set rsEntry.ActiveConnection = Nothing

						if NOT rsEntry.EOF then
							do while NOT rsEntry.EOF
								TGStrTop = TGStrTop & " OR tblTrainerSchedules.TG" & rsEntry("TypeGroupID") & " = 1 "
								rsEntry.MoveNext
							loop
						end if
						rsEntry.close
					end if

					curTmpDate = DateValue(CDATE(DateAdd("n", Session("tzOffset"),Now)))

					strSQL = "SELECT DISTINCT SortOrder, TrFirstName, TrLastName, DisplayName, CASE WHEN NOT TRAINERS.DisplayName IS NULL THEN TRAINERS.DisplayName ELSE TRAINERS.TrLastName END AS TrnDisName, tblTrainerSchedules.TrainerID FROM TRAINERS, tblTrainerSchedules WHERE TRAINERS.TrainerID=tblTrainerSchedules.TrainerID AND tblTrainerSchedules.ShowPublic=1 AND TRAINERS.[Delete]=0 AND TRAINERS.Active=1 AND tblTrainerSchedules.EndDate >= " & DateSep & DateAdd("y", -30, curTmpDate) & DateSep & "  AND TRAINERS.TrainerID<>-1  "
					strSQL = strSQL & " AND tblTrainerSchedules.Unavailable=0 "
					if TGStrTop<>"" then
						strSQL = strSQL & " AND ( 1 = 0 " & TGStrTop & " ) "
					end if
					strSQL = strSQL & "ORDER BY " 
					if ss_TrnOrderBy = 1 then
						strSQL = strSQL & "TRAINERS.SortOrder, "
					end if
					strSQL = strSQL & GetTrnOrderBy()

					rsEntry.CursorLocation = 3
					rsEntry.open strSQL, cnWS
					Set rsEntry.ActiveConnection = Nothing

					Do While Not rsEntry.EOF
%>
						<option value="<%=rsEntry("TrainerID")%>" title="<%=stringIfLengthOverLimit(Replace(FmtTrnNameNew(rsEntry, true),"&nbsp;", ""),24)%>" <% if curTrn=CLNG(rsEntry("TrainerID")) then Response.Write "selected" end if %>>
							<%=truncateString(Replace(FmtTrnNameNew(rsEntry, true),"&nbsp;", ""),24)%>
						</option>
<%
						rsEntry.MoveNext
					Loop
					rsEntry.close

					strSQL = "SELECT DISTINCT TrFirstName, TrLastName, CASE WHEN NOT TRAINERS.DisplayName IS NULL THEN TRAINERS.DisplayName ELSE TRAINERS.TrLastName END AS TrnDisName, tblTrainerSchedules.TrainerID FROM TRAINERS, tblTrainerSchedules WHERE TRAINERS.TrainerID=tblTrainerSchedules.TrainerID AND tblTrainerSchedules.MaskPublic=1 AND TRAINERS.[Delete]=0 AND TRAINERS.Active=1 AND TRAINERS.TrainerID<>-1 AND tblTrainerSchedules.EndDate >= " & DateSep & curTmpDate & DateSep & " ORDER BY TrLastName, TrFirstName"
					rsEntry.CursorLocation = 3
					rsEntry.open strSQL, cnWS
					Set rsEntry.ActiveConnection = Nothing

					Do While Not rsEntry.EOF
%>
                  <option value="<%=rsEntry("TrainerID")%>m"><%=xssStr(allHotWords(18))%>&nbsp;<%=trnCount%></option>
                  <%
                                                        rsEntry.MoveNext
                                                Loop
                                                rsEntry.close
%>
				</select>
				<script type="text/javascript">
					document.search2.optInstructor.options[0].text = "<%= allHotWords(149) %>" + " " + "<%= allHotWords(114) %>";
				</script>
<%
				if ss_ApptSchedGenderFilter then
%>
					<select name="optGender" id="optGender" class="bf-filter" onchange="subForm();">
						<option value="">
							<%= allHotWords(149) %>
						</option>
						<option value="Female" <% if request.Form("optGender") = "Female" then response.Write("selected") end if %>>
							<%=xssStr(allHotWords(139))%>
						</option>
						<option value="Male" <% if request.Form("optGender") = "Male" then response.Write("selected") end if %>>
							<%=xssStr(allHotWords(138))%>
						</option>
					</select>
<%
				end if
%>
				<input type="hidden" id="optView" name="optView" class="bf-filter" value="<%=ifb_view%>" />
				<!--<input type="hidden" name="optView" value="<%if request.Form("optView")="" AND ss_apptDefaultToWeekView then response.write "week" else response.write "day" end if%>">-->
			</div>
		</div>
	</div>
	</form>
	
<%
	If NOT Session("Pass") AND NOT UseVersionB(1) then

		'use lightbox login

		' hold dictionary of trainer IDs and their display names... this will get used when building
		' out the scheudle's booking strings
		dim TrainerNamesDict
		Set TrainerNamesDict = Server.CreateObject("Scripting.Dictionary")

		strSQL = " SELECT TrainerID, TrFirstName, TrLastName, DisplayName FROM Trainers WHERE Trainers.Active = 1"

		dim rsTrainerNames
		set rsTrainerNames = Server.CreateObject("ADODB.Recordset")
		rsTrainerNames.CursorLocation = 3
		rsTrainerNames.open strSQL, cnWS
		Set rsTrainerNames.ActiveConnection = Nothing

		do while NOT rsTrainerNames.EOF
			dim tmptrnid : tmptrnid = rsTrainerNames("TrainerID")
			'add trainer name (formatted by a function in inc_i18n.asp)
			TrainerNamesDict.add tmptrnid, FmtTrnNameNew(rsTrainerNames,false)
			rsTrainerNames.MoveNext
		loop

		rsTrainerNames.close
		set rsTrainerNames = nothing

		'end trainer names stuff

%>
		<script type="text/javascript" language="javascript">

			var trainersMap = {
<%

				dim x, arr_trnNames, arr_trnIDs
				arr_trnIDs = TrainerNamesDict.Keys
				For x = 0 to TrainerNamesDict.Count - 1
					if Left(CStr(arr_trnIDs(x)),1) = "-" then
						'precede keys that are negative numbers with "n", because "-" is invalid javascript. strip the "-" character
						arr_trnIDs(x) = Replace(CStr(arr_trnIDs(x)), "-", "n")
					end if

					Response.Write "" & arr_trnIDs(x) & ""  & ": '" & jsEscSingle(xssStr(TrainerNamesDict.Item(arr_trnIDs(x)))) & "'"

					if x <> UBound(arr_trnIDs) then
						Response.write "," & VbCrLf
					end if
				Next

%>
			};

			var typeGroupMap = {
<%
				dim xx
				For xx = 0 to UBound(tgIDs)
					if CStr(tgIDs(xx)) <> "" then

						if Left(CStr(tgIDs(xx)),1) = "-" then
							tgIDs(xx) = Replace(CStr(tgIDs(xx)),"-","n")
						end if

						if xx = 0 then
							Response.Write "" & tgIDs(xx) & "" & ": '" & jsEscSingle(xssStr(tgNames(xx))) & "'"
						else
							Response.Write "" & "," & tgIDs(xx) & "" & ": '" & jsEscSingle(xssStr(tgNames(xx))) & "'"
						end if
					end if
				Next
%>
			};

			$(document).ready(function () {
				var $links = $("a.apptLink");
				/*once doc is loaded, pull up all links that have class "apptLlink"*/

				//copied queryString parse code
				function getQueryStingValue(fullString, searchFor) {
					var parts = fullString.split("&");
					for ( var i=0; i < parts.length; i++ ) {
						var keyValPair = parts[i].split("=");
						if (keyValPair[0] == searchFor) {
							return keyValPair[1];
						}
					}
					return null;
				}

				$links.each(function () {
					var $this = $(this);

					//wrap up forwarding link and booking string into a closure that will become the link's clickhandler
					var clickHandler = (function(){
						var forwardingLink = "/ASP/"+ $this.attr("href");

						var typeGroupID, typeGroupName, theDate, theTrainerID, theTrainerName, bookingStr1, bookingStr2;

						typeGroupID = getQueryStingValue(forwardingLink, "tgid");
						typeGroupName = typeGroupMap[typeGroupID];
						bookingStr1 = typeGroupName;

						theDate = getQueryStingValue(forwardingLink, "localDate");

						theTrainerID = getQueryStingValue(forwardingLink, "trnid");

						theTrainerName = trainersMap[theTrainerID];

						bookingStr2 = "on " + theDate;

						if (typeof theTrainerName != "undefined" && theTrainerName != " ")
							bookingStr2 += ", with " + theTrainerName;

						//return closure function that has bookingStr1,2 and forwardingLink wrapped up...
						return function(evt){
							evt.preventDefault();

							if (forwardingLink == "undefined")
								throw "No forwarding link provided in <a> tag: " + $this.toString();

							promptLogin(bookingStr1, bookingStr2, forwardingLink);	
						}
					})();

					$this.click(clickHandler);

			});
		});
	</script>

	<!-- #include file="inc_login_content.asp" -->
	<!-- #include file="inc_fb_confirm_login_lb.asp" -->
<% end if %>




<script type="text/javascript">
	$(document).ready(function() {
		$('a.parentLink').click(function(event) {
			event.preventDefault();
			parent.location.href = $(this).attr('linkit');
		});
	});
</script>
<% pageEnd %>
<!-- #include file="post.asp" -->


