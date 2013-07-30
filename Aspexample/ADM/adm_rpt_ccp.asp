<%@ CodePage=65001 %>
<!-- #include file="init.asp" -->
<%
'Dim SessionFarm
'set SessionFarm = Server.CreateObject("SessionFarm.SFSession")
response.charset="utf-8"
%>
<%
dim phraseDictionary
set phraseDictionary = LoadPhrases("BusinessmodemerchantaccounttransactionsPage", 140)

%>
		<!-- #include file="inc_accpriv.asp" -->

<script type="text/javascript">
    function showConfirmVoidAllCheckedTransactions() {
        return '<%=DisplayPhraseJS(phraseDictionary,"Confirmvoidallcheckedtransactions")%>';
    }

    function showNoTransToBeSettled(){
        return '<%=DisplayPhraseJS(phraseDictionary,"Notransactionssettledtobebatched")%>';
    }

    function showAboutToSettleNumTrans(number){
        var AboutToSettle = '<%=DisplayPhraseJS(phraseDictionary,"Youareabouttosettlenumtransactions")%>';
        AboutToSettle = AboutToSettle.replace("<NUMBER>", number);
        return AboutToSettle;
    }
</script>
<%
	Dim ap_cc_void, ap_cc_settle, VoidOk
	ap_cc_void = validAccessPriv("CC_Void")
	ap_cc_settle = validAccessPriv("CC_Settle")

	VoidOk = validAccessPriv("TB_VOID")

if not Session("Pass") OR Session("Admin")="false" OR (NOT validAccessPriv("RPT_CCP") AND NOT ap_cc_settle) then 
%>
<script type="text/javascript">
	alert("<%=DisplayPhraseJS(systemMessagesErrorsDictionary,"Notauthorizedtoviewpage")%>");
	javascript:history.go(-1);
</script>
<%
else
%>
		<!-- #include file="../inc_i18n.asp" -->
		<!-- #include file="inc_rpt_tagging.asp" -->
		<!-- #include file="inc_hotword.asp" -->
<%
	if useVersionB(TEST_NAVIGATION_REDESIGN) then
		Session("tabID") = "97"
	end if

	function padZeros(val, NumDigits)
		padZeros = ""
		if val<>"" then
			for i=1 to NumDigits-Len(val)
				padZeros = padZeros & "0"
			next
			padZeros = padZeros & val
		end if
	end function

	Dim cSDate, cEDate, disMode, curLocName, SmodeCCP, cont, ss_EnableACH, oneMID, CCProcessor, tmpVisa, tmpMC, tmpAmex, tmpDisc, cSwiped
	SmodeCCP = checkStudioSetting("tblCCOpts", "SModeCCP")

	dim category : category = ""
	dim tmpCat : tmpCat = ""
	if (RQ("category"))<>"" then
		tmpCat = RQ("category")
		category = Replace(tmpCat, " ", "")
	elseif (RF("category"))<>"" then
		tmpCat = RF("category") 
		category = Replace(tmpCat, " ", "")
	end if

	set rsEntry = Server.CreateObject("ADODB.Recordset")

	CCProcessor =""
    if implementationSwitchIsEnabled("BluefinCanada") then
        CCProcessor2 =""
	    strSQL = "SELECT CCProcessor, CCProcessor2 FROM Studios WHERE StudioID=" & session("StudioID")
    else
        strSQL = "SELECT CCProcessor FROM tblCCOpts WHERE StudioID=" & session("StudioID")
    end if
	rsEntry.CursorLocation = 3
	rsEntry.open strSQL, cnWS
	Set rsEntry.ActiveConnection = Nothing
	if NOT rsEntry.EOF then
		if NOT isNULL(rsEntry("CCProcessor")) then
			CCProcessor = TRIM(rsEntry("CCProcessor"))
            if implementationSwitchIsEnabled("BluefinCanada") then
                if CCProcessor = "PMN" then
                    CCProcessor2 = rsEntry("CCProcessor2")
                end if
            end if            
		end if        
	end if
	rsEntry.close
	
	oneMID = false
	if CCProcessor<>"OP" then	''if Optimal Payments must select location to find config file
		strSQL = "SELECT MID, TID FROM Location GROUP BY MID, TID HAVING (NOT (MID IS NULL))"
		rsEntry.CursorLocation = 3
		rsEntry.open strSQL, cnWS
		Set rsEntry.ActiveConnection = Nothing
		if NOT rsEntry.EOF then
			if rsEntry.RecordCount = 1 then
				oneMID = true
			end if
		end if
		rsEntry.close
	end if
	
	if request.form("optCCLocation")<>"" then
		ccLoc = request.form("optCCLocation")
	else
		if NOT SmodeCCP then
			ccLoc = "98"		
		else
			ccLoc = "-1"
			if NOT oneMID then		
				ccLoc = session("curLocation")
			end if
		end if
	end if

	if request.form("requiredtxtDateStart")="" then	'First Load
		cSwiped = "true"
	end if
	if request.form("optSwiped")<>"" then
		cSwiped = request.form("optSwiped")
	end if
	
	if request.form("requiredtxtDateStart")<>"" then
		Call SetLocale(session("mvarLocaleStr"))
			cSDate = CDATE(request.form("requiredtxtDateStart"))
		Call SetLocale("en-us")
	else
		cSDate = DateAdd("m",-1,DateValue(CDATE(DateAdd("n", Session("tzOffset"),Now))))
	end if

	if request.form("requiredtxtDateEnd")<>"" then
		Call SetLocale(session("mvarLocaleStr"))
			cEDate = CDATE(request.form("requiredtxtDateEnd"))
		Call SetLocale("en-us")
	else
		cEDate = DateValue(CDATE(DateAdd("n", Session("tzOffset"),Now)))
	end if

	if request.form("optDate")="all" then
		disMode = "all"
	else
		disMode = "range"
	end if
	
	if ccLoc<>"-1" then
		strSQL = "SELECT LocationName FROM Location WHERE LocationID=" & ccLoc
		rsEntry.CursorLocation = 3
		rsEntry.open strSQL, cnWS
		Set rsEntry.ActiveConnection = Nothing
		if NOT rsEntry.EOF then
			curLocName = rsEntry("LocationName")				
		end if
		rsEntry.close
	else
		curLocName = xssStr(allHotWords(479))
	end if
	
	if NOT request.form("frmExpReport")="true" then
%>
<!-- #include file="pre.asp" -->
<!-- #include file="frame_bottom.asp" -->
<%= js(array("mb", "adm/adm_rpt_ccp", "MBS", "reportFavorites", "plugins/jquery.SimpleLightBox")) %>
<!-- American/Canada=2 format mm/dd/yyyy --> <!-- European/Rest of the world=1 format dd-mm-yyyy -->

<%= js(array("calendar" & dateFormatCode)) %>
<%= css(array("SimpleLightBox")) %>
<!-- #include file="inc_help_content.asp" -->
<!-- #include file="../inc_date_ctrl.asp" -->
<!-- #include file="../inc_ajax.asp" -->
<!-- #include file="../inc_val_date.asp" -->
<!-- #include file="css/site_setup.asp" -->
<!-- #include file="inc_user_options.asp" -->
<% pageStart %>
	<% if UseVersionB(TEST_REPORT_LANDING_PAGE) then %>
		<div class="headText breadcrumbs-old" valign="bottom">
		<span class="breadcrumb-item"><a href="/reportslandingpage/index"><%=DisplayPhrase(reportPageTitlesDictionary, "Reports")%></a></span>
		<%if category <>"" then%>
		<span class="breadcrumb-item">&raquo;</span>
		<span class="breadcrumb-item active"><a href="/reportslandingpage/<%=category%>category"><%=GetReportsCategoryPagePhrase(category)%></a></span>
		<%end if %>
		<span class="breadcrumb-item">&raquo;</span>
		<%=DisplayPhrase(reportPageTitlesDictionary,"Approvedtransactions") %>

		<div id="add-to-favorites">
			<span id="" class="favorite"></span><span class="text"><%= DisplayPhrase(reportPageTitlesDictionary,"Addtofavorites") %></span>
		</div>
		</div>
	<%end if %>
<table height="100%" width="<%=strPageWidth%>" cellspacing="0">
    <tr> 
     <td valign="top" height="100%" width="100%">
        <form name="frmCCP" method="POST" action="adm_rpt_ccp.asp">
		  <input type="hidden" name="frmGenReport" value="" />
		  <input type="hidden" name="frmExpReport" value="" />
		  <% if UseVersionB(TEST_REPORT_LANDING_PAGE) then %>
				<input type="hidden" name="reportUrl" id="reportUrl" value="<%=Request.ServerVariables("SCRIPT_NAME") %>" />
				<input type="hidden" name="category" value="<%=category%>">
		  <% end if %>
				<div id="topdiv">
					<table class="center" cellspacing="0" width="90%" height="100%">
			<% if NOT UseVersionB(TEST_REPORT_LANDING_PAGE) then %>
            <tr> 
              <td class="headText" align="left" valign="top"> 
                <table class="mainText" width="100%" cellspacing="0">
                  <tr> 
                    <td class="headText" valign="bottom"><b><%=DisplayPhrase(pageTitlesDictionary,"Batchreview") %></b>
						<!--JM - 48_2448-->
						<% showTrainingMovieIcon("21044446-managing-credit-card-processing#batch") %>
					</td>
                  </tr>
                </table>
              </td>
            </tr>
<%		end if
         dim rsEntry
         set rsEntry = Server.CreateObject("ADODB.Recordset")
%>
            <tr> 
              <td valign="bottom" class="mainText right" height="18"> 

			<table class="mainText" width="100%"  cellspacing="0">
			  <tr>
				<td valign="top" class="right" colspan="2">
				<!-- #include file="inc_batch_nav.asp" -->
				</td>
			  </tr>
			</table>
			
			  </td>
            </tr>
						<tr>
			<td>
			<table class="mainText border4" cellspacing="0" style="float:right;">
                        <tr> 
                          <td valign="bottom" style="background-color:#F2F2F2;" nowrap>
                          <b>
<%
			dim LastBatch, ActiveBatch
			LastBatch = null
			strSQL = "SELECT tblCCOpts.EnableACH, LastBatch, ActiveBatch FROM tblCCOpts WHERE tblCCOpts.StudioID=" & session("StudioID")
			rsEntry.CursorLocation = 3
			rsEntry.open strSQL, cnWS
			Set rsEntry.ActiveConnection = Nothing
			if NOT rsEntry.EOF then
				ss_EnableACH = rsEntry("EnableACH")
				if (NOT isNull(rsEntry("LastBatch"))) then
					LastBatch = DateAdd("n", Session("tzOffset"),rsEntry("LastBatch"))
				end if
				ActiveBatch = rsEntry("ActiveBatch")
			end if
			rsEntry.close

			pMID = ""
			if ccLoc<>"-1" then
				strSQL = "SELECT MID, TID FROM Location WHERE LocationID=" & ccLoc
				rsEntry.CursorLocation = 3
				rsEntry.open strSQL, cnWS
				Set rsEntry.ActiveConnection = Nothing
				if NOT rsEntry.EOF then
					pMID = rsEntry("MID")
				end if
				rsEntry.close
			end if
			
			'only add options or enabled card types
			strSQL = "SELECT ccVisa, ccMasterCard, ccAmericanExpress, ccDiscover "
			strSQL = strSQL & "FROM tblCCOpts"
			rsEntry.CursorLocation = 3
			rsEntry.open strSQL, cnWS
			Set rsEntry.ActiveConnection = Nothing
			if NOT rsEntry.EOF then	
				tmpAmex = rsEntry("ccAmericanExpress")
				tmpVisa = rsEntry("ccVisa")
				tmpMC = rsEntry("ccMasterCard")
				tmpDisc = rsEntry("ccDiscover")
				rsEntry.close
%>
							  <select name="optCCType">
<%
				if tmpAmex then
%>
                                  <option value="Credit (AMEX)" <%if request.form("optCCType")="Credit (AMEX)" then response.write "selected" end if%> ><%=xssStr(allHotWords(659))%></option>
<%
				end if
				if tmpVisa or tmpMC then
%>
                                  <option value="Credit (Visa/MC)" <%if request.form("optCCType")="Credit (Visa/MC)" then response.write "selected" end if%>><%=xssStr(allHotWords(660))%></option>
<%
				end if
				if tmpDisc then
%>
								  <option value="Credit (Discover)" <%if request.form("optCCType")="Credit (Discover)" then response.write "selected" end if%>><%=xssStr(allHotWords(661))%></option>
<%
				end if
%>
                              </select>

<%
			end if			
%>
							  <select name="optSwiped">
							  	<option value="true" <%if cSwiped="true" then response.write "selected" end if%>><%=xssStr(allHotWords(662))%></option>
							  	<option value="false" <%if cSwiped="false" then response.write "selected" end if%>><%=xssStr(allHotWords(663))%></option>
							  	<option value="" <%if cSwiped="" then response.write "selected" end if%>><%=xssStr(allHotWords(149))%></option>
							  </select>
							  <select name="optCCLocation">
<%
				strSQL = "SELECT LocationID, LocationName FROM Location WHERE Active=1 AND wsShow=1 AND LocationID<>98 ORDER BY LocationName"
				rsEntry.CursorLocation = 3
				rsEntry.open strSQL, cnWS
				Set rsEntry.ActiveConnection = Nothing

				do while NOT rsEntry.EOF
%>
                                  <option value="<%=rsEntry("LocationID")%>" <%if ccLoc=CSTR(rsEntry("LocationID")) then response.write "selected" end if%>><%=rsEntry("LocationName")%></option>
<%
					rsEntry.MoveNext
				loop
				rsEntry.close
%>
							<option value="98" <%if ccLoc="98" then response.write "selected" end if%>><%=xssStr(allHotWords(25))%></option>
							<option value="-1" <%if ccLoc="-1" then response.write "selected" end if%>><%=xssStr(allHotWords(479))%></option>
							</select>
                            <%=xssStr(allHotWords(77))%> 
                            <input onBlur="validateDate(this, '<%=FmtDateShort(cSDate)%>', true);" type="text"  name="requiredtxtDateStart" value="<%=FmtDateShort(cSDate)%>" class="date">
                        <script type="text/javascript">
                        	var cal1 = new tcal({ 'formname': 'frmCCP', 'controlname': 'requiredtxtDateStart' });
                        	cal1.a_tpl.yearscroll = true;
		</script>
                            &nbsp;<%=xssStr(allHotWords(79))%> 
                            <input onBlur="validateDate(this, '<%=FmtDateShort(cEDate)%>', true);" type="text"  name="requiredtxtDateEnd" value="<%=FmtDateShort(cEDate)%>" class="date">
                        <script type="text/javascript">
                        	var cal2 = new tcal({ 'formname': 'frmCCP', 'controlname': 'requiredtxtDateEnd' });
                        	cal2.a_tpl.yearscroll = true;
		</script>
                            &nbsp;
                            <% if validAccessPriv("RPT_TAG") then 
                                taggingFilter 
                            end if %>
							<input name="Button" type="button" value="<%=xssStr(allHotWords(226))%>" onClick="genCCP();">
							<span class="icon-button" style="vertical-align: middle;" title="<%=xssStr(allHotWords(658))%>" ><a onClick="exportCCP();" ><img src="<%= contentUrl("/asp/adm/images/export-to-excel-20px.png") %>" /></a></span>
                            <% if NOT validAccessPriv("RPT_TAG") then 
                            else 
                                taggingButtons("frmCCP") 
                            end if %>
                            </b>
							</td>
                        </tr>
                      </table>
			</td>
			</tr>
						</table>
			</div>
			</td>
			</tr>
			
            <tr>
              <td valign="top" class="mainTextBig center" align="left">
<%
	if ccLoc="" then
	else
%>
              <a style="color:black;"><b><%=DisplayPhrase(phraseDictionary,"Quickreference")%></b></a> 
              <table class="smallTextBlack center" width="400"  border="1" cellspacing="0">
                <tr>
                  <td>&nbsp;<%if session("numLocations")>1 then response.write UCASE(curLocName) else response.write xssStr(allHotWords(462)) end if%></td>
<%
		if tmpAmex then
%>
                  <td class="center-ch"><a onclick="allCurLoc();amexCurLoc();genCCP()" ><%=xssStr(allHotWords(673))%></a></td>
<%
		end if
		if tmpVisa or tmpMC then
%>
                  <td class="center-ch"><a onclick="allCurLoc();visaMcCurLoc();genCCP()" ><%=xssStr(allHotWords(660))%></a></td>
<%
		end if
		if tmpDisc then
%>
                  <td class="center-ch"><a onclick="allCurLoc();discoCurLoc();genCCP()" ><%=xssStr(allHotWords(661))%></a></td>
<%
		end if
%>
                </tr>
<%
			if session("mvarMIDs")>0 then
				'QUICK REFERENCE KEYED
				'''Get Stats for location
				'strSQL = "SELECT LEFT(RTRIM(LTRIM(ccNum)), 1) AS CardType, COUNT(LEFT(ccNum, 1)) AS NumTrans, SUM(ccAmt) AS Amt FROM tblCCTrans WHERE (Settled = 0) AND (Status = N'Approved') AND (TransTime >= " & DateSep & cSDate & DateSep & ") AND (TransTime <= " & DateSep & DateAdd("d", 1, cEDate) & DateSep & ") "
				strSQL = "SELECT CASE WHEN ccType=N'American Express' THEN 3 WHEN ccType=N'Visa' THEN 4 WHEN ccType=N'Master Card' THEN 5 WHEN ccType=N'Discover' THEN 6 END AS CardType, COUNT(ccType) AS NumTrans, SUM(CASE WHEN tblCCTrans.Status = 'Credit' THEN (tblCCTrans.ccAmt * -1) ELSE (tblCCtrans.ccAmt) END) AS Amt FROM tblCCTrans WHERE (Settled = 0) AND (Status = N'Approved'"
				if ccProcessor="TCI" then
					strSQL = strSQL & " OR tblCCTrans.Status = N'Credit'"
				end if
				strSQL = strSQL & ") AND (TransTime >= " & DateSep & cSDate & DateSep & ") AND (TransTime <= " & DateSep & DateAdd("d", 1, cEDate) & DateSep & ") "
				if ccLoc<>"-1" then
					strSQL = strSQL & "AND (tblCCTrans.LocationID = " & ccLoc & ") "
				end if
				strSQL = strSQL & "AND (tblCCTrans.CCSwiped=0) "
				strSQL = strSQL & " AND (NOT (ccType IS NULL)) GROUP BY CASE WHEN ccType=N'American Express' THEN 3 WHEN ccType=N'Visa' THEN 4 WHEN ccType=N'Master Card' THEN 5 WHEN ccType=N'Discover' THEN 6 END ORDER BY CASE WHEN ccType=N'American Express' THEN 3 WHEN ccType=N'Visa' THEN 4 WHEN ccType=N'Master Card' THEN 5 WHEN ccType=N'Discover' THEN 6 END"
			response.write debugSQL(strSQL, "SQL")
				rsEntry.CursorLocation = 3
				rsEntry.open strSQL, cnWS
				Set rsEntry.ActiveConnection = Nothing
%>
                <tr>
                  <td>&nbsp;&nbsp;&nbsp;<%=xssStr(allHotWords(664))%></td>
                  <%
				if tmpAmex then
					'Get AMEX
					tmpTransAmt = 0
					tmpTransCount = 0
					if NOT rsEntry.EOF then
						if rsEntry("CardType")=3 then
							tmpTransCount = rsEntry("NumTrans")
							tmpTransAmt = rsEntry("Amt")
							rsEntry.MoveNext
						end if
					end if
%>				  
                  <td class="center-ch" style="cursor:pointer"><a href="javascript:keyedCurLoc();amexCurLoc();genCCP();" ><%=tmpTransCount%> - <%=FormatNumber(tmpTransAmt*.01,2)%></a></td>
<%
				end if
				
				if tmpVisa or tmpMC then
					'Get Visa/MC
					tmpTransCount = 0
					tmpTransAmt = 0
					if NOT rsEntry.EOF then
						if rsEntry("CardType")=4 then
							tmpTransCount = rsEntry("NumTrans")
							tmpTransAmt = rsEntry("Amt")
							rsEntry.MoveNext
						end if
					end if
					if NOT rsEntry.EOF then
						if rsEntry("CardType")=5 then
							tmpTransCount = tmpTransCount + rsEntry("NumTrans")
							tmpTransAmt = tmpTransAmt + rsEntry("Amt")
							rsEntry.MoveNext
						end if
					end if
%>				  
                  <td class="center-ch" style="cursor:pointer"><a href="javascript:keyedCurLoc();visaMcCurLoc();genCCP()" ><%=tmpTransCount%> - <%=FormatNumber(tmpTransAmt*.01,2)%></a></td>
<%
				end if
				
				if tmpDisc then
					'Discover
					tmpTransCount = 0
					tmpTransAmt = 0
					if NOT rsEntry.EOF then
						if rsEntry("CardType")=6 then
							tmpTransCount = rsEntry("NumTrans")
							tmpTransAmt = rsEntry("Amt")
						end if
					end if
			
%>				  
                  <td class="center-ch" style="cursor:pointer"><a href="javascript:keyedCurLoc();discoCurLoc();genCCP()" ><%=tmpTransCount%> - <%=FormatNumber(tmpTransAmt*.01,2)%></a></td>
                </tr>	
<%
				end if 'if tmpDisc
				rsEntry.close

				'QUICK REFERENCE Swiped
				'''Get Stats for location
				'strSQL = "SELECT LEFT(RTRIM(LTRIM(ccNum)), 1) AS CardType, COUNT(LEFT(ccNum, 1)) AS NumTrans, SUM(ccAmt) AS Amt FROM tblCCTrans WHERE (Settled = 0) AND (Status = N'Approved') AND (TransTime >= " & DateSep & cSDate & DateSep & ") AND (TransTime <= " & DateSep & DateAdd("d", 1, cEDate) & DateSep & ") "
				strSQL = "SELECT CASE WHEN ccType=N'American Express' THEN 3 WHEN ccType=N'Visa' THEN 4 WHEN ccType=N'Master Card' THEN 5 WHEN ccType=N'Discover' THEN 6 END AS CardType, COUNT(ccType) AS NumTrans, SUM(CASE WHEN tblCCTrans.Status = 'Credit' THEN (tblCCTrans.ccAmt * -1) ELSE (tblCCtrans.ccAmt) END) AS Amt FROM tblCCTrans WHERE (Settled = 0) AND (Status = N'Approved' "
				if ccProcessor="TCI" then
					strSQL = strSQL & " OR tblCCTrans.Status = N'Credit'"
				end if
				strSQL = strSQL & ")  AND (TransTime >= " & DateSep & cSDate & DateSep & ") AND (TransTime <= " & DateSep & DateAdd("d", 1, cEDate) & DateSep & ") "
				if ccLoc<>"-1" then
					strSQL = strSQL & "AND (tblCCTrans.LocationID = " & ccLoc & ") "
				end if
				strSQL = strSQL & "AND (tblCCTrans.CCSwiped=1) "
				strSQL = strSQL & " AND (NOT (ccType IS NULL)) GROUP BY CASE WHEN ccType=N'American Express' THEN 3 WHEN ccType=N'Visa' THEN 4 WHEN ccType=N'Master Card' THEN 5 WHEN ccType=N'Discover' THEN 6 END ORDER BY CASE WHEN ccType=N'American Express' THEN 3 WHEN ccType=N'Visa' THEN 4 WHEN ccType=N'Master Card' THEN 5 WHEN ccType=N'Discover' THEN 6 END"
			response.write debugSQL(strSQL, "SQL")
				rsEntry.CursorLocation = 3
				rsEntry.open strSQL, cnWS
				Set rsEntry.ActiveConnection = Nothing
%>
                <tr>
                  <td>&nbsp;&nbsp;&nbsp;<%=xssStr(allHotWords(665))%></td>
<%
				if tmpAmex then
					'Get AMEX
					tmpTransAmt = 0
					tmpTransCount = 0
					if NOT rsEntry.EOF then
						if rsEntry("CardType")=3 then
							tmpTransCount = rsEntry("NumTrans")
							tmpTransAmt = rsEntry("Amt")
							rsEntry.MoveNext
						end if
					end if
%>				                       
                  <td class="center-ch" style="cursor:pointer"><a href="javascript:swipedCurLoc();amexCurLoc();genCCP()" ><%=tmpTransCount%> - <%=FormatNumber(tmpTransAmt*.01,2)%></a></td>
<%
				end if
				
				if tmpVisa or tmpMC then
					'Get Visa/MC
					tmpTransCount = 0
					tmpTransAmt = 0
					if NOT rsEntry.EOF then
						if rsEntry("CardType")=4 then
							tmpTransCount = rsEntry("NumTrans")
							tmpTransAmt = rsEntry("Amt")
							rsEntry.MoveNext
						end if
					end if
					if NOT rsEntry.EOF then
						if rsEntry("CardType")=5 then
							tmpTransCount = tmpTransCount + rsEntry("NumTrans")
							tmpTransAmt = tmpTransAmt + rsEntry("Amt")
							rsEntry.MoveNext
						end if
					end if
%>				  
                  <td class="center-ch" style="cursor:pointer"><a href="javascript:swipedCurLoc();visaMcCurLoc();genCCP()" ><%=tmpTransCount%> - <%=FormatNumber(tmpTransAmt*.01,2)%></a></td>
<%
				end if
				
				if tmpDisc then
					'Discover
					tmpTransCount = 0
					tmpTransAmt = 0
					if NOT rsEntry.EOF then
						if rsEntry("CardType")=6 then
							tmpTransCount = rsEntry("NumTrans")
							tmpTransAmt = rsEntry("Amt")
						end if
					end if
			
%>				  
                  <td class="center-ch" style="cursor:pointer"><a href="javascript:swipedCurLoc();discoCurLoc();genCCP()" ><%=tmpTransCount%> - <%=FormatNumber(tmpTransAmt*.01,2)%></a></td>
                </tr>	
<%
				end if 'if tmpDisc
				rsEntry.close

		end if
%>
              </table>
              <% end if ''not multi loc w/ all locs selected %>			
			  </td>
			</tr> 
		</table>
		</form>
		<form name="frmCCSettle" id="frmCCSettle" action="adm_rpt_ccp_p.asp" method="post">
			<input type="hidden" name="frmMID" value="<%=pMID%>" />
			<input type="hidden" name="frmBatchLoc" value="<%=ccLoc%>" />  
			<% if UseVersionB(TEST_REPORT_LANDING_PAGE) then %>
				<input type="hidden" name="reportUrl" id="Hidden1" value="<%=Request.ServerVariables("SCRIPT_NAME") %>" />
				<input type="hidden" name="category" value="<%=category%>">
			<% end if %>

		<table class="center" style="width:95%">      
            <tr> 
            <td valign="top" class="mainTextBig"> 
              <table class="mainText" width="95%" cellspacing="0" style="margin:auto;">
                <tr > 
                  <td  colspan="2" valign="top" class="mainTextBig center"> 
					
					  
                      <table class="mainText" cellspacing="0" width="90%">
                       <tr> 
                          <td valign="top" align=right> 
                            <table class="mainText" width="100%" cellspacing="0">
                              <tr> 
                              <td><b></b></td>
                              <td class="right" valign="bottom"> 
		<% if request.form("frmGenReport")="true" then %>							  
							  <a id="checkAll" href="javascript:checkAll(document.getElementById('frmCCSettle'), 'filecheck', true);"><%=xssStr(allHotWords(617))%></a> | <a id="uncheckAll" href="javascript:checkAll(document.getElementById('frmCCSettle'), 'filecheck', false);"><%=xssStr(allHotWords(618))%></a>
		<% end if %>&nbsp;
							   </td>
                              </tr>
							</table>
						</td>
						</tr>
					</table>
<%			else ' we are exporting %>
<%
				Dim stFilename
				stFilename="attachment; filename=ApprovedTransactions " & Replace(cSDate,"/","-") & " - " & Replace(cEDate,"/","-") & ".xls" 
				Response.ContentType = "application/vnd.ms-excel" 
				Response.AddHeader "Content-Disposition", stFilename
			end if

%>
              <table class="smallTextBlack"  border="1" cellspacing="0" style="width:100%;">
<%			'end if 'export check %>
                    
                        <tr> 
                          <td valign="top" align=right> 
                            <table class="mainText" cellspacing="0" width="100%">
                              <tr> 
                                <td colspan=11 style="background-color:<%=session("pageColor2")%>;"><img height="2" width="100%" src="<%= contentUrl("/asp/adm/images/trans.gif") %>"></td>
                              </tr>
                              <tr> 
                                <th class="whiteHeader" nowrap bgcolor="<%= session("pageColor4")%>"> 
                                  <b>&nbsp;<%=xssStr(allHotWords(57))%> / <%=xssStr(allHotWords(58))%>
</b></th>
                                <th class="whiteHeader" nowrap bgcolor="<%= session("pageColor4")%>"><b>&nbsp;<%=xssStr(allHotWords(12))%>
</b></th>
                                <th class="whiteHeader" nowrap bgcolor="<%= session("pageColor4")%>"><b>&nbsp;<%=DisplayPhrase(phraseDictionary,"Cardinfo")%></b></th>
                                <th  nowrap class="whiteHeader center-ch" bgcolor="<%= session("pageColor4")%>"><b>&nbsp;<%=xssStr(allHotWords(35))%>&nbsp;</b></th>
                                <th  nowrap class="whiteHeader center-ch" bgcolor="<%= session("pageColor4")%>"><b>&nbsp;<%=xssStr(allHotWords(50))%>&nbsp;</b></th>
                                <th  nowrap class="whiteHeader center-ch" bgcolor="<%= session("pageColor4")%>"><b>&nbsp;<%=xssStr(allHotWords(424))%>&nbsp;</b></th>
                                <th  nowrap class="whiteHeader center-ch" bgcolor="<%= session("pageColor4")%>"><b>&nbsp;<%=xssStr(allHotWords(666))%></b></th>
                                <th  nowrap class="whiteHeader center-ch" bgcolor="<%= session("pageColor4")%>"><b>&nbsp;<%=xssStr(allHotWords(667))%></b></th>
						<% if session("numLocations")>1 then %>
                                <th  nowrap class="whiteHeader center-ch" bgcolor="<%= session("pageColor4")%>"><b>&nbsp;<%=xssStr(allHotWords(8))%></b></th>
						<% end if %>
                                <th  nowrap class="whiteHeader center-ch" bgcolor="<%= session("pageColor4")%>"><b><%=xssStr(allHotWords(60))%></b></th>
						<% if NOT request.form("frmExpReport")="true" then  %>
                                <th class="whiteHeader" nowrap bgcolor="<%= session("pageColor4")%>" class="center-ch"><b><%=xssStr(allHotWords(668))%>/<%=xssStr(allHotWords(669))%>&nbsp;</b></th>
						<% 	end if %>
                              </tr>
						<%	if NOT request.form("frmExpReport")="true" then  %>
                              <tr> 
                                <td colspan=11 style="background-color:<%=session("pageColor2")%>;"><img height="2" width="100%" src="<%= contentUrl("/asp/adm/images/trans.gif") %>"></td>
                              </tr>
						<% 	end if %>
<%
if request.form("frmGenReport")="true" then
		Dim strTempName, intCount
		strSQL = "SELECT tblCCTrans.TransactionNumber, tblCCTrans.CCLastFour, tblCCTrans.ExpMoYr, tblCCTrans.Status, tblCCTrans.Settled, tblCCTrans.ccAmt, tblCCTrans.BatchNumber, tblCCTrans.ClientID, tblCCTrans.AuthCode, tblCCTrans.TransTime, tblCCTrans.SaleID, tblCCTrans.ccNum, tblCCTrans.ccType, tblCCTrans.OrderID, CLIENTS.LastName, CLIENTS.FirstName, tblCCTrans.LocationID, Location.LocationName, tblCCTrans.Cardholder "
		strSQL = strSQL & " FROM Location RIGHT OUTER JOIN CLIENTS INNER JOIN tblCCTrans ON CLIENTS.ClientID = tblCCTrans.ClientID ON Location.LocationID = tblCCTrans.LocationID LEFT OUTER JOIN Sales ON tblCCTrans.SaleID = Sales.SaleID "
		
		if request.form("optFilterTagged")="on" then
			strSQL = strSQL & " INNER JOIN tblClientTag ON (tblClientTag.ClientID = CLIENTS.ClientID "
			if session("mVarUserID")<>"" then
				strSQL = strSQL & " AND tblClientTag.smodeID = " & session("mVarUserID")
			end if
			strSQL = strSQL & " ) "
		end if
		
		strSQL = strSQL & "WHERE ((tblCCTrans.Settled = 0) AND (tblCCTrans.Status = 'Approved'" 
		if ccProcessor="TCI" then
			strSQL = strSQL & " OR tblCCTrans.Status = 'Credit'"
		end if
		strSQL = strSQL & ") "
		if ccLoc<>"-1" then
			strSQL = strSQL & " AND tblCCTrans.LocationID=" & ccLoc
		end if
		if cSwiped="true" then
			strSQL = strSQL & "AND (tblCCTrans.CCSwiped=1) "
		elseif cSwiped="false" then
			strSQL = strSQL & "AND (tblCCTrans.CCSwiped=0) "
		end if
		if disMode = "range" then
			strSQL = strSQL & " AND tblCCTrans.TransTime >= " & DateSep & cSDate & DateSep & " "
			strSQL = strSQL & " AND tblCCTrans.TransTime <= " & DateSep & DateAdd("d", 1, cEDate) & DateSep & " "
		end if
		if request.form("optCCType")<>"" AND request.form("optCCType")<>"-1" then
			if request.form("optCCType") = "Credit (Visa/MC)" then
				'strSQL = strSQL & "AND (tblCCTrans.ccNum LIKE N'4%' OR tblCCTrans.ccNum LIKE N'5%') "
				strSQL = strSQL & "AND (tblCCTrans.ccType=N'Visa' OR tblCCTrans.ccType=N'Master Card') "
			elseif request.form("optCCType") = "Credit (AMEX)" then
				'strSQL = strSQL & "AND (tblCCTrans.ccNum LIKE N'3%') "
				strSQL = strSQL & "AND (tblCCTrans.ccType=N'American Express') "
			elseif request.form("optCCType") = "Credit (Discover)" then
				'strSQL = strSQL & "AND (tblCCTrans.ccNum LIKE N'6%') "
				strSQL = strSQL & "AND (tblCCTrans.ccType=N'Discover') "
			end if
		end if
    strSQL = strSQL & ") "

    	if request.form("frmTagClients")="true" then 'tag clients sql
            if request.form("frmTagClientsNew")="true" then
                clearAndTagQuery(strSQL)
            else
                tagQuery(strSQL)
            end if
        end if

		strSQL = strSQL & " ORDER BY tblCCTrans.TransTime DESC;"
response.write debugSQL(strSQL, "SQL1")
        rsEntry.CursorLocation = 3
		rsEntry.open strSQL, cnWS
		Set rsEntry.ActiveConnection = Nothing


		rowcount = 0 
		totalNumTransactions = 0
		totalAmount = 0 
	if not rsEntry.EOF then
		Do While NOT rsEntry.EOF
			if rowcount=0 then
%>
                              <tr style="background-color:#F2F2F2;">
<%
               rowcount = 1
            else
%>
                              <tr style="background-color:#FAFAFA;"> 
<%
               rowcount = 0
            end if
%>
                                <td nowrap >&nbsp;<%=FmtDateTime(DateAdd("n", Session("tzOffset"),rsEntry("TransTime")))%></td>
                                <td nowrap ><% if NOT request.form("frmExpReport")="true" then  %>&nbsp;<a href="adm_clt_purch.asp?ID=<%=rsEntry("ClientID")%>&qParam=ph" title="<%=DisplayPhraseAttr(phraseDictionary,"Clicktoviewclientaccounthistory")%>"><% end if %><%=rsEntry("FirstName")%>&nbsp;<%=rsEntry("LastName")%><% if NOT request.form("frmExpReport")="true" then  %></a><% end if %></td>

                                <td nowrap align="left">&nbsp;
								
								<%if NOT isNULL(rsEntry("Cardholder")) then response.write rsEntry("Cardholder") & " / " end if%><%=xssStr(allHotWords(117))%>:<%=rsEntry("ExpMoYr")%>
								<%if NOT isNULL(rsEntry("CCLastFour")) then response.write " / "& xssStr(allHotWords(670)) & ": " & padZeros(rsEntry("CCLastFour"),4)%>
								</td>
								<td nowrap class="center-ch" ><% if NOT request.form("frmExpReport")="true" then  %>&nbsp;<% end if %><%if rsEntry("Status")="Credit" then response.write "<span style=""color:#990000;"">" end if%><%=FormatNumber(rsEntry("ccAmt")*.01,2)%></td>
                                <td nowrap class="center-ch" >&nbsp;
<%
						'if Left(rsEntry("ccNum"),1)=4 OR Left(rsEntry("ccNum"),1)=5  then
						'	response.write "Visa/MC"
						'elseif Left(rsEntry("ccNum"),1)=3 then
						'	response.write "AMEX"
						'elseif Left(rsEntry("ccNum"),1)=6 then
						'	response.write "Discover"
						'else
						'	response.write "n/a"
						'end if
						
						response.write rsEntry("ccType")
%>								
								</td>
                                <td nowrap class="center-ch">&nbsp;
<%
						if NOT isNull(rsEntry("SaleID")) then
							if VoidOk and NOT request.form("frmExpReport")="true" then
								response.write "<a title=""" & DisplayPhraseAttr(phraseDictionary,"Clicktogototransaction") & """ href=""adm_tlbx_voidedit.asp?saleno=" & rsEntry("SaleID") & """>" & xssStr(allHotWords(159)) & "/" & xssStr(allHotWords(379)) & "</a>"
							else
								response.write xssStr(allHotWords(159)) & "/" & xssStr(allHotWords(379))
							end if
						else
							response.write xssStr(allHotWords(246))
						end if
%>
								&nbsp;</td>
                                <td nowrap class="center-ch" >&nbsp;<%=rsEntry("authCode")%>
									<%if NOT isNull(rsEntry("OrderID")) then%>
										| <%=rsEntry("OrderID")%>
									<%end if%>
								</td>
                                <td nowrap class="center-ch" ><%=rsEntry("TransactionNumber")%></td>
							<% if session("numLocations")>1 then %>
                                <td nowrap class="center-ch" ><%=rsEntry("LocationName")%></td>
							<% end if %>
                                <td nowrap class="center-ch" >&nbsp;&nbsp;<%if rsEntry("Status")="Credit" then response.write "<span style=""color:#990000;"">" end if%><%=rsEntry("Status")%></td>
							<%	if NOT request.form("frmExpReport")="true" then  %>
                                <td class="center-ch" > 
                                  <input type="checkbox" name="chk_<%=rsEntry("TransactionNumber")%>"  class="filecheck">
                                </td>
							<% 	end if %>
                              </tr>
						<%	if NOT request.form("frmExpReport")="true" then  %>
                              <tr> 
                                <td colspan=11 style="background-color:<%=session("pageColor4")%>;"><img height="1" width="100%" src="<%= contentUrl("/asp/adm/images/trans.gif") %>"></td>
                              </tr>
						<% 	end if %>
<%
			totalNumTransactions = totalNumTransactions + 1
			if rsEntry("Status")="Credit" then
				totalAmount = totalAmount - rsEntry("ccAmt")
			else
				totalAmount = totalAmount + rsEntry("ccAmt")
			end if
			
			
	    rsEntry.MoveNext 
    
        Loop
%>
						<%	if NOT request.form("frmExpReport")="true" then  %>
                              <tr style="background-color:<%=session("pageColor4")%>;"> 
                                <td colspan=11><img height="1" width="100%" src="<%= contentUrl("/asp/adm/images/trans.gif") %>"></td>
                              </tr>
						<%	end if %>
                              <tr> 
                                <td colspan=11>&nbsp;<b><%=DisplayPhrase(phraseDictionary,"Totalnumberoftransactions")%>: 
                                  <%=totalNumTransactions%>&nbsp;&nbsp;&nbsp;&nbsp;<%=DisplayPhrase(phraseDictionary,"Totalamount")%>: <%=FormatNumber(totalAmount/100,2)%></b></td>
                              </tr>
<%		
	else '''rsEntry.EOF
%>
                              <tr> 
                                <td colspan="11">&nbsp;<%=DisplayPhrase(phraseDictionary,"Nopendingauthorizedtransactions")%></td>
                              </tr>
							<%	if NOT request.form("frmExpReport")="true" then  %>
                              <tr> 
                                <td colspan=11 style="background-color:<%=session("pageColor4")%>;"><img height="1" width="100%" src="<%= contentUrl("/asp/adm/images/trans.gif") %>"></td>
                              </tr>
							<% end if %>
<%	
    end if
    rsEntry.Close
	
else	'No Show Reulst
%>
                              <tr> 
                                <td class="center-ch" colspan="11">&nbsp;<span style="color:#990000;">-- <%=DisplayPhrase(phraseDictionary,"Selectoptionsandclickgenerate")%> --</span></td>
                              </tr>
							 <%	if NOT request.form("frmExpReport")="true" then  %>
                              <tr> 
                                <td colspan=11 style="background-color:<%=session("pageColor4")%>;"><img height="1" width="100%" src="<%= contentUrl("/asp/adm/images/trans.gif") %>"></td>
                              </tr>
							  <% end if %>
<%
end if
%>
                            </table>
                          </td>
                        </tr>
				</table>
<%			if NOT request.form("frmExpReport")="true" then
				if request.form("frmGenReport")="true" then 
%>
                      <table class="mainText" width="90%" cellspacing="0">
                        <tr>
                            <% if implementationSwitchIsEnabled("BluefinCanada") then
                                   if CCProcessor2 <> "ELV" then %> 
			                           <td class="right"><%if ap_cc_void then%><a href="javascript:rejectTrans();"><%=DisplayPhrase(phraseDictionary,"Voidcheckedtransactions")%></a><%end if%></td>
                                   <% end if
                            else %>
                                <td class="right"><%if ap_cc_void then%><a href="javascript:rejectTrans();"><%=DisplayPhrase(phraseDictionary,"Voidcheckedtransactions")%></a><%end if%></td>
                            <% end if %>
                        </tr>
                      </table>
                      <br />
					  <span id="lastBatch"><%=xssStr(allHotWords(671))%><%=LastBatch %></span>
					<%	if ap_cc_settle AND (ccLoc<>"-1" OR oneMID) AND cSwiped<>"" then %>
                    <input onClick="SendBatch();" type="button" name="settleButton" value="<%=DisplayPhraseAttr(phraseDictionary,"Batchandsettleselectedtransactions")%>" />				
					<% end if 
						if ActiveBatch then
					%>
						<span id="activeBatch"><%=DisplayPhrase(phraseDictionary,"Activelybatching")%></span>	
					<%
						end if
					 %>

<% 				end if ' genreport = true
 			end if ' exportreport != true
%>
			</form>
      </td>
                </tr>
              </table>
            </td>
          </tr>
        </table>
    </td>
    </tr>
		</table>
<% pageEnd %>
<!-- #include file="post.asp" -->

<%
	
    Set rsEntry = Nothing

	

end if ' end session
%>
