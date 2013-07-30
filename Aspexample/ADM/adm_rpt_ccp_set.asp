<%@ CodePage=65001 %>
<!-- #include file="init.asp" -->
<%
'Dim SessionFarm
'set SessionFarm = Server.CreateObject("SessionFarm.SFSession")
response.charset="utf-8"
Server.ScriptTimeout = 300    '5 min (value in seconds)
%>
<%

dim phraseDictionary
set phraseDictionary = LoadPhrases("BusinessmodesettledtransactionsPage", 142)

%>
		<!-- #include file="inc_accpriv.asp" -->
<%
if not Session("Pass") OR Session("Admin")="false" OR NOT validAccessPriv("RPT_CCP") then 
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

	dim category : category = ""
	dim tmpCat : tmpCat = ""
	if (RQ("category"))<>"" then
		tmpCat = RQ("category")
		category = Replace(tmpCat, " ", "")
	elseif (RF("category"))<>"" then
		tmpCat = RF("category") 
		category = Replace(tmpCat, " ", "")
	end if

	Dim cSDate, cEDate, disMode,  VoidOk, ss_EnableACH, tmpAmex, tmpVisa, tmpMC, tmpDisc, ACHHotWord, ccProcessor, pSaleID, genReport, SplitVisaMC, pLastFour

    dim rsEntry
	set rsEntry = Server.CreateObject("ADODB.Recordset")

	if request.form("delTransID")<>"" then
		strSQL = "UPDATE tblCCTrans SET Settled=0, Status=N'Manually Rejected', TransTime=" & DateSep & DateAdd("n", Session("tzOffset"),Now) & DateSep & " WHERE TransactionNumber=" & request.form("delTransID")
		cnWS.execute strSQL
	end if

	ss_EnableACH = checkStudioSetting("tblCCOpts", "EnableACH")
	ss_AutoBatchMethod = checkStudioSetting("Studios", "AutoBatchMethod")
	VoidOk = validAccessPriv("TB_VOID")
	
	ACHHotWord = "ACH"
	ACHHotWord =  getHotWord(109)
		
	if request.form("requiredtxtDateStart")<>"" then
		Call SetLocale(session("mvarLocaleStr"))
			cSDate = CDATE(request.form("requiredtxtDateStart"))
		Call SetLocale("en-us")
	else
		cSDate = DateValue(DateAdd("y",-3,DateAdd("n", Session("tzOffset"),Now)))
	end if
	if request.form("requiredtxtDateEnd")<>"" then
		Call SetLocale(session("mvarLocaleStr"))
			cEDate = CDATE(request.form("requiredtxtDateEnd"))
		Call SetLocale("en-us")
	else
		cEDate = DateValue(CDATE(DateAdd("n", Session("tzOffset"),Now)))
	end if
	'if request.form("optDate")="all" then
	'	disMode = "all"
	'else
		disMode = "range"
	'end if
	
	if request.querystring("loc")<>"" then
		ccLoc = request.querystring("loc")
	elseif request.form("optCCLocation")<>"" then
		ccLoc = request.form("optCCLocation")
	else
		'if session("curLocation")="0" then
			ccLoc = "-2"
		'else
		'	ccLoc = session("curLocation")
		'end if
	end if

	if request.querystring("saleID")<>"" then
		pSaleID = request.querystring("saleID")
	elseif request.form("txtSaleID")<>"" then
		pSaleID = request.form("txtSaleID")
	else
		pSaleID = ""
	end if
	if NOT isNumeric(pSaleID) then
		pSaleID = ""
	end if
	
	if request.form("txtCCLastFour")<>"" AND  isNum(request.form("txtCCLastFour")) then
		pLastFour = request.form("txtCCLastFour")
	else
		pLastFour = ""
	end if
	
		
	if request.form("frmGenReport")="true" or request.querystring("saleID")<>"" then
		genReport = true
	end if
	        if implementationSwitchIsEnabled("BluefinCanada") then
			    strSQL = "SELECT tblCCOpts.EnableACH, tblCCOpts.ccVisa, tblCCOpts.ccMasterCard, tblCCOpts.ccAmericanExpress, tblCCOpts.ccDiscover, Studios.ccProcessor "
			    strSQL = strSQL & "FROM tblCCOpts, Studios WHERE tblCCOpts.StudioID=" & session("StudioID")
            else
                strSQL = "SELECT tblCCOpts.EnableACH, tblCCOpts.ccVisa, tblCCOpts.ccMasterCard, tblCCOpts.ccAmericanExpress, tblCCOpts.ccDiscover, tblCCOpts.ccProcessor "
			    strSQL = strSQL & "FROM tblCCOpts WHERE tblCCOpts.StudioID=" & session("StudioID")
            end if
			rsEntry.CursorLocation = 3
			rsEntry.open strSQL, cnWS
			Set rsEntry.ActiveConnection = Nothing
			if NOT rsEntry.EOF then
				ss_EnableACH = rsEntry("EnableACH")
				tmpAmex = rsEntry("ccAmericanExpress")
				tmpVisa = rsEntry("ccVisa")
				tmpMC = rsEntry("ccMasterCard")
				tmpDisc = rsEntry("ccDiscover")
				ccProcessor = rsEntry("ccProcessor")
			end if
			rsEntry.close

            if ccProcessor = "MON" OR (ccProcessor = "OP" AND Session("countryCode")="CA") then
                SplitVisaMC = true
            else
                SplitVisaMC = false
            end if

	if NOT request.form("frmExpReport")="true" then 
%>
<!-- #include file="pre.asp" -->
<!-- #include file="frame_bottom.asp" -->
<%= js(array("adm/adm_rpt_ccp_set", "mb", "MBS", "reportFavorites", "plugins/jquery.SimpleLightBox")) %>
<!-- American/Canada=2 format mm/dd/yyyy --> <!-- European/Rest of the world=1 format dd-mm-yyyy -->

<%= js(array("calendar" & dateFormatCode)) %>
<%= css(array("SimpleLightBox")) %> 
		<!-- #include file="inc_help_content.asp" -->

<!-- #include file="../inc_date_ctrl.asp" -->
<!-- #include file="../inc_ajax.asp" -->
<!-- #include file="../inc_val_date.asp" -->
    <!-- #include file="css/site_setup.asp" -->
	<!-- #include file="inc_user_options.asp" -->  
    <%if ccProcessor="TCI" then %>
    <script type="text/javascript">
        $(function() {
            //alert($("select[name='optDisMode']").val());
            $("select[name='optSwipeOrKeyed']").data('prev', $("select[name='optSwipeOrKeyed']").val());
            $("select[name='optSwipeOrKeyed']").change(function() {
                $("select[name='optSwipeOrKeyed']").data('prev', $("select[name='optSwipeOrKeyed']").val());
            });
            if ($("select[name='optDisMode']").val() == "summary") {
                $("select[name='optSwipeOrKeyed']").prop('disabled', true);
            }
            $("select[name='optDisMode']").change(function() {
                if ($("select[name='optDisMode']").val() == "summary") {
                    $("select[name='optSwipeOrKeyed']").data('prev', $("select[name='optSwipeOrKeyed']").val());
                    $("select[name='optSwipeOrKeyed']").val("");
                    $("select[name='optSwipeOrKeyed']").prop('disabled', true);
                }
                else {
                    $("select[name='optSwipeOrKeyed']").prop('disabled', false);
                    $("select[name='optSwipeOrKeyed']").val($("select[name='optSwipeOrKeyed']").data('prev'));
                }
            });
        });
    </script>
    <%end if %>
<% pageStart %>
<% if UseVersionB(TEST_REPORT_LANDING_PAGE) then %>
	<div class="headText breadcrumbs-old" valign="bottom">
	<span class="breadcrumb-item"><a href="/reportslandingpage/index"><%=DisplayPhrase(reportPageTitlesDictionary, "Reports")%></a></span>
	<span class="breadcrumb-item">&raquo;</span>
	<%if category <> "" then%>
	<span class="breadcrumb-item active"><a href="/reportslandingpage/<%=category%>category"><%=GetReportsCategoryPagePhrase(category)%></a></span>
	<span class="breadcrumb-item">&raquo;</span>
	<%end if %>
	<%=DisplayPhrase(reportPageTitlesDictionary,"Settledtransactions")%>
         <% showTrainingMovieIcon("21044446-managing-credit-card-processing#batch") %>

	<div id="add-to-favorites">
		<span id="" class="favorite"></span><span class="text"><%=DisplayPhrase(reportPageTitlesDictionary, "Addtofavorites")%></span>
	</div>
	</div>
<%else %>
<div class="headText" valign="bottom"><b><%=DisplayPhrase(pageTitlesDictionary,"Settledtransactions")%></b></div>
<% showTrainingMovieIcon("21044446-managing-credit-card-processing#batch") %>
<%end if %>
<table height="100%" width="<%=strPageWidth%>" cellspacing="0">    
<tr> 
     <td valign="top" height="100%" width="100%">
        <table class="center" cellspacing="0" width="90%" height="100%">
          <form name="frmSales" action="adm_rpt_ccp_set.asp" method="POST">
		  <% if UseVersionB(TEST_REPORT_LANDING_PAGE) then %>
			<% if category <> "" then %>
				<input type="hidden" name="category" id="category" value="<%=category %>" />
			<%end if %>
			<input type="hidden" name="reportUrl" id="reportUrl" value="<%=Request.ServerVariables("SCRIPT_NAME") %>" />
		  <%end if %>
		  <input type="hidden" name="frmExpReport" value="">
		  <input type="hidden" name="frmGenReport" value="">
		  <input type="hidden" name="delTransID" value="">
				<div id="topdiv">
          <tr> 
              <td valign="bottom" class="mainText right" height="18">
              <!-- #include file="inc_batch_nav.asp" -->
              </td>
          </tr>
					</div>
					<tr>
					<td class="right" valign="bottom" height="26"> 
                      <table style="float:right;" class="mainText border4" cellspacing="0">
                        <tr> 
                          <td class="center-ch nowrap" valign="bottom" style="background-color:#F2F2F2;"><b>
							<select name="optDisMode">
								<option value="detail"<% if request.form("optDisMode")="detail" then response.write " selected" end if %>><%=xssStr(allHotWords(674))%></option>
								<option value="summary"<% if request.form("optDisMode")="summary" or request.form("optDisMode")="" then response.write " selected" end if %>><%=xssStr(allHotWords(675))%></option>
							</select>
                            <%=xssStr(allHotWords(77))%>: 
                            <input type="text"  name="requiredtxtDateStart" value="<%=FmtDateShort(cSDate)%>" onBlur="validateDate(this, '<%=FmtDateShort(cSDate)%>', true);" class="date">
                      <script type="text/javascript">
                      	var cal1 = new tcal({ 'formname': 'frmSales', 'controlname': 'requiredtxtDateStart' });
                      	cal1.a_tpl.yearscroll = true;
		</script>
                            &nbsp;<%=xssStr(allHotWords(79))%>: 
                            <input type="text"  name="requiredtxtDateEnd" value="<%=FmtDateShort(cEDate)%>" onBlur="validateDate(this, '<%=FmtDateShort(cEDate)%>', true);" class="date">
                      <script type="text/javascript">
                      	var cal2 = new tcal({ 'formname': 'frmSales', 'controlname': 'requiredtxtDateEnd' });
                      	cal2.a_tpl.yearscroll = true;
		</script>
                            &nbsp; </b>
							<% if validAccessPriv("RPT_TAG") then 
	  							taggingFilter 
							end if %>
                            <input name="Button" type="button" value="<%=xssStr(allHotWords(226))%>" onClick="genCCP();">
							<span class="icon-button" style="vertical-align: middle;" title="<%=xssStr(allHotWords(658))%>" ><a onClick="exportCCP();" ><img src="<%= contentUrl("/asp/adm/images/export-to-excel-20px.png") %>" /></a></span>
							<% if not Session("Pass") OR Session("Admin")="false" OR NOT validAccessPriv("RPT_TAG") then 
							else 
								taggingButtons("frmSales") 
							end if %>
						  </td>
                        </tr>

                    </table>
                  </td>
					</tr>
          <tr> 
            <td valign="top" class="mainTextBig">
              <table class="mainText center" width="95%" cellspacing="0">
                <tr> 
                  <td  colspan="2" valign="top" class="mainTextBig">
                    <br />
                    <table class="mainText center" cellspacing="0" width="90%">
              <tr>
                          <td valign="top" align=left>

								<select name="optCCType" >
                                  <option value="-1"><%=DisplayPhrase(phraseDictionary,"Allcreditcardsandach")%></option>
<%
			if tmpVisa or tmpMC then
%>
                                  <option value="4" <%if request.form("optCCType")="4" then response.write "selected" end if%>><%=xssStr(allHotWords(660))%></option>
<% 
			end if
			if tmpAmex then
%>
                                  <option value="3" <%if request.form("optCCType")="3" then response.write "selected" end if%>><%=xssStr(allHotWords(659))%></option>
<%
			end if
			if tmpDisc then
%>
                                  <option value="5" <%if request.form("optCCType")="5" then response.write "selected" end if%>><%=xssStr(allHotWords(661))%></option>
<%
			end if
			if ccProcessor = "TCI" then
%>
                                  <option value="151" <%if request.form("optCCType")="151" then response.write "selected" end if%>><%=xssStr(allHotWords(660))%> / <%=xssStr(allHotWords(661))%></option>
<%
			end if
%>
							<%if ss_EnableACH then %>
								  <option value="100" <%if request.form("optCCType")="100" then response.write "selected" end if%>><%=xssStr(allHotWords(681))%></option>
								  <option value="101" <%if request.form("optCCType")="101" then response.write "selected" end if%>><%=DisplayPhrase(phraseDictionary,"Creditcardsonly")%></option>
							<%end if%>
                                </select>
								<select name="optSwipeOrKeyed">
									<option value=""><%=DisplayPhrase(phraseDictionary,"Allkeyedorswiped")%></option>
									<option value="keyed" <%if request.form("optSwipeOrKeyed")="keyed" then response.write "selected" end if %>><%=DisplayPhrase(phraseDictionary,"Keyedonly")%></option>
									<option value="swiped" <%if request.form("optSwipeOrKeyed")="swiped" then response.write "selected" end if %>><%=DisplayPhrase(phraseDictionary,"Swipedonly")%></option>
								</select>
							  <select name="optCCLocation" >
<%
				strSQL = "SELECT LocationID, LocationName FROM Location WHERE  (NOT (MID IS NULL)) AND (Active = 1) ORDER BY LocationName"
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
							<option value="-2" <%if ccLoc="-2" then response.write "selected" end if%>><%=xssStr(allHotWords(479))%></option>
							</select>
<%
		Dim strTempName
		Dim intCount

		strSQL = "SELECT DISTINCT tblCCTrans.BatchNumber "
		strSQL = strSQL & "FROM tblCCTrans "
		strSQL = strSQL & "WHERE tblCCTrans.Settled=1 AND (NOT (BatchNumber IS NULL)) "
		if ccLoc<>"-2" then
			strSQL = strSQL & " AND tblCCTrans.LocationID=" & ccLoc
		end if
		if disMode = "range" then
			strSQL = strSQL & " AND TransTime >= " & DateSep & cSDate & DateSep & " "
			strSQL = strSQL & " AND TransTime <= " & DateSep & DateAdd("d", 1, cEDate) & DateSep & " "
		end if
		strSQL = strSQL & "ORDER BY BatchNumber DESC"
%>
                            <select name="optBatchNum" >
                              <option value="0"><%=DisplayPhrase(phraseDictionary,"Allsettlements")%></option>
								<option value="-1" <%if request.form("optBatchNum")="-1" then response.write "selected" end if%>><%=DisplayPhrase(phraseDictionary,"Creditsonly")%></option>
<%
	rsEntry.CursorLocation = 3
	rsEntry.open strSQL, cnWS
	Set rsEntry.ActiveConnection = Nothing

	do While NOT rsEntry.EOF
%>
                                                    <option value="<%=rsEntry("BatchNumber")%>" <% if request.form("optBatchNum")=CSTR(rsEntry("BatchNumber")) then response.write "selected" end if %>><%=xssStr(allHotWords(679))%><%=rsEntry("BatchNumber")%></option>
<%
		rsEntry.MoveNext
	loop
	rsEntry.close
%>
                           </select>

							&nbsp;&nbsp;&nbsp;<strong><%=UCASE(xssStr(allHotWords(170)))%></strong>&nbsp;&nbsp;&nbsp;
							<%=DisplayPhrase(phraseDictionary,"Searchbysaleid")%>
							<input type="text" name="txtSaleID" maxlength="14" size="9" value="<%=pSaleID%>" onKeyDown="return mb.chkKey(this, event, genCCP);" <%if pSaleID<>"" then response.write " style=""background-color:#FFFF99""" end if%>>
							<%=DisplayPhrase(phraseDictionary,"Searchbyccachlast4")%>
							<input type="text" name="txtCCLastFour" maxlength="4" size="9" value="<%=pLastFour%>" onKeyDown="return mb.chkKey(this, event, genCCP);" <%if pLastFour<>"" then response.write " style=""background-color:#FFFF99""" end if%>>

<%

			else ' we are exporting %>
<%
				Dim stFilename
				stFilename="attachment; filename=SettledTransactions " & Replace(cSDate,"/","-") & " - " & Replace(cEDate,"/","-") & ".xls" 
				Response.ContentType = "application/vnd.ms-excel" 
				Response.AddHeader "Content-Disposition", stFilename 
%>


<%			end if 'export check 

if genReport then

	if request.form("frmTagClients")="true" then 'tag clients sql
		strSQL = "SELECT tblCCTrans.ClientID "
		strSQL = strSQL & "FROM CLIENTS "
		strSQL = strSQL & "INNER JOIN tblCCTrans ON CLIENTS.ClientID = tblCCTrans.ClientID "
		strSQL = strSQL & "LEFT OUTER JOIN Sales ON tblCCTrans.SaleID = Sales.SaleID "
		strSQL = strSQL & "LEFT OUTER JOIN tblPayments ON tblPayments.CCTransID = tblCCTrans.TransactionNumber "
		strSQL = strSQL & "LEFT OUTER JOIN Location ON Location.LocationID = Sales.LocationID "
		if request.form("optFilterTagged")="on" then
			strSQL = strSQL & " INNER JOIN tblClientTag ON (tblClientTag.ClientID = CLIENTS.ClientID "
			if session("mVarUserID")<>"" then
				strSQL = strSQL & " AND tblClientTag.smodeID = " & session("mVarUserID")
			end if
			strSQL = strSQL & " ) "
		end if
		strSQL = strSQL & "WHERE (((tblCCTrans.Settled)=1) "
		if pSaleID<>"" OR pLastFour<>"" then	'search by saleID and CC/ACH last4
			if pSaleID<>"" then strSQL = strSQL & " AND tblCCTrans.SaleID=" & pSaleID
			if pLastFour<>"" then  
				if request.form("optCCType")="-1" then 'All
					strSQL = strSQL & " AND (tblCCTrans.CCLastFour=" & pLastFour
					strSQL = strSQL & " OR tblCCTrans.ACHLastFour LIKE '%" & pLastFour & "') "
				elseif  request.form("optCCType")= "100" then 'ACH Only
					strSQL = strSQL & " AND tblCCTrans.ACHLastFour LIKE '%" & pLastFour  & "'"
				else 'CC Only or specific CC
					strSQL = strSQL & " AND tblCCTrans.CCLastFour=" & pLastFour
				end if
			end if
		else
			if ccLoc<>"-2" then
				strSQL = strSQL & " AND (Sales.LocationID = " & ccLoc & ") "
			end if
			if request.form("optBatchNum")<>"" AND request.form("optBatchNum")<>"0" AND request.form("optBatchNum")<>"-1" then
				strSQL = strSQL & " AND BatchNumber=" & request.form("optBatchNum") & " "
			elseif request.form("optBatchNum")="-1" then
				strSQL = strSQL & " AND tblCCTrans.Status='Credit'"
			end if
			if disMode = "range" then
				strSQL = strSQL & " AND tblCCTrans.TransTime >= " & DateSep & cSDate & DateSep & " "
				strSQL = strSQL & " AND tblCCTrans.TransTime <= " & DateSep & DateAdd("d", 1, cEDate) & DateSep & " "
			end if
			if request.form("optSwipeOrKeyed")<>"" then
				strSQL = strSQL & " AND (tblCCTrans.ACHName IS NULL) "
				if request.form("optSwipeOrKeyed")="swiped" then
					strSQL = strSQL & " AND (tblCCTrans.CCSwiped=1) "
				else
					strSQL = strSQL & " AND (tblCCTrans.CCSwiped=0) "
				end if
			end if
			if request.form("optCCType")<>"" AND request.form("optCCType")<>"-1" then
				if request.form("optCCType")="4" then
					strSQL = strSQL & " AND (tblPayments.PaymentMethod=4 OR tblCCTrans.ccType = 'Visa' OR tblCCTrans.ccType = 'Master Card') "
				elseif request.form("optCCType")="5" then
					strSQL = strSQL & " AND (tblPayments.PaymentMethod=5 OR tblCCTrans.ccType = 'Discover') "
				elseif request.form("optCCType")="151" then
					strSQL = strSQL & " AND (tblPayments.PaymentMethod=4 OR tblCCTrans.ccType = 'Visa' OR tblCCTrans.ccType = 'Master Card' OR tblPayments.PaymentMethod=5 OR tblCCTrans.ccType = 'Discover') "
				elseif request.form("optCCType")="3" then
					strSQL = strSQL & " AND (tblPayments.PaymentMethod=3 OR tblCCTrans.ccType = 'American Express') "
				elseif request.form("optCCType")="6" then
					strSQL = strSQL & " AND (tblPayments.PaymentMethod=6) "
				elseif request.form("optCCType")="100" then
					strSQL = strSQL & " AND (NOT (tblCCTrans.ACHName IS NULL)) "
				elseif request.form("optCCType")="101" then
					strSQL = strSQL & " AND (tblCCTrans.ACHName IS NULL) "
				else
					strSQL = strSQL & " AND (tblPayments.PaymentMethod=" & request.form("optCCType") & ") "
				end if
			end if
			strSQL = strSQL & " AND (tblPayments.PaymentMethod <> 96 )"
		end if	'search by saleID or CC/ACH last 4
		strSQL = strSQL & ")"
		
	response.write debugSQL(strSQL, "SQL")
		'response.end
		'rsEntry.CursorLocation = 3
		'rsEntry.open strSQL, cnWS
		'Set rsEntry.ActiveConnection = Nothing
					
		if request.form("frmTagClientsNew")="true" then
			clearAndTagQuery(strSQL)
		else
			tagQuery(strSQL)
		end if
					
		'rsEntry.close
	end if 

	if request.form("optDisMode")="detail" or request.form("optDisMode")="" then
		strSQL = "SELECT tblCCTrans.Status, tblCCTrans.Settled, tblCCTrans.CCLastFour, tblCCTrans.ACHLastFour, tblCCTrans.ccAmt, tblCCTrans.BatchNumber, tblCCTrans.ClientID, tblCCTrans.TransactionNumber, tblCCTrans.AuthCode, tblCCTrans.TransTime, tblCCTrans.SaleID, tblCCTrans.ACHName, tblCCTrans.OrderID, tblCCTrans.CCSwiped, tblCCTrans.AuthTime, tblCCTrans.OutputFileNum, tblCCTrans.MerchantID, tblCCTrans.TerminalID, CLIENTS.LastName, CLIENTS.FirstName, tblPayments.PaymentMethod, tblCCTrans.Cardholder, Location.LocationName, Sales.LocationID "
		strSQL = strSQL & "FROM CLIENTS INNER JOIN tblCCTrans ON CLIENTS.ClientID = tblCCTrans.ClientID "
		strSQL = strSQL & "LEFT OUTER JOIN tblPayments ON tblPayments.CCTransID = tblCCTrans.TransactionNumber " 
		strSQL = strSQL & "LEFT OUTER JOIN Sales ON tblCCTrans.SaleID = Sales.SaleID "
		strSQL = strSQL & "LEFT OUTER JOIN Location ON Location.LocationID = Sales.LocationID "
		if request.form("optFilterTagged")="on" then 
			strSQL = strSQL & "INNER JOIN tblClientTag ON CLIENTS.ClientID = tblClientTag.clientID "
		end if
		strSQL = strSQL & "WHERE ((tblCCTrans.Settled=1) "
		if request.form("optFilterTagged")="on" then 
			if session("mvaruserID")<>"" then
				strSQL = strSQL & "AND (tblClientTag.smodeID = " & session("mvaruserID") & ") " 
			else
				strSQL = strSQL & "AND (tblClientTag.smodeID = 0) " 
			end if
		end if
		if pSaleID<>"" OR pLastFour<>"" then	'search by saleID and CC/ACH last4
			if pSaleID<>"" then strSQL = strSQL & " AND tblCCTrans.SaleID=" & pSaleID
				If pLastFour<>"" then
					if request.form("optCCType")="-1" then 'All
						strSQL = strSQL & " AND (tblCCTrans.CCLastFour=" & pLastFour
						strSQL = strSQL & " OR tblCCTrans.ACHLastFour LIKE '%" & pLastFour & "') "
					elseif  request.form("optCCType")= "100" then 'ACH Only
						strSQL = strSQL & " AND tblCCTrans.ACHLastFour LIKE '%" & pLastFour  & "'"
					else 'CC Only or specific CC
						strSQL = strSQL & " AND tblCCTrans.CCLastFour=" & pLastFour
					end if
				end if	
		else
			if ccLoc<>"-2" then
				strSQL = strSQL & " AND (Sales.LocationID = " & ccLoc & ") "
			end if
			if request.form("optBatchNum")<>"" AND request.form("optBatchNum")<>"0" AND request.form("optBatchNum")<>"-1" then
				strSQL = strSQL & " AND BatchNumber=" & request.form("optBatchNum") & " "
			elseif request.form("optBatchNum")="-1" then
				strSQL = strSQL & " AND tblCCTrans.Status='Credit'"
			end if
			if disMode = "range" then
				strSQL = strSQL & " AND tblCCTrans.TransTime >= " & DateSep & cSDate & DateSep & " "
				strSQL = strSQL & " AND tblCCTrans.TransTime <= " & DateSep & DateAdd("d", 1, cEDate) & DateSep & " "
			end if
			if request.form("optSwipeOrKeyed")<>"" then
				strSQL = strSQL & " AND (tblCCTrans.ACHName IS NULL) "
				if request.form("optSwipeOrKeyed")="swiped" then
					strSQL = strSQL & " AND (tblCCTrans.CCSwiped=1) "
				else
					strSQL = strSQL & " AND (tblCCTrans.CCSwiped=0) "
				end if
			end if
			if request.form("optCCType")<>"" AND request.form("optCCType")<>"-1" then
				if request.form("optCCType")="4" then
					strSQL = strSQL & " AND (tblPayments.PaymentMethod=4 OR tblCCTrans.ccType = 'Visa' OR tblCCTrans.ccType = 'Master Card') "
				elseif request.form("optCCType")="5" then
					strSQL = strSQL & " AND (tblPayments.PaymentMethod=5 OR tblCCTrans.ccType = 'Discover') "
				elseif request.form("optCCType")="151" then
					strSQL = strSQL & " AND (tblPayments.PaymentMethod=4 OR tblCCTrans.ccType = 'Visa' OR tblCCTrans.ccType = 'Master Card' OR tblPayments.PaymentMethod=5 OR tblCCTrans.ccType = 'Discover') "
				elseif request.form("optCCType")="3" then
					strSQL = strSQL & " AND (tblPayments.PaymentMethod=3 OR tblCCTrans.ccType = 'American Express') "
				elseif request.form("optCCType")="6" then
					strSQL = strSQL & " AND (tblPayments.PaymentMethod=6) "
				elseif request.form("optCCType")="100" then
					strSQL = strSQL & " AND (NOT (tblCCTrans.ACHName IS NULL)) "
				elseif request.form("optCCType")="101" then
					strSQL = strSQL & " AND (tblCCTrans.ACHName IS NULL) "
				else
					strSQL = strSQL & " AND (tblPayments.PaymentMethod=" & request.form("optCCType") & ") "
				end if
			end if
		end if	'search by saleID or CC/ACH last 4
		strSQL = strSQL & " ) ORDER BY tblCCTrans.TransTime DESC, tblCCTrans.TransactionNumber DESC"
	else ' summary mode
		strSQL = "SELECT COUNT(DISTINCT tblCCTrans.TransactionNumber) as NumTrans, SUM(CASE WHEN tblCCTrans.Status = 'Credit' THEN (tblCCTrans.ccAmt * -1) ELSE (tblCCtrans.ccAmt) END) as BatchTotal, MIN(tblCCTrans.TransTime) as TransTime, "
		if ccProcessor <> "TCI" then
			strSQL = strSQL & " tblCCTrans.BatchNumber, "
		else 
			strSQL = strSQL & " '' as BatchNumber, "
		end if
    if ccProcessor = "MON" then 'CB 7/23/09 - Moneris to report separately for VISA and MC
      strSQL = strSQL & "tblCCTrans.ccType "
    elseif ccProcessor = "TCI" then
      strSQL = strSQL & "CASE WHEN tblCCTrans.ccType=N'Visa' OR tblCCTrans.ccType=N'Master Card' OR tblCCTrans.ccType=N'Discover' then N'Visa/MC/Discover' ELSE tblCCTrans.ccType END AS ccType"
    else
      strSQL = strSQL & "CASE WHEN tblCCTrans.ccType=N'Visa' OR tblCCTrans.ccType=N'Master Card' then N'Visa/MC' ELSE tblCCTrans.ccType END AS ccType"
    end if
    if ccProcessor <> "TCI" then
    strSQL = strSQL & ", tblCCTrans.CCSwiped "
    end if
    if ccProcessor = "TCI" then 
			strSQL = strSQL & ", tblCCTrans.MerchantID AS LocationName "
		end if
		if ccProcessor <> "TCI" then 
			if ss_AutoBatchMethod<>2 then
				strSQL = strSQL & ", Location.LocationName "
			end if
		end if
		strSQL = strSQL & "FROM tblCCTrans "
		strSQL = strSQL & "LEFT OUTER JOIN Sales ON tblCCTrans.SaleID = Sales.SaleID "
		strSQL = strSQL & "LEFT OUTER JOIN tblPayments ON tblPayments.CCTransID = tblCCTrans.TransactionNumber " 
		strSQL = strSQL & "LEFT OUTER JOIN Location ON Location.LocationID = Sales.LocationID "
		if request.form("optFilterTagged")="on" then 
			strSQL = strSQL & "INNER JOIN tblClientTag ON tblCCTrans.ClientID = tblClientTag.clientID "
		end if
		strSQL = strSQL & "WHERE ((tblCCTrans.Settled=1) "
		if request.form("optFilterTagged")="on" then 
			if session("mvaruserID")<>"" then
				strSQL = strSQL & "AND (tblClientTag.smodeID = " & session("mvaruserID") & ") " 
			else
				strSQL = strSQL & "AND (tblClientTag.smodeID = 0) " 
			end if
		end if
		if pSaleID<>"" OR pLastFour<>"" then	'search by saleID and CC/ACH last4
			if pSaleID<>"" then strSQL = strSQL & " AND tblCCTrans.SaleID=" & pSaleID
				If pLastFour<>"" then
					if request.form("optCCType")="-1" then 'All
						strSQL = strSQL & " AND (tblCCTrans.CCLastFour=" & pLastFour
						strSQL = strSQL & " OR tblCCTrans.ACHLastFour LIKE '%" & pLastFour & "') "
					elseif  request.form("optCCType")= "100" then 'ACH Only
						strSQL = strSQL & " AND tblCCTrans.ACHLastFour LIKE '%" & pLastFour  & "'"
					else 'CC Only or specific CC
						strSQL = strSQL & " AND tblCCTrans.CCLastFour=" & pLastFour
					end if
				end if
		else
			if ccLoc<>"-2" then
				strSQL = strSQL & " AND (Sales.LocationID = " & ccLoc & ") "
			end if
			if request.form("optBatchNum")<>"" AND request.form("optBatchNum")<>"0" AND request.form("optBatchNum")<>"-1" then
				strSQL = strSQL & " AND BatchNumber=" & request.form("optBatchNum") & " "
			elseif request.form("optBatchNum")="-1" then
				strSQL = strSQL & " AND tblCCTrans.Status='Credit'"
			end if
			if disMode = "range" then
				strSQL = strSQL & " AND tblCCTrans.TransTime >= " & DateSep & cSDate & DateSep & " "
				strSQL = strSQL & " AND tblCCTrans.TransTime <= " & DateSep & DateAdd("d", 1, cEDate) & DateSep & " "
			end if
			if request.form("optSwipeOrKeyed")<>"" then
				strSQL = strSQL & " AND (tblCCTrans.ACHName IS NULL) "
				if request.form("optSwipeOrKeyed")="swiped" then
					strSQL = strSQL & " AND (tblCCTrans.CCSwiped=1) "
				else
					strSQL = strSQL & " AND (tblCCTrans.CCSwiped=0) "
				end if
			end if
			if request.form("optCCType")<>"" AND request.form("optCCType")<>"-1" then
				if request.form("optCCType")="4" then
					strSQL = strSQL & " AND (tblPayments.PaymentMethod=4 OR tblCCTrans.ccType = 'Visa' OR tblCCTrans.ccType = 'Master Card') "
				elseif request.form("optCCType")="5" then
					strSQL = strSQL & " AND (tblPayments.PaymentMethod=5 OR tblCCTrans.ccType = 'Discover') "
				elseif request.form("optCCType")="151" then
					strSQL = strSQL & " AND (tblPayments.PaymentMethod=4 OR tblCCTrans.ccType = 'Visa' OR tblCCTrans.ccType = 'Master Card' OR tblPayments.PaymentMethod=5 OR tblCCTrans.ccType = 'Discover') "
				elseif request.form("optCCType")="3" then
					strSQL = strSQL & " AND (tblPayments.PaymentMethod=3 OR tblCCTrans.ccType = 'American Express') "
				elseif request.form("optCCType")="6" then
					strSQL = strSQL & " AND (tblPayments.PaymentMethod=6) "
				elseif request.form("optCCType")="100" then
					strSQL = strSQL & " AND (NOT (tblCCTrans.ACHName IS NULL)) "
				elseif request.form("optCCType")="101" then
					strSQL = strSQL & " AND (tblCCTrans.ACHName IS NULL) "
				else
					strSQL = strSQL & " AND (tblPayments.PaymentMethod=" & request.form("optCCType") & ") "
				end if
			end if
		end if	'search by saleID
		strSQL = strSQL & " AND (tblPayments.PaymentMethod <> 96 )) GROUP BY "
		if ccProcessor <> "TCI" then
			strSQL = strSQL & " tblCCTrans.BatchNumber, "
		else
			strSQL = strSQL & " DATEPART(Y,tblCCTrans.TransTime), "
		end if
		
		if SplitVisaMC then 'CB 7/23/09 - Moneris to report separately for VISA and MC
			strSQL = strSQL & " tblCCTrans.ccType "
        elseif ccProcessor = "TCI" then
            strSQL = strSQL & " CASE WHEN tblCCTrans.ccType=N'Visa' OR tblCCTrans.ccType=N'Master Card' OR tblCCTrans.ccType=N'Discover' then N'Visa/MC/Discover' ELSE tblCCTrans.ccType END "
        else
            strSQL = strSQL & " CASE WHEN tblCCTrans.ccType=N'Visa' OR tblCCTrans.ccType=N'Master Card' then N'Visa/MC' ELSE tblCCTrans.ccType END "
        end if
		if ccProcessor <> "TCI" then 
		strSQL = strSQL & ", tblCCTrans.CCSwiped "
		end if
		if ccProcessor = "TCI" then 
			strSQL = strSQL & ", tblCCTrans.MerchantID "
		end if
		if ccProcessor <> "TCI" then 
			if ss_AutoBatchMethod<>2 then
				strSQL = strSQL & ", Location.LocationName "
			end if
		end if
		if ccProcessor <> "TCI" then
			strSQL = strSQL & "ORDER BY tblCCTrans.BatchNumber DESC"
		else
			strSQL = strSQL & "ORDER BY DATEPART(Y,tblCCTrans.TransTime) DESC"
		end if
		
	end if
	response.write debugSQL(strSQL, "SQL")
        rsEntry.CursorLocation = 3
		rsEntry.open strSQL, cnWS
		Set rsEntry.ActiveConnection = Nothing

			if request.form("optDisMode")="detail" OR request.form("optDisMode")="" then %>
                    <table class="mainText" cellspacing="0" width="100%">
				<% if NOT request.form("frmExpReport")="true" then  %>
                    <tr style="background-color:<%=session("pageColor2")%>;"> 
                      <td colspan=13><img height="2" width="100%" src="<%= contentUrl("/asp/adm/images/trans.gif") %>"></td>
                    </tr>
				<% end if %>
                      <tr bgcolor="<%= session("pageColor4")%>"> 
                        <th class="whiteHeader"><b>&nbsp;<%=xssStr(allHotWords(57))%> / <%=xssStr(allHotWords(58))%></b></th>
                        <th class="whiteHeader"><b>&nbsp;<%=xssStr(allHotWords(12))%></b></th>
                        <th class="whiteHeader" nowrap align="left"><b>&nbsp;<%=xssStr(allHotWords(44))%> / &nbsp;<%=xssStr(allHotWords(218))%>&nbsp;</b></th>
                        <th nowrap class="whiteHeader center-ch"><b>&nbsp;<%=xssStr(allHotWords(35))%>&nbsp;</b></th>
                        <th nowrap class="whiteHeader center-ch"><b>&nbsp;<%=xssStr(allHotWords(424))%>&nbsp;</b></th>
                        <th nowrap class="whiteHeader center-ch"><b><%=xssStr(allHotWords(666))%></b></th>
                        <th nowrap class="whiteHeader center-ch"><%=xssStr(allHotWords(218))%></th>
                        <th nowrap class="whiteHeader center-ch"><b><%=xssStr(allHotWords(667))%>&nbsp;</b></th>
                        <th nowrap class="whiteHeader center-ch"><b>&nbsp;<%=xssStr(allHotWords(679))%></b></th>
                        <th nowrap class="whiteHeader center-ch"><b>&nbsp;<%=xssStr(allHotWords(8))%></b></th>
                        <th nowrap class="whiteHeader center-ch"><b>&nbsp;<%=xssStr(allHotWords(682))%></b></th>
                        <th nowrap class="whiteHeader center-ch"><b><%=xssStr(allHotWords(60))%></b></th>
                        <th class="whiteHeader" nowrap bgcolor="<%= session("pageColor4")%>" class="center-ch">&nbsp;</th>
                      </tr>
<%			else ' summary mode %>
                    <table class="mainText center-ch" cellspacing="0" width="75%">
				<% if NOT request.form("frmExpReport")="true" then  %>
                    <tr style="background-color:<%=session("pageColor2")%>;"> 
                      <td colspan=13><img height="2" width="100%" src="<%= contentUrl("/asp/adm/images/trans.gif") %>"></td>
                    </tr>
				<% end if %>
                      <tr bgcolor="<%= session("pageColor4")%>"> 
                        <th class="whiteHeader" align="left"> <b>&nbsp;<%=xssStr(allHotWords(57))%> / <%=xssStr(allHotWords(58))%></b></th>
                        <th class="whiteHeader" nowrap align="left"><b>&nbsp;&nbsp;<%=xssStr(allHotWords(50))%></b></th>
                        <th class="whiteHeader right"><b>&nbsp;<%=xssStr(allHotWords(679))%></b></th>
                        <th class="whiteHeader" nowrap align="left"><b>&nbsp;&nbsp;<%=xssStr(allHotWords(218))%>&nbsp;</b></th>
                        <th class="whiteHeader right"><b>&nbsp;<%=xssStr(allHotWords(35))%>&nbsp;</b></th>
				<% if ss_AutoBatchMethod<>2 then %>
                        <th class="whiteHeader right"><b>&nbsp;<%if ccProcessor<>"TCI" then response.Write xssStr(allHotWords(8)) else response.Write xssStr(allHotWords(682)) end if%></b></th>
				<% end if %>
                        <th class="whiteHeader" nowrap bgcolor="<%= session("pageColor4")%>" class="center-ch">&nbsp;</th>
                      </tr>
<%			end if %>
				<% if NOT request.form("frmExpReport")="true" then  %>
                      <tr style="background-color:<%=session("pageColor2")%>;"> 
                        <td colspan=12><img height="2" width="100%" src="<%= contentUrl("/asp/adm/images/trans.gif") %>"></td>
                      </tr>
				<% end if %>
<%
		rowcount = 0
		totalNumTransactions = 0
		totalAmount = 0 
	if not rsEntry.EOF then		
		Do While NOT rsEntry.EOF
			if request.form("optDisMode")="detail" OR request.form("optDisMode")="" then
				if rowcount=0 then
	%>
						  <tr bgcolor=#F2F2F2> 
							<%
				   rowcount = 1
				else
	%>
						  <tr bgcolor=#FAFAFA> 
	<%
				   rowcount = 0
				end if
	%>
							<td nowrap>&nbsp;<%=FmtDateTime(DateAdd("n", Session("tzOffset"),rsEntry("TransTime")))%></td>
							<td nowrap>&nbsp;<% if NOT request.form("frmExpReport")="true" then %><a href="adm_clt_purch.asp?ID=<%=rsEntry("ClientID")%>&qParam=ph" title="<%=DisplayPhraseAttr(phraseDictionary, "Clicktoviewclientaccounthistory")%>"><% end if %><%=rsEntry("FirstName")%>&nbsp;<%=rsEntry("LastName")%><% if NOT request.form("frmExpReport")="true" then %></a><% end if %></td>
							<td nowrap align="left" >&nbsp;
	<%
								if NOT isNULL(rsEntry("ACHName")) then
									response.write rsEntry("ACHName") & " / " 
								elseif NOT isNULL(rsEntry("CardHolder")) then 
									response.write rsEntry("CardHolder") & " / " 
								end if 
	%>
							  <%
							if NOT isNull(rsEntry("PaymentMethod")) then
								if rsEntry("PaymentMethod")>=3 AND rsEntry("PaymentMethod")<=6 then
									tmpPayID = rsEntry("PaymentMethod")
								else
									tmpPayID = -1
								end if
							else
								tmpPayID = -1
							end if
							
							if NOT isNULL(rsEntry("ACHName")) then
								response.write "" & xssStr(allHotWords(109)) & ""
							elseif tmpPayID = -1 then
								response.write xssStr(allHotWords(246))
							elseif tmpPayID = 3 then
								response.write xssStr(allHotWords(673))
							elseif tmpPayID = 4 then
								response.write xssStr(allHotWords(660))
							elseif tmpPayID = 5 then
								response.write xssStr(allHotWords(661))
							end if
							if NOT isNULL(rsEntry("CCLastFour")) then 
								response.write " " & xssStr(allHotWords(670)) & ":" & padZeros(rsEntry("CCLastFour"),4)
							end if

							if session("Admin")="sa" AND ccProcessor<>"MON" then
								response.write " " & rsEntry("MerchantID")
								if NOT isNULL(rsEntry("TerminalID")) then
									response.write " / " & rsEntry("TerminalID")
								end if
							end if
							'add last4 for ACH transactions, x3 for EZI
							if NOT isNULL(rsEntry("ACHName")) AND rsEntry("ACHLastFour")<>""  AND NOT isNULL(rsEntry("ACHLastFour"))then
								if ccProcessor="EZI" then
									response.write " Last4: x" & Right(rsEntry("ACHLastFour"),3)
								else
									response.write " Last4: " & padZeros(rsEntry("ACHLastFour"),4)
								end if
							end if
%>
							</td>
							<td nowrap class="right"><% if NOT request.form("frmExpReport")="true" then  %>&nbsp;<% end if %><%if rsEntry("Status")="Credit" then response.write "<span style=""color:#990000;"">" end if%><%=FormatNumber(rsEntry("ccAmt")*.01,2)%></td>
							<td nowrap class="center-ch">&nbsp;
<%
							if NOT isNull(rsEntry("SaleID")) then
								if VoidOk and NOT request.form("frmExpReport")="true" then
									response.write "<a title=""" & DisplayPhraseAttr(phraseDictionary,"Clicktogototransaction") & """ href=""adm_tlbx_voidedit.asp?saleno=" & rsEntry("SaleID") & """>" & xssStr(allHotWords(159)) & "/" & xssStr(allHotWords(545)) & "</a>"
								else
									response.write xssStr(allHotWords(159)) & "/" & xssStr(allHotWords(545))
								end if
							else
								response.write xssStr(allHotWords(246))
							end if
							
							if NOT isNULL(rsEntry("ACHName")) AND (ccProcessor="MON" OR ccProcessor="OP") AND DateValue(rsEntry("TransTime")) < DateValue(DateAdd("y",-5,DateAdd("n", Session("tzOffset"),Now))) then
								response.write "&nbsp;<a title=""" & DisplayPhraseAttr(phraseDictionary,"Clicktorejecttransaction") & """ href=""javascript:rejTrx(" & rsEntry("TransactionNumber") & ")"">" & xssStr(DisplayPhrase(phraseDictionary,"Reject")) & "</a>"
							end if							
%>
							&nbsp;</td>
							<td nowrap class="center-ch"><%=rsEntry("AuthCode")%>
									<%if NOT isNull(rsEntry("OrderID")) then%>
										| <%=rsEntry("OrderID")%>
									<%end if%>
							</td>
							<td nowrap class="center-ch">
<%
							if isNULL(rsEntry("ACHName")) AND (ccProcessor="CCP" OR ccProcessor="PMN" OR ccProcessor="MON") then
								if rsEntry("CCSwiped") then 
									response.write xssStr(allHotWords(662))
								else
									response.write xssStr(allHotWords(663))
								end if
							end if
%>
							</td>
							<td nowrap class="center-ch"><%=rsEntry("TransactionNumber")%></td>
							<td nowrap class="center-ch">
<%
							if NOT isNULL(rsEntry("BatchNumber")) then
								response.write rsEntry("BatchNumber")
							elseif NOT isNULL(rsEntry("OutputFileNum")) then
								response.write rsEntry("OutputFileNum")
							end if
%>
							</td>
							<td nowrap class="center-ch"><%=rsEntry("LocationName")%>&nbsp;</td>
							<td nowrap class="center-ch"><%=rsEntry("MerchantID")%>&nbsp;</td>
							<td nowrap class="center-ch"><%if rsEntry("Status")="Credit" then response.write "<span style=""color:#990000;"">" end if%><%=rsEntry("Status")%></td>
							<td align=center>&nbsp; </td>
						  </tr>
					<% if NOT request.form("frmExpReport")="true" then  %>
						<tr style="background-color:<%=session("pageColor4")%>;"> 
						  <td colspan=13><img height="1" width="100%" src="<%= contentUrl("/asp/adm/images/trans.gif") %>"></td>
						</tr>
					<% end if %>
	<%			totalNumTransactions = totalNumTransactions + 1
				if rsEntry("Status")="Credit" then
					totalAmount = totalAmount - rsEntry("ccAmt")
				else
					totalAmount = totalAmount + rsEntry("ccAmt")
				end if
			else 'summary view 
				if tmpBatchNum="" or rsEntry("BatchNumber")<>tmpBatchNum then	%>
				<tr>
					<td nowrap align="left">&nbsp;<%=FmtDateTime(DateAdd("n", Session("tzOffset"),rsEntry("TransTime")))%></td>
					<td nowrap align="left">&nbsp;&nbsp;
					<%	if isNull(rsEntry("BatchNumber")) OR isNull(rsEntry("ccType")) then %>
						<%=xssStr(allHotWords(109))%>
					<% 	else %>
						<%=rsEntry("ccType")%>
					<% 	end if %>
					</td>
					<td nowrap class="right">
					<%	if NOT isNull(rsEntry("BatchNumber")) then %>
						<%=rsEntry("BatchNumber")%>
					<% 	end if %>
					</td>
					<td nowrap align="left">&nbsp;&nbsp;
					<%	
					if NOT isNull(rsEntry("BatchNumber")) then 
					    if ccProcessor <> "TCI" then 
					        if rsEntry("CCSwiped") then response.write xssStr(allHotWords(662)) else response.write xssStr(allHotWords(663)) end if
					    end if 
					end if%>
					</td>
					<td nowrap class="right"><%=FormatNumber(rsEntry("BatchTotal")/100, 2)%></td>
					<% if ss_AutoBatchMethod<>2 then %>
					<td nowrap class="right"><%=rsEntry("LocationName")%>&nbsp;</td>
					<% end if %>
				</tr>
<%					totalAmount = totalAmount + rsEntry("BatchTotal")
					totalNumTransactions = totalNumTransactions + rsEntry("NumTrans")
				end if
			end if ' summary vs detail
		    rsEntry.MoveNext 
        Loop
%>
				<% if NOT request.form("frmExpReport")="true" then  %>
                      <tr style="background-color:<%=session("pageColor4")%>;"> 
                        <td colspan=12><img height="1" width="100%" src="<%= contentUrl("/asp/adm/images/trans.gif") %>"></td>
                      </tr>
				<% end if %>
                      <tr> 
                        <td colspan="2" align="left">&nbsp;<%=DisplayPhrase(phraseDictionary,"Totalnumberoftransactions")%>: <%=totalNumTransactions%></td>
						<td colspan="3" class="right"><%=DisplayPhrase(phraseDictionary,"Totalamount")%>: <%=FormatNumber(totalAmount*.01,2)%></b></td>
						<td colspan="7">&nbsp;</td>
                      </tr>
<%
	else		''No Trans in range
%>	
                      <tr> 
                        <td colspan=12>&nbsp;<%=DisplayPhrase(phraseDictionary,"Notransactionsinselecteddaterange")%></td>
                      </tr>
				<% if NOT request.form("frmExpReport")="true" then  %>
                      <tr style="background-color:<%=session("pageColor4")%>;"> 
                        <td colspan=12><img height="1" width="100%" src="<%= contentUrl("/asp/adm/images/trans.gif") %>"></td>
                      </tr>
				<% end if %>
<%	
	end if
    
	    rsEntry.Close
	    Set rsEntry = Nothing
%>
                  </table>
<%end if	'frmGenReport%>				  
				<% if NOT request.form("frmExpReport")="true" then  %>
                </TD>
              </TR>
            </TABLE>
                  </td>
                </tr>
              </table>
            </td>
          </tr>
          
                   </form>
        </table>
    </td>
    </tr>
		</table>
<% pageEnd %>
<!-- #include file="post.asp" -->

<%
	end if

	end if
%>
