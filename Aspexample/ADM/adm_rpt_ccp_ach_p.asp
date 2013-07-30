<%@ CodePage=65001 %>
<% Server.ScriptTimeout = 600 %>
<%
'dim SessionFarm
'set SessionFarm = Server.CreateObject("SessionFarm.SFSession")
response.charset="utf-8"
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<%






'if Session("curSession") <> Session.SessionID then
if Session("StudioID") = "" then
%>
<script type="text/javascript">
	parent.resetSession();
</script>
<%
else
%>
		<!-- #include file="../inc_dbconn.asp" -->
		<!-- #include file="inc_accpriv.asp" -->
<%
if not Session("Pass") OR Session("Admin")="false" OR NOT validAccessPriv("RPT_CCP") then 
%>
<script type="text/javascript">
	alert("You are not authorized to view this page.\nPlease login to an account with Access Privileges granted for this feature.");
	javascript:history.go(-1);
</script>
<%
else
%>
		<!-- #include file="inc_chk_ss.asp" -->
		<!-- #include file="../inc_i18n.asp" -->
		<!-- #include file="inc_crypt.asp" -->
<%
	' Log folder 
	' Create the folder in your web and give read/write permissions to COM user
	const cLogFolder = "/ftp"

  Private Function IsLocalFolderWriteable(sFolder)	
	  Dim bRet : bRet = false
	  Dim fso
	  Set fso = CreateObject("Scripting.FileSystemObject")
	  Dim sTempFileName : sTempFileName = fso.GetTempName()
		Dim sFile
		sFile = sFolder & "\" & sTempFileName
		Dim oFile
		On Error Resume Next
		Set oFile = fso.CreateTextFile(sFile, True)
		If Err = 0 Then
			oFile.Close
			fso.DeleteFile(sFile)		
			bRet = true
		End If
		
		IsLocalFolderWriteable = bRet
	End Function


	dim testEnvironment, writeFTPLogFile, curDateJulianFmt, curDatei18nFmt, curTime, filename, fs, of, payDate, payDateJulianFmt, payDatei18nFmt, submitAttempts, attemptNum, ss_ccTestMode
	testEnvironment = false
	writeFTPLogFile = false

	ss_ccTestMode = checkStudioSetting("Studios","ccTestMode")
	
	function createOutputFile(batchNum)
		if Session("CCProcessor")="MON" OR Session("CCProcessor")="OP" OR (implementationSwitchIsEnabled("BluefinCanada") AND (Session("CCProcessor")="PMN" AND Session("CCProcessor2")="ELV")) then
			if testEnvironment OR ss_ccTestMode then
				filename = "RBC_" & session("studioID") & "_" & curDateJulianFmt & "_TEST" & ".txt"
			else
				filename = "RBC_" & session("studioID") & "_" & curDateJulianFmt & "_" & Right(batchNum, 4) & ".txt"
			end if            
		elseif Session("CCProcessor")="HSBC" then
			if request.form("optBatchTransType")="BANK" then
				filename = "DDA_" & session("studioID") & "_" & curDateJulianFmt & "_" & Right(batchNum, 4) & ".APC"
			elseif request.form("optBatchTransType")="VISA" then
			    if ss_CountryCode="TW"	then	'VISA/MC for Taiwan
			        filename = MerchantID & "." & Right(Year(payDate),2) & padZeros(Month(payDate),2) & padZeros(Day(payDate),2) & "." & Right(batchNum, 2) & ".in"
			    else
				    'filename = "VMC_" & session("studioID") & "_" & curDateJulianFmt & "_" & Right(batchNum, 4) & ".txt"
				    filename = "BA" & VISABatchAuthID
				end if
			else	'AMEX
				filename = "AMEX_" & session("studioID") & "_" & curDateJulianFmt & "_" & Right(batchNum, 4) & ".txt"
			end if
		end if
		
		'response.Write filename
		'response.end
		
		set fs = CreateObject("Scripting.FileSystemObject")
        set of = fs.CreateTextFile(studio_path & session("studioShort") & "\" & filename, true)
	end function

	function addLineToOutputFile(outStr)
		outStr = Replace(outStr, "&nbsp;", " ")
		if ss_Country="TW" AND session("ccProcessor")="HSBC" AND request.form("optBatchTransType")="VISA" then
		    of.write(outStr)
		else
		    of.writeLine(outStr)
		end if
		
        
		'Print output
		if ss_Country="TW" AND session("ccProcessor")="HSBC" AND request.form("optBatchTransType")="VISA" then
		    response.write outStr
		else
		    response.write outStr & "<br />" & VbCrLf
		end if
		if testEnvironment then
			response.write "length of line: " & LEN(Replace(outStr,"&nbsp;", " ")) & "<br />"
		end if
	end function

	function HTMLSpace(numSpaces)
		HTMLSpace = ""
		for i=1 to numSpaces
			HTMLSpace = HTMLSpace & "&nbsp;"
		next
	end function

	function padZeros(val, NumDigits)
		for i=1 to NumDigits-Len(val)
			padZeros = padZeros & "0"
		next
		padZeros = padZeros & val
	end function
	
	function padTS(val, NumDigits)
		for i=1 to NumDigits-Len(val)
			padTS = padTS & "&nbsp;"
		next
		padTS = val & padTS
	end function

	function padLS(val, NumDigits)
		for i=1 to NumDigits-Len(val)
			padLS = padLS & "&nbsp;"
		next
		padLS = padLS & val
	end function 

	'YYYYDDD - DDD is 1-365
	curDateJulianFmt = Year(DateAdd("n", Session("tzOffset"),Now)) & padZeros(DateDiff("y", "1/1/" & Year(DateAdd("n", Session("tzOffset"),Now)), DateAdd("n", Session("tzOffset"),Now)) + 1 , 3)
	'YYYYMMDD
	curDatei18nFmt = Year(DateAdd("n", Session("tzOffset"),Now)) & padZeros(Month(DateAdd("n", Session("tzOffset"),Now)), 2) & padZeros(Day(DateAdd("n", Session("tzOffset"),Now)), 2)
	curTime = DateAdd("n", Session("tzOffset"),Now)
	
	''Form Input
	dim csDate, cEDate, disMode
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
	if request.form("requiredtxtDateStart")<>"" then
		Call SetLocale(session("mvarLocaleStr"))
			payDate = CDATE(request.form("requiredtxtPaymentDate"))
		Call SetLocale("en-us")
	end if

	payDateJulianFmt = Year(payDate) & padZeros(DateDiff("y", "1/1/" & Year(payDate), payDate) + 1 , 3)
	payDatei18nFmt = Year(payDate) & padZeros(Month(payDate), 2) & padZeros(Day(payDate), 2)
	
	if request.form("optDate")="all" then
		disMode = "all"
	else
		disMode = "range"
	end if
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

	''Generate File	
	dim rsEntry, strFileOut, tmpStrOut, tmpStr, first, ftpUser, ftpPwd, ftpHeader, bankClientID, numRecords, intCounter, VISABatchAuthID, AMEXBatchSubmitterID, AMEXTerminalID, ss_CountryCode, tmpMerchantID,lastBatch,activeBatch,batchWait
	set rsEntry = Server.CreateObject("ADODB.Recordset")

	batchWait = 5
	strSQL = "SELECT ActiveBatch, LastBatch FROM tblCCOpts WHERE tblCCOpts.StudioID=" & session("StudioID")
	rsEntry.CursorLocation = 3
	rsEntry.open strSQL, cnWS
	Set rsEntry.ActiveConnection = Nothing
		lastBatch = rsEntry("LastBatch")
		activeBatch = rsEntry("ActiveBatch")
	rsEntry.close

	if (DateDiff("n",LastBatch,Now)<batchWait AND activeBatch) then
	%>
	<script type="text/javascript">
		document.write("Redirecting...");
		alert("Batching in progress.\n\nPlease try again at \n <%=DateAdd("n",batchWait,LastBatch) %>");
		javascript: history.go(-1);
	</script>
	<%
	response.end
	end if

	strSQL = "SELECT Studios.CountryCode FROM Studios"
	rsEntry.CursorLocation = 3
	rsEntry.open strSQL, cnWS
	Set rsEntry.ActiveConnection = Nothing
	if not rsEntry.EOF then
		ss_CountryCode = rsEntry("CountryCode")
	end if
	rsEntry.close

	strSQL = "UPDATE tblCCOpts SET LastBatch=" & DateSep & Now & DateSep &", ActiveBatch=1"
	cnWS.execute strSQL
%>
<html>
<head>
<title>MINDBODY Online&#8482;</title>
<meta http-equiv="Content-Type" content="text/html">
<!-- #include file="frame_bottom.asp" -->
<%= js(array("mb", "MBS")) %>

<!-- #include file="../inc_date_ctrl.asp" -->

</head>
<body>
<% pageStart %>
<table height="100%" width="<%=strPageWidth%>" cellspacing="0">    
<tr> 
     <td class="center-ch" valign="top" height="100%" width="100%">
        <table cellspacing="0" width="90%" height="100%">
          <tr> 
            <td class="headText" align="left" valign="top"> 
              <table class="mainText" width="100%" cellspacing="0">
                <tr> 
                  <td class="headText"><b>Send to Bank Results</b></td>
                  <td class="right" valign="top"> 
                    <table class="mainText border4" cellspacing="0">
                    </table>
                  </td>
                </tr>
              </table>
            </td>
          </tr>
          <tr> 
            <td valign="bottom" class="mainText right" height="18"></td>
          </tr>
          <tr> 
            <td valign="top" class="mainTextBig center-ch"> <br />
              <table class="mainText" cellspacing="0" width="90%">
              <tr> 
                <td valign="top" align="left"> 
<%
	first = true
	strFileOut = ""
	tmpStrOut = ""
	tmpStr = ""
	intCounter = 1
	totNumTrans = 0
	totAmt = 0
	pretotAmt = 0
	
	nextBatchFileNum = 1000
	strSQL = "SELECT MAX(BatchFileNum) AS MaxBatchNum FROM tblCCTrans"
	rsEntry.CursorLocation = 3
	rsEntry.open strSQL, cnWS
	Set rsEntry.ActiveConnection = Nothing
	if NOT rsEntry.EOF then
		if NOT isNULL(rsEntry("MaxBatchNum")) then
			nextBatchFileNum = rsEntry("MaxBatchNum") + 1
		end if
	end if
	rsEntry.close
	
	strSQL = "SELECT tblCCTrans.SaleID, tblCCTrans.Status, tblCCTrans.Settled, tblCCTrans.ccAmt, tblCCTrans.TransactionNumber, tblCCTrans.AuthCode, tblCCTrans.TransTime, tblCCTrans.OrderID, tblCCTrans.ClientID, CLIENTS.LastName, CLIENTS.FirstName, CLIENTS.RSSID, tblCCTrans.ACHName, tblCCTrans.ACHAccountNum, IsNull(tblCCTrans.ACHRoutingNum, 0) AS ACHRoutingNum, tblCCTrans.IsSavingsAcct, tblCCTrans.MerchantID, tblCCTrans.TerminalID, tblCCTrans.ccNum, tblCCTrans.ExpMoYr, tblCCTrans.AuthTime, Location.FTPHeaderRecord, Location.FTPPassword, Location.FTPUsername, Location.BankClientID, Location.OP_AcctNum, tblCCOpts.VISABatchAuthID, tblCCOpts.AMEXBatchSubmitterID, tblCCOpts.ACHAccountNumber, tblCCTrans.Cardholder "
	strSQL = strSQL & "FROM tblCCTrans INNER JOIN CLIENTS ON tblCCTrans.ClientID = CLIENTS.ClientID INNER JOIN Location ON tblCCTrans.LocationID = Location.LocationID LEFT OUTER JOIN Sales ON tblCCTrans.SaleID = Sales.SaleID CROSS JOIN tblCCOpts "
	strSQL = strSQL & "WHERE (tblCCOpts.StudioID = " & session("studioID") & ") AND (tblCCTrans.Settled = 0) AND (tblCCTrans.Status=N'Pending') "
	'CB 47_225 - Support for PAP/DDA ie BANK, VISA & AMEX
	if request.form("optBatchTransType")="BANK" then
		strSQL = strSQL & " AND (NOT (tblCCTrans.ACHName IS NULL)) "
	elseif request.form("optBatchTransType")="VISA" then
		strSQL = strSQL & "AND (tblCCTrans.ccType=N'Visa' OR tblCCTrans.ccType=N'Master Card') "
	elseif request.form("optBatchTransType")="AMEX" then
		strSQL = strSQL & "AND (tblCCTrans.ccType=N'American Express') "
	end if				
	if disMode = "range" then
		strSQL = strSQL & " AND TransTime >= " & DateSep & cSDate & DateSep
		strSQL = strSQL & " AND TransTime <= " & DateSep & DateAdd("d", 2, cEDate) & DateSep
	end if
	if ccLoc<>"-2" then
		strSQL = strSQL & " AND (tblCCTrans.LocationID = " & ccLoc & ") "
	end if
	if request.form("optBatchFileNum")<>"" AND request.form("optBatchFileNum")<>"0" then
		strSQL = strSQL & " AND (tblCCTrans.BatchFileNum = " & request.form("optBatchFileNum") & ") "
	end if
	strSQL = strSQL & " ORDER BY tblCCTrans.TransTime DESC, TransactionNumber DESC"
	response.write debugSQL(strSQL, "SQL")
	rsEntry.CursorLocation = 3
	rsEntry.open strSQL, cnWS
	Set rsEntry.ActiveConnection = Nothing

	''****************************** GENERATE FILE ********************************************''
	if NOT rsEntry.EOF then

		numRecords = 0

		''PASS 1 GET GLOBAL VALS / DOLLAR HEADER / RECORD COUNT
		do while NOT rsEntry.EOF
			if request.form("chk_"&rsEntry("TransactionNumber"))="on" then
				if first then
					first = false
					ftpUser = TRIM(rsEntry("FTPUsername"))
					ftpPwd = TRIM(rsEntry("FTPPassword"))
					ftpHeader = TRIM(rsEntry("FTPHeaderRecord"))
					bankClientID = TRIM(rsEntry("BankClientID"))
					VISABatchAuthID = rsEntry("VISABatchAuthID")
					AMEXBatchSubmitterID = rsEntry("AMEXBatchSubmitterID")
					MerchantID = TRIM(rsEntry("MerchantID"))
					'TT - Work Item 7387 - Add dim to hold AMEX Batch ID for TW and HBSC
					AMEXTerminalID = TRIM(rsEntry("TerminalID"))
				end if
				numRecords = numRecords + 1
				pretotAmt = pretotAmt + rsEntry("ccAmt")
			end if 	'transaction checked to send
			rsEntry.MoveNext
		loop
		''END PASS 1

		''PASS 2 
		rsEntry.MoveFirst

		'Create Output File
		createOutputFile(nextBatchFileNum)
        
		'================================================================================================
		'								BEGIN - ROYAL BANK OF CANADA STD152 FORMAT
		'================================================================================================
		if (Session("CCProcessor")="MON" OR Session("CCProcessor")="OP") OR ((Session("CCProcessor")="PMN" AND Session("CCProcessor2")="ELV") AND implementationSwitchIsEnabled("BluefinCanada")) then

			'Add $$$ Header		
			addLineToOutPutFile(ftpHeader)

			''*************************** BEGIN SETUP HEADER RECORD *****************************''
			'Header Field # 1 Record Count - pad w/zeros
			tmpStrOut = padZeros(intCounter,6)
	
			'Header Field # 2 Record Type use "A" for leading record
			tmpStrOut = tmpStrOut & "A"
	
			'Header Field # 3 Transaction Code - "HDR"
			tmpStrOut = tmpStrOut & "HDR"
	
			'Header Field # 4 Client Number 10 digits should already have last 4 digits padded w/zero's
			tmpStrOut = tmpStrOut & bankClientID
	
			'Header Field # 5 Client Name 30 chars A-Z and 0-9 ONLY
			'First stip off any non alpha numeric characters
			tmpStr = UCASE(TRIM(session("StudioName")))
			Dim regEx
			Set regEx = New RegExp
			regEx.Global = true
			regEx.Pattern = "[^0-9a-zA-Z]"
			tmpStr = regEx.Replace(tmpStr, "")
			'Add to output pad after with spaces
			tmpStrOut = tmpStrOut & LEFT(tmpStr,30)
			if LEN(tmpStr)<30 then
				tmpStrOut = tmpStrOut & HTMLSpace(30-LEN(tmpStr))
			end if
	
			'Header Field # 6 File Creation Number 4 digits "TEST" for test environment
			if TestEnvironment OR ss_ccTestMode then
				tmpStrOut = tmpStrOut & "TEST"
			else
				tmpStrOut = tmpStrOut & Right(nextBatchFileNum, 4)
			end if
	
			'Header Field # 7 File Creation Date YYYYDDD
			tmpStrOut = tmpStrOut & curDateJulianFmt
	
			'Header Field # 8 Currency Type CAD OR USD
			tmpStrOut = tmpStrOut & "CAD"
	
			'Header Field # 9 input type - "1"
			tmpStrOut = tmpStrOut & "1"
	
			'Header Field # 10-15	 blanks, total 84
			tmpStrOut = tmpStrOut & HTMLSpace(86)
	
			'Header Field # 16 Client Optional Record "Y" or "N"
			tmpStrOut = tmpStrOut & "N"
	
			addLineToOutPutFile(tmpStrOut)
			''*************************** END SETUP HEADER RECORD ********************************''

			do while NOT rsEntry.EOF
				if request.form("chk_"&rsEntry("TransactionNumber"))="on" then
	
					''*************************** BEGIN BASIC PAYMENT RECORDS ****************************''
					'Payment Record Field # 1 Record Count - increment
					intCounter = intCounter + 1
					tmpStrOut = padZeros(intCounter,6)
	
					'Payment Record Field # 2 Record Type - "D"
					tmpStrOut = tmpStrOut & "D"
	
					'Payment Record Field # 2 Transaction Code - 3 chars, leave blank if destined for CAN
					tmpStrOut = tmpStrOut & HTMLSpace(3)
	
					'Payment Record Field # 4 Client Number 10 digits should already have last 4 digits padded w/zero's
					tmpStrOut = tmpStrOut & bankClientID
	
					'Payment Record Field # 5 Filling 1 blank
					tmpStrOut = tmpStrOut & HTMLSpace(1)
	
                    tmpStr = LEFT(TRIM(rsEntry("ClientID")),19)
					'Payment Record Field # 6 Customer Number - 19 Chars (USE RSSID)
                    if NOT isNULL(rsEntry("RSSID")) then
                        if TRIM(rsEntry("RSSID"))<>"" then
                            tmpStr = LEFT(TRIM(rsEntry("RSSID")),19)
                        end if
                    end if
                    'Bug 2434 - CCP 12/31/09
				    tmpStrOut = tmpStrOut & tmpStr & HTMLSpace(19-LEN(tmpstr))
	
					'Payment Record Field # 7 Payment Number 2 digits
					if LEN(rsEntry("TransactionNumber"))<2 then	'if single digit transaciton number pad with leading zero
						tmpStrOut = tmpStrOut & "0"
					end if
					tmpStrOut = tmpStrOut & RIGHT(rsEntry("TransactionNumber"),2)
	
					'Payment Record Field # 8/9 Routing Number 9 digits, might need leading zero
					if LEN(TRIM(rsEntry("ACHRoutingNum")))=8 then	'reverse order add in leading 0
						tmpStrOut = tmpStrOut & "0" & RIGHT(TRIM(rsEntry("ACHRoutingNum")), 3) & LEFT(TRIM(rsEntry("ACHRoutingNum")), 5)
					elseif LEN(TRIM(rsEntry("ACHRoutingNum")))=9 then	'leading zero/value entered
						tmpStrOut = tmpStrOut & RIGHT(TRIM(rsEntry("ACHRoutingNum")), 4) & LEFT(TRIM(rsEntry("ACHRoutingNum")), 5)
					else	'send what was entered... CB 12/11/08 Padd with Zeros
						tmpStrOut = tmpStrOut & padZeros(RIGHT(rsEntry("ACHRoutingNum"),9), 9)
					end if
	
					'Payment Record Field # 10 Account Number up to 18 digits right fill blanks
					tmpStr = rsEntry("ACHAccountNum")
					tmpStr = DES_Decrypt(tmpStr,true,null)
					tmpStrOut = tmpStrOut & tmpStr
					if LEN(tmpStr)<18 then
						tmpStrOut = tmpStrOut & HTMLSpace(18-LEN(tmpStr))
					end if
	
					'Payment Record Field # 11 Filling 1 blank
					tmpStrOut = tmpStrOut & HTMLSpace(1)
					
					'Payment Record Field # 10 Payment Amount 12 digits format $$$$$$cc w/leading blanks
					tmpStrOut = tmpStrOut & padZeros(rsEntry("ccAmt"), 10)
	
					'Payment Record Field # 13 Reserved 6 blanks
					tmpStrOut = tmpStrOut & HTMLSpace(6)
					
					'Payment Record Field # 14 Payment Date - YYYYDDD
					'tmpStrOut = tmpStrOut & curDateJulianFmt
					tmpStrOut = tmpStrOut & payDateJulianFmt
					
					'Payment Record Field # 15 Customer Name 30 Chars
					tmpStrOut = tmpStrOut & LEFT(TRIM(rsEntry("ACHName")),30)
					if LEN(TRIM(rsEntry("ACHName")))<30 then
						tmpStrOut = tmpStrOut & HTMLSpace(30-LEN(TRIM(rsEntry("ACHName"))))
					end if
	
					'Payment Record Field # 16 Language Code "E" English, "F" French
					tmpStrOut = tmpStrOut & "E"
	
					'Payment Record Field # 17 Payment Medium / Payment Route Method "E" electronic... can be left blank to get default
					tmpStrOut = tmpStrOut & "E"
	
					'Payment Record Field # 18 Client Short Name - 15 chars.. can leave blank for default on profile, which makes sense... they can pick 15 chars better than stripping Studio Name
					tmpStrOut = tmpStrOut & HTMLSpace(15)
	
					'Payment Record Field # 19 Destination Currency "CAD" or "USD"
					tmpStrOut = tmpStrOut & "CAD"
					
					'Payment Record Field # 20 Reserved blank
					tmpStrOut = tmpStrOut & HTMLSpace(1)
	
					'Payment Record Field # 21 Destination Country "CAN" or "USA"
					tmpStrOut = tmpStrOut & "CAN"
	
					'Payment Record Field # 22/23 Filler/Reserved 4 blanks
					tmpStrOut = tmpStrOut & HTMLSpace(4)
	
					'Payment Record Field # 24 Optional Record Indicator "Y" or "N"
					tmpStrOut = tmpStrOut & "N"
	
					addLineToOutPutFile(tmpStrOut)
	
					totNumTrans = totNumTrans + 1
					totAmt = totAmt + rsEntry("ccAmt")
					''*************************** END BASIC PAYMENT RECORDS ******************************''
	
				end if	'checked
				rsEntry.MoveNext
			loop
			''END PASS 2
			
			''*************************** BEGIN TRAILER RECORD ********************************''
			'Trailer Record Field # 1 Record Count - increment
			intCounter = intCounter + 1
			tmpStrOut = padZeros(intCounter,6)
	
			'Trailer Record Field # 2 Record Type - "Z"
			tmpStrOut = tmpStrOut & "Z"
	
			'Trailer Record Field # 2 Transaction Code - 3 chars, leave blank if destined for CAN
			tmpStrOut = tmpStrOut & "TRL"
	
			'Trailer Record Field # 4 Client Number 10 digits should already have last 4 digits padded w/zero's
			tmpStrOut = tmpStrOut & bankClientID
	
			'Trailer Record Field # 5/6 Reserved 20 blanks
			tmpStrOut = tmpStrOut & HTMLSpace(20)
	
			'Trailer Record Field # 7 Total Number of Transactions 6 digi's leading zeros
			tmpStrOut = tmpStrOut & padZeros(totNumTrans, 6)
	
			'Trailer Record Field # 8 Total Amount of Transactions 14 digi's leading spaces
			tmpStrOut = tmpStrOut & padZeros(totAmt, 14) 
	
			'Trailer Record Field # 9 Total Number of Option Client Records - 2 digits "00"
			tmpStrOut = tmpStrOut & "00"
	
			'Trailer Record Field # 10 Total Number of Option Customer Records - 6 digits "000000"
			tmpStrOut = tmpStrOut & "000000"
	
			'Trailer Record Field # 11-15 Reserved/Filler 84 blanks
			tmpStrOut = tmpStrOut & HTMLSpace(84)
	
			addLineToOutPutFile(tmpStrOut)
			''*************************** END TRAILER RECORD **********************************''

		'================================================================================================
		'								END - ROYAL BANK OF CANADA STD152 FORMAT
		'================================================================================================





		'================================================================================================
		'								BEGIN - HSBC MULTIPLE FORMATS
		'================================================================================================
		elseif session("ccProcessor")="HSBC" then
		
			'****************** BEGIN - HSBC BA FILE 172 FORMAT *****************************************
			if request.form("optBatchTransType")="BANK" then

				'---------------- BEGIN HEADER --------------------------------
				tmpStrOut = "G"                         'Autoplan Code - G autoPay-In (autoplan 1)
				'tmpStrOut = tmpStrOut & HTMLSpace(12)	'Account Number 12 digits auto filled
				tmpStrOut = tmpStrOut & padTS(rsEntry("ACHAccountNumber"),12)	'Account Number 12 digits
				tmpStrOut = tmpStrOut & HTMLSpace(3)	'Payment Code blank in PURE example
				tmpStrOut = tmpStrOut & Left(MonthName(Month(payDate)),3) & HTMLSpace(1) & padZeros(Day(payDate),2) & HTMLSpace(6)	'Reference 12
				tmpStrOut = tmpStrOut & padZeros(Day(payDate), 2) & padZeros(Month(payDate), 2) & Right(Year(payDate), 2)            'Value Date - DDMMYY
				tmpStrOut = tmpStrOut & "K"             'Input Medium - K - diskette
				tmpStrOut = tmpStrOut & "********"      'File Name (8)
				tmpStrOut = tmpStrOut & padZeros(numRecords, 5)      'Number of Records (5)
				tmpStrOut = tmpStrOut & padZeros(pretotAmt, 10)      'Amount of Records (10)
				tmpStrOut = tmpStrOut & HTMLSpace(7)	'Overflow count
				tmpStrOut = tmpStrOut & HTMLSpace(12)	'Overflow amount
				tmpStrOut = tmpStrOut & HTMLSpace(2)	'Unused
				tmpStrOut = tmpStrOut & "1"	            'Centre Code - '1'1
				'Header Total - 80
				'addLineToOutPutFile(tmpStrOut)
				'---------------- END HEADER --------------------------------

				do while NOT rsEntry.EOF
					if request.form("chk_"&rsEntry("TransactionNumber"))="on" AND rsEntry("ccAmt")<>0 then

						'---------------- BEGIN DETAIL  -----------------------------
						intCounter = intCounter + 1

						tmpStrOut = tmpStrOut & HTMLSpace(1)	'Filler - 1 Space
												
						tmpStrOut = tmpStrOut & LEFT(TRIM(rsEntry("RSSID")),12)                     'Particular /Debtor Reference - 12 OR LESS Chars (USE RSSID)
						if LEN(TRIM(rsEntry("RSSID")))<12 then
							tmpStrOut = tmpStrOut & HTMLSpace(12-LEN(TRIM(rsEntry("RSSID"))))
						end if

						tmpStrOut = tmpStrOut & LEFT(TRIM(rsEntry("ACHName")),20)                   'Account Name
						if LEN(TRIM(rsEntry("ACHName")))<20 then
							tmpStrOut = tmpStrOut & HTMLSpace(20-LEN(TRIM(rsEntry("ACHName"))))
						end if

						'tmpStrOut = tmpStrOut & TRIM(rsEntry("ACHRoutingNum"))                      '
						'if LEN(TRIM(rsEntry("ACHRoutingNum")))<6 then
						'	tmpStrOut = tmpStrOut & HTMLSpace(6-LEN(TRIM(rsEntry("ACHRoutingNum"))))
						'end if
						
						tmpStr = rsEntry("ACHAccountNum")                                           'Bank Number (3) Branch Number (3) Bank Account Number (9)
						tmpStr = TRIM(rsEntry("ACHRoutingNum")) & DES_Decrypt(tmpStr,true,null)
						tmpStrOut = tmpStrOut & tmpStr
						if LEN(tmpStr)<15 then
							tmpStrOut = tmpStrOut & HTMLSpace(15-LEN(tmpStr))
						end if
			
						tmpStrOut = tmpStrOut & padZeros(rsEntry("ccAmt"), 10)                      'Amount 8 Digits + 2 Decimals
						tmpStrOut = tmpStrOut & HTMLSpace(4)	                                    'Value Date - Blank for AutoPlan1
						tmpStrOut = tmpStrOut & HTMLSpace(6)	                                    'Continuation of Second Party Identifier - Spaces if not used

						tmpStrOut = tmpStrOut & RIGHT(TRIM(rsEntry("TransactionNumber")),12)        'Second Party Reference	 - 12 or less
						if LEN(TRIM(rsEntry("TransactionNumber")))<12 then
							tmpStrOut = tmpStrOut & HTMLSpace(12-LEN(TRIM(rsEntry("TransactionNumber"))))
						end if

        				'addLineToOutPutFile(tmpStrOut)

        				'Detail Total - 80
						totNumTrans = totNumTrans + 1
						totAmt = totAmt + rsEntry("ccAmt")
						'---------------- END DETAIL  -----------------------------

					end if	'checked
					rsEntry.MoveNext
				loop

                'no line feeds add all at the end
   				'addLineToOutPutFile(tmpStrOut)
   				'CB Updated to write output without line feed
   				tmpStrOut = Replace(tmpStrOut, "&nbsp;", " ")
   				of.write(tmpStrOut)


			'****************** BEGIN - HSBC VISA GLOBAL PAYMENTS 253 FORMAT ****************************
			elseif request.form("optBatchTransType")="VISA" then
			    
			    if ss_CountryCode="TW"	then	'VISA/MC for Taiwan
			        
    				
				    '---------------- BEGIN HEADER --------------------------------
				    tmpStrOut = "H"	'RECORD IDENTIFIER 1, place
				    tmpStrOut = tmpStrOut & rsEntry("TerminalID")	'TERMINAL ID, 8 places
				    tmpStrOut = tmpStrOut & padTS(rsEntry("MerchantID"),15) 'MERCHANT ID, 15 places
				    tmpStrOut = tmpStrOut & "T" 'CURRENCY CODE, 'T' for TWD or 'U' for USD, 1 place
				    tmpStrOut = tmpStrOut & padZeros(numRecords,6) 'TOTAL COUNT, 6 places
				    tmpStrOut = tmpStrOut & padZeros("0",6) 'AUTH ONLY TOTAL COUNT, 6 places
				    tmpStrOut = tmpStrOut & padZeros("0",12) 'AUTH ONLY TOTAL AMOUNT, 12 places
				    tmpStrOut = tmpStrOut & padZeros(numRecords,6) 'AUTH & SETTLEMENT ONLY TOTAL COUNT, assumming this is the same as total count, 6 places
				    tmpStrOut = tmpStrOut & padZeros(preTotAmt,12) 'AUTH & SETTLEMENT ONLY TOTAL AMOUNT, assumming this is the same as total amount, 12 places
				    tmpStrOut = tmpStrOut & padZeros("0",6) 'OFFLINE TOTAL COUNT, 6 places
				    tmpStrOut = tmpStrOut & padZeros("0",12) 'OFFLINE TOTAL AMOUNT, 12 places
				    tmpStrOut = tmpStrOut & padZeros("0",6) 'REFUND TOTAL COUNT, 6 places
				    tmpStrOut = tmpStrOut & padZeros("0",12) 'REFUND ONLY TOTAL AMOUNT, 12 places
				    tmpStrOut = tmpStrOut & HTMLSpace(151)  ' FILLER 151 places
				    'tmpStrOut = tmpStrOut & HTMLSpace(2) ' 2 places
				    'Header Total - 256
				    addLineToOutPutFile(tmpStrOut)
				    '---------------- END HEADER --------------------------------
			        
			        do while NOT rsEntry.EOF
					    if request.form("chk_"&rsEntry("TransactionNumber"))="on" AND rsEntry("ccAmt")<>0 then

						    '---------------- BEGIN DETAIL  -----------------------------
						    
    						
    						tmpStrOut = "D"	'RECORD IDENTIFIER, 1 place
    						tmpStrOut = tmpStrOut & padZeros(intCounter,6) 'SEQUENCE NUMBER, 6 places
    						tmpStrOut = tmpStrOut & rsEntry("TerminalID")	'TERMINAL ID, 8 places
				            tmpStrOut = tmpStrOut & padTS(rsEntry("MerchantID"),15) 'MERCHANT ID, 15 places
				            tmpStrOut = tmpStrOut & Right(Year(payDate),2) & padZeros(Month(payDate),2) & padZeros(Day(payDate),2)  	'Transaction Date - YYMMDD, 6 places
				            tmpStrOut = tmpStrOut & "S" 'TRANSACTION TYPE, A: AUTH ONLY, S: AUTH & SETTLEMENT, O: OFFLINE, R: REFUND, 1place
				            'Card Number
						    tmpStr = rsEntry("ccNum")
						    tmpStr = DES_Decrypt(tmpStr,true,null)
						    tmpStrOut = tmpStrOut & padTS(tmpStr, 19) 'CARD NUMBER, 19 places
						    tmpStrOut = tmpStrOut & rsEntry("ExpMoYr")	'Expiry Date MMYY, 4 places
						    tmpStrOut = tmpStrOut & HTMLSpace(3) 'CVC2 FILL SPACE, 3 places
						    tmpStrOut = tmpStrOut & padZeros(rsEntry("ccAmt"), 12) 'AMOUNT, 12 places
						    tmpStrOut = tmpStrOut & "999999" 'APPROVED CODE, 999999 for trans type "S", 6 places
						    tmpStrOut = tmpStrOut & "99" 'RESPONSE CODE, 99 for trans type "S", 2 places
						    tmpStrOut = tmpStrOut & padLS(rsEntry("RSSID"),20)'RECEIPT RESERVED FIELD, 20 places
						    tmpStrOut = tmpStrOut & padTS(rsEntry("TransactionNumber"),20)'MERCHANT RESERVED FIELD, 20 places
						    tmpStrOut = tmpStrOut & HTMLSpace(40)'REFFERAL DATA, 40 places
						    tmpStrOut = tmpStrOut & "0" 'INSTALLMENT FLAG, 1 place
						    tmpStrOut = tmpStrOut & HTMLSpace(10)'PRODUCTION CODE, 10 places
						    tmpStrOut = tmpStrOut & HTMLSpace(3)'INSTALLMENT DURATION,3 places
						    tmpStrOut = tmpStrOut & HTMLSpace(9)'MONTHLY REPAYMENT AMOUNT, 9 places
						    tmpStrOut = tmpStrOut & HTMLSpace(9)'FIRST PAYMENT AMOUNT, 9 places
						    tmpStrOut = tmpStrOut & HTMLSpace(7)'FORMALITY FEE, 7 places
						    tmpStrOut = tmpStrOut & HTMLSpace(16)'CHIP CODE, 16 places
						    tmpStrOut = tmpStrOut & HTMLSpace(36)' RESERVED FILL SPACE, 36 places
						    'tmpStrOut = tmpStrOut & HTMLSpace(2) '2 places
						    'Total 256
						    addLineToOutPutFile(tmpStrOut)
				            
						    intCounter = intCounter + 1 ' counter starts at 1
						    totNumTrans = totNumTrans + 1
						    totAmt = totAmt + rsEntry("ccAmt")
						    '---------------- END DETAIL  -----------------------------

					    end if	'checked
					    rsEntry.MoveNext
				    loop


				    '---------------- BEGIN TRAILER --------------------------------
					    'not defined in spec?
				    '---------------- END TRAILER --------------------------------
			        
			        
			    elseif ss_CountryCode="SG"	then	'VISA/MC for Singapore
						'//fsdfgsdfgsdfg
						'//asdfgasdfg
						'---------------- BEGIN HEADER --------------------------------
				    tmpStrOut = "030"	'Transaction Code - Sale 3 digits
				    tmpStrOut = tmpStrOut & rsEntry("MerchantID")	'Merchant Number 12 digits
				    tmpStrOut = tmpStrOut & HTMLSpace(2)	'Filler - 2 spaces
				    tmpStrOut = tmpStrOut & Right(nextBatchFileNum, 3)	'Batch Number 3 digits
				    tmpStrOut = tmpStrOut & HTMLSpace(7)	'Deposit Control Number 7 spaces
				    tmpStrOut = tmpStrOut & padZeros(pretotAmt, 12)	'Batch Total - 12
				    tmpStrOut = tmpStrOut & padZeros("0",11)	'Discount Amount - 11
				    tmpStrOut = tmpStrOut & "12"	'Transaction Currency 2 Digits - 12 for SGD
				    tmpStrOut = tmpStrOut & Right(Year(DateAdd("n", Session("tzOffset"),Now)),2)	'Microfilm Year
				    tmpStrOut = tmpStrOut & "M"	'Special Transaction Indicator - M for Mail/Phone Order
				    tmpStrOut = tmpStrOut & HTMLSpace(25)	'Filler - 23 spaces
				    'Header Total - 80
				    addLineToOutPutFile(tmpStrOut)
				    '---------------- END HEADER --------------------------------
				    
				    do while NOT rsEntry.EOF
					    if request.form("chk_"&rsEntry("TransactionNumber"))="on" AND rsEntry("ccAmt")<>0 then

						    '---------------- BEGIN DETAIL  -----------------------------
						    intCounter = intCounter + 1
    						
						    tmpStrOut = HTMLSpace(1)	'Filler - 1 Space
						    tmpStrOut = tmpStrOut & HTMLSpace(2)	'Transaction Code - 2 spaces - Sales
						    tmpStr = rsEntry("ccNum") 
						    tmpStr = DES_Decrypt(tmpStr,true,null)
						    tmpStrOut = tmpStrOut & tmpStr
						    if LEN(tmpStr)<16 then
							    tmpStrOut = tmpStrOut & HTMLSpace(16-LEN(tmpStr))
						    end if 'Card Number 16 digits
						    tmpStrOut = tmpStrOut & rsEntry("ExpMoYr")	'Expiry Date MMYY 4 digits
						    tmpStrOut = tmpStrOut & padZeros(rsEntry("ccAmt"), 12)	'Transaction Amount - 12 digits
						    tmpStrOut = tmpStrOut & padZeros(Day(payDate),2) & padZeros(Month(payDate),2) & Right(Year(payDate),2) 	'Transaction Date - DDMMYY 6 digits
						    tmpStrOut = tmpStrOut & HTMLSpace(6)	'Authorization Code - 6 spaces
						    tmpStrOut = tmpStrOut & "000"	'JCB Installment Number - 000 for non-JCB instalment
						    tmpStrOut = tmpStrOut & FmtPadString(rsEntry("TransactionNumber"), 50, " ", false)	'Remarks 50 chars updated to track TransactionNumber
						    tmpStrOut = tmpStrOut & HTMLSpace(20)	'Payer ID - Loinheart merchant only - 20 spaces 
						    tmpStrOut = tmpStrOut & HTMLSpace(120)	'Payer Name - Loinheart merchant only - 120 spaces 
						    tmpStrOut = tmpStrOut & HTMLSpace(20)	'Invoice Number Loinheart merchant only - 20 spaces 
						    tmpStrOut = tmpStrOut & HTMLSpace(13)	'Invoice Amount Loinheart merchant only - 13 spaces
						    'Total 253
						    addLineToOutPutFile(tmpStrOut)
    						
						    totNumTrans = totNumTrans + 1
						    totAmt = totAmt + rsEntry("ccAmt")
						    '---------------- END DETAIL  -----------------------------

					    end if	'checked
					    rsEntry.MoveNext
				    loop
						   
			    else
				    'Spec in HSBC-VSIA_MC.pdf
				    'VISABatchAuthID - used in file naming
    				
				    '---------------- BEGIN HEADER --------------------------------
				    tmpStrOut = "030"	'Transaction Code - Sale
				    tmpStrOut = tmpStrOut & rsEntry("MerchantID")	'Merchant Number 9 digits
				    tmpStrOut = tmpStrOut & "90"	'Outlet Number 2 digits - HARDCODED FOR PURE
				    tmpStrOut = tmpStrOut & "393"	'Reel Number 3 digits   - HARDCODED FOR PURE
				    tmpStrOut = tmpStrOut & Right(nextBatchFileNum, 3)	'Batch Number 3 digits
				    tmpStrOut = tmpStrOut & HTMLSpace(7)	'Deposit Control Number
				    tmpStrOut = tmpStrOut & padZeros(pretotAmt, 12)	'Batch Total - 12
				    tmpStrOut = tmpStrOut & padZeros("0",11)	'Discount Amount - 11
				    tmpStrOut = tmpStrOut & "00"	'Transaction Currency 2 Digits - 00 for HKD
				    tmpStrOut = tmpStrOut & Right(Year(DateAdd("n", Session("tzOffset"),Now)),2)	'Microfilm Year
				    tmpStrOut = tmpStrOut & "M"	'Special Transaction Indicator - M for Mail/Phone Order
				    tmpStrOut = tmpStrOut & HTMLSpace(1)	'Filler - 1 space
				    tmpStrOut = tmpStrOut & HTMLSpace(1)	'Card Type - Space for Visa/MC/JCB
				    tmpStrOut = tmpStrOut & HTMLSpace(23)	'Filler - 23 spaces
				    'Header Total - 80
				    addLineToOutPutFile(tmpStrOut)
				    '---------------- END HEADER --------------------------------

				    do while NOT rsEntry.EOF
					    if request.form("chk_"&rsEntry("TransactionNumber"))="on" AND rsEntry("ccAmt")<>0 then

						    '---------------- BEGIN DETAIL  -----------------------------
						    intCounter = intCounter + 1
    						
						    tmpStrOut = HTMLSpace(1)	'Filler - 1 Space
						    tmpStrOut = tmpStrOut & HTMLSpace(2)	'Transaction Code - Space - Sales
						    'Card Number
						    tmpStr = rsEntry("ccNum")
						    tmpStr = DES_Decrypt(tmpStr,true,null)
							tmpStr = LEFT(tmpStr, 16)
						    tmpStrOut = tmpStrOut & tmpStr
						    if LEN(tmpStr)<16 then
							    tmpStrOut = tmpStrOut & HTMLSpace(16-LEN(tmpStr))
						    end if
						    tmpStrOut = tmpStrOut & rsEntry("ExpMoYr")	'Expiry Date MMYY
						    tmpStrOut = tmpStrOut & padZeros(rsEntry("ccAmt"), 12)	'Transaction Amount - 12

						    'tmpStrOut = tmpStrOut & padZeros(Day(rsEntry("AuthTime")),2) & padZeros(Month(rsEntry("AuthTime")),2) & Right(Year(rsEntry("AuthTime")),2) 	'Transaction Date - DDMMYY
						    tmpStrOut = tmpStrOut & padZeros(Day(payDate),2) & padZeros(Month(payDate),2) & Right(Year(payDate),2) 	'Transaction Date - DDMMYY
						    tmpStrOut = tmpStrOut & HTMLSpace(6)	'Authorization Code - 6 spaces
						    tmpStrOut = tmpStrOut & HTMLSpace(1)	'Filler - 1 space
    						
						    'tmpStrOut = tmpStrOut & FmtPadString(rsEntry("RSSID"), 15, " ", false)	'Remarks - 15 use RSSID padded with spaces on the right
						    tmpStrOut = tmpStrOut & FmtPadString(rsEntry("TransactionNumber"), 15, " ", false)	'Remarks - 15 updated to track TransactionNumber
    						
						    tmpStrOut = tmpStrOut & "000"	'JCB Installment Number - 000 for non-JCB instalment
						    tmpStrOut = tmpStrOut & HTMLSpace(4)	'Telephone ID
						    tmpStrOut = tmpStrOut & HTMLSpace(6)	'Telephone Ducation
						    tmpStrOut = tmpStrOut & HTMLSpace(4)	'Filler - 4 spaces
						    tmpStrOut = tmpStrOut & HTMLSpace(173)	'Loinheart merchant only 
						    'Total 253
						    addLineToOutPutFile(tmpStrOut)
    						
						    totNumTrans = totNumTrans + 1
						    totAmt = totAmt + rsEntry("ccAmt")
						    '---------------- END DETAIL  -----------------------------

					    end if	'checked
					    rsEntry.MoveNext
				    loop


				    '---------------- BEGIN TRAILER --------------------------------
					    'not defined in spec?
				    '---------------- END TRAILER --------------------------------
				end if

			'****************** BEGIN- HSBC AMEX BAS FORMAT *********************************************
			elseif request.form("optBatchTransType")="AMEX" then

				if ss_CountryCode="TW" OR ss_CountryCode="SG"	then	'AMEX for Tawain
					'---------------- BEGIN FILE HEADER --------------------------------
					tmpStrOut = "01"	'Record Type
					tmpStrOut = tmpStrOut & padZeros(intCounter,6)		'Sequence Number
					tmpStrOut = tmpStrOut & HTMLSpace(24)	'Contact Name
					tmpStrOut = tmpStrOut & HTMLSpace(24)	'Contact Address Line 1
					tmpStrOut = tmpStrOut & HTMLSpace(24)	'Contact Address Line 2
					tmpStrOut = tmpStrOut & HTMLSpace(24)	'Contact Address Line 3
					tmpStrOut = tmpStrOut & HTMLSpace(10)	'Contact Telephone
					'tmpStrOut = tmpStrOut & AMEXBatchSubmitterID	'AMEX Submitter ID
					tmpStrOut = tmpStrOut & AMEXTerminalID	'AMEX Submitter ID 'TT - Work Item 7387 - Changed to AMEXTerminalID to use newly mapped field for TW AMEX
					tmpStrOut = tmpStrOut & padZeros(Day(DateAdd("n", Session("tzOffset"),Now)),2) & padZeros(Month(DateAdd("n", Session("tzOffset"),Now)),2) & Right(Year(DateAdd("n", Session("tzOffset"),Now)),2)	'File Creation Date
					tmpStrOut = tmpStrOut & "P"	'Test Flag, P - Production, T - Test
					tmpStrOut = tmpStrOut & HTMLSpace(223)	'Reserved
	
					'Total - 350
					addLineToOutPutFile(tmpStrOut)
					'---------------- END FILE HEADER --------------------------------
	
					'---------------- BEGIN BATCH HEADER --------------------------------
					intCounter = intCounter + 1
	
					tmpStrOut = "02"	'Record Type
					tmpStrOut = tmpStrOut & padZeros(intCounter,6)		'Sequence Number
					tmpStrOut = tmpStrOut & padZeros(nextBatchFileNum, 6)	'Submission Reference
					'tmpStrOut = tmpStrOut & rsEntry("TerminalID")	'Merchant Number 10 digits
					'tmpStrOut = tmpStrOut & rsEntry("MerchantID")	'Merchant Number 10 digits 'TT - Work Item 7387 - AMEX merchant ID stored in MerchantID for TW AMEX instead of TerminalID
					'tmpStrOut = tmpStrOut & HTMLSpace(30)	'Merchant Name
					tmpStrOut = tmpStrOut & padTS(rsEntry("MerchantID"), 40) 'Merchant Number and Merchant Name stored together 10 + 30 = 40 digits
					tmpStrOut = tmpStrOut & "69"	'Industry Code - Mail/Telephone order/Misellaneous
					tmpStrOut = tmpStrOut & "599"	'Sub Industry Code
					tmpStrOut = tmpStrOut & HTMLSpace(291)	'Reserved
	
					'Total 350
					addLineToOutPutFile(tmpStrOut)
					'---------------- END BATCH HEADER --------------------------------
	
					do while NOT rsEntry.EOF
						if request.form("chk_"&rsEntry("TransactionNumber"))="on" AND rsEntry("ccAmt")<>0 then
	
							'---------------- BEGIN DETAIL  -----------------------------
							intCounter = intCounter + 1
							
							tmpStrOut = "03"	'Record Type
							tmpStrOut = tmpStrOut & padZeros(intCounter,6)		'Sequence Number
							'Card Number - Always 15 digits
							tmpStr = rsEntry("ccNum")
							tmpStr = DES_Decrypt(tmpStr,true,null)
							tmpStrOut = tmpStrOut & tmpStr
							if isNULL(rsEntry("Cardholder")) then 
								tmpStrOut = tmpStrOut & HTMLSpace(26)	'Card Holder Name
							else
								tmpStrOut = tmpStrOut & LEFT(rsEntry("Cardholder"),26)	'Card Holder Name
								if LEN(rsEntry("Cardholder"))<26 then
									tmpStrOut = tmpStrOut & HTMLSpace(26-LEN(rsEntry("Cardholder")))
								end if
							end if
							tmpStrOut = tmpStrOut & padZeros(rsEntry("ccAmt"), 10)	'Charge Amount - 10
							tmpStrOut = tmpStrOut & "+"	'Charge Sign "-" for Credit, "+" Debit

							'tmpStrOut = tmpStrOut & padZeros(Day(rsEntry("AuthTime")),2) & padZeros(Month(rsEntry("AuthTime")),2) & Right(Year(rsEntry("AuthTime")),2) 	'Charge Date - DDMMYY
							tmpStrOut = tmpStrOut & padZeros(Day(payDate),2) & padZeros(Month(payDate),2) & Right(Year(payDate),2) 	'Charge Date - DDMMYY

							tmpStrOut = tmpStrOut & rsEntry("ExpMoYr")	'Card Expiry Date MMYY
							tmpStrOut = tmpStrOut & HTMLSpace(3)	'Authorization Code
							tmpStrOut = tmpStrOut & "03"	'Transaction Type, 03 - Mail Order
							tmpStrOut = tmpStrOut & padZeros(Right(rsEntry("TransactionNumber"),7),7)	'Transaction Reference Code
							tmpStrOut = tmpStrOut & HTMLSpace(2)	'Special Program Code
							'tmpStrOut = tmpStrOut & Left(session("StudioName"),40)	'Transaction Description
							'if LEN(session("StudioName"))<40 then
							'	tmpStrOut = tmpStrOut & HTMLSpace(40-LEN(session("StudioName")))
							'end if
							'tmpStrOut = tmpStrOut & HTMLSpace(40)	'Transaction Description (Formating by industry?)
							tmpStrOut = tmpStrOut & FmtPadString(rsEntry("RSSID"), 40, " ", false)	'Remarks - 15 use RSSID padded with spaces on the right
	
							tmpStrOut = tmpStrOut & HTMLSpace(226)	'Additional Transaction Description
	
							'Total 350
							addLineToOutPutFile(tmpStrOut)
							
							totNumTrans = totNumTrans + 1
							totAmt = totAmt + rsEntry("ccAmt")
							'---------------- END DETAIL  -----------------------------
	
						end if	'checked
						rsEntry.MoveNext
					loop
	
					'---------------- BEGIN BATCH TRAILER --------------------------------
					intCounter = intCounter + 1
	
					tmpStrOut = "04"	'Record Type
					tmpStrOut = tmpStrOut & padZeros(intCounter,6)		'Sequence Number
					tmpStrOut = tmpStrOut & padZeros(totNumTrans, 6)	'Batch Total Number of Trans
					tmpStrOut = tmpStrOut & padZeros(totAmt, 10)	'Batch Tota Number of DR Amount
					tmpStrOut = tmpStrOut & padZeros("", 10)		'Batch Tota Number of CR Amount
					tmpStrOut = tmpStrOut & HTMLSpace(316)	'Reserved
	
					'Total 350
					addLineToOutPutFile(tmpStrOut)
					'---------------- END BATCH TRAILER --------------------------------
	
					'---------------- BEGIN FILE TRAILER --------------------------------
					intCounter = intCounter + 1
	
					tmpStrOut = "05"	'Record Type
					tmpStrOut = tmpStrOut & padZeros(intCounter,6)		'Sequence Number
					tmpStrOut = tmpStrOut & padZeros(totNumTrans, 6)	'Batch Total Number of Trans
					tmpStrOut = tmpStrOut & padZeros(totAmt, 10)	'Batch Tota Number of DR Amount
					tmpStrOut = tmpStrOut & padZeros("", 10)		'Batch Tota Number of CR Amount
					tmpStrOut = tmpStrOut & HTMLSpace(316)	'Reserved
	
					'Total 350
					addLineToOutPutFile(tmpStrOut)
					'---------------- END FILE TRAILER --------------------------------

				else 'then	'AMEX Hong Kong  (NOT Tawain or Singapore)
					'Spec in JAPA Financial Settlement Guide RFG - OCT 08.doc
					intCounter = 1

					'---------------- BEGIN FILE HEADER --------------------------------
					tmpStrOut = "TFH"												'Record Type
					tmpStrOut = tmpStrOut & padZeros(intCounter,8)					'Record Number
					tmpStrOut = tmpStrOut & AMEXBatchSubmitterID					'AMEX Submitter ID
					if LEN(AMEXBatchSubmitterID)<11 then
						tmpStrOut = tmpStrOut & HTMLSpace(11-LEN(AMEXBatchSubmitterID))
					end if
					tmpStrOut = tmpStrOut & HTMLSpace(21)							'Reserved
					tmpStrOut = tmpStrOut & nextBatchFileNum						'Subitter File Reference Number
					if LEN(nextBatchFileNum)<9 then
						tmpStrOut = tmpStrOut & HTMLSpace(9-LEN(nextBatchFileNum))
					end if
					tmpStrOut = tmpStrOut & HTMLSpace(9)							'SUBMITTER_FILE_SEQUENCE_NUMBER - Blank in PURE Example File
					'tmpStrOut = tmpStrOut & payDatei18nFmt							'FILE_CREATION_DATE - YYYYMMDD
					'CB 4/17/09 Updated to current date
					tmpStrOut = tmpStrOut & curDatei18nFmt							'FILE_CREATION_DATE - YYYYMMDD
					'if TimeValue(curTime) >= TimeValue("1:00:00 PM") then			'FILE_CREATION_TIME - HHMMSS (24-hour clock)
					'	tmpStrOut = tmpStrOut & padZeros(CSTR(DatePart("h", curTime) + 12), 2)
					'else
						tmpStrOut = tmpStrOut & padZeros(CSTR(DatePart("h", curTime)), 2)
					'end if
					tmpStrOut = tmpStrOut & padZeros(CSTR(DatePart("n", curTime)), 2)
					tmpStrOut = tmpStrOut & padZeros(CSTR(DatePart("s", curTime)), 2)
					tmpStrOut = tmpStrOut & "08020000"								'FILE_VERSION_NUMBER - Constant
					tmpStrOut = tmpStrOut & HTMLSpace(617)							'Reserved

					'Total 700
					addLineToOutPutFile(tmpStrOut)
					'---------------- END BATCH HEADER --------------------------------
	
					do while NOT rsEntry.EOF
						if request.form("chk_"&rsEntry("TransactionNumber"))="on" AND rsEntry("ccAmt")<>0 then
	
							'---------------- BEGIN DETAIL  -----------------------------
							intCounter = intCounter + 1
							
							tmpStrOut = "TAB"										'Record Type
							tmpStrOut = tmpStrOut & padZeros(intCounter,8)			'Record Number

							'CB 4/17/09 
							tmpStrOut = tmpStrOut & padZeros(rsEntry("TransactionNumber"), 15)	'TRANSACTION_IDENTIFIER

							tmpStrOut = tmpStrOut & "02"							'FORMAT_CODE - Updated to '02' so TAA record is not required '20' from PURE Exmaple for General Retail
							tmpStrOut = tmpStrOut & "03"							'MEDIA_CODE - '03' from PURE Exmaple for Mail Order
							tmpStrOut = tmpStrOut & "06"							'SUBMISSION_METHOD - '03' from PURE Exmaple for PurchaseExpress
							tmpStrOut = tmpStrOut & HTMLSpace(10)					'Reserved
							'tmpStrOut = tmpStrOut & HTMLSpace(6)					'APPROVAL_CODE - (6) Per Benny Space Filled
							tmpStrOut = tmpStrOut & "000000"    					'APPROVAL_CODE - (6) Per Benny Zero Filled 6/1/09

							'Card Number - Always 15 digits
							tmpStr = rsEntry("ccNum")
							tmpStr = DES_Decrypt(tmpStr,true,null)
							tmpStrOut = tmpStrOut & tmpStr							'PRIMARY_ACCOUNT_NUMBER
							if LEN(tmpStr)<19 then
								tmpStrOut = tmpStrOut & HTMLSpace(19-LEN(tmpStr))	
							end if
							tmpStrOut = tmpStrOut & RIGHT(rsEntry("ExpMoYr"),2) & LEFT(rsEntry("ExpMoYr"),2)	'CARD_EXPIRY_DATE - YYMM

							'tmpStrOut = tmpStrOut & payDatei18nFmt					'TRANSACTION_DATE - YYYYMMDD
							'tmpStrOut = tmpStrOut & "111111"						'TRANSACTION_TIME - HHMMSS - set to 111111 in PURE example
							'CB 4/17/09 - Updated to use current date
							tmpStrOut = tmpStrOut & curDatei18nFmt					'TRANSACTION_DATE - YYYYMMDD
							tmpStrOut = tmpStrOut & "000000"						'TRANSACTION_TIME - HHMMSS - set to 111111 in PURE example
							tmpStrOut = tmpStrOut & "000"       					'Reserved
							tmpStrOut = tmpStrOut & padZeros(rsEntry("ccAmt"), 12)	'TRANSACTION_AMOUNT
							tmpStrOut = tmpStrOut & "000000"						'PROCESSING_CODE - "000000" for Debit

							'tmpStrOut = tmpStrOut & "392"							'TRANSACTION_CURRENCY_CODE - 344 – HK, 702 – SIN, 901 – TW
							'CB 4/17/09 Per Benny - 344 – HK, 702 – SIN, 901 – TW
							if ss_CountryCode="CN" OR ss_CountryCode="HK"	then 'Cn/HK
								tmpStrOut = tmpStrOut & "344"						'TRANSACTION_CURRENCY_CODE
							else	'SG
								tmpStrOut = tmpStrOut & "702"						'TRANSACTION_CURRENCY_CODE
							end if
							tmpStrOut = tmpStrOut & "01"							'EXTENDED_PAYMENT_DATA - Numer of monthly payments, 01 in PURE example

							tmpStr = ""
							tmpMerchantID = ""
							if NOT isNULL(rsEntry("TerminalID")) then
								tmpStr = CSTR(rsEntry("TerminalID"))
							end if
							tmpStrOut = tmpStrOut & tmpStr							'MERCHANT_ID (15) left justified space filled
							tmpMerchantID = tmpStr
							if LEN(tmpStr)<15 then
								tmpStrOut = tmpStrOut & HTMLSpace(15-LEN(tmpStr))
							end if

							'CB 4/17/09 Per Benny Space Filled
							tmpStrOut = tmpStrOut & HTMLSpace(15)					'MERCHANT_LOCATION_ID (15) Alphanumeric, upper case, left justified, character space filled

							tmpStr = ""
							if NOT isNULL(rsEntry("OP_AcctNum")) then
								tmpStr = CSTR(rsEntry("OP_AcctNum"))
							end if
							tmpStrOut = tmpStrOut & tmpStr							'MERCHANT_CONTACT_INFORMATION (40) left justified, character space filled
							if LEN(tmpStr)<40 then
								tmpStrOut = tmpStrOut & HTMLSpace(40-LEN(tmpStr))	
							end if

							'CB 4/17/09 Per Benny Space Filled
							tmpStrOut = tmpStrOut & HTMLSpace(8)						'TERMINAL_ID (8)
							'CB 4/17/09 Per Benny Should be 000090100000
							tmpStrOut = tmpStrOut & "000090100000"					'POINT_OF_SERVICE_DATA_CODE (12)
							tmpStrOut = tmpStrOut & "000"							'Field 23 - Reserved (3)
							tmpStrOut = tmpStrOut & "000000000000"					'Field 24 - Reserved (12)
							tmpStrOut = tmpStrOut & HTMLSpace(3)					'Field 25 - Reserved (3)

							tmpStr = ""
							if NOT isNULL(rsEntry("SaleID")) then
								tmpStr = CSTR(rsEntry("SaleID"))
							end if
							tmpStrOut = tmpStrOut & tmpStr							'INVOICE/REFERENCE_NUMBER (30) left justified space padded
							if LEN(tmpStr)<30 then
								tmpStrOut = tmpStrOut & HTMLSpace(30-LEN(tmpStr))	
							end if
							tmpStrOut = tmpStrOut & HTMLSpace(15)					'Transaction Advice Basic 
							tmpStrOut = tmpStrOut & HTMLSpace(436)					'Reserved

							totNumTrans = totNumTrans + 1
							totAmt = totAmt + rsEntry("ccAmt")
							'---------------- END DETAIL  -----------------------------

							'Total 700
							addLineToOutPutFile(tmpStrOut)
	
						end if	'checked
						rsEntry.MoveNext
					loop
					
					'---------------- BEGIN TRANSACTION BATCH TRAILER (TBT) ------------------
					intCounter = intCounter + 1
							
					tmpStrOut = "TBT"												'Record Type
					tmpStrOut = tmpStrOut & padZeros(intCounter,8)			        'Record Number
					tmpStrOut = tmpStrOut & tmpMerchantID                           'MERCHANT_ID
					if LEN(tmpMerchantID)<15 then
						tmpStrOut = tmpStrOut & HTMLSpace(15-LEN(tmpMerchantID))
					end if
					tmpStrOut = tmpStrOut & HTMLSpace(15)				        	'RESERVED
					tmpStrOut = tmpStrOut & padZeros(nextBatchFileNum, 15)	        'TBT_IDENTIFICATION_NUMBER 
					tmpStrOut = tmpStrOut & curDatei18nFmt 	                        'TBT_CREATION_DATE (8)
    				tmpStrOut = tmpStrOut & padZeros(totNumTrans, 8)                'Number of Records (8)
					tmpStrOut = tmpStrOut & "000"       				        	'RESERVED (3)
    				tmpStrOut = tmpStrOut & padZeros(totAmt, 20)                    'TBT_AMOUNT (20)
					tmpStrOut = tmpStrOut & "+"         				        	'+/- for Debit or Credit
					if ss_CountryCode="CN" OR ss_CountryCode="HK"	then 'CN/HK
						tmpStrOut = tmpStrOut & "344"						        'TRANSACTION_CURRENCY_CODE - 344 – HK, 702 – SIN, 901 – TW
					else	'SG
						tmpStrOut = tmpStrOut & "702"						        'TRANSACTION_CURRENCY_CODE
                    end if
					tmpStrOut = tmpStrOut & "000"							        'Field 12 - Reserved (3)
					tmpStrOut = tmpStrOut & "00000000000000000000"					'Field 13 - Reserved (20)
					tmpStrOut = tmpStrOut & HTMLSpace(3)				        	'RESERVED
					tmpStrOut = tmpStrOut & HTMLSpace(575)      					'RESERVED

				    'Total 700
				    addLineToOutPutFile(tmpStrOut)


					'---------------- BEGIN TRANSACTION FILE SUMMARY (TFS) ------------------
					intCounter = intCounter + 1
							
					tmpStrOut = "TFS"												'Record Type
					tmpStrOut = tmpStrOut & padZeros(intCounter,8)			        'Record Number
    				tmpStrOut = tmpStrOut & padZeros(totNumTrans, 8)                'Number of Debits (8)
					tmpStrOut = tmpStrOut & "000"							        'Field 4 - Reserved (3)
    				tmpStrOut = tmpStrOut & padZeros(totAmt, 20)                    'HASH_TOTAL_DEBIT_AMOUNT (20)
					tmpStrOut = tmpStrOut & "00000000"		    			        'NUMBER_OF_CREDITS (8)
					tmpStrOut = tmpStrOut & "000"							        'Field 7 - Reserved (3)
    				tmpStrOut = tmpStrOut & "00000000000000000000"                  'HASH_TOTAL_CREDIT_AMOUNT (20)
					tmpStrOut = tmpStrOut & "000"							        'Field 9 - Reserved (3)
    				tmpStrOut = tmpStrOut & padZeros(totAmt, 20)                    'HASH_TOTAL_AMOUNT (20)
					tmpStrOut = tmpStrOut & HTMLSpace(604)      					'RESERVED

				    'Total 700
				    addLineToOutPutFile(tmpStrOut)


				end if 	'TW/SNG vs HK
			end if 'TransType

		end if	'session("ccProcessor")

		'================================================================================================
		'								END - HSBC MULTIPLE FORMATS
		'================================================================================================


		'Finish Writing File
		of.close
		set of = nothing




			'***************TEST HSBC************************
			'if Session("CCProcessor")="HSBC" then
				'response.end
			'end if
			'***************************************





		''*************************** BEGIN FTP UPLOAD **********************************''
		dim ftpCon, Global
		dim ftpFailed : ftpFailed = false

		if writeFTPLogFile then
			' Log File
			Dim FSO
			Set FSO = Server.CreateObject("Scripting.FileSystemObject")
			Dim sLogFile
			Dim sTempFolder 
			Dim sTempFileName : sTempFileName = FSO.GetTempName()
	
			' Use const temp folder. 
			' Important: Create the temp folder in your web and give read/write permissions to the IWAM_ user if you want the log
			sTempFolder = Server.MapPath(cLogFolder)
			sLogFile = sTempFolder & "\" & sTempFileName & ".log"
			'Response.Write("LogFile = " & sTempFile & vbCrLf)
	
		end if

		attemptNum = 1
		submitAttempts = 1
		submitSuccess = false

		do while NOT submitSuccess AND attemptNum <= submitAttempts


			Set Global = CreateObject("sfFTPLib.Global")
			Global.LoadLicenseKeyFile(Request.ServerVariables("APPL_PHYSICAL_PATH") & "SmartFTPLicense.txt")
			
			
			if session("ccProcessor")="HSBC" then 'PURE
			    
				strSQL = "INSERT INTO tblContactLogs(ContactLog, ClientID, TrainerID, SystemGenerated, EmailStatus, Deleted, ContactDate) VALUES ('" & filename  & "',1,-3,1,1,1," & DateSep & Now & DateSep &" ) "
				cnWS.execute strSQL

			    Set ftpCon = Server.CreateObject("sfFTPLib.SFTPConnection")

			    if writeFTPLogFile then
				    ftpCon.LogFile = sLogFile
			    end if

			    'PURE FTP Info
			    '220.232.190.39 port 22 
			    'U: mbo
			    'P: mbo123456
    						
	            'Local testing w/ filezilla
			    if testEnvironment then
				    ftpCon.Host = "127.0.0.1"
				    ftpCon.Port = "21"
				    'ftpCon.Protocol = "0"	'FTP Normal
				    'ftpCon.Protocol = "2"	'FTP SSL Explicit
				    ftpCon.Username = "chet"
				    ftpCon.Password = "brandenburg"
				    'ftpCon.Passive = true
				    'ftpCon.HidePassword = true
    	        
    	        
    	        
			    'Temp send to dev for pre-testing
			    'if testEnvironment then
				    'ftpCon.Host = "backup.mindbodyonline.com"
				    'ftpCon.Port = "22"
				    ''ftpCon.Protocol = "0"	'FTP Normal
				    'ftpCon.Protocol = "2"	'FTP SSL Explicit
				    'ftpCon.Userna me = "chet"
				    'ftpCon.Password = "brandenburg"
				    'ftpCon.Passive = true
				    'ftpCon.HidePassword = true
			    else
				    ftpCon.Host = ftpHeader
				    ftpCon.Port = "22"
				    'ftpCon.Protocol = 2	'FTP SSL Explicit
				    ftpCon.Username = ftpUser
				    ftpCon.Password = ftpPwd
				    'ftpCon.Passive = true
				    'ftpCon.PROTFallback = false
    	
				    'tried
				    'ftpCon.DataProtection = ftpDataProtectionPrivate	'numeric 0,1... 2 seems to be consistent
				    'ftpCon.DataProtection = 2 
				    'ftpCon.HidePassword = true
    	
			    end if
    		
			    response.write ("FTP Connecting...<br />")
			    response.flush
	            'response.end
			    result = ftpCon.Connect
    	
			    if result = ftpErrorSuccess then
				    response.write("FTP Connected.<br />")
				    'result = ftpCon.ReadDirectory
    	
				    if result = ftpErrorSuccess then
					    response.write "read directory"
					    response.flush

					    result = ftpCon.UploadFile(studio_path & session("studioShort") & "\" & filename, filename, 0, 0, 0, 0)
    			
					    if result = ftpErrorSuccess then
						    response.write "Uploaded file.<br />"
					    else
						    response.write "Upload FAILED. Error: " & ftpCon.LastError
						    ftpFailed = true
					    end if
					    response.flush
				    else
					    response.write "FTP Read Directory FAILED. Error = " & ftpCon.LastError
					    ftpFailed = true
				    end if
			    else
				    response.write "FTP Connect() FAILED. Error = " & ftpCon.LastError & "<br /><br />"
				    ftpFailed = true
				    response.Write "Host: " & ftpCon.Host & "<br />"
				    response.Write "Port: " & ftpCon.Port & "<br />"
				    'response.Write "Protocol: " & ftpCon.Protocol & "<br />"
				    response.Write "Username: " & ftpCon.Username & "<br />"
				    response.Write "Password: " & ftpCon.Password & "<br />"
				    'response.Write "Passive: " & ftpCon.Passive & "<br />"
				    'response.Write "PROTFallback: " & ftpCon.PROTFallback & "<br />"
				    'response.Write "DataProtection: " & ftpCon.DataProtection & "<br />"
				    'response.Write "HidePassword: " & ftpCon.HidePassword & "<br />"
    				
			    end if
			    response.flush
			
			
			
			else 'NOT HSBC
			
			
			    Set ftpCon = Server.CreateObject("sfFTPLib.FTPConnectionSTA")

			    if writeFTPLogFile then
				    ftpCon.LogFile = sLogFile
			    end if

	            'Local testing w/ filezilla
			    if testEnvironment then
				    ftpCon.Host = "127.0.0.1"
				    ftpCon.Port = "21"
				    ftpCon.Protocol = "0"	'FTP Normal
				    'ftpCon.Protocol = "2"	'FTP SSL Explicit
				    ftpCon.Username = "chet"
				    ftpCon.Password = "brandenburg"
				    ftpCon.Passive = true
				    ftpCon.HidePassword = true
    	        
    	        
    	        
			    'Temp send to dev for pre-testing
			    'if testEnvironment then
				    'ftpCon.Host = "backup.mindbodyonline.com"
				    'ftpCon.Port = "22"
				    ''ftpCon.Protocol = "0"	'FTP Normal
				    'ftpCon.Protocol = "2"	'FTP SSL Explicit
				    'ftpCon.Userna me = "chet"
				    'ftpCon.Password = "brandenburg"
				    'ftpCon.Passive = true
				    'ftpCon.HidePassword = true
			    else
				    if session("ccProcessor")="MON" OR Session("CCProcessor")="OP" OR (Session("CCProcessor")="PMN" AND Session("CCProcessor2")="ELV") then				
					    ftpCon.Host = "ftpssl.rbc.com"
				    elseif session("ccProcessor")="HSBC" then
					    ftpCon.Host = ftpHeader
				    end if
					    ftpCon.Port = "21"
				    ftpCon.Protocol = 2	'FTP SSL Explicit
				    ftpCon.Username = ftpUser
				    ftpCon.Password = ftpPwd
				    ftpCon.Passive = true
				    ftpCon.PROTFallback = false
    	
				    'tried
				    'ftpCon.DataProtection = ftpDataProtectionPrivate	'numeric 0,1... 2 seems to be consistent
				    ftpCon.DataProtection = 2 
				    ftpCon.HidePassword = true
    	
			    end if
    		
			    response.write ("FTP Connecting...<br />")
			    response.flush
	            'response.end
			    result = ftpCon.Connect
    	
			    if result = ftpErrorSuccess then
				    response.write("FTP Connected.<br />")
				    result = ftpCon.ReadDirectory
    	
				    if result = ftpErrorSuccess then
					    response.write "read directory"
					    response.flush

					    result = ftpCon.UploadFile(studio_path & session("studioShort") & "\" & filename, filename, 0, 0)
    			
					    if result = ftpErrorSuccess then
						    response.write "Uploaded file.<br />"
					    else
						    response.write "Upload FAILED. Error: " & ftpCon.LastError
							logASPError "ftpComponent", "FTP Connection Error - UPLOAD FAILED - RBC","adm_rpt_ccp_ach_p.asp", ftpCon.LastError, "0" 
						    ftpFailed = true
					    end if
					    response.flush
				    else
					    response.write "FTP Read Directory FAILED. Error = " & ftpCon.LastError
					    ftpFailed = true
				    end if
			    else
				    response.write "FTP Connect() FAILED. Error = " & ftpCon.LastError & "<br /><br />"
				    ftpFailed = true
				    response.Write "Host: " & ftpCon.Host & "<br />"
				    response.Write "Port: " & ftpCon.Port & "<br />"
				    response.Write "Protocol: " & ftpCon.Protocol & "<br />"
				    response.Write "Username: " & ftpCon.Username & "<br />"
				    response.Write "Password: " & ftpCon.Password & "<br />"
				    response.Write "Passive: " & ftpCon.Passive & "<br />"
				    response.Write "PROTFallback: " & ftpCon.PROTFallback & "<br />"
				    response.Write "DataProtection: " & ftpCon.DataProtection & "<br />"
				    response.Write "HidePassword: " & ftpCon.HidePassword & "<br />"
    				
			    end if
			    response.flush
			
			
			
			
			end if 'HSBC
	
			Set ftpCon = Nothing
			Set Global = Nothing

			attemptNum = attemptNum + 1
		loop	'resubmit up to 5 times

		''*************************** END FTP UPLOAD *******************************''
		if fs.FileExists(studio_path & session("studioShort") & "\" & filename) AND NOT session("ccProcessor")="HSBC" then	
			fs.DeleteFile(studio_path & session("studioShort") & "\" & filename)
		end if
		set fs = nothing

		''*************************** UPDATE DB WITH UPLOAD ***********************''
		if NOT ftpFailed then

			''Pass 3
			rsEntry.MoveFirst
			do while NOT rsEntry.EOF
				if request.form("chk_"&rsEntry("TransactionNumber"))="on" then
					strSQL = "UPDATE tblCCTrans SET Status = N'Sent to Bank', BatchFileNum = " & nextBatchFileNum & ", TransTime=" & DateSep & Now & DateSep & ", BankPaymentDate=" & DateSep & payDate & DateSep & " WHERE (TransactionNumber = " & rsEntry("TransactionNumber") & ")"
					cnWS.execute strSQL
				end if
				rsEntry.MoveNext
			loop
		end if
%>
		<script type="text/javascript">
			<%if ftpFailed then%>
				alert("File Uploaded FAILED. \n\nPlease send the output to MINDBODY Online");
			<%else%>
				alert("The File was Sent Successfully!");
				document.location.replace("adm_rpt_ccp_ach.asp");
			<%end if%>
		</script>
<%
	else	''No Transactions Found... this case should be caught on UI side
		
	end if	''No Transactions Found... this case should be caught on UI side
	rsEntry.close

	strSQL = "UPDATE tblCCOpts SET ActiveBatch=0"
	cnWS.execute strSQL
%>

                  </td>
                </tr>
              </table>
            </td>
          </tr>
          <tr valign="middle" height="100%" class="printHide"> 
            <td class="headText center-ch" width="100%" valign="middle"> 
              <br />
              <br />
            </td>
          </tr>
        </table>
    </td>
    </tr>
		</table>
<% pageEnd %>
</body>
</html>
<%
	end if	'chk AP
end if	'chk session
%>
