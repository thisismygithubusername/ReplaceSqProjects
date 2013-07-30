<%@ CodePage=65001 %>
<!-- #include file="init.asp" -->
<%
'Dim SessionFarm
'set SessionFarm = Server.CreateObject("SessionFarm.SFSession")

Server.ScriptTimeout = 1200    '20 min (value in seconds)
response.charset="utf-8"
%>

<!-- #include file="../inc_dbconn_wsMaster.asp" -->
<!-- #include file="inc_accpriv.asp" -->

<%

Dim ss_UseEFT, tmpCCNum
ss_UseEFT = checkStudioSetting("Studios","UseEFT")

		dim rsEntry, strTempName, curMess, intCount, cashCount, settler, cMID, cTID, ccProcessor, oneMID, newOrderID, Moneris_ECR, origResponseSTR, origResponseList, errMsg, activeBatch, lastBatch, batchWait
		dim ccp_i, xmldom, xmlhttp, ccp_testing, op_TxnNumber
		dim ccp_ccTestMode : ccp_ccTestMode = checkStudioSetting("Studios","ccTestMode")
		ccp_testing = false
		op_TxnNumber = ""
		batchWait = 5 'minutes between unfinished batching attempts
		newOrderID = ""
		set rsEntry = Server.CreateObject("ADODB.Recordset")
		if implementationSwitchIsEnabled("BluefinCanada") then
    		strSQL = "SELECT Studios.CCProcessor, ActiveBatch, LastBatch FROM tblCCOpts, Studios WHERE tblCCOpts.StudioID=" & session("StudioID")
        else
            strSQL = "SELECT tblCCOpts.CCProcessor, ActiveBatch, LastBatch FROM tblCCOpts WHERE tblCCOpts.StudioID=" & session("StudioID")
        end if
		rsEntry.CursorLocation = 3
		rsEntry.open strSQL, cnWS
		Set rsEntry.ActiveConnection = Nothing
			ccProcessor = rsEntry("CCProcessor")
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





if not Session("Pass") OR Session("Admin")="false" OR NOT validAccessPriv("RPT_CCP") then 
%>
<script type="text/javascript">
	alert("You are not authorized to view this page.\nPlease login to an account with Access Privileges granted for this feature.");
	javascript:history.go(-1);
</script>
<%
else
%>
		<!-- #include file="../inc_i18n.asp" -->
		<!-- #include file="inc_chk_ss.asp" -->
		<!-- #include file="inc_ws_stats.asp" -->
		<!-- #include file="inc_crypt.asp" -->
		<!-- #include file="../inc_localization.asp" -->
<%
ok = -1

		oneMID = false
		'if ccProcessor<>"OP" then	'optimal needs location for config file
			strSQL = "SELECT MID, TID, OP_AcctNum FROM Location GROUP BY MID, TID, OP_AcctNum HAVING (NOT (MID IS NULL))"
			rsEntry.CursorLocation = 3
			rsEntry.open strSQL, cnWS
			Set rsEntry.ActiveConnection = Nothing
			if NOT rsEntry.EOF then
				if rsEntry.RecordCount=1 then
					oneMID = true
					cMID = rsEntry("MID")
					cTID = rsEntry("TID")
					Moneris_ECR = rsEntry("OP_AcctNum")
				end if
			end if
			rsEntry.close
		'end if
		
		if NOT oneMID then
			strSQL = "SELECT MID, TID, OP_AcctNum FROM Location WHERE LocationID=" & request.form("frmBatchLoc")
			rsEntry.CursorLocation = 3
			rsEntry.open strSQL, cnWS
			Set rsEntry.ActiveConnection = Nothing
				cMID = 	rsEntry("MID")
				cTID = 	rsEntry("TID")
				Moneris_ECR = rsEntry("OP_AcctNum")
			rsEntry.close
		end if

		'''Get New BatchNumber
		strSQL = "SELECT MAX(BatchNumber) AS MaxBatchNum FROM tblCCTrans"
		rsEntry.CursorLocation = 3
		rsEntry.open strSQL, cnWS
		Set rsEntry.ActiveConnection = Nothing

		if NOT rsEntry.EOF then
			if NOT isNull(rsEntry("MaxBatchNum")) then
				newBatchNum = rsEntry("MaxBatchNum") + 1
			else
				newBatchNum = 1			
			end if
		else
			newBatchNum = 1
		end if
		rsEntry.close

		strSQL = "UPDATE tblCCOpts SET LastBatch=" & DateSep & Now & DateSep &", ActiveBatch=1"
		cnWS.execute strSQL
%>

<!-- #include file="pre.asp" -->
<!-- #include file="frame_bottom.asp" -->

<%= js(array("mb", "MBS")) %>
<!-- #include file="../inc_date_ctrl.asp" -->
<!-- #include file="css/site_setup.asp" -->
<% pageStart %>


<table height="100%" width="<%=strPageWidth%>" cellspacing="0">

<tr> 
     <td valign="top" height="100%" width="100%">
        <table class="center" border="0" cellspacing="0" cellpadding="0" width="90%" height="100%">
          <tr> 
            <td class="headText" align="left" valign="top"> 
				<div id="topdiv">
              <table class="mainText" width="100%" border="0" cellspacing="0" cellpadding="0">
                <tr> 
                  <td class="headText"><b>Batch &amp; Settle Results</b></td>
                  <td align="right" valign="top"> 
                    <table class="mainText" border="1" cellspacing="0" cellpadding="0" bordercolor="<%=session("pageColor4")%>">
                    </table>
                  </td>
                </tr>
              </table>
				</div>
            </td>
          </tr>
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
            <td valign="top" class="mainTextBig"> <br />
              <table class="mainText center" border="0" cellspacing="0" cellpadding="0" width="90%">
              <tr> 
                <td valign="top" align=right> 
                  <table class="mainText" border="0" cellpadding="0" cellspacing="0" width="100%">
                    <tr style="background-color:<%=session("pageColor2")%>;"> 
                      <td colspan=5><img height="2" width="100%" src="<%= contentUrl("/asp/adm/images/trans.gif") %>"></td>
                    </tr>
                    <tr bgcolor="<%Response.Write session("pageColor4")%>"> 
                      <td class="whiteHeader" nowrap> <b>&nbsp;Transaction#</b></td>
                      <td class="whiteHeader" nowrap> <b>&nbsp;<%=session("ClientHW")%> <%=xssStr(allHotWords(40))%></b> 
                      </td>
                      <td class="whiteHeader" nowrap align="center"><b>&nbsp;<%=xssStr(allHotWords(35))%>&nbsp;</b></td>
                      <td class="whiteHeader" nowrap align="center"><%if ccProcessor<>"PMN" then%><b>&nbsp;Batch Number</b><b>&nbsp;</b><%end if%></td>
                      <td class="whiteHeader" nowrap align="center"><b>&nbsp;Result</b><b>&nbsp;</b></td>
                    </tr>
                    <tr style="background-color:<%=session("pageColor2")%>;"> 
                      <td colspan=5><img height="2" width="100%" src="<%= contentUrl("/asp/adm/images/trans.gif") %>"></td>
                    </tr>
<%
		if ccProcessor="CCP" then
			Set settler = CreateObject("ccp.settler")
			settler.merchantId = cMID
			settler.terminalId = cTID
			settler.batchNum = newBatchNum
		elseif ccProcessor="PMN" then
			Set settler = CreateObject("ATS.SecurePost")
			'settler.DevMode = true
			'settler.ATSID = "PNH69"
			'settler.ATSSubID = cMID
			settler.ATSID = cMID
			if NOT isNULL(cTID) then
				settler.ATSSubID = cTID
			end if
		elseif ccProcessor="MON" then
			Dim monerisReq, crypt_type
			crypt_type = "7"
		
		elseif ccProcessor="OP" OR ccProcessor="TEL" then
		
			Dim Dict, ConfigFile, op_ccp
			Set Dict = Server.CreateObject ("Scripting.Dictionary")
		
		end if

	if ss_UseEFT then	''Left Join w/tblEFTSchedules
		strSQL = "SELECT DISTINCT tblCCTrans.Status, tblCCTrans.Settled, tblCCTrans.ccAmt, tblCCTrans.ClientID, tblCCTrans.ccNum, tblCCTrans.ExpMoYr, tblCCTrans.TransactionNumber, tblCCTrans.AuthCode, tblCCTrans.CCSwiped, tblCCTrans.TransTime, tblCCTrans.op_TxnNumber, tblCCTrans.OrderID, tblCCTrans.responseStr, CLIENTS.LastName, CLIENTS.FirstName, tblEFTSchedule.StatusCode "
		strSQL = strSQL & "FROM (CLIENTS INNER JOIN tblCCTrans ON CLIENTS.ClientID = tblCCTrans.ClientID) LEFT JOIN tblEFTSchedule ON tblCCTrans.TransactionNumber = tblEFTSchedule.CCTransID "
		strSQL = strSQL & "WHERE (((tblCCTrans.Settled)=0) AND (tblCCTrans.Status = 'Approved' "
		if ccProcessor="TCI" then
			strSQL = strSQL & " OR tblCCTrans.Status = 'Credit'"
		end if
		strSQL = strSQL & ") "
		strSQL = strSQL & "AND ( ((tblCCTrans.MerchantID)=N'" & cMID & "')) OR ((tblCCTrans.MerchantID)=N'" & cTID & "')) "
	else
		strSQL = "SELECT tblCCTrans.Status, tblCCTrans.Settled, tblCCTrans.ccAmt, tblCCTrans.ClientID, tblCCTrans.ccNum, tblCCTrans.ExpMoYr, tblCCTrans.TransactionNumber, tblCCTrans.AuthCode, tblCCTrans.TransTime, tblCCTrans.op_TxnNumber, tblCCTrans.OrderID, tblCCTrans.responseStr, CLIENTS.LastName, CLIENTS.FirstName "
		strSQL = strSQL & "FROM CLIENTS INNER JOIN tblCCTrans ON CLIENTS.ClientID = tblCCTrans.ClientID "
		strSQL = strSQL & "WHERE (((tblCCTrans.Settled)=0) AND (tblCCTrans.Status = 'Approved' "
		if ccProcessor="TCI" then
			strSQL = strSQL & " OR tblCCTrans.Status = 'Credit'"
		end if
		strSQL = strSQL & ") "
		strSQL = strSQL & "AND ( ((tblCCTrans.MerchantID)=N'" & cMID & "')) OR ((tblCCTrans.MerchantID)=N'" & cTID & "')) "
	end if
	strSQL = strSQL & "ORDER BY tblCCTrans.CCSwiped, tblCCTrans.TransTime DESC;"

	'response.Write strSQL
	'response.end

	rsEntry.CursorLocation = 3
	rsEntry.open strSQL, cnWS
	Set rsEntry.ActiveConnection = Nothing

		rowcount = 0
		intcount = 0
		cashCount = 0

		Do While NOT rsEntry.EOF

			if request.form("chk_"&rsEntry("TransactionNumber"))="on" then

				''Added 9_12_06 by CB - checks for legacy unencrypted credit cards
				if isNumeric(rsEntry("ccNum")) then
					tmpCCNum = rsEntry("ccNum")
				else
					tmpCCNum = rsEntry("ccNum")
					tmpCCNum = DES_Decrypt(tmpCCNum,true,null)
				end if
	
				if ccProcessor = "CCP" then
					ok = settler.addRecord(CLNG(rsEntry("TransactionNumber")), "Sale", CSTR(tmpCCNum), CSTR(rsEntry("ExpMoYr")), CLNG(rsEntry("ccAmt")), CSTR(rsEntry("AuthCode")))
	
						'Debugging Code
						'Set auth = CreateObject("ccp.authorizer")
						'response.write "x" & tmpCCNum & "x"
						'response.write "result is: " & ok & "<br />"
						'response.write  auth.errIdToMsg(ok) & "<br />"
	
					if ok = 0 then
						curMess = "success"
					else
						curMess = "failed"
					end if
				elseif ccProcessor="TCI" then
					set xmldom = server.CreateObject("Microsoft.XMLDOM")
					set xmlhttp = server.CreateObject("MSXML2.ServerXMLHTTP")
						if ccp_ccTestMode then	'Test
							SoapServer = "https://web.cert.transfirst.com/prigateway/creditcard.asmx"
						else 'Live
							SoapServer = "https://webservices.primerchants.com/creditcard.asmx"
						end if
					
						xmlmsg = "<?xml version=""1.0"" encoding=""utf-8""?>"
						xmlmsg = xmlmsg & "<soapenv:Envelope xmlns:soapenv=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:web=""http://www.paymentresources.com/webservices/"">"
							xmlmsg = xmlmsg & "<soapenv:Header/>"
							xmlmsg = xmlmsg & "<soapenv:Body>"
								xmlmsg = xmlmsg & "<web:CloseBatch>"
									xmlmsg = xmlmsg & "<web:MerchantID>" & cMID & "</web:MerchantID>"
									xmlmsg = xmlmsg & "<web:RegKey>" & cTID & "</web:RegKey>"
									xmlmsg = xmlmsg & "<web:TransID>" & TRIM(rsEntry("OrderID")) & "</web:TransID>"
									xmlmsg = xmlmsg & "<web:Amount>" & Replace(FormatNumber(rsEntry("ccAmt")/100,2),",","") & "</web:Amount>"
									xmlmsg = xmlmsg & "<web:ForceSettlement>A</web:ForceSettlement>" 
									if rsEntry("Status")="Credit" then
										xmlmsg = xmlmsg & "<web:CreditID>" & TRIM(rsEntry("AuthCode")) & "</web:CreditID>"
									else
										xmlmsg = xmlmsg & "<web:CreditID></web:CreditID>"
									end if
								xmlmsg = xmlmsg & "</web:CloseBatch>"
							xmlmsg = xmlmsg & "</soapenv:Body>"
						xmlmsg = xmlmsg & "</soapenv:Envelope>"
					
						if ccp_testing then
							response.Write xmlmsg
						end if
					
						xmlhttp.open "POST", SoapServer, false
						xmlhttp.setRequestHeader "Content-Type", "text/xml; charset=utf-8"
						xmlhttp.setRequestHeader "Content-Length", Len(xmlmsg)
						if true then 'CC Real Time Processing TODO
								xmlhttp.setRequestHeader "SOAPAction", "http://www.paymentresources.com/webservices/CloseBatch"
						end if				

						xmlhttp.send xmlmsg

						if xmlhttp.Status = 200 then
							Set xmldom = xmlhttp.responseXML
							Set objLst = xmldom.getElementsByTagName("*")
							
							
							for ccp_i = 0 to (objLst.length - 1)
								if ccp_testing then
									response.write objLst.item(ccp_i).nodeName & ": "
									response.write objLst.item(ccp_i).text & vbCrLF & "<br /> "
								end if
								
								Select Case objLst.item(ccp_i).nodeName
									Case "Status"
										if objLst.item(ccp_i).text = "Settled" then	'Settled
											ok = 0
											curMess = "success"
										end if
									Case "Message"
										ok = -1
										curMess = objLst.item(ccp_i).text
									'Case else
										'Nothing else we want
								End Select
							next

						else	'Processor failed to response or unhandled error
							if ccp_testing then
								Response.Write "The request was unsuccessful.<br /><br />" & vbCrLF
								Response.Write "status:" & xmlhttp.status & "<br /><br />" & vbCrLF
								Response.write xmlhttp.statusText
							end if
							reasonText = "There was a problem communicating with the gateway."
						end if					
					
					set xmlhttp = nothing
					set xmldom = nothing
				
				elseif ccProcessor="PMN" then	'PAY ME NOW

					settler.Amount = CLNG(Replace(Replace(FormatNumber(rsEntry("ccAmt")/100,2),".",""),",",""))
					settler.ProcessPost(CSTR(rsEntry("AuthCode")))

					if settler.ResultAccepted AND newBatchNum<>"" then
						'response.write "BatchNumber is: " & settler.BatchNumber & "<br />"
						'response.write "ResponseString is: " & settler.ResultAuthCode & "<br />"
						responseSTR = settler.ResultAuthCode
						resultList = split(responseSTR,":")
						'response.write "BatchNumber from response string: " & resultList(3)
			
						newBatchNum = resultList(3)
						if newBatchNum<>"" then
							ok = 0
							curMess = "success"
						else
							ok = -1
							curMess = "Problem Getting Batch Number"
						end if
					else
						if settler.ResultErrorFlag Then
							curMess = settler.LastError
							'response.write "Error: " & settler.LastError
						else
							curMess = settler.ResultAuthCode
							'response.write "Declined: " & settler.ResultAuthCode
						end if
						ok = -1
					end if
				elseif ccProcessor="MON" then

                    origResponseSTR = rsEntry("responseStr")
                    origResponseList = split(origResponseSTR,"|")

					Set monerisReq = server.CreateObject("Moneris.Request")
				    'TESTING URL!!
				    'monerisReq.initRequest cMID, cTID, "https://esplusqa.moneris.com/gateway2/servlet/MpgRequest"	'storeID,api_token
					monerisReq.initRequest cMID, cTID, "https://www3.moneris.com/gateway2/servlet/MpgRequest"
		
					Set settler = server.CreateObject("Moneris.Completion")

					monerisReq.setRequest settler.formatRequest( CSTR(origResponseList(0)), Replace(FormatNumber(rsEntry("ccAmt")/100,2),",",""), CSTR(rsEntry("OrderID")), crypt_type )
					monerisReq.sendRequest
	
					ok = -1	'default failed
					if monerisReq.getResponseCode="null" then
						curMess = "Failed: " & monerisReq.getMessage
					else
						if isNumeric(monerisReq.getResponseCode) then
							if CINT(monerisReq.getResponseCode)<50 then	'approved
								curMess = "success"
								ok = 0
								newOrderID = monerisReq.getTransID
							else	'>= 50 declined
								curMess = monerisReq.getMessage & "Response Code:" & monerisReq.getResponseCode
							end if
						else
							curMess = monerisReq.getMessage & "Response Code:" & monerisReq.getResponseCode
						end if
					end if
				elseif ccProcessor="OP" OR ccProcessor="TEL" then
				
					set xmldom = server.CreateObject("Microsoft.XMLDOM")
					set xmlhttp = server.CreateObject("MSXML2.ServerXMLHTTP")
					
					ok = -1
					curMess = "failed"
					op_TxnNumber = ""
					errMsg = ""
					if ccp_ccTestMode then	'Test
						SoapServer = "https://webservices.test.optimalpayments.com/creditcardWS/CreditCardService/v1"
					else 'Live
						SoapServer = "https://webservices.optimalpayments.com/creditcardWS/CreditCardService/v1"
					end if
					
					xmlmsg = "<?xml version=""1.0"" encoding=""utf-8""?>"
					xmlmsg = xmlmsg & "<soapenv:Envelope xmlns:soapenv=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:v1=""http://www.optimalpayments.com/creditcard/v1"" xmlns:v11=""http://www.optimalpayments.com/creditcard/xmlschema/v1"">"
					xmlmsg = xmlmsg & "<soapenv:Header/>"
					xmlmsg = xmlmsg & "<soapenv:Body>"
					xmlmsg = xmlmsg & "<v1:ccSettlement>"
					xmlmsg = xmlmsg & "<v11:ccPostAuthRequestV1>"
					xmlmsg = xmlmsg & "<v11:merchantAccount>"
					xmlmsg = xmlmsg & "<v11:accountNum>" & Moneris_ECR & "</v11:accountNum>"
					xmlmsg = xmlmsg & "<v11:storeID>" & cMID & "</v11:storeID>"
					xmlmsg = xmlmsg & "<v11:storePwd>" & cTID & "</v11:storePwd>"
					xmlmsg = xmlmsg & "</v11:merchantAccount>"
					xmlmsg = xmlmsg & "<v11:confirmationNumber>" & TRIM(rsEntry("OrderID")) & "</v11:confirmationNumber>"
					xmlmsg = xmlmsg & "<v11:merchantRefNum>" & TRIM(rsEntry("TransactionNumber")) & "</v11:merchantRefNum>"
					xmlmsg = xmlmsg & "<v11:amount>" & Replace(FormatNumber(rsEntry("ccAmt")/100,2),",","") & "</v11:amount>"
					xmlmsg = xmlmsg & "<v11:dupeCheck>false</v11:dupeCheck>"
					xmlmsg = xmlmsg & "</v11:ccPostAuthRequestV1>"
					xmlmsg = xmlmsg & "</v1:ccSettlement>"
					xmlmsg = xmlmsg & "</soapenv:Body>"
					xmlmsg = xmlmsg & "</soapenv:Envelope>"
					
					
					if ccp_testing then
						response.Write xmlmsg
					end if
				
					xmlhttp.open "POST", SoapServer, false
					xmlhttp.setRequestHeader "Content-Type", "text/xml; charset=utf-8"
					xmlhttp.setRequestHeader "Content-Length", Len(xmlmsg)
					'xmlhttp.setRequestHeader "SOAPAction", ""		

					xmlhttp.send xmlmsg
					
					

					if xmlhttp.Status = 200 then
						Set xmldom = xmlhttp.responseXML
						Set objLst = xmldom.getElementsByTagName("*")
						for ccp_i = 0 to (objLst.length - 1)
							if ccp_testing then
								 
								response.write objLst.item(ccp_i).nodeName & ": "
								response.write objLst.item(ccp_i).text & vbCrLF
							end if
							
							Select Case objLst.item(ccp_i).nodeName
								Case "confirmationNumber"
									newOrderID = objLst.item(ccp_i).text
								Case "authCode"
									'approvalCode = objLst.item(ccp_i).text
								Case "decision"
									if UCASE(objLst.item(ccp_i).text) = "ACCEPTED" then	'Payment Accespted Successfully
										ok = 0
										curMess = "Success"
									elseif UCASE(objLst.item(ccp_i).text) = "ERROR" then
										ok = -1
									elseif UCASE(objLst.item(ccp_i).text) = "DECLINED" then
										ok = -1
									end if
								Case "description"
									'responseStr = objLst.item(ccp_i).text
									errMsg = objLst.item(ccp_i).text
								Case "code"
									'errorCode = objLst.item(ccp_i).text
							End Select
						next

					else	'Processor failed to response or unhandled error
						if ccp_testing then
							Response.Write "The request was unsuccessful.<br /><br />" & vbCrLF
							Response.Write "status:" & xmlhttp.status & "<br /><br />" & vbCrLF
							Response.write xmlhttp.statusText
						end if
						reasonText = "There was a problem communicating with the gateway."
					end if					
				
					set xmlhttp = nothing
					set xmldom = nothing
					
					if ok<>0 AND errMsg<>"" then
						curMess = curMess & ": " & errMsg
					end if
				end if
				if ok = 0 then
						intcount = intcount + 1
						cashCount = cashCount + rsEntry("ccAmt")
						
						strSQL = "Update tblCCTrans SET Settled=1, BatchNumber=" & newBatchNum & ", TransTime=" & DateSep & Now & DateSep
						strSQL = strSQL & ", ccNum='xxxxxxxxxxxxxxxx'"
	
						'Not a good idea for MON... need to keep original orderID/trx_num
						'Add back on 7_26_07 - this number is required to refund must be udpated
						if newOrderID<>"" then
							strSQL = strSQL & ", OrderID=N'" & newOrderID & "'"
						end if
						
						if op_TxnNumber<>"" then
							strSQL = strSQL & ", op_TxnNumber = " & op_TxnNumber
						end if
	
						strSQL = strSQL & " WHERE TransactionNumber=" & rsEntry("TransactionNumber")
						cnWS.execute strSQL
				end if

					if rowcount=1 then
					   rowcount=0
%>
                    <tr bgcolor=#FAFAFA> 
<%
					elseif rowcount=0 then
					   rowcount=1
%>
                    <tr bgcolor=#F2F2F2> 
<%
					end if
%>
                      <td nowrap>&nbsp;&nbsp;&nbsp;<%=rsEntry("TransactionNumber")%></td>
                      <td nowrap>&nbsp;<%=rsEntry("FirstName")%>&nbsp;<%=rsEntry("LastName")%></td>
                      <td nowrap align="center">&nbsp;<%=FormatNumber(rsEntry("ccAmt")*.01,2)%></td>
					  <td nowrap align="center">
						<%'if ccProcessor<>"PMN" then 'CB Removed 12_19_2007%>
						  	<%if curMess="Success" then response.write newBatchNum end if%>
						<%'end if %>
					  </td>
                      <td nowrap align="center">&nbsp;<%=curMess%></td>
                    </tr>
                    <tr style="background-color:<%=session("pageColor4")%>;"> 
                      <td colspan=5><img height="1" width="100%" src="<%= contentUrl("/asp/adm/images/trans.gif") %>"></td>
                    </tr>
<%
				if ss_UseEFT then
					if NOT isNull(rsEntry("StatusCode")) then
						if rsEntry("StatusCode")=3 then
							updateStat 6, 1		'''Increment EFT Trans Metric
						end if
					end if
				end if
			end if	'Selected for this batch/settle

			rsEntry.MoveNext 
    
        Loop
    
		if intcount > 0 then	'transactions successful
			if ccProcessor="CCP" then
				result = settler.sendBatchNow()

					'returns - transcount, totalTranAmount
					'0, 0.00 when failed
					
			elseif ccProcessor="PMN" then
				settler.BatchNumber = newBatchNum
				settler.MerchantPIN = "GNKFMaNP23JKii9HAhrLQ8UqxjL7QagF"
				settler.BatchSettle

				responseSTR = ""
				if NOT isNULL(settler.ResultAuthCode) then
					responseSTR = settler.ResultAuthCode
				end if
				
				strSQL = "INSERT INTO tblCCBatches (BatchNumber, ResponseStr, BatchTime) "
				strSQL = strSQL & "VALUES (" & newBatchNum & ", N'" & sqlInjectStr(responseSTR) & "', GetDate())"
				cnWS.execute strSQL
				
			elseif ccProcessor="MON" then
				Dim batchClose
				Set monerisReq = server.CreateObject("Moneris.Request")
			    'TESTING URL!!
			    'monerisReq.initRequest cMID, cTID, "https://esplusqa.moneris.com/gateway2/servlet/MpgRequest"	'storeID,api_token
				monerisReq.initRequest cMID, cTID, "https://www3.moneris.com/gateway2/servlet/MpgRequest"
				
				Set batchClose = server.CreateObject("Moneris.BatchClose")
				monerisReq.setRequest batchClose.formatRequest( Moneris_ECR )
				monerisReq.sendRequest
				
			'these may not apply to batch close
			'Response.Write "start out<br />"	
			'Response.Write "Receipt ID:  " & monerisReq.getReceiptID & "<br />"
			'Response.Write "Response Code:  " & monerisReq.getResponseCode & "<br />"
			'Response.Write "Transaction Type:  " & monerisReq.getTransType & "<br />"
			'Response.Write "Message:  " & monerisReq.getMessage & "<br />"
			'Response.Write "Amount:  " & monerisReq.getTransAmount & "<br />"
			'Response.Write "Bank Totals:  " & monerisReq.getBankTotals & "<br />"
			'Response.Write "Card Type:  " & monerisReq.getCardType & "<br />"
			'Response.Write "Reference Number:  " & monerisReq.getReferenceNum & "<br />"
			'Response.Write "Transaction ID:  " & monerisReq.getTransID & "<br />"
			'Response.Write "ISO:  " & monerisReq.getISO & "<br />"
			'Response.Write "Auth Code:  " & monerisReq.getAuthCode & "<br />"
			'Response.Write "Transaction Time:  " & monerisReq.getTransTime & "<br />"
			'Response.Write "Transaction Date:  " & monerisReq.getTransDate & "<br />"
			'Response.Write "Complete:  " & monerisReq.getCompleteStatus & "<br />"
			'Response.Write "Timeout:  " & monerisReq.getTimedoutStatus & "<br />"
			'Response.Write "Ticket:  " & monerisReq.getTicket & "<br />"
				
			end if
		end if
		
		strSQL = "UPDATE tblCCOpts SET ActiveBatch=0"
		cnWS.execute strSQL

	    rsEntry.close
	    Set rsEntry = Nothing
		
	    cnMB.Close
	    Set cnMB = Nothing
		
%>
                  </table>
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
end if ' end session
%>

