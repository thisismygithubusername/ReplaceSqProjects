<%@ CodePage=65001 %>
<%
'Dim SessionFarm
'set SessionFarm = Server.CreateObject("SessionFarm.SFSession")
response.charset="utf-8"
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<%
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
<!-- #include file="../inc_i18n.asp" -->
<!-- #include file="inc_chk_ss.asp" -->
<%
         dim rsEntry, useEFT, result, reasonText, xmlmsg
         set rsEntry = Server.CreateObject("ADODB.Recordset")
		 useEFT = checkStudioSetting("Studios", "UseEFT")
			result = 1
		Dim CCProcessor : CCProcessor = ""
		Dim CCProcessor2 : CCProcessor2 = null
		strSQL = "SELECT CCProcessor, CCProcessor2 FROM tblCCOpts WHERE StudioID=" & session("StudioID")
		rsEntry.CursorLocation = 3
		rsEntry.open strSQL, cnWS
		Set rsEntry.ActiveConnection = Nothing
		if NOT rsEntry.EOF then
			if NOT isNULL(rsEntry("CCProcessor")) then
				CCProcessor = TRIM(rsEntry("CCProcessor"))
			end if
			if NOT isNULL(rsEntry("CCProcessor2")) then
				CCProcessor2 = TRIM(rsEntry("CCProcessor2"))
			end if
		end if
		rsEntry.close
		

         Dim strTempName, intCount, cashCount, settler, tmpResponseArr, tmpAmt
         dim ccp_ccTestMode : ccp_ccTestMode = checkStudioSetting("Studios","ccTestMode")
         dim ccp_testing : ccp_testing = false

		strSQL = "SELECT tblCCTrans.Status, tblCCTrans.Settled, tblCCTrans.ccAmt, tblCCTrans.ClientID, tblCCTrans.ccNum, tblCCTrans.ExpMoYr, tblCCTrans.TransactionNumber, tblCCTrans.AuthCode, tblCCTrans.TransTime, tblCCTrans.MerchantID, tblCCTrans.TerminalID, tblCCTrans.OrderID, CLIENTS.LastName, CLIENTS.FirstName FROM CLIENTS, tblCCTrans "
		strSQL = strSQL & "WHERE CLIENTS.ClientID = tblCCTrans.ClientID AND tblCCTrans.Settled=0 "
		rsEntry.CursorLocation = 3
		rsEntry.open strSQL, cnWS
		Set rsEntry.ActiveConnection = Nothing

		rowcount = 0    
		intcount = 0
		cashCount = 0

		Do While NOT rsEntry.EOF

			if request.form("chk_"&rsEntry("TransactionNumber"))="on" then
						
					

					if CCProcessor="PMN" then
						if NOT isNULL(rsEntry("MerchantID")) AND NOT isNULL(rsEntry("ccAmt")) AND NOT IsNULL(rsEntry("OrderID")) AND NOT isNULL(rsEntry("AuthCode")) then

							tmpAmt = ""
							tmpAmt = TRIM(rsEntry("ccAmt"))
							tmpResponseArr = Split(rsEntry("AuthCode"), "-")
	
							Set voidauth = CreateObject("ATS.SecurePost")
		
							voidauth.ATSID = rsEntry("MerchantID")
							if NOT isNULL(rsEntry("TerminalID")) then
								voidauth.ATSSubID = rsEntry("TerminalID")
							end if
							voidauth.Amount = tmpAmt
							if CCProcessor2 = "FD" then
								voidauth.AuthReverse tmpResponseArr(0)
							else
								voidauth.ProcessVoid tmpResponseArr(0)
							end if
							'response.write "PreAuth Voided Successfully! ID:" & tmpResponseArr(0) & "<br />"
						
							Set voidauth = nothing
						end if
					elseif CCProcessor="TCI" then		
		
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
								xmlmsg = xmlmsg & "<web:VoidCreditCardSale>"
									xmlmsg = xmlmsg & "<web:MerchantID>" & rsEntry("MerchantID") & "</web:MerchantID>"
									xmlmsg = xmlmsg & "<web:RegKey>" & rsEntry("TerminalID") & "</web:RegKey>"
									xmlmsg = xmlmsg & "<web:TransID>" & rsEntry("OrderID") & "</web:TransID>"
								xmlmsg = xmlmsg & "</web:VoidCreditCardSale>"
							xmlmsg = xmlmsg & "</soapenv:Body>"
						xmlmsg = xmlmsg & "</soapenv:Envelope>"
					
						if ccp_testing then
							response.Write xmlmsg
						end if
					
						xmlhttp.open "POST", SoapServer, false
						xmlhttp.setRequestHeader "Content-Type", "text/xml; charset=utf-8"
						xmlhttp.setRequestHeader "Content-Length", Len(xmlmsg)
						if true then 'CC Real Time Processing TODO
								xmlhttp.setRequestHeader "SOAPAction", "http://www.paymentresources.com/webservices/VoidCreditCardSale"
						end if				

						xmlhttp.send xmlmsg

						if xmlhttp.Status = 200 then
							Set xmldom = xmlhttp.responseXML
							Set objLst = xmldom.getElementsByTagName("*")
							
							
							for i = 0 to (objLst.length - 1)
								if ccp_testing then
									response.write objLst.item(i).nodeName & ": "
									response.write objLst.item(i).text & vbCrLF & "<br /> "
								end if
								
								Select Case objLst.item(i).nodeName
									Case "Status"
										if objLst.item(i).text <> "Voided" then	'Void Unsuccesful
											result = -1
										end if
									Case "Message"
										reasonText = objLst.item(i).text
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
					end if
					
					if result = 1 then
						strSQL = "Update tblCCTrans SET Status='Approved (Voided)'"
						strSQL = strSQL & ", TransTime=" & DateSep & Now & DateSep
						strSQL = strSQL & ", SaleID=null"
						strSQL = strSQL & " WHERE TransactionNumber=" & rsEntry("TransactionNumber")
						cnWS.execute strSQL
					else
						Response.Write "Void Failed: " & reasonText
					end if
			end if	'Selected for this batch/settle					
					
		    rsEntry.MoveNext 
        Loop
	
	    rsEntry.Close
	    Set rsEntry = Nothing

	end if
	cnWS.close
	set cnWS = nothing

	end if

	response.redirect "adm_rpt_ccp.asp"
%>
