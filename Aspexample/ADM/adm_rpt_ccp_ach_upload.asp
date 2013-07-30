<%@ CodePage=65001 %>
<!-- #include file="init.asp" -->
<%
'Dim SessionFarm
'set SessionFarm = Server.CreateObject("SessionFarm.SFSession")
response.charset="utf-8"
%>
<%
	%>
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

		Dim Upload, fso, filename, filechars, char, rso, newFileName
		Set Upload = Server.CreateObject("csASPUpload.Process")
		Set fso = CreateObject("Scripting.FileSystemObject")

		%>
		<!-- #include file="../inc_i18n.asp" -->
		<!-- #include file="inc_rpt_tagging.asp" -->
		<!-- #include file="inc_hotword.asp" -->
		<!-- #include file="inc_row_colors.asp" -->		
		<!-- #include file="inc_return_acct.asp" -->
		<%
		VoidOk = validAccessPriv("TB_VOID")
	
		Dim cSDate, cEDate, disMode, ccProcessor, isBatchProcessor, numTrx, numFTPAccounts, ACHHotWord
		ccProcessor = checkStudioSetting("tblCCOpts", "ccProcessor")
		numTrx = 0
		if ccProcessor = "MON" OR ccProcessor="OP" OR ccProcessor = "HSBC" then
			isBatchProcessor = true
		else
			isBatchProcessor = false
		end if
		
		dim rsEntry
		set rsEntry = Server.CreateObject("ADODB.Recordset")
	
		ACHHotWord = "ACH"
		ACHHotWord = getHotWord(109)
		
		' set the row colors
        setRowColors "#F2F2F2", "#FAFAFA"

		%>
<!-- #include file="pre.asp" -->
		<!-- #include file="frame_bottom.asp" -->
<%= js(array("mb", "MBS")) %>
		
		<!-- #include file="../inc_date_ctrl.asp" -->
		<!-- #include file="../inc_ajax.asp" -->
		
<% pageStart %>
			<table height="100%" width="<%=strPageWidth%>" border="0" cellspacing="0" cellpadding="0">    
				<tr>
					<td align="center" valign="top" height="100%" width="100%">
						<table class="center" border="0" cellspacing="0" cellpadding="0" width="90%" height="100%">
							<form name="frmFileUpload" action="adm_rpt_ccp_ach_upload_p.asp?sr=false" method="post" enctype="multipart/form-data">
                                <input type="hidden" name="frmSubmitted" value="<%=session("StudioID")%>" />
                                <input type="hidden" name="frmExpReport" value="">
                                <input type="hidden" name="frmGenReport" value="">
                                <input type="hidden" name="frmResetBatch" value=""/>
								<tr>
									<td class="headText" align="left" valign="top">
										<table class="mainText" width="100%" border="0" cellspacing="0" cellpadding="0">
											<tr>
												<td class="headText" valign="bottom"><b>Upload Results</b></td>
												<td align="right" valign="bottom" height="26">
											</tr>
										</table>
									</td>
								</tr>
								<tr>
									<td valign="bottom" class="mainText" align="right" height="18">
										<b>
											<a href="adm_rpt_ccp.asp"><%= pp_PageTitle("Approved Transactions") %></a> |
											<a href="adm_rpt_ccp_ach.asp"><%= pp_PageTitle("Pending Transactions") %></a> | 
											<a href="adm_rpt_ccp_rej.asp"><%= pp_PageTitle("Voided Rejected Transactions") %></a> | <a href="adm_rpt_ccp_set.asp"><%= pp_PageTitle("Settled Transactions") %></a>&nbsp; 
										</b>
									</td>
								</tr>
								<tr>
									<td valign="top" class="mainTextBig" align="center" height="100%">
										<table class="mainText center" width="95%" border="0" cellspacing="0" cellpadding="0" height="100%">
											<tr>
												<td class="mainText" colspan="2" valign="top" align="center">
													<table class="mainText center" border="0" cellspacing="0" cellpadding="0" width="90%">
														<tr valign="top">
															<td align="center">

													<%if request.QueryString("filename")="" then %>
					                                            <table class="mainText center" width="70%" border="0" cellspacing="0" cellpadding="0">
							                                            <input type="hidden" name="fileAction" value="Upload">
						                                            <tr> 
							                                            <td colspan="14" style="background-color:<%=session("pageColor2")%>;"><img height="2" width="100%" src="<%= contentUrl("/asp/adm/images/trans.gif") %>"></td>
						                                            </tr>
						                                            <tr class="whiteSmallText" style="background-color:<%=Session("pageColor4")%>;"> 
							                                            <td align="left" nowrap><b>&nbsp;Upload New File&nbsp;</b></td>
						                                            </tr>
						                                            <tr> 
							                                            <td colspan="14" style="background-color:<%=session("pageColor2")%>;"><img height="2" width="100%" src="<%= contentUrl("/asp/adm/images/trans.gif") %>"></td>
						                                            </tr>
						                                            <tr class="mainText" style="background-color:<%=getRowColor(true)%>;"> 
							                                            <td nowrap><input type="file" name="resultFile" size="40"></td>
						                                            </tr>
						                                            <tr class="mainText" style="background-color:<%=getRowColor(true)%>;"> 
							                                            <td nowrap><input type="button" value="Upload" onClick="javascript: if(document.frmFileUpload.resultFile.value == '') { alert('Please select a file.'); document.frmFileUpload.resultFile.focus(); } else { document.frmFileUpload.submit(); }"></td>
						                                            </tr>
						                                            <tr> 
							                                            <td colspan="14" style="background-color:<%=session("pageColor4")%>;"><img height="2" width="100%" src="<%= contentUrl("/asp/adm/images/trans.gif") %>"></td>
						                                            </tr>
					                                            </table>

                                                    <% else 'process file %>





					                                            <table class="mainText center" width="70%" border="0" cellspacing="0" cellpadding="0">
						                                            <tr><td colspan="14" style="background-color:<%=session("pageColor2")%>;"><img height="2" width="100%" src="<%= contentUrl("/asp/adm/images/trans.gif") %>"></td></tr>
						                                            <tr class="whiteSmallText" style="background-color:<%=Session("pageColor4")%>;"> 
							                                            <td colspan="14" align="left" nowrap><b>&nbsp;File Upload Results&nbsp;</b></td>
						                                            </tr>
						                                            <tr><td colspan="14" style="background-color:<%=session("pageColor2")%>;"><img height="2" width="100%" src="<%= contentUrl("/asp/adm/images/trans.gif") %>"></td></tr>
<%
												        set objStream = Server.CreateObject("ADODB.Stream")
												        objStream.Open
												        objStream.Type = 2 'text
												        objStream.CharSet = "ascii"
												        objStream.LoadFromFile studio_path & session("studioShort") & "\" & request.querystring("filename")

										                tmpLine = objStream.ReadText(-2)

                                                        tmpFileType = "unknown"
										                if LEFT(tmpLine,3) = "030"  then    'VISA
											                tmpFileType = "VISA"
											                BatchNumber = Mid(tmpLine, 18, 20)
										                elseif LEFT(tmpLine,3) = "TFH" then
											                tmpFileType = "AMEX"
											              elseif LEFT(tmpLine,3) = "010" then
																			tmpFileType = "AMEXTW"
																		elseif LEFT(tmpLine,1) = "H" then
											                tmpFiletype = "VISATW"
										                end if
										                
										                dim dispFileType
										                select case tmpFileType
																			case "AMEXTW"
																				dispFileType = "AMEX"
																			case "VISATW"
																				dispFileType = "VISA"
																			case else
																				dispFileType = tmpFileType
																		end select
%>
						                                            <tr class="mainText" style="background-color:<%=getRowColor(true)%>;">
						                                                <td width="150" nowrap>File Type: <strong><%=dispFileType%></strong></td>
						                                                <td nowrap><strong>Transaction Number</strong></td>
						                                                <td nowrap><strong>Status</strong></td>
						                                            </tr>
						                                            <tr><td colspan="14" style="background-color:<%=session("pageColor4")%>;"><img height="1" width="100%" src="<%= contentUrl("/asp/adm/images/trans.gif") %>"></td></tr>
<%
                                                        objStream.Close
												        objStream.Open
												        objStream.LoadFromFile studio_path & session("studioShort") & "\" & request.querystring("filename")                                                       

                                                        tmpCounter = 0
                                                        tmpCounterSuccess = 0
                                                        
												        do while NOT objStream.EOS
													        
													        'read in line
													        if tmpFileType = "VISA" then    'read 129 chars at a time
    													        tmpLine = objStream.ReadText(130)
    													            													        
    													        if Mid(tmpLine, 81, 8) = "ACCEPTED" then
    													            ccTransStatus = "Approved"
        													        ccTransID = TRIM(Mid(tmpLine, 49, 15))
    													        else
    													            ccTransStatus = LEFT("Declined: " & TRIM(Mid(tmpLine, 59, 27)), 30)
        													        ccTransID = TRIM(Mid(tmpLine, 44, 15))
    													        end if
    													    elseif tmpFileType = "VISATW" then    'read line at a time
    													        tmpLine = objStream.ReadText(-2)
    													        if LEFT(tmpLine,1) = "D" then
    													            ccTransStatus = fmtresponse(Mid(tmpLine, 82, 2))
        													        ccTransID = TRIM(Mid(tmpLine, 104, 20))
        													    end if
													        elseif tmpFileType = "AMEXTW" then    'read line at a time
    													        tmpLine = objStream.ReadText(-2)
    													        if LEFT(tmpLine,2) = "03" then
																	if TRIM(Mid(tmpLine, 351, 8)) = "Approved" then
    																	ccTransStatus = "Approved"
																	else
    																	ccTransStatus = fmtresponseAmexTw(Mid(tmpLine, 71, 3))
																	end if
        													        ccTransID = TRIM(Mid(tmpLine, 76, 7))
        													    end if
    													        
													        else    'read line at a time
													            
    													        tmpLine = objStream.ReadText(-2)
    													        
    													        'AMEX - HK/SNG
    													        ccTransID = Mid(tmpLine, 12, 15)
    													        if Mid(tmpLine, 265, 8) = "Approved" then
    													            ccTransStatus = "Approved"
    													        else
    													            ccTransStatus = "Declined: " & TRIM(Mid(tmpLine, 265, 28))
    													        end if
    													        
    													        
													        end if
													        
													        
																	if (tmpFileType = "VISA" AND (Mid(tmpLine, 1, 1)="0" OR Mid(tmpLine, 1, 1)="Z" OR Mid(tmpLine, 1, 1)="T")) OR (tmpFileType="AMEX" AND Mid(tmpLine, 1, 3)<>"TAB") OR (tmpFileType="VISATW" AND Mid(tmpLine, 1, 1)<>"D") OR (tmpFileType="AMEXTW" AND LEFT(tmpLine,2)<>"03")then 'skip header & footer
																	else


																			if NOT isNum(ccTransID) then
																					ccTransStatus = "Invalid Transaction Number: " & ccTransID
																			else

																	tmpCounter = tmpCounter + 1

																	if ccTransStatus <> "Approved" then
																		'Get Info About Sale and Convert to Account Debit
																		strSQL = "SELECT tblPayments.PaymentID FROM tblPayments INNER JOIN tblCCTrans ON tblPayments.CCTransID = tblCCTrans.TransactionNumber WHERE (tblPayments.CCTransID = " & ccTransID & ") AND (NOT (tblCCTrans.Status LIKE N'Decline%'))"
																		rsEntry.CursorLocation = 3
																		rsEntry.open strSQL, cnWS
																		Set rsEntry.ActiveConnection = Nothing
																		if NOT rsEntry.EOF then
																			returnToAccount rsEntry("PaymentID"), false, 0.00
																		end if
																		rsEntry.close
																	end if

                                                                    strSQL = "UPDATE tblCCTrans SET Settled="
                                                                    if ccTransStatus = "Approved" then
                                                                       strSQL = strSQL & "1"
                                                                        tmpCounterSuccess = tmpCounterSuccess + 1
                                                                    else
                                                                       strSQL = strSQL & "0"
                                                                    end if
                                                                    strSQL = strSQL & ", Status=N'" & sqlInjectStr(ccTransStatus) & "' "
                                                                    strSQL = strSQL & "WHERE TransactionNumber=" & ccTransID
                                                                   response.write debugSQL(strSQL, "SQL")
                                                                    cnWS.execute strSQL

%>
						                                            <tr class="mainText" style="background-color:<%=getRowColor(true)%>;">
						                                                <td nowrap><%=tmpCounter %>.</td>
						                                                <td nowrap><%=ccTransID%></td>						                                                
						                                                <td nowrap><%=ccTransStatus%></td>
						                                            </tr>
<%
                                                                end if 'isNum(ccTransID)
                                                            end if  'Skip Header / Footer
                                                        loop
%>
						                                            <tr> 
							                                            <td colspan="14" style="background-color:<%=session("pageColor4")%>;"><img height="2" width="100%" src="<%= contentUrl("/asp/adm/images/trans.gif") %>"></td>
						                                            </tr>
						                                            <tr class="mainText">
						                                                <td colspan="14" nowrap>Approved: <%=tmpCounterSuccess%> &nbsp;&nbsp;&nbsp; Declined: <%=tmpCounter - tmpCounterSuccess %></td>
						                                            </tr>
					                                            </table>

                                                    <% end if %>
																
															</td>
														</tr>
													</table>
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

		Set Upload = nothing


end if
%>

<%
function fmtresponse(rescode)
    Select Case rescode
        Case "00"
        fmtresponse=("Approved")
        Case "01"
        fmtresponse=("Call Bank")
        Case "02"
        fmtresponse=("Call Bank")
        Case "03"
        fmtresponse=("Decline: Invalid Merchant")
        Case "04"
        fmtresponse=("Pickup")
        Case "05"
        fmtresponse=("Decline: Do Not Honor")
        Case "07"
        fmtresponse=("Decline: HOLD CALL, Pickup Card")
        Case "11"
        fmtresponse=("Approved")
        Case "12"
        fmtresponse=("Decline: Invalid Trans Type")
        Case "13"	
        fmtresponse=("Decline: Invalid Amount")
        Case "14"		
        fmtresponse=("Decline: Invalid card number (no such number)")
        Case "19"		
        fmtresponse=("Try Again: Re-Enter Transaction	Try again")
        Case "21"		
        fmtresponse=("Decline: NO ACTION TAKEN")
        Case "30"	
        fmtresponse=("Decline: Format error")
        Case "31"	
        fmtresponse=("Decline: Bank not supported by switch")
        Case "41"	
        fmtresponse=("Decline: Pick-up Card for Lost card")
        Case "43"	
        fmtresponse=("Decline: Pick-up Card for Stolen card, Pickup")
        Case "51"	
        fmtresponse=("Decline: Not sufficient funds")
        Case "54"	
        fmtresponse=("Decline: Expired card or card expire date error")
        Case "58"	
        fmtresponse=("Decline: Card has not been activated or others ")
        Case "61"	
        fmtresponse=("Decline: Exceeds withdrawal amount limit")
        Case "63"	
        fmtresponse=("Decline: CVC error")
        Case "75"	
        fmtresponse=("Decline: Allowable number of PIN tries exceeded")
        Case "77"	
        fmtresponse=("Decline: Reversal data are inconsistent with Original Message")
        Case "76"	
        fmtresponse=("Decline: REFERENCE ERROR or product code error")
        Case "79"	
        fmtresponse=("Decline: ALREADY REVERSED")
        Case "85"	
        fmtresponse=("Decline: Card ok, declined by issue bank ")
        Case "89"	
        fmtresponse=("Decline: System error (reserved for BASE24 use)")
        Case "90"	
        fmtresponse=("Decline: Cut-off is in process (switch ending a day?s business and starting Then ext. Transaction can be sent again in a few minutes)")
        Case "92"	
        fmtresponse=("Decline: BIN Error")
        Case "91"	
        fmtresponse=("Decline: Issuer or switch is inoperative")
        Case "94"	
        fmtresponse=("Decline: Duplicate transaction")
        Case "96"	
        fmtresponse=("Decline: System Malfunction")
        Case "99"	
        fmtresponse=("Approved")
		Case else
		fmtresponse=("Approved")
    End Select
end function

function fmtresponseAmexTw(rescode)
	Select Case rescode
		Case "992"
		fmtresponseAmexTw = ("Decline: Zero Transaction Amount ")
		Case "995"
		fmtresponseAmexTw = ("Decline: STM Account is not enrolled with AMEX ")
		Case "998"
		fmtresponseAmexTw = ("Decline: No Authorisation Response Received ")
		Case "000"
		fmtresponseAmexTw = ("Decline: Code 16000 not assigned ")
		Case "100"
		fmtresponseAmexTw = ("Decline: Terminal not found on file ")
		Case "101"
		fmtresponseAmexTw = ("Decline: Error in reading magnetic stripe ")
		Case "102"
		fmtresponseAmexTw = ("Decline: Error in stat message ")
		Case "103"
		fmtresponseAmexTw = ("Decline: Found duplicate sequence number ")
		Case "104"
		fmtresponseAmexTw = ("Decline: Terminal not active ")
		Case "107"
		fmtresponseAmexTw = ("Decline: Card capture type invalid ")
		Case "108"
		fmtresponseAmexTw = ("Decline: Amount is bad ")
		Case "109"
		fmtresponseAmexTw = ("Decline: Cardmember number is bad ")
		Case "110"
		fmtresponseAmexTw = ("Decline: Date/time is bad ")
		Case "111"
		fmtresponseAmexTw = ("Decline: Service establishment number is bad ")
		Case "112"
		fmtresponseAmexTw = ("Decline: Sequence number is bad ")
		Case "113"
		fmtresponseAmexTw = ("Decline: Transaction code is bad ")
		Case "114"
		fmtresponseAmexTw = ("Decline: Process code is bad ")
		Case "115"
		fmtresponseAmexTw = ("Decline: Effective date is bad ")
		Case "116"
		fmtresponseAmexTw = ("Decline: Expiration date is bad ")
		Case "117"
		fmtresponseAmexTw = ("Decline: Track two is bad ")
		Case "118"
		fmtresponseAmexTw = ("Decline: Product code is bad ")
		Case "119"
		fmtresponseAmexTw = ("Decline: Record of Charge invoice is bad ")
		Case "120"
		fmtresponseAmexTw = ("Decline: Input/output error ")
		Case "121"
		fmtresponseAmexTw = ("Decline: Old Record of Charge not found ")
		Case "122"
		fmtresponseAmexTw = ("Decline: Reversal not found ")
		Case "123"
		fmtresponseAmexTw = ("Decline: Void of a reverse ")
		Case "124"
		fmtresponseAmexTw = ("Decline: RRN is bad ")
		Case "125"
		fmtresponseAmexTw = ("Decline: Response sent as approved, inquiry only ")
		Case "126"
		fmtresponseAmexTw = ("Decline: Host Unavailable ")
		Case "127"
		fmtresponseAmexTw = ("Decline: No parms found ")
		Case "128"
		fmtresponseAmexTw = ("Decline: Length in message is bad ")
		Case "129"
		fmtresponseAmexTw = ("Decline: Input source name equal spaces ")
		Case "130"
		fmtresponseAmexTw = ("Decline: Invalid cipher type passed ")
		Case "131"
		fmtresponseAmexTw = ("Decline: AMEX mod 10 number is invalid ")
		Case "132"
		fmtresponseAmexTw = ("Decline: Cardnumber does not start with a 34 or 37 ")
		Case "133"
		fmtresponseAmexTw = ("Decline: Personal Identification Number is invalid ")
		Case "134"
		fmtresponseAmexTw = ("Decline: Exceeded the max num of tries in entering PIN ")
		Case "135"
		fmtresponseAmexTw = ("Decline: Sequence number is zero ")
		Case "136"
		fmtresponseAmexTw = ("Decline: Cardnumber is zeroes ")
		Case "137"
		fmtresponseAmexTw = ("Decline: Card not valid for terminal ")
		Case "138"
		fmtresponseAmexTw = ("Decline: Maximum amount exceeded ")
		Case "139"
		fmtresponseAmexTw = ("Decline: Record of Charge invoice number is zero ")
		Case "140"
		fmtresponseAmexTw = ("Decline: Statement of Charges invoice number is zero ")
		Case "141"
		fmtresponseAmexTw = ("Decline: Tip amount invalid ")
		Case "142"
		fmtresponseAmexTw = ("Decline: Statement of Charges batch number is bad ")
		Case "150"
		fmtresponseAmexTw = ("Decline: Invalid transaction requested for fsi ")
		Case "151"
		fmtresponseAmexTw = ("Decline: Unexpected message type identifier ")
		Case "152"
		fmtresponseAmexTw = ("Decline: FSI invalid (not found or not active) ")
		Case "153"
		fmtresponseAmexTw = ("Decline: A key exchange reply was received ")
		Case "154"
		fmtresponseAmexTw = ("Decline: No keys are in the FSI record to exchange ")
		Case "155"
		fmtresponseAmexTw = ("Decline: FSI record does not allowed key exchanges ")
		Case "156"
		fmtresponseAmexTw = ("Decline: Card reported lost ")
		Case "157"
		fmtresponseAmexTw = ("Decline: Card reported stolen ")
		Case "158"
		fmtresponseAmexTw = ("Decline: Approved for partial amount ")
		Case "159"
		fmtresponseAmexTw = ("Decline: Approved, cardholder is a VIP ")
		Case "160"
		fmtresponseAmexTw = ("Decline: Invalid reconciliation count ")
		Case "161"
		fmtresponseAmexTw = ("Decline: Card acceptor contact card acquirer ")
		Case "162"
		fmtresponseAmexTw = ("Decline: Card is restricted card ")
		Case "163"
		fmtresponseAmexTw = ("Decline: Card acceptor call acquirer's security ")
		Case "164"
		fmtresponseAmexTw = ("Decline: Cardholder contact card issuer ")
		Case "165"
		fmtresponseAmexTw = ("Decline: Bank not supported by switch ")
		Case "166"
		fmtresponseAmexTw = ("Decline: Transaction not permitted to terminal ")
		Case "167"
		fmtresponseAmexTw = ("Decline: Transaction not completed, legal violation ")
		Case "168"
		fmtresponseAmexTw = ("Decline: Re-enter entire transaction ")
		Case "169"
		fmtresponseAmexTw = ("Decline: Cutover (settlement) in progress ")
		Case "170"
		fmtresponseAmexTw = ("Decline: Issuer or switch is inoperative ")
		Case "171"
		fmtresponseAmexTw = ("Decline: Financial institution can not be found ")
		Case "172"
		fmtresponseAmexTw = ("Decline: Not sufficient funds ")
		Case "173"
		fmtresponseAmexTw = ("Decline: Transaction not permitted to cardholder ")
		Case "176"
		fmtresponseAmexTw = ("Decline: Cipher key not found ")
		Case "177"
		fmtresponseAmexTw = ("Decline: Cipher convert key not found ")
		Case "178"
		fmtresponseAmexTw = ("Decline: Bad length in pin block ")
		Case "179"
		fmtresponseAmexTw = ("Decline: Bad user data length ")
		Case "180"
		fmtresponseAmexTw = ("Decline: Bad fill type ")
		Case "181"
		fmtresponseAmexTw = ("Decline: Bad block type ")
		Case "182"
		fmtresponseAmexTw = ("Decline: The zone key does not exist ")
		Case "183"
		fmtresponseAmexTw = ("Decline: There is no master key ")
		Case "184"
		fmtresponseAmexTw = ("Decline: Cipher DBIO error ")
		Case "185"
		fmtresponseAmexTw = ("Decline: The resource array is full ")
		Case "186"
		fmtresponseAmexTw = ("Decline: The replacement string contains invalid chars ")
		Case "187"
		fmtresponseAmexTw = ("Decline: Error in verifying message authentication code ")
		Case "188"
		fmtresponseAmexTw = ("Decline: Link State is not operational ")
		Case "189"
		fmtresponseAmexTw = ("Decline: TCU session keys update error ")
		Case "190"
		fmtresponseAmexTw = ("Decline: TCU terminal initialisation error ")
		Case "191"
		fmtresponseAmexTw = ("Decline: Update out of market updates not allowed ")
		Case "192"
		fmtresponseAmexTw = ("Decline: Password was used previously ")
		Case "193"
		fmtresponseAmexTw = ("Decline: Security DBIO error ")
		Case "194"
		fmtresponseAmexTw = ("Decline: Update out of application not allowed ")
		Case "195"
		fmtresponseAmexTw = ("Decline: Update out of city not allowed ")
		Case "196"
		fmtresponseAmexTw = ("Decline: Update out of country updates not allowed ")
		Case "197"
		fmtresponseAmexTw = ("Decline: Invalid authority level ")
		Case "198"
		fmtresponseAmexTw = ("Decline: New passwords do not match ")
		Case "199"
		fmtresponseAmexTw = ("Decline: Current password has expired ")
		Case "200"
		fmtresponseAmexTw = ("Decline: System not available ")
		Case "201"
		fmtresponseAmexTw = ("Decline: MAA save not in chain ")
		Case "202"
		fmtresponseAmexTw = ("Decline: Bad message id in MAA save ")
		Case "203"
		fmtresponseAmexTw = ("Decline: Bad pointer in MAA save ")
		Case "230"
		fmtresponseAmexTw = ("Decline: Unable to allocate in ISO utility ")
		Case "231"
		fmtresponseAmexTw = ("Decline: Expiration date error in ISO utility ")
		Case "232"
		fmtresponseAmexTw = ("Decline: Track 1 error in ISO utility ")
		Case "233"
		fmtresponseAmexTw = ("Decline: Privilege use field 1 error in ISO utility ")
		Case "234"
		fmtresponseAmexTw = ("Decline: Privilege use field 2 error in ISO utility ")
		Case "235"
		fmtresponseAmexTw = ("Decline: Privilege use field 3 error in ISO utility ")
		Case "236"
		fmtresponseAmexTw = ("Decline: Privilege use field 4 error in ISO utility ")
		Case "237"
		fmtresponseAmexTw = ("Decline: Terminal parameter error in ISO utility ")
		Case "238"
		fmtresponseAmexTw = ("Decline: Audit data error in ISO utility ")
		Case "239"
		fmtresponseAmexTw = ("Decline: Header length error in ISO utility ")
		Case "301"
		fmtresponseAmexTw = ("Decline: Record already exists on add by dbio ")
		Case "302"
		fmtresponseAmexTw = ("Decline: Record not found on file by dbio ")
		Case "303"
		fmtresponseAmexTw = ("Decline: Record not written, changed since last access ")
		Case "399"
		fmtresponseAmexTw = ("Decline: Unable to process request by dbio ")
		Case "400"
		fmtresponseAmexTw = ("Decline: DC Route/Edit table error ")
		Case "401"
		fmtresponseAmexTw = ("Decline: Bit map table error ")
		Case "402"
		fmtresponseAmexTw = ("Decline: Required resource for process not found on resource array ")
		Case "403"
		fmtresponseAmexTw = ("Decline: Response code table error ")
		Case "404"
		fmtresponseAmexTw = ("Decline: Vendor table path function not found on process resource ")
		Case "405"
		fmtresponseAmexTw = ("Decline: Error code table error ")
		Case "406"
		fmtresponseAmexTw = ("Decline: Card definition table error ")
		Case "407"
		fmtresponseAmexTw = ("Decline: Vendor table source type not found for vendor transaction ")
		Case "408"
		fmtresponseAmexTw = ("Decline: Table input exceeds the defined table size ")
		Case "409"
		fmtresponseAmexTw = ("Decline: DES key source name not found on application configuration ")
		Case "410"
		fmtresponseAmexTw = ("Decline: DES key name not found on DES key table ")
		Case "411"
		fmtresponseAmexTw = ("Decline: Link counter table is full ")
		Case "412"
		fmtresponseAmexTw = ("Decline: Required vendor for process not found in vendor table ")
		Case "413"
		fmtresponseAmexTw = ("Decline: Currency Code not found or Invalid ")
		Case "414"
		fmtresponseAmexTw = ("Decline: Table load error - Current table not loaded ")
		Case "420"
		fmtresponseAmexTw = ("Decline: Network request message accepted ")
		Case "421"
		fmtresponseAmexTw = ("Decline: Reversal request message accepted ")
		Case "488"
		fmtresponseAmexTw = ("Decline: On error count exceeded ")
		Case "500"
		fmtresponseAmexTw = ("Decline: Batch number is bad on Data Capture process ")
		Case "501"
		fmtresponseAmexTw = ("Decline: Batch # length invalid on Data Capture process ")
		Case "502"
		fmtresponseAmexTw = ("Decline: Batch error in reconciliation ")
		Case "503"
		fmtresponseAmexTw = ("Decline: Totals length invalid ")
		Case "504"
		fmtresponseAmexTw = ("Decline: Batch already open ")
		Case "505"
		fmtresponseAmexTw = ("Decline: No closed Statement of Charges slots ")
		Case "506"
		fmtresponseAmexTw = ("Decline: No suspended Statement of Charges slots ")
		Case "507"
		fmtresponseAmexTw = ("Decline: Bad reconciliation, send detail ")
		Case "508"
		fmtresponseAmexTw = ("Decline: Bad reconciliation, no detail available ")
		Case "509"
		fmtresponseAmexTw = ("Decline: Terminal identification is bad ")
		Case "510"
		fmtresponseAmexTw = ("Decline: Batch not found ")
		Case "511"
		fmtresponseAmexTw = ("Decline: Format error on reconciliation ")
		Case "512"
		fmtresponseAmexTw = ("Decline: Batch balance error on closed suspend ")
		Case "513"
		fmtresponseAmexTw = ("Decline: Bit map error on input ")
		Case "514"
		fmtresponseAmexTw = ("Decline: Bit map error on output ")
		Case "515"
		fmtresponseAmexTw = ("Decline: Totals not numeric ")
		Case "516"
		fmtresponseAmexTw = ("Decline: Totals maximum exceeded ")
		Case "517"
		fmtresponseAmexTw = ("Decline: Bad business date ")
		Case "518"
		fmtresponseAmexTw = ("Decline: Terminal WORK/HOST SUSPENDED SOC cannot be active ")
		Case "600"
		fmtresponseAmexTw = ("Decline: DCP queue not available ")
		Case "601"
		fmtresponseAmexTw = ("Decline: TSS System internal error ")
		Case "602"
		fmtresponseAmexTw = ("Decline: Invalid pin ")
		Case "603"
		fmtresponseAmexTw = ("Decline: Pin errors exceeded ")
		Case "604"
		fmtresponseAmexTw = ("Decline: Not enrolled in TSS ")
		Case "605"
		fmtresponseAmexTw = ("Decline: Invalid card number ")
		Case "606"
		fmtresponseAmexTw = ("Decline: Velocity limit exceeded ")
		Case "607"
		fmtresponseAmexTw = ("Decline: Format error ")
		Case "608"
		fmtresponseAmexTw = ("Decline: Negative item present on account ")
		Case "609"
		fmtresponseAmexTw = ("Decline: Cancel code present on account ")
		Case "610"
		fmtresponseAmexTw = ("Decline: Supp denied by basic ")
		Case "611"
		fmtresponseAmexTw = ("Decline: Invalid card number ")
		Case "612"
		fmtresponseAmexTw = ("Decline: Inhibit code not changed ")
		Case "613"
		fmtresponseAmexTw = ("Decline: SE No or SE type not found on se file/ SE type table ")
		Case "614"
		fmtresponseAmexTw = ("Decline: Invalid SE option Check trap_card_def table ")
		Case "615"
		fmtresponseAmexTw = ("Decline: Amount Over Floor Limit ")
		Case "616"
		fmtresponseAmexTw = ("Decline: New Card issued to Card Member ")
		Case "617"
		fmtresponseAmexTw = ("Decline: Deny_Acct_Cancelled ")
		Case "618"
		fmtresponseAmexTw = ("Decline: Forced Denial ")
		Case "619"
		fmtresponseAmexTw = ("Decline: Amex inhibit, enrollment denied negative ")
		Case "620"
		fmtresponseAmexTw = ("Decline: Amex inhibit, enrollment denied delinquent ")
		Case "630"
		fmtresponseAmexTw = ("Decline: Amex inhibit, credit ")
		Case "650"
		fmtresponseAmexTw = ("Decline: Amex transmit error ")
		Case "651"
		fmtresponseAmexTw = ("Decline: Amex CAS communications down ")
		Case "652"
		fmtresponseAmexTw = ("Decline: Amex CAS down ")
		Case "653"
		fmtresponseAmexTw = ("Decline: TRAP request to TPFE found invalid ")
		Case "654"
		fmtresponseAmexTw = ("Decline: TRAP command process down ")
		Case "660"
		fmtresponseAmexTw = ("Decline: Amex CAS reply- please wait ")
		Case "661"
		fmtresponseAmexTw = ("Decline: Amex CAS reply- invalid card type ")
		Case "662"
		fmtresponseAmexTw = ("Decline: Amex CAS reply- invalid merchant ")
		Case "663"
		fmtresponseAmexTw = ("Decline: Amex CAS reply- use before effective date ")
		Case "664"
		fmtresponseAmexTw = ("Decline: Amex CAS reply- not on file 1 ")
		Case "665"
		fmtresponseAmexTw = ("Decline: Amex CAS reply- invalid function code ")
		Case "666"
		fmtresponseAmexTw = ("Decline: Amex CAS reply- invalid auth amount ")
		Case "667"
		fmtresponseAmexTw = ("Decline: Amex CAS reply- invalid card number ")
		Case "668"
		fmtresponseAmexTw = ("Decline: Amex CAS reply-auth request denied ")
		Case "669"
		fmtresponseAmexTw = ("Decline: Amex CAS reply- invalid auth code ")
		Case "670"
		fmtresponseAmexTw = ("Decline: Amex CAS reply- expired card ")
		Case "671"
		fmtresponseAmexTw = ("Decline: Amex CAS reply- prohibits on card ")
		Case "672"
		fmtresponseAmexTw = ("Decline: Amex CAS reply- bad status ")
		Case "673"
		fmtresponseAmexTw = ("Decline: Amex CAS reply- service error ")
		Case "674"
		fmtresponseAmexTw = ("Decline: Amex CAS reply- call issuer ")
		Case "675"
		fmtresponseAmexTw = ("Decline: Amex CAS reply- call issuer with code ")
		Case "676"
		fmtresponseAmexTw = ("Decline: Amex CAS reply- not on file 2 ")
		Case "677"
		fmtresponseAmexTw = ("Decline: Amex CAS reply- timeout ")
		Case "678"
		fmtresponseAmexTw = ("Decline: Amex CAS reply- try later ")
		Case "680"
		fmtresponseAmexTw = ("Decline: Amex CAS reply- approve with positive id ")
		Case "681"
		fmtresponseAmexTw = ("Decline: Amex CAS reply- CID invalid ")
		Case "682"
		fmtresponseAmexTw = ("Decline: Amex CAS reply code unknown to TPFE ")
		Case "683"
		fmtresponseAmexTw = ("Decline: Invalid slot number received ")
		Case "684"
		fmtresponseAmexTw = ("Decline: SE limit %-ratio data invalid or missing ")
		Case "685"
		fmtresponseAmexTw = ("Decline: AMEX CAS-ref queue number invalid/out-of-range ")
		Case "686"
		fmtresponseAmexTw = ("Decline: AMEX CAS-invalid currency/currency-amount conversion ")
		Case "687"
		fmtresponseAmexTw = ("Decline: AMEX CAS-invalid/illogical/unsupported trans request ")
		Case "688"
		fmtresponseAmexTw = ("Decline: Transaction referred to authorizer ")
		Case "689"
		fmtresponseAmexTw = ("Decline: Transaction denied by Host ")
		Case "690"
		fmtresponseAmexTw = ("Decline: Transaction not processed -no authorizer available ")
		Case "691"
		fmtresponseAmexTw = ("Decline: Invalid alpha code in service establishment field ")
		Case "692"
		fmtresponseAmexTw = ("Decline: Invalid travel agent# in SE field ")
		Case "693"
		fmtresponseAmexTw = ("Decline: Invalid or out-of-range referral queue number ")
		Case "694"
		fmtresponseAmexTw = ("Decline: Transaction denied - cardmember left se ")
		Case "695"
		fmtresponseAmexTw = ("Decline: Optima reversal trans (14) processed by Host ")
		Case "696"
		fmtresponseAmexTw = ("Decline: Optima reversal trans (14) rejected by Host ")
		Case "697"
		fmtresponseAmexTw = ("Decline: CID field invalid - please reenter CID data ")
		Case "698"
		fmtresponseAmexTw = ("Decline: CID data req, but is missing - enter CID data ")
		Case "701"
		fmtresponseAmexTw = ("Decline: Unable to process transaction ")
		Case "702"
		fmtresponseAmexTw = ("Decline: Transaction source is invalid or not present ")
		Case "703"
		fmtresponseAmexTw = ("Decline: Function code is invalid or not present ")
		Case "704"
		fmtresponseAmexTw = ("Decline: Cardnumber is invalid or not present ")
		Case "705"
		fmtresponseAmexTw = ("Decline: Cardmember account number not enrolled ")
		Case "706"
		fmtresponseAmexTw = ("Decline: Cardmember account number already enrolled ")
		Case "707"
		fmtresponseAmexTw = ("Decline: Cardmember name not present, entry required ")
		Case "708"
		fmtresponseAmexTw = ("Decline: Zip code invalid or not present ")
		Case "709"
		fmtresponseAmexTw = ("Decline: City or state invalid or not present ")
		Case "710"
		fmtresponseAmexTw = ("Decline: Product code invalid or not present ")
		Case "711"
		fmtresponseAmexTw = ("Decline: Billing currency invalid or not present ")
		Case "712"
		fmtresponseAmexTw = ("Decline: Enrollment source code invalid or not present ")
		Case "713"
		fmtresponseAmexTw = ("Decline: Operation id for PIN update invalid or not present ")
		Case "714"
		fmtresponseAmexTw = ("Decline: PIN invalid or not present ")
		Case "715"
		fmtresponseAmexTw = ("Decline: AMEX inhibit code invalid ")
		Case "716"
		fmtresponseAmexTw = ("Decline: AMEX inhibit on operation id invalid/not present ")
		Case "717"
		fmtresponseAmexTw = ("Decline: AMEX inhibit off operation id invalid/not present ")
		Case "718"
		fmtresponseAmexTw = ("Decline: AMEX inhibit code not present ")
		Case "719"
		fmtresponseAmexTw = ("Decline: 30 day delinquent code invalid or not present ")
		Case "720"
		fmtresponseAmexTw = ("Decline: 60 day delinquent code invalid or not present ")
		Case "721"
		fmtresponseAmexTw = ("Decline: 90 day delinquent code invalid or not present ")
		Case "722"
		fmtresponseAmexTw = ("Decline: Denial override code invalid or not present ")
		Case "723"
		fmtresponseAmexTw = ("Decline: Velocity rate code invalid or not present ")
		Case "724"
		fmtresponseAmexTw = ("Decline: Update transaction contains no fields for update ")
		Case "725"
		fmtresponseAmexTw = ("Decline: Accnt not enrolled,negative recd found for card no. ")
		Case "726"
		fmtresponseAmexTw = ("Decline: Account is not eligible for enrollment ")
		Case "727"
		fmtresponseAmexTw = ("Decline: Update of inhibit code with lesser priorty invalid ")
		Case "728"
		fmtresponseAmexTw = ("Decline: Cardmember country code invalid or not present ")
		Case "729"
		fmtresponseAmexTw = ("Decline: Userid function invalid or not present ")
		Case "730"
		fmtresponseAmexTw = ("Decline: Cancel code invalid or not present ")
		Case "731"
		fmtresponseAmexTw = ("Decline: Account not enrolled because account cancelled ")
		Case "732"
		fmtresponseAmexTw = ("Decline: Amex inhibit cannot be turned off by this tranxn ")
		Case "733"
		fmtresponseAmexTw = ("Decline: Card Nos. for this product type cannot be enrolled ")
		Case "734"
		fmtresponseAmexTw = ("Decline: Account not enrolled, cardmember is delinquent ")
		Case "735"
		fmtresponseAmexTw = ("Decline: IMS is not available for enrollment process ")
		Case "736"
		fmtresponseAmexTw = ("Decline: Enroll inhibit code invalid or not present ")
		Case "737"
		fmtresponseAmexTw = ("Decline: TPF system can not get CAS record ")
		Case "738"
		fmtresponseAmexTw = ("Decline: Received response from vendor ")
		Case "739"
		fmtresponseAmexTw = ("Decline: Amex CAS reply - approve print no vat receipt ")
		Case "801"
		fmtresponseAmexTw = ("Decline: No header record on CMM input ")
		Case "802"
		fmtresponseAmexTw = ("Decline: CMM batch sequence number not consecutive ")
		Case "803"
		fmtresponseAmexTw = ("Decline: CMM control file is empty ")
		Case "804"
		fmtresponseAmexTw = ("Decline: CMM input record type is invalid ")
		Case "805"
		fmtresponseAmexTw = ("Decline: No trailer record on CMM input ")
		Case "806"
		fmtresponseAmexTw = ("Decline: CMM trailer counts do not balance ")
		Case "820"
		fmtresponseAmexTw = ("Decline: Invalid error code 16820 ")
		Case "821"
		fmtresponseAmexTw = ("Decline: Invalid error code 16821 ")
		Case "822"
		fmtresponseAmexTw = ("Decline: Host connection established ")
		Case "823"
		fmtresponseAmexTw = ("Decline: Host connection broken ")
		Case "850"
		fmtresponseAmexTw = ("Decline: &a1& paging partition is &a2&% full ")
		Case "851"
		fmtresponseAmexTw = ("Decline: &a1& file partition is &a2&% full ")
		Case "852"
		fmtresponseAmexTw = ("Decline: &a1& had fatal error, total fatal errors: &a2& ")
		Case "853"
		fmtresponseAmexTw = ("Decline: &a1& has aborted on module &a2& ")
		Case "854"
		fmtresponseAmexTw = ("Decline: Restart of &a1& of module &a2& was successful ")
		Case "855"
		fmtresponseAmexTw = ("Decline: &a1& is not available to module &a2& ")
		Case "856"
		fmtresponseAmexTw = ("Decline: &a1& cpu utilization is &a2&% ")
		Case "857"
		fmtresponseAmexTw = ("Decline: &a1& is out of service, attempting to restart ")
		Case "858"
		fmtresponseAmexTw = ("Decline: &a1& has been placed back in service ")
		Case "859"
		fmtresponseAmexTw = ("Decline: Monitoring has been enabled on &a1& on module &a2& ")
		Case "860"
		fmtresponseAmexTw = ("Decline: Monitoring has been disabled on &a1& on module &a2 ")
		Case "861"
		fmtresponseAmexTw = ("Decline: &a1& has logged into module &a2& ")
		Case "862"
		fmtresponseAmexTw = ("Decline: &a1& has logged out of module &a2& ")
		Case "863"
		fmtresponseAmexTw = ("Decline: The reconfigure command issued for module &a1& ")
		Case "864"
		fmtresponseAmexTw = ("Decline: X25_exchange extn &a1& not running on module &a2& ")
		Case "865"
		fmtresponseAmexTw = ("Decline: X25_exchange extn &a1& restarted on module &a2& ")
		Case "866"
		fmtresponseAmexTw = ("Decline: &a1& exceeded threshold limit,removed from service ")
		Case "867"
		fmtresponseAmexTw = ("Decline: Do you really wish to disable &a1&? ")
		Case "868"
		fmtresponseAmexTw = ("Decline: Do you really wish to enable &a1&? ")
		Case "869"
		fmtresponseAmexTw = ("Decline: Remote Negative Traffic has stopped de-queing ")
		Case "870"
		fmtresponseAmexTw = ("Decline: Remote Negative Traffic has started de-queing ")
		Case "871"
		fmtresponseAmexTw = ("Decline: Localized System Support is not logged in ")
		Case "872"
		fmtresponseAmexTw = ("Decline: Do you really, really want to delete &a1&? ")
		Case "873"
		fmtresponseAmexTw = ("Decline: A red light is illuminated on module &a1& ")
		Case "874"
		fmtresponseAmexTw = ("Decline: &a1& on module &a2& ")
		Case "875"
		fmtresponseAmexTw = ("Decline: &a1& is disabled on module &a1& ")
		Case "876"
		fmtresponseAmexTw = ("Decline: &a1& rebooted on module &a2& ")
		Case "877"
		fmtresponseAmexTw = ("Decline: &a1& has no specified restart path,restart aborted ")
		Case "878"
		fmtresponseAmexTw = ("Decline: &a1& was rebooted on module &&& ")
		Case "879"
		fmtresponseAmexTw = ("Decline: Reboot of &a1& failed on module &a2&,code was &a3& ")
		Case "880"
		fmtresponseAmexTw = ("Decline: &a1& is not registered to module &a2& ")
		Case "881"
		fmtresponseAmexTw = ("Decline: Update of &a1& failed, not place in service ")
		Case "882"
		fmtresponseAmexTw = ("Decline: &a1& is simplexed on module &a2& ")
		Case "883"
		fmtresponseAmexTw = ("Decline: LSS ""HUB"" queue &a1& is full ")
		Case "884"
		fmtresponseAmexTw = ("Decline: LSS queue &a1& not available ")
		Case "885"
		fmtresponseAmexTw = ("Decline: LSS_SVR has been normally shutdown on &a1& ")
		Case "886"
		fmtresponseAmexTw = ("Decline: LSS_SVR has aborted on &a1&, restarting now ")
		Case "901"
		fmtresponseAmexTw = ("Decline: normal process shutdown requested ")
		Case "950"
		fmtresponseAmexTw = ("Decline: input gt_tbl empty file ")
		Case "951"
		fmtresponseAmexTw = ("Decline: gt_tbl file less than 1002 recs ")
		Case "952"
		fmtresponseAmexTw = ("Decline: gt_tbl not translated from MVS ")
		Case "953"
		fmtresponseAmexTw = ("Decline: gt_tbl not sequential ")
		Case "954"
		fmtresponseAmexTw = ("Decline: mod10 check acct invalid ")
		Case "955"
		fmtresponseAmexTw = ("Decline: gt acct range invalid ")
		Case "956"
		fmtresponseAmexTw = ("Decline: GT control file is empty ")
		Case "957"
		fmtresponseAmexTw = ("Decline: gt_tbl file greater than 1002 recs ")
		Case "958"
		fmtresponseAmexTw = ("Decline: amex cardnumber prefix not valid ")
		Case "959"
		fmtresponseAmexTw = ("Decline: Invalid error code 16959 ")
		Case "960"
		fmtresponseAmexTw = ("Decline: Unrecognized message id ")
		Case "961"
		fmtresponseAmexTw = ("Decline: Invalid settlement date ")
		Case "962"
		fmtresponseAmexTw = ("Decline: Invalid NII ")
		Case "963"
		fmtresponseAmexTw = ("Decline: Invalid POS entry mode ")
		Case "964"
		fmtresponseAmexTw = ("Decline: Invalid condition code ")
		Case "965"
		fmtresponseAmexTw = ("Decline: Invalid retrieval reference number ")
		Case "966"
		fmtresponseAmexTw = ("Decline: Virtual memory connection has not been established ")
		Case "967"
		fmtresponseAmexTw = ("Decline: Front End shuntback ")
		Case "968"
		fmtresponseAmexTw = ("Decline: TAC Control File Empty ")
		Case "969"
		fmtresponseAmexTw = ("Decline: TAC Table Empty ")
		Case "970"
		fmtresponseAmexTw = ("Decline: TAC Table Format Error ")
		Case "971"
		fmtresponseAmexTw = ("Decline: Refer to card issuer- special conditions ")
		Case "972"
		fmtresponseAmexTw = ("Decline: Invalid error code 16972 ")
		Case "973"
		fmtresponseAmexTw = ("Decline: Pick-up card special condition ")
		Case "974"
		fmtresponseAmexTw = ("Decline: Request in progress ")
		Case "975"
		fmtresponseAmexTw = ("Decline: Approve for partial amount ")
		Case "001"
		fmtresponseAmexTw = ("Decline: Invalid Card Expiry Date ")
		Case "002"
		fmtresponseAmexTw = ("Decline: Invalid Card Member Number ")
		Case "003"
		fmtresponseAmexTw = ("Decline: Invalid Merchant Number ")
		Case "004"
		fmtresponseAmexTw = ("Decline: Invalid Transaction Date ")
		Case "006"
		fmtresponseAmexTw = ("Decline: Invalid Industry Code ")
		Case "007"
		fmtresponseAmexTw = ("Decline: Invalid SubIndustry Code ")
		Case "008"
		fmtresponseAmexTw = ("Decline: Invalid Transaction Type ")
		Case "998"
		fmtresponseAmexTw = ("Decline: No Authorisation Response Received ")
		Case "999"
		fmtresponseAmexTw = ("Decline: Incorrect CAN/SE Number ")
		Case else
		fmtresponseAmexTw = ("Approved")
	End Select
end function
 %>
