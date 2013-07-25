<%@ CodePage=65001 %>
<!-- #include file="init.asp" -->
<%
'Dim SessionFarm
'set SessionFarm = Server.CreateObject("SessionFarm.SFSession")
response.charset="utf-8"
%>
<%
if not Session("Pass") then
	response.redirect "su1.asp"
else
%>
		<!-- #include file="inc_i18n.asp" -->
		<!-- #include file="inc_localization.asp" -->
		<!-- #include file="adm/inc_chk_ss.asp" -->
		<!-- #include file="adm/inc_acct_balance.asp" -->		
		<!-- #include file="adm/inc_crypt.asp" -->
        <% if session("CR_Memberships") <> 0 then %>
            <!-- #include file="inc_dbconn_regions.asp" -->
            <!-- #include file="inc_dbconn_wsMaster.asp" -->
            <!-- #include file="adm/inc_masterclients_util.asp" -->
        <% end if %>
		<!-- #include file="adm/inc_chk_membership.asp" -->
<%
	if session("mvarMIDs")=2 then
		ccMode = true
	else
		ccMode = false
	end if
	if ccMode = false then
		response.redirect "cc_unabled.asp"
	end if
%>
		<!-- #include file="inc_loading.asp" -->
<%

dim noProd, pClassID, classDate, typeGroup, typeGroupID, prodName
dim prodSelected, pAddAmount, pProdCost, DepositProdID, DepositAmount, isEditable, CltBalance, CltBalance2, balChoice, balanceApplied, partialBalance, pAllowCltNegBal, unPaidCase, unPaidCaseRec
dim pUseEnrollResReq, pUseEnroll, pUsePmtPlanOnly, pBuyCredit, isMember, mem_AllowNonMemberPurchases, cMembershipID, mem_ServiceDiscountPerc, mem_ProdDiscountPerc
dim payMode	''1-Reg 2-Buy Credit 3-EnrollBasic 4-EnrollPmtPlan 5-EnrollReqRes 6-Gift Card

if isNum(request.form("tabID"))then
	session("tabID") = request.form("tabID")
elseif isNum(request.QueryString("tabID")) then
	session("tabID") = request.querystring("tabID")
end if

dim phraseDictionary
set phraseDictionary = LoadPhrases("Mainshop2Page", 15)

''''Set initial Params
prodSelected = false
pAllowCltNegBal = checkStudioSetting("tblGenOpts","AllowCltNegBal")
noProd=""
typeGroupID = Request.QueryString("typeGroup")
if Request.Querystring("prod")<>"" then
	productID = CDBL(Request.Querystring("prod"))
else
	productID = 0
end if

if request.querystring("applyBal")="" and request.form("frmApplyBalance")="" then
	balChoice = 0
else ' applyBal is 'yes' or 'no' OR frmApplyBalance is 'yes' or 'no'  - that is, they've chosen whether or not to apply their balance.
	balChoice = 1
end if

if session("pClientID")<>"" then
	clientID = session("pClientID")
else
	clientID = session("mvarUserID")
end if

set rsWSDB = Server.CreateObject("ADODB.Recordset")

mem_ServiceDiscountPerc = 0
mem_ProdDiscountPerc = 0

'BJD: 6/12/08 - new membership logic
if clientID<>"" then
	isMember = checkMembership(clientID, "")
	
	if isMember then
		cMembershipID = MemSeriesTypeID
	
		strSQL = "SELECT ServiceDiscountPerc, ProdDiscountPerc, AllowNonMemberPurchases FROM tblSeriesType WHERE SeriesTypeID = " & sqlClean(cMembershipID)
		rsWSDB.CursorLocation = 3
		rsWSDB.open strSQL, cnWS
		Set rsWSDB.ActiveConnection = Nothing
		
		if NOT rsWSDB.EOF then
			mem_ServiceDiscountPerc = rsWSDB("ServiceDiscountPerc")
			mem_ProdDiscountPerc = rsWSDB("ProdDiscountPerc")
			mem_AllowNonMemberPurchases = rsWSDB("AllowNonMemberPurchases")
		end if
		rsWSDB.close
	else
		cMembershipID = -1
	end if
	
else
	isMember = false
	cMembershipID = -1
end if


if request.form("requiredopt" & session("ResourceDisplayName") & "Type")<>"" then
	pResourceType = CLNG(request.form("requiredopt" & session("ResourceDisplayName") & "Type"))
end if

if request.form("requiredoptDeposit")<>"" then
	DepositProdID = CLNG(request.form("requiredoptDeposit"))
else
	DepositProdID = 0
end if

if request.querystring("clsDate")<>"" then
	classDate = request.querystring("clsDate")
elseif request.form("clsDate")<>"" then
	classDate = request.form("clsDate")
end if

	''''Determine PayMode - Check for Enrollment
	pEnroll = false
	pEnrollResReq = false
	pUsePmtPlanOnly = false
	payMode = 0

	'' Get ClassID, if no class then assume payMode = 1 - Reg
	if request.form("cid")<>"" then
		pClassID = request.form("cid")
	elseif request.querystring("cid")<>"" then
		pClassID = request.querystring("cid")
	else
		payMode = 1
	end if
	'' Credit Mode
	if request.form("frmBuyCredit")="true" then
		payMode = 2
	end if

	if typeGroupID ="GC" then
		payMode = 6
	end if

	'' If No PayMode Yet, Check Enroll Type PayModes

	if payMode=0 then
		 strSQL = "SELECT wsEnrollment FROM tblTypeGroup WHERE TypeGroupID=" & sqlClean(typeGroupID)
		 rsWSDB.CursorLocation = 3
		rsWSDB.open strSQL, cnWS
		Set rsWSDB.ActiveConnection = Nothing

		 if not rsWSDB.EOF then
			if rsWSDB("wsEnrollment") then
				payMode = 3
				if checkStudioSetting("tblResvOpts","EnrollReqResource") then
					payMode = 5
				else
	 			    rsWSDB.close
					strSQL = "SELECT ISNULL(tblCourses.PmtPlan, tblClasses.PmtPlan) as PmtPlan FROM tblClasses LEFT OUTER JOIN tblCourses ON tblCourses.CourseID = tblClasses.CourseID WHERE ClassID=" & sqlClean(pClassID)
					rsWSDB.CursorLocation = 3
					rsWSDB.open strSQL, cnWS
					Set rsWSDB.ActiveConnection = Nothing

					if rsWSDB("PmtPlan") then
						payMode = 4
					end if
				end if
			end if
		end if
		rsWSDB.close		 
	end if

	if payMode = 4 or payMode = 5 then	''If EnrollPmtPlan or EnrollReqRes
		''If Deposit has been selected
		if request.form("requiredoptDeposit")<>"0" AND request.form("requiredoptDeposit")<>"-1" AND request.form("requiredoptDeposit")<>"" then
			strSQL = "Select OnlinePrice, Editable From Products WHERE ProductID=" & sqlClean(request.form("requiredoptDeposit")) 
			rsWSDB.CursorLocation = 3
			rsWSDB.open strSQL, cnWS
			Set rsWSDB.ActiveConnection = Nothing

			if not rsWSDB.EOF then
				if rsWSDB("Editable") then
					DepositAmount = request.form("requiredtxtPaymentAmount")
				else
					DepositAmount = rsWSDB("OnlinePrice")
				end if
			end if
			rsWSDB.close
		end if
		frmActionStr = "purchase_res_p.asp"
	else
		frmActionStr = "purchase_p.asp"
	end if	

	if payMode = 4 or payMode = 5 then
		strSQL = "SELECT ClassName FROM tblClassDescriptions, tblClasses WHERE tblClasses.DescriptionID=tblClassDescriptions.ClassDescriptionID AND tblClasses.ClassID=" & sqlClean(pclassID)
		rsWSDB.CursorLocation = 3
		rsWSDB.open strSQL, cnWS
		Set rsWSDB.ActiveConnection = Nothing

		className = rsWSDB("ClassName")
		rsWSDB.close
	end if
	
	''''''Unpaid Check'''''
	if request.form("frmUnpaid")<>"" then
		unPaidCase = request.form("frmUnpaid")
	else
		unPaidCase = "unknown"
	end if
	
	if payMode=0 then
		payMode = 1
	end if
	if unPaidCase="unknown" AND TypeGroupID<>"" AND payMode=1 then
		''''Check for unpaids for this client

	    strSQL = "SELECT PmtRefNo, Remaining FROM [PAYMENT DATA] WHERE [PAYMENT DATA].TypeGroup=" & sqlClean(TypeGroupID) & " AND ClientID=" & sqlClean(clientID) & " AND type=9 AND ExpDate > " & DateSep & sqlClean(DateValue(CDATE(DateAdd("n", Session("tzOffset"),Now)))) & DateSep & " AND [Current Series]=1"
		rsWSDB.CursorLocation = 3
		rsWSDB.open strSQL, cnWS
		Set rsWSDB.ActiveConnection = Nothing

		if rsWSDB.EOF then		
			unPaidCase = "false"
		else
			dim absValRemaining, oldPmtRefNo, ElapsedTime
			absValRemaining = abs(rsWSDB("Remaining"))
			oldPmtRefNo = rsWSDB("PmtRefNo")
			unPaidCase = "true"
			rsWSDB.close
				strSQL = "SELECT [PAYMENT DATA].PmtRefNo, [PAYMENT DATA].Remaining, [PAYMENT DATA].Type, Max([VISIT DATA].ClassDate) AS LastVisit, "
				strSQL = strSQL & "Min([VISIT DATA].ClassDate) AS FirstVisit, [PAYMENT DATA].TypeGroup "
				strSQL = strSQL & "FROM [PAYMENT DATA] INNER JOIN [VISIT DATA] ON [PAYMENT DATA].PmtRefNo = [VISIT DATA].PmtRefNo "
				strSQL = strSQL & "GROUP BY [PAYMENT DATA].PmtRefNo, [PAYMENT DATA].Remaining, [PAYMENT DATA].Type, [PAYMENT DATA].TypeGroup "
				strSQL = strSQL & "HAVING ((([PAYMENT DATA].PmtRefNo)=" & sqlClean(oldPmtRefNo) & ") AND (([PAYMENT DATA].Type)=9));"
				rsWSDB.CursorLocation = 3
				rsWSDB.open strSQL, cnWS
				Set rsWSDB.ActiveConnection = Nothing
				if not rsWSDB.EOF then
					ElapsedTime = DateDiff("d", CDate(rsWSDB("FirstVisit")), CDate(rsWSDB("LastVisit")))
					upSDate = CDATE(rsWSDB("FirstVisit"))
					if ElapsedTime = 0 then
						upEDate = upSDate
					else
						upEDate = CDATE(rsWSDB("LastVisit"))
					end if
				else
					unPaidCase = "false"
				end if
		end if
		rsWSDB.close
	end if
	if unPaidCase="true" then
		unPaidCaseRec = "true"
	else
		unPaidCaseRec = "unknown"
	end if

Dim nTax1, nTax2, nTax3, nTax4, nTax5, TTax
strSQL = "SELECT CAST(Tax1 AS DECIMAL(20,8)) AS Tax1, CAST(Tax2 AS DECIMAL(20,8)) AS Tax2, CAST(Tax3 AS DECIMAL(20,8)) AS Tax3, CAST(Tax4 AS DECIMAL(20,8)) AS Tax4, CAST(Tax5 AS DECIMAL(20,8)) AS Tax5 FROM Location WHERE LocationID=98"
rsWSDB.CursorLocation = 3
rsWSDB.open strSQL, cnWS
Set rsWSDB.ActiveConnection = Nothing

nTax1 = rsWSDB("Tax1")
nTax2 = rsWSDB("Tax2")
nTax3 = rsWSDB("Tax3")
nTax4 = rsWSDB("Tax4")
nTax5 = rsWSDB("Tax5")
rsWSDB.close

strSQL = "SELECT tblGenOpts.PurchPolicy FROM tblGenOpts WHERE tblGenOpts.StudioID=" & sqlClean(session("StudioID"))
rsWSDB.CursorLocation = 3
rsWSDB.open strSQL, cnWS
Set rsWSDB.ActiveConnection = Nothing

	strPurchPolicy = rsWSDB("PurchPolicy")
rsWSDB.close
%>
		
<html>
<head>
<title><%=Session("StudioName")%> Online</title>
<meta http-equiv="Content-Type" content="text/html">
<!-- #include file="frame_bottom.asp" -->
<%= js(array("mb", "formval", "VCC1", "VCC2", "valcur", "main_shop_old")) %>

<script type="text/javascript">
function popImageWindow(imageFileName) {
	myHeight = 175;
	myWidth = 250;
	var height = screen.height;
	var width = screen.width;
	var leftpos = width / 2 - myWidth / 2;
	var toppos = height / 2 - myHeight / 2; 
    recWindow=window.open('<% response.write "http" & addS & "://" & request.servervariables("SERVER_NAME") & "/images/" %>' + imageFileName,"ImageWindow","toolbar=no,location=no,directories=no,resizable=yes,menubar=yes,scrollbars=yes,status=no,width=" + myWidth + ",height=" + myHeight + ", left=" + leftpos + ",top=" + toppos);
    setTimeout('recWindow.focus()',1);
}
</script>


	
	<!-- #include file="inc_date_ctrl.asp" -->


</head>

<body>
<% pageStart %>
<form name="frmCC" onSubmit="return processSubmit(this);" method="POST" action="<%=frmActionStr%>?typeGroupID=<%=typeGroupID%>&classID=<%=pClassID%>">
<!-- #include file="inc_chk_holiday.asp" -->
<!-- #include file="adm/inc_frm_rtn_pmt.asp"  -->
	<input type="hidden" name="frmUsePmtPlan" value="<%=request.form("frmUsePmtPlan")%>">
	<input type="hidden" name="frmBuyCredit" value="<%=request.form("frmBuyCredit")%>">
    <input type="hidden" name="frmUnpaid" value="<%=unPaidCase%>">
    <input type="hidden" name="frmUnpaidRec" value="<%=unPaidCaseRec%>">
	<input type="hidden" name="frmApplyBalance" value="<%=request.querystring("applyBal")%>">

<table height="100%" width="<%=strPageWidth%>" cellspacing="0">
  <tr> 
      <td class="center-ch" valign="top" height="100%" width="100%"> <br />
          <table cellspacing="0" width="90%" height="100%">
            <tr> 
<%
	giftCardsAvailable = false
							
	strSQL = "SELECT PRODUCTS.ProductID, PRODUCTS.Description "
	strSQL = strSQL & "FROM PRODUCTS WHERE [Delete]=0 AND wsShow = 1 AND Products.CategoryID = 22"

	set rsProduct2 = Server.CreateObject("ADODB.Recordset")
	rsProduct2.CursorLocation = 3
	rsProduct2.open strSQL, cnWS
	Set rsProduct2.ActiveConnection = Nothing

	if not rsProduct2.EOF then
		giftCardsAvailable = true
	end if 
	giftCardsAvailable = false


	rsProduct2.close

	if payMode=2 then
		displayPurchaseType = "Buy Credit"
		displayOtherLink = "Buy Gift Cards"
	elseif payMode=6 then
		displayGiftCardText = "Buy Gift Cards"
		displayPurchaseType = "Buy Services"
	else
		displayPurchaseType = "Buy Services"
		displayGiftCardText = "Buy Gift Cards"
	end if
%>

	<% if giftCardsAvailable then %>
              <td class="headText" align="left">
			  		<table  cellpadding="3" cellspacing="0" class="border4 center-ch">
						<tr>
							<td class="headText" <% if payMode = 6 then %>onClick="reloadPage4('')" style="cursor:pointer"<% else %>style="background-color:<%=session("pageColor4")%>;" <% end if %> width="50%" class="center-ch">
							<% if payMode=6 then %>
								&nbsp;<b><%=displayPurchaseType%></b>&nbsp;
							<% else %>
								&nbsp;<span style="color:white;"><b><%=displayPurchaseType%></b></span>&nbsp;
							<% end if %>
							</td>
					
							<td class="headText" <% if payMode <> 6 then %>onClick="reloadPage4('GC')" style="cursor:pointer"<% else %>style="background-color:<%=session("pageColor4")%>;" <% end if %> class="center-ch" width="50%">
							<% if payMode<>6 then %>
								&nbsp;<b><%=displayGiftCardText%></b>&nbsp;
							<% else %>
								&nbsp;<span style="color:white;"><b><%=displayGiftCardText%></b></span>&nbsp;
							<% end if %>
							</td>
						</tr>
					</table>			  		
			  </td>
            </tr>
	<% else %>
              <td class="headText" align="left">
			  	<b><%=displayPurchaseType%></b>  		
			  </td>
            </tr>
	<% end if %>
	
	
	<% if payMode<3 then %>
            <tr>
              <td class="mainText right"><b> </b></td>
            </tr>
	<% end if %>

            <tr> 
              <td valign="top" class="mainTextBig center-ch"> 
                <table class="mainText" width="95%" cellspacing="0">
                  <tr > 
                    <td class="mainTextBig" colspan="2" valign="top"> 
                      <table class="mainText" width="100%" cellspacing="0">
					<% if NOT isNull(strPurchPolicy) then %>
					<% if strPurchPolicy<>"" then %>
                        <tr> 
                          <td colspan="2" class="mainText" align="left"><div id="purchPol"><br />&nbsp;<%=strPurchPolicy%></div></td>
						</tr>						
					<% end if %>
					<% end if %>
                        <% if request.form("requiredoptPaymentMethod")="16" then %>
                        <tr> 
                          <td COLSPAN="2" height="1" width="100%" align="left" style="background-color:<%=session("pageColor4")%>;"><img src="<%= contentUrl("/asp/images/trans.gif") %>" width="100%" height="1"></td>
                        </tr>
                        <tr> 
                          <td colspan="2" class="mainText" align="left"><b>&nbsp; 
                            <%						  
							CltBalance = getAccountBalance(session("pClientID"),"", "")
							  Response.Write DisplayPhrase(phraseDictionary,"Accountbalancetext") & ":&nbsp;" 
							  Response.write FmtCurrency(CltBalance) 
%>
                            &nbsp;</b></td>
                        </tr>
                        <tr> 
                          <td COLSPAN="2" height="1" width="100%" align="left" style="background-color:<%=session("pageColor4")%>;"><img src="<%= contentUrl("/asp/images/trans.gif") %>" width="100%" height="1"></td>
                        </tr>
                        <% end if %>
                        <% if payMode=5 then  %>
                        <tr> 
                          <td COLSPAN="2" height="6" width="100%" align="left"></td>
                        </tr>
                        <% if payMode<>4 then %>
                        <tr> 
                          <td COLSPAN="2" height="1" width="100%" align="left" style="background-color:<%=session("pageColor4")%>;"><img src="<%= contentUrl("/asp/images/trans.gif") %>" width="100%" height="1"></td>
                        </tr>
                        <tr> 
                          <td class="mainText2" width="30%" nowrap style="background-color:<%=session("pageColor2")%>;"><b>&nbsp;
						  	<%=DisplayPhrase(phraseDictionary,"Buildyour")%> <%=xssStr(allHotWords(3))%></b></td>
                          <td align="left" style="background-color:<%=session("pageColor2")%>;">&nbsp;</td>
                        </tr>
                        <tr> 
                          <td COLSPAN="2" height="1" width="100%" align="left" style="background-color:<%=session("pageColor4")%>;"><img src="<%= contentUrl("/asp/images/trans.gif") %>" width="100%" height="1"></td>
                        </tr>
                        <% end if ''usePmtPlan %>
                        <% end if %>
                        <tr> 
                          <td COLSPAN="2" height="6" width="100%" align="left"></td>
                        </tr>
                        <% if payMode=1 then %>
                        <tr> 
                          <td width="1%" height="18"> 
                            <%
					set rsProduct2 = Server.CreateObject("ADODB.Recordset")
					if checkStudioSetting("tblResvOpts","EnrollReqResource") then
					''If EnrollReqRes then don't show products of Enroll TG
							strSQL = "SELECT TypeGroup, TypeGroupID FROM tblTypeGroup WHERE wsEnrollment=0 AND (wsAppointment=1 OR wsReservation=1) AND Active=1"
					else
							'strSQL = "SELECT TypeGroup, TypeGroupID FROM tblTypeGroup WHERE (wsEnrollment=1 OR wsAppointment=1 OR wsReservation=1) AND Active=1"
							strSQL = "SELECT DISTINCT tblTypeGroup.TypeGroupID, tblTypeGroup.TypeGroup, PRODUCTS.TypeGroup AS PRODTG FROM tblTypeGroup LEFT OUTER JOIN PRODUCTS ON tblTypeGroup.TypeGroupID = PRODUCTS.TypeGroup WHERE (tblTypeGroup.wsEnrollment = 1) AND (tblTypeGroup.Active = 1) AND (NOT (PRODUCTS.TypeGroup IS NULL)) AND (PRODUCTS.Discontinued = 0) AND (PRODUCTS.[Delete] = 0) AND (PRODUCTS.Class = 1) AND (PRODUCTS.Type <> 9) AND (PRODUCTS.wsShow = 1) OR (tblTypeGroup.Active = 1) AND (NOT (PRODUCTS.TypeGroup IS NULL)) AND (PRODUCTS.Discontinued = 0) AND (PRODUCTS.[Delete] = 0) AND "
							strSQL = strSQL & "(PRODUCTS.Class = 1) AND (PRODUCTS.Type <> 9) AND (tblTypeGroup.wsAppointment = 1) AND (PRODUCTS.wsShow = 1) OR (tblTypeGroup.Active = 1) AND (NOT (PRODUCTS.TypeGroup IS NULL)) AND (PRODUCTS.Discontinued = 0) AND (PRODUCTS.[Delete] = 0) AND (PRODUCTS.Class = 1) AND (PRODUCTS.Type <> 9) AND (tblTypeGroup.wsReservation = 1) AND (PRODUCTS.wsShow = 1) OR (tblTypeGroup.Active = 1) AND (NOT (PRODUCTS.TypeGroup IS NULL)) AND (PRODUCTS.Discontinued = 0) AND (PRODUCTS.[Delete] = 0) AND (PRODUCTS.Class = 1) AND (PRODUCTS.Type <> 9) AND (tblTypeGroup.wsResource = 1) AND (PRODUCTS.wsShow = 1) ORDER BY tblTypeGroup.TypeGroup"
					end if
					rsProduct2.CursorLocation = 3
					rsProduct2.open strSQL, cnWS
					Set rsProduct2.ActiveConnection = Nothing

%>
                            <b>&nbsp;<%=DisplayPhrase(phraseDictionary,"Whatservice")%>&nbsp;&nbsp;</b> 
                          </td>
                          <td height="18" align="left"> 
                            <select name="purchaseType" onChange="reloadPageTG();">
                              <option value="0"><%=DisplayPhrase(phraseDictionary,"Selecttype")%></option>
                              <% if not rsProduct2.EOF then
						
							Do While not rsProduct2.EOF
						%>
                              <option value="<%=rsProduct2("TypeGroupID")%>" 
						<%
								if typeGroupID=Trim(rsProduct2("TypeGroupID")) then
									Response.Write " selected"
								end if 
						%>
								><%=rsProduct2("TypeGroup")%></option>
                              <%		  
								rsProduct2.MoveNext
							
								Loop
							end if
						
							rsProduct2.close
							'Set rsProduct2 = Nothing
							
							'strSQL = "SELECT PRODUCTS.ProductID, PRODUCTS.Description "
							'strSQL = strSQL & "FROM PRODUCTS WHERE [Delete]=0 AND wsShow = 1 AND Products.CategoryID = 22"

							'rsProduct2.CursorLocation = 3
							'rsProduct2.open strSQL, cnWS
							'Set rsProduct2.ActiveConnection = Nothing

							'if not rsProduct2.EOF then
						%>
								<!--option value="GC">Gift Card</option-->
						<% 	'end if 

							'rsProduct2.close
							Set rsProduct2 = Nothing

						%>
                            </select>
<script type="text/javascript"> 
      if (document.frmCC.purchaseType.options.length == 2 && document.frmCC.purchaseType.options[0].selected ) { 
           document.frmCC.purchaseType.options[1].selected = true;
		   reloadPageTG();
      } 
</script>
                          </td>
                        </tr>
                        <% else %>
                        <input type="hidden" name="purchaseType" value="<%=TypeGroupID%>">
						<% 		if payMode = 6 then  ' Gift Card %>
									<input type="hidden" name="giftCardPurchase" value="1">
									<input type="hidden" name="giftCardAmount" value="<%=FIXME%>">
						<% 		end if  ' payMode = 6 %>
                        <% end if '''''''pUseEnroll %>
						

							<% if payMode=1 AND unPaidCase="true" AND unPaidCaseRec="unknown" then %>
							<% else %>

								<% if unPaidCaseRec="true" then %>
									<tr valign="middle" height="22"> 
									  
                          <td colspan="2"> 
                            <% if request.form("frmUPRem")="" then %>
                            <b>&nbsp;<span style="color:#FF0000;"><%=absValRemaining%></span> 
                            <%=DisplayPhrase(phraseDictionary,"Unpaidsfrom")%> <%=FmtDateShort(upSDate)%> 
                            <% if ElapsedTime<>"0" then response.write "&nbsp;to&nbsp;" & FmtDateShort(upEDate) & "&nbsp;spanning <span style=""color:#FF0000;"">" & ElapsedTime & "</span>&nbsp;days" end if%>
                            </b> 
                            <% else %>
                            <b>&nbsp;<%=DisplayPhrase(phraseDictionary,"Reconciling")%> <span style="color:#FF0000;"><%=request.form("frmUPRem")%></span> 
                            <%=DisplayPhrase(phraseDictionary,"Unpaidsfrom")%> <%=FmtDateShort(request.form("frmUPSDate"))%> 
                            <% if request.form("frmUPElapsedTime")<>"0" then response.write "&nbsp;to&nbsp;" & FmtDateShort(request.form("frmUPEDate")) & "&nbsp;spanning <span style=""color:#FF0000;"">" & request.form("frmUPElapsedTime") & "</span>&nbsp;days" end if%>
                            </b> 
                            <% end if %>
                          </td>
									</tr>

										<% if request.form("frmUPRem")="" then %>
											<input type="hidden" name="frmUPRem" value="<%=absValRemaining%>">
											<input type="hidden" name="frmUPOldPmtRefNo" value="<%=oldPmtRefNo%>">									
											<input type="hidden" name="frmUPSDate" value="<%=upSDate%>">									
											<input type="hidden" name="frmUPEDate" value="<%=upEDate%>">
											<input type="hidden" name="frmUPElapsedTime" value="<%=ElapsedTime%>">									
										<% else %>
											<input type="hidden" name="frmUPRem" value="<%=request.form("frmUPRem")%>">
											<input type="hidden" name="frmUPOldPmtRefNo" value="<%=request.form("frmUPOldPmtRefNo")%>">									
											<input type="hidden" name="frmUPSDate" value="<%=request.form("frmUPSDate")%>">									
											<input type="hidden" name="frmUPEDate" value="<%=request.form("frmUPEDate")%>">
											<input type="hidden" name="frmUPElapsedTime" value="<%=request.form("frmUPElapsedTime")%>">									
										<% end if %>
								<% end if %>

						
                        <% if typeGroupID<>"" OR payMode=2 then %>
                        <tr valign="middle"> 
                          <%
								 set rsProduct = Server.CreateObject("ADODB.Recordset")
						
								'create SQL select query string
								if payMode=2 then

									strSQL = "SELECT Products.ProductID, Description, OnlinePrice, PRODUCTS.PricePerSession, ClientCredit, Editable, ProductNotes FROM Products WHERE Discontinued=0 AND wsShow=1 AND [Delete]=0 AND GiftCertificate=0 AND Products.ClientCredit>0;"
								''Enrollment Cases!?!
								elseif payMode>=3 and payMode < 6 then
									strSQL = "SELECT Products.ProductID, Description, OnlinePrice, PRODUCTS.PricePerSession, ClientCredit, ProductNotes, Editable, EnableTax1, EnableTax2, EnableTax3, EnableTax4, EnableTax5 "
									strSQL = strSQL & "FROM Products INNER JOIN tblServiceCosts ON Products.ProductID=tblServiceCosts.ProductID LEFT OUTER JOIN (SELECT ProductID FROM tblProductSeriesTypeSetting WHERE (SeriesTypeID = " & sqlClean(cMembershipID) & ") AND (Setting = 1)) AS MemLevel ON PRODUCTS.ProductID = MemLevel.ProductID LEFT OUTER JOIN (SELECT COUNT(*) AS NumRestrictions, ProductID FROM tblProductSeriesTypeSetting WHERE (Setting = 1) GROUP BY ProductID) AS MemRestrict ON PRODUCTS.ProductID = MemRestrict.ProductID "
									' BQL 49_2629 - Removed Typegroup restriction out of this query
									strSQL = strSQL & "WHERE tblServiceCosts.RefType=1 AND tblServiceCosts.RefID=" & sqlClean(pClassID) & " AND Discontinued=0 AND wsShow=1 AND CategoryID<>22 AND CategoryID<>23 AND [Delete]=0 "
									if NOT isMember then ' CB - 45_1276 Members Only
										strSQL = strSQL & " AND ((MemRestrict.NumRestrictions = 0) OR (MemRestrict.NumRestrictions IS NULL)) "
									else
										strSQL = strSQL & " AND ((MemLevel.ProductID IS NOT NULL) "
										if mem_AllowNonMemberPurchases then
											strSQL = strSQL & " OR (MemRestrict.NumRestrictions = 0) OR (MemRestrict.NumRestrictions IS NULL) "
										end if
										strSQL = strSQL & ") "
									end if
									strSQL = strSQL & " ORDER BY Description"
								elseif payMode = 6 then ' Gift Cards
						 			strSQL = "SELECT PRODUCTS.ProductID, Description, OnlinePrice, PRODUCTS.PricePerSession, ClientCredit, ProductNotes, Editable, EnableTax1, EnableTax2, EnableTax3, EnableTax4, EnableTax5 "
									strSQL = strSQL & "FROM PRODUCTS WHERE Discontinued=0 AND wsShow=1 AND Products.GiftCertificate = 1 AND [Delete]=0 AND Products.CategoryID = 22 "
									strSQL = strSQL & "ORDER BY PRODUCTS.Description;"
								else
									if unPaidCaseRec = "true" then
										if request.form("frmUPRem")="" then
											strSQL = "SELECT PRODUCTS.ProductID, PRODUCTS.ProductNotes, PRODUCTS.ClientCredit, PRODUCTS.Editable, PRODUCTS.OnlinePrice, PRODUCTS.PricePerSession, PRODUCTS.Description, PRODUCTS.Class, PRODUCTS.wsShow, PRODUCTS.Type, PRODUCTS.Count, PRODUCTS.Duration, PRODUCTS.TypeGroup "
											strSQL = strSQL & "FROM PRODUCTS WHERE Products.Class=1 AND Discontinued=0 AND [Delete]=0 AND wsShow=1 AND [Type]<>9 AND TypeGroup=" & sqlClean(TypeGroupID) & " AND PRODUCTS.Count>=" & sqlClean(absValRemaining) & " AND PRODUCTS.Duration>=" & sqlClean(ElapsedTime) & " ORDER BY Description;"
										else
											strSQL = "SELECT PRODUCTS.ProductID, PRODUCTS.ProductNotes, PRODUCTS.ClientCredit, PRODUCTS.Editable, PRODUCTS.OnlinePrice, PRODUCTS.PricePerSession, PRODUCTS.Description, PRODUCTS.Class, PRODUCTS.wsShow, PRODUCTS.Type, PRODUCTS.Count, PRODUCTS.Duration, PRODUCTS.TypeGroup "
											strSQL = strSQL & "FROM PRODUCTS WHERE Products.Class=1 AND Discontinued=0 AND [Delete]=0 AND wsShow=1 AND [Type]<>9 AND TypeGroup=" & sqlClean(TypeGroupID) & " AND PRODUCTS.Count>=" & sqlClean(request.form("frmUPRem")) & " AND PRODUCTS.Duration>=" & sqlClean(request.form("frmUPElapsedTime")) & " ORDER BY Description;"
										end if
									else
										if request.form("frmProdVTID")<>"" AND request.form("frmProdVTID")<>"0" then
											strSQL = "SELECT Products.ProductID, Description, OnlinePrice, PRODUCTS.PricePerSession, ProductNotes, ClientCredit, Editable FROM Products, tblProductVisitTypes "
											strSQL = strSQL & "WHERE Products.ProductID=tblProductVisitTypes.ProductID AND tblProductVisitTypes.VisitTypeID=" & sqlClean(request.form("frmProdVTID")) & " AND Class=1 AND Discontinued=0 AND wsShow=1 AND [Delete]=0 AND TypeGroup=" & sqlClean(TypeGroupID)
											strSQL = strSQL & " ORDER BY Description;"
										else									
											strSQL = "SELECT ProductID, Description, OnlinePrice, PRODUCTS.PricePerSession, ProductNotes, ClientCredit, Editable FROM Products "
											strSQL = strSQL & "WHERE Class=1 AND Discontinued=0 AND wsShow=1 AND [Delete]=0 AND TypeGroup=" & sqlClean(TypeGroupID)
											strSQL = strSQL & " ORDER BY Description;"
										end if
									end if
								end if
						'response.write debugSQL(strSQL, "SQL")
								rsProduct.CursorLocation = 3
								rsProduct.open strSQL, cnWS
								Set rsProduct.ActiveConnection = Nothing

								
								if not rsProduct.EOF then
								
									if rsProduct.RecordCount = 1 then 'if payMode>4 and payMode < 6 then
										productID = rsProduct("ProductID")
										prodSelected = true
										prodNotes = rsProduct("ProductNotes")
										if rsProduct("PricePerSession") then
											pProdCost = rsProduct("OnlinePrice")*frmRtnNumSessions
										else
											pProdCost = rsProduct("OnlinePrice")
										end if
										pClientCredit = rsProduct("ClientCredit")
										isEditable = rsProduct("Editable")
										noProd = "false"
										prodName = rsProduct("Description")
										TTax = 0
									   if rsProduct("EnableTax1") then
											TTax = TTax + CDbl(nTax1)
									   end if
									   if rsProduct("EnableTax2") then
											TTax = TTax + CDbl(nTax2)
									   end if
									   if rsProduct("EnableTax3") then
											TTax = TTax + CDbl(nTax3)
									   end if
									   if rsProduct("EnableTax4") then
											TTax = TTax + CDbl(nTax4)
									   end if
									   if rsProduct("EnableTax5") then
											TTax = TTax + CDbl(nTax5)
									   end if
%>
                          <input type="hidden" name="requiredoptPurchaseItem" value="<%=rsProduct("ProductID")%>">
                          <td colspan="2"> <b>&nbsp;<%=rsProduct("Description")%> 
                            at <% if rsProduct("PricePerSession") then response.write FmtCurrency(frmRtnNumSessions*rsProduct("OnlinePrice")) else response.write FmtCurrency(rsProduct("OnlinePrice")) end if %></b> 
                            <%									
									else
%>
                          <td width="1%" height="18"><b>&nbsp;Which <% if payMode=6 then %><%=DisplayPhrase(phraseDictionary,"Giftcardtext")%><% else %>service<% end if %> would you like?&nbsp;</b></td>
                          <td height="18" align="left"> 
                            <select name="requiredoptPurchaseItem" onChange="reloadPage2();">
                              <option value="0">Select Item</option>
                              <%
								dim tmpProdStr
								dim tmpProdID		
								intCount=0
						
								 Do While not rsProduct.EOF
						
									intCount = intCount + 1
					 				if rsProduct("PricePerSession") then %>
                              <option value="<%=rsProduct("ProductID")%>" name="<% response.write rsProduct("Description") & " for " & FmtCurrency(frmRtnNumSessions*rsProduct("OnlinePrice")) %>"
								<% 	else %>
                              <option value="<%=rsProduct("ProductID")%>" name="<% response.write rsProduct("Description") & " for " & FmtCurrency(rsProduct("OnlinePrice")) %>"
								<%	end if
									if productID=rsProduct("ProductID") then
										Response.Write " selected" 
										prodSelected = true
										prodNotes = rsProduct("ProductNotes")
										pProdCost = rsProduct("OnlinePrice")
										pClientCredit = rsProduct("ClientCredit")
										isEditable = rsProduct("Editable")
										noProd = "false"
										prodName = rsProduct("Description")										
									end if
						%>
						><%=rsProduct("Description")%> at <% if rsProduct("PricePerSession") then response.write FmtCurrency(frmRtnNumSessions*rsProduct("OnlinePrice")) else response.write FmtCurrency(rsProduct("OnlinePrice")) end if %></option>
                              <%		
									rsProduct.MoveNext
								Loop
						%>
                            </select>
<script type="text/javascript"> 
      if (document.frmCC.requiredoptPurchaseItem.options.length == 2 && document.frmCC.requiredoptPurchaseItem.options[0].selected ) { 
           document.frmCC.requiredoptPurchaseItem.options[1].selected = true;
		   reloadPage2();
      } 
</script>
                            <input type="hidden" name="frmProdCost" value="<%=pProdCost%>">
                            <%
										if prodSelected then
											if NOT isNull(prodNotes) then
												response.write "<br /><b>Notes:</b></b>&nbsp;" & prodNotes
											end if
										end if

							end if	'''payMode > 3 and < 6
						else
							if payMode<>6 then 
								if payMode=2 then
									%><td colspan="2"><br /><b><%=DisplayPhrase(phraseDictionary,"Nocreditpayments")%></b><%
									noProd = "true"
							
								elseif typeGroupID<>0 then
							' BQL 45_2323
							ss_ClientContactEmail = checkStudioSetting("tblGenOpts", "ClientContactEmail")
									%><td colspan="2"><br /><span style="font-size:1.2em;color:red;"><%=DisplayPhrase(phraseDictionary,"Notenabled")%></span>
									<%	if ss_ClientContactEmail<>"" then %>
										<br /><br /><img src="<%= contentUrl("/asp/images/email2.gif") %>"><a href="mailto:<%=ss_ClientContactEmail%>">Contact <%=session("StudioName")%></a>
									<%	end if %>
										<script type="text/javascript">document.getElementById('purchPol').style.display = 'none';</script>
								<%	noProd = "true"
								else
									%><td colspan="2"><br /><%=DisplayPhrase(phraseDictionary,"Pleaseselecttype")%><%
									noProd = "true"
								end if
							end if
						end if
						
		rsProduct.close
%>
                          </td>
                        </tr>
                        <% if payMode=2 AND noProd="false" then %>
                        <%

				if request.form("requiredtxtPrice")<>"" then
					if isNumeric(request.form("requiredtxtPrice")) then
						frmPrice = FormatNumber(request.form("requiredtxtPrice"))
					else
						if isEditable then
							frmPrice = ""
						else
							frmPrice = FormatNumber(pProdCost)
						end if
					end if
				else
					if isEditable then
						frmPrice = ""
					else
						frmPrice = FormatNumber(pProdCost)
					end if
				end if

				if request.form("requiredtxtCreditAmount")<>"" then
					if isNumeric(request.form("requiredtxtCreditAmount")) then
						frmCreditAmt = FormatNumber(request.form("requiredtxtCreditAmount"))
					else
						frmCreditAmt = FormatNumber(pClientCredit)
					end if
				else
					frmCreditAmt = FormatNumber(pClientCredit)
				end if


							if NOT isEditable then
%>
                        <tr> 
                          <td class="mainText" width="1%" nowrap ><b>&nbsp;<%=xssStr(allHotWords(69))%>:</b></td>
                          <td align="left"> 
                            <input type="text" name="requiredtxtPrice" size="6" maxlength="12" value="<%=frmPrice%>" disabled>
                          </td>
                        </tr>
                        <input type="hidden" name="requiredtxtPrice" value="<%=frmPrice%>">						
                        <tr> 
                          <td class="mainText" width="1%" nowrap ><b>&nbsp;<%=DisplayPhrase(phraseDictionary,"Creditamounttext")%>:</b></td>
                          <td align="left"> 
                            <input type="text" name="requiredtxtCreditAmount" size="6" maxlength="12" value="<%=frmCreditAmt%>" disabled>
                          </td>
                        </tr>
                        <input type="hidden" name="requiredtxtCreditAmount" value="<%=frmCreditAmt%>">
            			  <% else '''editable %>

                        <tr> 
                          <td class="mainText" width="1%" nowrap ><b>&nbsp;<%=xssStr(allHotWords(69))%>:</b></td>
                          <td align="left"> 
                            <input type="text" name="requiredtxtPrice" size="6" maxlength="12" value="<%=frmPrice%>">
                            <span class="textSmall"> &nbsp;(ex.: 100.00)</span> 
                          </td>
                        </tr>
                        <input type="hidden" name="requiredtxtCreditAmount" value="<%=frmPrice%>">
			               <% end if %>

                        <% end if ''Credit Case, Prod Selected %>
                        <% if payMode=5 AND noProd="false" then %>
                        <tr valign="middle"> 
                          <td width="1%" height="18"> <b>&nbsp;<%=DisplayPhrase(phraseDictionary,"Pleasechoose")%> <%=xssStr(allHotWords(0))%>.&nbsp;</b></td>
                          <td height="18"> 
                            <select name="requiredopt<%=session("ResourceDisplayName")%>Type" onChange="reloadPage2();">
                              <option value="0">Select</option>
                              <%
		''Query ResourceTypes for this Product and the count of Resources for each Type

		strSQL = "SELECT DISTINCT Count(tblResourceTypes.Name) AS CountOfResources, tblResourceTypes.Name, tblResourceTypes.ResourceTypeID, tblProductResourceTypes.AddAmount "
		strSQL = strSQL & "FROM tblResources INNER JOIN (tblProductResourceTypes INNER JOIN tblResourceTypes ON tblProductResourceTypes.ResourceTypeID = tblResourceTypes.ResourceTypeID) ON tblResources.ResourceTypeID = tblResourceTypes.ResourceTypeID "
		strSQL = strSQL & "WHERE (((tblProductResourceTypes.ProductID)=" & sqlClean(productID) & ")) "
		strSQL = strSQL & "GROUP BY tblResourceTypes.Name, tblResourceTypes.ResourceTypeID, tblProductResourceTypes.AddAmount ORDER BY tblProductResourceTypes.AddAmount;"

		rsProduct.CursorLocation = 3
		rsProduct.open strSQL, cnWS
		Set rsProduct.ActiveConnection = Nothing

		Do While NOT rsProduct.EOF


			strSQL = "SELECT Count(ResourceTypeID) AS ResTypeCount FROM [VISIT DATA] WHERE ClassDate=" & DateSep & sqlClean(classDate) & DateSep & " AND ResourceTypeID=" & sqlClean(rsProduct("ResourceTypeID"))
			rsWSDB.CursorLocation = 3
			rsWSDB.open strSQL, cnWS
			Set rsWSDB.ActiveConnection = Nothing


			ResourceTypeAvail = false

			if rsWSDB.EOF then
				ResourceTypeAvail = true
			else
				if rsProduct("CountOfResources") > rsWSDB("ResTypeCount") then
					ResourceTypeAvail = true
				end if
			end if
			rsWSDB.close
			
			if ResourceTypeAvail then
%>
                              <option value="<%=rsProduct("ResourceTypeID")%>" name="<%=rsProduct("Name")%> Additional Cost: <%=FmtCurrency(rsProduct("AddAmount"))%>"
<%
				if pResourceType = rsProduct("ResourceTypeID") then
					response.write " selected "
					pAddAmount = rsProduct("AddAmount")
				end if
%>

				><%=rsProduct("Name")%> Additional Cost: <%=FmtCurrency(rsProduct("AddAmount"))%></option>
                              <%
			end if
			rsProduct.MoveNext
		Loop
		rsProduct.close
%>
                            </select>
				<script type="text/javascript">
					document.frmCC.requiredopt<%=session("ResourceDisplayName")%>Type.options[0].text = "Select <%=jsEscDouble(allHotWords(6))%> Type";
				</script>
                            <input type="hidden" name="frmAddAmount" value="<%=pAddAmount%>">
                          </td>
                        </tr>
                        <tr> 
                          <td COLSPAN="2" height="1" width="100%" align="left" style="background-color:<%=session("pageColor4")%>;"><img src="<%= contentUrl("/asp/images/trans.gif") %>" width="100%" height="1"></td>
                        </tr>
                        <tr valign="middle"> 
                          <td width="1%" height="18" nowrap style="background-color:#F2F2F2;"> 
                            <b>&nbsp;<%=DisplayPhrase(phraseDictionary,"Subtotal")%>:&nbsp;</b></td>
                          <td width="70%" height="18" style="background-color:#F2F2F2;"><b>&nbsp;<%=FmtCurrency(pProdCost + pAddAmount)%></b></td>
                        </tr>
                        <tr valign="middle"> 
                          <td width="1%" height="18" nowrap style="background-color:#FAFAFA;"> 
                            <b>&nbsp;<%=xssStr(allHotWords(71))%>:&nbsp;</b></td>
                          <td width="70%" height="18" style="background-color:#FAFAFA;"><b>&nbsp;<%=FmtCurrency( (pProdCost + pAddAmount)*TTax )%></b></td>
                        </tr>
                        <tr> 
                          <td COLSPAN="2" height="1" width="100%" align="left" style="background-color:<%=session("pageColor4")%>;"><img src="<%= contentUrl("/asp/images/trans.gif") %>" width="100%" height="1"></td>
                        </tr>
                        <tr valign="middle"> 
                          <td width="1%" height="18" nowrap style="background-color:#F2F2F2;"> 
                            <b>&nbsp;<%=DisplayPhrase(phraseDictionary,"Totalcosttext")%>: <%=xssStr(allHotWords(3))%> 
                            &nbsp;</b></td>
                          <td width="70%" height="18" style="background-color:#F2F2F2;"><b>&nbsp;<%=FmtCurrency( (pProdCost + pAddAmount) + ((pProdCost + pAddAmount)*TTax) )%></b></td>
                        </tr>
                        <tr valign="middle"> 
                          <td colspan="2" height="6" width="100%" align="left"></td>
                        </tr>
                        <% end if ''payMode = 5 %>
                        <% if (payMode = 4 or payMode = 5) AND noProd="false" then %>
                        <tr valign="middle"> 
                          <td width="1%" height="18"> <b>&nbsp;<%=DisplayPhrase(phraseDictionary,"Selectinitialpayment")%>:&nbsp;</b></td>
                          <td height="18"><b></b> 
                            <select name="requiredoptDeposit" onChange="reloadPage2();">
                              <option value="0">Select Payment</option>
                              <%
								noProd = "true"							  

		dim pmtProdCount
		strSQL = "SELECT COUNT(DISTINCT tblServiceCosts.ProductID) as ProdCount FROM tblServiceCosts INNER JOIN PRODUCTS ON Products.ProductID=tblServiceCosts.ProductID WHERE tblServiceCosts.RefType=1 AND tblServiceCosts.RefID=" & sqlClean(pClassID) & " AND Products.ClientCredit<>0 "
'response.write debugSQL(strSQL, "SQL")
		rsProduct.CursorLocation = 3
		rsProduct.open strSQL, cnWS
		Set rsProduct.ActiveConnection = Nothing

		pmtProdCount = clng(rsProduct("ProdCount"))
		
		rsProduct.close

		if pmtProdCount = 0 then
			strSQL = "SELECT DISTINCT Products.ProductID, Products.Description, Products.OnlinePrice, Products.PricePerSession, Products.ProductNotes, Products.Editable FROM Products LEFT OUTER JOIN tblServiceCosts ON (Products.ProductID=tblServiceCosts.ProductID AND tblServiceCosts.RefID=" & sqlClean(pClassID) & ") WHERE (Products.Discontinued=0 AND Products.GiftCertificate=0 AND Products.[Delete]=0 AND wsShow = 1) AND ((Products.ClientCredit>0 AND tblServiceCosts.ProductID IS NULL) or (Products.ProductID = " & sqlClean(productID) & ") or ((tblServiceCosts.ProductID IS NOT NULL) AND tblServiceCosts.RefType=1 AND tblServiceCosts.RefID=" & sqlClean(pClassID) & "))"
		else
			strSQL = "SELECT Products.ProductID, tblServiceCosts.ProductID as SCPID, Description, OnlinePrice, PricePerSession, ClientCredit, ProductNotes, Editable, EnableTax1, EnableTax2, EnableTax3, EnableTax4, EnableTax5 FROM Products INNER JOIN tblServiceCosts "
			strSQL = strSQL & " ON Products.ProductID=tblServiceCosts.ProductID WHERE tblServiceCosts.RefType=1 AND tblServiceCosts.RefID=" & sqlClean(pClassID) & " AND Products.GiftCertificate=0 AND Discontinued=0 AND [Delete]=0 AND wsShow = 1 "
			strSQL = strSQL & " ORDER BY Description;"
		end if

		'response.write debugSQL(strSQL, "SQL")
		rsProduct.CursorLocation = 3
		rsProduct.open strSQL, cnWS
		Set rsProduct.ActiveConnection = Nothing

		Do While NOT rsProduct.EOF
%>
                              <% if rsProduct("Editable") then %><br />
								<option value="<%=rsProduct("ProductID")%>" name="<% response.write rsProduct("Description") %>"							
							<% else 
								if rsProduct("PricePerSession") then %>
                            		<option value="<%=rsProduct("ProductID")%>" name="<% response.write rsProduct("Description") & " at " & FmtCurrency(frmRtnNumSessions*rsProduct("OnlinePrice")) %>"
							<%	else %>
                            		<option value="<%=rsProduct("ProductID")%>" name="<% response.write rsProduct("Description") & " at " & FmtCurrency(rsProduct("OnlinePrice")) %>"
							<% end if %>
						<% end if %>
						<%
									if DepositProdID=rsProduct("ProductID") then
										Response.Write " selected" 
										noProd = "false"
										InitPayEditable = rsProduct("Editable")
									end if
						%>
						><%=rsProduct("Description")%> 
                              <% if NOT rsProduct("Editable") then %>
                              at 
							  <% if rsProduct("PricePerSession") then response.write FmtCurrency(frmRtnNumSessions*rsProduct("OnlinePrice")) else response.write FmtCurrency(rsProduct("OnlinePrice")) end if %> 
                              <% end if %>
                              </option>
                              <%
			rsProduct.MoveNext
		Loop
		rsProduct.close
%>
                            </select>
                          </td>
                        </tr>
                        <%  if InitPayEditable then %>
                        <tr> 
                          <td width="1%" nowrap style="background-color:#FAFAFA;"><b>&nbsp;<%=DisplayPhrase(phraseDictionary,"Paymentamounttext")%>:</b></td>
                          <td width="82%" style="background-color:#FAFAFA;"> 
                            <input type="text" name="requiredtxtPaymentAmount" maxlength="10" size="4" value="<%=request.form("requiredtxtPaymentAmount")%>" onBlur="updateCreditAmt(this);" class="date">
                            &nbsp;<span class="textSmall">(ex.: 100.00)</span> 
                          </td>
                        </tr>
                        <input type="hidden" name="requiredtxtCreditAmount" value="<%=request.form("requiredtxtPaymentAmount")%>">
                        <%
				end if '''payment editable
		end if ''payMode >= 4

if noProd="false" then


	' Determine whether or not we need to ask them about their account balance.
	partialBalance = 1 ' default to displaying credit card payment form
	CltBalance2 = getAccountBalance(clientID, "", "")
	
	if balChoice = 0 then
		if (cdbl(CltBalance2) > 0 and payMode < 4) OR ((payMode = 4 OR payMode = 5) AND CltBalance2 >= DepositAmount) then %>

                        <tr>
                          <td align="left"> 
                            <b>&nbsp;<%=DisplayPhrase(phraseDictionary,"Useaccountbalance")%></b><br />&nbsp;<%=DisplayPhrase(phraseDictionary,"Currentbalance")%>:<br />&nbsp;<%=FmtCurrency(CltBalance2)%>
							</td>
							<td valign="top">
								<input name="optApplyBalance" type="radio" value="yes" onClick="reloadPage3()">Yes
								<input name="optApplyBalance" type="radio" value="no" onClick="reloadPage3()">No
							</td>
                        </tr>
                        <tr> 
                          <td colspan="2" height="6" width="100%" align="left" valign="middle"></td>
                        </tr>
                        <tr> 
                          <td COLSPAN="2" height="1" width="100%" align="left" style="background-color:<%=session("pageColor4")%>;"><img src="<%= contentUrl("/asp/images/trans.gif") %>" width="100%" height="1"></td>
                        </tr>
<% 		else ' We haven't made a choice, but we don't have one to make, set balChoice to 1
			balChoice = 1
		end if
		
	else  ' balChoice = 1
		if request.querystring("applyBal")="yes" then 
			if pProdCost >= CltBalance2 then ' only apply partial balance
				balanceApplied = CltBalance2
				partialBalance = 1
			else ' pProdCost < CltBalance2
				balanceApplied = pProdCost
				partialBalance = 0
			end if		
%>
                        <tr>
                          <td align="left"> 
                            <b>&nbsp;<%=DisplayPhrase(phraseDictionary,"Appliedbalancetext")%>:</b>
						  </td>
						  <td>
								<%=FmtCurrency(balanceApplied)%>
								<input type="hidden" name="frmAccountBalanceApplied" value="<%=balanceApplied%>">
		<%  if partialBalance = 0 then %>
								<input type="hidden" name="requiredoptPaymentMethod" value="16">
		<%	end if %>
						  </td>
                        </tr>
		<% if partialBalance = 1 then %>
                        <tr>
                          <td align="left"> 
                            <b>&nbsp;<%=DisplayPhrase(phraseDictionary,"Remainingbalance")%>:</b>
						  </td>
						  <td>
								<%=FmtCurrency(pProdCost - balanceApplied)%>
						  </td>
                        </tr>						
		<% end if %>
						<tr><td colspan="2">&nbsp;</td></tr>
<%		end if
	end if

	if balChoice = 1 and partialBalance = 1 then  
	'	if balanceApplied > 0 then	
%>
<!-- input type="hidden" name="requiredoptPaymentMethod" value="CC-Acct"-->
<%	'	else %>
<input type="hidden" name="requiredoptPaymentMethod" value="CC">
<% 	'	end if %>

			<% if payMode = 6 then ' Gift Card Purchase %>
                        <tr> 
                          <td COLSPAN="2" height="1" width="100%" align="left" style="background-color:<%=session("pageColor4")%>;"><img src="<%= contentUrl("/asp/images/trans.gif") %>" width="100%" height="1"></td>
                        </tr>
                        <tr> 
                          <td class="mainText2" nowrap style="background-color:<%=session("pageColor2")%>;" colspan="2"><b>&nbsp;<%=DisplayPhrase(phraseDictionary,"Whogiftcard")%></b></td>
                        </tr>
                        <tr> 
                          <td COLSPAN="2" height="1" width="100%" align="left" style="background-color:<%=session("pageColor4")%>;"><img src="<%= contentUrl("/asp/images/trans.gif") %>" width="100%" height="1"></td>
                        </tr>
                        <tr> 
                          <td colspan="2" height="4" width="100%" align="left"></td>
                        </tr>
						<tr id="eGiftCardRowOn">
							<td colspan="2">
								<table class="mainText" width="100%" cellspacing="0">
									<tr>
										<td width="2%">&nbsp;</td>
										<td width="30%">To:</td>
										<td align="left"><input type="text" size="40" maxlength="50" name="requiredtxtGiftCardTo"></td>
									</tr>
									<tr height="1"><td colspan="3"><img src="<%= contentUrl("/asp/images/trans.gif") %>" height="1" width="1"></td></tr>
									<tr>
										<td>&nbsp;</td>
										<td>From:</td>
										<td align="left"><input type="text" size="40" maxlength="50" name="requiredtxtGiftCardFrom"></td>
									</tr>
									<tr>
										<td>&nbsp;</td>
										<td valign="top">Message:</td>
										<td align="left"><textarea rows=4 cols="50" name="egiftCardMessage"></textarea></td>
									</tr>
									<tr height="10"><td colspan="3"><img src="<%= contentUrl("/asp/images/trans.gif") %>" height="20" width="1"></td></tr>
									<tr>
										<td>&nbsp;</td>
										<td colspan="2"><B><%=DisplayPhrase(phraseDictionary,"Giftcardemail")%>:</B></td>
									</tr>
									<tr height="10"><td colspan="3"><img src="<%= contentUrl("/asp/images/trans.gif") %>" height="20" width="1"></td></tr>
									<tr>
										<td>&nbsp;</td>
										<td><%=DisplayPhrase(phraseDictionary,"Recipemailaddr")%>:</td>
										<td align="left"><input type="text" size="40" name="egiftCardMail1"></td>
									</tr>
									<tr height="1"><td colspan="3"><img src="<%= contentUrl("/asp/images/trans.gif") %>" height="1" width="1"></td></tr>
									<tr>
										<td>&nbsp;</td>
										<td><%=DisplayPhrase(phraseDictionary,"Confirmrecipemailaddr")%>:</td>
										<td align="left"><input type="text" size="40" name="egiftCardMail2"></td>
									</tr>
									<tr height="5"><td colspan="3"><img src="<%= contentUrl("/asp/images/trans.gif") %>" height="5" width="1"></td></tr>
								</table>
							</td>
						</tr>
			<% end if %>


                        <tr> 
                          <td COLSPAN="2" height="1" width="100%" align="left" style="background-color:<%=session("pageColor4")%>;"><img src="<%= contentUrl("/asp/images/trans.gif") %>" width="100%" height="1"></td>
                        </tr>
                        <tr> 
                          <td class="mainText2" nowrap style="background-color:<%=session("pageColor2")%>;" colspan="2"><b>&nbsp;<%=DisplayPhrase(phraseDictionary,"Paymentinformation")%> </b></td>
                        </tr>
                        <tr> 
                          <td COLSPAN="2" height="1" width="100%" align="left" style="background-color:<%=session("pageColor4")%>;"><img src="<%= contentUrl("/asp/images/trans.gif") %>" width="100%" height="1"></td>
                        </tr>
                        <tr> 
                          <td colspan="2" height="4" width="100%" align="left"></td>
                        </tr>			

                        <%
	if noProd="false" then
         set rsCCInfo = Server.CreateObject("ADODB.Recordset")


        'create SQL select query string
        strSQL = "SELECT CreditCardNo, ExpMonth, ExpYear, ccType FROM tblCCNumbers WHERE ClientID = " & sqlClean(clientID)

        rsCCInfo.CursorLocation = 3
rsCCInfo.open strSQL, cnWS
Set rsCCInfo.ActiveConnection = Nothing


 if not rsCCInfo.EOF then
	bCCNum = rsCCInfo("CreditCardNo")
 	if (Not isNull(bCCNum)) AND (NOT isNull(rsCCInfo("ExpMonth"))) AND (NOT isNull(rsCCInfo("ExpYear"))) then
		bCCNum = DES_Decrypt(bCCNum,true,null)
		bCCNum = Replace(bCCNum, " ", "")
		if Len(bCCNum) >= 11 then
%>
                        <tr> 
                          <td COLSPAN="2" align="left"> 
                            <input type="radio" name="UseOnFile" value="yes" checked>
                            <b><%=DisplayPhrase(phraseDictionary,"Option1")%>: </b><%=DisplayPhrase(phraseDictionary,"Usemybilling")%> 
		<% if Len(bCCNum)>=15 then %>
			<br />&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<%=rsCCInfo("ccType")%>: xxxxxxxxxxxx<%=Right(bCCNum,4)%> Exp. <%=rsCCInfo("ExpMonth")%>/<%=TRIM(rsCCInfo("ExpYear"))%>
		<% end if %>							
                          </td>
                        </tr>
		<% if Len(bCCNum)<13 then %>
                        <tr> 
                          <td COLSPAN="2" align="left">&nbsp;&nbsp;&nbsp;&nbsp;
						  <%=DisplayPhrase(phraseDictionary,"Enterlast4")%>:&nbsp; 
                            <input type="text" size=4 maxlength="4" name="requiredccLastFour" value="<%=request.form("requiredccLastFour")%>">
                            &nbsp;&nbsp;&nbsp;&nbsp;<%=DisplayPhrase(phraseDictionary,"Thenclick")%> 
                          </td>
                        </tr>
			<% end if %>
                        <tr> 
                          <td colspan="2" height="4" width="100%" align="left"></td>
                        </tr>
                        <tr> 
                          <td COLSPAN="2" height="1" width="100%" align="left" style="background-color:<%=session("pageColor4")%>;"><img src="<%= contentUrl("/asp/images/trans.gif") %>" width="100%" height="1"></td>
                        </tr>
                        <tr> 
                          <td colspan="2" height="4" width="100%" align="left"></td>
                        </tr>
                        <tr> 
                          <td COLSPAN="2" align="left"> 
                            <input type="radio" name="UseOnFile" value="no">
                            <b><%=DisplayPhrase(phraseDictionary,"Option2")%>: </b><%=DisplayPhrase(phraseDictionary,"Supmybilling")%>.&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
                          </td>
                        </tr>
                        <%
		else
%>
                        <input type="hidden" name="UseOnFile" value="new">
<%
		end if	''Bad CC No check
	 end if		''Null CC INfo check
 end if		''Null CC RS
 rsCCInfo.close
 set rsCCInfo = nothing
%>
                        <tr> 
                          <td class="smallTextBlack" colspan="3" align="left"> 
                            <table width="100%" cellspacing="0">
                              <tr class="mainText"> 
                                <td align="left" nowrap width="1%">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
								<%=DisplayPhrase(phraseDictionary,"Creditcardtext")%>:&nbsp;</td>
                                <td colspan="1"> 
                                  <table class="mainText" cellspacing="0">
                                    <tr> 
                                      <td> 
                                        <select class="textSmall" name="requiredoptCCType" value=""
                   onChange="return checkSelect(this, 'a Credit Card');">
                                          <option value="0" selected>Select Credit 
                                          Card 
<%
	set rsMBDB = Server.CreateObject("ADODB.Recordset")
	strSQL = "SELECT tblCCOpts.ccVisa, tblCCOpts.ccCVV, tblCCOpts.ccMasterCard, tblCCOpts.ccAmericanExpress, tblCCOpts.ccDiscover FROM tblCCOpts WHERE tblCCOpts.StudioID = " & sqlClean(session("StudioID"))
	rsMBDB.CursorLocation = 3
	rsMBDB.open strSQL, cnWS
	Set rsMBDB.ActiveConnection = Nothing
%>
                                          <% if rsMBDB("ccVisa") then %>
                                          <option value="Visa">Visa</option>
                                          <% end if %>
                                          <% if rsMBDB("ccMasterCard") then %>
                                          <option value="Master Card">Master Card</option>
                                          <% end if %>
                                          <% if rsMBDB("ccAmericanExpress") then %>
                                          <option value="American Express">American 
                                          Express</option>
                                          <% end if %>
                                          <% if rsMBDB("ccDiscover") then %>
                                          <option value="Discover">Discover</option>
                                          <% end if %>

                                        </select>
                                      </td>
                                      <td class="right">
                                          <% if rsMBDB("ccVisa") then %>
                                          <img src="<%= contentUrl("/asp/images/cclogo_visa.gif") %>" width="42" height="24">
                                          <% end if %>
                                          <% if rsMBDB("ccMasterCard") then %>
                                          <img src="<%= contentUrl("/asp/images/cclogo_mastercard.gif") %>" width="42" height="24">
                                          <% end if %>
                                          <% if rsMBDB("ccAmericanExpress") then %>
                                          <img src="<%= contentUrl("/asp/images/cclogo_amex.gif") %>" width="61" height="24">
                                          <% end if %>
                                          <% if rsMBDB("ccDiscover") then %>
                                          <img src="<%= contentUrl("/asp/images/cclogo_discover.gif") %>" width="44" height="24">
                                          <% end if %>
									  </td>
                                    </tr>
                                  </table>
                                </td>
                              </tr>
                              <tr class="mainText"> 
                                <td nowrap align="left" width="1%"> &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<%=xssStr(allHotWords(51))%>:&nbsp; </td>
                                <td colspan="3"> 
                                  <input type=text name="requiredtxtCCNumber" size=20 value=""
                        onChange="return checkCCLength(this, 'Credit Card Number');">
                                  <input type=hidden name="CCNumber" value="">
                                </td>
                              </tr>
                              <tr class="mainText"> 
                                <td align="left" nowrap width="1%"> &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<%=xssStr(allHotWords(52))%>:&nbsp; </td>
                                <td colspan="3"> 
                                  <select class="textSmall" name="requiredoptExpMonth" value=''
                   onChange="return checkSelect(this, ' a Credit Card Expiration Month');">
                                    <option value="0" selected>Month</option>
                                    <option value="01">1 - January</option>
                                    <option value="02">2 - February</option>
                                    <option value="03">3 - March</option>
                                    <option value="04">4 - April</option>
                                    <option value="05">5 - May</option>
                                    <option value="06">6 - June</option>
                                    <option value="07">7 - July</option>
                                    <option value="08">8 - August</option>
                                    <option value="09">9 - September</option>
                                    <option value="10">10 - October</option>
                                    <option value="11">11 - November</option>
                                    <option value="12">12 - December</option>
                                  </select>
                                  &nbsp;&nbsp; 
                                  <select class="textSmall" name="requiredoptExpYear" value=''
                    onChange="return checkSelect(this, ' a Credit Card Expiration Year');">
                                    <option value="0" selected>Year 
<%
	Dim tmpThisYear, tmpYearCount
	tmpThisYear = Now
	for tmpYearCount = 0 to 20
%>				
			                        <option value="<%=Year(DateAdd("yyyy", tmpYearCount, tmpThisYear))%>"><%=Year(DateAdd("yyyy", tmpYearCount, tmpThisYear))%></option>
<%
	next
%>						
                                  </select>
                                  <input type=hidden name="ExpYear" value=''>
                                </td>
                              </tr>
					<%	if rsMBDB("ccCVV") then %>
                              <tr class="mainText"> 
                                <td align="left" nowrap width="1%">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;CVV2: 
                                </td>
                                <td colspan="3"> 
                                  <input type=text name="requiredtxtCVV2" size=4 value=""
                  maxlength="4">
                                  <span class="textSmallBlack"><%=DisplayPhrase(phraseDictionary,"Wherecvv")%> <a href="#" onClick="popImageWindow('CVV2-amex.gif')">[AMEX]</a>&nbsp;<a href="#" onClick="popImageWindow('CVV2-visa.gif')">[Other]</a></span> </td>
                              </tr>
<%						end if 
						rsMBDB.close
%>
                              <tr class="mainText"> 
                                <td nowrap align="left" width="1%"> &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<%=xssStr(allHotWords(44))%>:&nbsp; 
                                </td>
                                <td colspan="3"> 
                                  <input type=text name="requiredtxtCCName" size="40" value="" onChange="return checkProperNames(this, 'Credit Card Account Name');">
                                </td>
                              </tr>
                              <tr class="mainText"> 
                                <td align="left" nowrap width="1%"> &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<%=xssStr(allHotWords(45))%>:&nbsp; 
                                </td>
                                <td colspan="3"> 
                                  <input type=text name="requiredtxtPayAddress1" size=40 value=""
                onChange="return checkAddress(this, 'Address');">
                                </td>
                              </tr>
                              <tr class="mainText"> 
                                <td nowrap align="left" width="1%"> &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<%=xssStr(allHotWords(47))%>:&nbsp; 
                                </td>
                                <td colspan="3"> 
                                  <input type=text name="requiredtxtPayCity" size=40 value=""
                onChange="return checkCity(this, 'City');">
                                </td>
                              </tr>
                              <tr class="mainText"> 
                                <td nowrap align="left" width="1%"> &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<%=xssStr(allHotWords(48))%>:&nbsp; 
                                </td>
                                <td colspan="3"> 
                                  <select class="textSmall" name="requiredoptPayState"
             onChange="return checkSelect(this, 'a State');">
                                    <option value="0" selected>State 
                                    <option value="OU">Outside USA 
                                    <option value="AL">Alabama 
                                    <option value="AK">Alaska 
                                    <option value="AZ">Arizona 
                                    <option value="AR">Arkansas 
                                    <option value="CA">California 
                                    <option value="CO">Colorado 
                                    <option value="CT">Connecticut 
                                    <option value="DC">District of Columbia 
                                    <option value="DE">Delaware 
                                    <option value="FL">Florida 
                                    <option value="GA">Georgia 
                                    <option value="HI">Hawaii 
                                    <option value="ID">Idaho 
                                    <option value="IL">Illinois 
                                    <option value="IN">Indiana 
                                    <option value="IA">Iowa 
                                    <option value="KS">Kansas 
                                    <option value="KY">Kentucky 
                                    <option value="LA">Louisiana 
                                    <option value="ME">Maine 
                                    <option value="MD">Maryland 
                                    <option value="MA">Massachusetts 
                                    <option value="MI">Michigan 
                                    <option value="MN">Minnesota 
                                    <option value="MS">Mississippi 
                                    <option value="MO">Missouri 
                                    <option value="MT">Montana 
                                    <option value="NE">Nebraska 
                                    <option value="NV">Nevada 
                                    <option value="NH">New Hampshire 
                                    <option value="NJ">New Jersey 
                                    <option value="NM">New Mexico 
                                    <option value="NY">New York 
                                    <option value="NC">North Carolina 
                                    <option value="ND">North Dakota 
                                    <option value="OH">Ohio 
                                    <option value="OK">Oklahoma 
                                    <option value="OR">Oregon 
                                    <option value="PA">Pennsylvania 
                                    <option value="RI">Rhode Island 
                                    <option value="SC">South Carolina 
                                    <option value="SD">South Dakota 
                                    <option value="TN">Tennessee 
                                    <option value="TX">Texas 
                                    <option value="UT">Utah 
                                    <option value="VT">Vermont 
                                    <option value="VA">Virginia 
                                    <option value="WA">Washington 
                                    <option value="WV">West Virginia 
                                    <option value="WI">Wisconsin 
                                    <option value="WY">Wyoming 
                                  </select>
                                  &nbsp;&nbsp; &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<%=xssStr(allHotWords(49))%>:&nbsp; 
                                  <input type=text name="requiredtxtPayZip" size=10 value="">
                                  &nbsp;&nbsp; </td>
                              </tr>
                              <tr class="mainText"> 
                                <td align="left" nowrap width="1%"> &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<%=DisplayPhrase(phraseDictionary,"Countrytext")%>:&nbsp; 
                                </td>
                                <td colspan="3"> 
                                  <select name="requiredoptPayCountry" onChange="return checkSelect(this, 'a Country');">
                                    <option value=0>Select a Country 
                                    <option value="USA" selected>United States 
                                    <option value="Canada">Canada 
                                    <option value="OU">Other 
                                  </select>
                                </td>
                              </tr>
                              <tr align="left" class="mainText"> 
                                <td colspan=2> &nbsp;&nbsp;&nbsp;&nbsp; 
                                  <input type="checkbox" name="SaveInfo">
                                  <%=DisplayPhrase(phraseDictionary,"Storemybillinginfo")%></td>
                              </tr>
                            </table>
                          </td>
                        </tr>

                        <% end if 'no page params %>
                        <%
	end if '''select unpaid to rec or not 
	
end if 'noProd


end if ''no Prod/ResReq


if noProd="false" and balChoice = 1 then  

%>
                        <tr> 
                          <td colspan="3"  height="26" class="right headText">  
                            <% if request.form("requiredoptPaymentMethod")<>"CC" then response.write "<br />" end if %>
								<table align="right" width="25%" cellpadding="3" cellspacing="2" style="cursor:pointer" onClick="checkForm(document.frmCC)">
									<tr><td class="center-ch" style="background-color:<%=session("pageColor4")%>;" class="headText"><B><span style="color:white;"><%=DisplayPhrase(phraseDictionary,"Checkout")%></span></B></td></tr>
								</table>
                            <!--input class="mainText" type="Submit" Value="Check Out" style="font-family:Trebuchet MS; font-size:14pt" id=Submit1 name="Submit1" -->
                          </td>
                        </tr>
                        <tr> 
                          <td colspan="2" height="6" width="100%" align="left" valign="middle"></td>
                        </tr>
                        <tr> 
                          <td COLSPAN="2" height="1" width="100%" align="left" style="background-color:<%=session("pageColor4")%>;"><img src="<%= contentUrl("/asp/images/trans.gif") %>" width="100%" height="1"></td>
                        </tr>
                        <tr class="right"> 
                          <td colspan="3" height="26">&nbsp;</td>
                        </tr>
                        <% end if %>
                        <% end if %>
                      </table>
                    </td>
                  </tr>
                </table>
              </td>
            </tr>
            
          </table>
    </td>
    </tr>

		<!-- #include file="inc_footer.asp" -->

  </table>

</form>
<script type="text/javascript">
var submitOne = false;
function checkForm(sbmtForm) {
	if (!submitOne) {
		submitOne = true;
		if (processSubmit(sbmtForm) == true) {
			sbmtForm.submit();
		}
	}
}

function processSubmit(sbmtForm) {
	var returnVal

<% if payMode=5 then %>
	if (checkSelect(document.frmCC.requiredopt<%=session("ResourceDisplayName")%>Type, "a <%=jsEscDouble(allHotWords(0))%> Type", true)!="")
	{
		alert(checkSelect(document.frmCC.requiredopt<%=session("ResourceDisplayName")%>Type, "a <%=jsEscDouble(allHotWords(0))%> Type", true));
		return false;
	} else if (checkSelect(document.frmCC.requiredoptDeposit, "a Deposit/Inital Payment", true)!="")
	{
		alert(checkSelect(document.frmCC.requiredoptDeposit, "a Deposit/Inital Payment", true));
		return false;
	} else {
<% elseif payMode=6 then %>
	if (document.frmCC.egiftCardMail1.value != document.frmCC.egiftCardMail2.value) 
	{
		alert("Recipient Email address does not match, please correct this error.");
		frmCC.egiftCardMail1.focus();
		return false
	} else {
<% end if %>
		returnVal = checkRequired(sbmtForm);

		if (returnVal==true) {

<% if payMode=2 then %>
			if (confirm("Are you sure you want to purchase this credit for " + frmCC.requiredtxtPrice.value))  {
				returnVal = true;
			} else {
				returnVal = false;
			}

<% elseif request.form("requiredoptPaymentMethod")="16" AND pProdCost > CltBalance AND NOT pAllowCltNegBal then %>
		alert("This <%=jsEscDouble(allHotWords(12))%> does not have enought credit to buy the selected service.\nPlease select a different service, payment method, or first buy more credit for this <%=jsEscDouble(allHotWords(12))%>.");
					returnVal = false;
<% elseif payMode=5 then %>
		if (confirm("Are you sure you want to reserve <%=className%> for <%=FmtCurrency(pProdCost)%>" + "\n<%=jsEscDouble(allHotWords(0))%> Type: " + frmCC.requiredopt<%=session("ResourceDisplayName")%>Type.options[frmCC.requiredopt<%=session("ResourceDisplayName")%>Type.selectedIndex].name + "\nTotal <%=jsEscDouble(allHotWords(3))%> Cost: <%=Replace(FmtCurrency( (pProdCost + pAddAmount) + ((pProdCost + pAddAmount)*TTax) ), "&nbsp;", "")%>\n\nClicking OK will add this reservation to your schedule.\nYour credit card will be charged for the <%=FmtCurrency(DepositAmount)%> initial payment.")){
					returnVal = true;
			} else {
				returnVal = false;
			}
<% elseif payMode=4 then %>
			if (confirm("Are you sure you want to reserve <%=className%> for <%=FmtCurrency(pProdCost)%>" + "\n\nClicking OK will add this reservation to your schedule.\nYour credit card will be charged for the <%=Replace(FmtCurrency(DepositAmount), "&nbsp;", "")%> initial payment."))  {
					returnVal = true;
			} else {
				returnVal = false;
			}
<% else %>
			if (confirm("Are you sure you want to purchase <%=prodName%> for <%=FmtCurrency(pProdCost)%> <% if partialBalance = 1 then %>\nCredit Card will be charged: <%=FmtCurrency(pProdCost - balanceApplied)%> <% end if %>" ))  {
					returnVal = true;
			} else {
				returnVal = false;
			}
<% end if %>

				
		} else {
			returnVal=false;
		}

<% if payMode=5 or payMode=6 then %>
	}
<% end if %>


	return returnVal;
}
</script>

</div>
</body>
</html>
<%
	cnWS.close
	set cnWS = nothing
	

		end if	
%>
