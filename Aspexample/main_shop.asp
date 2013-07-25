<%@ CodePage=65001 %>
<%Option Explicit%>
<!-- #include file="init.asp" -->
<!-- #include file="json/JSON.asp" -->
<!-- #include file="inc_tax_calcs.asp" -->
<%
'Dim SessionFarm
'set SessionFarm = Server.CreateObject("SessionFarm.SFSession")
response.charset="utf-8"
%>

<!-- #include file="inc_dbconn_wsMBPS.asp" -->
<!-- #include file="inc_i18n.asp" -->
<% if session("CR_Memberships") <> 0 then %>
    <!-- #include file="inc_dbconn_regions.asp" -->
    <!-- #include file="inc_dbconn_wsMaster.asp" -->
    <!-- #include file="adm/inc_masterclients_util.asp" -->
<% end if %>
<!-- #include file="adm/inc_acct_balance.asp" -->
<!-- #include file="adm/inc_crypt.asp" -->
<!-- #include file="adm/inc_contract_text.asp" -->
<!-- #include file="inc_chk_cltol.asp" -->
<!-- #include file="adm/inc_discount.asp" -->
<!-- #include file="adm/inc_hotword.asp" -->
<!-- #include file="inc_hack_str.asp" -->
<!-- #include file="adm/inc_chk_holiday.asp" -->
<!-- #include file="inc_tinymcesetup.asp" -->
<!-- #include file="inc_dbconn.asp" -->
<!-- #include file="inc_localization.asp" -->
<%
if request.form("tabID")<>"" AND isNum(request.form("tabID")) then
    session("tabID") = request.form("tabID")
elseif request.querystring("tabID")<>"" AND isNum(request.querystring("tabID")) then
    session("tabID") = request.querystring("tabID")
else
    session("tabID") = 3
end if

	dim PmtSelected, ss_DefaultStreetCorner, first, selectedExpDate
	dim frmRtnClassDate, frmRtnClassEDate, frmRtncont, frmRtntmpDate, frmRtnNumSessions, tmpNumSessions, frmRtnNumDeducted, pRegStartDate, deliveryDate
	dim mem_ServiceDiscountPerc, mem_ProdDiscountPerc, mem_AllowNonMemberPurchases, cMembershipID, curDate, dayCount, quantity

	dim phraseDictionary
	set phraseDictionary = LoadPhrases("ConsumermodeonlinestorePage", 51)

	dim reSchedule, origId
	reSchedule = request.Item("reSchedule")
	origId = request.Item("origId")


    Dim ccMode, pMode, lastpMode, rsEntry, rsEntry2,  clientID, typeGroupID, visitTypeID, productID, pmtProductID, partnerID, unPaidCase, unPaidCaseRec, prodSelected, isGC, isPPGC, isEditable, frmCredit, tmpCount, tmpStr, ss_DemoSite, ContractsOk, PackagesOk, contractStartDate, isContract, contractAPTot, tmpTax, contractTot, sortOrderID, depositSortOrderID, groupID, showFirstAP, contractDiscount, contractDiscountAmt, contractConfirm, strContractText,  ServicesOk
    Dim ss_EnrollResReq, ss_UseApptProdVT, ss_OnlineStore, sCartNumItems, numItemsUpdated, totTax, RetailItemsOk, isMember, noCCAndNoAccountCredit, isEnrollTG,isAppointmentTG, showMakePurchaseButton, UsePerStaffPricing
    Dim ss_ShowServiceNotesOnline, ss_UseOnlineStore, ss_EnableOnlineStore, ss_ClientContactEmail, ss_UseConsumerModeFlash, ss_useUPS, ss_useFedEx, ss_useUSPS, ss_AllowPurchaseWithoutEnroll, ss_EnableSemesters
    Dim ss_AllowFasterShippingMethods, ss_FreeShippingThreshold, ss_ContractBillDays, ss_ShowPackageItemsOnReceipts, prodFound, ss_AllowInStorePickup, ss_AllowOrderOutofStock, ss_AllowDupRes, ss_AccountPaymentsConsumerMode

    set rsEntry = Server.CreateObject("ADODB.Recordset")
    set rsEntry2 = Server.CreateObject("ADODB.Recordset")
        
    dim ss_CountryCode, ss_AllowNegativeBalConsMode, InitPayEditable, relClientID, optReservedFor, ss_reserveForOtherClient
	dim perSessionPriced
    relClientID = ""

    'JM-51_2801
    ss_CountryCode = checkStudioSetting("Studios", "countryCode")
    if ss_countryCode<>"US" AND ss_countryCode<>"GB" AND ss_countryCode<>"CA" AND ss_countryCode<>"AU" then
        ss_AllowNegativeBalConsMode = checkStudioSetting("tblGenOpts", "AllowNegativeBalConsMode")
    else
        ss_AllowNegativeBalConsMode = false
    end if
    ss_AccountPaymentsConsumerMode = checkStudioSetting("tblGenOpts", "AccountPaymentsConsumerMode")
    if session("PartnersEnabled") then
        connectToMBPS
    end if
	  
    strSQL = "SELECT tblGenOpts.UseUPS, tblGenOpts.UseFedEx, tblGenOpts.UseUSPS, tblGenOpts.UseProdVT, tblGenOpts.UseOnlineStore, tblGenOpts.EnableOnlineStore, tblGenOpts.AllowPurchaseWithoutEnroll, "
    strSQL = strSQL & "tblResvOpts.EnrollReqResource, tblResvOpts.CltModeAllowDupRes, tblGenOpts.ShowServiceNotesOnline, tblGenOpts.ClientContactEmail, tblGenOpts.DefaultStreetCorner, "
    strSQL = strSQL & "tblGenOpts.UseConsumerModeFlash, tblGenOpts.AllowFasterShippingMethods, tblGenOpts.FreeShippingThreshold, tblGenOpts.AllowInStorePickup, tblGenOpts.ShowPackageItemsOnReceipts, tblGenOpts.AllowOrderOutofStock, tblGenOpts.UsePerStaffPricing, Studios.SiteType, tblResvOpts.EnableSemesters FROM tblGenOpts INNER JOIN tblResvOpts ON tblGenOpts.StudioID = tblResvOpts.StudioID, Studios WHERE (tblGenOpts.StudioID = " & sqlClean(session("StudioID")) & ")"
    rsEntry.CursorLocation = 3
    rsEntry.open strSQL, cnWS
    Set rsEntry.ActiveConnection = Nothing
        ss_EnrollResReq = rsEntry("EnrollReqResource")
        ss_UseApptProdVT = rsEntry("UseProdVT")
        ss_ShowServiceNotesOnline = CBool(rsEntry("ShowServiceNotesOnline"))
        ss_ClientContactEmail = rsEntry("ClientContactEmail")
        ss_UseOnlineStore = rsEntry("UseOnlineStore")
        ss_EnableOnlineStore = rsEntry("EnableOnlineStore")
        ss_DefaultStreetCorner = rsEntry("DefaultStreetCorner")
        ss_UseConsumerModeFlash  = rsEntry("UseConsumerModeFlash")
        ss_DemoSite = rsEntry("SiteType")
        ss_useUPS = rsEntry("UseUPS")
        ss_useFedEx = rsEntry("UseFedEx")
        ss_useUSPS = rsEntry("UseUSPS")
        ss_AllowFasterShippingMethods = rsEntry("AllowFasterShippingMethods")
        ss_FreeShippingThreshold = rsEntry("FreeShippingThreshold")
        'ss_ContractBillDays = rsEntry("ContractBillDays")
        ss_ShowPackageItemsOnReceipts = rsEntry("ShowPackageItemsOnReceipts")
        ss_AllowInStorePickup = rsEntry("AllowInStorePickup")
        ss_AllowOrderOutofStock = rsEntry("AllowOrderOutofStock")
        UsePerStaffPricing = rsEntry("UsePerStaffPricing")
        ss_AllowDupRes = rsEntry("CltModeAllowDupRes")
        ss_AllowPurchaseWithoutEnroll = rsEntry("AllowPurchaseWithoutEnroll")
        ss_EnableSemesters = rsEntry("EnableSemesters")

    rsEntry.close

    prodSelected = false
    isEnrollTG = false
	isAppointmentTG = false
    showMakePurchaseButton = true
    isGC = false
    isPPGC = false
    isEditable = false
    numItemsUpdated = false
    if session("mvarMIDs")=2 then
            ccMode = true
    else
            ccMode = false
    end if
        
    ' First things first, if we are logging in as a guest, set up the session 
    if request.Form("optLoginAsGuest") <> "" then
        ' Set session vars
        Session("mvarUserId") = 0 ' Online Guest is always 0
            Session("mvarNameFirst") = "Online"
            Session("mvarNameLast") = "Guest"
            Session("Pass") = true
            Session("Admin") = "false"
    end if
        
    ' More importantly, if we are logged in as the guest and there is bad stuff in the cart
    ' Log us out
    dim guestOK
	strSQL = "SELECT COUNT(tblShoppingCart.ProductID) as NonGiftCerts "&_
			"FROM tblShoppingCart "&_
			"LEFT OUTER JOIN PRODUCTS on tblShoppingCart.ProductID = PRODUCTS.ProductID "&_
			"WHERE ((PRODUCTS.GiftCertificate = 0 OR (PRODUCTS.GiftCertificate = 1 AND PRODUCTS.DebitCard = 0)) OR tblShoppingCart.PartnerID <> 0) "&_
			"AND (SessionID = N'" & sqlClean(getSessionGUID()) & "') AND SaleID IS NULL "
    rsEntry.CursorLocation = 3
    rsEntry.open strSQL, cnWS
    Set rsEntry.ActiveConnection = Nothing
    'response.Write(strSQL)
        
    guestOK = CINT(rsEntry("NonGiftCerts")) = 0
    if NOT guestOK AND Session("mvarUserID") = "0" then
        ' Log out the user
        Session("mvarUserId") = empty
        Session("mvarNameFirst") = empty
        Session("mvarNameLast") = empty
        Session("Pass") = false
        Session("Admin") = "false"
        'sessionFarm.Abandon()
        
        ' Add a javascript to refresh the header (This doesn't seem to work, but it's here for
        ' posterity)
        response.Write("<script type=""text/javascript"">")
        'response.Write("$(document).ready(function() {")
        response.Write("parent.mainFrame.location.href = parent.mainFrame.location.href;")
        'response.Write("});") 
        response.Write("</script>")
    end if
    rsEntry.Close

    'Check if services setup to sell online
    strSQL = "SELECT COUNT(*) AS NumRetailItems FROM PRODUCTS WHERE (Discontinued = 0) AND (wsShow = 1) AND ([Delete] = 0) AND (CategoryID >25)"
    rsEntry.CursorLocation = 3
    rsEntry.open strSQL, cnWS
    Set rsEntry.ActiveConnection = Nothing
    RetailItemsOk = false
    if NOT rsEntry.EOF then
            if rsEntry("NumRetailItems")>0 AND ss_EnableOnlineStore then
                    RetailItemsOk = true
            end if
    end if
    rsEntry.close

    'Check if services setup to sell online
    strSQL = "SELECT COUNT(*) AS NumServices FROM PRODUCTS INNER JOIN tblTypeGroup ON PRODUCTS.TypeGroup = tblTypeGroup.TypeGroupID WHERE (Discontinued = 0) AND (wsShow = 1) AND ([Delete] = 0) AND (ContractOnly = 0) AND (CategoryID <=20) AND tblTypeGroup.Active = 1 "
    rsEntry.CursorLocation = 3
    rsEntry.open strSQL, cnWS
    Set rsEntry.ActiveConnection = Nothing
    ServicesOk = false
    if NOT rsEntry.EOF then
            if rsEntry("NumServices")>0 then
                    ServicesOk = true
            end if
    end if
    rsEntry.close

    'Check if contracts setup to sell online
    strSQL = "SELECT COUNT(*) AS NumContracts FROM tblContract INNER JOIN tblContractItem ON tblContract.ContractID = tblContractItem.ContractID WHERE (tblContract.Discontinued = 0) AND (tblContract.Deleted = 0) AND (tblContract.SellOnline = 1) AND (IsPackage=0)"
    rsEntry.CursorLocation = 3
    rsEntry.open strSQL, cnWS
    Set rsEntry.ActiveConnection = Nothing
    ContractsOk = false
    if NOT rsEntry.EOF then
            if rsEntry("NumContracts")>0 then
                    ContractsOk = true
            end if
    end if
    rsEntry.close

    'CB 45_1276
    'Check if packages setup to sell online
    strSQL = "SELECT COUNT(*) AS NumPackages FROM tblContract INNER JOIN tblContractItem ON tblContract.ContractID = tblContractItem.ContractID WHERE (tblContract.Discontinued = 0) AND (tblContract.Deleted = 0) AND (tblContract.SellOnline = 1) AND (IsPackage=1)"
    rsEntry.CursorLocation = 3
    rsEntry.open strSQL, cnWS
    Set rsEntry.ActiveConnection = Nothing
    PackagesOk = false
    if NOT rsEntry.EOF then
            if rsEntry("NumPackages")>0 then
                    PackagesOk = true
            end if
    end if
    rsEntry.close
        
    ' BJD: Replaced (rsEntry("UseFedEx") OR rsEntry("UseUPS")) with RetailItemsOk
    if (ss_UseOnlineStore AND ss_EnableOnlineStore AND RetailItemsOk) OR session("PartnersEnabled") then
            ss_OnlineStore = true
    else
            ss_OnlineStore = false
    end if
        

    
    'Count Items in Shopping Cart
	strSQL = "SELECT COUNT(SortOrderID) AS NumItems FROM tblShoppingCart WHERE (SaleID IS NULL) AND DepositSortOrderID IS NULL GROUP BY SessionID HAVING (SessionID = N'" &  sqlClean(getSessionGUID()) & "')"
	rsEntry.CursorLocation = 3
	rsEntry.open strSQL, cnWS
	Set rsEntry.ActiveConnection = Nothing
	if NOT rsEntry.EOF then
		sCartNumItems = rsEntry("NumItems")
	else
		sCartNumItems = 0
	end if
	rsEntry.close

	''SET DEFAULT LAUNCH TAB - 'Default to Retail 1-Services | 2-GC/Credit | 3-Retail | 0-Contracts & Packages
	lastpMode = request.form("frmLastMode")
	if ss_OnlineStore AND RetailItemsOk then
		if request.querystring("justloggedin")<>"" AND sCartNumItems>0 then
			checkPrevPurch(DisplayPhrase(phraseDictionary,"Previntropurchase"))
			pMode=4
		else
			pMode = 3
		end if
	else
		if request.querystring("justloggedin")<>"" AND sCartNumItems>0 then
			checkPrevPurch(DisplayPhrase(phraseDictionary,"Previntropurchase"))
			pMode=4
		elseif ServicesOk then
			pMode = 1
		elseif ContractsOk OR PackagesOk then
			pMode = 0
		else
			pMode = 2
		end if
	end if

	if request.querystring("justloggedin")="" then
		if request.querystring("pMode")<>"" then
				pMode = CINT(request.querystring("pMode"))
		elseif request.form("frmMode")<>"" then
				pMode = CINT(request.form("frmMode"))
		elseif request.querystring("typeGroup")<>"" then
				pMode = 1       ''If Comming from Resv/Booking Service
		end if
	end if


    'JM-54_2847
    if request.querystring("relClientID") <> "" AND isNum(request.querystring("relClientID")) then
        relClientID = request.querystring("relClientID")
    elseif request.form("frmrelClientID") <> "" AND isNum(request.form("frmrelClientID")) then
        relClientID = request.form("frmrelClientID")
    else
        relClientID = ""
    end if

    clientID = session("mvarUserID")
    'JM-55_2848
    ss_reserveForOtherClient = checkStudioSetting("tblGenOpts","reserveForOtherClient")
    if request.querystring("optResfor")<>"" and ss_reserveForOtherClient and relClientID ="" then
        optReservedFor = request.querystring("optResfor")
    elseif request.form("optReservedFor")<>"" and ss_reserveForOtherClient and relClientID ="" then
        optReservedFor = request.form("optReservedFor")
    else
        optReservedFor = ""
    end if
    'response.Write optReservedFor
        
    mem_ServiceDiscountPerc = 0
    mem_ProdDiscountPerc = 0

    'RI - Bug 2843 enrollment register fix
    if NOT ss_AllowDupRes then
		dim classID, courseID, clsStartTime, clsEndTime, clsLocation, chkClientOverLap, classDate, registerStartDate
		classID = request.QueryString("cid")
		courseID = request.QueryString("courseid")
		classDate = Request.QueryString("classDate")
		if request.form("txtRegStartDate")<>"" then
			Call SetLocale(session("mvarLocaleStr"))
				registerStartDate = CDATE(request.form("txtRegStartDate"))
			Call SetLocale("en-us")         
		end if

		if request.Form("optRegisterType") = "openSchedule" OR request.Form("optRegisterType") = "dateForward" then 'Open enrollment/Date forward
            strSQL = "SELECT tblClassSch.ClassID, tblClassSch.ClassDate, tblClassSch.TrainerID, tblClassSch.PayScaleID, tblClassSch.StartTime, tblClassSch.EndTime, tblClasses.LocationID, tblClassDescriptions.VisitTypeID, tblClassDescriptions.ClassPayment, tblVisitTypes.NumDeducted FROM tblClasses INNER JOIN tblClassSch ON tblClasses.ClassID = tblClassSch.ClassID INNER JOIN tblClassDescriptions ON tblClassSch.DescriptionID = tblClassDescriptions.ClassDescriptionID INNER JOIN tblVisitTypes ON tblClassDescriptions.VisitTypeID = tblVisitTypes.TypeID "
            if courseID<>"" then 
						strSQL = strSQL & "WHERE (tblClasses.CourseID = " & sqlClean(courseID) & ")"
            else
						strSQL = strSQL & "WHERE (tblClassSch.ClassID = " & sqlClean(classID) & ")"
            end if
            if request.Form("optRegisterType") = "dateForward" AND registerStartDate <> "" then				
						strSQL = strSQL & " AND (tblClassSch.ClassDate>=" & DateSep & sqlClean(registerStartDate) & DateSep & ")"     
            end if
            strSQL = strSQL & " ORDER BY tblClassSch.ClassDate "
            rsEntry.CursorLocation = 3
            rsEntry.open strSQL, cnWS
            Set rsEntry.ActiveConnection = Nothing

            dim firstEnroll, overLapFound, checkCltOverlap
            firstEnroll = true
            overLapFound = false
            Do While not rsEntry.EOF
                classID = rsEntry("ClassID")
                clsStartTime = rsEntry("StartTime")
                clsEndTime = rsEntry("EndTime")
                clsLocation = rsEntry("LocationID")
                if request.form("optDay" & FmtDateShort(rsEntry("ClassDate")) & "-" & rsEntry("ClassID")) = "on" OR request.Form("optRegisterType") = "dateForward" then
                    if isNull(rsEntry("StartTime")) then
                        chkClientOverLap = false
                    else
                        if relClientID <> "" then
                            chkClientOverLap = checkClientOverlap(relClientID, clsStartTime, clsEndTime, rsEntry("ClassDate"), false, 0, null )
                        elseif optReservedFor = "" then
                            chkClientOverLap = checkClientOverlap(ClientID, clsStartTime, clsEndTime, rsEntry("ClassDate"), false, 0, null )
                        end if
                    end if
                    if chkClientOverLap then
                        overLapFound = true
                    end if
                end if
                rsEntry.MoveNext
            loop
            rsEntry.Close   

        else 'Normal enrollment
            if clsStartTime <> "null" then
                if relClientID <> "" then
                    checkCltOverlap = checkClientOverlap(relClientID, clsStartTime, clsEndTime, classDate, false, 0, null ) 
                elseif optReservedFor = "" then
                    checkCltOverlap = checkClientOverlap(ClientID, clsStartTime, clsEndTime, classDate, false, 0, null ) 
                end if
            else 
                if relClientID <> "" then
                    checkCltOverlap = checkClientOverlap2(relClientID, classDate, classID, null)  
                elseif optReservedFor = "" then
                    checkCltOverlap = checkClientOverlap2(ClientID, classDate, classID, null)  
                end if
            end if
        end if
        if checkCltOverlap OR overLapFound then
                response.redirect ("res_ae.asp?classId="& classID & "&classDate=" & classDate)
        end if
    end if 

    'BJD: 6/12/08 - new membership logic
    if clientID<>"" then
        if relClientID <> "" then
            isMember = checkMembership(relClientID, "")
        else
            isMember = checkMembership(clientID, "")
        end if
                
        if isMember then
            cMembershipID = MemSeriesTypeID
                
            strSQL = "SELECT ServiceDiscountPerc, ProdDiscountPerc, AllowNonMemberPurchases FROM tblSeriesType WHERE SeriesTypeID = " & sqlClean(cMembershipID)
            rsEntry.CursorLocation = 3
            rsEntry.open strSQL, cnWS
            Set rsEntry.ActiveConnection = Nothing
                        
            if NOT rsEntry.EOF then
                    mem_ServiceDiscountPerc = rsEntry("ServiceDiscountPerc")
                    mem_ProdDiscountPerc = rsEntry("ProdDiscountPerc")
                    mem_AllowNonMemberPurchases = rsEntry("AllowNonMemberPurchases")
            end if
            rsEntry.close
        else
            cMembershipID = -1
        end if
    else
        isMember = false
        cMembershipID = -1
    end if

    if IsNum(request.querystring("tg")) then
            typeGroupID = request.querystring("tg")
    elseif IsNum(request.querystring("typeGroup")) then
            typeGroupID = request.querystring("typeGroup")
    elseif isNum(request.form("optTG")) then
            typeGroupID = request.form("optTG")
    else
		' check to see if there is only one available
		' added a check to get the first available IF there is only one (matches optTG logic below)
		if ss_EnrollResReq then
                ''If EnrollReqRes then don't show products of Enroll TG
                strSQL = "SELECT DISTINCT tblTypeGroup.TypeGroup, tblTypeGroup.wsEnrollment, tblTypeGroup.TypeGroupID FROM tblTypeGroup LEFT OUTER JOIN PRODUCTS ON tblTypeGroup.TypeGroupID = PRODUCTS.TypeGroup WHERE "
                strSQL = strSQL & "((PRODUCTS.wsShow = 1) AND (PRODUCTS.Discontinued = 0) AND (NOT (PRODUCTS.TypeGroup IS NULL)) AND (PRODUCTS.[Delete] = 0) AND (PRODUCTS.Class = 1) AND (PRODUCTS.ContractOnly=0) AND (PRODUCTS.Type <> 9)) AND "
                strSQL = strSQL & "  tblTypeGroup.wsEnrollment=0 AND (tblTypeGroup.wsAppointment=1 OR tblTypeGroup.wsReservation=1 OR tblTypeGroup.wsMedia=1) AND tblTypeGroup.Active=1 ORDER BY tblTypeGroup.TypeGroup"
        else
                strSQL = "SELECT DISTINCT tblTypeGroup.TypeGroupID, tblTypeGroup.wsEnrollment, tblTypeGroup.TypeGroup, PRODUCTS.TypeGroup AS PRODTG FROM tblTypeGroup INNER JOIN PRODUCTS ON tblTypeGroup.TypeGroupID = PRODUCTS.TypeGroup "
                strSQL = strSQL & "WHERE (tblTypeGroup.wsEnrollment=1 OR tblTypeGroup.wsAppointment=1 OR tblTypeGroup.wsReservation=1 OR tblTypeGroup.wsResource = 1 OR tblTypeGroup.wsMedia=1 OR tblTypeGroup.wsArrival=1) AND (tblTypeGroup.Active = 1) AND (NOT (PRODUCTS.TypeGroup IS NULL)) AND (PRODUCTS.Discontinued = 0) AND (PRODUCTS.[Delete] = 0) AND (PRODUCTS.Class = 1) AND (PRODUCTS.ContractOnly=0) AND (PRODUCTS.Type <> 9) AND (PRODUCTS.wsShow = 1) "
                strSQL = strSQL & "ORDER BY tblTypeGroup.TypeGroup"
        end if
            
        rsEntry.CursorLocation = 3
        rsEntry.open strSQL, cnWS
        Set rsEntry.ActiveConnection = Nothing
            
        if NOT rsEntry.EOF then
			if rsEntry.RecordCount = 1 then
				TypeGroupID = rsEntry("TypeGroupID")
			else
				TypeGroupID = 0 
			end if
        else
			TypeGroupID = 0 
        end if
        rsEntry.close
    end if
        
    '''Check for enroll case
    if IsNum(TypeGroupID) then
            strSQL = "SELECT wsEnrollment, wsAppointment FROM tblTypeGroup WHERE TypeGroupID=" & sqlClean(TypeGroupID)
            rsEntry.CursorLocation = 3
            rsEntry.open strSQL, cnWS
            Set rsEntry.ActiveConnection = Nothing
            if NOT rsEntry.EOF then
                    isEnrollTG = rsEntry("wsEnrollment")
					isAppointmentTG = rsEntry("wsAppointment")
            end if
            rsEntry.close
    end if

    if IsNum(request.form("requiredoptPurchaseItem")) then
            productID = request.form("requiredoptPurchaseItem")
    elseif IsNum(request.form("frmProdID")) then
            productID = request.form("frmProdID")
    elseif IsNum(request.querystring("prodid")) then        'CB 51_2776 - Contract & Packages Links
            productID = request.querystring("prodid")
    else
            productID = 0
    end if
        
    dim stype
    if request.querystring("stype") <> "" then
        stype = request.QueryString("stype")
    elseif request.Form("stype") <> "" then
        stype = request.Form("stype")
    else
        stype = ""
    end if


        dim productDiscontinued
        productDiscontinued = false
        
        'Make sure that this product is not discontinued or only sold as a contract.  
        'It is possible to have a direct link to this page that will otherwise let you purchace discontinued or contract only products.
        ' Series and Memberships have an stype of 41
        ' Contracts and Packages have an stype of 40 
        if productID <> "" AND productID <> 0 AND productID <> "0" then
            strSQL = ""
            'response.Write "stype: " & stype & "<br/>"
            'response.end
            if stype = "41" then
                ' Series and Memberships
                strSQL =  "Select productID, Discontinued " &_
                            "FROM " &_
                            "PRODUCTS " &_
            	                "INNER JOIN tblTypeGroup ON PRODUCTS.TypeGroup = tblTypeGroup.TypeGroupID " &_
	                            "INNER JOIN tblSeriesType ON PRODUCTS.Type = tblSeriesType.SeriesTypeID " &_
                            "WHERE " &_
                            "PRODUCTS.productID = " & sqlInjectStr(productID) & " " &_
                            "AND (" &_
                                "Discontinued = 1 OR ContractOnly = 1 " &_
                                "OR PRODUCTS.[Delete] = 1 " &_
                                "OR tblTypeGroup.Active = 0 " &_
                                "OR tblSeriesType.Active = 0 " &_
                            ")"
            elseif stype = "43" then
                ' Products
                strSQL = "Select productID, Discontinued from PRODUCTS where productID = " & sqlInjectStr(productID) & " AND (Discontinued = 1 OR ContractOnly = 1)"
            elseif stype = "40" then
                ' Contracts and packages
                strSQL =  "SELECT ContractID, Discontinued from tblContract " &_ 
                            "WHERE (tblContract.ContractID = " & sqlInjectStr(productID) & ") " &_
                            "AND (Discontinued = 1)"    
            end if

            if strSQL <> "" then
                rsEntry.CursorLocation = 3
                rsEntry.open strSQL, cnWS
                Set rsEntry.ActiveConnection = Nothing
                if NOT rsEntry.EOF then
                    productDiscontinued = true
                end if
                rsEntry.close
            end if
        end if

		'Get Per session pricing
		perSessionPriced = 0
		if productID <> "" AND productID <> 0 then
			strSQL = "SELECT PricePerSession FROM PRODUCTS WHERE ProductID = " & sqlClean(productID) 
			rsEntry.CursorLocation = 3
            rsEntry.open strSQL, cnWS
            Set rsEntry.ActiveConnection = Nothing
            if NOT rsEntry.EOF then
				perSessionPriced = rsEntry("PricePerSession")
            end if
            rsEntry.close
		end if

        if IsNum(request.form("requiredoptDeposit")) then
                pmtProductID = request.form("requiredoptDeposit")
        else
                pmtProductID = 0
        end if
        
        if isNum(request.form("frmPartnerID")) then
                partnerID = CINT(request.form("frmPartnerID"))
        else
                partnerID = 0
        end if
        if request.form("requiredtxtCreditAmount")<>"" AND isNum(request.form("requiredtxtCreditAmount")) then
            Call SetLocale(session("mvarLocaleStr"))
				if isNum(request.form("requiredtxtCreditAmount")) then
					frmCredit = CDBL(request.form("requiredtxtCreditAmount"))
				else
					frmCredit = 0
				end if
            Call SetLocale("en-us")
        else
                frmCredit = 0
        end if

        '''''''''''BEGIN Pass Thru Vars for Resv/Appt/Enroll/WaitList''''''''''''
        Dim pVD_Date, pVD_ClassID, pVD_CourseID, pVD_LeadOrFollow, pVD_Rec_EDate, pVD_Rec_DayStr, pVD_TrnID, pVD_RTrnID, pVD_STime, pVD_ETime, pVD_TG, pVD_VT, pVD_Loc, pVD_Notes, pVD_WaitList, pEnrollType, pUsePmtPlan, pRegisterType, pVD_recType, pVD_recNum
        if request.querystring("clsDate")<>"" then
                pVD_Date = CDATE(request.querystring("clsDate"))
        elseif request.form("frmVD_Date")<>"" then
                pVD_Date = CDATE(request.form("frmVD_Date"))
        end if
        if request.form("txtRegStartDate")<>"" then
			Call SetLocale(session("mvarLocaleStr"))
				if isDate(request.form("txtRegStartDate")) then
					pVD_Date = CDATE(request.form("txtRegStartDate"))
				end if
			Call SetLocale("en-us")
        end if
        if isNum(request.querystring("cid")) then
                pVD_ClassID = request.querystring("cid")
        elseif isNum(request.form("frmVD_ClassID")) then
                pVD_ClassID = request.form("frmVD_ClassID")
        end if
        if isNum(request.querystring("courseid")) then
                pVD_CourseID = request.querystring("courseid")
        elseif isNum(request.form("courseid")) then
                pVD_CourseID = request.form("courseid")
        elseif isNum(request.form("frmVD_CourseID")) then
                pVD_CourseID = request.form("frmVD_CourseID")
        end if
        if request.form("optLeadOrFollow")<>"" then
                pVD_LeadOrFollow = request.form("optLeadOrFollow")
        elseif request.form("frmVD_LeadOrFollow")<>"" then
                pVD_LeadOrFollow = request.form("frmVD_LeadOrFollow")
        end if
				''set days of the week
				tmpCount = 1
				if request.Form("frmVD_Rec_DayStr") <> "" then
					pVD_Rec_DayStr = request.Form("frmVD_Rec_DayStr")
				else 
					pVD_Rec_DayStr = ""
					do while tmpCount<=7
						if request.form("optDay"&tmpCount)="on" then
							pVD_Rec_DayStr = pVD_Rec_DayStr & tmpCount & ","
						end if
						tmpCount = tmpCount + 1
					loop
				end if
        if request.form("txtEDate")<>"" then    ''For Recurring Resv/Appts - BUT NOT ENROLL
                if pVD_ClassID="" then
                        Call SetLocale(session("mvarLocaleStr"))
                                if isDate(request.form("txtEDate")) then
                                        pVD_Rec_EDate = CDATE(request.form("txtEDate"))
                                end if
                        Call SetLocale("en-us")
                end if
        elseif request.form("optClassSchEDate")<>"" and pVD_ClassID<>"" then
                pVD_Rec_EDate = CDATE(request.form("optClassSchEDate"))
        elseif request.form("frmVD_Rec_EDate")<>"" then
                pVD_Rec_EDate = CDATE(request.form("frmVD_Rec_EDate"))
                pVD_Rec_DayStr = request.form("frmVD_Rec_DayStr")
        end if
        if isNum(request.form("optInstructor")) OR isNum(request.querystring("trnid")) then
                if isNum(request.form("optInstructor")) then
                        pVD_TrnID = request.form("optInstructor")
                else
                        pVD_TrnID = request.querystring("trnid")
                end if
        else
                pVD_TrnID = request.form("frmVD_TrnID")
        end if
        if isNum(request.querystring("rtrnid")) then
                pVD_RTrnID = request.querystring("rtrnid")
        else
                pVD_RTrnID = request.form("frmVD_RTrnID")
        end if
        if request.form("optStartTime")<>"" then
                pVD_STime = request.form("optStartTime")
                pVD_ETime = request.form("optEndTime")
        elseif request.querystring("stime")<>"" then
                pVD_STime = request.querystring("stime")
                pVD_ETime = request.querystring("etime")
        else
                pVD_STime = request.form("frmVD_STime")
                pVD_ETime = request.form("frmVD_ETime")
        end if
        if isNum(request.querystring("typeGroup")) then
                pVD_TG = request.querystring("typeGroup")
        elseif isNum(request.form("frmVD_TG")) then
                pVD_TG = request.form("frmVD_TG")
        else
                pVD_TG = ""
        end if
        if isNum(request.form("optVisitType")) OR isNum(request.querystring("vt")) then 'single vs rec appt
                if request.form("optVisitType")<>"" then
                        pVD_VT = request.form("optVisitType")
                else
                        pVD_VT = request.querystring("vt")
                end if
        else
                pVD_VT = request.form("frmVD_VT")
        end if
        visitTypeID = pVD_VT
        if isNum(request.form("optLocation")) then
                pVD_Loc = request.form("optLocation")
        elseif isNum(request.querystring("loc")) then
                pVD_Loc = request.querystring("loc")
        elseif isNum(request.form("frmVD_Loc")) then
                pVD_Loc = request.form("frmVD_Loc")
        else
                pVD_Loc = session("curLocation")
        end if
        if request.form("txtNotes")<>"" OR request.form("txtResNotes")<>"" then
                if request.form("txtNotes")<>"" then
                        pVD_Notes = request.form("txtNotes")
                else
                        pVD_Notes = request.form("txtResNotes")
                end if
        else
                pVD_Notes = request.form("frmVD_Notes")
        end if
        if request.querystring("waitlist")<>"" then
                pVD_WaitList = request.querystring("waitlist")
        else
                pVD_WaitList = request.form("frmVD_WaitList")
        end if
        if request.form("EnrollType")<>"" then
                pEnrollType = request.form("EnrollType")
        else 
                pEnrollType = request.form("frmEnrollType")
        end if
        if request.form("frmUsePmtPlan")<>"" then
                pUsePmtPlan = request.form("frmUsePmtPlan")
        end if
        
        if request.form("optRegisterType")<>"" then
                pRegisterType = request.form("optRegisterType")
        elseif request.form("frmRegisterType")<>"" then
                pRegisterType = request.form("frmRegisterType")
        else
                pRegisterType = ""
        end if
        
        if request.form("txtRegStartDate")<>"" then
                Call SetLocale(session("mvarLocaleStr"))
                        if isDate(request.form("txtRegStartDate")) then
                                pRegStartDate = CDATE(request.form("txtRegStartDate"))
                        end if
                Call SetLocale("en-us")
        end if
        
        dim recurringQString
        if request.form("frmVD_recType") <> "" then
            recurringQString = recurringQString & "&recType=" & request.form("frmVD_recType")
        else
            recurringQString = recurringQString & "&recType=" & request.querystring("recType")
        end if
        if request.form("frmVD_recNum") <> "" then
            recurringQString = recurringQString & "&recNum=" & request.form("frmVD_recNum")
        else
            recurringQString = recurringQString & "&recNum=" & request.querystring("recNum")
        end if

        dim enrollDayQString
        enrollDayQString = ""
        if request.form("optEnrollDay1")="on" then
                enrollDayQString = enrollDayQString & "&optEnrollDay1=on"
        end if
        if request.form("optEnrollDay2")="on" then
                enrollDayQString = enrollDayQString & "&optEnrollDay2=on"
        end if
        if request.form("optEnrollDay3")="on" then
                enrollDayQString = enrollDayQString & "&optEnrollDay3=on"
        end if
        if request.form("optEnrollDay4")="on" then
                enrollDayQString = enrollDayQString & "&optEnrollDay4=on"
        end if
        if request.form("optEnrollDay5")="on" then
                enrollDayQString = enrollDayQString & "&optEnrollDay5=on"
        end if
        if request.form("optEnrollDay6")="on" then
                enrollDayQString = enrollDayQString & "&optEnrollDay6=on"
        end if
        if request.form("optEnrollDay7")="on" then
                enrollDayQString = enrollDayQString & "&optEnrollDay7=on"
        end if
        contractStartDate = DateValue((DateAdd("n", Session("tzOffset"),Now)))
        if request.form("requiredtxtStartDate")<>"" then
				'if contract start date is client specified, have the date inputted be converted back to usdate
				if (request.form("notDateSpecificContract") = "True") then
					Call SetLocale(session("mvarLocaleStr"))
				end if
				if isDateValid(request.form("requiredtxtStartDate")) then
					if CDATE(request.form("requiredtxtStartDate")) > DateValue((DateAdd("n", Session("tzOffset"),Now))) then
						contractStartDate = CDATE(request.form("requiredtxtStartDate"))
					end if
				end if
				if (request.form("notDateSpecificContract") = "True") then
					Call SetLocale("en-us")
				end if
        end if
        'BJD: 45_2313 - retain recurring schedule info
        if request.querystring("recType")<>"" then
                pVD_recType = request.querystring("recType")
        elseif request.form("frmVD_recType")<>"" then
                pVD_recType = request.form("frmVD_recType")
        end if
        if request.querystring("recNum")<>"" then
                pVD_recNum = request.querystring("recNum")
        elseif request.form("frmVD_recNum")<>"" then
                pVD_recNum = request.form("frmVD_recNum")
        end if
        

        '''''''''''END Pass Thru Vars for Resv/Appt/Enroll/WaitList''''''''''''

        if pMode=1 then 'SERVICES ONLY
                '''''''''''''''BEGIN UNPAID CHECK''''''''''''''''''
                if request.form("frmUnpaid")<>"" then
                        unPaidCase = request.form("frmUnpaid")
                else
                        unPaidCase = "unknown"
                end if
                
                if unPaidCase="unknown" AND pMode=1 AND clientID<>"" AND TypeGroupID<>"" AND TypeGroupID<>0 then
                    ''''Check for unpaids for this client
                    dim cID
                    cID = clientID
                    if relClientID <> "" then
                            cID = relClientID
                    end if'RI Bug 2963 - adding reconcile of related
                    strSQL = join(array("SELECT MAX(PmtRefNo) AS PmtRefNo, SUM(Remaining) AS Remaining, SUM(RealRemaining) AS RealRemaining, MIN(MinUPVD) AS MinUPVD, MAX(MaxUPVD) AS MaxUPVD ",_
                         "FROM  (SELECT DISTINCT [PAYMENT DATA].PmtRefNo, [PAYMENT DATA].Remaining, [PAYMENT DATA].RealRemaining, UPVisits.MinVDForPD AS MinUPVD, UPVisits.MaxVDForPD AS MaxUPVD, b.CartPmtRefNo ",_
                                 "FROM [PAYMENT DATA] LEFT OUTER JOIN (SELECT UP_PmtRefNo AS CartPmtRefNo FROM tblShoppingCart WHERE SessionID = N'" & sqlClean(getSessionGUID()) & "' AND (SaleID IS NULL)) b ",_
                           "ON [PAYMENT DATA].PmtRefNo = b.CartPmtRefNo ",_
                           "INNER JOIN (SELECT MIN(ClassDate) AS MinVDForPD, MAX(ClassDate) AS MaxVDForPD, PmtRefNo FROM [VISIT DATA] ",_
                                                                  "WHERE (ClientID = ", sqlClean(cID), ") GROUP BY PmtRefNo) AS UPVisits ",_
                                 "ON [PAYMENT DATA].PmtRefNo = UPVisits.PmtRefNo LEFT OUTER JOIN tblTGRelate ON [PAYMENT DATA].TypeGroup = tblTGRelate.TG1 ",_
                                 "WHERE ([PAYMENT DATA].TypeGroup = ", sqlClean(TypeGroupID), ") ",_
                                       "AND ([PAYMENT DATA].ClientID = ", sqlClean(cID), ") ",_
                                       "AND ([PAYMENT DATA].Type = 9) ",_
                                       "AND ([PAYMENT DATA].ExpDate > ", DateSep, sqlClean(DateValue(DateAdd("n", Session("tzOffset"),Now))), DateSep, ") ",_
                                       "AND ([PAYMENT DATA].[Current Series] = 1) ",_
                                       "OR ([PAYMENT DATA].TypeGroup = tblTGRelate.TG1) ",_
                                       "AND ([PAYMENT DATA].ClientID = ", cID, ") ",_
                                       "AND ([PAYMENT DATA].Type = 9) ",_
                                       "AND ([PAYMENT DATA].ExpDate > ", DateSep, sqlClean(DateValue(DateAdd("n", Session("tzOffset"),Now))), DateSep, ") ",_
                                       "AND ([PAYMENT DATA].[Current Series] = 1) ",_
                                       "AND (tblTGRelate.TG2 = ", sqlClean(TypeGroupID), ")) ",_
                                "AS UPData ",_
                         "HAVING (NOT (SUM(Remaining) IS NULL) AND MAX(CartPmtRefNo) IS NULL)"))
                    'response.write strSQL
                    rsEntry.CursorLocation = 3
                    rsEntry.open strSQL, cnWS
                    Set rsEntry.ActiveConnection = Nothing
    
                    if rsEntry.EOF then             
                            unPaidCase = "false"
                    else
                            dim absValRemaining, ElapsedTime, upSDate, upEDate, oldPmtRefNo
                            oldPmtRefNo = rsEntry("PmtRefNo")
                            absValRemaining = abs(rsEntry("Remaining"))
                            unPaidCase = "false"                    
                            if absValRemaining > 0 then
                                    unPaidCase = "true"
                            end if
                            ElapsedTime = DateDiff("d", CDate(rsEntry("MinUPVD")), CDate(rsEntry("MaxUPVD")))
                            upSDate = CDATE(rsEntry("MinUPVD"))
                            if ElapsedTime = 0 then
                                    upEDate = upSDate
                            else
                                    upEDate = CDATE(rsEntry("MaxUPVD"))
                            end if
                    end if
                    rsEntry.close
                end if ' unpaid state is unknown
                
                if unPaidCase="true" then
                        unPaidCaseRec = "true"
                else
                        unPaidCaseRec = "unknown"
                end if
                '''''''''''''''END UNPAID CHECK''''''''''''''''''''''''''''''''''''''''''''
        end if  'SERVICES ONLY
        
        if pMode <= 5 then      ''ADD ITEM STEPS 1-3 || CB: 8_22_07 - added 4 so linking can add products and land on check out page || JG: 11_5_2012 - Added 5 for Account Credits so they can land on check out page.

                '''''''''''' BEGIN ADD TO CART """"""""""""""""""""""
                ' Do not add it to the cart if this is a discontinued product.  The productID still be 0 or not there if it is a contract or maybe other reasons.
                if request.form("frmAddCart")<>"" AND NOT productDiscontinued  then
                        if pMode=0 then 'CONTRACTS & PACKAGES SPECIAL CASE
                        
                                sortOrderID = 1
                                'Get Next SortOrderID
                                strSQL = "SELECT MAX(SortOrderID) AS NextID, SessionID FROM tblShoppingCart GROUP BY SessionID HAVING (SessionID = N'" & sqlClean(getSessionGUID()) & "')"
                                rsEntry.CursorLocation = 3
                                rsEntry.open strSQL, cnWS
                                Set rsEntry.ActiveConnection = Nothing
                                if NOT rsEntry.EOF then
                                        sortOrderID = rsEntry("NextID") + 1
                                end if
                                rsEntry.close
                                groupID = sortOrderID
                        
                                strSQL = "SELECT tblContractItem.ProductID, tblContractItem.Price, tblContractItem.ContractItemID, tblContract.Deposit, IsNull(PRODUCTS.ClientCredit, 0) AS ClientCredit FROM tblContract INNER JOIN tblContractItem ON tblContract.ContractID = tblContractItem.ContractID INNER JOIN PRODUCTS ON tblContractItem.ProductID = PRODUCTS.ProductID WHERE (tblContract.ContractID = " & sqlClean(productID) & ")"
                                rsEntry.CursorLocation = 3
                                rsEntry.open strSQL, cnWS
                                Set rsEntry.ActiveConnection = Nothing
                                if NOT rsEntry.EOF then

                                        if rsEntry("Deposit")>0 then    'Insert Contract Deposit into Cart
                                                strSQL = "INSERT INTO tblShoppingCart (SessionID, SortOrderID, GroupID, Created, ProductID, Quantity, CreditAmount, PartnerID, ContractID, ContractStartDate, RecClientID, reservedFor) VALUES ("
                                                strSQL = strSQL & "N'" & sqlClean(getSessionGUID()) & "'"
                                                strSQL = strSQL & ", " & sqlClean(sortOrderID)
                                                strSQL = strSQL & ", " & sqlClean(groupID)
                                                strSQL = strSQL & ", " & DateSep & sqlClean(DateAdd("n", Session("tzOffset"),Now)) & DateSep
                                                strSQL = strSQL & ", -5" 'Contract Deposit Product
                                                strSQL = strSQL & ", 1"         'Quantity
                                                strSQL = strSQL & ", " & sqlClean(rsEntry("Deposit"))
                                                strSQL = strSQL & ", 0"         'PartnerID
                                                strSQL = strSQL & ", " & sqlClean(productID)      'contractID
                                                strSQL = strSQL & ", " & DateSep & sqlClean(contractStartDate) & DateSep 
                                                if relClientID <> "" then
                                        strSQL = strSQL & ", "& sqlClean(relClientID)
                                    else
                                        strSQL = strSQL & ", null"
                                    end if
                                    'JM-55_2848
                                    if optReservedFor <> "" then
                                        strSQL = strSQL & ", N'"& sqlInjectStr(optReservedFor) &"'"
                                    else
                                        strSQL = strSQL & ", null"
                                    end if
                                                strSQL = strSQL & ")"
                                                cnWS.execute strSQL
        
                                                sortOrderID = sortOrderID + 1
                                        end if

                                        do while NOT rsEntry.EOF                        
                                
                                                strSQL = "INSERT INTO tblShoppingCart (SessionID, SortOrderID, GroupID, Created, ProductID, Quantity, CreditAmount, PartnerID, ContractID, ContractItemID, ContractStartDate, RecClientID, reservedFor) VALUES ("
                                                strSQL = strSQL & "N'" & sqlClean(getSessionGUID()) & "'"
                                                strSQL = strSQL & ", " & sqlClean(sortOrderID)
                                                strSQL = strSQL & ", " & sqlClean(groupID)
                                                strSQL = strSQL & ", " & DateSep & sqlClean(DateAdd("n", Session("tzOffset"),Now)) & DateSep
                                                strSQL = strSQL & ", " & sqlClean(rsEntry("ProductID"))
                                                strSQL = strSQL & ", 1"         'Quantity
                                                strSQL = strSQL & ", " & sqlClean(rsEntry("ClientCredit"))
                                                strSQL = strSQL & ", 0"         'PartnerID
                                                strSQL = strSQL & ", " & sqlClean(productID)      'contractID
                                                strSQL = strSQL & ", " & sqlClean(rsEntry("ContractItemID"))
                                                strSQL = strSQL & ", " & DateSep & sqlClean(contractStartDate) & DateSep 
                                                if relClientID <> "" then
                                        strSQL = strSQL & ", "& sqlClean(relClientID)
                                    else
                                        strSQL = strSQL & ", null"
                                    end if
                                     'JM-55_2848
                                    if optReservedFor <> "" then
                                        strSQL = strSQL & ", N'"& sqlInjectStr(optReservedFor) &"'"
                                    else
                                        strSQL = strSQL & ", null"
                                    end if
                                                strSQL = strSQL & ")"
                                                cnWS.execute strSQL
        
                                                sortOrderID = sortOrderID + 1
                                                rsEntry.MoveNext
                                        loop
                                end if  'EOF
                                rsEntry.close

                                %>
                                <script type="text/javascript">
                                	document.location.replace('main_shop.asp?stype=<%= Server.URLEncode(stype) %>&pMode=4<%=enrollDayQString%>&reSchedule=<%=Server.URLEncode(reSchedule)%>&origId=<%=Server.URLEncode(origId)%><%=recurringQString%>');
                                </script>
                                <%              
                        
                        else
                                'BJD: 3/24/08 - do not add membership series/retail unless user is logged in and a member
                                dim mem_MemberHasAccess, prod_MembersOnly
                                mem_MemberHasAccess = false
                                prod_MembersOnly = false
                                
                                strSQL = "SELECT COUNT(*) AS NumRestrictions FROM tblProductSeriesTypeSetting WHERE (Setting = 1) AND (ProductID = " & sqlClean(productID) & ") "
        
                                rsEntry.open strSQL, cnWS
                                Set rsEntry.ActiveConnection = Nothing
                                if NOT rsEntry.EOF then
                                        if rsEntry("NumRestrictions")>0 then
                                                prod_MembersOnly = true
                                        end if
                                end if
                                rsEntry.close
                                
                                if isMember AND prod_MembersOnly then
                                        strSQL = "SELECT ProductID FROM tblProductSeriesTypeSetting WHERE (ProductID = " & sqlClean(productID) & ") AND (SeriesTypeID = " & sqlClean(cMembershipID) & ") AND (Setting = 1) "
                                        rsEntry.CursorLocation = 3
                                        rsEntry.open strSQL, cnWS
                                        Set rsEntry.ActiveConnection = Nothing
                                        
                                        if NOT rsEntry.EOF then
                                                mem_MemberHasAccess = true
                                        end if
                                        rsEntry.close
                                end if
                                
                                'BJD: 5/26/08 - do not add intro series if its in the cart already
                                dim introInCart
                                introInCart = false
                                strSQL = "SELECT tblShoppingCart.ProductID FROM tblShoppingCart INNER JOIN PRODUCTS ON tblShoppingCart.ProductID = PRODUCTS.ProductID WHERE (SessionID = N'" & sqlClean(getSessionGUID()) & "') AND (tblShoppingCart.ProductID = " & sqlClean(productID) & ") AND (PRODUCTS.Introductory = 1) AND (SaleID IS NULL)  "
                                rsEntry.CursorLocation = 3
                                rsEntry.open strSQL, cnWS
                                Set rsEntry.ActiveConnection = Nothing
                                if NOT rsEntry.EOF then
                                        introInCart = true
                                        %><script type="text/javascript" >alert("<%=DisplayPhraseJS(phraseDictionary,"Dupintro")%>")</script> <%
                                end if
                                rsEntry.close
                                
                                
								if session("mVarUserID")<>"" then
									checkPrevPurch(DisplayPhrase(phraseDictionary,"Previntropurchase"))
								end if
                                
                                
                                'IF (member has access through membership OR product has no members only restrictions) AND intro series is not in the cart
                                if (mem_MemberHasAccess OR NOT prod_MembersOnly) AND NOT introInCart then
                        
                                        sortOrderID = 1
                                        'Get Next SortOrderID
                                        strSQL = "SELECT MAX(SortOrderID) AS NextID, SessionID FROM tblShoppingCart GROUP BY SessionID HAVING (SessionID = N'" & sqlClean(getSessionGUID()) & "')"
                                        rsEntry.CursorLocation = 3
                                        rsEntry.open strSQL, cnWS
                                        Set rsEntry.ActiveConnection = Nothing
                                        if NOT rsEntry.EOF then
                                                sortOrderID = rsEntry("NextID") + 1
                                        end if
                                        rsEntry.close
        
                                        Dim removeAddAction : removeAddAction = false
                                        if pVD_Date<>"" AND pVD_TG<>"" then     'Has AddAction
                                                strSQL = "SELECT ProductID, TypeGroup FROM PRODUCTS WHERE (ProductID = " & sqlClean(productID) & ") AND (TypeGroup = " & sqlClean(pVD_TG) & ") OR ((ProductID = " & sqlClean(productID) & ") AND (TypeGroup IN (SELECT TG2 FROM tblTGRelate WHERE (TG1 = " & sqlClean(pVD_TG) & "))))"
                                                rsEntry.CursorLocation = 3
                                                rsEntry.open strSQL, cnWS
                                                Set rsEntry.ActiveConnection = Nothing
                                                if rsEntry.EOF then     'AddAction Can Not Be Paid For With This Product
                                                        removeAddAction = true
                                                end if
                                                rsEntry.close                                                                           
                                        end if
                                        
                                        '******************************************** ADD Studio Product OR MBPS *******************************************
                                if pUsePmtPlan AND productID<>pmtProductID then 
                                            strSQL = "INSERT INTO tblShoppingCart (SessionID, SortOrderID, GroupID, Created, ProductID, Quantity, CreditAmount, VD_Date, VD_ClassID, VD_CourseID, VD_Rec_EDate, VD_Rec_DayStr, VD_TrnID, VD_RTrnID, VD_STime, VD_ETime, VD_TG, VD_VT, VD_Loc, VD_Notes, VD_WaitList, VD_EnrollType, GC_ToName, GC_ToEmail, GC_FromName, GC_Msg, GC_ToFullName, GC_Title, GC_DeliveryDate, GCLayoutID, UP_SDate, UP_PmtRefNo, VD_Leader, PartnerID, DayMon, DayTue, DayWed, DayThu, DayFri, DaySat, DaySun, EnrollRegisterType, EnrollRegStartDate, NumSessions, RecClientID, reservedFor "
                                            strSQL = strSQL & ") VALUES (N'" & sqlClean(getSessionGUID()) & "'"
                                            strSQL = strSQL & ", " & sqlClean(sortOrderID)
                                            strSQL = strSQL & ", " & sqlClean(sortOrderID)
                                            strSQL = strSQL & ", " & DateSep & sqlClean(DateAdd("n", Session("tzOffset"),Now)) & DateSep
                                            strSQL = strSQL & ", " & sqlClean(pmtProductID)
                                            strSQL = strSQL & ", 1" 'Quantity
                                            strSQL = strSQL & ", " & sqlClean(frmCredit)
                                            strSQL = strSQL & ", null"
                                            if pVD_ClassID<>"" AND NOT removeAddAction then
                                                    strSQL = strSQL & ", " & sqlClean(pVD_ClassID)
                                            else
                                                    strSQL = strSQL & ", null"
                                            end if
                                            strSQL = strSQL & ", null, null, null, null, null, null, null, null, null, null, null"
                                            strSQL = strSQL & ", 0, 0, null, null, null, null, null, null, null, null, null, null, 0, " & sqlClean(partnerID)
                                            strSQL = strSQL & ", 0, 0, 0, 0, 0, 0, 0, NULL, NULL, 1"    
                                            if relClientID <> "" then
                                        strSQL = strSQL & ", "& sqlClean(relClientID)
                                    else
                                        strSQL = strSQL & ", null"
                                    end if
                                     'JM-55_2848
                                    if optReservedFor <> "" then
                                        strSQL = strSQL & ", N'"& sqlInjectStr(optReservedFor) &"'"
                                    else
                                        strSQL = strSQL & ", null"
                                    end if
                                                strSQL = strSQL & ")"
                                            'response.write debugSQL(strSQL, "SQL")
                                            'response.end
                                            cnWS.execute strSQL
                                            
                                            frmCredit = "0" ' blank out credit for the 'real' product
                                            depositSortOrderID = sortOrderID
                                            sortOrderID = sortOrderID + 1
                                    end if 

                                        strSQL = "INSERT INTO tblShoppingCart (SessionID, SortOrderID, GroupID, Created, ProductID, Quantity, CreditAmount, VD_Date, VD_ClassID, VD_CourseID, VD_Rec_EDate, VD_Rec_DayStr, VD_TrnID, VD_RTrnID, VD_STime, VD_ETime, VD_TG, VD_VT, VD_Loc, VD_Notes, VD_WaitList, VD_EnrollType, GC_ToName, GC_ToEmail, GC_FromName, GC_Msg, GC_ToFullName, GC_Title, GC_DeliveryDate, GCLayoutID, UP_SDate, UP_PmtRefNo, VD_Leader, PartnerID, DayMon, DayTue, DayWed, DayThu, DayFri, DaySat, DaySun, EnrollRegisterType, EnrollRegStartDate, NumSessions "
                                        if pUsePmtPlan AND productID<>pmtProductID then 
                                            strSQL = strSQL & ", DepositSortOrderID "
                                        end if 
                                        strSQL = strSQL & ", RecClientID, reservedFor) VALUES (N'" & sqlClean(getSessionGUID()) & "'"
                                        strSQL = strSQL & ", " & sqlClean(sortOrderID)
                                        strSQL = strSQL & ", " & sqlClean(sortOrderID)
                                        strSQL = strSQL & ", " & DateSep & sqlClean(DateAdd("n", Session("tzOffset"),Now)) & DateSep
                                        strSQL = strSQL & ", " & sqlClean(productID)
                                        strSQL = strSQL & ", 1"         'Quantity
                                        strSQL = strSQL & ", " & sqlClean(frmCredit)
                                        if pVD_Date<>"" AND NOT removeAddAction and  isDateValid(pVD_Date) then
                                                strSQL = strSQL & ", " & DateSep & sqlClean(pVD_Date) & DateSep
                                        else
                                                strSQL = strSQL & ", null"
                                        end if
                                        if pVD_ClassID<>"" AND NOT removeAddAction and isNum(pVD_ClassID) then
                                                strSQL = strSQL & ", " & sqlClean(pVD_ClassID)
                                        else
                                                strSQL = strSQL & ", null"
                                        end if
                                        if pVD_CourseID<>"" AND NOT removeAddAction and isNum(pVD_CourseID) then
                                                strSQL = strSQL & ", " & sqlClean(pVD_CourseID)
                                        else
                                                strSQL = strSQL & ", null"
                                        end if
                                        if pVD_Rec_EDate<>"" AND NOT removeAddAction and isDateValid(pVD_Rec_EDate) then
                                                strSQL = strSQL & ", " & DateSep & sqlClean(pVD_Rec_EDate) & DateSep
                                        else
                                                strSQL = strSQL & ", null"
                                        end if
                                        if pVD_Rec_DayStr<>"" AND NOT removeAddAction then
                                                strSQL = strSQL & ", N'" & sqlInjectStr(pVD_Rec_DayStr) & "'"
                                        else
                                                strSQL = strSQL & ", null"
                                        end if
                                        if pVD_TrnID<>"" AND NOT removeAddAction and isNum(pVD_TrnID) then
                                                strSQL = strSQL & ", " & sqlClean(pVD_TrnID)
                                        else
                                                strSQL = strSQL & ", null"
                                        end if
                                        if pVD_RTrnID<>"" AND NOT removeAddAction and isNum(pVD_RTrnID) then
                                                strSQL = strSQL & ", " & sqlClean(pVD_RTrnID)
                                        else
                                                strSQL = strSQL & ", null"
                                        end if
                                        if pVD_STime<>"" AND NOT removeAddAction and isTime(pVD_STime) then
                                                strSQL = strSQL & ", " & TimeSepB & sqlClean(pVD_STime) & TimeSepA
                                        else
                                                strSQL = strSQL & ", null"
                                        end if
                                        if pVD_ETime<>"" AND NOT removeAddAction and isTime(pVD_ETime) then
                                                strSQL = strSQL & ", " & TimeSepB & sqlClean(pVD_ETime) & TimeSepA
                                        else
                                                strSQL = strSQL & ", null"
                                        end if
                                        if pVD_TG<>"" AND NOT removeAddAction and isNum(pVD_TG) then
                                                strSQL = strSQL & ", " & sqlClean(pVD_TG)
                                        else
                                                strSQL = strSQL & ", null"
                                        end if
                                        if pVD_VT<>"" AND NOT removeAddAction and isNum(pVD_VT) then
                                                strSQL = strSQL & ", " & sqlClean(pVD_VT)
                                        else
                                                strSQL = strSQL & ", null"
                                        end if
                                        if pVD_Loc<>"" AND NOT removeAddAction and isNum(pVD_Loc) then
                                                strSQL = strSQL & ", " & sqlClean(pVD_Loc)
                                        else
                                                strSQL = strSQL & ", null"
                                        end if
                                        if pVD_Notes<>"" AND NOT removeAddAction then
                                                strSQL = strSQL & ", N'" & sqlInjectStr(pVD_Notes) & "'"
                                        else
                                                strSQL = strSQL & ", null"
                                        end if
                                        if pVD_WaitList<>"" AND NOT removeAddAction then
                                                strSQL = strSQL & ", 1"
                                        else
                                                strSQL = strSQL & ", 0"
                                        end if
                                        if pEnrollType<>"" AND NOT removeAddAction then
                                                strSQL = strSQL & ", 1"
                                        else
                                                strSQL = strSQL & ", 0"
                                        end if
                                        if request.form("requiredtxtGiftCardTo")<>"" then
                                                strSQL = strSQL & ", N'" & sqlInjectStr(request.form("requiredtxtGiftCardTo")) & "'"
                                        else
                                                strSQL = strSQL & ", null"
                                        end if
										'need to use changed(required) field names, see lines 1829-1844
                                        if request.form("requiredEm_RecipientEmail1")<>"" then
                                                strSQL = strSQL & ", N'" & sqlInjectStr(request.form("requiredEm_RecipientEmail1")) & "'"
                                        else
                                                strSQL = strSQL & ", null"
                                        end if
                                        if request.form("requiredtxtGiftCardFrom")<>"" then
                                                strSQL = strSQL & ", N'" & sqlInjectStr(request.form("requiredtxtGiftCardFrom")) & "'"
                                        else
                                                strSQL = strSQL & ", null"
                                        end if
                                        if request.form("egiftCardMessage")<>"" then
                                                strSQL = strSQL & ", N'" & sqlInjectStr(request.form("egiftCardMessage")) & "'"
                                        else
                                                strSQL = strSQL & ", null"
                                        end if
                                        if request.form("requiredtxtGiftCardToFullName")<>"" then
                                                strSQL = strSQL & ", N'" & sqlInjectStr(request.form("requiredtxtGiftCardToFullName")) & "'"
                                        else
                                                strSQL = strSQL & ", null"
                                        end if
                                        if request.form("requiredtxtGiftCardTitle")<>"" then
                                                strSQL = strSQL & ", N'" & sqlInjectStr(request.form("requiredtxtGiftCardTitle")) & "'"
                                        else
                                                strSQL = strSQL & ", null"
                                        end if
                                        
										'DeliveryDate
										deliveryDate = ReadNullableDateFromLocalForm(request.form("requiredtxtDeliveryDate"))
										if request.form("optSendGiftCardByEmail")= "on" and not isNull(deliveryDate) then
												strSQL = strSQL & ", " & DateSep & sqlClean(deliveryDate) & DateSep
                                        else
                                                strSQL = strSQL & ", null"
                                        end if
                                        
                                        if request.form("selectedImageId")<>"" AND isNum(request.form("selectedImageId")) then
                                                strSQL = strSQL & ", " & sqlClean(request.form("selectedImageId"))
                                        else
                                                strSQL = strSQL & ", null"
                                        end if
                                        if request.form("frmUPSDate")<>"" and isDateValid(request.form("frmUPSDate")) then
                                                strSQL = strSQL & ", " & DateSep & sqlClean(request.form("frmUPSDate")) & DateSep
                                        else
                                                strSQL = strSQL & ", null"
                                        end if
										
										' isNum() for SQL injection protection. This column type is a nullable bigInt
                                        if request.form("frmUPOldPmtRefNo")<>"" AND isNum(request.form("frmUPOldPmtRefNo")) then
                                                strSQL = strSQL & ", " & sqlClean(request.form("frmUPOldPmtRefNo"))
                                        else
                                                strSQL = strSQL & ", null"
                                        end if
                                        if request.form("frmVD_LeadOrFollow")<>"" AND NOT removeAddAction then
                                                if request.form("frmVD_LeadOrFollow")="follow" then
                                                        strSQL = strSQL & ", 0"
                                                else
                                                        strSQL = strSQL & ", 1"
                                                end if
                                        else
                                                strSQL = strSQL & ", 0"
                                        end if
                                        'if request.form("frmPartnerID")<>"" then
                                        '       strSQL = strSQL & ", " & request.form("frmPartnerID")
                                        'else
                                        '       strSQL = strSQL & ", null"
                                        'end if
                                        strSQL = strSQL & ", " & partnerID
                                        if request.form("optEnrollDay1")="on" AND NOT removeAddAction then
                                                strSQL = strSQL & ", 1"
                                        else
                                                strSQL = strSQL & ", 0"
                                        end if
                                        if request.form("optEnrollDay2")="on" AND NOT removeAddAction then
                                                strSQL = strSQL & ", 1"
                                        else
                                                strSQL = strSQL & ", 0"
                                        end if
                                        if request.form("optEnrollDay3")="on" AND NOT removeAddAction then
                                                strSQL = strSQL & ", 1"
                                        else
                                                strSQL = strSQL & ", 0"
                                        end if
                                        if request.form("optEnrollDay4")="on" AND NOT removeAddAction then
                                                strSQL = strSQL & ", 1"
                                        else
                                                strSQL = strSQL & ", 0"
                                        end if
                                        if request.form("optEnrollDay5")="on" AND NOT removeAddAction then
                                                strSQL = strSQL & ", 1"
                                        else
                                                strSQL = strSQL & ", 0"
                                        end if
                                        if request.form("optEnrollDay6")="on" AND NOT removeAddAction then
                                                strSQL = strSQL & ", 1"
                                        else
                                                strSQL = strSQL & ", 0"
                                        end if
                                        if request.form("optEnrollDay7")="on" AND NOT removeAddAction then
                                                strSQL = strSQL & ", 1"
                                        else
                                                strSQL = strSQL & ", 0"
                                        end if
                                        if request.form("frmRegisterType")<>"" AND NOT removeAddAction then
                                                strSQL = strSQL & ", N'" & sqlInjectStr(request.form("frmRegisterType")) & "'"
                                        else
                                                strSQL = strSQL & ", NULL"
                                        end if
                                        if pRegStartDate<>"" AND NOT removeAddAction and isDateValid(pRegStartDate) then
                                                strSQL = strSQL & ", " & DateSep & sqlClean(pRegStartDate) & DateSep
                                        else
                                                strSQL = strSQL & ", NULL"
                                        end if
                                        if request.form("optNumSessions")<>"" and isNum(request.form("optNumSessions")) then
                                            strSQL = strSQL & ", " & sqlClean(request.form("optNumSessions"))
                                        elseif isNum(request.Form("txtServiceCount")) and request.Form("txtServiceCount") > "1" then ' count
											strSQL = strSQL & ", " & sqlClean(request.Form("txtServiceCount"))
										else
											strSQL = strSQL & ", 1"
                                        end if

                                        if pUsePmtPlan AND productID <> pmtProductID then 
                                            strSQL = strSQL & ", " & sqlClean(depositSortOrderID)
                                        end if 
                                        if relClientID <> "" then
                                strSQL = strSQL & ", "& sqlClean(relClientID)
                            else
                                strSQL = strSQL & ", null"
                            end if
                     'JM-55_2848
                            if optReservedFor <> "" then
                                strSQL = strSQL & ", N'"& sqlInjectStr(optReservedFor) &"'"
                            else
                                strSQL = strSQL & ", null"
                            end if
                                        strSQL = strSQL & ")"   
                                        'response.write debugSQL(strSQL, "SQL")
                                       ' RW strSQL
										'response.end
                                        cnWS.execute strSQL
                                        
                                        'BJD: 45_2313 - update recurring reservations for new scheduling logic
                                        ' add open enrollment stuff 
                                        'if request.form("frmRegisterType")="openSchedule" then

                                        'CB 12/31/2008 - added InStr ' BQL 1/2/9 removed, upstream logic updated
                                        if request.form("frmRegisterType") = "openSchedule" then
                                                if pVD_CourseID<>"" then
                                                        
                                                        strSQL = " SELECT tblClasses.ClassID, tblClasses.DaySunday as Day1, tblClasses.DayMonday as Day2, tblClasses.DayTuesday as Day3, tblClasses.DayWednesday as Day4, "
                                                        strSQL = strSQL & " tblClasses.DayThursday as Day5, tblClasses.DayFriday as Day6, tblClasses.DaySaturday as Day7, tblClassDescriptions.ClassName, tblClasses.ClassDateStart, "
                                                        strSQL = strSQL & " tblClasses.ClassDateEnd FROM tblClasses INNER JOIN "
                                                        strSQL = strSQL & " tblClassDescriptions ON tblClasses.DescriptionID = tblClassDescriptions.ClassDescriptionID "
                                                        strSQL = strSQL & " WHERE tblClasses.CourseID = " & sqlClean(pVD_CourseID)
                                                        strSQL = strSQL & " ORDER BY ClassDateStart "
                                                        'response.write debugSQL(strSQL, "SQL")
                                                        rsEntry.CursorLocation = 3
                                                        rsEntry.open strSQL, cnWS
                                                        Set rsEntry.ActiveConnection = Nothing
                                                        if NOT rsEntry.EOF then 
                                                                do while NOT rsEntry.EOF
                
                                                                        frmRtnClassDate = pVD_Date
                                                                        if request.form("optClassSchEDate")<>"" then
                                                                                frmRtnClassEDate = CDATE(rsEntry("ClassDateEnd"))
                                                                        else
                                                                                frmRtnClassEDate = CDATE(frmRtnClassDate)
                                                                        end if
                                                                        
                                                                        frmRtncont = true
                                                                        frmRtntmpDate = frmRtnClassDate
                                                                        if CDATE(frmRtntmpDate) < CDATE(rsEntry("ClassDateStart")) then
                                                                                frmRtntmpDate = rsEntry("ClassDateStart")
                                                                        end if
                        
                                                                        do while frmRtncont
                                                                                if request.form("optDay" & FmtDateShort(frmRtntmpDate) & "-" & rsEntry("ClassID"))="on" then 
                                                                                        strSQL = "INSERT INTO tblShoppingCartEnrollDate (SessionID, SortOrderID, EnrollDate, ClassID) VALUES ("
                                                                                        strSQL = strSQL & "N'" & getSessionGUID() & "'"
                                                                                        strSQL = strSQL & ", " & sortOrderID
                                                                                        strSQL = strSQL & ", " & DateSep & frmRtntmpDate & DateSep
                                                                                        strSQL = strSQL & ", " & rsEntry("ClassID")
                                                                                        strSQL = strSQL & ")"
                                                                                        cnWS.execute strSQL
                                                                                end if
                                                                                frmRtntmpDate = CDATE(DATEADD("d", 1, frmRtntmpDate))
                                                
                                                                                if frmRtntmpDate > frmRtnClassEDate then
                                                                                        frmRtncont = false
                                                                                end if
                                                                        loop
                                                                        rsEntry.MoveNext
                                                                loop
                                                        end if
                                                        rsEntry.close


                                                else            
                                                        frmRtnClassDate = pVD_Date
                                                        if request.form("optClassSchEDate")<>"" then
                                                                frmRtnClassEDate = CDATE(request.form("optClassSchEDate"))
                                                        elseif request.form("frmVD_Rec_EDate")<>"" then
                                                                frmRtnClassEDate = CDATE(request.form("frmVD_Rec_EDate"))
                                                        else
                                                                frmRtnClassEDate = CDATE(frmRtnClassDate)
                                                        end if
                                                        
                                                        frmRtncont = true
                                                        frmRtntmpDate = frmRtnClassDate
                                                        
                                                        'response.write " Start: " & frmRtntmpDate & " End: " & frmRtnClassEDate 
                                                        'response.end
                                                        do while frmRtncont
                                                                if request.form("optDay" & FmtDateShort(frmRtntmpDate) & "-" & pVD_ClassID)="on" then
                                                                        strSQL = "INSERT INTO tblShoppingCartEnrollDate (SessionID, SortOrderID, EnrollDate, ClassID) VALUES ("
                                                                        strSQL = strSQL & "N'" &  sqlClean(getSessionGUID()) & "'"
                                                                        strSQL = strSQL & ", " & sqlClean(sortOrderID)
                                                                        strSQL = strSQL & ", " & DateSep & sqlClean(frmRtntmpDate) & DateSep
                                                                        strSQL = strSQL & ", " & sqlClean(pVD_ClassID)
                                                                        strSQL = strSQL & ")"
                                                                        cnWS.execute strSQL
                                                                end if
                                                                frmRtntmpDate = CDATE(DATEADD("d", 1, frmRtntmpDate))
                                                
                                                                if frmRtntmpDate > frmRtnClassEDate then
                                                                        frmRtncont = false
                                                                end if
                                                        loop
                                                end if
                                        elseif pVD_TrnID<>"" then ' appt
                                                curDate = CDATE(pVD_Date)
                                                
                                                if pVD_recType<>"" AND pVD_recNum<>"" then 'recurring appt
                                                        dayCount = 1
                                                        Do While dayCount <= 7 
                                                                if inStr(pVD_Rec_DayStr, WeekDay(curDate))>0 then
                                                                        Do While CDATE(curDate) <= CDATE(pVD_Rec_EDate)
                                                                                strSQL = "INSERT INTO tblShoppingCartEnrollDate (SessionID, SortOrderID, EnrollDate, ClassID) VALUES ("
                                                                                strSQL = strSQL & "N'" &  sqlClean(getSessionGUID()) & "'"
                                                                                strSQL = strSQL & ", " & sqlClean(sortOrderID)
                                                                                strSQL = strSQL & ", " & DateSep & sqlClean(curDate) & DateSep
                                                                                strSQL = strSQL & ", 0)"
                                                                                cnWS.execute strSQL
                                                                                
                                                                                if pVD_recType=1 then 'weeks
                                                                                        curDate = CDATE(DateAdd("ww", pVD_recNum, curDate))
                                                                                else 'months
                                                                                        curDate = CDATE(DateAdd("ww", pVD_recNum * 4, curDate))
                                                                                end if
                                                                        loop
                                                                end if
                                                                dayCount = dayCount + 1
                                                        loop
                                                else 'single appt
                                                        strSQL = "INSERT INTO tblShoppingCartEnrollDate (SessionID, SortOrderID, EnrollDate, ClassID) VALUES ("
                                                        strSQL = strSQL & "N'" &  sqlClean(getSessionGUID()) & "'"
                                                        strSQL = strSQL & ", " & sqlClean(sortOrderID)
                                                        strSQL = strSQL & ", " & DateSep & sqlClean(curDate) & DateSep
                                                        strSQL = strSQL & ", 0)"
                                                        cnWS.execute strSQL
                                                end if
                                        end if
                                        
                                        if request.form("frmAddCart")="1" then
                                                if pMode=3 then
                                                        'no op - keep same search
                                                else
                        %>
                        <script type="text/javascript">
                            document.location.replace('main_shop.asp?stype=<%= Server.URLEncode(stype) %>&pMode=<%=Server.URLEncode(pMode)%><%=enrollDayQString%>&reSchedule=<%=Server.URLEncode(reSchedule)%>&origId=<%=Server.URLEncode(origId)%><%=recurringQString%>');
                        </script>
                        <%              
                                                end if
                                        elseif request.form("frmAddCart")="2" then
                                        
                        %>
                        <script type="text/javascript">
                            document.location.replace('main_shop.asp?stype=<%= Server.URLEncode(stype) %>&pMode=4<%=enrollDayQString%>&reSchedule=<%=Server.URLEncode(reSchedule)%>&origId=<%=Server.URLEncode(origId)%><%=recurringQString%>');
                        </script>
                        <%              
                                        end if
                                end if
                        end if  'contracts & packages
                end if  'request.form("frmAddCart")<>""
                '''''''''''' END ADD TO CART """"""""""""""""""""""
        end if          ''ADD ITEM STEPS 1-3

        if pMode=3 then 'Retail Add
                Dim lastSearch
                if request.form("txtSearchStr")<>"" then
                        lastSearch      = request.form("txtSearchStr")
                else
                        lastSearch = request.form("lastSearch")
                end if
        end if

        if pMode=4 OR pMode=0 then
                '''''''''''' SHOPPING CART """"""""""""""""""""""
                Dim tmpCount2, rowColor, nTax1, nTax2, nTax3, nTax4, nTax5, tmpCalcTaxRate, tmpPrice, tmpDisc, tmpDiscP, grandTot, studioTotal
                'Dim tmpTaxAmt, studioTotTax, studioTotal, studioTotalUntaxed ' removed
                nTax1 = 0
                nTax2 = 0
                nTax3 = 0
                nTax4 = 0
                nTax5 = 0
                grandTot = 0
                strSQL = "SELECT CAST(Tax1 AS DECIMAL(20,8)) AS Tax1, CAST(Tax2 AS DECIMAL(20,8)) AS Tax2, CAST(Tax3 AS DECIMAL(20,8)) AS Tax3, CAST(Tax4 AS DECIMAL(20,8)) AS Tax4, CAST(Tax5 AS DECIMAL(20,8)) AS Tax5 "&_
						" FROM Location "&_
						" WHERE (LocationID = 98) "
                rsEntry.CursorLocation = 3
                rsEntry.open strSQL, cnWS
                Set rsEntry.ActiveConnection = Nothing
                if NOT rsEntry.EOF then 
                        if NOT isNULL(rsEntry("Tax1")) then
                                nTax1 = CDbl(rsEntry("Tax1"))
                        end if
                        if NOT isNULL(rsEntry("Tax2")) then
                                nTax2 = CDbl(rsEntry("Tax2"))
                        end if
                        if NOT isNULL(rsEntry("Tax3")) then
                                nTax3 = CDbl(rsEntry("Tax3"))
                        end if
                        if NOT isNULL(rsEntry("Tax4")) then
                                nTax4 = CDbl(rsEntry("Tax4"))
                        end if
                        if NOT isNULL(rsEntry("Tax5")) then
                                nTax5 = CDbl(rsEntry("Tax5"))
                        end if
                end if
                rsEntry.close
    end if  'if pMode=4 OR pMode=0 then
                
    if pMode=4 then   
        ''Update Cart - STUDIO PRODUCTS
        strSQL = "SELECT SortOrderID, Quantity, PricePerSession, ISNULL(tblShoppingCart.NumSessions, 1) as NumSessions, "
         if UsePerStaffPricing then
            strSQL = strSQL & "ISNULL(tblContractItem.Price, CASE WHEN tblTrainerVisitRates.Price IS NOT NULL AND PRODUCTS.[COUNT] = 1 THEN tblTrainerVisitRates.Price ELSE Products.OnlinePrice END) AS OnlinePrice, "
         else
	        strSQL = strSQL & "ISNULL(tblContractItem.Price, PRODUCTS.OnlinePrice) AS OnlinePrice, "
         end if
    
        strSQL = strSQL & "CreditAmount, PromotionID, tblShoppingCart.ContractID "&_
                 "FROM tblShoppingCart "&_
                 "INNER JOIN Products ON tblShoppingCart.ProductID = Products.ProductID "&_
                 "LEFT OUTER JOIN tblContract ON tblContract.ContractID = tblShoppingCart.ContractID "&_
                 "LEFT OUTER JOIN tblContractItem ON tblContractItem.ContractItemID = tblShoppingCart.ContractItemID "
        if UsePerStaffPricing then
            strSQL = strSQL & "LEFT OUTER JOIN tblTrainerVisitRates ON tblShoppingCart.VD_VT = tblTrainerVisitRates.VisitTypeID AND tblShoppingCart.VD_TrnID = tblTrainerVisitRates.TrainerID AND tblTrainerVisitRates.Active = 1 " 
        end if
       strSQL = strSQL & "WHERE (SessionID = N'" & sqlClean(getSessionGUID()) & "') AND (SaleID IS NULL) AND PartnerID=0 "&_
                 "ORDER BY SortOrderID"
        rsEntry.CursorLocation = 3
        rsEntry.open strSQL, cnWS
        Set rsEntry.ActiveConnection = Nothing

        do while NOT rsEntry.EOF
            if request.form("optDelete"&rsEntry("SortOrderID"))="on" then

                strSQL = "DELETE FROM tblShoppingCartEnrollDate FROM tblShoppingCartEnrollDate INNER JOIN tblShoppingCart on tblShoppingCart.SessionID = tblShoppingCartEnrollDate.SessionID "&_
                         "AND tblShoppingCart.SortOrderID = tblShoppingCartEnrollDate.SortOrderID "&_
                         "WHERE (tblShoppingCartEnrollDate.SessionID = N'" & sqlClean(getSessionGUID()) & "') AND tblShoppingCart.SortOrderID=" & sqlClean(rsEntry("SortOrderID")) &_
                         " OR tblShoppingCart.DepositSortOrderID = " & sqlClean(rsEntry("SortOrderID"))
                cnWS.execute strSQL

                'CB 45_1276
                if isNULL(rsEntry("ContractID")) then   'Individual Item
                        strSQL = "DELETE FROM tblShoppingCart WHERE (SessionID = N'" & sqlClean(getSessionGUID()) & "') AND SortOrderID=" & sqlClean(rsEntry("SortOrderID")) & " OR DepositSortOrderID = " & sqlClean(rsEntry("SortOrderID"))
                        cnWS.execute strSQL
                else    'Contract or Package - Delete All Items in Contract/Package
                        strSQL = "DELETE FROM tblShoppingCart WHERE (SessionID = N'" & sqlClean(getSessionGUID()) & "') AND ContractID=" & sqlClean(rsEntry("ContractID"))
                        cnWS.execute strSQL
                end if
                                
                numItemsUpdated = true
            elseif ((request.form("txtQty"&rsEntry("SortOrderID"))<>CSTR(rsEntry("NumSessions")) AND rsEntry("PricePerSession")) OR (NOT rsEntry("PricePerSession") AND request.form("txtQty"&rsEntry("SortOrderID"))<>CSTR(rsEntry("Quantity")))) AND request.form("txtQty"&rsEntry("SortOrderID"))<>"" then
				quantity = request.form("txtQty"&rsEntry("SortOrderID"))
                
				if NOT ss_AllowOrderOutofStock then
					strSQL = "SELECT (tblInvOnHand.UnitsLogged - tblInvOnHand.UnitsSold - tblInvOnHand.UnitsOnOrder) AS Available "&_
                             "FROM tblInvOnHand "&_
                             "INNER JOIN PRODUCTS ON PRODUCTS.ProductID = tblInvOnHand.ProductID "&_
                             "INNER JOIN tblShoppingCart ON tblShoppingCart.ProductID = tblInvOnHand.ProductID "&_
                             "WHERE PRODUCTS.ItemTypeID = 2 AND (tblShoppingCart.SessionID = N'" & sqlClean(getSessionGUID()) & "') AND tblShoppingCart.SortOrderID=" & sqlClean(rsEntry("SortOrderID")) &_
                             " AND LocationID = CASE WHEN PRODUCTS.OnlineStoreInvLoc IS NULL THEN (SELECT OnlineStoreLoc FROM tblGenOpts) ELSE PRODUCTS.OnlineStoreInvLoc END"
					rsEntry2.CursorLocation = 3
					rsEntry2.open strSQL, cnWS
					Set rsEntry2.ActiveConnection = Nothing
					if NOT rsEntry2.EOF then
						if rsEntry2("Available") < CINT(request.form("txtQty"&rsEntry("SortOrderID"))) then
							quantity = rsEntry2("Available")
							if quantity < 0 then
								quantity = 0
							end if
							%>
                            <%ReplaceInPhrase phraseDictionary, "Maxqtyavailable", "<QUANTITY>", quantity %>
                            <script type="text/javascript">alert("<%=DisplayPhraseJS(phraseDictionary,"Maxqtyavailable")%>");</script><%
						end if
					end if
					rsEntry2.Close
				end if
				if NOT rsEntry("PricePerSession") then
					strSQL = "UPDATE tblShoppingCart SET Quantity=" & sqlClean(quantity) & " WHERE (SessionID = N'" & sqlClean(getSessionGUID()) & "') AND SortOrderID=" & sqlClean(rsEntry("SortOrderID"))
				else 
					strSQL = "UPDATE tblShoppingCart SET NumSessions=" & sqlClean(quantity) & " WHERE (SessionID = N'" & sqlClean(getSessionGUID()) & "') AND SortOrderID=" & sqlClean(rsEntry("SortOrderID"))
				end if
                cnWS.execute strSQL

                ' update promotion amounts with quantity
                if NOT isNull(rsEntry("PromotionID")) then
                    strSQL = "SELECT PromotionID, DiscPerc, DiscAmt FROM tblPromotion WHERE (PromotionID = " & sqlClean(rsEntry("PromotionID")) & ")"
                    rsEntry2.CursorLocation = 3
                    rsEntry2.open strSQL, cnWS
                    Set rsEntry2.ActiveConnection = Nothing
                    if NOT rsEntry2.EOF then
                        if NOT isNULL(rsEntry2("DiscPerc")) then
                            curDiscPerc = rsEntry2("DiscPerc")

                            if rsEntry("CreditAmount")<>0 then
                                curDiscAmt = CSng((request.form("txtQty"&rsEntry("SortOrderID")) * rsEntry("CreditAmount")) * rsEntry2("DiscPerc") * 0.01)
                            else
                                if rsEntry("PricePerSession") then
                                        curDiscAmt = CSng((request.form("txtQty"&rsEntry("SortOrderID")) * (rsEntry("OnlinePrice")*rsEntry("NumSessions"))) * rsEntry2("DiscPerc") * 0.01)
                                else
                                        curDiscAmt = CSng((request.form("txtQty"&rsEntry("SortOrderID")) * (rsEntry("OnlinePrice"))) * rsEntry2("DiscPerc") * 0.01)
                                end if
                            end if
                        else
                            curDiscAmt = CSng((request.form("txtQty"&rsEntry("SortOrderID")) * rsEntry2("DiscAmt")))
                            if rsEntry("CreditAmount")<>0 then
                                 curDiscPerc = FormatNumber((rsEntry2("DiscAmt") / rsEntry("CreditAmount")) * 100)
                            else
                                if rsEntry("OnlinePrice") <> 0 then 
                                    if rsEntry("PricePerSession") then
                                        curDiscPerc = FormatNumber((rsEntry2("DiscAmt") / (rsEntry("OnlinePrice")*quantity)) * 100)
                                    else
                                        curDiscPerc = FormatNumber((rsEntry2("DiscAmt") / (rsEntry("OnlinePrice"))) * 100)
                                    end if
                                else
                                    curDiscPerc = 0
                                    curDiscAmt = 0
                                end if
                            end if
                        end if

                        'prevent promotions from giving money away
                        if curDiscPerc > 100 then
                            curDiscPerc = FormatNumber(100)
                        end if
                        if curPromoDiscAmt >  rsEntry("OnlinePrice") then
                            curPromoDiscAmt = FormatNumber(rsEntry("OnlinePrice"))
                        end if
                                                
                        strSQL = "UPDATE tblShoppingCart SET PromoDiscPerc=" & sqlClean(curDiscPerc) & ", PromoDiscAmt=" & sqlClean(curDiscAmt) & " WHERE (SessionID = N'" & sqlClean(getSessionGUID()) & "') AND SortOrderID=" & sqlClean(rsEntry("SortOrderID"))
                        cnWS.execute strSQL
                    end if
                    rsEntry2.close
                end if
            end if
                rsEntry.MoveNext
        loop
        rsEntry.close
                
        ' Update Cart - PARTNER STORES
        strSQL = "SELECT SortOrderID, Quantity FROM tblShoppingCart WHERE (SessionID = N'" & sqlClean(getSessionGUID()) & "') AND (SaleID IS NULL) AND PartnerID<>0 ORDER BY SortOrderID "
        rsEntry.CursorLocation = 3
        rsEntry.open strSQL, cnWS
        Set rsEntry.ActiveConnection = Nothing
        do while NOT rsEntry.EOF
            if request.form("optDelete"&rsEntry("SortOrderID"))="on" then
                strSQL = "DELETE FROM tblShoppingCart WHERE (SessionID = N'" & sqlClean(getSessionGUID()) & "') AND SortOrderID=" & sqlClean(rsEntry("SortOrderID"))
                cnWS.execute strSQL

                numItemsUpdated = true
            elseif request.form("txtQty"&rsEntry("SortOrderID"))<>CSTR(rsEntry("Quantity")) AND request.form("txtQty"&rsEntry("SortOrderID"))<>"" then
                strSQL = "UPDATE tblShoppingCart SET Quantity=" & sqlClean(request.form("txtQty"&rsEntry("SortOrderID"))) & " WHERE (SessionID = N'" & sqlClean(getSessionGUID()) & "') AND SortOrderID=" & sqlClean(rsEntry("SortOrderID"))
                cnWS.execute strSQL
            end if
            rsEntry.MoveNext
        loop
        rsEntry.close
                
        'BJD: 51_1003 - Charity donation
        if session("PartnersEnabled") then
                dim charityPartnerID, charityPartnerName, charityPartnerDesc, charityWebsite, charityDonationProductID
                strSQL = "SELECT tblPartner.PartnerID, tblPartner.PartnerName, tblStudioPartner.StudioID, tblPartner.PartnerDesc, tblPartner.Website, tblCatalog.ProductID "&_
                         "FROM tblPartner "&_
                         "INNER JOIN tblStudioPartner ON tblPartner.PartnerID = tblStudioPartner.PartnerID "&_
                         "INNER JOIN tblCatalog ON tblPartner.PartnerID = tblCatalog.PartnerID "&_
                         "WHERE (tblPartner.Active = 1) AND (tblPartner.SetupMode = 0) AND (tblPartner.Charity=1) AND (tblCatalog.vmsProductID=-9) AND (tblStudioPartner.StudioID = " & sqlClean(session("studioID")) & ") "
                if ss_DemoSite="0" then
                     strSQL = strSQL & "AND (tblPartner.TestMode = 0) "
                end if
                strSQL = strSQL & "ORDER BY tblPartner.PartnerName "
                rsEntry.CursorLocation = 3
                rsEntry.open strSQL, cnMBPS
                Set rsEntry.ActiveConnection = Nothing
        
                if NOT rsEntry.EOF then
                    charityPartnerID = rsEntry("PartnerID")
                    charityPartnerName = rsEntry("PartnerName")
                    charityPartnerDesc = rsEntry("PartnerDesc")
                    charityWebsite = rsEntry("Website")
                    charityDonationProductID = rsEntry("ProductID")
                end if
                rsEntry.close
                                        
                'BJD: 51_1003 - update donation in shopping cart
                if request.form("frmDonationChanged")="true" then
                    strSQL = "DELETE FROM tblShoppingCart WHERE PartnerID=" & sqlClean(charityPartnerID)
                    cnWS.execute strSQL
                                
                    if IsNum(request.form("optDonationAmount")) then
                        if CSNG(request.form("optDonationAmount"))>0 then
                            sortOrderID = 1
                            'Get Next SortOrderID
                            strSQL = "SELECT MAX(SortOrderID) AS NextID, SessionID FROM tblShoppingCart GROUP BY SessionID HAVING (SessionID = N'" & sqlClean(getSessionGUID()) & "')"
                            rsEntry.CursorLocation = 3
                            rsEntry.open strSQL, cnWS
                            Set rsEntry.ActiveConnection = Nothing
                            if NOT rsEntry.EOF then
                                    sortOrderID = rsEntry("NextID") + 1
                            end if
                            rsEntry.close
                                        
                            strSQL = "INSERT INTO tblShoppingCart (SessionID, SortOrderID, Created, ProductID, Quantity, CreditAmount, PartnerID) VALUES ("
                            strSQL = strSQL & "N'" & sqlClean(getSessionGUID()) & "'"
                            strSQL = strSQL & ", " & sqlClean(sortOrderID)
                            strSQL = strSQL & ", " & DateSep & sqlClean(DateAdd("n", Session("tzOffset"),Now)) & DateSep
                            strSQL = strSQL & ", " & sqlClean(charityDonationProductID)
                            strSQL = strSQL & ",1"
                            strSQL = strSQL & ", " & sqlClean(request.form("optDonationAmount"))
                            strSQL = strSQL & ", " & sqlClean(charityPartnerID)
                            strSQL = strSQL & ") "
                            cnWS.execute strSQL
                    end if
                end if
            end if
        end if ' if partners are enabled
                
                '''''''''''' END SHOPPING CART """"""""""""""""""""""
    end if

    strSQL = "SELECT COUNT(SortOrderID) AS NumItems FROM tblShoppingCart WHERE (SaleID IS NULL) GROUP BY SessionID HAVING (SessionID = N'" &  sqlClean(getSessionGUID()) & "')"
    rsEntry.CursorLocation = 3
    rsEntry.open strSQL, cnWS
    Set rsEntry.ActiveConnection = Nothing
    if NOT rsEntry.EOF then
            sCartNumItems = rsEntry("NumItems")
    else
            sCartNumItems = 0
    end if
    rsEntry.close
%>
<!-- #include file="pre.asp" -->
<!-- #include file="frame_bottom.asp" -->
<!-- begin client alerts -->
<%
'client alert context vars
focusFrmElement = ""
cltAlertList = setClientAlertsList(session("mvarUserID"))
%>
<!-- #include file="inc_ajax.asp" -->
<!-- #include file="adm/inc_alert_js.asp" -->
<!-- end client alerts  -->

			<%= css(array("calendar")) %>
<%= js(array("MBS", "valcur", "calendar" & dateFormatCode, "VCC2", "plugins/jquery.placeholder", "plugins/jquery.lightboxLib", "plugins/jquery.SimpleLightBox", "plugins/jquery.moreless")) %>
<%= css(array("gallery","SimpleLightBox")) %>
<!-- #include file="inc_date_ctrl.asp" -->

<%= js(array("mb", "main_shop")) %>

<script type="text/javascript">

function donationUpdated() {
        document.frmShop.frmDonationChanged.value = "true";
        SelView(4);
}

function goToPage(pNum) {
        document.frmShop.frmPageNum.value = pNum;
        document.frmShop.lastSearch.value = "<%=xssStr(lastSearch)%>";
        doSearch();
}
function addRetailItem(ndx, addMode, partnerID) {
        if (document.getElementById("optProdGroup"+ndx)!=null) {
                document.frmShop.frmProdID.value = document.getElementById("optProdGroup"+ndx).options[document.getElementById("optProdGroup"+ndx).selectedIndex].value;
        } else if (addMode == 3) {
                document.frmShop.frmProdID.value = ndx;
                addMode = 2;
        } else {
                document.frmShop.frmProdID.value = document.getElementById("frmProdGroup"+ndx).value;
        }
        document.frmShop.frmPartnerID.value = partnerID;
        document.frmShop.frmPageNum.value = "<%=xssStr(request.form("frmPageNum"))%>";
        document.frmShop.lastSearch.value = "<%=xssStr(lastSearch)%>";
        addToCart(addMode);
}


//END - JS For Retail Online Store

function addToCart(addMode) {
	//check for valid input on editable gift card
	var $amount = $('input[name=requiredtxtCreditAmount]');
	//if it exists, and it doesn't have a valid currency,
	//set it's value back to form's starting value, and prevent form submission
	if ($amount[0]){
		if (!validateCurrency($amount[0])){
			$amount.val('<%=frmCredit%>')
			return false;
		}
	}

	if (document.getElementById("opt_ContractIAgree") != null && !document.getElementById("opt_ContractIAgree").checked) {
			alert("<%=DisplayPhraseJS(phraseDictionary,"Pleasecheckiagree")%>");
	} else {

		<%if stype = "43" then%>
			document.frmShop.action = "main_shop.asp?pMode=4&stype=<%= stype %><%=recurringQString%>";
		<%end if%>

		//If the user has chosen to send the gift card to the recipiant via email,
		//change the necessary field names so that they will be verified not empty
		//along with the other required fields in the function 'checkrequired'
		var $emailCheckbox = $('#optSendGiftCardByEmail'),
            $email1 = $('#egiftCardMail1'),
            $email2 = $('#egiftCardMail2'),
            $deliveryDate = $('#optGCDeliveryDate');

		if ($emailCheckbox.is(':checked')) {
			$email1.attr('name', 'requiredEm_RecipientEmail1');
			$email2.attr('name', 'requiredEm_RecipientEmail2');
			$deliveryDate.attr('name', 'requiredtxtDeliveryDate');
		}
		else {
			$email1.attr('name', 'egiftCardMail1')
			$email2.attr('name', 'egiftCardMail2')
			$deliveryDate.attr('name', 'optGCDeliveryDate');
		}

		if (checkrequired(document.frmShop)) {
			<%if pMode=2 then%>
				//Gift Card Case
				// Validate gift card email
				var emailRegEx = /^[a-zA-Z0-9._-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,4}$/;

				if ($emailCheckbox.is(':checked')) {
					if ($email1.val().toLowerCase() != $email2.val().toLowerCase()) {
					    alert('<%=DisplayPhraseJS(phraseDictionary,"Emailaddressesdonotmatch")%>');
					} else if (emailRegEx.test($email1.val()) == false && $email1.val() != "") {
						alert('<%=DisplayPhraseJS(phraseDictionary,"Emailaddressnotvalid")%>');
					} else {
						document.frmShop.frmAddCart.value = addMode;
						document.frmShop.submit();
					}
				} else {
					document.frmShop.frmAddCart.value = addMode;
					document.frmShop.submit();
				}
			<%else%>
				document.frmShop.frmAddCart.value = addMode;
				document.frmShop.submit();
			<%end if%>
		}
	}
}

function CheckOut() {
<%if NOT session("pass") then%>
        document.frmShop.action = "su1.asp";
<%else%>
        if (document.frmShop.frmShipItems.value=="True") {
                document.frmShop.action = "shop_chkout_ship.asp?<%=recurringQString%>";
        } else {
                document.frmShop.action = "shop_chkout.asp?<%=recurringQString%>";
        }
<%end if%>
        document.frmShop.submit();
}

function chkPromo(field, onfocus) {
        if (onfocus && field.value=="<%=DisplayPhraseJS(phraseDictionary,"Promotioncode")%>") {
                field.value = "";
        } else if (!onfocus && field.value=="") {
                field.value = "<%=DisplayPhraseJS(phraseDictionary,"Promotioncode")%>";
        }
}

        
$(document).ready(function() {
		// defined in siteLteIe7.less only important in IE7
		$('div.wrapper').addClass("main_shop-cm-wrapper"); 

        $('input[name=optSendGiftCardByEmail]').click(function () {
                                     
            if ($(this).is(':checked')) {
               $('.giftCardDeliveryInfo').show();                
            }
            else {
               $('.giftCardDeliveryInfo').hide();
                
            }
        }).triggerHandler('click');
});

 
</script>
<%
        dim cssColor3

        strSQL = "SELECT CSSColor1, CSSColor2, CSSColor3, CSSColor4, TopBgColor FROM tblAppearance"

        rsEntry.CursorLocation = 3
        rsEntry.open strSQL, cnWS
        Set rsEntry.ActiveConnection = Nothing
        
        if NOT rsEntry.EOF then
                cssColor3 = rsEntry("CSSColor3")
        end if
        rsEntry.close
%>

<%= css(array("inc_retail_links", "main_shop")) %>
</head>
<body>
<!-- #include file="inc_retail_links.asp" -->
<% ShowSubTabLinks ("") %>
<!-- #include file="adm/inc_alert_content.asp" -->
<% pageStart %>
<script type="text/javascript">
	//parent.topFrame.document.getElementById('top_shopDiv').OnComplete = reloadTopShop();

	function reloadTopShop() {
		//parent.topFrame.location.href = parent.topFrame.location.href;
	}
</script>
<div id="shoppingCart" class="group">
    
<form id="frmShop" name="frmShop" action="main_shop.asp" method="post">
    <input type="hidden" name="reSchedule" value="<%=xssStr(reSchedule)%>" />
    <input type="hidden" name="origId" value="<%=xssStr(origId)%>" />
    <input type="hidden" name="frmMode" value="<%=xssStr(pMode)%>">
    <input type="hidden" name="frmrelClientID" value="<%=xssStr(relClientID)%>">
    <input type="hidden" name="frmLastMode" value="<%=xssStr(pMode)%>">
    <input type="hidden" name="frmSubmitted" value="<%=session("StudioID")%>">
    <input type="hidden" name="frmAddCart" value="">
    <input type="hidden" name="frmUnpaid" value="<%=xssStr(unPaidCase)%>">
    <input type="hidden" name="frmUnpaidRec" value="<%=xssStr(unPaidCaseRec)%>">
    <input type="hidden" name="frmVD_Date" value="<%=xssStr(pVD_Date)%>">
    <input type="hidden" name="frmStupid" value="<%=xssStr(request.form("txtRegStartDate"))%>">     
    <input type="hidden" name="frmVD_ClassID" value="<%=xssStr(pVD_ClassID)%>">
    <input type="hidden" name="frmVD_CourseID" value="<%=xssStr(pVD_CourseID)%>">
    <input type="hidden" name="frmVD_LeadOrFollow" value="<%=xssStr(pVD_LeadOrFollow)%>">
    <input type="hidden" name="frmVD_Rec_EDate" value="<%=xssStr(pVD_Rec_EDate)%>">
    <input type="hidden" name="frmVD_Rec_DayStr" value="<%=xssStr(pVD_Rec_DayStr)%>">
    <input type="hidden" name="frmVD_TrnID" value="<%=xssStr(pVD_TrnID)%>">
    <input type="hidden" name="frmVD_RTrnID" value="<%=xssStr(pVD_RTrnID)%>">
    <input type="hidden" name="frmVD_STime" value="<%=xssStr(pVD_STime)%>">
    <input type="hidden" name="frmVD_ETime" value="<%=xssStr(pVD_ETime)%>">
    <input type="hidden" name="frmVD_TG" value="<%=xssStr(pVD_TG)%>">
    <input type="hidden" name="frmVD_VT" value="<%=xssStr(pVD_VT)%>">
    <input type="hidden" name="frmVD_Loc" value="<%=xssStr(pVD_Loc)%>">
    <input type="hidden" name="frmVD_Notes" value="<%=xssStr(pVD_Notes)%>">
    <input type="hidden" name="frmVD_WaitList" value="<%=xssStr(pVD_WaitList)%>">
    <input type="hidden" name="frmEnrollType" value="<%=xssStr(pEnrollType)%>">
    <input type="hidden" name="frmUsePmtPlan" value="<%=xssStr(pUsePmtPlan)%>">
    <input type="hidden" name="frmEditMode" value="">
    <input type="hidden" name="frmRegisterType" value="<%=xssStr(pRegisterType)%>">
    <input type="hidden" name="optEnrollDay1" value="<%if request.form("optEnrollDay1")="on" or request.querystring("optEnrollDay1")="on" then response.write "on" end if%>">
    <input type="hidden" name="optEnrollDay2" value="<%if request.form("optEnrollDay2")="on" or request.querystring("optEnrollDay2")="on" then response.write "on" end if%>">
    <input type="hidden" name="optEnrollDay3" value="<%if request.form("optEnrollDay3")="on" or request.querystring("optEnrollDay3")="on" then response.write "on" end if%>">
    <input type="hidden" name="optEnrollDay4" value="<%if request.form("optEnrollDay4")="on" or request.querystring("optEnrollDay4")="on" then response.write "on" end if%>">
    <input type="hidden" name="optEnrollDay5" value="<%if request.form("optEnrollDay5")="on" or request.querystring("optEnrollDay5")="on" then response.write "on" end if%>">
    <input type="hidden" name="optEnrollDay6" value="<%if request.form("optEnrollDay6")="on" or request.querystring("optEnrollDay6")="on" then response.write "on" end if%>">
    <input type="hidden" name="optEnrollDay7" value="<%if request.form("optEnrollDay7")="on" or request.querystring("optEnrollDay7")="on" then response.write "on" end if%>">
    <input type="hidden" name="txtRegStartDate" value="<%=xssStr(request.form("txtRegStartDate"))%>">
<% 'BJD: 45_2313 - new hiddens %>
    <input type="hidden" name="frmVD_recType" value="<%=xssStr(pVD_recType)%>">
    <input type="hidden" name="frmVD_recNum" value="<%=xssStr(pVD_recNum)%>">
<% 'BJD: 6/30/08 - Replacing form variables that got nuked when I removed the old shop code %>  
    <input type="hidden" name="frmProdID" value="">
    <input type="hidden" name="frmPartnerID" value="">
    <input type="hidden" name="frmPageNum" value="">
    <input type="hidden" name="lastSearch" value="">
    <input type="hidden" name="optReservedFor" value="<%=xssStr(optReservedFor) %>">
    <input type="hidden" name="stype" value="<%=xssStr(stype) %>" />
        <% if request.QueryString("recCount") <> "" OR request.Form("frmRecCount")<>"" then %>
				<input type="hidden" name="frmRecCount" value="<%if request.QueryString("recCount") <> "" then response.write request.QueryString("recCount") else response.write request.Form("frmRecCount") end if %>" />
				<% end if %>
        <% 'For guest logins %>
    <input type="hidden" name="optLoginAsGuest" value="" />
<%      
        'added for open enrollment 3/21/08 by Brad  

        'CB 12/31/2008 - Added to Preserve Days of Week from res_a
        for dayCount=1 to 7
                if request.form("Day"&dayCount)="true" then 
%>
        <input type="hidden" name="Day<%=dayCount%>" value="true">
<%
                end if
        next

        
        frmRtnClassDate = pVD_Date
        frmRtnNumSessions = 0 

        if pVD_CourseID<>"" then
            strSQL = " SELECT tblClasses.ClassID, tblClasses.DaySunday as Day1, tblClasses.DayMonday as Day2, tblClasses.DayTuesday as Day3, tblClasses.DayWednesday as Day4, "
            strSQL = strSQL & " tblClasses.DayThursday as Day5, tblClasses.DayFriday as Day6, tblClasses.DaySaturday as Day7, tblClassDescriptions.ClassName, tblClasses.ClassDateStart, "
            strSQL = strSQL & " tblClasses.ClassDateEnd, ISNULL(tblVisitTypes.NumDeducted, 0) as NumDeducted "
            strSQL = strSQL & " FROM tblClasses INNER JOIN "
            strSQL = strSQL & " tblClassDescriptions ON tblClasses.DescriptionID = tblClassDescriptions.ClassDescriptionID "
            strSQL = strSQL & " INNER JOIN tblVisitTypes ON tblClassDescriptions.VisitTypeID = tblVisitTypes.TypeID "
            strSQL = strSQL & " WHERE tblClasses.CourseID = " & sqlClean(pVD_CourseID)
            strSQL = strSQL & " ORDER BY ClassDateStart "
            rsEntry.CursorLocation = 3
            rsEntry.open strSQL, cnWS
            Set rsEntry.ActiveConnection = Nothing
            if NOT rsEntry.EOF then
                tmpNumSessions = 0
                
                do while NOT rsEntry.EOF
                    if request.form("optClassSchEDate")<>"" then
                        if isDate(request.form("optClassSchEDate")) then
                            frmRtnClassEDate = CDATE(request.form("optClassSchEDate"))
                        else
                            frmRtnClassEDate = CDATE(frmRtnClassDate)
                        end if
                    'pVD_Rec_EDate = frmRtnClassEDate
                    elseif request.form("frmVD_Rec_EDate")<>"" then
                        frmRtnClassEDate = CDATE(request.form("frmVD_Rec_EDate"))
                        'pVD_Rec_EDate = frmRtnClassEDate
                    elseif request.form("txtEDate")<>"" then
                        frmRtnClassEDate = CDATE(request.form("txtEDate"))
                    else
                        frmRtnClassEDate = CDATE(frmRtnClassDate)
                    end if
                    
                    frmRtncont = true
                    frmRtntmpDate = frmRtnClassDate
                    if CDATE(frmRtntmpDate) < CDATE(rsEntry("ClassDateStart")) then
                        frmRtntmpDate = rsEntry("ClassDateStart")
                    end if

                    'MB bug #4516
                    if CDATE(frmRtnClassEDate) > CDATE(rsEntry("ClassDateEnd")) then
                        frmRtnClassEDate = rsEntry("ClassDateEnd")
                    end if
                    
                    do while frmRtncont
                        if rsEntry("Day" & WeekDay(frmRtntmpDate)) then %>
                            <input type="hidden" name="optDay<%=FmtDateShort(frmRtntmpDate)%>-<%=rsEntry("ClassID")%>" value="<%=xssStr(request.form("optDay" & FmtDateShort(frmRtntmpDate) & "-" & rsEntry("ClassID")))%>">
<%
                            tmpNumSessions = tmpNumSessions + rsEntry("NumDeducted")
                            if request.form("optDay" & FmtDateShort(frmRtntmpDate) & "-" & rsEntry("ClassID")) = "on" OR isNull(request.form("optDay" & FmtDateShort(frmRtntmpDate) & "-" & rsEntry("ClassID"))) then
                                frmRtnNumSessions = frmRtnNumSessions + rsEntry("NumDeducted")
                            end if
                        end if
                        frmRtntmpDate = CDATE(DATEADD("d", 1, frmRtntmpDate))

                        if frmRtntmpDate > CDATE(rsEntry("ClassDateEnd")) then
                            frmRtncont = false
                        end if
                    loop
                    rsEntry.MoveNext
                loop
                if frmRtnNumSessions = 0 then
                    frmRtnNumSessions = tmpNumSessions
                end if
            end if
            rsEntry.close
        else
            if request.form("optClassSchEDate")<>"" then
                    if isDate(request.form("optClassSchEDate")) then
                            frmRtnClassEDate = CDATE(request.form("optClassSchEDate"))
                    else
                            frmRtnClassEDate = CDATE(frmRtnClassDate)
                    end if
                    'pVD_Rec_EDate = frmRtnClassEDate
            elseif request.form("frmVD_Rec_EDate")<>"" then
                    frmRtnClassEDate = CDATE(request.form("frmVD_Rec_EDate"))
                    'pVD_Rec_EDate = frmRtnClassEDate
            elseif request.form("txtEDate")<>"" then
                    frmRtnClassEDate = CDATE(request.form("txtEDate"))
            else
                    frmRtnClassEDate = CDATE(frmRtnClassDate)
            end if
            
            frmRtncont = true
            frmRtntmpDate = frmRtnClassDate

            if frmRtntmpDate<>"" and frmRtnClassEDate<>"" and pVD_ClassID<>"" then
				dim holidaysInRR : holidaysInRR = GetHolidaysInRange(frmRtntmpDate, frmRtnClassEDate, typeGroupID)
				strSQL = "SELECT SUM(CC.ClassCount * CC.NumDeducted) as NumSessions, SUM(CC.NumDeducted) as NumDeducted "&_
						" FROM "&_
						" ( "&_
						"	SELECT COUNT(tblClassSch.ClassDate) as ClassCount, ISNULL(tblVisitTypes.NumDeducted, 0) as NumDeducted "&_
						"	FROM tblClassSch "&_
						"	INNER JOIN tblClassDescriptions ON tblClassSch.DescriptionID = tblClassDescriptions.ClassDescriptionID "&_
						"	INNER JOIN tblVisitTypes ON tblClassDescriptions.VisitTypeID = tblVisitTypes.TypeID "&_
						"	WHERE tblClassSch.ClassID = " & sqlClean(pVD_ClassID) & " AND TrainerID<>-1 "&_
						"		AND tblClassSch.ClassDate >= " & DateSep & sqlClean(frmRtntmpDate) & DateSep &_
						"		AND tblClassSch.ClassDate <= " & DateSep & sqlClean(frmRtnClassEDate) & DateSep
				if pVD_Rec_DayStr<>"" then
					strSQL = strSQL & " AND (datepart(WEEKDAY, tblClassSch.ClassDate) IN (" & sqlClean(pVD_Rec_DayStr) & "-999)) "
				end if
				if (not isNull(holidaysInRR)) then
					strSQL = strSQL & " AND tblClassSch.ClassDate NOT IN ( " & sqlClean(DateArrayToInSQL(holidaysInRR)) & " ) "
				end if
				strSQL = strSQL &_
						"	GROUP BY tblVisitTypes.NumDeducted "&_
						" ) AS CC "
				'response.write debugSQL(strSQL, "SQL")
				rsEntry.CursorLocation = 3
				rsEntry.open strSQL, cnWS
				Set rsEntry.ActiveConnection = Nothing
				if NOT rsEntry.EOF then
					frmRtnNumDeducted = rsEntry("NumDeducted")
					tmpNumSessions = rsEntry("NumSessions")
				end if
				rsEntry.close
            end if

            do while frmRtncont
                if request.form("Day" & WeekDay(frmRtntmpDate))="true" then 
%>
                    <input type="hidden" name="optDay<%=FmtDateShort(frmRtntmpDate)%>-<%=pVD_ClassID%>" value="<%=xssStr(request.form("optDay" & FmtDateShort(frmRtntmpDate) & "-" & pVD_ClassID))%>">
    <%              if request.form("optDay" & FmtDateShort(frmRtntmpDate) & "-" & pVD_ClassID) = "on" then
                        frmRtnNumSessions = frmRtnNumSessions + frmRtnNumDeducted
                    end if
                end if
                frmRtntmpDate = CDATE(DATEADD("d", 1, frmRtntmpDate))

                if frmRtntmpDate > frmRtnClassEDate then
                        frmRtncont = false
                end if
            loop
            ' end added 3/21/08 brad
            if frmRtnNumSessions = 0 then
                frmRtnNumSessions = tmpNumSessions
            end if

        end if
	%>
    <input type="hidden" name="optNumSessions" value="<%=frmRtnNumSessions%>">
    <input type="hidden" name="optClassSchEDate" value="<%=xssStr(frmRtnClassEDate)%>">
    <input type="hidden" name="frmDonationChanged" value="">

    <div id="mainShop">
        <div id="pageTitle">
            <h1>
            <%  dim ContractsAndPackages
				if pMode = 0 then ' Contracts and Packages
                    if ContractsOk AND PackagesOk then
                        ContractsAndPackages = DisplayPhrase(phraseDictionary,"Contracts") & " / " & DisplayPhrase(phraseDictionary,"Packages")
                    elseif ContractsOk then
                        ContractsAndPackages = DisplayPhrase(phraseDictionary,"Contracts")
                    else
                        ContractsAndPackages = DisplayPhrase(phraseDictionary,"Packages")
                    end if
                    response.Write(ContractsAndPackages)
                elseif pMode = 1 then 'SERIES & MEMBERSHIPS
                    response.Write DisplayPhrase(phraseDictionary,"Services")
                elseif pMode = 2 then ' Gift cards
                    response.Write DisplayPhrase(phraseDictionary,"Giftcards")
                elseif pMode = 3 then 'Retail store
                elseif pMode = 4 then ' Shopping Cart
                    response.Write("<img src=""" & contentUrl("/asp/images/shopping_cart_icon.png") & """ />&nbsp;" & DisplayPhrase(phraseDictionary,"Shoppingcart"))
				elseif pMode = 5 then ' Acct Credit
					RW xssStr(allHotWords(767))
                end if %>
            </h1>
            <% if Session("mvarMIDs") = 0 then 
            'JM-51_2801
                if NOT ccMode AND NOT ss_AccountPaymentsConsumerMode then%>
                    
                <div ><strong><span class="errorColor"><img src="<%= contentUrl("/asp/adm/images/alert-red-16px.png") %>">&nbsp;<%=DisplayPhrase(phraseDictionary,"Wearenotacceptingpayments")%></span></strong></div>
                        
                <%elseif NOT ccMode AND ss_AccountPaymentsConsumerMode AND NOT ss_AllowNegativeBalConsMode then%>
                        
                <div ><strong><span class="errorColor"><img src="<%= contentUrl("/asp/adm/images/alert-red-16px.png") %>">&nbsp;<%=DisplayPhrase(phraseDictionary,"Accountcreditrequired")%></span></strong></div>
                        
                <%end if 
            end if %>
        </div>
<%
			'''''''''''''BEGIN CONTRACTS & PACKAGES''''''''''''''''''''''''''
			if pMode=0 then 'CB 45_1276
%>
			<!--#include file="inc_main_shop_contracts_packages.asp"-->
<%
			end if 'CB 45_1276
			'''''''''''''END CONTRACTS & PACKAGES''''''''''''''''''''''''''

			''''''''''''''BEGIN SERIES & MEMBERSHIPS''''''''''''''''''''''''''
			if pMode=1 then
%>
			<!-- #include file="inc_main_shop_services.asp" -->
<%
			end if
			''''''''''''''END SERIES & MEMBERSHIPS''''''''''''''''''''''''''

			''''''''''''''BEGIN GC/CREDIT''''''''''''''''''''''''''
			if pMode=2 OR pMode=5 then
%> 
			<div id="leftPanel" class="group">
				<div class="section group <%if not session("PartnersEnabled") then response.write("no-partners") end if%>">
					<div id="CreditPanel">
<%                
				strSQL = "SELECT Products.ProductID, ItemTypeID, Description, UnitPrice, ClientCredit, GiftCardEditableInConsumerMode, ProductNotes, GiftCertificate, DebitCard, ( SELECT COUNT(GCLayoutID) FROM tblGCLayout WHERE Active = 1  ) as Layouts  FROM Products "
				strSQL = strSQL & "LEFT OUTER JOIN (SELECT ProductID FROM tblProductSeriesTypeSetting WHERE (SeriesTypeID = " & sqlClean(cMembershipID) & ") AND (Setting = 1)) AS MemLevel ON PRODUCTS.ProductID = MemLevel.ProductID LEFT OUTER JOIN (SELECT COUNT(*) AS NumRestrictions, ProductID FROM tblProductSeriesTypeSetting WHERE (Setting = 1) GROUP BY ProductID) AS MemRestrict ON PRODUCTS.ProductID = MemRestrict.ProductID "
				strSQL = strSQL & "WHERE Discontinued=0 AND wsShow=1 AND [Delete]=0 AND Products.CategoryID BETWEEN 22 AND 23 "
				if NOT isMember then ' BJD: 6/10/08 - new 'members only' logic
					strSQL = strSQL & " AND ((MemRestrict.NumRestrictions = 0) OR (MemRestrict.NumRestrictions IS NULL)) "
				else
					strSQL = strSQL & " AND ((MemLevel.ProductID IS NOT NULL)"
					if mem_AllowNonMemberPurchases then
						strSQL = strSQL & " OR (MemRestrict.NumRestrictions = 0) OR (MemRestrict.NumRestrictions IS NULL)"
					end if
					strSQL = strSQL & ")"
				end if
				strSQL = strSQL & " ORDER BY Description"
				rsEntry.CursorLocation = 3
				rsEntry.open strSQL, cnWS
				Set rsEntry.ActiveConnection = Nothing
				if NOT rsEntry.EOF then
    
					Dim giftcardname
                   
					if rsEntry.RecordCount = 1 then
						productID = rsEntry("ProductID")
					end if
                    if session("giftCardID")> 0 then
                        productID = CINT(session("giftCardID"))
                    end if
                    session("giftCardID") = 0
                   
%>
						<label for="requiredoptPurchaseItem">							
<% 							if pMode=2 then		'Gift Cards				%>
								<%=DisplayPhraseAttr(phraseDictionary,"Whichgiftcard")%>
<%							elseif pMode=5 then 'Account Payments		%>
								<%=DisplayPhraseAttr(phraseDictionary,"Whichacctpayment")%>
<%							end if										%>							
						</label>
						<select name="requiredoptPurchaseItem" id="requiredoptPurchaseItem" onchange="document.frmShop.submit();" >
							<option value="0">
<% 							if pMode=2 then		'Gift Cards				%>
								<%=allHotWords(802)%>
<%							elseif pMode=5 then 'Account Payments		%>
								<%=allHotWords(801)%>
<%							else										%>	
								<%=allHotWords(727)%>
<%							end if										%>
							</option>
<%                  'Gift Cards/Certificates
                    if pMode=2 then
					    Do While not rsEntry.EOF
						     ''if  rsEntry("GiftCertificate") AND rsEntry("Layouts") > 0  then
                               if (rsEntry("ItemTypeID")=4) then
    %>           
							    <option value="<%=rsEntry("ProductID")%>" name="<%=rsEntry("Description") & " for " & FmtCurrency(rsEntry("UnitPrice")) %>" <%if CSTR(productID)=CSTR(rsEntry("ProductID")) then response.write " selected" end if%>>
								    <%=rsEntry("Description")%>
								    <%if NOT rsEntry("GiftCardEditableInConsumerMode") OR rsEntry("GiftCertificate") then %>&nbsp;<%=allHotWords(56)%>&nbsp;
								    <%= FmtCurrency(rsEntry("UnitPrice"))%><% end if%>
							    </option>
    <%
							    if CSTR(productID)=CSTR(rsEntry("ProductID")) then
								    prodSelected = true
								    if rsEntry("GiftCertificate") then
									    isGC = true
								    end if
								    if rsEntry("DebitCard") then
									    isPPGC = true
								    end if
								    if rsEntry("GiftCardEditableInConsumerMode") then
									    isEditable = true
								    end if
								    frmCredit = Replace(FormatNumber(rsEntry("ClientCredit")), ",", "")
								    giftcardname = rsEntry("Description")
							    end if
						    end if
						    rsEntry.MoveNext
					    Loop
                     end if
                    'Account Credits
                    if pMode=5 then 
                     Do While not rsEntry.EOF
                            if (rsEntry("ItemTypeID")=3) then
                            ''if NOT (rsEntry("GiftCertificate") AND rsEntry("Layouts") > 0)  then 
    %>           
							    <option value="<%=rsEntry("ProductID")%>" name="<%=rsEntry("Description") & " for " & FmtCurrency(rsEntry("UnitPrice")) %>" <%if CSTR(productID)=CSTR(rsEntry("ProductID")) then response.write " selected" end if%>>
								    <%=rsEntry("Description")%>
								    <%if NOT rsEntry("GiftCardEditableInConsumerMode") OR rsEntry("GiftCertificate") then %>&nbsp;<%=allHotWords(56)%>&nbsp;
								    <%= FmtCurrency(rsEntry("UnitPrice"))%><% end if%>
							    </option>
    <%
							    if CSTR(productID)=CSTR(rsEntry("ProductID")) then
								    prodSelected = true
								    if rsEntry("GiftCertificate") then
									    isGC = true
								    end if
								    if rsEntry("DebitCard") then
									    isPPGC = true
								    end if
								    if rsEntry("GiftCardEditableInConsumerMode") then
									    isEditable = true
								    end if
								    frmCredit = Replace(FormatNumber(rsEntry("ClientCredit")), ",", "")
								    giftcardname = rsEntry("Description")
							    end if
						    end if
						    rsEntry.MoveNext
					    Loop
                     end if
%>
						</select>
<%
				end if 'rsEntr.EOF
				rsEntry.close
				if isEditable then
%>
						<div class="forie7">
							<label for="requiredtxtCreditAmount" class="clear-both">
								<% 	if pMode=2 then 'if Gift Cards			%>
									<%=DisplayPhraseAttr(phraseDictionary,"Pleaseentergiftcardamount")%>
								<%	elseif pMode=5 then ' if Account Credit %>
									<%=DisplayPhraseAttr(phraseDictionary,"Pleaseenteracctpaymentamount")%>
								<% end if %>								
							</label>
							<span class="required">*</span>
							<input type="text" name="requiredtxtCreditAmount" id="requiredtxtCreditAmount" class="required credit"maxlength="12" value="<%=FmtNumber(CDBL(frmCredit))%>" />
							<span class="exampleAmount">&nbsp;<%=DisplayPhrase(phraseDictionary,"Eg100")%>
							</span>
						</div>
<%
				end if ''''is editable
				if isGC then
					strSQL = "SELECT GCLayoutID, LayoutName FROM tblGCLayout WHERE (Active = 1) ORDER BY LayoutName"
					rsEntry.CursorLocation = 3
					rsEntry.open strSQL, cnWS
					Set rsEntry.ActiveConnection = Nothing
%>
					<!-- #include file="inc_main_shop_gift_cards_credits.asp" -->
					<% ' cant wait to get rid of IE7 %>
					<% if isIE7 then %>
					<style type="text/css">
						#wrapper-frame 
						{
							position: static !important;
						}
					</style>
					<% end if %>
<%
					'Load JS outside of the preview element so they don't get re-evaluated when moved in the dom
					loadFilmstripPlugin
					loadCharCounterPlugin
					consumerModeDocReady
%>
					<div id="preview" class="LightBox">
                        <div class="water-mark-overlay"></div>
						<!-- #include file="inc_gift_card_content.asp" -->
<%
					createCardPreviewHtmlStructure
%>
						</div> <!--#preview-->
<%
					rsEntry.close
				end if


				if prodSelected AND ((NOT pUsePmtPlan) OR PmtSelected) AND showMakePurchaseButton then
%>
						<div class="makePurchase">
							<input onclick="addToCart(2);" class="makePurchaseButton" type="button" name="Button" value="<%=DisplayPhraseAttr(phraseDictionary,"Makepurchase")%>" />
<%
					if isGC then
%>
							<input onclick="" class="previewButton preview" type="button" name="Button" value="<%=DisplayPhraseAttr(phraseDictionary,"Previewgiftcard")%>" />
<%
					end if
%>
						</div>
<%
				end if 'ProdSelected
%>
					</div>
				</div> <!-- .section -->
			</div> <!-- .leftPanel -->
<%
			end if

			''''''''''''''END GC/CREDIT''''''''''''''''''''''''''
%>

        <%if pMode=3 then       '''''''BEGIN RETAIL STORE'''''''''''''''%>
<%
' *********************************************************************************************************************
                if request.querystring("prodid")="" then
                        'if (session("PartnersEnabled") AND ss_DefaultStreetCorner AND request.querystring("catid")="") OR request.querystring("showSC")="true" then
                        '       response.redirect "../../MBPS"
                        'else
                                ' BJD: redirect to studio's retail store in .NET - with catid querystring var
                                response.redirect "/Pages/OnlineStore.aspx?partnerID=0&catid=" & request.querystring("catid")
                        'end if
                end if
                
                'BJD: 6/12/08 - OLD ONLINE STORE CODE REMOVED
                
        end if                  '''''''END RETAIL STORE''''''''''''''' 
%>


    <%if pMode = 4 then ''''''''BEGIN SHOPPING CART'''''''''''' %>
            <!-- #include file="inc_main_shop_cart.asp" -->
    <%end if ''''''''END SHOPPING CART'''''''''''' 


	


	
	if session("Unamerican") and (pMode = 1 or pMode = 0 or pMode = 4) then
    %>
	<div class="notice-taxes-inclusive-tax"><%=DisplayPhrase(phraseDictionary,"Noticepricesincludetax")%></div>
	<%
	end if
	%>

    <!--</div>--> <!-- .section --> 
</div> <!-- .shoppingCartPanel -->             
<%		if rsEntry.State = 1 then ' close the connection if it is open
			rsEntry.Close
		end if

		if isNum(request.querystring("prodid")) then
                if request.querystring("partnerID")="" OR request.querystring("partnerID")="0" then
            if pMode = 0 then   'Contracts & Packages, if package Ok to auto add to cart
                strSQL = "SELECT SellOnline AS wsShow FROM tblContract WHERE (IsPackage = 1) AND (ContractID = " & sqlClean(request.querystring("prodid")) & ")"
            else
                            strSQL = "SELECT wsShow FROM PRODUCTS WHERE ProductID = " & sqlClean(request.querystring("prodid"))
            end if

                        rsEntry.CursorLocation = 3
                        rsEntry.open strSQL, cnWS
                        Set rsEntry.ActiveConnection = Nothing

                else ' partner store product
                        'connectToMBPS
                        strSQL = "Select Active as wsShow FROM tblCatalog Where ProductID = " & sqlClean(request.querystring("prodid"))
                        
                        rsEntry.CursorLocation = 3
                        rsEntry.open strSQL, cnMBPS
                        Set rsEntry.ActiveConnection = Nothing
                end if
                
                if NOT rsEntry.EOF then
                        if rsEntry("wsShow") then
 %>             <script type="text/javascript">
                <%if pMode<>0 then '' not a contract/package %>
                    addRetailItem(<%=request.querystring("prodid")%>, 3, <%if request.querystring("partnerID")="" then response.write "0" else response.write request.querystring("partnerID") end if%>);
                <%else %>
            addToCart(2);
                <%end if %>
                <% if request.querystring("partnerID")<>0 then %>
                        document.location.replace("/Pages/OnlineStore.aspx?partnerID=" + request.querystring("partnerID"));
                <% end if %>
                </script>
        <%              end if          
                end if
        end if %>

</form>
</div>
<!-- #include file="inc_login_content.asp" -->
<!-- #include file="inc_fb_confirm_login_lb.asp" -->
<% pageEnd %>
<!-- #include file="post.asp" -->

<%

function checkPrevPurch(str)

    dim introPrevPurch
    dim alertStr
	dim canPurchase, canPurchaseStr, jsonParams

    introPrevPurch = false
    if session("mVarUserID")<>"" then
        strSQL = "SELECT tblShoppingCart.ProductID, PRODUCTS.Description FROM tblShoppingCart INNER JOIN PRODUCTS ON tblShoppingCart.ProductID = PRODUCTS.ProductID WHERE (SessionID = N'" & sqlClean(getSessionGUID()) & "') AND ((tblShoppingCart.ProductID IN (SELECT [PAYMENT DATA].ProductID FROM [PAYMENT DATA] INNER JOIN PRODUCTS ON [PAYMENT DATA].ProductID = PRODUCTS.ProductID WHERE (tblShoppingCart.SaleID IS NULL AND PRODUCTS.Introductory = 1) AND ([PAYMENT DATA].Returned = 0) AND ([PAYMENT DATA].ClientID = " & sqlClean(session("mVarUserID")) & ")))  OR (PRODUCTS.IntroNewClient = 1 AND PRODUCTS.Introductory =1 ))"
        'response.write debugSQL(strSQL, "checkPrevPurch() SQL")
		'loginfo "###########checkPrevPurch() ##" & strSQL
        'response.end
        rsEntry.CursorLocation = 3
        rsEntry.open strSQL, cnWS
        Set rsEntry.ActiveConnection = Nothing
        if NOT rsEntry.EOF then
                do while NOT rsEntry.EOF
				canPurchase = true

				if session("mVarUserID") <> "0" then
					canPurchase = false

					set jsonParams = JSON.parse("{}")
					jsonParams.set "ClientID", session("mVarUserID")
					jsonParams.set "ItemID", rsEntry("ProductID")
					canPurchaseStr = CallMethodWithJSON("mb.Core.BLL.Service.CanPurchase",jsonParams)
					if canPurchaseStr = "" then
						canPurchase = true
					end if
				end if

				if NOT canPurchase then
					alertStr = ReplaceInPhrase(phraseDictionary, "Previntropurchase","<PRODUCTNAME>", rsEntry("Description"))
						%><script type="text/javascript" >                    alert("<%=alertStr %>");</script> <%
					strSQL = "DELETE FROM tblShoppingCart WHERE (SessionID = N'" & getSessionGUID() & "') AND ProductID=" & rsEntry("ProductID")
					cnWS.execute strSQL
				end if
                    rsEntry.MoveNext
            loop
        end if
        rsEntry.close
    end if
    
    checkPrevPurch=true
end function
 %>
