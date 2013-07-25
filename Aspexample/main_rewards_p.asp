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
<%
if not Session("Pass") then
%>
<script type="text/javascript">
alert("Please Log In");
self.close();
</script>
<%
else
%>
<!-- #include file="inc_dbconn.asp" -->
<!-- #include file="adm/inc_crypt.asp" -->
<!-- #include file="adm/inc_chk_ss.asp" -->
<!-- #include file="adm/controls/adm_clt_referral.asp" -->
<%
	if request.form("frmSubmitted")=session("StudioID") then
         dim rsEntry, rsIndex
         set rsEntry = Server.CreateObject("ADODB.Recordset")
		 set rsIndex = Server.CreateObject("ADODB.Recordset")

        'create SQL select query string
        strSQL = "SELECT StudioID, reqAddress, reqCity, reqState, reqZip, reqPhone, reqWorkPhone, reqCellPhone, reqReferredBy, reqBirthday, reqMiddleName, reqEmail, reqEmergContact, reqRSSID, reqHeight, reqBust, reqWaist, reqHip, reqGirth, reqInseam, reqHead, reqShoe, reqTights FROM RequiredFields WHERE StudioID = " & session("studioID")
        rsEntry.CursorLocation = 3
		rsEntry.open strSQL, cnWS
		Set rsEntry.ActiveConnection = Nothing
		Set rsUser = Server.CreateObject("ADODB.Recordset")


		'create insert statement to add a new users starting profile info
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        strSQL = "UPDATE Clients SET RewardsOptIn = 1"
		if rsEntry("reqMiddleName") then
	        strSQL = strSQL & ", MiddleName=N'" & sqlInjectStr(Request.Form("requiredtxtMiddleName")) & "'"
		end if
		
		if rsEntry("reqAddress") then
	        strSQL = strSQL & ", Address=N'" & sqlInjectStr(Request.Form("requiredtxtAddress")) & "'"
	        if checkStudioSetting("tblGenOpts","UseAddressLineTwo") and request.Form("txtAddress2")<>"" then
	        strSQL = strSQL & ", Address2=N'" & sqlInjectStr(Request.Form("txtAddress2")) & "'"
	        end if
		end if

		if rsEntry("reqCity") then
	        strSQL = strSQL & ", City=N'" & sqlInjectStr(Request.Form("requiredtxtCity")) & "'"
		end if

		if rsEntry("reqState") then
	        strSQL = strSQL & ", State=N'" & sqlInjectStr(Request.Form("requiredoptState")) & "'"
		end if
		
		if (rsEntry("reqBirthday") AND request.form("requiredtxtBirthday")<>"") then
			Call SetLocale(session("mvarLocaleStr"))
				if isDate(request.form("requiredtxtBirthday")) then
					tmpDate = CDATE(request.form("requiredtxtBirthday"))
					Call SetLocale("en-us")
					strSQL = strSQL & ", Birthdate=" & DateSep & tmpDate & DateSep
				end if
			Call SetLocale("en-us")
		end if
		
		' Client Measurements
		if (rsEntry("reqHeight") AND request.Form("requiredtxtHeight") <> "") then
		    strSQL = strSQL & ", Height=" & Replace(Trim(request.Form("requiredtxtHeight")),"'","''")
		end if
		if (rsEntry("reqBust") AND request.Form("requiredtxtBust") <> "") then
		    strSQL = strSQL & ", Bust=" & Replace(Trim(request.Form("requiredtxtBust")),"'","''")
		end if
		if (rsEntry("reqWaist") AND request.Form("requiredtxtWaist") <> "") then
		    strSQL = strSQL & ", Waist=" & Replace(Trim(request.Form("requiredtxtWaist")),"'","''")
		end if
		if (rsEntry("reqHip") AND request.Form("requiredtxtHip") <> "") then
		    strSQL = strSQL & ", Hip=" & Replace(Trim(request.Form("requiredtxtHip")),"'","''")
		end if
		if (rsEntry("reqGirth") AND request.Form("requiredtxtGirth") <> "") then
		    strSQL = strSQL & ", Girth=" & Replace(Trim(request.Form("requiredtxtGirth")),"'","''")
		end if
		if (rsEntry("reqInseam") AND request.Form("requiredtxtInseam") <> "") then
		    strSQL = strSQL & ", Inseam=" & Replace(Trim(request.Form("requiredtxtInseam")),"'","''")
		end if
		if (rsEntry("reqHead") AND request.Form("requiredtxtHead") <> "") then
		    strSQL = strSQL & ", Head=" & Replace(Trim(request.Form("requiredtxtHead")),"'","''")
		end if
		if (rsEntry("reqShoe") AND request.Form("requiredtxtShoe") <> "") then
		    strSQL = strSQL & ", Shoe=" & Replace(Trim(request.Form("requiredtxtShoe")),"'","''")
		end if
		if (rsEntry("reqTights") AND request.Form("requiredtxtTights") <> "") then
		    strSQL = strSQL & ", Tights=" & Replace(Trim(request.Form("requiredtxtTights")),"'","''")
		end if

		if (rsEntry("reqPhone") and request.form("requiredtxtHomePhone")<>"") then
			tmpPhone = Replace(Trim(request.form("requiredtxtHomePhone")),"(","")
			tmpPhone = Replace(tmpPhone,")","")
			tmpPhone = Replace(tmpPhone,"-","")
			tmpPhone = Replace(tmpPhone,".","")
			tmpPhone = Replace(tmpPhone," ","")
			tmpPhone = Replace(tmpPhone,"+","")
			tmpPhone = Left(tmpPhone,20)
			if isNumeric(tmpPhone) then
			 	strSQL = strSQL & ", HomePhone='" & tmpPhone & "'"
			else
			 	strSQL = strSQL & ", HomePhone=null"
			end if
		end if

		if (rsEntry("reqWorkPhone") and request.form("requiredtxtWorkPhone")<>"") then
			tmpPhone = Replace(Trim(request.form("requiredtxtWorkPhone")),"(","")
			tmpPhone = Replace(tmpPhone,")","")
			tmpPhone = Replace(tmpPhone,"-","")
			tmpPhone = Replace(tmpPhone,".","")
			tmpPhone = Replace(tmpPhone," ","")
			tmpPhone = Replace(tmpPhone,"+","")
			tmpPhone = Left(tmpPhone,20)
			if isNumeric(tmpPhone) then
			 	strSQL = strSQL & ", WorkPhone='" & tmpPhone & "'"
			else
			 	strSQL = strSQL & ", WorkPhone=null"
			end if
		end if

		if (rsEntry("reqCellPhone") and request.form("requiredtxtMobilePhone")<>"") then
			tmpPhone = Replace(Trim(request.form("requiredtxtMobilePhone")),"(","")
			tmpPhone = Replace(tmpPhone,")","")
			tmpPhone = Replace(tmpPhone,"-","")
			tmpPhone = Replace(tmpPhone,".","")
			tmpPhone = Replace(tmpPhone," ","")
			tmpPhone = Replace(tmpPhone,"+","")
			tmpPhone = Left(tmpPhone,20)
			if isNumeric(tmpPhone) then
			 	strSQL = strSQL & ", CellPhone='" & tmpPhone & "'"
			else
			 	strSQL = strSQL & ", CellPhone=null"
			end if
		end if


		if rsEntry("reqZip") then
	        strSQL = strSQL & ", postalcode=N'" & sqlInjectStr(Request.Form("requiredtxtZip")) & "'"
		end if
		if rsEntry("reqEmail") then
	        strSQL = strSQL & ", EmailName=N'" & sqlInjectStr(Request.Form("requiredtxtEmail")) & "'"

			if Request.Form("optSendReminders") = "on" then
    			strSQL = strSQL & ", SendMeReminders=1"
			else
    			strSQL = strSQL & ", SendMeReminders=0"
			end if
		end if
		
		if rsEntry("reqEmergContact") then
			if request.form("requiredtxtEmergContact")<>"" then
				strSQL = strSQL & ", EmergContact=N'" & sqlInjectStr(request.form("requiredtxtEmergContact")) & "'"
			end if
			if request.form("requiredtxtEmergRelationship")<>"" then
				strSQL = strSQL & ", EmergRela=N'" & sqlInjectStr(request.form("requiredtxtEmergRelationship")) & "'"
			end if
			if request.form("txtEmergEmail")<>"" then
				strSQL = strSQL & ", EmergEmail=N'" & sqlInjectStr(request.form("txtEmergEmail")) & "'"
			end if

			if TRIM(request.form("requiredtxtEmergPhone"))<>"" then
				tmpPhone = Replace(Trim(request.form("requiredtxtEmergPhone")),"(","")
				tmpPhone = Replace(tmpPhone,")","")
				tmpPhone = Replace(tmpPhone,"-","")
				tmpPhone = Replace(tmpPhone,".","")
				tmpPhone = Replace(tmpPhone," ","")
				tmpPhone = Replace(tmpPhone,"+","")
			
				if isNumeric(tmpPhone) then
					strSQL = strSQL & ", EmergPhone='" & tmpPhone & "'"
				else
					strSQL = strSQL & ", EmergPhone=null"
				end if
			else
				strSQL = strSQL & ", EmergPhone=null"
			end if
		end if
		
		dim referralStr : referralStr = GetReferredBy(request.form("optReferralType"), request.form("txtOtherReferral"))
		dim referrerId : referrerId = GetRefererIdByType(request.form("optReferralType"))
		
		if referralStr<>"" then
			strSQL = strSQL & ", ReferredBy=N'" & sqlInjectStr(referralStr) & "'"
		else
			strSQL = strSQL & ", ReferredBy=null"
		end if
		if referrerId<>"" then
			strSQL = strSQL & ", ReferrerID=N'" & referrerId & "'"
		else
			strSQL = strSQL & ", ReferrerID=null"
		end if
		
			
        strSQL = strSQL & " WHERE ClientId=" & Session("mvarUserId")
	'response.write debugSQL(strSQL, "SQL")
        cnWS.Execute strSQL 
		'' set client indexes
			
		
		dim indexFormName : indexFormName = ""
		dim i : i=0
		'' delete index data for the ones I am updating ONLY (Active & Shown in Consumer Mode)
		strSQL = "DELETE FROM tblClientIndexData WHERE (ClientIndexValueID IN (SELECT tblClientIndexData_1.ClientIndexValueID FROM tblClientIndexData AS tblClientIndexData_1 INNER JOIN tblClientIndexValue ON tblClientIndexData_1.ClientIndexValueID = tblClientIndexValue.ClientIndexValueID INNER JOIN tblClientIndex ON tblClientIndexValue.ClientIndexID = tblClientIndex.ClientIndexID WHERE (tblClientIndexData_1.ClientID = " & Session("mvarUserId") & ") AND (tblClientIndex.Active = 1) AND (tblClientIndex.ConsumerMode = 1))) AND (ClientID = " & Session("mvarUserId") & ")"
		cnWS.execute strSQL
		
		strSQL = "SELECT ClientIndexID, ClientIndexName, Required FROM tblClientIndex WHERE (Active = 1 AND ConsumerMode = 1 AND Required=1) ORDER BY SortOrderID, ClientIndexName"
		rsIndex.CursorLocation = 3
		rsIndex.open strSQL, cnWS
		Set rsIndex.ActiveConnection = Nothing
		do while NOT rsIndex.EOF
		i = i+1
			indexFormName = "requiredoptClientIndex"&i
			
			if request.form(indexFormName)<>"0" AND isNum(request.form(indexFormName)) then
				strSQL = "INSERT INTO tblClientIndexData (ClientIndexValueID, ClientID) VALUES (" & request.form(indexFormName) & "," & Session("mvarUserId") & ")"
				'response.write debugSQL(strSQL, "SQL")
				cnWS.execute strSQL
			end if
			rsIndex.MoveNext
		loop
		rsIndex.close
        
		if request.form("requiredtxtEmail")<>"" AND request.form("requiredtxtEmail")<>request.form("prevEmail") then
%>
		<!-- #include file="inc_dbconn_wsMaster.asp" -->
		<!-- #include file="adm/inc_masterclients.asp" -->
<%
			removeMasterClient Session("mvarUserId")
			strSQL = "SELECT ClientID, RSSID, FirstName, LastName, EmailName, LoginName FROM CLIENTS WHERE ClientID=" & Session("mvarUserId")
			rsUser.CursorLocation = 3
			rsUser.open strSQL, cnWS
			Set rsUser.ActiveConnection = Nothing
				addMasterClient rsUser("ClientID"), rsUser("FirstName"), rsUser("LastName"), rsUser("EmailName"), rsUser("LoginName"), session("StudioID"), rsUser("RSSID")
			rsUser.close
		end if	'email entered and changed


		
		
		set rsUser = Nothing

	rsEntry.Close
	Set rsEntry = Nothing
	end if	'frmSubmitted

	response.redirect "main_rewards.asp?frmSubmitted=true"

	
	


	cnWS.close
	
	set cnWS = nothing
end if
end if
%>

