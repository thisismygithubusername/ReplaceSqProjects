<%@ CodePage=65001 %>
<!-- #include file="init.asp" -->
<%
'Dim SessionFarm
'set SessionFarm = Server.CreateObject("SessionFarm.SFSession")
response.charset="utf-8"
%>

		<!-- #include file="inc_internet_guest.asp" -->
		<!-- #include file="inc_i18n.asp" -->
		<!-- #include file="adm/inc_chk_ss.asp" -->
        <% if session("CR_Memberships") <> 0 then %>
            <!-- #include file="inc_dbconn_regions.asp" -->
            <!-- #include file="inc_dbconn_wsMaster.asp" -->
            <!-- #include file="adm/inc_masterclients_util.asp" -->
        <% end if %>
		<!-- #include file="adm/inc_chk_membership.asp" -->
		<!-- #include file="adm/inc_acct_balance.asp" -->
		<!-- #include file="adm/inc_crypt.asp" -->
		<!-- #include file="inc_localization.asp" --> 
		<!-- #include file="adm/inc_hotword.asp" -->
		<!-- #include file="adm/controls/adm_clt_referral.asp" -->
<%
session("TabID") = 100
session("pageID")="_rewards"


dim phraseDictionary
set phraseDictionary = LoadPhrases("ConsumermoderewardssignupPage", 75)

dim mainInfoDictionary
set mainInfoDictionary = LoadPhrases("ConsumermodemyinfomyprofilePage", 21)

dim reqReferral
reqReferral = false

if not Session("Pass") then
	response.redirect "su1.asp"
else

	dim rsUser, rsEntry, rsEntry2, rsValue, rsIndex, AlertRequiredIndexBiz, AlertIndexIds
	set rsUser = Server.CreateObject("ADODB.Recordset")
	set rsEntry = Server.CreateObject("ADODB.Recordset")
	set rsEntry2 = Server.CreateObject("ADODB.Recordset")
	set rsIndex = Server.CreateObject("ADODB.Recordset")
	set rsValue = Server.CreateObject("ADODB.Recordset")

	Dim ss_UseStates, ss_contactEmail, ss_StoreGenderPreferences
	strSQL = "SELECT tblGenOpts.UseStates, tblGenOpts.ClientContactEmail, tblApptOpts.StoreGenderPreferences FROM tblGenOpts INNER JOIN tblApptOpts ON tblGenOpts.StudioID = tblApptOpts.StudioID WHERE tblGenOpts.StudioID=" & session("StudioID")
    rsUser.CursorLocation = 3
	rsUser.open strSQL, cnWS
	Set rsUser.ActiveConnection = Nothing
	if NOT rsUser.EOF then
		ss_UseStates = rsUser("UseStates")
		ss_contactEmail = rsUser("ClientContactEmail")
		ss_StoreGenderPreferences = rsUser("StoreGenderPreferences")
	else
		ss_UseStates = true
		ss_contactEmail = false
		ss_StoreGenderPreferences = false
	end if
	rsUser.close
%>
<!-- #include file="pre.asp" -->
<!-- #include file="frame_bottom.asp" -->
<%= js(array("MBS", "controls")) %>
<!-- #include file="inc_ajax.asp" -->
<!-- #include file="inc_date_ctrl.asp" -->

<%= js(array("mb", "main_rewards", "calendar" & dateFormatCode, "VCC2")) %>
<style type="text/css">
    table#rewards td { color: #555;}
    .section { width: 600px; }
</style>
<!-- #include file="inc_cm_header_bar.asp" -->
<% ShowCMHeader %> 
<% pageStart %>
<form action="main_rewards_p.asp" method="post" name="frmRewardsSignup">
<input name="frmSubmitted" type="hidden" value="<%=session("studioID")%>">

<h1><%= DisplayPhrase(phraseDictionary,"Rewards") %></h1> 
<div class="section">
<table id="rewards" height="100%" width="550" cellspacing="0">
  <tr> 
      
      <td class="" valign="top" height="100%" width="100%">  
        <table cellspacing="0" width="100%" height="100%" class="">
          <tr> 
            <td valign="top" class="" align="left">  
              <table class="" width="100%" cellspacing="0">
                <tr > 
                  <td class="" colspan="2" valign="top">

<%
    'create SQL select query string
	strSQL = "SELECT * FROM CLIENTS WHERE (CLIENTS.ClientID=" & session("mvarUserID") & ")"
    rsUser.CursorLocation = 3
	rsUser.open strSQL, cnWS
	Set rsUser.ActiveConnection = Nothing
	if rsUser.EOF then
		response.redirect "su1.asp"
	end if
	
	if rsUser("wspending") then
		Response.Write "<br /><br />" & DisplayPhrase(mainInfoDictionary,"Unabletoprovide") &".<br />"& DisplayPhrase(mainInfoDictionary,"Waitingforsecurityverification") &"."
	else
%>
                <table class="" width="100%" cellspacing="0">
				
				<%
				
				strSQL = "SELECT StudioName FROM STUDIOS WHERE STUDIOS.StudioID = " & session("studioID") 
				'response.write debugSQL(strSQL, "SQL")
				rsEntry.CursorLocation = 3
				rsEntry.open strSQL, cnWS
				Set rsEntry.ActiveConnection = Nothing
%>
				<!-- <table  cellspacing="0" class="mainText center-ch" width="70%"> -->
					
<%				if NOT rsEntry.EOF then
				
					if request.form("frmSubmitted")=session("StudioID") or request.querystring("frmSubmitted")="true" or rsUser("RewardsOptIn") then%>
				<tr>
					<td colspan="2"><%=DisplayPhrase(phraseDictionary,"Rewardsconfirmmessage")%><br /><br /></td>
				</tr>
				<%	else
						ReplaceInPhrase phraseDictionary, "Rewardswelcomemessage", "<STUDIONAME>", rsEntry("StudioName") %>
				<tr>
					<td colspan="2"><%=DisplayPhrase(phraseDictionary,"Rewardswelcomemessage")%><br /><br /></td>
				</tr>
<%					end if
				end if
				rsEntry.Close 
				
				if request.form("frmSubmitted")="" and request.querystring("frmSubmitted")="" and rsUser("RewardsOptIn")=false then
				
					strSQL = "SELECT RequiredFields.StudioID, RequiredFields.reqAddress, RequiredFields.reqCity, RequiredFields.reqState, RequiredFields.reqZip, RequiredFields.reqPhone, RequiredFields.reqWorkPhone, RequiredFields.reqCellPhone, RequiredFields.reqReferredBy, RequiredFields.reqBirthday, RequiredFields.reqMiddleName, RequiredFields.reqEmail, RequiredFields.reqEmergContact, RequiredFields.reqRSSID, "
					strSQL = strSQL & "RequiredFields.reqHeight, RequiredFields.reqBust, RequiredFields.reqWaist, RequiredFields.reqHip, RequiredFields.reqGirth, RequiredFields.reqInseam, RequiredFields.reqHead, RequiredFields.reqShoe, RequiredFields.reqTights, "
					strSQL = strSQL & "CLIENTS.RSSID, CLIENTS.ClientID, CLIENTS.Address,  CLIENTS.Address2, CLIENTS.LastName,  CLIENTS.FirstName, CLIENTS.City, CLIENTS.State, CLIENTS.PostalCode, CLIENTS.HomePhone, CLIENTS.WorkPhone, CLIENTS.CellPhone, CLIENTS.ReferredBy, clients.referrerId, CLIENTS.Birthdate, CLIENTS.MiddleName, CLIENTS.EmailName, CLIENTS.RefusedEmail, CLIENTS.EmergContact, CLIENTS.EmergRela, CLIENTS.EmergPhone, CLIENTS.EmergEmail, CLIENTS.SendMeReminders FROM RequiredFields CROSS JOIN CLIENTS "
					strSQL = strSQL & "WHERE RequiredFields.studioID = " & session("studioID") & " AND CLIENTS.ClientID=" & session("mvarUserID")
				
					rsEntry.CursorLocation = 3
					rsEntry.open strSQL, cnWS
					Set rsEntry.ActiveConnection = Nothing
%>
				<!-- <table  cellspacing="0" class="mainText center-ch" width="70%"> -->
					
<%
				if NOT rsEntry.EOF then
				  reqReferral = rsEntry("reqReferredBy")
				%>
				<tr>
						<td width="35%"><strong><%=xssStr(allHotWords(80))%>:&nbsp;</strong></td>
						<td width="65%"><input type="text" name="txt_First" id="txt_First" maxlength="50" size="22" value="<%=rsEntry("FirstName")%>" disabled></td>
					</tr>
					
				<tr>
						<td width="35%"><strong><%=xssStr(allHotWords(81))%>:&nbsp;</strong></td>
						<td width="65%"><input type="text" name="txt_Last" id="txt_Last" maxlength="50" size="22" value="<%=rsEntry("LastName")%>" disabled></td>
					</tr>
					<% if rsEntry("reqMiddleName") then%>
					<tr>
						<td width="35%"><strong><!-- JM-51_2725 --><%=xssStr(allHotWords(204))%>:&nbsp;</strong></td>
						<td width="65%"><input type="text" name="requiredtxtMiddleName" id="txt_MiddleName" maxlength="50" size="22" value="<%= rsEntry("MiddleName") %>"></td>
					</tr>
					<%end if%>
					<% if rsEntry("reqAddress") then%>
					<tr>
						<td width="35%"><strong><%=xssStr(allHotWords(46))%>:&nbsp;</strong></td>
						<td width="65%"><input type="text" name="requiredtxtAddress" id="txt_Address" maxlength="50" size="22" value="<%= rsEntry("Address") %>"></td>
					</tr>
					<% if checkStudioSetting("tblGenOpts","UseAddressLineTwo") then %>
					<tr>
						<td width="35%"><strong><%=xssStr(allHotWords(46))%>&nbsp;2&nbsp;:&nbsp;</strong></td>
						<td width="65%"><input type="text" name="txtAddress2" id="txt_Address2" maxlength="50" size="22" value="<%= rsEntry("Address2") %>"></td>
					</tr>
					<% end if %>
					<%end if%>
					<% if rsEntry("reqCity") then%>
					<tr>
						<td width="35%"><strong><%=xssStr(allHotWords(47))%>:&nbsp;</strong></td>
						<td width="65%"><input type="text" name="requiredtxtCity" id="txt_City" maxlength="50" size="22" value="<%= rsEntry("City") %>"></td>
					</tr>
					<%end if%>
					<% if rsEntry("reqState") then%>
					<tr>
						<td width="35%"><strong>
					  <%if session("mvarLocaleStr")="en-gb" then%><%=xssStr(allHotWords(207))%>: <%else%><%=xssStr(allHotWords(48))%>: <%end if%>&nbsp;</strong></td>
						<td width="65%">
                            <select name="requiredoptState" id=="opt_State">
							  	<option value="">Select State/Prov</option>
<%
						strSQL = "SELECT CountryCode, StateProvCode, StateProvName FROM tblWrldStateProv WHERE (CountryCode = N'" & checkStudioSetting("Studios", "countryCode") & "') ORDER BY StateProvName"
						rsEntry2.CursorLocation = 3
						rsEntry2.open strSQL, cnWS
						Set rsEntry2.ActiveConnection = Nothing
						do while not rsEntry2.EOF
%>						  
							  	<option value="<%=rsEntry2("StateProvCode")%>" <%if rsEntry("State") = rsEntry2("StateProvCode") then%>selected<%end if%> ><%=UCASE(rsEntry2("StateProvName"))%></option>
<%
						rsEntry2.MoveNext
					loop
					rsEntry2.close
%>
                            </select>
					  </td>
					</tr>
					<%end if%>
					<% if rsEntry("reqZip") then%>					
<tr>
						<td width="35%"><strong><%=xssStr(allHotWords(49))%>:&nbsp;</strong></td>
						<td width="65%"><input type="text" name="requiredtxtZip" id="txt_Zip" maxlength="50" size="22" value="<%= rsEntry("PostalCode") %>"></td>
					</tr>
					<%end if%>
					<% if rsEntry("reqPhone") then%>
					<tr>
						<td width="35%"><strong><%=xssStr(allHotWords(82))%>:&nbsp;</strong></td>
						<td width="65%"><input type="text" name="requiredtxtHomePhone" id="txt_HomePhone" maxlength="50" size="15" value="<%= rsEntry("HomePhone") %>"></td>
					</tr>
					<%end if%>
					<% if rsEntry("reqWorkPhone") then%>
					<tr>
						<td width="35%"><strong><%=xssStr(allHotWords(84))%>:&nbsp;</strong></td>
						<td width="65%"><input type="text" name="requiredtxtWorkPhone" id="txt_WorkPhone" maxlength="50" size="15" value="<%= rsEntry("WorkPhone") %>"></td>
					</tr>
					<%end if%>
					<% if rsEntry("reqCellPhone") then%>
					<tr>
						<td width="35%"><strong><%=xssStr(allHotWords(83))%>: </strong>&nbsp;</td>
						<td width="65%"><input type="text" name="requiredtxtMobilePhone" id="txt_CellPhone" maxlength="50" size="15" value="<%= rsEntry("CellPhone") %>"></td>
					</tr>
					<%end if%>
					<% if rsEntry("reqEmail") then%>
					<input type="hidden" name="prevEmail" value="<%=rsEntry("EmailName")%>">
					<tr>
						<td width="35%"><strong><%=xssStr(allHotWords(39))%>:</strong></td>
					  <td width="65%"><input type="text" name="requiredtxtEmail" id="txt_Email" maxlength="200" size="36" value="<%= rsEntry("EmailName") %>" onBlur="jsCheckEmail(this);">
					</tr>
					<tr>
						<td width="35%"><strong><%= DisplayPhrase(phraseDictionary,"Sendmeemailconfirmations") %>:</strong></td>
						<td width="65%"> <input type="checkbox" name="optSendReminders" id="opt_EmailOptIn" <%if rsEntry("SendMeReminders") then%>checked><%end if%>&nbsp;<b></b></td>
					</tr>
					<%end if%>
					
					<% if rsEntry("reqBirthday") then%>
					<tr>
						<td width="35%"><strong><%=xssStr(allHotWords(124))%>:&nbsp;</strong>&nbsp;</td>
						<td width="65%"><input type="text" name="requiredtxtBirthday" id="requiredtxtBirthday" maxlength="50" size="15" value="<%= rsEntry("Birthdate") %>" class="date">
							<script type="text/javascript">
								var calBD = new tcal({'formname':'frmRewardsSignup', 'controlname':'requiredtxtBirthday'});
								calBD.a_tpl.yearscroll = true;
							</script>
					  </td>
					</tr>
					
<%					
					end if
					if rsEntry("reqEmergContact") then
%>
					<tr>
						<td width="35%"><strong><%= DisplayPhrase(phraseDictionary,"Emergencycontactname") %>:&nbsp;</strong></td>
						<td width="65%"><input type="text" name="requiredtxtEmergContact" id="txt_EmergContact" maxlength="50" size="22" value="<%= rsEntry("EmergContact") %>"></td>
					</tr>

					<tr>
						<td width="35%"><strong><%= DisplayPhrase(phraseDictionary,"Emergencycontactrelationship") %>:&nbsp;</strong></td>
						<td width="65%"><input type="text" name="requiredtxtEmergRelationship" id="txt_EmergRela" maxlength="50" size="22" value="<%= rsEntry("EmergRela") %>"></td>
					</tr>

					<tr>
						<td width="35%"><strong><%= DisplayPhrase(phraseDictionary,"Emergencycontactphone") %>:&nbsp;</strong></td>
						<td width="65%"><input type="text" name="requiredtxtEmergPhone" id="txt_EmergPhone" maxlength="15" size="22" value="<%= rsEntry("EmergPhone") %>"></td>
					</tr>


					<tr>
						<td width="35%"><strong><%= DisplayPhrase(phraseDictionary,"Emergencycontactemail") %>:&nbsp;</strong></td>
						<td width="65%"><input type="text" name="txtEmergEmail" id="txt_EmergEmail" maxlength="255" size="22" value="<%= rsEntry("EmergEmail") %>"></td>
					</tr>

<%				end if

                ' Client Measurements
                ' I'm leaving this out because it's silly. 
                ' Requiring measurements should not be part of the rewards program.
                
                
				if reqReferral then
%>
					<tr>
						<td width="35%"><strong><%=GetReferralLabel(xssStr(allHotWords(127))) %></strong></td>
						<td width="65%">
              <%= GetReferralSelect ("", rsEntry("ReferredBy"))%>
					  </td>
					</tr>
<%					
					end if %>
					
					
<%
			strSQL = "SELECT ClientIndexID, ClientIndexName, Required FROM tblClientIndex WHERE (Active = 1 AND ConsumerMode = 1 AND Required=1) ORDER BY SortOrderID, ClientIndexName"
			rsIndex.CursorLocation = 3
			rsIndex.open strSQL, cnWS
			Set rsIndex.ActiveConnection = Nothing
			dim i
			if NOT rsIndex.EOF then
			i=0
%>
				<tr><td colspan="2"><strong><br /><%= DisplayPhrase(phraseDictionary,"Indexes") %>:<br /></strong> </td></tr>
<%
				do while NOT rsIndex.EOF
				i = i + 1
%>
								<tr>
								  <td nowrap width="1%"><strong><%=i%>.&nbsp;<%=rsIndex("ClientIndexName")%>:</strong>&nbsp;&nbsp;</td>
								  <td>
									<select name="requiredoptClientIndex<%=i%>">
										<option value="0">Not Assigned</option>
<%
					strSQL = "SELECT tblClientIndexValue.ClientIndexValueName, tblClientIndexValue.ClientIndexValueID, CltData.ClientIndexValueID AS Selected FROM tblClientIndexValue LEFT OUTER JOIN (SELECT ClientID, ClientIndexValueID FROM tblClientIndexData WHERE (ClientID = " & session("mvarUserID") & ")) CltData ON tblClientIndexValue.ClientIndexValueID = CltData.ClientIndexValueID WHERE (tblClientIndexValue.ClientIndexID = " & rsIndex("ClientIndexID") & ") AND (tblClientIndexValue.Active = 1) ORDER BY tblClientIndexValue.ClientIndexValueName"
					rsEntry2.CursorLocation = 3
					rsEntry2.open strSQL, cnWS
					Set rsEntry2.ActiveConnection = Nothing
					do while NOT rsEntry2.EOF
%>
										<option value="<%=rsEntry2("ClientIndexValueID")%>" <%if NOT isNULL(rsEntry2("Selected")) then response.write "selected" end if%>><%=rsEntry2("ClientIndexValueName")%></option>
<%
						rsEntry2.MoveNext
					loop
					rsEntry2.close
%>
				</select>
								  </td>
								  <td colspan="2">&nbsp;</td>
							    </tr>
<%
					rsIndex.MoveNext
				loop
			end if
			rsIndex.close ' end if eof%>

				
				
<%
			end if ' end reqFields AND index EOF
			rsEntry.close
			
%> 
					
					<tr>
					<td></td>	<td colspan="2" class="" ><br /><br /> <input type="submit"  id="signUpButton" name= "UpdateButton1" value='<%= DisplayPhraseAttr(phraseDictionary,"Signupnow") %>'> 
						<br />
				<br />
					  </td>
					</tr>
					<%end if%>
					</table>
					
					
	 <%end if'frmSubmitted%>				
  
                  </td>
                </tr>
              </table>
            </td>
          </tr>
         
        </table>
    </td>
    </tr>
  </table>
					 
                      </div>
  </form>
<script type="text/javascript">
	$(document).ready(function(){
		referSelectCtrl.ready(function() {});
		<% if reqReferral then %>
		  referSelectCtrl.setIsRequired(true, function(){alert("Please select a Referral.");});
		<% end if %>
		referSelectCtrl.setCustomText({ 'client': '<%=jsEscSingle(session("ClientHW"))%>' });
		$('select[name=optReferralType]').change(function() {
			referSelectCtrl.change();
		});
		//onSubmit="javascript: return validateForm(this);"
		$('form[name=frmRewardsSignup]').submit(function() {
			return (checkrequired(this) && referSelectCtrl.canSubmit());
		});
	});
</script>
<br />

<% pageEnd %>

<!-- #include file="post.asp" -->

<%
end if
%>
