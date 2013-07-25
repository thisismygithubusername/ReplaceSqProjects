<%@ codepage="65001" %>
<!-- #include file="init.asp" -->
<!-- #include file="json/JSON.asp" -->
<%
response.charset="utf-8"
%>
<!-- #include file="inc_internet_guest.asp" -->
<%
if NOT Session("Pass") then
	response.redirect "su1.asp" & "?" & request.ServerVariables("QUERY_STRING")
end if
%>
<!-- #include file="inc_i18n.asp" -->
<%
if session("CR_Memberships") <> 0 then
%>
<!-- #include file="inc_dbconn_regions.asp" -->
<!-- #include file="inc_dbconn_wsMaster.asp" -->
<!-- #include file="adm/inc_masterclients_util.asp" -->
<%
end if
%>
<!-- #include file="adm/inc_acct_balance.asp" -->
<!-- #include file="adm/inc_crypt.asp" -->
<!-- #include file="adm/inc_hotword.asp" -->
<!-- #include file="inc_tinymcesetup.asp" -->
<%
session("TabID") = "2"


Dim phraseDictionary
set phraseDictionary = LoadPhrases("ConsumermodemyinfomyprofilePage", 21)


%>
<script type="text/javascript">
    var CMProfilePagePhrases =  
    {
        PleaseCompleteRequired: function() {
            return '<%=DisplayPhraseJS(phraseDictionary,"Pleasecompleterequired")%>';
        },
        EmailNotValid: function () {
            return '<%=DisplayPhraseJS(phraseDictionary, "Emailnotvalid")%>';
        },
        PasswordRequiresNumber: function () {
            return '<%=DisplayPhraseJS(phraseDictionary, "Passwordrequiresnumber")%>';
        },
        PasswordRequiresLetter: function () {
            return '<%=DisplayPhraseJS(phraseDictionary, "Passwordrequiresletter")%>';
        },
        PasswordNeeds6Chars: function () {
            return '<%=DisplayPhraseJS(phraseDictionary, "Passwordneeds6chars")%>';
        },
        PasswordNoSpaces: function () {
            return '<%=DisplayPhraseJS(phraseDictionary, "Passwordsnospaces")%>';
        },
        InvalidCC: function () {
            return '<%=DisplayPhraseJS(phraseDictionary, "Invalidcc")%>';
        },
    };
</script>

<!-- #include file="pre.asp" -->
	<%= js(array("mb")) %>
	<!-- #include file="frame_bottom.asp" -->
	<%= js(array("MBS")) %>
	<!-- begin client alerts -->
	<%
	'client alert context vars
	focusFrmElement = ""
	cltAlertList = setClientAlertsList(session("mvarUserID"))
	%>
	<!-- #include file="inc_ajax.asp" -->
	<!-- #include file="adm/inc_alert_js.asp" -->
	<!-- end client alerts  -->
	<%= js(array("VCC2","main_info","jquery.placeholder")) %>

<%
'localization for query datepicker....
dim langCode : langCode = Left(Session("mvarLocaleStr"),2)
if langCode <> "en" then

	dim fso : set fso = Server.CreateObject("Scripting.FileSystemObject")
	dim chkPath : chkPath = Server.MapPath("..\work\scripts\plugins\jquery.ui.datepicker-" & langCode & ".js")

	if fso.FileExists(chkPath) then
		response.write js(array("plugins/jquery.ui.datepicker-" & langCode))
	end if
	set fso=nothing
end if

'build some JS functions for validation!!!

strCCSQL = "SELECT ccAmericanExpress, ccDiscover, ccMasterCard, ccVisa FROM tblCCOpts WHERE StudioID = " & Session("StudioID")
dim rsCCTypes
set rsCCTypes = Server.CreateObject("ADODB.RecordSet")

rsCCTypes.CursorLocation = 3
rsCCTypes.Open strCCSQL, cnWS
Set rsCCTypes.ActiveConnection = Nothing

' Create the accepted types string
dim acceptedTypes : acceptedTypes = ""
dim sep : sep = ""
if CBOOL(rsCCTypes("ccAmericanExpress")) then
	acceptedTypes = acceptedTypes & "American Express"
	sep = ", "
end if
if CBOOL(rsCCTypes("ccVisa")) then
	acceptedTypes = acceptedTypes & sep & "Visa"
	sep = ", "
end if
if CBOOL(rsCCTypes("ccMasterCard")) then
	acceptedTypes = acceptedTypes & sep & "MasterCard"
	sep = ", "
end if
if CBOOL(rsCCTypes("ccDiscover")) then
	acceptedTypes = acceptedTypes & sep & "Discover"
end if

%>
	<script type="text/javascript">
	$(document).ready(function () {

		$.datepicker.setDefaults({ dateFormat: '<%=FmtDatePickerDateCode%>' });
		
		
		var acceptedCCTypes = {"ccAmericanExpress"	: <%=lcase(rsCCTypes("ccAmericanExpress"))%> ,
								"ccVisa"			: <%=lcase(rsCCTypes("ccVisa")) %>,
								"ccMasterCard"		: <%=lcase(rsCCTypes("ccMasterCard"))%>,
								"ccDiscover"		: <%=lcase(rsCCTypes("ccDiscover")) %>};
		
		//set error messages on "checkRequired" extension of jquery object
		//this is defined in main_info.js
		$(document).checkRequired(
			{
				PleaseCompleteRequired	: <%= "'" & DisplayPhraseJS(phraseDictionary,"Pleasecompleterequired") & "'"%>,
				PasswordRequiresNumber	: <%= "'" & DisplayPhraseJS(phraseDictionary,"Passwordrequiresnumber") & "'" %>,
				PasswordRequiresLetter	: <%= "'" & DisplayPhraseJS(phraseDictionary,"Passwordrequiresletter") & "'" %>,
				PasswordNeeds6Chars			: <%= "'" & DisplayPhraseJS(phraseDictionary,"Passwordneeds6chars") & "'" %>,
				PasswordsNoSpaces				: <%= "'" & DisplayPhraseJS(phraseDictionary,"Passwordsnospaces") & "'" %>,
				WeOnlyAccept						: <%= "'" & DisplayPhraseJS(phraseDictionary,"Weonlyaccept") & "'" %>,
				InvalidCC								: <%= "'" & DisplayPhraseJS(phraseDictionary,"Invalidcc") & "'" %>,
				AcceptedCCTypesList			: <%= "'" & acceptedTypes & "'" %>
			},
			{
				chkCCTypeAccepted : function(ccNumber)
				{
					return CheckCCNum(ccNumber, acceptedCCTypes, false);
				}
			}
		);
	});




   function DisplayStateProvince() {
        return '<%=jsEsc(allHotWords(244))%>'
   }


	</script>
<% 'end localization for jquery datepicker and password/cc validation 
rsCCTypes.close
%>
	<%= css(array("inc_sub_links","main_info")) %>
	<style type="text/css">
		.clearfix
		{
			content: ".";
			display: block;
			height: 0;
			clear: both;
			visibility: hidden;
		}
	</style>
	<!--[if IE 7]>
<style type="text/css">
.clearfix { display:inline-block; zoom:1; }
</style>
<![endif]-->
	<input type="hidden" id="thisStudioID" value="<%=Session("StudioID")%>" />
	<input type="hidden" id="thisClientID" value="<%=Session("mvarUserId")%>" />
	<!-- #include file="adm/inc_alert_content.asp" -->
	<!-- #include file="inc_sub_links.asp" -->
	<!-- #include file="inc_render_main_info.asp" -->
	<% ShowSubTabLinks ("main_info") %>
	<% pageStart %>
	<div id="myInfoContainer">
<%

' THIS FUNCTION DOES SO MUCH MAGIC!!!!
' It loads rsUser, which will be used throughout all the info screens 
populateRsUser
if rsUser.EOF then
	response.redirect "su1.asp"
end if

dim isPending : isPending = rsUser("wspending")

if isPending then
	Response.Write "<br /><br />" & DisplayPhrase(phraseDictionary, "Unabletoprovide") &".<br />"&DisplayPhrase(phraseDictionary, "Waitingforsecurityverification") &"."
else
%>
		<h1>
			<%=DisplayPhrase(phraseDictionary, "Myinfo")%>
		</h1>
		<div class="infoPane" id="PersonalPane">
			<% renderHeader DisplayPhrase(phraseDictionary, "Personal"), "Personal" %>
			<div class="infoSubContainer">
				<% renderPersonalInfo		%>
				<% renderPersonalInfoEdit	%>
				<% 'response.write("<div class=""spacer cancelSpacer"">.</div>")%>
				<% 'renderCancelEditButton("CancelPersonalInfoEdit") %>
			</div>
		</div>
<%
	if Session("mvarMIDs")<>0 then
%>
		<div class="infoPane" id="BillingPane">
			<% renderHeader allHotWords(89), "Billing"	%>
			<div class="infoSubContainer">
				<% renderBillingInfo		%>
				<% renderBillingInfoEdit	%>
				<% 'response.write("<div class=""spacer cancelSpacer"">.</div>")%>
				<% 'renderCancelEditButton("CancelBillingInfoEdit") %>
			</div>
		</div>
<%
	end if
	if siteHasActiveRelationships OR userHasRelationships then
%>
		<div class="infoPane" id="FamilyPane">
			<% renderHeader DisplayPhrase(signUpDictionary, "Familymemberinformation") , "Family"	%>
			<div class="infoSubContainer">
				<% renderFamilyInfo		%>
				<% renderFamilyInfoEdit	%>
			</div>
		</div>
		<div style="clear: both">
		</div>
		<div id="requiredFieldExplanation">
			*<%=xssStr(allHotWords(791))%>
		</div>
<%
	end if
end if
%>
	</div> <%'id="myinfocontainer" %>
	<% pageEnd %>
<!-- #include file="post.asp" -->
 
