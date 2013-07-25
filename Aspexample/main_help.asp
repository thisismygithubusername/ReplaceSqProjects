<%@ CodePage=65001 %>
<%	' This file is still directed to by home.asp. No easy way to circumvent that without a major
	' overhaul, so we'll just redirect to the proper location
	response.Redirect("/help") 
	response.End	
%>
<!-- #include file="init.asp" -->
<%
'Dim SessionFarm
'set SessionFarm = Server.CreateObject("SessionFarm.SFSession")
response.charset="utf-8"
%>


<%

%>
		<!-- #include file="inc_dbconn.asp" -->
		<!-- #include file="inc_i18n.asp" -->
		<!-- #include file="inc_localization.asp" -->
		<!-- #include file="adm/inc_hotword.asp" -->
<%

dim phraseDictionary
set phraseDictionary = LoadPhrases("ConsumermodehelpPage", 52)

dim loginDictionary
set loginDictionary = LoadPhrases("TopframeloginPage", 13)

dim rsEntry, ss_HideCltForgotPwd
set rsEntry = Server.CreateObject("ADODB.Recordset")

strSQL = "SELECT Studios.StudioURL, tblGenOpts.StudioLinkTab, tblGenOpts.HideCltHelp, tblGenOpts.HideCltForgotPwd, tblGenOpts.UpperCaseTabs, tblAppearance.LogoHeight, tblAppearance.LogoWidth, "&_
		"tblAppearance.topBGColor, tblGenOpts.ClientModeLockTGRemoveBuy, tblGenOpts.CltModeSigup, tblGenOpts.CltModeSigupExisting, tblGenOpts.BBenabled, tblGenOpts.ClientModeLockLoc "&_
		"FROM Studios "&_
		"INNER JOIN tblGenOpts ON Studios.StudioID = tblGenOpts.StudioID "&_
		"INNER JOIN tblAppearance ON Studios.StudioID = tblAppearance.StudioID "&_
		"INNER JOIN tblApptOpts ON Studios.StudioID = tblApptOpts.StudioID "&_
		"WHERE (Studios.StudioID = " & session("StudioID") & ")"
rsEntry.CursorLocation = 3
rsEntry.open strSQL, cnWS
Set rsEntry.ActiveConnection = Nothing
ss_HideCltForgotPwd = rsEntry("HideCltForgotPwd")
rsEntry.close
%>

<!-- #include file="frame_bottom.asp" -->
<!-- #include file="pre.asp" -->
<!-- #include file="inc_date_ctrl.asp" -->

<%= js(array("mb")) %>
<%= css(array("inc_sub_links")) %>
<style type="text/css">
    body .section
    {
        color: #555;
    }
    body .section b
    {
        font-size: 1.2em;
    }
    .frgt-pwd-ctr
    {
    	float: right; width:300px; text-align:right; padding-top: 25px;
    }
    .frgt-pwd-ctr a
    {
    	text-decoration: none;
    }
    .frgt-pwd-ctr a:hover
    {
    	text-decoration: underline;
    }
</style>

<!-- #include file="inc_cm_header_bar.asp" -->
<% ShowCMHeader %> 
<% pageStart %>
<div class="frgt-pwd-ctr">
	<% if NOT ss_HideCltForgotPwd and not session("pass") then %>
		<a href="/PasswordRecovery/"><%=DisplayPhrase(loginDictionary,"Forgotyourlogin")%></a>
	<% end if %>
</div>

<h1><%=xssStr(allHotWords(24))%></h1>
<div class="section" style="width: 600px;">
	<%=phraseDictionary("Helpcontent")%>
</div>
<% pageEnd %>
<!-- #include file="post.asp" -->
<%

%>
