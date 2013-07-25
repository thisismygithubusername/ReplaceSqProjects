<%@ CodePage=65001 %>
<!-- #include file="init.asp" -->
<!-- #include file="json/JSON.asp" -->

<%
'Dim SessionFarm
'set SessionFarm = Server.CreateObject("SessionFarm.SFSession")
response.charset="utf-8"
%>

		<!-- #include file="inc_internet_guest.asp" -->
		<!-- #include file="inc_dbconn.asp" -->
		<!-- #include file="inc_i18n.asp" -->
<%
dim headlineID, desc, title
headlineID = Request.QueryString("id")

%>	<!-- #include file="inc_localization.asp" --> <%
dim phraseDictionary
set phraseDictionary = LoadPhrases("ConsumermodenewsPage", 22)

         dim rsEntry, topMessage, studioID
         set rsEntry = Server.CreateObject("ADODB.Recordset")
        'create SQL select query string
        strSQL = "SELECT tblBB.HeadlineTitle, tblBB.Description FROM tblBB WHERE tblBB.HeadlineID=" & headlineID
        rsEntry.CursorLocation = 3
		rsEntry.open strSQL, cnWS
		Set rsEntry.ActiveConnection = Nothing
    
		if not rsEntry.EOF then
			title = rsEntry("HeadlineTitle")
			desc = rsEntry("Description")
		end if
		rsEntry.close
%>
<!-- #include file="pre.asp" -->
<!-- #include file="frame_bottom.asp" -->

<!-- #include file="inc_date_ctrl.asp" -->
<!-- #include file="inc_cm_header_bar.asp" -->
<%= js(array("mb")) %>
<%= css(array("bootstrap/MbCoreStyles")) %>

<!-- #include file="inc_tinymcesetup.asp" -->
<% pageStart %>
<% showCMHeader %>
 <div class="content" style="padding-top:10px;">
            <h2><%=DisplayPhrase(phraseDictionary,"Newsevents")%></h2>
         
                    <h3><%=xssStr(title)%></h3><br />
                    <br />
                    <%=HtmlPurifyForDisplay(desc)%><br />
</div>
<% pageEnd %>
<!-- #include file="post.asp" -->
