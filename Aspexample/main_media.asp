<%@ CodePage=65001 %>
<!-- #include file="init.asp" -->
<!-- #include file="json/JSON.asp" -->
<%
'Dim SessionFarm
'set SessionFarm = Server.CreateObject("SessionFarm.SFSession")
response.charset="utf-8"
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<!-- #include file="inc_dbconn.asp" -->
<!-- #include file="inc_jquery.asp" -->
<!-- #include file="inc_i18n.asp" -->
<!-- #include file="inc_localization.asp" -->
<!-- #include file="adm/inc_hotword.asp" -->
<!-- #include file="adm/inc_chk_ss.asp" -->
<!-- #include file="inc_tinymcesetup.asp" -->

<%
session("pageID")="_media"
if isNum(request.form("tabID"))then
	session("tabID") = request.form("tabID")
elseif isNum(request.QueryString("tabID")) then
	session("tabID") = request.querystring("tabID")
end if

' Load the page phrases.
dim phraseDictionary
set phraseDictionary = LoadPhrases("ConsumermodemediaPage", 85)

dim ss_CheckActivationDates

ss_CheckActivationDates = checkStudioSetting("tblGenOpts", "CheckActivationDates")

if NOT session("pass") then ' not logged in so require user to login/create new account
   %>
   <!-- #include file="pre.asp" -->
    <!-- #include file="frame_bottom.asp" -->

    

    <script type="text/javascript">
    	document.location.href = "SU1.asp?<%=buildLinkVars(false) %>";
    </script>
   
    <!-- #include file="inc_cm_header_bar.asp" -->
    <% ShowCMHeader %> 
    <% pageStart %>
    <% pageEnd %>
    </body>
    </html>
    
    <%
	response.end
end if ' end if not logged in

%>
	<!-- #include file="inc_internet_guest.asp" -->
	<% if session("CR_Memberships") <> 0 then %>
		<!-- #include file="inc_dbconn_regions.asp" -->
		<!-- #include file="inc_dbconn_wsMaster.asp" -->
		<!-- #include file="adm/inc_masterclients_util.asp" -->
	<% end if %>
	<!-- #include file="adm/inc_chk_membership.asp" -->
	<!-- #include file="adm/inc_acct_balance.asp" -->
	<!-- #include file="adm/inc_crypt.asp" -->
	<!-- #include file="inc_loading.asp" -->
	<!-- #include file="adm/inc_recalc_pd.asp" -->
   <!-- #include file="pre.asp" -->
    <!-- #include file="frame_bottom.asp" -->
    <!-- #include file="inc_date_ctrl.asp" -->
    <!-- #include file="inc_fixed_bar.asp" -->

    <%= js(array("mb", "MBS")) %>
    <!-- begin client alerts -->
    <%
    'client alert context vars
    focusFrmElement = ""
    cltAlertList = setClientAlertsList(session("mvarUserID"))
    %>
	
    <!-- #include file="inc_ajax.asp" -->
    <!-- #include file="adm/inc_alert_js.asp" -->
    <!-- end client alerts  -->

    <%= js(array("calendar" & dateFormatCode, "VCC2", "main_media", "plugins/jquery.selectBox")) %>
    <%= css(array("jquery.selectBox")) %>
    <script type="text/javascript">
	    function subFormClrPNum() {
		    document.search2.pageNum.value = "1";
		    subForm();
	    }
    function showMedia(mediaID) {	    
		$.ajax({
			url: "http<%=addS%>://<%=request.ServerVariables("HTTP_HOST")%>/asp/media_geturl.asp",
			type: "POST",
			data: { "id" : mediaID },
			dataType: 'text',
			success: function(data) {			
				data = data.replace(/&amp;/g, '&'); // RegEx replace all occurances. Needed for goofy Chrome behavior
				popup(data);
			}
		});
    }
    $(document).ready(function () {

        //
        // Enable selectBox control and bind events
        //

        $(".filterList SELECT").selectBox({ fixed: true });

    });
	</script>
<%

	dim rsEntry
	set rsEntry = Server.CreateObject("ADODB.Recordset")
	
	dim newVDID, PmtRefNo, PaymentName, PaymentTG, Clicked, curMediaID, curTG, clientID
	
	if request.form("optTG")<>"" then
		curTG = CINT(request.form("optTG"))
	elseif request.querystring("tg")<>"" then
		curTG = CINT(request.querystring("tg"))
	else
		curTG=0
	end if
	
	' Very particular case- if you are logged in as the internet guest account (clientID 0)
	' and you come to this page, it logs you out above. However, this query depends on a valid
	' ID, so we use a temp ID, just so that this will not bitch.
	if isEmpty(session("mvarUserID")) or isNull(session("mvarUserID")) then
	    clientID = 0
	else 
	    clientID = session("mvarUserID")
	end if
	
	' check to see if there is only 1 media link available (and can pay for it) - if so, autoselect it
	strSQL = "SELECT DISTINCT tblMedia.MediaID "&_
             "FROM ("&_
                 "SELECT ISNULL(tblTypeGroup.TypeGroupID, PD.TypeGroup) AS TypeGroupID, PD.PmtRefNo "&_
                 "FROM tblTGRelate "&_
                 "RIGHT OUTER JOIN tblTypeGroup ON (tblTGRelate.TG1 = tblTypeGroup.TypeGroupID OR tblTGRelate.TG2 = tblTypeGroup.TypeGroupID) "&_
		         "RIGHT OUTER JOIN ("&_
                     "SELECT [PAYMENT DATA].TypeGroup, [PAYMENT DATA].PmtRefNo "&_
                     "FROM [PAYMENT DATA] "&_
                     "INNER JOIN tblTypeGroup ON [PAYMENT DATA].TypeGroup = tblTypeGroup.TypeGroupID "&_
 		             "WHERE ([PAYMENT DATA].Expired = 0) AND ([PAYMENT DATA].ClientID = " & clientID & ") AND (tblTypeGroup.Active = 1) AND (tblTypeGroup.wsMedia = 1 OR "&_
					 "tblTypeGroup.TypeGroupID IN "&_
						"(SELECT TG2 FROM tblTGRelate INNER JOIN tblTypeGroup ON tblTGRelate.TG1 = tblTypeGroup.TypeGroupID WHERE tblTypeGroup.wsMedia = 1)) "
					if ss_CheckActivationDates then
                     strSQL = strSQL & "AND (" & DateSep & DateValue(DateAdd("n", Session("tzOffset"),Now)) & DateSep & " <= "&_
                     "[PAYMENT DATA].ExpDate) AND (Remaining > 0)"
					else
                     strSQL = strSQL & "AND (" & DateSep & DateValue(DateAdd("n", Session("tzOffset"),Now)) & DateSep & " BETWEEN [PAYMENT DATA].ActiveDate "&_
                     "AND [PAYMENT DATA].ExpDate) AND (Remaining > 0)"
					 end if
                 strSQL = strSQL & ") PD ON ISNULL(tblTGRelate.TG2, tblTypeGroup.TypeGroupID) = PD.TypeGroup"&_
             ") Payment "&_
 	         "INNER JOIN tblMedia ON Payment.TypeGroupID = tblMedia.TypegroupID "&_
             "INNER JOIN tblTypeGroupTab ON tblTypeGroupTab.TabID = " & session("TabID") & " AND tblMedia.TypegroupID = tblTypeGroupTab.TypeGroupID "&_
             "WHERE (" & DateSep & DateValue(DateAdd("n", Session("tzOffset"),Now)) & DateSep & " BETWEEN tblMedia.StartDate AND tblMedia.EndDate) "
	
	rsEntry.CursorLocation = 3
	rsEntry.open strSQL, cnWS
	Set rsEntry.ActiveConnection = Nothing
	
	if NOT rsEntry.EOF then
		if rsEntry.recordCount=1 then
			curMediaID = rsEntry("MediaID")
		end if
	end if
	rsEntry.close
	
	Clicked = false
	
	if request.form("frmAccessLink")<>"" then
		curMediaID = request.form("frmAccessLink")
	end if
	
	'if request.form("frmAccessLink")<>"" then
	if curMediaID<>"" then
	
		' BJD: Complicated Query
		' This query gets info for existing visits and valid services that can/have already paid for a link.
		strSQL = "SELECT Media.MediaID, Media.Name, Media.TypegroupID, Media.URL, Media.StartDate, Media.EndDate, Media.HasService, Media.Clicked, Media.TypeGroupName, Media.PmtRefNo, Media.TypePurch, "&_
                 "tblMedia.Description "&_
		         "FROM ("&_
                     "SELECT DISTINCT tblMedia.MediaID, tblMedia.Name, tblMedia.TypegroupID, tblMedia.URL, tblMedia.StartDate, tblMedia.EndDate, Payment.TypeGroupID AS HasService, Payment.Remaining, "&_
                     "Visits.MediaID AS Clicked, tblTypeGroup.TypeGroup AS TypeGroupName, Payment.PmtRefNo, Payment.TypePurch, Payment.Priority, Payment.ActiveDate, Payment.PaymentDate "&_
			         "FROM ("&_
                         "SELECT ISNULL(tblTypeGroup.TypeGroupID, PD.TypeGroup) AS TypeGroupID, PD.PmtRefNo, PD.TypePurch, PD.Remaining, PD.Priority, PD.ActiveDate, PD.PaymentDate "&_
                         "FROM tblTGRelate "&_
                         "RIGHT OUTER JOIN tblTypeGroup ON (tblTGRelate.TG1 = tblTypeGroup.TypeGroupID OR tblTGRelate.TG2 = tblTypeGroup.TypeGroupID) "&_
				         "RIGHT OUTER JOIN ("&_
                             "SELECT [PAYMENT DATA].TypeGroup, [PAYMENT DATA].PmtRefNo, [PAYMENT DATA].TypePurch, [PAYMENT DATA].Remaining, "&_
								"[Payment Data].Priority, [Payment Data].ActiveDate, [Payment Data].PaymentDate " &_
                             "FROM [PAYMENT DATA] "&_
                             "INNER JOIN tblTypeGroup ON [PAYMENT DATA].TypeGroup = tblTypeGroup.TypeGroupID "&_
				             "WHERE ([PAYMENT DATA].Expired = 0) AND ([PAYMENT DATA].ClientID = " & session("mvarUserID") & ") AND (tblTypeGroup.Active = 1) AND (tblTypeGroup.wsMedia = 1  OR "&_
							 "tblTypeGroup.TypeGroupID IN "&_
								"(SELECT TG2 FROM tblTGRelate INNER JOIN tblTypeGroup ON tblTGRelate.TG1 = tblTypeGroup.TypeGroupID WHERE tblTypeGroup.wsMedia = 1)) "
						     if ss_CheckActivationDates then			 
								strSQL = strSQL & "AND ([PAYMENT DATA].ActiveDate <= " & DateSep & DateValue(DateAdd("n", Session("tzOffset"),Now)) & DateSep & ") "
							 end if
                             strSQL = strSQL & "AND ([PAYMENT DATA].ExpDate >= " & DateSep & DateValue(DateAdd("n", Session("tzOffset"),Now)) & DateSep & ")"&_
                         ") PD ON ISNULL(tblTGRelate.TG2, tblTypeGroup.TypeGroupID) = PD.TypeGroup "&_
				     ") Payment "&_
                     "RIGHT OUTER JOIN tblMedia INNER JOIN tblTypeGroup ON tblMedia.TypegroupID = tblTypeGroup.TypeGroupID ON Payment.TypeGroupID = tblMedia.TypegroupID "&_
                     "LEFT OUTER JOIN ("&_
                         "SELECT [VISIT DATA].MediaID "&_
                         "FROM [PAYMENT DATA] "&_
                         "INNER JOIN [VISIT DATA] ON [PAYMENT DATA].PmtRefNo = [VISIT DATA].PmtRefNo "&_
                         "INNER JOIN tblTypeGroup ON [PAYMENT DATA].TypeGroup = tblTypeGroup.TypeGroupID "&_
				         "WHERE ([PAYMENT DATA].Expired = 0) AND ([PAYMENT DATA].ClientID = " & session("mvarUserID") & ") "
						 if ss_CheckActivationDates then
							strSQL = strSQL & "AND ([PAYMENT DATA].ActiveDate <= " & DateSep & DateValue(DateAdd("n", Session("tzOffset"),Now)) & DateSep & ") "
						 end if
                         strSQL = strSQL & "AND ([PAYMENT DATA].ExpDate >= " & DateSep & DateValue(DateAdd("n", Session("tzOffset"),Now)) & DateSep & ") AND (tblTypeGroup.wsMedia = 1 OR "&_
					     "tblTypeGroup.TypeGroupID IN "&_
						 "(SELECT TG2 FROM tblTGRelate INNER JOIN tblTypeGroup ON tblTGRelate.TG1 = tblTypeGroup.TypeGroupID WHERE tblTypeGroup.wsMedia = 1)) "&_
                     ") Visits ON tblMedia.MediaID = Visits.MediaID "&_
                 ") Media "&_
		         "INNER JOIN tblMedia ON Media.MediaID = tblMedia.MediaID "&_
		         "WHERE (Media.MediaID = " & curMediaID & ") AND (((Media.Remaining > 0) AND (Media.Clicked IS NULL)) OR (Media.Clicked IS NOT NULL)) "&_
				 "ORDER BY " &_
					"Media.Priority DESC, " 

		if curTG<>"" and curTG<>"0" then
			strSQL = strSQL & "CASE WHEN Media.TypeGroup=" & curTG & " THEN 1 ELSE 0 END DESC, " 
		end if

		strSQL = strSQL & "Media.ActiveDate, Media.PaymentDate, Media.PmtRefNo"
		         '"ORDER BY Media.TypeGroupName, Media.Name "

		'response.write debugSQL(strSQL, "SQL")
		rsEntry.CursorLocation = 3
		rsEntry.open strSQL, cnWS
		Set rsEntry.ActiveConnection = Nothing
		
		if NOT rsEntry.EOF then
			
			if NOT isNull(rsEntry("Clicked")) then
				Clicked = true
			end if
			PmtRefNo = rsEntry("PmtRefNo")
			PaymentName = rsEntry("TypePurch")
			PaymentTG = rsEntry("TypeGroupID")
			
			MediaName = rsEntry("Name")
			MediaURL = rsEntry("URL")
			
		end if
		rsEntry.close
		
		if NOT Clicked then ' has not clicked this link yet
			''' Insert new VD record
			strSQL = "INSERT INTO [VISIT DATA] (ClientID, MediaID, TypeTaken, PmtRefNo, TypeGroup, RequestDate, Location, OldClientID, Value, ClassType, ClassDate, ClassTime, CreationDateTime) VALUES ("
			strSQL = strSQL & session("mvarUserID")
			strSQL = strSQL & ", " & curMediaID
			strSQL = strSQL & ", N'" & Replace(PaymentName, "'", "''") & "'"
			strSQL = strSQL & ", " & PmtRefNo
			strSQL = strSQL & ", " & PaymentTG
			strSQL = strSQL & ", " & DateSep & DateAdd("n", Session("tzOffset"),Now) & DateSep
			strSQL = strSQL & ", 98"
			strSQL = strSQL & ", 0"
			strSQL = strSQL & ", 1"
			strSQL = strSQL & ", N'media'"
			strSQL = strSQL & ", " & DateSep & DateValue(DateAdd("n", Session("tzOffset"),Now)) & DateSep
			strSQL = strSQL & ", " & TimeSepB & TimeValue(DateAdd("n", Session("tzOffset"),Now)) & TimeSepA
			strSQL = strSQL & ", " & DateSep & DateAdd("n", Session("tzOffset"),Now) & DateSep
			strSQL = strSQL & ")"
			'response.write debugSQL(strSQL, "SQL")
			cnWS.execute strSQL
			
			'' Update PD
			call ReCalcPaymentData(PmtRefNo, false)
			'cnWS.execute strSQL
			
		else ' has been clicked
		
			'' update request date
		
		end if ' link clicked vs not clicked
		
	end if
	
	' popup link
	if curMediaID<>"" then
%>
	<script type="text/javascript">
		showMedia('<%=curMediaID%>');
	</script>
<%
	end if
%>
    <style type="text/css">
    .topFilters {
        margin: 0 auto;
        position: relative;
        width: 960px;
    }
    .topFiltersInner {
        float: right;
        padding: 10px 0 32px 10px;
    }

    .dateControls { 
        background-clip: padding-box;
        border-radius: 10px 10px 10px 10px;
        margin: 0 auto;
        padding: 5px 0;
        position: absolute;
        right: 10px;
        top: 25px;
        width: 335px;
        z-index: 700;
    }

    .dateControls .leftSide {
        float: left;
        padding-top: 1px;
    }
    .dateControls .rightSide {
        float: right;
        padding-top: 1px;
    }

    .cur-date {
        display: none;
    }
    #day-tog-c, #week-tog-c {
        color: #555;
    }
    h1 {
        background: #FFFFFF;
        font-size: 24px;
        line-height: 32px;
        margin-top: 10px;
        padding: 21px 0 15px 5px;
    }
    .wrapper {
        padding-top: 127px;
    }
    #main-content 
    {
    	margin-top: 0px;
    }
    .section {
        margin-bottom: 10px;
    }
    table#mediaInnerTableVideos td { color: #555;}
    </style>
    </head>
    <body>
    <!-- #include file="adm/inc_alert_content.asp" -->

    <form name="search2" method="post" action="main_media.asp" style="width:100%" />
      <input type="hidden" name="pageNum" value="1" />
      <input type="hidden" name="requiredtxtUserName" value="" />
      <input type="hidden" name="requiredtxtPassword" value="" />
      <input type="hidden" name="optForwardingLink" value="" />
	  <input type="hidden" name="optRememberMe" value="" />
	  <input type="hidden" name="tabID" value="<%=session("tabID")%>"/>

	  <div class="wrapperTop" >
        <div class="wrapperTopInner">
            <div class="pageTop">
                <div class="pageTopLeft">
                    &nbsp; 
                </div>
                <div class="pageTopRight">
                    &nbsp; 
                </div>
                    <h1><%=phraseDictionary("Media") %></h1> 
                    <div id="dateControls" class="dateControls">
                    <div  class="leftSide" >
                        <% dayAndWeekControls %>
                    </div>
                    <div class="rightSide" >
                        <% dayInfo %>
                    </div>
                </div>
            </div><!-- pageTop -->
        </div><!-- wrapperTopInner -->
    </div> <!-- wrapperTop -->

    <div class="fixedHeader">
        <div class="topFilters">
            <div class="topFiltersInner filterList"> 
                <%
				dim boolShowTG
				boolShowTG = false
				
				''Check for > 1 TG otherwise don't display TG selector
				strSQL = "SELECT tblTypeGroup.TypeGroupID, tblTypeGroup.TypeGroup FROM tblTypeGroup "
				if session("tabID")<>"" AND isNumeric(session("tabID")) then
					strSQL = strSQL & " INNER JOIN tblTypeGroupTab ON tblTypeGroupTab.TypeGroupID = tblTypeGroup.TypeGroupID AND tblTypeGroupTab.TabID = " & session("tabID")
				end if 
				strSQL = strSQL & "  WHERE active=1 AND wsMedia=1 AND wsDisable<>1 ORDER BY TypeGroup"
				'response.write debugSQL(strSQL, "SQL")
				rsEntry.CursorLocation = 3
				rsEntry.open strSQL, cnWS
				Set rsEntry.ActiveConnection = Nothing

				if NOT rsEntry.EOF then
					rsEntry.MoveNext
					if NOT rsEntry.EOF then
						boolShowTG = true
						rsEntry.MoveFirst
					end if
				end if

				if boolShowTG then
%>
                <select name="optTG" onChange="subFormClrPNum();">
                  <option value="0">All</option>
                  <%
						Do While NOT rsEntry.EOF
%>
                  <option value="<%=rsEntry("TypeGroupID")%>" <%if curTG=rsEntry("TypeGroupID") then response.write "selected" end if%>><%=Left(rsEntry("TypeGroup"),18)%></option>
                  <%
							rsEntry.MoveNext
						Loop
%>
                </select>
				<script type="text/javascript">

					$(document).ready(function () {
						document.search2.optTG.options[0].text = '<%=allHotWords(149) %>' + " " + '<%=allHotWords(503)%>';
					});
				</script>
<%				
				end if '''1 or less hide
			    rsEntry.close
%>
                
                <%	if ss_SchedShowDay then %>
				<select name="optView" onChange="subForm();">
					<option value="day" <%if request.querystring("view")="day" OR (request.querystring("view")="" AND NOT ss_apptDefaultToWeekView) then response.write "selected" end if %>>Daily</option>
					<option value="week" <%if request.querystring("view")="week" OR (request.querystring("view")="" AND ss_apptDefaultToWeekView) then response.write "selected" end if %>>Weekly</option>
				</select>
				<% else %>
						<input type="hidden" name="optView" value="week">
				<% end if '''day/week view %>
            </div>
        </div>
	</div> <%' fixedHeaderdiv %>
  </form>
  <div class="clear:both"></div>
<% pageStart %>

  <table id="mediaOuterTable" width="<%=strPageWidth%>" cellspacing="0" height="100%">
    <tr>
	    <td colspan="2" class="mainTextBig"><b>Click on a Link to View!</b></td>
    </tr>
	<tr> 
		<td valign="top" height="100%" width="100%" style="background-color:#FFFFFF;"> 
			<form name="frmMedia" action="main_media.asp" method="post">
			<input type="hidden" name="frmAccessLink" value="">
<%
			dim curProgram : curProgram = ""
			
			strSQL = "SELECT Media.MediaID, Media.Name, Media.TypegroupID, Media.URL, Media.StartDate, Media.EndDate, Media.HasService, Media.Clicked, Media.TypeGroupName, tblMedia.Description "&_
				     "FROM ("&_
                            "SELECT DISTINCT tblMedia.MediaID, tblMedia.Name, tblMedia.TypegroupID, tblMedia.URL, tblMedia.StartDate, tblMedia.EndDate, Payment.TypeGroupID AS HasService, "&_
                            "Visits.MediaID AS Clicked, tblTypeGroup.TypeGroup AS TypeGroupName "&_
					        "FROM ("&_
                                "SELECT  ISNULL(tblTypeGroup.TypeGroupID, PD.TypeGroup) AS TypeGroupID, PD.PmtRefNo "&_
                                "FROM tblTGRelate "&_
                                "RIGHT OUTER JOIN tblTypeGroup ON (tblTGRelate.TG1 = tblTypeGroup.TypeGroupID OR tblTGRelate.TG2 = tblTypeGroup.TypeGroupID) "&_
					            "RIGHT OUTER JOIN ("&_
                                                    "SELECT [PAYMENT DATA].TypeGroup, [PAYMENT DATA].PmtRefNo "&_
                                                    "FROM [PAYMENT DATA] "&_
                                                    "INNER JOIN tblTypeGroup ON [PAYMENT DATA].TypeGroup = tblTypeGroup.TypeGroupID "&_
					                                "WHERE ([PAYMENT DATA].Expired = 0) AND ([PAYMENT DATA].ClientID = " & clientID & ") "&_
                                                    "AND (tblTypeGroup.Active = 1) AND (tblTypeGroup.wsMedia = 1  OR "&_
													"tblTypeGroup.TypeGroupID IN "&_
													"(SELECT TG2 FROM tblTGRelate INNER JOIN tblTypeGroup ON tblTGRelate.TG1 = tblTypeGroup.TypeGroupID WHERE tblTypeGroup.wsMedia = 1)) "
													if ss_CheckActivationDates then		
														strSQL = strSQL & "AND ([PAYMENT DATA].ActiveDate <= " & DateSep & DateValue(DateAdd("n", Session("tzOffset"),Now)) & DateSep & ") "
													end if
                                                    strSQL = strSQL & "AND ([PAYMENT DATA].ExpDate >= " & DateSep & DateValue(DateAdd("n", Session("tzOffset"),Now)) & DateSep & ") "&_
                                                    "AND (Remaining > 0)"&_
                                                    ") PD ON ISNULL(tblTGRelate.TG2, tblTypeGroup.TypeGroupID) = PD.TypeGroup "&_
					            ") AS Payment "&_
                            "RIGHT OUTER JOIN tblMedia INNER JOIN tblTypeGroup ON tblMedia.TypegroupID = tblTypeGroup.TypeGroupID ON Payment.TypeGroupID = tblMedia.TypegroupID "&_
                            "LEFT OUTER JOIN ("&_
                                            "SELECT DISTINCT [VISIT DATA].MediaID "&_
                                            "FROM [PAYMENT DATA] "&_
                                            "INNER JOIN [VISIT DATA] ON [PAYMENT DATA].PmtRefNo = [VISIT DATA].PmtRefNo "&_
                                            "INNER JOIN tblTypeGroup ON [PAYMENT DATA].TypeGroup = tblTypeGroup.TypeGroupID "&_
					                        "WHERE ([PAYMENT DATA].Expired = 0) AND ([PAYMENT DATA].ClientID = " & clientID & ") "
											if ss_CheckActivationDates then
												strSQL = strSQL & "AND ([PAYMENT DATA].ActiveDate <= " & DateSep & DateValue(DateAdd("n", Session("tzOffset"),Now)) & DateSep & ") "
											end if
                                            strSQL = strSQL & "AND ([PAYMENT DATA].ExpDate >= " & DateSep & DateValue(DateAdd("n", Session("tzOffset"),Now)) & DateSep & ") "&_
                                            "AND (tblTypeGroup.wsMedia = 1 OR "&_
											"tblTypeGroup.TypeGroupID IN "&_
											"(SELECT TG2 FROM tblTGRelate INNER JOIN tblTypeGroup ON tblTGRelate.TG1 = tblTypeGroup.TypeGroupID WHERE tblTypeGroup.wsMedia = 1)) "&_
                                            ") Visits ON tblMedia.MediaID = Visits.MediaID "
			if session("tabID")<>"" AND isNumeric(session("tabID")) then
				strSQL = strSQL & "INNER JOIN tblTypeGroupTab ON tblTypeGroupTab.TypeGroupID = tblTypeGroup.TypeGroupID AND tblTypeGroupTab.TabID = " & session("tabID")
			end if 
			strSQL = strSQL & " WHERE (" & DateSep & DateAdd("n", Session("tzOffset"),Now) & DateSep & " BETWEEN tblMedia.StartDate AND tblMedia.EndDate) "
			if curTG<>"" and curTG<>"0" then
				strSQL = strSQL & "AND tblMedia.TypeGroupID = " & curTG & " "
			end if
					
			strSQL = strSQL & ") AS Media "&_
			         "INNER JOIN tblMedia ON Media.MediaID = tblMedia.MediaID "&_
			         "ORDER BY Media.TypeGroupName, Media.Name "
			
			'response.write debugSQL(strSQL, "SQL")
            'response.Write("sql " & strSQL)
			rsEntry.CursorLocation = 3
			rsEntry.open strSQL, cnWS
			Set rsEntry.ActiveConnection = Nothing
			
			if NOT rsEntry.EOF then%>
            <div id="innerContentDiv" class="section" style="width:<% if curMediaID="" then response.write "60%" else response.write "90%" end if %>"><%
				do while NOT rsEntry.EOF
					if curProgram<>rsEntry("TypeGroupName") then
						curProgram = rsEntry("TypeGroupName")
%>
			            <table id="mediaInnerTable<%=trim(replace(curProgram, " ", ""))%>" class="mainText center" cellspacing="0" width="100%">
					        <tr>
						        <td colspan="2" class="" style=""><h2><%=curProgram%></h2></td>
					        </tr>
<%
					end if
%>
					<tr height="20px">
						<td width="40%" valign="top" class="mainTextBig">&nbsp;&nbsp;&nbsp;&nbsp;
							<strong>
							<%
                            if NOT isNull(rsEntry("HasService")) OR NOT isNull(rsEntry("Clicked")) then
                                response.Write("<a ")
								if NOT isNull(rsEntry("Clicked")) then 
                                    response.write("style='color:#CC0000' href='javascript:showMedia(" & rsEntry("MediaID") & ");'")
                                else 
                                   response.write("href='javascript:accessLink(" & rsEntry("MediaID") & ");'")
                                end if
                                response.Write(">")
							end if
							response.write(rsEntry("Name"))

							if NOT isNull(rsEntry("HasService")) OR NOT isNull(rsEntry("Clicked")) then
                                response.write("</a>")
                            end if %>
							</strong>
						</td>
						<td style="padding: 5px 5px 5px 5px; "><div style="border:1px solid <%=session("PageColor2")%>; background-color:#EFEFEF; width: 80%; padding: 5px 5px 5px 5px; "><%=rsEntry("Description")%></div></td>
					</tr>
<%
					rsEntry.moveNext
				loop
%>
			</table>
            </div>
<%
			end if ' eof
%>
		</form>
		</td>
    </tr>
</table>
<% pageEnd %>
<!-- #include file="post.asp"-->
