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
if not Session("Pass") OR Session("Admin")="false" OR NOT validAccessPriv("RPT_LOGS") then 
%>
<script type="text/javascript">
	alert("You are not authorized to view this page.\nPlease login to an account with Access Privileges granted for this feature.");
	javascript:history.go(-1);
</script>
<%
else
%>
		<!-- #include file="../inc_i18n.asp" -->
		<!-- #include file="inc_row_colors.asp" -->
		<!-- #include file="inc_rpt_tagging.asp" -->
		<!-- #include file="inc_date_arrows.asp" -->
		<!-- #include file="inc_utilities.asp" -->
        <!-- #include file="inc_rpt_save.asp" -->
<%
	if useVersionB(TEST_NAVIGATION_REDESIGN) then
		Session("tabID") = "97"
	end if

	dim category : category = ""
	if RQ("category") <> "" then
		category = RQ("category")
	elseif RF("category") <> "" then
		category = RF("category")
	end if

    ' Setup the general-purpose recordset (used all over the place)
    dim rsEntry
    set rsEntry = Server.CreateObject("ADODB.Recordset")
    
    ' Variables taken from the submitted form ------------------------------------------------------------
	Dim cSDate, cEDate, ap_RPT_DATAACCESS
	Dim cSalesRecordedStartDate, cSalesRecordedEndDate
	Dim typeGroupID, visitTypeID
	Dim trainerID, showAllTrainers, showNoResults
	Dim filterByRevenueCategory
	Dim trainerIDCache, trainerIDCacheSize, trainerNameCache

	ap_RPT_DATAACCESS = validAccessPriv("RPT_DATAACCESS")

    ' Retrieve the form variables ------------------------------------------------------------------------
	if request.form("requiredtxtDateStart")<>"" then
		Call SetLocale(session("mvarLocaleStr"))
			cSDate = CDATE(request.form("requiredtxtDateStart"))
		Call SetLocale("en-us")
	else
		'cSDate = DateAdd("ww",-1,DateValue(CDATE(DateAdd("n", Session("tzOffset"),Now))))
		cSDate = DateAdd("y",-30,DateValue(CDATE(DateAdd("n", Session("tzOffset"),Now)))) 'Minus 14 days
	end if

	if request.form("requiredtxtDateEnd")<>"" then
		Call SetLocale(session("mvarLocaleStr"))
			cEDate = CDATE(request.form("requiredtxtDateEnd"))
		Call SetLocale("en-us")
	else
		cEDate = DateValue(CDATE(DateAdd("n", Session("tzOffset"),Now)))
	end if
	
	if request.form("txtSalesRecordedStart")<>"" then
		Call SetLocale(session("mvarLocaleStr"))
			cSalesRecordedStartDate = CDATE(request.form("txtSalesRecordedStart"))
		Call SetLocale("en-us")
	else
		cSalesRecordedStartDate = DateAdd("y",-14,DateValue(CDATE(DateAdd("n", Session("tzOffset"),Now)))) 'Minus 14 days
	end if

	if request.form("txtSalesRecordedEnd")<>"" then
		Call SetLocale(session("mvarLocaleStr"))
			cSalesRecordedEndDate = CDATE(request.form("txtSalesRecordedEnd"))
		Call SetLocale("en-us")
	else
		cSalesRecordedEndDate = DateValue(CDATE(DateAdd("n", Session("tzOffset"),Now)))
	end if
	
	' Typegroup and Visit Type
	if request.Form("requiredtxtVisitType") <> "" then
	    dim val : val = split(request.Form("requiredtxtVisitType"), "#")
	    typeGroupID = CLNG(val(0))
	    visitTypeID = CLNG(val(1))
	end if
	
	' Trainer
	if request.Form("optTrainer") <> "" then
	    trainerID = CLNG(request.Form("optTrainer"))
	    showAllTrainers = false
	else
	    trainerID = 0
	    showAllTrainers = true
	end if
	
	if request.Form("optShowNoResults") <> "" then
	    showNoResults = true
	else
	    showNoResults = false
	end if
	
	' Filter by Revenue Category
	filterByRevenueCategory = (request.Form("optFilterByRevenueCategory") <> "")
	
	' Filter by Session Date
	filterBySessionDate = not (request.Form("optShowAll") <> "")

	' set the row colors- access via getRowColor
	setRowColors "#FAFAFA", "#F2F2F2"
	
	
	' Load the trainers into a cache (avoids multiple queries)
	'Output the trainer list
    strSQL = "SELECT TRAINERS.TrFirstName + ' ' + TRAINERS.TrLastName AS name, TRAINERS.TrainerID "
    strSQL = strSQL & "FROM TRAINERS WHERE TRAINERS.Active = 1 AND TRAINERS.TrainerID > 0 AND TRAINERS.TrainerID <> 0 AND TRAINERS.isSystem=0 "
   response.write debugSQL(strSQL, "SQL")
    rsEntry.CursorLocation = 3
    rsEntry.open strSQL, cnWS
    Set rsEntry.ActiveConnection = Nothing
    
    ' This is an optimization for later so we don't need to do this query again.
    ' Cache the trainer IDs
    trainerIDCacheSize = rsEntry.RecordCount
    ReDim trainerIDCache(trainerIDCacheSize)
    ReDim trainerNameCache(trainerIDCacheSize)
    
    dim cacheCount : cacheCount = 0

    Do While NOT rsEntry.EOF
        ' Cache the trainer info
        trainerIDCache(cacheCount) = CLNG(rsEntry("TrainerID"))
        trainerNameCache(cacheCount) = rsEntry("Name")
        cacheCount = cacheCount + 1
        
        ' Nothing left to see here...
        rsEntry.moveNext
    loop
    
    ' Close the recordset
    rsEntry.close
%>
<% if request.form("frmExpReport")<>"true" then %>

<% 'Start of HTML Output ---------------------------------------------------------------------------------------------- %>

<% 'Javascript Goes Here ---------------------------------------------------------------------------------------------- %>

<!-- #include file="pre.asp" -->
<!-- #include file="frame_bottom.asp" -->
<%= js(array("mb", "adm/adm_rpt_conversions")) %>
<script type="text/javascript">
    function exportReport() {
	    document.frmConversions.frmExpReport.value = "true";
	    document.frmConversions.frmGenReport.value = "true";
	    <% iframeSubmit "frmConversions", "adm_rpt_conversions.asp" %>
    }
</script>

<%= js(array("calendar" & dateFormatCode, "reportFavorites", "plugins/jquery.SimpleLightBox")) %>


<!-- #include file="../inc_date_ctrl.asp" -->
<!-- #include file="../inc_ajax.asp" -->
<!-- #include file="../inc_val_date.asp" -->

<%
    'CSS Files ----------------------------------------------------------------------------------------------
    css(array("SimpleLightBox"))
%>


<% 'Start of Body Section ---------------------------------------------------------------------------------------------- %>
<% pageStart %>
<style type="text/css">
#tableFrmConversions TD {
    padding: 5px;
}
#tableFrmConversions TD TD{
    padding: 0;
}
</style>
<% if UseVersionB(TEST_REPORT_LANDING_PAGE) then %>
	<div class="headText breadcrumbs-old" align="left">
	<span class="breadcrumb-item"><a href="/reportslandingpage/index"><%=DisplayPhrase(reportPageTitlesDictionary, "Reports")%></a></span>
	<span class="breadcrumb-item">&raquo;</span>
	<%if category <> "" then%>
	<span class="breadcrumb-item active"><a href="/reportslandingpage/<%=category%>category"><%=GetReportsCategoryPagePhrase(category)%></a></span>
	<span class="breadcrumb-item">&raquo;</span>
	<%end if %>
	<%=DisplayPhrase(reportPageTitlesDictionary, "Trainerconversions")%>
	<div id="add-to-favorites">
		<span id="" class="favorite"></span><span class="text"><%=DisplayPhrase(reportPageTitlesDictionary, "Addtofavorites")%></span>
	</div>
	</div>
<%end if %>
<table height="100%" width="<%=strPageWidth%>" cellspacing="0">    
<tr> 
     <td valign="top" height="100%" width="100%"> <br />
        <table cellspacing="0" width="90%" height="100%" style="margin: 0 auto;">
	<% if NOT UseVersionB(TEST_REPORT_LANDING_PAGE) then %>
		  <tr> 
            <td class="headText" align="left"><b><%= pp_PageTitle("Trainer Conversions") %></b></td>
          </tr>
	<%end if %>
          <tr> 
            <td valign="top" class="mainText"> 
              <table id="tableFrmConversions" class="mainText border4" cellspacing="0" width=100% style="margin: 0 auto;">
                <form name="frmConversions" action="adm_rpt_conversions.asp" method="POST">
				<% if UseVersionB(TEST_REPORT_LANDING_PAGE) then %>
					<% if category <> "" then %>
						<input type="hidden" name="category" id="category" value="<%=category %>" />
					<%end if %>
					<input type="hidden" name="reportUrl" id="reportUrl" value="<%=Request.ServerVariables("SCRIPT_NAME") %>" />
				<%end if %>
					<input type="hidden" name="frmGenReport" value="">
					<input type="hidden" name="frmExpReport" value="">
                  <tr>
                  <% ' Start of filter section ------------------------------------------------------------------------------- %>
                    <td align="left" valign="bottom" style="background-color:#FAFAFA;" nowrap id="Step1"><b>Step 1:</b></td>
                    <td style="background-color:#FAFAFA;">
                    <%  strSQL = "SELECT tblTypeGroup.TypeGroup, tblTypeGroup.TypeGroupID, tblVisitTypes.TypeID, tblVisitTypes.TypeName FROM tblVisitTypes INNER JOIN tblTypeGroup ON tblVisitTypes.Typegroup = tblTypeGroup.TypeGroupID "
		                strSQL = strSQL & "WHERE (tblVisitTypes.Active = 1) AND (tblVisitTypes.[Delete] = 0) AND (tblTypeGroup.Active = 1) "
		                strSQL = strSQL & "ORDER BY tblTypeGroup.TypeGroup, tblVisitTypes.SortOrder, tblVisitTypes.TypeName"
		               response.write debugSQL(strSQL, "SQL")
		                'response.End
		                rsEntry.CursorLocation = 3
		                rsEntry.open strSQL, cnWS
		                Set rsEntry.ActiveConnection = Nothing
		                
		                ' Need to output the TypeGroup and then all the sub-sessions indented. The SQL is not formated nicely for this
		                ' So I use an old trick. As we iterate through the result set, check to see if TypeGroup is different from the
		                ' last entry. If it is, output an option of just the TypeGroup, otherwise, indent and output only the VisitType.
		                dim lastTypeGroup : lastTypeGroup = ""
                    %>
			            Select Promotional Session Type: <select name="requiredtxtVisitType" id="requiredtxtVisitType">
						
                    <%
		                Do While NOT rsEntry.EOF
		                    ' Check if the TypeGroup is different from the last
		                    if rsEntry("TypeGroupID") <> lastTypeGroup then
		                        response.Write("<option value=""" & rsEntry("TypeGroupID") & "#0"" ")
		                        lastTypeGroup = rsEntry("TypeGroupID")
		                        
		                        ' If this is the selected typegroup (and there is no selected VisitType),
		                        ' mark it as selected
		                        if typeGroupID = lastTypeGroup AND visitTypeID = 0 then
		                            response.Write("selected")
		                        end if
		                        
		                        response.Write(">" & rsEntry("TypeGroup") & "</option>")
		                        
		                        ' Do not advance the SQL recordset; the next time it loops the typegroup will be the
		                        ' same and it will go to the else statement, and print the first option.
		                    else 
		                        ' Same type group, writing out visit types
			                    response.Write("<option value=""" & rsEntry("TypeGroupID") & "#" & rsEntry("TypeID") & """ ")
			                    
			                    ' If this is the selected VisitType, mark it as selected
		                        if visitTypeID = rsEntry("TypeID") then
		                            response.Write("selected")
		                        end if
			                    
			                    response.Write(">&nbsp;&nbsp;&nbsp;&nbsp;" & rsEntry("TypeName") & "</option>")
			                    
			                    ' Since the SQL only includes these sub-options, advance the record
			                    rsEntry.MoveNext
			                end if
			            Loop
		                
		                rsEntry.close
		            %>
		                </select></b>
		            </td></tr>
		               
			        <% 'Step 2: Session Date Filter --------------------------------------------------------------------- %>
			        <tr><td align="left" valign="bottom" style="background-color:#F2F2F2;" nowrap id="Step2"><b>Step 2:</b></td>
                    <td style="background-color:#F2F2F2;">
			            Show All? <img src="<%= contentUrl("/asp/adm/images/tech-tip-16px.png") %>" align="middle" title="Uncheck this box to filter the date of the initial session."> <input type="checkbox" name="optShowAll" <% if not filterBySessionDate then response.write "checked" end if %> onclick="$('#SessionDateFilters').toggle()" />
                        <div id="SessionDateFilters" style="display: <% if filterBySessionDate then response.write "block" else response.write "none" end if %>">
                            &nbsp;For All Sessions Delivered Between: 
                            <input onBlur="validateDate(this, '<%=FmtDateShort(cSDate)%>', true);" type="text"  name="requiredtxtDateStart" value="<%=FmtDateShort(cSDate)%>" class="date">
                <script type="text/javascript">
		            var cal1 = new tcal({'formname':'frmConversions', 'controlname':'requiredtxtDateStart'});
		            cal1.a_tpl.yearscroll = true;
	            </script>
					        &nbsp;And: 
					        <input onBlur="validateDate(this, '<%=FmtDateShort(cEDate)%>', true);" type="text"  name="requiredtxtDateEnd" value="<%=FmtDateShort(cEDate)%>" class="date">
                <script type="text/javascript">
		            var cal2 = new tcal({'formname':'frmConversions', 'controlname':'requiredtxtDateEnd'});
		            cal2.a_tpl.yearscroll = true;
	            </script>
					        &nbsp;
					        <table cellspacing="0">
					            <tr>
					                <td>
            					        <br /><% showDateArrows("frmConversions") %>
					                </td>
					            </tr>
					        </table>
					    </div>
				    </td></tr>
		          
		            <% 'Step 3: Sales Date Filter --------------------------------------------------------------------- %>
		            <tr><td align="left" valign="bottom" style="background-color:#FAFAFA;" nowrap id="Step3"><b>Step 3:</b></td>
                    <td style="background-color:#FAFAFA;">
		                Show subsequent purchases made between: <input onBlur="validateDate(this, '<%=FmtDateShort(cSalesRecordedStartDate)%>', true);" type="text"  name="txtSalesRecordedStart" value="<%=FmtDateShort(cSalesRecordedStartDate)%>" class="date">
                <script type="text/javascript">
		            var cal3 = new tcal({'formname':'frmConversions', 'controlname':'txtSalesRecordedStart'});
		            cal3.a_tpl.yearscroll = true;
	            </script>
				        &nbsp;And: 
				        <input onBlur="validateDate(this, '<%=FmtDateShort(cSalesRecordedEndDate)%>', true);" type="text"  name="txtSalesRecordedEnd" value="<%=FmtDateShort(cSalesRecordedEndDate)%>" class="date">
                <script type="text/javascript">
		            var cal4 = new tcal({'formname':'frmConversions', 'controlname':'txtSalesRecordedEnd'});
		            cal4.a_tpl.yearscroll = true;
	            </script>
				        <br />
				        <%= showDateArrowsCustomInputNames("frmConversions", "txtSalesRecordedStart", "txtSalesRecordedEnd") %>
			        </td></tr>
				    
				    <%  'Trainer Filter --------------------------------------------------------------------------- %>
				    <tr><td align="left" valign="bottom" style="background-color:#F2F2F2;" nowrap id="Step4"><b>Step 4:</b></td>
                    <td style="background-color:#F2F2F2;">
				        Select <%=xssStr(allHotWords(6))%> <img src="<%= contentUrl("/asp/adm/images/tech-tip-16px.png") %>" align="middle" title="Select the trainer to view conversions for,">: 
                        <select name="optTrainer">
		                <option value="" style="background-color:#FFFF99">All <%=xssStr(allHotWords(113))%></option>%>
        				
    				<%  'Output the trainer list
		         
		                'trainerIDCacheSize = rsEntry.RecordCount
		                'ReDim trainerIDCache(trainerIDCacheSize)
		                'ReDim trainerNameCache(trainerIDCacheSize)
		                
		                cacheCount = 0
                    
    				    Do While cacheCount < trainerIDCacheSize
    				        response.Write("<option value=""" & trainerIDCache(cacheCount) & """ ")
    				        
    				        ' If a trainer has been selected, mark it
    				        if trainerID = trainerIDCache(cacheCount) then
    				            response.Write("selected")
    				        end if
    				        
    				        response.Write(">" & trainerNameCache(cacheCount) & "</option>")
    				        cacheCount = cacheCount + 1
    				    loop
    				%>
    				    </select> 
    				    &nbsp;&nbsp;
    				    Show <%=xssStr(allHotWords(113))%> with no results: 
    				    <input type="checkbox" name="optShowNoResults" <% if showNoResults then response.Write("checked") end if %>/>
			        </td></tr>
				    
				    <%  'Step 5: Sales from Original Category --------------------------------------------------------------------------- %>
				    <tr><td align="left" valign="bottom" style="background-color:#FAFAFA;" nowrap id="Step5"><b>Step 5:</b></td>
                    <td style="background-color:#FAFAFA;">
				        Show only subsequent sales from the original promotional session's revenue category: <input name="optFilterByRevenueCategory" type=checkbox <% if filterByRevenueCategory then response.Write("checked") end if %>/>
				    </td></tr>
				    
				    <%  'Generate and controls --------------------------------------------------------------------------- %>
				    <tr><td align="left" valign="bottom" style="background-color:#F2F2F2;" nowrap id="Step6"><b>Step 6:</b></td>
                    <td style="background-color:#F2F2F2;"> 
					    <input type="button" name="Button" value="Generate" onClick="genReport();">
					<%  if not Session("Pass") OR Session("Admin")="false" OR NOT validAccessPriv("RPT_EXPORT") then
					    else
					        exportToExcelButton
						    savingButtons "frmConversions", "Trainer Conversions"
					    end if %>
					</td>
                </tr></b></form>
              </table><br />

			</td>
          </tr>
          <tr> 
            <td valign="top" class="mainTextBig"> 
<%  end if ' hide in export

    if request.form("frmGenReport")="true" then

	    if request.form("frmExpReport")="true" then
		    Dim stFilename
		    stFilename="attachment; filename=Trainer_Conversions_" & Replace(cSDate,"/","-") & " to " & Replace(cEDate,"/","-") & ".xls" 
		    Response.ContentType = "application/vnd.ms-excel" 
		    Response.AddHeader "Content-Disposition", stFilename 
	    end if
	    
	    ' This table should be printed as long as we are generating a report, even for export
%>
        <table class="mainText" width="100%" cellspacing="0" style="margin: 0 auto;">
        <% 'Output of report starts here! --------------------------------------------------------
        ' If we are outputting all the trainers, loop through them
        if showAllTrainers then
            ' Damn VB array semantics not being 0 based
            for trainerCount = 0 to (trainerIDCacheSize - 1)
                outputTrainerConversionRecord trainerIDCache(trainerCount), trainerNameCache(trainerCount)
            next
        else
            ' Need to get the trainer name that matches the selected trainer ID
            for trainerCount = 0 to (trainerIDCacheSize - 1)
                if trainerIDCache(trainerCount) = trainerID then
                    outputTrainerConversionRecord trainerID, trainerNameCache(trainerCount)
                end if
            next				    
        end if
        %>
        <tr> 
            <td colspan="5">&nbsp;</td>
        </tr>
        <tr > 
          <td  colspan="2" valign="top" class="mainTextBig center-ch"> 
            </td>
        </tr>
        </table>
	<% end if ' genReport %>
	<% if request.form("frmExpReport")<>"true" then %>
            </td>
          </tr>
          
        </table>
    </td>
    </tr>

		</table>
<% pageEnd %>
<!-- #include file="post.asp" -->


<%
	end if
	
end if

' Output one element of a trainer conversion record. Uses the global parameters of the page,
' varies based on the trainer. For "All Trainers" report, call this for each trainer.
' tid = Trainer ID to output
function outputTrainerConversionRecord(tid, trainerName)
    ' Setup the record set
    dim rsTrainer
    set rsTrainer = Server.CreateObject("ADODB.Recordset")
    
    ' First, prep the inner query, which will find the client IDs of all the clients who had sessions
    ' with trainer tid within the specified class dates
    clientSQL = "SELECT min(vd.ClientID) as ClientID, min(vd.ClassDate) as ClassDate, MIN(PRODUCTS.CategoryID) as RevenueCategory "
    clientSQL = clientSQL & "FROM [VISIT DATA] vd "
    clientSQL = clientSQL & "INNER JOIN [PAYMENT DATA] ON vd.PmtRefNo = [PAYMENT DATA].PmtRefNo INNER JOIN PRODUCTS ON [PAYMENT DATA].ProductID = PRODUCTS.ProductID "
    clientSQL = clientSQL & "WHERE vd.TrainerID = " & tid
    
    ' If the visitTypeID is nothing, then we are looking for entire programs
    if visitTypeID = 0 then
        clientSQL = clientSQL & " AND vd.TypeGroup = " & typeGroupID
    else
        clientSQL = clientSQL & " AND vd.VisitType = " & visitTypeID
    end if
    
    ' If we are filtering by session date, do that
    if filterBySessionDate then
        clientSQL = clientSQL & " AND vd.ClassDate <= '" & cEDate & "' AND vd.ClassDate >= '" & cSDate & "' "
    end if
    
    ' Throw on the group by
    clientSQL = clientSQL & " GROUP BY vd.ClientID"
    'response.write clientSQL
    
    ' Now we have a working list of clients
    ' Make a much bigger select from the sales tables. 
    totalSQL = "SELECT s.ClientID, c.FirstName + ' ' + c.LastName as Name, eligibleClients.ClassDate as SessionDelivered, s.SaleDate, sd.Description, ((sd.UnitPrice * sd.Quantity) + sd.ItemTax1 + sd.ItemTax2 + sd.ItemTax3 + sd.ItemTax4 + sd.ItemTax5 - sd.DiscAmt) as LinePrice "
    totalSQL = totalSQL & "FROM Sales s "
	totalSQL = totalSQL & "INNER JOIN (" & clientSQL & ") eligibleClients ON eligibleClients.ClientID = s.ClientID "
	totalSQL = totalSQL & "INNER JOIN CLIENTS c on c.ClientID = s.ClientID "
	totalSQL = totalSQL & "INNER JOIN [Sales Details] sd ON sd.SaleID = s.SaleID "
	totalSQL = totalSQL & "INNER JOIN [Categories] cat ON cat.CategoryID = sd.CategoryID "
    totalSQL = totalSQL & "WHERE s.SaleDate >= eligibleClients.ClassDate "
    totalSQL = totalSQL & "AND sd.CategoryID < 21 "
    totalSQL = totalSQL & "AND s.SaleDate >= '" & cSalesRecordedStartDate & "' AND s.SaleDate <= '" & cSalesRecordedEndDate & "' "
    
    ' If we are filtering only by sales in the original revenue category
    if filterByRevenueCategory then
        totalSQL = totalSQL & "AND sd.CategoryID = eligibleClients.RevenueCategory "
    end if
    
    totalSQL = totalSQL & "ORDER BY s.ClientID"

    'response.write totalSQL
    'response.End
    
    rsTrainer.CursorLocation = 3
    rsTrainer.open totalSQL, cnWS
    Set rsTrainer.ActiveConnection = Nothing
    
    'Abort quick if there are no results and we're not showing them
    if rsTrainer.RecordCount < 1 AND NOT showNoResults then
        rsTrainer.Close
        exit function
    end if
    
    if request.form("frmExpReport")<>"true" then
        response.Write("<tr><td class=""mainText"" colspan=""2"" valign=""top"" style=""border-color:" & session("pageColor4") & "; border-width: 1px; border-style: solid;"">")
        response.Write("<table class=""TrainerResults mainText"" width=""100%"" cellspacing=""0"" cellpadding=""5px"">")
    else ' Simplified version without style
        response.Write("<tr><td class=""mainText"" colspan=""2"" valign=""top"">")
        response.Write("<table class=""TrainerResults mainText"" width=""100%"" cellspacing=""0"" cellpadding=""5px"">")
    end if
        
    ' Trainer line, need to pull trainer name from cache. 
    response.Write("<tr><td colspan=""3"" align=""left"" class=""whiteHeader"" ")
    if request.form("frmExpReport")<>"true" then
        response.Write("style=""background-color:" & session("pageColor4") & ";""")
    end if
    response.Write(">" & xssStr(allHotWords(6)) & ": ")
    response.Write(trainerName & "</td></tr>")
    
    ' Write out nothing for this client
    if rsTrainer.RecordCount < 1 AND showNoResults then
        response.Write("<tr><td colspan=""3"" align=""center"">No Results.</td></tr>")
        response.Write("</table></td></tr>")
        rsTrainer.close
        exit function
    end if
    
    ' Write out the legend
    %>
    <tr>
        <td><strong>Client ID</strong></td>
        <td><strong>Name</strong></td>
        <td><strong>Initial Session Delivered</strong></td>
    </tr>
    <tr>
        <td>Subsequent Series Purchased</td>
        <td>Sale Date</td>
        <td>Dollar Value</td>
    </tr>
    
    <%
    
    ' Loop through the results.
    ' For each client, print their name on the first line, and then print each line item sale
    ' when the client changes, write out the first line
    dim currentClient : currentClient = 0
    dim totalSales : totalSales = 0
    Do While NOT rsTrainer.EOF
        ' Increment the total sales
        if NOT isNULL(rsTrainer("LinePrice")) then
            totalSales = totalSales + CLNG(rsTrainer("LinePrice"))
        end if
        
        ' Output the client line
        if currentClient <> CLNG(rsTrainer("ClientID")) then
            currentClient = CLNG(rsTrainer("ClientID"))
            response.Write("<tr><td colspan=3>&nbsp;</td></tr>")
            response.Write("<tr><td><strong>" & rsTrainer("ClientID") & "</strong></td>")
            response.Write("<td><strong>" & rsTrainer("Name") & "</strong></td>")
            response.Write("<td><strong>" & rsTrainer("SessionDelivered") & "</strong></td></tr>")
        end if
        
        ' Write out the line items
        response.Write("<tr><td>" & rsTrainer("Description") & "</td>")
        response.Write("<td>" & rsTrainer("SaleDate") & "</td>")
        response.Write("<td>" & FmtCurrency(rsTrainer("LinePrice")) & "</td></tr>")
        rsTrainer.MoveNext 
    loop
    
    ' Output the total sales for this trainer
    response.Write("<tr><td colspan=3 align=""right"" class=""whiteHeader"" ")
    if request.form("frmExpReport")<>"true" then
        response.Write("style=""background-color:" & session("pageColor4") & ";""")
    end if
    response.Write(">Total " & trainerName & " " & xssStr(allHotWords(219)) & ": " & FmtCurrency(totalSales) & "</td></tr>")
    
    response.Write("</table></td></tr>")
    response.Write("<tr><td>&nbsp;</td></tr>") ' Little space between trainers
    
    rsTrainer.close

end function

%>
