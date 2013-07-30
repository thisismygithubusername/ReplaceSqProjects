<%@CodePage=65001 %>
<% Option Explicit %>
<!-- #include file="init.asp" -->
<%
' The Costumes Report has two major modes: Class and Student. Class is like a summary report,
' it shows all students in a given class and displays each student, (their measurements if
' those are used), and the items(s) purchased for that class.
' In the Student report, a single student's purchase history is shown. This report is driven
' by the querystring arguments classid and clientid, with classid taking preference if both are
' given.

'Dim SessionFarm
'set SessionFarm = Server.CreateObject("SessionFarm.SFSession")
response.charset="utf-8"
%>
<%
	dim rsEntry, rsEntry2
	set rsEntry = Server.CreateObject("ADODB.Recordset")
	set rsEntry2 = Server.CreateObject("ADODB.Recordset")
	Dim stFilename
	%>
	<!-- #include file="inc_accpriv.asp" -->
	<%
	if not Session("Pass") OR Session("Admin")="false" OR NOT validAccessPriv("RPT_MARKETING") then 
		%>
		<script type="text/javascript">
		    alert("You are not authorized to view this page.\nPlease login to an account with Access Privileges granted for this feature.");
		    javascript: history.go(-1);
		</script>
		<%
	else
		%>
			<!-- #include file="../inc_i18n.asp" -->
			<!-- #include file="inc_rpt_tagging.asp" -->
			<!-- #include file="inc_utilities.asp" -->
			<!-- #include file="inc_rpt_save.asp" -->
			<!-- #include file="inc_hotword.asp" -->
		<%
		Dim ap_view_all_locs, ss_UseAssignedEquipment, ss_UseClientMeasurements
		Dim cClassID, cClientID, cMode
		
		' Pull up studio settings
		ss_UseAssignedEquipment = checkStudioSetting("tblGenOpts", "UseAssignedEquipment")
		ss_UseClientMeasurements  = checkStudioSetting("tblGenOpts", "UseClientMeasurements")
		
		' -------------------------------------------------------------------------------------------
		' Parse querystring/POST data into variables to control the display
		
		' Parse out the class ID
		If request.QueryString("pClsID") <> "" then
		    cClassID = CLNG(request.QueryString("pClsID"))
		    cMode = "class"
		elseif request.Form("requiredtxtClassID") <> "" then
		    cClassID = CLNG(request.Form("requiredtxtClassID"))
		    cMode = "class"
		else
		    cClassID = 0
		    cMode = "student" ' If we don't have a class, we must be doing a student
		end if
		
		' Parse out the Client ID
		If request.QueryString("clientid") <> "" then
		    cClientID = CLNG(request.QueryString("clientid"))
		elseif request.Form("requiredtxtClientID") <> "" then
		    cClientID = CLNG(request.Form("requiredtxtClientID"))
		else
		    cClientID = 0
		end if
		
		' Show all the head and options if we are not generating an excel report
		if NOT request.form("frmExpReport")="true" then
		%>
<!-- #include file="pre.asp" -->
			    <!-- #include file="frame_bottom.asp" -->
			
<%= js(array("mb", "MBS", "calendar" & dateFormatCode, "adm/adm_rpt_costumes")) %>
			    <script type="text/javascript">
			    function exportReport() {
				    document.frmParameter.frmExpReport.value = "true";
				    document.frmParameter.frmGenReport.value = "true";
				    <% iframeSubmit "frmParameter", "adm_rpt_costumes.asp" %>
			    }
			    </script>
			    
			    <!-- #include file="../inc_date_ctrl.asp" -->
    
<% pageStart %>
			<table height="100%" width="<%=strPageWidth%>" cellspacing="0">    
			<tr> 
			  <td class="center-ch" valign="top" height="100%" width="100%"> 
				<table cellspacing="0" width="90%" height="100%">
				  <tr>
					<td class="headText" align="left" valign="top">
					  <table width="100%" cellspacing="0">
						<tr>
						  <td class="headText" valign="bottom"><b id="costumeHeader"><%= pp_PageTitle("Costumes") %></b></td>
						  <td valign="bottom" class="right" height="26"></td>
						</tr>
					  </table>
					</td>
				  </tr>
				  <tr height="30"> 
					<td  valign="bottom" class="center-ch headText">
					  <form name="frmParameter" action="adm_rpt_prospects.asp" method="POST">
					  <input type="hidden" name="frmGenReport" value="">
					  <input type="hidden" name="frmExpReport" value="">
					  </form>
					</td>
				  </tr>
				  <tr> 
					<td valign="top" class="mainTextBig center-ch"> 
					<table class="mainText" width="95%" cellspacing="0">
<%                    ' Switch on the mode.
                    if cMode = "class" then
                        ' First, select information about each product this class is supposed to purchase
                        strSQL = "SELECT MIN(p.Description) AS Description, p.ProductGroupID FROM tblClassProductGroups tcpg "
                        strSQL = strSQL & "INNER JOIN PRODUCTS p ON tcpg.ProductGroupID = p.ProductGroupID "
                        strSQL = strSQL & "WHERE tcpg.ClassID = " & cClassID & " "
                        strSQL = strSQL & "GROUP BY p.ProductGroupID"
                        
                       response.write debugSQL(strSQL, "SQL")
						rsEntry.CursorLocation = 3
						rsEntry.open strSQL, cnWS
						set rsEntry.ActiveConnection = Nothing
						
						' Store this information into some arrays, we'll be going through these alot
						redim requiredProductNames(rsEntry.RecordCount - 1)
						redim requiredProductGroups(rsEntry.RecordCount - 1)
						dim i : i = 0
						dim lastClient
						
						do while NOT rsEntry.EOF
						    requiredProductNames(i) = rsEntry("Description")
						    requiredProductGroups(i) = rsEntry("ProductGroupID")
						    i = i + 1
						    rsEntry.MoveNext
						loop
						rsEntry.Close
						
						' Print out the header line, need to find the class name
						strSQL = "SELECT ClassName, ClassDateStart, ClassDateEnd, TrFirstName, TrLastName, DisplayName, TrainerID FROM tblClasses "
						strSQL = strSQL & "INNER JOIN tblClassDescriptions ON tblClassDescriptions.ClassDescriptionID = tblClasses.DescriptionID "
						strSQL = strSQL & "INNER JOIN TRAINERS ON TRAINERS.TrainerID = tblClasses.ClassTrainerID "
						strSQL = strSQL & "WHERE ClassID = " & cClassID & " "
						
					response.write debugSQL(strSQL, "SQL")
						rsEntry.CursorLocation = 3
						rsEntry.open strSQL, cnWS
						set rsEntry.ActiveConnection = Nothing
						
						if NOT rsEntry.EOF then
					        response.Write("<tr><td class=""headText""><strong><a href=""adm_cs_e.asp?classID=" & cClassID & """>" & rsEntry("ClassName") & "</a>&nbsp;w/&nbsp;<a href=""adm_trn_e.asp?trnID=" & rsEntry("TrainerID") & """>" & FmtTrnNameNew(rsEntry, false) & "</a>&nbsp;")
					        if rsEntry("ClassDateStart") = rsEntry("ClassDateEnd") then
					            response.Write("(" & FmtDateShort(rsEntry("ClassDateStart")) & ")")
					        else
					            response.Write("(" & FmtDateShort(rsEntry("ClassDateStart")) & "&nbsp;-&nbsp;" & FmtDateShort(rsEntry("ClassDateEnd")) & ")")
					        end if
					        response.Write("</strong></td></tr>")
					    end if
					    rsEntry.Close
                    
                        ' We are displaying the information for an entire class. Find all the clients
                        ' in the class, and display the required products and what each student has
                        ' purchased.
                        strSQL = "SELECT DISTINCT c.FirstName, c.LastName, c.ClientID, p.Description, p.ProductID, p.ProductGroupID, Sizes.SizeName, Colors.ColorName "
                        
                        ' If we are using Measurements, select them as well
                        if ss_UseClientMeasurements then
                            strSQL = strSQL & ", c.MeasurementsTaken, c.Height, c.Bust, c.Waist, c.Hip, c.Girth, c.Inseam, c.Head, c.Shoe, c.Tights "
                        end if
                        
                        strSQL = strSQL & "FROM [VISIT DATA] vd "
                        strSQL = strSQL & "INNER JOIN CLIENTS c ON vd.ClientID = c.ClientID "
                        
                        ' Get the products that were purchased. We'll still have null entries for each person in the class
                        strSQL = strSQL & "LEFT OUTER JOIN [Sales Details] sd "
                        strSQL = strSQL & "INNER JOIN Colors INNER JOIN PRODUCTS AS p ON Colors.ColorID = p.ColorID INNER JOIN Sizes ON Sizes.SizeID = p.SizeID ON sd.ProductID = p.ProductID ON vd.ClientID = sd.RecClientID AND vd.ClassID = sd.ClassID "
                        
                        strSQL = strSQL & "WHERE vd.ClassID = " & cClassID & " "
                        strSQL = strSQL & "ORDER BY c.ClientID, c.LastName, c.FirstName"
                        
                       response.write debugSQL(strSQL, "SQL")
						rsEntry.CursorLocation = 3
						rsEntry.open strSQL, cnWS
						Set rsEntry.ActiveConnection = Nothing
						
						if rsEntry.EOF then
							response.Write("<tr class=""headText"" align=""center"">")
							response.Write("    <td colspan=""7"">No <script type=""text/javascript"">document.write('" & jsEscSingle(allHotWords(12)) & "'); // Clients Enrolled</script> </td>")
							response.Write("</tr>")
					    else
					        ' Set up the grouping variables.
					        lastClient = 0
					        
					        ' Loop through the record set. The first entry for a new client may or may not be null with respect
					        ' products, in any case, pull out the measurements
					        do while NOT rsEntry.EOF
					            ' It's a new client
				                lastClient = CLNG(rsEntry("ClientID"))
				                
				                ' Start a new table row (and table)
				                response.Write("<tr><td><table class=""mainText"" width=""100%"" cellspacing=""0"" cellpadding=""5px"">" & VbCrLf)
				                
				                ' Print out the header row, include measurements if we're using them
				                response.Write("<tr class=""whiteHeader""align=""center"" style=""background-color:" & session("pageColor4") & ";"" style=""border-color:" & session("pageColor4") & "; border-width: 1px; border-style: solid;"">")
				                if ss_UseClientMeasurements then
				                    response.Write("<td align=""left"" width=""20%""><a style=""color: #FFFFFF;"" href=""adm_clt_profile.asp?id=" & rsEntry("ClientID") & """>" & rsEntry("LastName") & ",&nbsp;" & rsEntry("FirstName") & "</a></td>")
				                    response.Write("<td>" & getHotWord(236) & ":&nbsp;" & rsEntry("Height") & "</td>")
				                    response.Write("<td>" & getHotWord(237) & ":&nbsp;" & rsEntry("Bust") & "</td>")
				                    response.Write("<td>" & getHotWord(229) & ":&nbsp;" & rsEntry("Waist") & "</td>")
				                    response.Write("<td>" & getHotWord(230) & ":&nbsp;" & rsEntry("Hip") & "</td>")
				                    response.Write("<td>" & getHotWord(231) & ":&nbsp;" & rsEntry("Girth") & "</td>")
				                    response.Write("<td>" & getHotWord(232) & ":&nbsp;" & rsEntry("Inseam") & "</td>")
				                    response.Write("<td>" & getHotWord(233) & ":&nbsp;" & rsEntry("Head") & "</td>")
				                    response.Write("<td>" & getHotWord(234) & ":&nbsp;" & rsEntry("Shoe") & "</td>")
				                    response.Write("<td>" & getHotWord(235) & ":&nbsp;" & rsEntry("Tights") & "</td>")
				                    response.Write("<td>Taken: " & FmtDateShort(rsEntry("MeasurementsTaken")) & "</td>")
				                else
				                    response.Write("<td align=""left"" colspan=""11""><a style=""color: #FFFFFF;"" href=""adm_clt_profile.asp?id=" & rsEntry("ClientID") & """>" & rsEntry("LastName") & ",&nbsp;" & rsEntry("FirstName") & "</a></td>")
				                end if
				                response.Write("</tr>")
                                
                                ' Read in the products in the result set.
                                ' If the ProductID is null, then the client hasn't purchased anything.
                                if NOT isNull(rsEntry("ProductID")) then
                                    ' output potential products
                                    
                                    ' Read all the products this client has bought, we'll compare them
                                    ' to the ones he should have bought.
                                    dim continue : continue = true
                                    
                                    redim clientProductName(-1)
                                    redim clientProductGroup(-1)
                                    redim clientProductSize(-1)
                                    redim clientProductColor(-1)
                                    
                                    ' Read in all the products for this client
                                    do while continue
                                        if NOT rsEntry.EOF then
                                            if CLNG(rsEntry("ClientID")) = lastClient then
                                                ' Finally, we get to a known good state
                                                ' Read in the product information stored here
                                                redim preserve clientProductName(UBound(clientProductName) + 1)
                                                redim preserve clientProductGroup(UBound(clientProductGroup) + 1)
                                                redim preserve clientProductSize(UBound(clientProductSize) + 1)
                                                redim preserve clientProductColor(UBound(clientProductColor) + 1)
                                                clientProductName(UBound(clientProductName)) = rsEntry("Description")
                                                clientProductGroup(UBound(clientProductGroup)) = rsEntry("ProductGroupID")
                                                clientProductSize(UBound(clientProductSize)) = rsEntry("SizeName")
                                                clientProductColor(UBound(clientProductColor)) = rsEntry("ColorName")
                                                
                                                ' Onward and upward
                                                rsEntry.MoveNext
                                            else
                                                continue = false
                                            end if
                                        else
                                            continue = false
                                        end if
                                    loop
                                    
                                    ' We've collected all their purchased items
                                    ' Loop through the required items, match them to the purchased items
                                    ' And print them
                                    dim match
                                    
                                    for i = 0 to UBound(requiredProductGroups)
                                        ' See if productGroup matches anything in clientProductGroup
                                        match = inArray(requiredProductGroups(i), clientProductGroup)

                                        if (match > -1) then
                                            ' Match, product purchased
                                            response.Write("<tr><td>" & clientProductName(match) & "</td>" & VbCrLf)
                                            response.Write("<td>" & clientProductSize(match) & "</td>" & VbCrLf)
                                            response.Write("<td>" & clientProductColor(match) & "</td>" & VbCrLf)
                                            response.Write("<td colspan=""8"">&nbsp;</td></tr>" & VbCrLf)
                                        else
                                            ' No match, product not purchased
                                            response.Write("<tr style=""color:#FF0000""><td>" & requiredProductNames(i) & "</td><td>Not Purchased</td><td colspan=""9""><a href=""main_retail.asp?cltID=" & lastClient & """>Make Purchase</a></td></tr>")
                                        end if
                                    next
                                else 'ProductID is NULL
                                    ' Generic output of all products, list as NONE
                                    for i = 0 to UBound(requiredProductNames)
                                        response.Write("<tr style=""color:#FF0000""><td>" & requiredProductNames(i) & "</td><td >Not Purchased</td><td colspan=""9""><a href=""main_retail.asp?cltID=" & lastClient & """>Make Purchase</a></td></tr>")
                                    next
                                    
                                    ' Move to the next client
                                    rsEntry.MoveNext
                                end if' CLNG(rsEntry("ClientID")) = lastClient
                                
                                ' Close the client table and record
                                response.Write("</table><tr><td>&nbsp;</td></tr></td></tr>" & VbCrLf)
                                
                                ' Done with this client, we will be at the next client by now
					        loop ' 
					    
					    ' All looped through
					    end if ' rsEntry.EOF (Initial check for any results)
                        
                    elseif cMode = "student" then
                        ' This is currently not used, but may be in the future. Fill it in here
                        
                        'strSQL = "SELECT c.FirstName, c.LastName, tcd.ClassName, p.Description, Colors.ColorName, Sizes.SizeName "
                        'strSQL = strSQL & "FROM [VISIT DATA] vd "
                        'strSQL = strSQL & "INNER JOIN Sales s ON s.ClientID = vd.ClientID "
                        'strSQL = strSQL & "INNER JOIN [Sales Details] sd ON s.SaleID = sd.SaleID AND sd.ClassID = vd.ClassID "
                        'strSQL = strSQL & "INNER JOIN CLIENTS c ON vd.ClientID = c.ClientID "
                        'strSQL = strSQL & "INNER JOIN PRODUCTS p ON p.ProductID = sd.ProductID "
                        'strSQL = strSQL & "INNER JOIN Colors ON Colors.ColorID = p.ColorID "
                        'strSQL = strSQL & "INNER JOIN Sizes ON Sizes.SizeID = p.SizeID "
                        'strSQL = strSQL & "INNER JOIN tblClasses ON tblClasses.ClassID = vd.ClassID "
                        'strSQL = strSQL & "INNER JOIN tblClassDescriptions tcd ON tcd.ClassDescriptionID = tblClasses.DescriptionID "
                        'strSQL = strSQL & "WHERE vd.ClientID = " & cClientID
                        
                        response.write debugSQL(strSQL, "SQL")
						'rsEntry.CursorLocation = 3
						'rsEntry.open strSQL, cnWS
						'Set rsEntry.ActiveConnection = Nothing

						'if rsEntry.EOF then
						'	response.Write("<tr class=""headText"" align=""center"">")
						'	response.Write("    <td colspan=""7"">No result.</td>")
						'	response.Write("</tr>")
					    'else
					    
					    'end if
                    end if
 %>
                      </table>
                    </td>
				  </tr>		
	            
	          </table>
			</td>
			</tr>
				</table>
<% pageEnd %>
<!-- #include file="post.asp" -->

<%'end of frmexport value check
	
end if ' Got some mismatched IF somewhere, probably related to a nested frmExpReport problem
end if ' It works, though
%>
