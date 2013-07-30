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
if not Session("Pass") OR Session("Admin")="false" OR NOT validAccessPriv("RPT_CLTWO") then 
%>
<script type="text/javascript">
	alert("You are not authorized to view this page.\nPlease login to an account with Access Privileges granted for this feature.");
	javascript:history.go(-1);
</script>
<%
else
%>
		<!-- #include file="../inc_i18n.asp" -->
		<!-- #include file="inc_rpt_tagging.asp" -->

		
<!-- #include file="pre.asp" -->
<!-- #include file="frame_bottom.asp" -->



<%= js(array("mb", "adm/adm_rpt_acct_y")) %>



	<!-- #include file="../inc_date_ctrl.asp" -->




<% pageStart %>
<table height="100%" width="<%=strPageWidth%>" cellspacing="0">    
<tr> 
     <td class="center-ch" valign="top" height="100%" width="100%"> <br />
        <table cellspacing="0" width="90%" height="100%">
          <tr> 
            <td class="headText" align="left"><b><%=session("ClientHW")%>s with Logins</b> </td>
          </tr>
          <tr>
            <td valign="top" class="mainText right">&nbsp;</td>
          </tr>
          <tr> 
            <td valign="top" class="mainTextBig center-ch"> 
              <table class="mainText" width="95%" cellspacing="0">
				<tr>
					<td class="center-ch textSmall">
					<form name="frmParameter" action="adm_rpt_acct_y.asp" method="POST">
						
						<input type="hidden" name="frmTagClients" value="false">
						<input type="hidden" name="frmTagClientsNew" value="false">
							<b>Show <%=session("ClientHW")%> Created Only: </b><input type="checkbox" name="optCltCreated" onClick="submit();" <% if request.form("optCltCreated")="on" then response.write " checked" end if %>>&nbsp;&nbsp;
							<%if not Session("Pass") OR Session("Admin")="false" OR NOT validAccessPriv("RPT_TAG") then
							 else%>
							 <b>Tagged <%=session("ClientHW")%>s Only (<span id="TaggedCount"><%=getTaggedCount()%></span>):</b>
								<input type="checkbox" name="optFilterTagged" <%if request.form("optFilterTagged")="on" then response.write "checked" end if%> onClick="toggleTagButton(); submit();">
								<br />
								<input type="button"  name="TagAdd" value="Tag <%=session("ClientHW")%>s (Add)" onClick="tagClientsAdd();" <%if request.form("optFilterTagged")="on" then response.write "disabled" end if%>>
								<input type="button"  name="TagNew" value="Tag <%=session("ClientHW")%>s (New)" onClick="tagClientsNew();">
								&nbsp;&nbsp;&nbsp;
								
								<br /><br />
							<%end if%>
					
					</td>
				
				</tr>

		
				<tr >
					<td  colspan="2" valign="top" class="mainTextBig center-ch">
		<%
			if 1=0 then
			'if request.form("frmTagClients")="true" then
				on error resume next
				strSQLTag = "INSERT INTO tblClientTag "
				strSQLTag = strSQLTag & "SELECT ClientID, " 
				if session("mvarUserID")<>"" then
					strSQLTag = strSQLTag & session("mvarUserID") & " "
				else
					strSQLTag = strSQLTag & "0 "
				end if
				strSQLTag = strSQLTag & "FROM Clients WHERE NOT (LoginName IS NULL) AND [Deleted]=0 "
				if request.form("optCltCreated")="on" then
					strSQLTag = strSQLTag & "AND Location=98 "
				end if 
				cnWS.execute strSQLTag
		%>
				<script type="text/javascript">
					alert("Resulting <%=jsEscDouble(allHotWords(12))%>s are tagged.");
				</SCRIPT>
		<%
			end if
		%>

<!--#include file="adovbs.inc"-->
<%
Const intPageSize = 50
Dim intCurrentPage, objConn, objRS, opener
dim intTotalPages, intI
Dim strTempName, intCount



if Request.ServerVariables("Content_Length") = 0 Then
  intCurrentPage = 1
Else
  intCurrentPage = CInt(Request.Form("CurrentPage"))
  Select Case Request.Form("Submit")
   Case "Previous"
      intCurrentPage = intCurrentPage -1
   Case "Next"
      intCurrentPage = intCurrentPage + 1
  End Select
End if
%>
			<input type="hidden" name="CurrentPage" value="<%=intCurrentPage%>">
		</form>
<%

         set objRS = Server.CreateObject("ADODB.Recordset")
         objRs.CursorLocation = adUseClient
         objRs.CursorType = adOpenStatic
         objRs.CacheSize = intPageSize
 
        'create SQL select query string
        strSQL = "SELECT  Clients.ClientID, FirstName, LastName, LoginName, Address, HomePhone, WorkPhone FROM Clients "
		if request.form("optFilterTagged")="on" then
			strSQL = strSQL & " INNER JOIN tblClientTag ON (tblClientTag.ClientID = CLIENTS.ClientID "
			if session("mVarUserID")<>"" then
				strSQL = strSQL & " AND tblClientTag.smodeID = " & session("mVarUserID")
			end if
			strSQL = strSQL & " ) "
		end if
		strSQL = strSQL & "WHERE NOT (LoginName IS NULL) AND [Deleted]=0 "
		if request.form("optCltCreated")="on" then
			strSQL = strSQL & "AND OnlineSignUp=1 "
		end if
		
		if request.form("frmTagClients")="true" then
			if request.form("frmTagClientsNew")="true" then
				clearAndTagQuery(strSQL)
			else
				tagQuery(strSQL)
			end if
		end if
	
		strSQL = strSQL & "ORDER BY lastname, firstname"
	
        objRs.Open strSQL, cnWS, , , adCmdText
        objRs.PageSize = intPageSize
        if Not(objRS.EOF) then objRS.AbsolutePage = intCurrentPage
        intTotalPages = objRS.PageCount

        intCount = 0


        %>
			<!--<input type="hidden" name="CurrentPage" value="<%=intCurrentPage%>">
		</form>-->


<b>Click on a <%=session("ClientHW")%>'s name to edit their Contact Information</b> 
<br />
<br />

            <Form action="<%=Request.ServerVariables("Script_Name") %>" method="post">
			
            <input type="hidden" name="CurrentPage" value="<%=intCurrentPage%>">
            <%
            if intCurrentPage > 1 then
            %>
                      <input type="Submit" name="submit" Value="Previous">
            <%
            End if

            if intCurrentPage <> intTotalPages Then
            %>
                      <input type="submit" name="submit" value="Next">
            <%
            end if
            %>
            </form>


            Page <%= intCurrentPage %> of <%= intTotalPages %>
        <br />

        
                    <table class="mainText" cellspacing="0">
                      <tr style="background-color:<%=session("pageColor2")%>;">
                        <td colspan=6><img height="2" width="100%" src="<%= contentUrl("/asp/adm/images/trans.gif") %>"></td>
                      </tr>
                      <tr style="background-color:<%=session("pageColor4")%>;"> 
                        <td class="whiteHeader" width="55">&nbsp;<b>Count</b></td>
                        <td class="whiteHeader" width="140">&nbsp;<b><%=xssStr(allHotWords(40))%></b></td>
                        <td class="whiteHeader" width="160">&nbsp;<b><%=xssStr(allHotWords(41))%></b></td>
                        <td class="whiteHeader" width="160">&nbsp;<b><%=xssStr(allHotWords(46))%></b></td>
                        <td class="whiteHeader" width="60">&nbsp;<b><%=xssStr(allHotWords(82))%><b>&nbsp;</td>
                      </tr>
                      <tr style="background-color:<%=session("pageColor2")%>;">
                        <td colspan=6><img height="2" width="100%" src="<%= contentUrl("/asp/adm/images/trans.gif") %>"></td>
                      </tr>
                      <%
   	rowcount = 0

	if NOT objRS.EOF then		
        For intI = 1 to objRS.PageSize

            intCount = intCount + 1

	            if rowcount=0 then
%>
                      <tr bgcolor=#F2F2F2> 
                        <%
               rowcount = 1
            else
%>
                      <tr bgcolor=#FAFAFA> 
                        <%
               rowcount = 0
            end if
				
            Response.Write "<td align=""center"">" & intCount  &  ". &nbsp;&nbsp;</td>"
            Response.Write "<td><a href=""main_info.asp?id=" & objRS("ClientID") & "&fl=true"">"
            Response.Write objRS("FirstName") & "&nbsp;" 
            Response.Write objRS("lastname") 
            Response.Write "</a></td>"

            Response.Write "<td>"
            Response.Write objRS("LoginName") & "&nbsp;"
            Response.Write "</td>"

            Response.Write "<td>"
            Response.Write Left(objRS("Address"),24) & "&nbsp;"
            Response.Write "</td>"

            Response.Write "<td align=""center"">"
            Response.Write FmtPhoneNum(objRS("HomePhone")) & "&nbsp;"
            Response.Write "</td>"

            'Response.Write "<td>"
            'Response.Write FmtPhoneNum(objRS("WorkPhone")) & "&nbsp;"
            'Response.Write "</td>"


            Response.Write "</tr>"

            objRs.MoveNext
            if objRS.EOF Then Exit For

    
         Next
	end if ''''EOF		 
%>
                    </table>
<%
            objRS.Close : cnWS.Close
            set objRS = nothing : set cnWS = Nothing


	%>
  <Form action="<%=Request.ServerVariables("Script_Name") %>" method="post">
    
    <%
            if intCurrentPage > 1 then
            %>
                      <input type="Submit" name="submit" Value="Previous">
    <%
            End if

            if intCurrentPage <> intTotalPages Then
            %>
                      <input type="submit" name="submit" value="Next">
    <%
            end if
            %>
  </form>


  Page <%= intCurrentPage %> of <%= intTotalPages %> <br />

                  </td>
                </tr>
              </table>
            </td>
          </tr>
          
        </table>
    </td>
    </tr>

		</table>
<% pageEnd %>
<!-- #include file="post.asp" -->

<%
	''

	end if
%>
