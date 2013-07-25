<%@ CodePage=65001 %>
<% Option Explicit %>
<!-- #include file="init.asp" -->
<%
'Dim SessionFarm
'set SessionFarm = Server.CreateObject("SessionFarm.SFSession")
response.charset="utf-8"
%>

		<!-- #include file="inc_dbconn.asp" -->
		<!-- #include file="inc_i18n.asp" -->
<%
dim pLocationID, rsEntry
set rsEntry = Server.CreateObject("ADODB.Recordset")
if isNum(Request.QueryString("id")) then
	pLocationID = Request.QueryString("id")
else
	pLocationID = 0
end if
%>


<!-- #include file="frame_bottom.asp" -->
<!-- #include file="pre.asp" -->
<!-- #include file="inc_date_ctrl.asp" -->
<%= js(array("mb")) %>


<% pageStart %>
<table height="100%" width="90%" cellspacing="0" class="center">
  <tr> 
<%      
strSQL = "SELECT LocationName, Phone, Ext, Address, Email FROM Location WHERE LocationID=" & pLocationID
rsEntry.CursorLocation = 3
rsEntry.open strSQL, cnWS
Set rsEntry.ActiveConnection = Nothing
if NOT rsEntry.EOF then
%>
      <td class="center-ch" valign="top" height="100%" width="100%"> <br />
        <table cellspacing="0" width="90%" height="100%">
          <tr> 
            <td class="headText" align="left"><b><%=session("StudioName")%> @ <%=rsEntry("LocationName")%></b></td>
          </tr>
          <tr> 
            <td valign="top" class="mainTextBig center-ch"> 
              <table class="mainText" width="50%" cellspacing="0">
                <tr >
                  <td class="mainText" width="20%" valign="top">&nbsp;</td>
                  <td class="mainText" colspan="2" valign="top" width="75%">&nbsp;</td>
                </tr>
                <tr > 
                  <td class="mainText" width="20%" valign="top">&nbsp;</td>
                  <td class="mainText" colspan="2" valign="top" width="75%">&nbsp;</td>
                </tr>
                <tr > 
                  <td class="mainText" width="20%" valign="top"><b><%=xssStr(allHotWords(86))%>:</b></td>
                  <td class="mainText" colspan="2" valign="top" width="75%"> 
<%
				  	if NOT isNull(rsEntry("Address")) then
						if Trim(rsEntry("Address"))<>"" then
							response.write rsEntry("Address")
						end if
					end if
%>
                  </td>
                </tr>
                <tr > 
                  <td class="mainText" width="20%" valign="top"><b></b></td>
                  <td class="mainText" colspan="2" valign="top" width="75%">&nbsp;</td>
                </tr>
<%
				if NOT isNull(rsEntry("Phone")) then
					dim ext : ext = ""
					if NOT IsNull(rsEntry("Ext")) then
						ext = " Ext: " & rsEntry("Ext")
					end if
					if rsEntry("Phone")<>"" then
%>
                <tr> 
                  <td class="mainText" width="20%" valign="top"><b><%=xssStr(allHotWords(85))%>:</b></td>
                  <td class="mainText" colspan="2" valign="top" width="75%"><%=FmtPhoneNum(rsEntry("Phone")) & ext %></td>
                </tr>
<%
					end if
				end if

				if NOT isNull(rsEntry("Email")) then
					if rsEntry("Email")<>"" then
%>
                <tr > 
                  <td class="mainText" width="20%" valign="top"><b><%=xssStr(allHotWords(87))%>:</b></td>
                  <td class="mainText" colspan="2" valign="top" width="75%"><a href="mailto:<%=rsEntry("Email")%>"><%=rsEntry("Email")%></a></td>
                </tr>
<%
					end if
				end if

end if
rsEntry.close
%>
                <tr > 
                  <td class="mainText" width="20%" valign="top">&nbsp;</td>
                  <td class="mainText" colspan="2" valign="top" width="75%"> <br />
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

%>
