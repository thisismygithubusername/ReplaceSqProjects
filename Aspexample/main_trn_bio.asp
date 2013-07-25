<%@ CodePage=65001 %>
<!-- #include file="init.asp" -->
<%
'Dim SessionFarm
'set SessionFarm = Server.CreateObject("SessionFarm.SFSession")
response.charset="utf-8"

Dim Image
Set Image = Server.CreateObject("csImageFile.Manage")
%>
		<!-- #include file="inc_i18n.asp" -->
<%
dim rsTrainers
dim strRequestQuery

if isNum(Request.QueryString("trainer")) then
	strRequestQuery =  CLNG(Request.QueryString("trainer"))
else
	strRequestQuery =  -999
end if
'response.write strRequestQuery

set rsTrainers = Server.CreateObject("ADODB.Recordset")

strSQL= "SELECT TrainerID, Bio FROM TRAINERS WHERE TrainerID =" & strRequestQuery
rsTrainers.CursorLocation = 3
rsTrainers.open strSQL, cnWS
Set rsTrainers.ActiveConnection = Nothing
if rsTrainers.EOF then
	response.write "Trainer Not Found"
	response.End
end if
%>
<html>
<head>
<title><%=Session("StudioName")%> Online</title>
<meta http-equiv="Content-Type" content="text/html">
	<!-- #include file="frame_bottom.asp" -->
	<!-- #include file="inc_back_ctrl.asp" -->
	<!-- #include file="inc_date_ctrl.asp" -->
<%= js(array("mb")) %>
</head>
<body>
<% pageStart %>
<table height="100%" width="<%=strPageWidth%>" cellspacing="0" cellpadding="4">
  <tr> 
    <td class="center-ch" valign="top" height="100%" width="100%"> 
      <br />
      <table width="85%" cellspacing="0" class="center">
        <tr height="100%">
          <td height="100%" class="headText" align="left"><b><%=FmtTrnName(rsTrainers("TrainerID"))%></b></td>
        </tr>
        <tr height="100%"> 
          <td height="100%" class="mainTextBig center-ch"> 
            <table height="100%" class="mainText" width="95%" cellspacing="0" cellpadding="10">
             <tr height="100%"> 
			  	<td><br />
<%
					if Image.FileExists(studio_path & session("studioShort") & "\staff\" & rsTrainers("TrainerID") & "_large.jpg") then
						Image.ReadFile studio_path & session("studioShort") & "\staff\" & rsTrainers("TrainerID") & "_large.jpg"
%>
						<img src="<% response.write "http" & addS & "://" & request.servervariables("SERVER_NAME") & "/studios" & session("ClusterID") & "/" & Session("studioShort") & "/staff/" & rsTrainers("TrainerID") & "_large.jpg?imageversion=" & session("imageVersion")%>">
					<%elseif Image.FileExists(studio_path & session("studioShort") & "\staff\" & rsTrainers("TrainerID") & ".jpg") then
						Image.ReadFile studio_path & session("studioShort") & "\staff\" & rsTrainers("TrainerID") & ".jpg"
%>
						<img src="<% response.write "http" & addS & "://" & request.servervariables("SERVER_NAME") & "/studios" & session("ClusterID") & "/" & Session("studioShort") & "/staff/" & rsTrainers("TrainerID") & ".jpg?imageversion=" & session("imageVersion")%>">
					<% end if %>
				</td>
                <td height="100%" align="left" valign="top"><br />
                  <br />
<%
Dim trnBio

trnBio = rsTrainers("bio")
%>

<%=trnBio%>

<%
if isNull(trnBio) then
%>
No bio yet for this <%=xssStr(allHotWords(6))%>.
<%
end if
%>

                </td>
              </tr>
            </table>
          </td>
          
          </tr>
      </table>
      <br />
      <br />
      <br />
      <br />
      <table width="85%" cellspacing="0" height="20" class="center">
        <tr> 
          <td valign="middle"><input type="button" onclick="javascript:goBack();" value="< Back" /></td> 
        </tr>
      </table>
    </td>
  </tr>
</table>
<% pageEnd %>
</body>
</html>
<%
	cnWS.close
	set cnWS = nothing
	
%>
