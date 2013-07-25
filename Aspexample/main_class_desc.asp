<%@ CodePage=65001 %>
<!-- #include file="init.asp" -->
<!-- #include file="json/JSON.asp" -->
<%
'Dim SessionFarm
'set SessionFarm = Server.CreateObject("SessionFarm.SFSession")
response.charset="utf-8"
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<%
Dim Image
Set Image = Server.CreateObject("csImageFile.Manage")
%>
		<!-- #include file="inc_i18n.asp" -->
		<!-- #include file="inc_localization.asp" -->			
		<!-- #include file="inc_tinymcesetup.asp" -->
<%

dim phraseDictionary
set phraseDictionary = LoadPhrases("ClassdescriptionPage", 20)

dim rsClass
dim strRequestQuery, cView
Dim clsNotes, clsDescription, clsName, clsRreReq, CltBadPrereq, pp_Prerequisite

CltBadPrereq = ""
CltBadPrereq = request.querystring("badprereq")

if isNum(Request.QueryString("ID")) then
	strRequestQuery =  CLNG(Request.QueryString("ID"))
else
	strRequestQuery =  -999
end if

set rsClass = Server.CreateObject("ADODB.Recordset")

strSQL = "SELECT ClassNotes, ClassDescriptionID, ClassName, ClassDescription, ClassPrereq FROM tblClassDescriptions WHERE ClassDescriptionID=" & strRequestQuery
rsClass.CursorLocation = 3
rsClass.open strSQL, cnWS
Set rsClass.ActiveConnection = Nothing
if NOT rsClass.EOF then
	
	'Sanitize Class Notes, Description, and PreReq Notes	
	clsNotes = HtmlPurifyForDisplay(rsClass("ClassNotes"))		
	clsDescription = HtmlPurifyForDisplay(rsClass("ClassDescription"))
	clsPreReq = HtmlPurifyForDisplay(rsClass("ClassPrereq"))
	
	clsName = rsClass("ClassName")
end if
rsClass.close
%>
<!-- #include file="pre.asp" -->
	<!-- #include file="frame_bottom.asp" -->
	<!-- #include file="inc_back_ctrl.asp" -->
	<!-- #include file="inc_date_ctrl.asp" -->

<%= js(array("mb")) %>
<style type="text/css">
    #suContainerDiv p
    {
        padding-left: 10px;
    }
    #classDesc
    {
        float:left;
        width:70%;
        margin-right:10px;
    }
    #imgDiv
    {
        float:left;
    }

	
	</style>
</head>
<body>
<!-- #include file="inc_cm_header_bar.asp" -->
<% ShowCMHeader %> 
<% pageStart %>
<div id="suContainerDiv">
	<h1><%=clsName%></h1>
	<div id="mainClassDescInner" class="section" style="width:400px">
		<table width="100%" cellspacing="0">
       
<%if CltBadPrereq<>"" then %>
			<tr>
				<td>
					<h3 style="color:#990000;"><%=DisplayPhrase(phraseDictionary,"Resnotmade")%></h3>
					<p style="color:#990000;">
						<%=clsName%>&nbsp;<%=DisplayPhrase(phraseDictionary,"Noprereq1")%>&nbsp;<%=session("StudioName")%>.
						<%=DisplayPhrase(phraseDictionary,"Noprereq2")%>&nbsp;<%=session("StudioName")%>.
					</p>
				</td>
			</tr>
<% end if %>

			<tr> 
				<td> 
					<h3><%= xssStr(allHotWords(65)) %></h3>
					<div  style="width:100%">
						<div id="classDesc" class="userHTML"><%=clsDescription%></div>
						<% 
					' check for image
					if Image.FileExists(studio_path & "\" & session("studioShort") & "\reservations\" & strRequestQuery & ".jpg") then 
		%>
						<div id="imgDiv">
						<img src="<% response.write "http" & addS & "://" & request.servervariables("SERVER_NAME") & "/studios" & session("ClusterID") & "/" & Session("studioShort") & "/reservations/" & strRequestQuery & ".jpg"%>?imageVersion=<%=session("imageVersion")%>">
						</div>
		<%
					end if
		%>
					</div>
					<div style="float:left;">
	<%
				if clsPreReq<>"" then
					response.write("<h3>" & DisplayPhrase(phraseDictionary,"Prerequisitenotes") & "</h3>")
					response.write("<div  class="" userHTML"">" & clsPreReq & "</div>")
				end if

			strSQL = "SELECT [Student Types].TypeName, tblClassPrereq.ClassDescriptionID FROM [Student Types] INNER JOIN tblClassPrereq ON [Student Types].TypeID = tblClassPrereq.TypeID WHERE (tblClassPrereq.ClassDescriptionID = " & strRequestQuery & ")"
			rsClass.CursorLocation = 3
			rsClass.open strSQL, cnWS
			Set rsClass.ActiveConnection = Nothing
			if NOT rsClass.EOF then
				response.write("<h3>"& DisplayPhrase(phraseDictionary,"Prerequisite") &" </h3>")
				do while not rsClass.EOF
					if CltBadPrereq<>"" then
						response.write "&nbsp;&nbsp;<span style=""color:#990000;"">" & rsClass("TypeName") & "</span><br />"
					else
						response.write "&nbsp;&nbsp;" & rsClass("TypeName")
					end if
					rsClass.MoveNext
				loop
			end if
			rsClass.close
	%>
			<% 	if clsNotes<>"" then %>
						<h3><%= xssStr(allHotWords(90)) %>&nbsp;</h3>
						<div class="userHTML"><%=clsNotes%></div>
			<% end if %>
					</div>
				</td>
			</tr>
			<tr>
				<td>
				  <table width="50" cellspacing="0" height="20" class="center">
					<tr> 
					  <td class="center-ch" valign="middle"><input type="button" onclick="javascript:goBack();" value="< Back" /></td> 
					</tr>
				  </table>
				</td>
			</tr>
		</table>
	</div>
</div>
<% pageEnd %>
<%
	set rsClass = nothing
%>
<!-- #include file="post.asp" -->
