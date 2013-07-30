<%@ CodePage=65001 %>
<!-- #include file="init.asp" -->
<%
'Dim SessionFarm : set SessionFarm = Server.CreateObject("SessionFarm.SFSession")
response.charset="utf-8"
%>
<%
    dim rsEntry : set rsEntry = Server.CreateObject("ADODB.Recordset")
    dim rsEntry2 : set rsEntry2 = Server.CreateObject("ADODB.Recordset")
    %>
    <!-- #include file="inc_accpriv.asp" -->
    <!-- #include file="inc_rpt_tagging.asp" -->
    <!-- #include file="../inc_ajax.asp" -->
    <!-- #include file="../inc_val_date.asp" -->
    <!-- #include file="../inc_jquery.asp" -->
    <%
    if NOT Session("Pass") OR Session("Admin")="false" then 'OR NOT validAccessPriv("RPT_DAY") then
        response.write "<script type=""text/javascript"">alert('You are not authorized to view this page.\nPlease login to an account with Access Privileges granted for this feature.');javascript:history.go(-1);</script>"
        cnWS.close
        set cnWS = nothing
        response.end
    end if
    %>
    <!-- #include file="../inc_i18n.asp" -->
    <!-- #include file="inc_hotword.asp" -->
    <%
	
	dim category : category = ""
	if (RQ("category"))<>"" then
		category = RQ("category")
	elseif (RF("category"))<>"" then
		category = RF("category") 
	end if

    if request.form("requiredtxtDateStart")<>"" then
        Call SetLocale(session("mvarLocaleStr"))
            CSDate = CDATE(sqlInjectStr(request.form("requiredtxtDateStart")))
        Call SetLocale("en-us")
    else
        CSDate = DateValue(CDATE(DateAdd("n", Session("tzOffset"),Now)))
    end if

    if request.form("requiredtxtDateEnd")<>"" then
        Call SetLocale(session("mvarLocaleStr"))
            CEDate = CDATE(sqlInjectStr(request.form("requiredtxtDateEnd")))
        Call SetLocale("en-us")
    else
        CEDate = DateValue(CDATE(DateAdd("n", Session("tzOffset"),Now)))
    end if

    ' Hotwords
    'dim arrHW : arrHW = getHotWords(array(80, 81))
        
    dim firstName_hw       : firstName_hw       = allHotWords(80)
    dim lastName_hw        : lastName_hw        = allHotWords(81)

    %>

<% if NOT request.form("frmExpReport")="true" then %>
<!-- #include file="pre.asp" -->
        <!-- #include file="frame_bottom.asp" -->
        
        <%= js(array("MBS", "calendar" & dateFormatCode, "adm/sorttable", "adm/adm_rpt_class_test", "reportFavorites", "plugins/jquery.SimpleLightBox")) %>
		<%= css(array("SimpleLightBox")) %> 
        <script type="text/javascript">
         function exportReport() {
                document.frmParameter.frmGenReport.value = "true";
                document.frmParameter.frmExpReport.value = "true";
                <% iframeSubmit "frmParameter", "adm_rpt_class_tests.asp" %>
         }
        </script>
<% end if %>
        <!-- #include file="css/report.asp" -->

<% if NOT request.form("frmExpReport")="true" then %>
<% pageStart %>
	<% if UseVersionB(TEST_REPORT_LANDING_PAGE) then %>
		<div class="headText breadcrumbs-old">
		<span class="breadcrumb-item"><a href="/reportslandingpage/index"><%=DisplayPhrase(reportPageTitlesDictionary, "Reports")%></a></span>
		<%if category<>"" then%>
		<span class="breadcrumb-item">&raquo;</span>
		<span class="breadcrumb-item active"><a href="/reportslandingpage/<%=category%>category"><%=GetReportsCategoryPagePhrase(category)%></a></span>
		<% end if %>
		<span class="breadcrumb-item">&raquo;</span>
			<%= DisplayPhrase(reportPageTitlesDictionary,"Classtests") %>
		<div id="add-to-favorites">
			<span id="" class="favorite"></span><span class="text"><%= DisplayPhrase(reportPageTitlesDictionary,"Addtofavorites") %></span>
		</div>
		</div>
	<%end if %>

        <div id="container">
<% if NOT UseVersionB(TEST_REPORT_LANDING_PAGE) then %>
		<div id="head" class="headText">
              <%=pp_PageTitle("Class Test")%>
          </div>
<%end if %>
          <div id="options" class="mainText" style="margin: 0 auto;">
            <form name="frmParameter" action="adm_rpt_class_tests.asp" method="POST">
              <input type="hidden" name="frmGenReport" value="" />
              <input type="hidden" name="frmExpReport" value="" />
              <input type="hidden" name="frmCurClientID" value="" />
			<% if UseVersionB(TEST_REPORT_LANDING_PAGE) then %>
				<input type="hidden" name="reportUrl" id="reportUrl" value="<%=Request.ServerVariables("SCRIPT_NAME") %>" />
				<input type="hidden" name="category" value="<%=category%>">
			<% end if %>
              
              <label for="requiredtxtDateStart">
                <input onBlur="if(this.value != ''){validateDate(this, '<%=FmtDateShort(cSDate)%>', true);};" type="text"  id="requiredtxtDateStart" name="requiredtxtDateStart" value="<%=FmtDateShort(CSDate)%>" class="date">
                <script type="text/javascript">
                var cal1 = new tcal({'formname':'frmParameter', 'controlname':'requiredtxtDateStart'});
                cal1.a_tpl.yearscroll = true;
                </script>
              </label>

              <label for="requiredtxtDateEnd">
                <input onBlur="validateDate(this, '<%=FmtDateShort(cEDate)%>', true);" type="text"  name="requiredtxtDateEnd" id="requiredtxtDateEnd" value="<%=FmtDateShort(CEDate)%>" class="date">
                <script type="text/javascript">
                var cal2 = new tcal({'formname':'frmParameter', 'controlname':'requiredtxtDateEnd'});
                cal2.a_tpl.yearscroll = true;
                </script>
              </label>

              <label for="optTG">
                <%
                strSQL = "SELECT DISTINCT tblTypeGroup.TypeGroupID, tblTypeGroup.TypeGroup "
                strSQL = strSQL & "FROM tblTypeGroup WHERE Active = 1 "
                strSQL = strSQL & "ORDER BY tblTypeGroup.TypeGroup "
                rsEntry.CursorLocation = 3
                rsEntry.open strSQL, cnWS
                Set rsEntry.ActiveConnection = Nothing
                %>
                <select id="optTG" name="optTG">
                  <option value="0">All Programs</option>
                  <%
                  do while NOT rsEntry.EOF
                      %>
                      <option value="<%=rsEntry("TypeGroupID")%>" <% if request.form("optTG")=CSTR(rsEntry("TypeGroupID")) then response.write " selected" end if %>><%=rsEntry("TypeGroup")%></option>
                      <%
                      rsEntry.moveNext
                  loop
                  rsEntry.close
                  %>
                </select>
              </label>

              <label for="optClassType">
                <%
                strSQL = "SELECT tblVisitTypes.TypeID, tblVisitTypes.TypeName FROM tblVisitTypes "
                strSQL = strSQL & "WHERE (tblVisitTypes.[Delete]=0) AND tblVisitTypes.Active = 1 "
                if request.form("optTG")<>"" AND request.form("optTG")<>"0" then
                    strSQL = strSQL & "AND tblVisitTypes.TypeGroup=" & request.form("optTG")
                end if
                rsEntry.CursorLocation = 3
                rsEntry.open strSQL, cnWS
                Set rsEntry.ActiveConnection = Nothing
                %>
                <select id="optClassType" name="optClassType">
                  <option value="0">All ClassTypes</option>
                  <%
                  do while NOT rsEntry.EOF
                      %>
                      <option value="<%=rsEntry("TypeID")%>" <% if request.form("optClassType")=CSTR(rsEntry("TypeID")) then response.write " selected" end if %>><%=rsEntry("TypeName")%></option>
                      <%
                      rsEntry.moveNext
                  loop
                  rsEntry.close
                  %>
                </select>
              </label>

              <label for="optSortBy">
                Sort by:
                <select name="optSortBy" id="optSortBy">
                  <option value="date" <% if request.form("optSortBy")="date" then response.write " selected" end if %>>Class Date</option>
                  <option value="name" <% if request.form("optSortBy")="name" then response.write " selected" end if %>>Client Last Name</option>
                  <option value="age" <% if request.form("optSortBy")="age" then response.write " selected" end if %>>Client Age</option>
                  <option value="class" <% if request.form("optSortBy")="class" then response.write " selected" end if %>>Class Name</option>
                </select>
              </label>

              <% taggingFilter %>
              <input type="button" name="Button" value="Generate" onClick="genReport();" />
              <% exportToExcelButton %>
              <% if Session("Pass") AND Session("Admin")<>"false" AND validAccessPriv("RPT_TAG") then %>
                <% taggingButtons("frmParameter") %>
              <% end if %>
            </form>
          </div>
<% end if %>
          <div id="report" class="mainText" style="margin: 0 auto;">
              <%
              if request.form("frmGenReport")="true" then

                  dim tagJoin, tagWhere, orderBy

                  if request.form("frmExpReport")="true" then
                      Dim stFilename : stFilename = "attachment; filename=Client Index Report.xls"
                      Response.ContentType = "application/vnd.ms-excel"
                      Response.AddHeader "Content-Disposition", stFilename
                  end if

                  showHeader = "false"

                  ' filter for sort by
                  select case request.form("optSortBy")
                      case "date"    orderBy = "ORDER BY tblTestClass.ClassDate, CLIENTS.LastName, CLIENTS.FirstName"
                      case "name"    orderBy = "ORDER BY CLIENTS.LastName, CLIENTS.FirstName, tblTestClass.ClassDate"
                      case "age"     orderBy = "ORDER BY CLIENTS.Birthdate DESC, tblTestClass.ClassDate"
                      case "class"   orderBy = "ORDER BY Q1.ClassName, tblTestClass.ClassDate"
                  end select

                  '' clear order by if tagging clients
                  if request.form("frmTagClients")="true" Then orderBy = "" end if

                  '' filter tagged clients only
                  if request.form("optFilterTagged")="on" Then
                      tagJoin = "INNER JOIN tblClientTag ON CLIENTS.ClientID = tblClientTag.ClientID"
                      if session("mvaruserID")<>"" then
                          tagWhere = "WHERE (tblClientTag.smodeID = " & session("mvaruserID") & ")" 
                      else
                          tagWhere = "WHERE (tblClientTag.smodeID = 0)" 
                      end if
                  end if

                  strSQL = join(array("SELECT CLIENTS.LastName, CLIENTS.FirstName, CLIENTS.Birthdate, Q1.TestDataNumber, Q1.TestDataAlpha, Q1.Question1, CLIENTS.ClientID, Q1.TestDate, Q1.ClassName",_
                                      "FROM CLIENTS",_
                                      tagJoin,_
                                      "INNER JOIN (SELECT tblTestQuestion.TestQuestionID, tblTestQuestion.Name AS Question1, tblTestData.ClientID, tblTestData.TestDataNumber, tblTestData.TestDataAlpha, tblTestData.ClassID, tblTestData.TestDate, tblClassDescriptions.ClassName",_
                                                  "FROM tblTestQuestion",_
                                                  "INNER JOIN tblTestData ON tblTestQuestion.TestQuestionID = tblTestData.TestQuestionID",_
                                                  "INNER JOIN tblClasses ON tblTestData.ClassID = tblClasses.ClassID",_
                                                  "INNER JOIN tblClassDescriptions ON tblClasses.DescriptionID = tblClassDescriptions.ClassDescriptionID",_
                                                  "WHERE tblTestData.TestDate <= '", cEDate, "' AND tblTestData.TestDate >= '", cSDate, "') Q1 ON CLIENTS.ClientID = Q1.ClientID",_
                                      "INNER JOIN tblTestClass ON Q1.ClassID = tblTestClass.ClassID AND Q1.TestDate = tblTestClass.ClassDate",_
                                      tagWhere,_
                                      orderBy), " ")

                  if request.form("frmTagClients")="true" then
                      ' tagging sql
                      if request.form("frmTagClientsNew")="true" then
                          clearAndTagQuery(strSQL)
                      else
                          tagQuery(strSQL)
                      end if
                  end if

                 response.write debugSQL(strSQL, "SQL")
                  'response.end
                  rsEntry.CursorLocation = 3
                  rsEntry.open strSQL, cnWS
                  Set rsEntry.ActiveConnection = Nothing

                  %>
                  <table class="sortable" id="classTestQuestions">
                    <thead>
                      <tr class="mainText">
                        <th>Date</th>
                        <th><%=firstName_hw%></th>
                        <th><%=lastName_hw%></th>
                        <th>Age</th>
                        <th>Class Name</th>
                        <th>Test Questions</th>
                        <th>Test Answers</th>
                      </tr>
                    </thead>
                    <tbody>
                      <%
                      do while NOT rsEntry.EOF
                          age = Left(DateDiff("d", rsEntry("BirthDate"), Now)/365,2)

                          if rsEntry("TestDataAlpha")<>"" then
                              answer = rsEntry("TestDataAlpha")
                          else
                              answer = rsEntry("TestDataNumber")
                          end if

                          %>
                          <tr class="mainText">
                            <td><%=rsEntry("TestDate")%></td>
                            <td><%=rsEntry("FirstName")%></td>
                            <td><%=rsEntry("LastName")%></td>
                            <td><%=age%></td>
                            <td><%=rsEntry("ClassName")%></td>
                            <td><%=rsEntry("Question1")%></td>
                            <td><%=answer%></td>
                          </tr>
                          <%
                          rsEntry.MoveNext
                      loop 'end rsEntry.EOF loop
                      %>
                    </tbody>
                  </table>
                  <%
              end if 'end of generate report if statement

              set rsEntry = nothing
              %>
<% if NOT request.form("frmExpReport")="true" then %>
              </div> <% '#report %>
            </div> <% '#container %>

<% pageEnd %>
<!-- #include file="post.asp" -->

    <%

end if
%>
