<%@ CodePage=65001 %>
<%
Dim SessionFarm
set SessionFarm = Server.CreateObject("SessionFarm.SFSession")
response.charset="utf-8"

'no longer necassary
'on error resume next
'if NOT DISABLE_COM_INTEROP then 
'	dim initializer : set initializer = Server.CreateObject("mb.Core.Tools.InitializeCOM")
'	initializer.OnError()
'end if	
'on error goto 0 
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<%

'if Session("curSession") <> Session.SessionID then
if Session("StudioID") = "" then
    Response.Write "<script type=""text/javascript"">parent.resetSession();</script>"
else
%>
		<!-- #include file="inc_dbconn.asp" -->
		<!-- #include file="inc_dbconn_wsMaster.asp" -->
		<!-- #include file="inc_i18n.asp" -->
		<!-- #include file="inc_localization.asp" -->		
<%
    function ShowErrorMessage(strMsg)
        'if true then
        if Session("Admin") ="sa" and NOT showSpecialMsg then
            Response.Write strMsg
        end if
    end function

    function showMsgTitle()
        if showSpecialMsg then
            response.Write specialMsgTitle
        else
            response.Write "Oops, this web page just hiccupped." 
	    end if
    end function

    function showMsgDetail()
        if showSpecialMsg then
            response.Write specialMsgDetail
        else
            if Session("Admin") ="false" then
                response.Write "<p>Would you email us a list of things you clicked on or typed before this error appeared? Let us know the error code below, too, so that we can email it to our software company (MINDBODY, Inc.). </p><p>Thanks!</p>" 
            elseif not session("admin") = "sa" then
                 response.Write "<p>Please email our Tech Support Team at <a href=""mailto:errors@mindbodyonline.com"">errors@mindbodyonline.com</a> a list of things you clicked on or typed before this error appeared. Email in the error code below too. It helps us fix the problem faster. </p><p>Thanks!</p>"
	        end if
	    end if
    end function

    if request.querystring("js")="true" then
        '' Handle Client Side Javascript errors coming from an ajax call
        
        dim message : message = join(array(request.querystring("errMsg") & " --- ",_
                                           "Referrer: " & request.querystring("httpRef") & " --- ",_
                                           "Querystring: " & request.querystring("queryString")))
        logError Now,_
                 "Javascript",_
                 message,_
                 request.querystring("url") ,_
                 session("studioID"),_
                 request.ServerVariables("HTTP_HOST"),_
                 getIPAddress,_
				 null
        
        response.end
    end if

  Const lngMaxFormBytes = 200

  Dim objASPError, blnErrorWritten, strServername, strServerIP, strRemoteIP
  Dim strMethod, lngPos, datNow, strQueryString, strURL
    dim rsError, curError
    Dim bakCodepage

  Dim flagSpecialError : flagSpecialError = 0
  Dim strErrorType : strErrorType = ""
  Dim strErrorMsg : strErrorMsg = ""
  Dim strErrorMsgLong : strErrorMsgLong = ""
  Dim strErrorFile : strErrorFile = ""
  Dim specialMsgErrNumber : specialMsgErrNumber = "80040E09" 'must be a string!
  Dim  showSpecialMsg : showSpecialMsg = false
  Dim  specialMsgTitle : specialMsgTitle = "You're using a read-only version of your site while we do some updates..."
  Dim  specialMsgDetail: specialMsgDetail = "As we speak, we're moving MINDBODY's data storage to a new facility. This means your site is about to get even better. Completing this scheduled maintenance will pave the way for continued world-class security, improved speeds, and fewer maintenance interruptions. Woo hoo!<br /><br />" &_
    "Between 7 pm and 11 pm Pacific Standard Time, you'll be using a read-only version of your site. This site contains your business's backed-up data, which could be up to 24 hours out of date. Because it's a read-only site, you won't be able to sign-in clients, book appointments, complete sales, etc. until our updates are complete.<br /><br />" &_
    "<strong>Are you a customer trying to book a class or appointment?</strong><br />" &_
    "If so, please contact the business directly by phone or through their website. This site cannot process bookings until we're done with our updates."
  
  '0 - No special case error
  '1 - CDATE()
  '2 - Response Buffer Exceed Limit

  If Response.Buffer Then
    Response.Clear
    Response.Status = "500 Internal Server Error"
    Response.ContentType = "text/html"
    Response.Expires = 0
  End If
  showSpecialMsg = false
  Set objASPError = Server.GetLastError
  if Hex(objASPError.Number) = specialMsgErrNumber then
    showSpecialMsg = true
  end if
%>
<html>
<head>
<title><%=Session("StudioName")%> Online</title>
    <meta http-equiv="Content-Type" content="text/html" />
    <!-- #include file="frame_bottom.asp" -->
	
	<%= js(array("mb", "main_error")) %>
    <%= css(array("bootstrap/MBCoreStyles","site","buttons","grid","error")) %>
</head>
<body>
<div id="wrapper-frame" style="position:relative;">
<div id="wrapper-minheight">
<div id="wrapper-bottompad">
<div style="display:block;height:0;clear:both;">&nbsp;</div>
<div class="wrapper">
<div id="main-content">
<div class="container no-section-margins classes-wrapper">
    <div class="error-container well alert">
        <h1 class="page-header"><%=showMsgTitle() %></h1>
        <p class="error-message">
            <%=showMsgDetail() %>
            <%if Session("Admin") ="sa" AND NOT showSpecialMsg then %>
                <strong>Technical Information</strong><br />
                    Error Type:<br />
            <% end if %>
<%
  on error resume next
	  bakCodepage = Session.Codepage
	  Session.Codepage = 1252
  on error goto 0


	
  ShowErrorMessage Server.HTMLEncode(objASPError.Category) & "<br />"

    if objASPError.ASPCode > "" Then 
        ShowErrorMessage Server.HTMLEncode(", " & objASPError.ASPCode) & "<br />"
    end if
    
	strErrorCode = Server.HTMLEncode("Error code: 0x" & Hex(objASPError.Number) )
    if NOT showSpecialMsg then response.Write strErrorCode & "<br />" end if

    strErrorType = Server.HTMLEncode(objASPError.Category)

    if objASPError.ASPDescription > "" Then 
        ShowErrorMessage Server.HTMLEncode(objASPError.ASPDescription) & "<br />"
		strErrorMsg = Server.HTMLEncode(objASPError.ASPDescription)
    elseif (objASPError.Description > "") Then 
        ShowErrorMessage Server.HTMLEncode(objASPError.Description) & "<br />"
		strErrorMsg =  Server.HTMLEncode(objASPError.Description)
    end if

	strErrorMsg = strErrorMsg & "<br/> URL: " & Request.ServerVariables("URL") 

	strErrorMsgLong = strErrorMsg & "<br/>Query String: " & request.ServerVariables("QUERY_STRING") & "<br/> Form: <br/>" 
	For Each varItem in request.form
		' Strip any form variables that contain cc or password
		if InStr(1, varitem, "cc", 1) > 0 OR InStr(1, varitem, "password", 1) > 0 OR InStr(1, varitem, "swipe", 1) > 0then
			strErrorMsgLong = strErrorMsgLong & Server.HTMLEncode(varItem) & ": " & Server.HTMLEncode("---STRIPPED---") & "<br/>"
		else
			strErrorMsgLong = strErrorMsgLong & Server.HTMLEncode(varItem) & ": " & Server.HTMLEncode(request.form(varItem)) & "<br/>"
		end if
    Next

	strErrorMsgLong = strErrorMsgLong & "<br/> Session: " & session.SessionID & "<br/>" 
	For Each varItem in session.contents
		strErrorMsgLong = strErrorMsgLong & Server.HTMLEncode(varItem) & ": " & Server.HTMLEncode(session.contents(varItem)) & "<br/>"
    Next

	strErrorMsgLong = strErrorMsgLong & "<br/> Server Variables: <br/>" 
	For Each varItem in request.ServerVariables
		strErrorMsgLong = strErrorMsgLong & Server.HTMLEncode(varItem) & ": " & Server.HTMLEncode(request.ServerVariables(varItem)) & "<br/>"
	Next



	
    blnErrorWritten = False

    ' Only show the Source if it is available and the request is from the same machine as IIS
    If objASPError.Source > "" Then
        strServername = LCase(Request.ServerVariables("SERVER_NAME"))
        strServerIP = Request.ServerVariables("LOCAL_ADDR")
        strRemoteIP =  getIPAddress
        If (strServername = "localhost" Or strServerIP = strRemoteIP) And objASPError.File <> "?" Then
            If objASPError.Line > 0 Then ShowErrorMessage ", line " & objASPError.Line end if
            If objASPError.Column > 0 Then ShowErrorMessage ", column " & objASPError.Column end if
            ShowErrorMessage "<br />"
            ShowErrorMessage "<span style=""color:#000;font: 8pt/11pt courier new;""><b>"
            ShowErrorMessage Server.HTMLEncode(objASPError.Source) & "<br />"
            If objASPError.Column > 0 Then ShowErrorMessage String((objASPError.Column - 1), "-") & "^<br />" end if
            ShowErrorMessage "</b></span>"
            blnErrorWritten = True
        End If
    End If

    If Not blnErrorWritten And objASPError.File <> "?" Then
        ShowErrorMessage "<b>" & Server.HTMLEncode(objASPError.File)
        strErrorFile = Server.HTMLEncode(objASPError.File)
        If objASPError.Line > 0 Then 
		    ShowErrorMessage Server.HTMLEncode(", line " & objASPError.Line)
		    strErrorFile = strErrorFile & " " & Server.HTMLEncode(", line " & objASPError.Line)
	    end if
	
        If objASPError.Column > 0 Then 
		    ShowErrorMessage ", column " & objASPError.Column
	    end if

        ShowErrorMessage "</b><br />"
    End If

    curError = logError(now,_
                           strErrorType,_ 
                           strErrorCode & " " & strErrorMsg,_ 
                           strErrorFile,_ 
                           session("StudioID"),_
                           request.ServerVariables("HTTP_HOST"),_
                           getIPAddress,_
						   strErrorMsgLong)

    if session("strSQL")<>"" then
        logErrorSQL curError, replace(session("strSQL"), "'", "''")
    end if

    '' Is the error special
    if InStr(strErrorMsg,"CDate") > 0 then
        flagSpecialError = 1
    elseif InStr(strErrorMsg,"Response Buffer") > 0 then
        flagSpecialError = 2
    elseif InStr(strErrorMsg,"Error during transfer:") > 0 then
        flagSpecialError = 3
    end if

    '' Alert if special
    Response.Write "<br /><span style=""color:990000;""><strong>"
    if flagSpecialError=1 then
        response.write "Please check that a valid date was entered."
    elseif flagSpecialError=2 then
        response.write "The processing of this page produced too much data to display.<br /><br />Please try limiting the date range, filtering by more criteria or selecting a summary view to reduce the amount of results."
    elseif flagSpecialError=3 then
        response.Write "The credit card processor is temporarily unavailable.  If you have tried charging the card repeatedly without success, please run the sale using the account payment method, and then save the " & session("ClientHW") & "'s credit card information to his or her Profile Screen. Once your credit card processor is back online, you can pay off the negative balance by selling credit to the client's saved credit card. <br /><br /> To learn more about when the processor will be back online <a href=""http://www.mindbodyonline.com/en/systeminformation"" target=""new"">click here</a>."
    end If
    Response.Write "</strong></span>"
%>
   </p>
        <a class="secondaryBtn arrow warningBtn" onclick="history.go(-1)"><span class="left"></span> Go Back</a>
    </div><!-- .error-container -->
</div><!-- .container -->    
</div>
</div> 
</div> 
</div> 
</div>               
</body>
</html>
<%

	set rsClass = nothing

	cnWS.close
	set cnWS = nothing
	cnMB.close
	set cnMB = nothing

	end if

function logError(errorDate, errorType, errorMsg, errorFile, studioID, host, ipAddr, errorMsgLong)
	Dim categoryID, errID
    strErrSQL = "INSERT INTO tblErrorLog (ErrorDateTime, ErrorType, ErrorMsg, ErrorFile, StudioID, Host, IPAddr, webServer, ErrorMsgLong, CategoryID ) VALUES ("
    strErrSQL = strErrSQL & DateSep & errorDate & DateSep
    strErrSQL = strErrSQL & ", N'" & sqlInjectStr(errorType) & "'"
    strErrSQL = strErrSQL & ", N'" & sqlInjectStr(left(errorMsg, 500)) & "'"
    strErrSQL = strErrSQL & ", N'" & sqlInjectStr(errorFile) & "'"
  
    strErrSQL = strErrSQL & ", " & studioID
    strErrSQL = strErrSQL & ", N'" & sqlInjectStr(host) & "'"
    strErrSQL = strErrSQL & ", N'" & sqlInjectStr(ipAddr) & "'"
  
    strErrSQL = strErrSQL & ", N'" & getWebServerID() & "'" '' Web server

	if isNull(errorMsgLong) then
		strErrSQL = strErrSQL & ", null "
	else
		strErrSQL = strErrSQL & ", N'" & sqlInjectStr(errorMsgLong) & "'"
	end if
	if inStr(ErrorFile,"inc_ccp.asp") > 0 then
		categoryID = "2"
	else
		categoryID = "0"
	end if
	strErrSQL = strErrSQL & "," & categoryID
    strErrSQL = strErrSQL & ")"

	dim cnEL : set cnEL = connectToErrorLog()
	dim rs

    set rs = cnEL.execute(strErrSQL & ";SELECT @@identity").nextrecordset
	logError = rs(0)

    cnEL.close
	set cnEL = nothing
end function

sub logErrorSQL(sError, sSQL)
    strErrSQL = " INSERT INTO tblErrorSQL (ErrorID, ErrorSQL) VALUES (" & sError & ", N'" & sSQL & "')"
    
	dim cnEL : set cnEL = connectToErrorLog()

	cnEL.execute strErrSQL

	cnEL.close
	set cnEL = nothing
end sub
%>
