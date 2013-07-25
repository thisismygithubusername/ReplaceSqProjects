<!-- #include file="init.asp" -->


<%

	dim startDate : startDate = Timer
	dim counter
	for counter=1 to 1000 step 1
		if EndpointToVariant(cryptoApiUrl() & "/test", "POST", "") <> true then
			Err.Raise 51,"InternalService","api call should return null"
		end if
	next
	dim endDate : endDate = Timer

	dim startDate2 : startDate2 = Timer
	for counter=1 to 1000 step 1
		
		dim fs : set fs=Server.CreateObject("Scripting.FileSystemObject")
		if fs.FileExists(server.MapPath ("/implementationSwitches.xml")) then
			dim f : set f=fs.GetFile(server.MapPath ("/implementationSwitches.xml"))
			dim implementationSwitchesXmlDateLastModified : implementationSwitchesXmlDateLastModified = f.DateLastModified
		else 
			Err.Raise 51,"InternalService","need an /implementationSwitches.xml"
		end if
	next
	dim endDate2 : endDate2 = Timer
%>


<h3>Results</h3>

<p class="">endpoint timing <span style="color:green"><%= endDate - startDate %></span></p>
<p class="">last modified check timing <span style="color:green"><%= endDate2 - startDate2 %></span></p>