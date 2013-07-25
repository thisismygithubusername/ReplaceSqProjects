<!-- #include file="json2.asp" -->
<!-- #include file="inc_post.asp" -->
<!-- #include file="inc_i18n.asp" -->

<%
''''''''''''''''''''''''''''''''''''
' setup some global html
''''''''''''''''''''''''''''''''''''
if request.QueryString("runTest")="" then
%>
<!-- #include file="frame_bottom.asp" -->
<!-- #include file="inc_jquery.asp" -->
<style type="text/css">
BODY 
{
	padding: 10px;
}
.test 
{
	border: 1px solid #999999;
	background-color: #FAFAFA;
	padding: 10px;
	margin-bottom: 25px;
}
.test h3 
{
	margin-top: 0;
}

</style>
<script src="Scripts/Plugins/jquery.libasync.js"></script>
<%
'response.write js(array("plugins/jquery.libasync"))
end if



''''''''''''''''''''''''''''''''''''
' test1 - libasync in it's simplest form
''''''''''''''''''''''''''''''''''''
if request.QueryString("runTest")="" then
%>
<div class="test">
	<h3>test1 - libasync in it's simplest form</h3>
	<p>Click the button, an async call will be made, as you should expect json data to be returned like {"success":true}, and absolutely nothing will happen</p>
	<button onclick="$.libasync.process({url: '/asp/libasync_examples.asp?runTest=test1',studioID: <%= session("studioid") %>});">Run Test 1</button>
</div>
<%
end if
if request.QueryString("runTest")="test1" then
	' do a whole bunch of work
	' return json data that libasync can work worth
	' at the very least we need to return an object with it's 
	response.ContentType = "application/json"
	response.Write("{")
	response.Write("""success"": true")
	response.Write("}")
	response.End
end if





''''''''''''''''''''''''''''''''''''
' test2 - libasync running an anonymous script on success
''''''''''''''''''''''''''''''''''''
if request.QueryString("runTest")="" then
%>
<div class="test">
	<h3>test2 - libasync running an anonymous script on success</h3>
	<p>Click the button, an async call will be made returning json data {"success":true, "script": "alert('we totally rock!');"}, and the script will fire that alert</p>
	<button onclick="$.libasync.process({url: '/asp/libasync_examples.asp?runTest=test2',studioID: <%= session("studioid") %>});">Run Test 2</button>
</div>
<%
end if
if request.QueryString("runTest")="test2" then
	' do a whole bunch of work
	' return json data that libasync can work worth
	' at the very least we need to return an object with it's 
	response.ContentType = "application/json"
	response.Write("{")
	response.Write("""success"": true,")
	response.Write("""script"": ""alert('we totally rock!');""")
	response.Write("}")
	response.End
end if




''''''''''''''''''''''''''''''''''''
' test3 - handling success by passing json to a custom callback function
''''''''''''''''''''''''''''''''''''
if request.QueryString("runTest")="" then
%>
<script type="text/javascript">
	function success(jsonData) {
		alert('callback success: ' + JSON.stringify(jsonData));
	}
</script>
<div class="test">
	<h3>test3 - handling success by passing json to a custom callback function</h3>
	<p>Click the button, an async call will be made returning json data {"success":true, "script": "alert('we totally rock!');"}, and the script will fire that alert</p>
	<button onclick="$.libasync.process({url: '/asp/libasync_examples.asp?runTest=test3',callback: success,studioID: <%= session("studioid") %>});">Run Test 3</button>
</div>
<%
end if

if request.QueryString("runTest")="test3" then
	' do a whole bunch of work
	' return json data that libasync can work worth
	' at the very least we need to return an object with it's 
	response.ContentType = "application/json"
	response.Write("{")
	response.Write("""success"": true,")
	response.Write("""json"": {""var1"":""a"",""var2"":""b""}")
	response.Write("}")
	response.End
end if










''''''''''''''''''''''''''''''''''''
' test4 - an internal error occurred
''''''''''''''''''''''''''''''''''''
if request.QueryString("runTest")="" then
%>
<div class="test">
	<h3>test4 - an internal error occurred</h3>
	<p>Click the button, an async call will be made returning json data {}, and the library will throw an error because we didn't return success true.</p>
	<button onclick="$.libasync.process({url: '/asp/libasync_examples.asp?runTest=test4',studioID: <%= session("studioid") %>});">Run Test 4</button>
</div>
<%
end if
if request.QueryString("runTest")="test4" then
	' do a whole bunch of work
	' return json data that libasync can work worth
	' at the very least we need to return an object with it's 
	response.ContentType = "application/json"
	response.Write("{")
	response.Write("}")
	response.End
end if
















''''''''''''''''''''''''''''''''''''
' test4 - session expired
''''''''''''''''''''''''''''''''''''
if request.QueryString("runTest")="" then
%>
<div class="test">
	<h3>test5 - session expired</h3>
	<p>Click the button, an async call will be made returning json data {"success":false, "sessionExpired": true}, and we will handle the session expired</p>
	<button onclick="$.libasync.process({url: '/asp/libasync_examples.asp?runTest=test5',studioID: <%= session("studioid") %>});">Run Test 5</button>
</div>
<%
end if
if request.QueryString("runTest")="test5" then
	' do a whole bunch of work
	' return json data that libasync can work worth
	' at the very least we need to return an object with it's 
	response.ContentType = "application/json"
	response.Write("{")
	response.Write("""success"": false,")
	response.Write("""sessionExpired"": true")
	response.Write("}")
	response.End
end if


'if NOT isNull(request.QueryString("test2")) then
'end if
'if NOT isNull(request.QueryString("test3")) then
'end if


''''''''''''''''''''''''''''''''''''
' test6 - calling a C Sharp endpoint  
''''''''''''''''''''''''''''''''''''
if request.QueryString("runTest")="" then
%>
<div class="test">
	<h3>test6 - calling a C Sharp endpoint</h3>
	<p>Click the button, an async call will be made to C Sharp returning json data {"success":true, "script": "alert('you just called mb.Web.Controllers.AppointmentsController.Test');"}, session security will be handled</p>
	<button onclick="$.libasync.process({url: '/Ajax/LibAsync?action=Appointment.Test',studioID: <%= session("studioid") %>});">Run Test 6</button>
</div>
<%
end if








''''''''''''''''''''''''''''''''''''
' test7 - calling an ASP endpoint  
''''''''''''''''''''''''''''''''''''
if request.QueryString("runTest")="" then
%>
<div class="test">
	<h3>test7 - calling an ASP endpoint</h3>
	<p>Due to limitation with reviving an expired session in ASP, we can't support async calls to ASP any more.  PORT YOUR CODE</p>
</div>
<%
end if



%>
