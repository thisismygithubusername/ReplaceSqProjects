<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<!-- #include file="json/JSON.asp" -->
<!-- #include file="inc_dbconn.asp" -->

<html>
	<head>
		<title>JSON Examples</title>
	</head>
	<body>
	<style>
	p{font-size:18px;}</style>
		<h2> JSON Examples </h2>
		<p> The JSON library is from our friends over at Web Dev Bros. For their site and more documentation, go <a href="http://www.webdevbros.net/2007/04/26/generate-json-from-asp-datatypes/">here</a>.</p>
		<h3> Steps/Setup </h3>
<%
	set d = server.CreateObject("scripting.dictionary")	
	
	d.add "string", "This is indexed by a string"
	d.add 5, "This is indexed by a number"
	d.add "boolean", false
	d.add "array", array(10, "hi", false)
%>
<ol>
<li> Include the JSON.asp library to your ASP page.</li>
<li>Create a vb dictionary to hold your name/value pairs:<br /><code>dim d : set d = server.CreateObject("scripting.dictionary")</code></li>
<li>Add your name/values. Names can be either numbers or strings. Values can be booleans, strings, numbers, doubles, arrays, or other dictionaries.<br /><br /><code>d.add "string", "This is indexed by a string"<br />d.add 5, "This is indexed by a number"<br />d.add "boolean", false<br />d.add "array", array(10, "hi", false)</code></li><br />
<li>JSONify it! The parts this is done by:<br /><br /> <code>(new JSON_v1).toJSON("nameOfIdentifier", d, false)</code><br /><br />Note the parameters to the toJSON function: first, the name that you wish to define for this json statement. You may use empty (WITHOUT QUOTES) if you don't want an identifier. The second argument is the dictionary/array/value you wish to JSONify. The third argument is true if the value you're converting to JSON is already within another object or array, and false otherwise.</li>
</ol>
<p>The result of the above code is:</p>
<code><%= (new JSON_v1).toJSON("nameOfIdentifier", d, false) %></code>
<p> If you were to JSONify something with the last parameter set to "true", let's say...</p>
<code>(new JSON_v1).toJSON(empty, array(10, "hi"), true)</code>
<p>It would return...</p>
<code><%= (new JSON_v1).toJSON(empty, array(10, "hi"), true)%></code>
<h3>Using SQL with JSON</h3>
<p> Do your normal setup, and create a dictionary for holding the info: </p>
<code>set e = server.CreateObject("scripting.dictionary")<br /><br />

dim rsEntry<br />
set rsEntry = Server.CreateObject("ADODB.Recordset")<br /><br />

dim strSQL<br />
strSQL = "SELECT LocationName, Address FROM Location WHERE LocationID = 1"<br /><br />

rsEntry.CursorLocation = 3<br />
rsEntry.open strSQL, cnWS<br />
Set rsEntry.ActiveConnection = Nothing</code>
<p> Then, when you remove content from the query result, append .Value to retrieve the value (and not the Field object).</p>
<code>if NOT rsEntry.EOF then<br />
e.add "LocationName", rsEntry("LocationName").Value<br />
e.add "Address", rsEntry("Address").Value<br />
end if<br />
rsEntry.close<br /><br />(new JSON_v1).toJSON(empty, e, false)</code>
<%

set e = server.CreateObject("scripting.dictionary")	

dim rsEntry
set rsEntry = Server.CreateObject("ADODB.Recordset")

dim strSQL
strSQL = "SELECT LocationName, Address FROM Location WHERE LocationID = 1"

rsEntry.CursorLocation = 3
rsEntry.open strSQL, cnWS
Set rsEntry.ActiveConnection = Nothing
if NOT rsEntry.EOF then
e.add "LocationName", rsEntry("LocationName").Value
e.add "Address", rsEntry("Address").Value
%>
<p>This produces: <code><%=(new JSON_v1).toJSON(empty, e, false) %></code></p>
<% end if
rsEntry.close %>
	</body>
</html>
