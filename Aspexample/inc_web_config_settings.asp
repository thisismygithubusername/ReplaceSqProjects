<%
On Error Resume Next 
if VarType(WEB_CONFIG_SETTINGS)<>0 then
	Execute "dim WEB_CONFIG_SETTINGS"
end if
On Error Goto 0 

' Determine if web.config was changed and re-cache if any were.
function webConfigSettingsNeedRegen()
	dim fs,f
	dim webConfigDateLastModified

	set fs=Server.CreateObject("Scripting.FileSystemObject")
	set f=fs.GetFile(server.MapPath ("/web.config"))
	webConfigDateLastModified = f.DateLastModified

	if application("webConfigDateLastModified") = webConfigDateLastModified then
		webConfigSettingsNeedRegen = false
	else
		webConfigSettingsNeedRegen = true
	end if
	'Response.Write "<div style=""border: 1px solid red; padding: 10px;"">" & webConfigDateLastModified & "</div>"
end function


' function getWebConfigSettings()
' returns a jscript object that resembles:
'	{
'		connectionStrings:
'		{
'			wsMBPS: 'Server=tcp:slo-sqldev01;Database=wsMBPS;Uid=mbo_app;Pwd=12mboapp24;',
'		},
'		appSettings:
'		{
'			UseDomainNameForInternalUrls: 'true',
'			EnvironmentName: 'LOC'
'			
'		}
'	}
function getWebConfigSettings() 
' application level caching didn't work... it wasn't reset when modifying web.config as expected
	dim fs,f

	' Request level caching.
	if isObject(WEB_CONFIG_SETTINGS) AND NOT isNull(WEB_CONFIG_SETTINGS) then
		set getWebConfigSettings = WEB_CONFIG_SETTINGS
		'set getWebConfigSettings = JSON.parse(session("getWebConfigSettings"))
	else
		' application level caching 
		if application("connectionStrings")<>"" AND NOT webConfigSettingsNeedRegen() then
			set getWebConfigSettings = JSON.parse("{}")
			getWebConfigSettings.set "connectionStrings", JSON.parse(application("connectionStrings"))
			getWebConfigSettings.set "appSettings", JSON.parse(application("appSettings"))
			set WEB_CONFIG_SETTINGS = getWebConfigSettings
		else
			application.Lock
			' catch for subsequent threads waiting on lock
			if application("connectionStrings")<>"" AND NOT webConfigSettingsNeedRegen() then
				set getWebConfigSettings = JSON.parse("{}")
				getWebConfigSettings.set "connectionStrings", JSON.parse(application("connectionStrings"))
				getWebConfigSettings.set "appSettings", JSON.parse(application("appSettings"))
				set WEB_CONFIG_SETTINGS = getWebConfigSettings
			else
				'Response.Write "<div style=""border: 1px solid red; padding: 10px;"">loading webConfigSettings</div>"
				'state variables
				dim onByDefault : onByDefault = false
				set getWebConfigSettings = JSON.parse("{}")
				dim appSettings : set appSettings = JSON.parse("{}")
				getWebConfigSettings.set "appSettings", appSettings
				dim connectionStrings : set connectionStrings = JSON.parse("{}")
				getWebConfigSettings.set "connectionStrings", connectionStrings

	
				'parse web.config and populate state variables above
				dim xmlDoc,xmlSettings,xmladd
				dim x, key, value
				set xmlDoc=server.CreateObject("MSXML2.DOMDocument.3.0")
				xmlDoc.async="false"
				'set xmlSettings=server.CreateObject("MSXML2.DOMDocument.3.0")
				'set xmladd=server.CreateObject("MSXML2.DOMDocument.3.0")

				xmlDoc.load(server.MapPath ("/web.config"))
				''''
				' need to cache all settings
				''''
				set xmlSettings = xmlDoc.GetElementsByTagName("appSettings").Item(0) 
				set xmladd = xmlSettings.GetElementsByTagName("add")
				for each x in xmladd 
					key = x.getAttribute("key")
					value = x.getAttribute("value")
					appSettings.set key, value
				next

				set xmlSettings = xmlDoc.GetElementsByTagName("connectionStrings").Item(0) 
				set xmladd = xmlSettings.GetElementsByTagName("add")
				for each x in xmladd 
					key = x.getAttribute("name")
					value = x.getAttribute("connectionString")
					connectionStrings.set key, value
				next


				' store date last modified in application cache
				set fs=Server.CreateObject("Scripting.FileSystemObject")
				set f=fs.GetFile(server.MapPath ("/web.config"))
				application("webConfigDateLastModified") = f.DateLastModified


				' store webConfig in application cache
				application("appSettings") = JSON.stringify(appSettings)
				application("connectionStrings") = JSON.stringify(connectionStrings)
				'response.Write"<div style=""border: 1px solid red; padding: 10px"">" & application("appSettings") & "</div>"
				'response.Write"<div style=""border: 1px solid red; padding: 10px"">" & application("connectionStrings") & "</div>"
				set WEB_CONFIG_SETTINGS = getWebConfigSettings
			end if
		end if
	end if
end function
%>
