<!-- #include file="inc_simple_logging.asp" -->
<%
' This file includes css and server snippets for common UI components
' CSS for these buttons are generated in adm_gs_cssp.asp
%>
<%
' good ol spacer image
function spacerImg(width, height)
	response.Write "<img src=""" & contentUrl("/asp/images/trans.gif") & """ style='width:" & width & "px; height: " & height & "px;' />"
end function

function smallRoundButtonWidthEnabled(text, onclick, width, enabled)
	smallRoundButtonWidthEnabledID "", text, onclick, width, enabled
end function

function smallRoundButtonWidthEnabledID(id, text, onclick, width, enabled)
%>
<table <% if id<>"" then response.write "id=""" & id & """ " end if %> class="simple-round-button inline-block <% if not enabled then response.write "simple-round-button-disabled" end if %>" <% if enabled then %>onclick="<%= onclick %>"<% end if %>>
	<tr>
		<td class="simple-round-button-l">&nbsp;</td>
		<td class="simple-round-button-c" <% if isNum(width) then %>style="width: <%=width%>px;"<% end if %>><%= text %></td>
		<td class="simple-round-button-r">&nbsp;</td>
	</tr>
</table>
<%
end function 

function smallRoundButton(text, onclick)
	smallRoundButtonWidthEnabled text, onclick, null, true
end function

function smallRoundButtonID(id, text, onclick)
	smallRoundButtonWidthEnabledID id, text, onclick, null, true
end function
%>

<%
function complexRoundButtonWidthEnabled(text, onclick, width, enabled)
	complexRoundButtonWidthEnabled text,onclick,width,enabled,false
end function
function complexRoundButtonWidthEnabled(text, onclick, width, enabled, dontOutput)
' create a styled button with round corners and a gradient bg
' below I illustrate how the left cron is styled
''''''''
' XXab11
' Xc2222
' d33333
' e44444
' 555555
' 555555
' 555555
' 555555
' 555555
' 555555
' 555555
' 555555
' 555555
' 555555
' 555555
' 555555
' f66666
' g77777
' Xh8888
' XXij99
''''''''''
' below I illustrate how the right cron is styled
' 11abXX
' 2222cX
' 33333d
' 44444e
' 555555
' 555555
' 555555
' 555555
' 555555
' 555555
' 555555
' 555555
' 555555
' 555555
' 555555
' 555555
' 66666f
' 77777g
' 8888hX
' 99ijXX
''''''''

dim widthString : widthString = ""
if width<>"" AND isNumeric(width) then 
   widthString = " style=""width: " + width + "px;"
end if

dim extraClasses : extraclasses = ""
if NOT enabled then
	extraClasses = " crb-disabled"
end if


dim onclickString : onclickString = ""
if enabled then
	onclickString = " onclick=""" + onclick + """"
end if	

complexRoundButtonWidthEnabled =									"<table cellspacing=""0"" class=""crb" + extraClasses + """" + onclickString + ">" 
complexRoundButtonWidthEnabled = complexRoundButtonWidthEnabled &	"<tr>"  
complexRoundButtonWidthEnabled = complexRoundButtonwidthEnabled &	"	<td class=""crb-l"">" 
complexRoundButtonWidthEnabled = complexRoundButtonwidthEnabled &	"		<div style=""height:20px;width:6px;float: left;position:relative;"">"
complexRoundButtonWidthEnabled = complexRoundButtonwidthEnabled &	"    	<div class=""crb-l1 crb-bg"">&nbsp;</div>"
complexRoundButtonWidthEnabled = complexRoundButtonwidthEnabled &	"    	<div class=""crb-l2 crb-bg"">&nbsp;</div>"
complexRoundButtonWidthEnabled = complexRoundButtonwidthEnabled &	"    	<div class=""crb-l3 crb-bg"">&nbsp;</div>"
complexRoundButtonWidthEnabled = complexRoundButtonwidthEnabled &	"    	<div class=""crb-l4 crb-bg"">&nbsp;</div>"
complexRoundButtonWidthEnabled = complexRoundButtonwidthEnabled &	"    	<div class=""crb-l5 crb-bg"">&nbsp;</div>"
complexRoundButtonWidthEnabled = complexRoundButtonwidthEnabled &	"    	<div class=""crb-l6 crb-bg"">&nbsp;</div>"
complexRoundButtonWidthEnabled = complexRoundButtonwidthEnabled &	"    	<div class=""crb-l7 crb-bg"">&nbsp;</div>"
complexRoundButtonWidthEnabled = complexRoundButtonwidthEnabled &	"    	<div class=""crb-l8 crb-bg"">&nbsp;</div>"
complexRoundButtonWidthEnabled = complexRoundButtonwidthEnabled &	"    	<div class=""crb-l9 crb-bg"">&nbsp;</div>"
complexRoundButtonWidthEnabled = complexRoundButtonwidthEnabled &	"    	<div class=""crb-la"">&nbsp;</div>"
complexRoundButtonWidthEnabled = complexRoundButtonwidthEnabled &	"    	<div class=""crb-lb crb-bg"">&nbsp;</div>"
complexRoundButtonWidthEnabled = complexRoundButtonwidthEnabled &	"    	<div class=""crb-lb-overlay"">&nbsp;</div>"
complexRoundButtonWidthEnabled = complexRoundButtonwidthEnabled &	"    	<div class=""crb-lc"">&nbsp;</div>"
complexRoundButtonWidthEnabled = complexRoundButtonwidthEnabled &	"    	<div class=""crb-ld"">&nbsp;</div>"
complexRoundButtonWidthEnabled = complexRoundButtonwidthEnabled &	"    	<div class=""crb-le crb-bg"">&nbsp;</div>"
complexRoundButtonWidthEnabled = complexRoundButtonwidthEnabled &	"    	<div class=""crb-le-overlay"">&nbsp;</div>"
complexRoundButtonWidthEnabled = complexRoundButtonwidthEnabled &	"    	<div class=""crb-lf crb-bg"">&nbsp;</div>"
complexRoundButtonWidthEnabled = complexRoundButtonwidthEnabled &	"    	<div class=""crb-lf-overlay"">&nbsp;</div>"
complexRoundButtonWidthEnabled = complexRoundButtonwidthEnabled &	"    	<div class=""crb-lg"">&nbsp;</div>"
complexRoundButtonWidthEnabled = complexRoundButtonwidthEnabled &	"    	<div class=""crb-lh"">&nbsp;</div>"
complexRoundButtonWidthEnabled = complexRoundButtonwidthEnabled &	"    	<div class=""crb-li"">&nbsp;</div>"
complexRoundButtonWidthEnabled = complexRoundButtonwidthEnabled &	"    	<div class=""crb-lj crb-bg"">&nbsp;</div>"
complexRoundButtonWidthEnabled = complexRoundButtonwidthEnabled &	"    	<div class=""crb-lj-overlay"">&nbsp;</div>"
complexRoundButtonWidthEnabled = complexRoundButtonwidthEnabled &	"	    </div>"
complexRoundButtonWidthEnabled = complexRoundButtonwidthEnabled &	"	</td>"
complexRoundButtonWidthEnabled = complexRoundButtonwidthEnabled &	"	<td class=""crb-c crb-bg""" + widthString + ">" + text + "</td>"
complexRoundButtonWidthEnabled = complexRoundButtonwidthEnabled &	"	<td class=""crb-r"">"
complexRoundButtonWidthEnabled = complexRoundButtonwidthEnabled &	"	    <div style=""height:20px;width:6px;float: left;position:relative;"">"
complexRoundButtonWidthEnabled = complexRoundButtonwidthEnabled &	"    	<div class=""crb-r1 crb-bg"">&nbsp;</div>"
complexRoundButtonWidthEnabled = complexRoundButtonwidthEnabled &	"    	<div class=""crb-r2 crb-bg"">&nbsp;</div>"
complexRoundButtonWidthEnabled = complexRoundButtonwidthEnabled &	"    	<div class=""crb-r3 crb-bg"">&nbsp;</div>"
complexRoundButtonWidthEnabled = complexRoundButtonwidthEnabled &	"    	<div class=""crb-r4 crb-bg"">&nbsp;</div>"
complexRoundButtonWidthEnabled = complexRoundButtonwidthEnabled &	"    	<div class=""crb-r5 crb-bg"">&nbsp;</div>"
complexRoundButtonWidthEnabled = complexRoundButtonwidthEnabled &	"    	<div class=""crb-r6 crb-bg"">&nbsp;</div>"
complexRoundButtonWidthEnabled = complexRoundButtonwidthEnabled &	"    	<div class=""crb-r7 crb-bg"">&nbsp;</div>"
complexRoundButtonWidthEnabled = complexRoundButtonwidthEnabled &	"    	<div class=""crb-r8 crb-bg"">&nbsp;</div>"
complexRoundButtonWidthEnabled = complexRoundButtonwidthEnabled &	"    	<div class=""crb-r9 crb-bg"">&nbsp;</div>"
complexRoundButtonWidthEnabled = complexRoundButtonwidthEnabled &	"    	<div class=""crb-ra crb-bg"">&nbsp;</div>"
complexRoundButtonWidthEnabled = complexRoundButtonwidthEnabled &	"    	<div class=""crb-ra-overlay"">&nbsp;</div>"
complexRoundButtonWidthEnabled = complexRoundButtonwidthEnabled &	"    	<div class=""crb-rb"">&nbsp;</div>"
complexRoundButtonWidthEnabled = complexRoundButtonwidthEnabled &	"    	<div class=""crb-rc"">&nbsp;</div>"
complexRoundButtonWidthEnabled = complexRoundButtonwidthEnabled &	"    	<div class=""crb-rd"">&nbsp;</div>"
complexRoundButtonWidthEnabled = complexRoundButtonwidthEnabled &	"    	<div class=""crb-re crb-bg"">&nbsp;</div>"
complexRoundButtonWidthEnabled = complexRoundButtonwidthEnabled &	"    	<div class=""crb-re-overlay"">&nbsp;</div>"
complexRoundButtonWidthEnabled = complexRoundButtonwidthEnabled &	"    	<div class=""crb-rf crb-bg"">&nbsp;</div>"
complexRoundButtonWidthEnabled = complexRoundButtonwidthEnabled &	"    	<div class=""crb-rf-overlay"">&nbsp;</div>"
complexRoundButtonWidthEnabled = complexRoundButtonwidthEnabled &	"    	<div class=""crb-rg"">&nbsp;</div>"
complexRoundButtonWidthEnabled = complexRoundButtonwidthEnabled &	"    	<div class=""crb-rh"">&nbsp;</div>"
complexRoundButtonWidthEnabled = complexRoundButtonwidthEnabled &	"    	<div class=""crb-ri crb-bg"">&nbsp;</div>"
complexRoundButtonWidthEnabled = complexRoundButtonwidthEnabled &	"    	<div class=""crb-ri-overlay"">&nbsp;</div>"
complexRoundButtonWidthEnabled = complexRoundButtonwidthEnabled &	"    	<div class=""crb-rj"">&nbsp;</div>"
complexRoundButtonWidthEnabled = complexRoundButtonwidthEnabled &	"	    </div>"
complexRoundButtonWidthEnabled = complexRoundButtonwidthEnabled &	"	</td>"
complexRoundButtonWidthEnabled = complexRoundButtonwidthEnabled &	"	</tr>"
complexRoundButtonWidthEnabled = complexRoundButtonwidthEnabled &	"</table>"

if NOT dontOutput then
	response.Write complexRoundButtonWidthEnabled
end if
%>
<%
end function

function complexRoundButton(text, onclick)
	complexRoundButtonWidthEnabled text, onclick, null, true, false
end function
%>

