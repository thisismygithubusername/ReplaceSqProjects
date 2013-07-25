<!-- #include file="base64.asp" -->
<%
function floor(x)
	dim t
	t = Round(x)
	if t > x then
		t = t - 1
	end if
	floor = t
end function

'' color functions
function HexToR(h)
	err.number = 0
	on Error resume Next
	if len(h)=4 or len(h)=3 then 'Hex Shorthand, eg. #fff
		HexToR = CLng("&H" & MID(cutHex(h), 1, 1) & MID(cutHex(h), 1, 1))
	else 'Normal hex color, eg. #fafafa
		HexToR = CLng("&H" & MID(cutHex(h), 1, 2))
	end if
	on Error GoTo 0
	if err.number <> 0 then
		HexToR = 0
		err.number = 0
	end if
end function

function HexToG(h)
	err.number = 0
	on Error resume Next
	if len(h)=4 or len(h)=3 then 'Hex Shorthand, eg. #fff
		HexToG = CLng("&H" & MID(cutHex(h), 2, 1) & MID(cutHex(h), 2, 1))
	else 'Normal hex color, eg. #fafafa
		HexToG = CLng("&H" & MID(cutHex(h), 3, 2))
	end if
	on Error GoTo 0
	if err.number <> 0 then
		HexToG = 0
		err.number = 0
	end if
end function

function HexToB(h)
	err.number = 0
	on Error resume Next
	if len(h)=4 or len(h)=3 then 'Hex Shorthand, eg. #fff
		HexToB = CLng("&H" & MID(cutHex(h), 3, 1) & MID(cutHex(h), 3, 1))
	else 'Normal hex color, eg. #fafafa
		HexToB = CLng("&H" & MID(cutHex(h), 5, 2))
	end if
	on Error GoTo 0
	if err.number <> 0 then
		HexToB = 0
		err.number = 0
	end if
end function

function cutHex(h)
	cutHex = Replace(h, "#", "")
end function

function HexToCSS(h, fade)
	HexToCSS = "rgb(" & CLng(fade * HexToR(h)) &  ", " & CLng(fade * HexToG(h)) & ", " & CLng(fade * HexToB(h)) & ")"
end function
function HexToRGBA(h, fade, alpha)
	HexToRGBA = "rgba(" & CLng(fade * HexToR(h)) &  ", " & CLng(fade * HexToG(h)) & ", " & CLng(fade * HexToB(h)) & "," & alpha & ")"
end function


' morphs between 2 hex codes given parameter 0 <= t <= 1
function HexMorph(h1, h2, t)
	dim rI, rF, gI, gF, bI, bF
	dim rT, gT, bT, strR, strG, strB
	rI = HexToR(h1)
	gI = HexToG(h1)
	bI = HexToB(h1)
	rF = HexToR(h2)
	gF = HexToG(h2)
	bF = HexToB(h2)
	
	rT = CLng(rI + t*(rF-rI))
	gT = CLng(gI + t*(gF-gI))
	bT = CLng(bI + t*(bF-bI))
	
	strR = cstr(Hex(rT))
	strG = cstr(Hex(gT))
	strB = cstr(Hex(bT))
	
	if Len(Hex(rT)) = 1 then strR = "0" & strR
	if Len(Hex(gT)) = 1 then strG = "0" & strG
	if Len(Hex(bT)) = 1 then strB = "0" & strB
	
	HexMorph = "#" & strR & strG & strB
end function

function HexMorphColors(colors, t)
	dim numColors, ndx, nT
	dim rI, rF, gI, gF, bI, bF
	dim rT, gT, bT, strR, strG, strB
	
	numColors = ubound(colors)
	ndx = floor(t * (numColors-1))
	if ndx=numColors-1 then
		ndx = ndx-1
	end if
	
	rI = HexToR(colors(ndx))
	gI = HexToG(colors(ndx))
	bI = HexToB(colors(ndx))
	rF = HexToR(colors(ndx+1))
	gF = HexToG(colors(ndx+1))
	bF = HexToB(colors(ndx+1))
	
	nT = (t - (ndx/(numColors-1)))/(((ndx+1)/(numColors-1)) - (ndx/(numColors-1))) ' local t
	rT = CLng(rI + nT*(rF-rI))
	gT = CLng(gI + nT*(gF-gI))
	bT = CLng(bI + nT*(bF-bI))
	
	strR = cstr(Hex(rT))
	strG = cstr(Hex(gT))
	strB = cstr(Hex(bT))
	
	if Len(Hex(rT)) = 1 then strR = "0" & strR
	if Len(Hex(gT)) = 1 then strG = "0" & strG
	if Len(Hex(bT)) = 1 then strB = "0" & strB

	HexMorphColors = "#" & strR & strG & strB
end function

function TwoDigitHex(color)
	if color > 255 then
		color = 255
	end if
	if color < 0 then
		color = 0
	end if
	TwoDigitHex = Hex(color)
	if len(TwoDigitHex) = 1 then
		TwoDigitHex = "0" & TwoDigitHex
	end if
end function

' Use this function to sanitize a color from the database - if all cases fail it returns the input
function HexCode(color)
	dim regex,rgbRegex

	Set rgbRegex = New RegExp
	rgbRegex.Pattern = "^\s*rgb\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)\s*\)\s*$"
	rgbRegex.IgnoreCase = True
	Set regex = New RegExp
	regex.Pattern = "^(#?)([a-fA-F0-9]{6}|[a-fA-F0-9]{3})$"
	regex.IgnoreCase = True
	if rgbRegex.Test(color) then ' valid rgb(r,g,b)
		HexCode = "#" & TwoDigitHex(rgbRegex.replace(color,"$1")) & TwoDigitHex(rgbRegex.replace(color,"$2")) & TwoDigitHex(rgbRegex.replace(color,"$3"))
	elseif regex.Test(color) then ' valid hex
		HexCode = regex.Replace(color, "#$2")
	else
		select case lcase(color)
			case "steelblue" HexCode="#4682B4" 
			case "royalblue" HexCode="#041690" 
			case "cornflowerblue" HexCode="#6495ED" 
			case "lightsteelblue" HexCode="#B0C4DE" 
			case "mediumslateblue" HexCode="#7B68EE" 
			case "slateblue" HexCode="#6A5ACD" 
			case "darkslateblue" HexCode="#483D8B" 
			case "midnightblue" HexCode="#191970" 
			case "navy" HexCode="#000080" 
			case "darkblue" HexCode="#00008B" 
			case "mediumblue" HexCode="#0000CD" 
			case "blue" HexCode="#0000FF" 
			case "dodgerblue" HexCode="#1E90FF" 
			case "deepskyblue" HexCode="#00BFFF" 
			case "lightskyblue" HexCode="#87CEFA" 
			case "skyblue" HexCode="#87CEEB" 
			case "lightblue" HexCode="#ADD8E6" 
			case "powderblue" HexCode="#B0E0E6" 
			case "azure" HexCode="#F0FFFF" 
			case "lightcyan" HexCode="#E0FFFF" 
			case "paleturquoise" HexCode="#AFEEEE" 
			case "mediumturquoise" HexCode="#48D1CC" 
			case "lightseagreen" HexCode="#20B2AA" 
			case "darkcyan" HexCode="#008B8B" 
			case "teal" HexCode="#008080" 
			case "cadetblue" HexCode="#5F9EA0" 
			case "darkturquoise" HexCode="#00CED1" 
			case "aqua" HexCode="#00FFFF" 
			case "cyan" HexCode="#00FFFF" 
			case "turquoise" HexCode="#40E0D0" 
			case "aquamarine" HexCode="#7FFFD4" 
			case "mediumaquamarine" HexCode="#66CDAA" 
			case "darkseagreen" HexCode="#8FBC8F" 
			case "mediumseagreen" HexCode="#3CB371" 
			case "seagreen" HexCode="#2E8B57" 
			case "darkgreen" HexCode="#006400" 
			case "green" HexCode="#008000" 
			case "forestgreen" HexCode="#228B22" 
			case "limegreen" HexCode="#32CD32" 
			case "lime" HexCode="#00FF00" 
			case "chartreuse" HexCode="#7FFF00" 
			case "lawngreen" HexCode="#7CFC00" 
			case "greenyellow" HexCode="#ADFF2F" 
			case "yellowgreen" HexCode="#9ACD32" 
			case "palegreen" HexCode="#98FB98" 
			case "lightgreen" HexCode="#90EE90" 
			case "springgreen" HexCode="#00FF7F" 
			case "mediumspringgreen" HexCode="#00FA9A" 
			case "darkolivegreen" HexCode="#556B2F" 
			case "olivedrab" HexCode="#6B8E23" 
			case "olive" HexCode="#808000" 
			case "darkkhaki" HexCode="#BDB76B" 
			case "darkgoldenrod" HexCode="#B8860B" 
			case "goldenrod" HexCode="#DAA520" 
			case "gold" HexCode="#FFD700" 
			case "yellow" HexCode="#FFFF00" 
			case "khaki" HexCode="#F0E68C" 
			case "palegoldenrod" HexCode="#EEE8AA" 
			case "blanchedalmond" HexCode="#FFEBCD" 
			case "moccasin" HexCode="#FFE4B5" 
			case "wheat" HexCode="#F5DEB3" 
			case "navajowhite" HexCode="#FFDEAD" 
			case "burlywood" HexCode="#DEB887" 
			case "tan" HexCode="#D2B48C" 
			case "rosybrown" HexCode="#BC8F8F" 
			case "sienna" HexCode="#A0522D" 
			case "saddlebrown" HexCode="#8B4513" 
			case "chocolate" HexCode="#D2691E" 
			case "peru" HexCode="#CD853F" 
			case "sandybrown" HexCode="#F4A460" 
			case "darkred" HexCode="#8B0000" 
			case "maroon" HexCode="#800000" 
			case "brown" HexCode="#A52A2A" 
			case "firebrick" HexCode="#B22222" 
			case "indianred" HexCode="#CD5C5C" 
			case "lightcoral" HexCode="#F08080" 
			case "salmon" HexCode="#FA8072" 
			case "darksalmon" HexCode="#E9967A" 
			case "lightsalmon" HexCode="#FFA07A" 
			case "coral" HexCode="#FF7F50" 
			case "tomato" HexCode="#FF6347" 
			case "darkorange" HexCode="#FF8C00" 
			case "orange" HexCode="#FFA500" 
			case "orangered" HexCode="#FF4500" 
			case "crimson" HexCode="#DC143C" 
			case "red" HexCode="#FF0000" 
			case "deeppink" HexCode="#FF1493" 
			case "fuchsia" HexCode="#FF00FF" 
			case "magenta" HexCode="#FF00FF" 
			case "hotpink" HexCode="#FF69B4" 
			case "lightpink" HexCode="#FFB6C1" 
			case "pink" HexCode="#FFC0CB" 
			case "palevioletred" HexCode="#DB7093" 
			case "mediumvioletred" HexCode="#C71585" 
			case "purple" HexCode="#800080" 
			case "darkmagenta" HexCode="#8B008B" 
			case "mediumpurple" HexCode="#9370DB" 
			case "blueviolet" HexCode="#8A2BE2" 
			case "indigo" HexCode="#4B0082" 
			case "darkviolet" HexCode="#9400D3" 
			case "darkorchid" HexCode="#9932CC" 
			case "mediumorchid" HexCode="#BA55D3" 
			case "orchid" HexCode="#DA70D6" 
			case "violet" HexCode="#EE82EE" 
			case "plum" HexCode="#DDA0DD" 
			case "thistle" HexCode="#D8BFD8" 
			case "lavender" HexCode="#E6E6FA" 
			case "ghostwhite" HexCode="#F8F8FF" 
			case "aliceblue" HexCode="#F0F8FF" 
			case "mintcream" HexCode="#F5FFFA" 
			case "honeydew" HexCode="#F0FFF0" 
			case "lightgoldenrodyellow" HexCode="#FAFAD2" 
			case "lemonchiffon" HexCode="#FFFACD" 
			case "cornsilk" HexCode="#FFF8DC" 
			case "lightyellow" HexCode="#FFFFE0" 
			case "ivory" HexCode="#FFFFF0" 
			case "floralwhite" HexCode="#FFFAF0" 
			case "linen" HexCode="#FAF0E6" 
			case "oldlace" HexCode="#FDF5E6" 
			case "antiquewhite" HexCode="#FAEBD7" 
			case "bisque" HexCode="#FFE4C4"  
			case "peachpuff" HexCode="#FFDAB9" 
			case "papayawhip" HexCode="#FFEFD5" 
			case "beige" HexCode="#F5F5DC" 
			case "seashell" HexCode="#FFF5EE" 
			case "lavenderblush" HexCode="#FFF0F5" 
			case "mistyrose" HexCode="#FFE4E1" 
			case "snow" HexCode="#FFFAFA" 
			case "white" HexCode="#FFFFFF" 
			case "whitesmoke" HexCode="#F5F5F5" 
			case "gainsboro" HexCode="#DCDCDC" 
			case "lightgrey" HexCode="#D3D3D3" 
			case "silver" HexCode="#C0C0C0" 
			case "darkgray" HexCode="#A9A9A9" 
			case "gray" HexCode="#808080" 
			case "lightslategray" HexCode="#778899" 
			case "slategray" HexCode="#708090" 
			case "dimgray" HexCode="#696969" 
			case "darkslategray" HexCode="#2F4F4F" 
			case "black" HexCode="#000000" 
			case else HexCode="#000000"
		end select
	end if
end function 

''''''''''''''''''''''''''''
' CssGradient generates all the browser specific gradient css styles for simple vertical
' and horizontal styles
'
' Params:
'	id: a unique identifier for the internal svg image used by IE9
'	gradientType: either "vertical" or "horizonatal"
'	hex1,fade1,alpha1: the starting color
'	hex2,fade2,alpha2: the ending color
'	important: whither to add !important tag
''''''''''''''''''''''''''''
function CssGradient(id,gradientType,hex1,fade1,alpha1,hex2,fade2,alpha2,important)
	if gradientType = "vertical" then
		if isIE7 OR isIE8 then
			CssGradient=CssGradient&"filter: progid:DXImageTransform.Microsoft.gradient(gradientType=0,startColorstr=#" & TwoDigitHex(CInt(alpha2 * 255)) & cutHex(HexCode(HexToCSS(hex2, fade2))) & ", endColorstr=#" & TwoDigitHex(CInt(alpha1 * 255)) & cutHex(HexCode(HexToCSS(hex1, fade1))) & ")"
			if important then CssGradient=CssGradient & " !important" end if 
			CssGradient=CssGradient&";"&vbCrLf
		else 
			CssGradient=CssGradient&"background-image: url(data:image/svg+xml;base64," &_
				Base64Encode("" &_
					"<svg xmlns=""http://www.w3.org/2000/svg"" preserveAspectRatio=""none"" width=""100%"" height=""100%"">" &_
					"<linearGradient id=""" & id & """ x1=""0%"" y1=""0%"" x2=""0%"" y2=""100%"" gradientUnits=""userSpaceOnUse"">" &_
					"<stop offset=""0%"" stop-color=""" & HexCode(HexToCSS(hex2, fade2)) & """ stop-opacity=""" & alpha2 & """/>" &_
					"<stop offset=""100%"" stop-color=""" & HexCode(HexToCSS(hex1, fade1)) & """ stop-opacity=""" & alpha1 & """/>" &_
					"</linearGradient>" &_
					"<rect x=""0"" y=""0"" width=""100%"" height=""100%"" fill=""url(#" & id &")"" />" &_
					"</svg>" &_
				"",true) &_
				")"
			if important then CssGradient=CssGradient & " !important" end if 
			CssGradient=CssGradient&";"&vbCrLf
		end if


		' Chrome has good support for SVG
		CssGradient=CssGradient&"background-image: -webkit-gradient(linear, left bottom, left top, from(" & HexToRGBA(hex1,fade1,alpha1) & "), to(" & HexToRGBA(hex2,fade2,alpha2) & "))"
		if important then CssGradient=CssGradient & " !important" end if 
		CssGradient=CssGradient&";"&vbCrLf

		' Firefox support for SVG is less than perfect
		CssGradient=CssGradient&"background-image: -moz-linear-gradient(center bottom, " & HexToRGBA(hex1,fade1,alpha1) & " 0%, " & HexToRGBA(hex2,fade2,alpha2) & " 100%)"
		if important then CssGradient=CssGradient & " !important" end if 
		CssGradient=CssGradient&";"&vbCrLf


	else
		if isIE7 OR isIE8 then
			CssGradient=CssGradient&"filter: progid:DXImageTransform.Microsoft.gradient(gradientType=1,startColorstr=#" & TwoDigitHex(CInt(alpha2 * 255)) & cutHex(HexCode(HexToCSS(hex2, fade2))) & ", endColorstr=#" & TwoDigitHex(CInt(alpha1 * 255)) & cutHex(HexCode(HexToCSS(hex1, fade1))) & ")"
			if important then CssGradient=CssGradient & " !important" end if 
			CssGradient=CssGradient&";"&vbCrLf
		else
			CssGradient=CssGradient&"background-image: url(data:image/svg+xml;base64," &_
				Base64Encode("" &_
					"<svg xmlns=""http://www.w3.org/2000/svg"" preserveAspectRatio=""none"" width=""100%"" height=""100%"">" &_
					"<linearGradient id=""" & id & """ x1=""0%"" y1=""0%"" x2=""100%"" y2=""0%"" gradientUnits=""userSpaceOnUse"">" &_
					"<stop offset=""0%"" stop-color=""" & HexCode(HexToCSS(hex2, fade2)) & """ stop-opacity=""" & alpha2 & """/>" &_
					"<stop offset=""100%"" stop-color=""" & HexCode(HexToCSS(hex1, fade1)) & """ stop-opacity=""" & alpha1 & """/>" &_
					"</linearGradient>" &_
					"<rect x=""0"" y=""0"" width=""100%"" height=""100%"" fill=""url(#" & id &")"" />" &_
					"</svg>" &_
				"",true) &_
				")"
			if important then CssGradient=CssGradient & " !important" end if 
			CssGradient=CssGradient&";"&vbCrLf
		end if

		' Chrome has good support for SVG
		CssGradient=CssGradient&"background-image: -webkit-gradient(linear, right top, left top, from(" & HexToRGBA(hex1,fade1,alpha1) & "), to(" & HexToRGBA(hex2,fade2,alpha2) & "))"
		if important then CssGradient=CssGradient & " !important" end if 
		CssGradient=CssGradient&";"&vbCrLf

		' Firefox support for SVG is less than perfect
		CssGradient=CssGradient&"background-image: -moz-linear-gradient(right bottom, " & HexToRGBA(hex1,fade1,alpha1) & " 0%, " & HexToRGBA(hex2,fade2,alpha2) & " 100%)"
		if important then CssGradient=CssGradient & " !important" end if 
		CssGradient=CssGradient&";"&vbCrLf


	end if
end function

%>
