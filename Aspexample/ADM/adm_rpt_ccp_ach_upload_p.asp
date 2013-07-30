<%@ CodePage=65001 %>
<%
'Dim SessionFarm
'set SessionFarm = Server.CreateObject("SessionFarm.SFSession")
response.charset="utf-8"

Set Image = Server.CreateObject("csImageFile.Manage")

'if Session("curSession") <> Session.SessionID then
if Session("StudioID") = "" then
%>
<script type="text/javascript">
	parent.resetSession();
</script>
<%
else
%>
<!-- #include file="../inc_dbconn.asp" -->
<!-- #include file="inc_accpriv.asp" -->
<%
	if not Session("Pass") OR Session("Admin")="false" OR NOT validAccessPriv("CLT_DOCS_A") then 
%>
<script type="text/javascript">
	alert("You are not authorized to view this page.\nPlease login to an account with Access Privileges granted for this feature.");
	javascript:history.go(-1);
</script>

<%
	else

		Dim Upload, fso, filename, filechars, char, rso, newFileName
		Set Upload = Server.CreateObject("csASPUpload.Process")
		Set fso = CreateObject("Scripting.FileSystemObject")

		Function UTF8FileName(strIn)   
			dim newFN
			newFN = ""
			dim sl: 	sl = 1
			dim tl: 	tl = 1
			dim key: 	key = "&#"
			dim kl: 	kl = Len(key)
			dim back_key: back_key = ";"
			dim scNdx:		scNdx = 1
			dim front_hex, back_hex, prefix_hex, post_hex
			sl = InStr(sl, strIn, key, 1)
			scNdx = InStr(1, strIn, ";", 1)
			do while sl > 0 and scNdx > 0
				if (tl=1 And sl<>1) OR tl<sl then
					newFN = newFN & Mid(strIn, tl, sl-tl)
				end if
				a = Mid(strIn, sl+2, scNdx-(sl+2))
				newFN = newFN & ChrW(a)
				sl = sl + (scNdx-sl) + 1
				tl = sl
				sl = InStr(sl, strIn, key, 1)
				if sl>0 then
					scNdx = InStr(sl, strIn, ";", 1)
				end if
			Loop
			UTF8FileName = newFN & Mid(strIn, tl)
		End Function


		newFileName = "ResultFile" & Replace(Replace(Replace(Now, " ", ""), "/", "_"), ":", "_") & ".txt"

		if Upload.FileQty > 0 then

			if Upload.Filesize(0) <= 5242880 then
				if NOT fso.FolderExists(studio_path & session("studioShort")) then
					fso.CreateFolder studio_path & session("studioShort")
				end if
	    		
				newFileName = UTF8FileName(newFileName)

				Upload.FileSave studio_path & session("studioShort") & "\" & newFilename, 0

			else %>
				<script type="text/javascript">
					alert("File too large.");
					window.location = 'adm_rpt_ccp_ach_upload.asp';
				</script>
		<%		response.flush
			end if
		end if
		
		Set Upload = nothing
		
		response.redirect "adm_rpt_ccp_ach_upload.asp?filename=" & newFileName
	end if
end if
%>
