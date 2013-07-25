<%
    ' requires JSON.asp
    ' requires inc_init_functions.asp for TinyMCEInclude

    dim inc_tinyMCEMode                 : inc_tinyMCEMode               = "textareas"
    dim inc_tinyMCESpecificSelector     : inc_tinyMCESpecificSelector   = ""
    dim inc_tinyMCESpecificDeSelector   : inc_tinyMCESpecificDeSelector = ""
    dim inc_tinyMCEPlugins              : inc_tinyMCEPlugins            = "paste,autolink,hoverclicklinks" 'advlink
    dim inc_tinyMCEButtonsRow1          : inc_tinyMCEButtonsRow1        = "bold,italic,underline,strikethrough,|,justifyleft,justifycenter,justifyright,justifyfull,|,bullist,numlist,separator,redo,undo,|,fontsizeselect,forecolor,|,code" ',|,link,unlink"
    dim inc_tinyMCEButtonsRow2          : inc_tinyMCEButtonsRow2        = ""
    dim inc_tinyMCEButtonsRow3          : inc_tinyMCEButtonsRow3        = ""
    dim inc_tinyMCEButtonsRow4          : inc_tinyMCEButtonsRow4        = ""
    dim inc_tinyMCEHeight               : inc_tinyMCEHeight             = 250
    dim inc_tinyMCEWidth                : inc_tinyMCEWidth              = 300
    dim inc_tinyMCEFontSizes            : inc_tinyMCEFontSizes          = "10px,12px,13px,14px,16px,18px,20px"

    function TinyMCEInclude()
        if isMobileBrowser() then
            exit function
        end if
        %>
        <script type="text/javascript" src="/scriptplugins/tiny_mce/tiny_mce.js"></script>
        <script type="text/javascript">
            var tinyMCESetupFunction = function (ed) {
                ed.onInit.add(function(ed) {
				  ed.pasteAsPlainText = true;
				});
            };
            tinyMCE.init({
                mode: "<%=inc_tinyMCEMode %>",
                
                <% if inc_tinyMCEMode = "specific_textareas" then
                    if inc_tinyMCESpecificSelector <> "" then %>
                        editor_selector: "<%=inc_tinyMCESpecificSelector %>",
                    <%else %>
                        editor_deselector: "<%=inc_tinyMCESpecificDeSelector %>",
                    <%end if %>
                <% elseif inc_tinyMCEMode = "exact" then %>
                    elements: "<%=inc_tinyMCESpecificSelector %>",
                <% end if %>

                // skins
                skin: "o2k7",
                skin_variant: "silver",

                // dimensions
                height: "<%=inc_tinyMCEHeight %>",
                width: "<%=inc_tinyMCEWidth %>",

                plugins: "<%=inc_tinyMCEPlugins %>",

				// Paste options
				paste_auto_cleanup_on_paste : true,
				paste_remove_styles: true,
				paste_remove_styles_if_webkit: true,
				paste_strip_class_attributes: true,
				paste_text_sticky: true,

                // General options
                theme: "advanced",

                theme_advanced_buttons1: "<%=inc_tinyMCEButtonsRow1%>",
                theme_advanced_buttons2: "<%=inc_tinyMCEButtonsRow2%>",
                theme_advanced_buttons3: "<%=inc_tinyMCEButtonsRow3%>",
                theme_advanced_buttons4: "<%=inc_tinyMCEButtonsRow4%>",
                theme_advanced_toolbar_location: "top",
                theme_advanced_toolbar_align: "left",
                invalid_elements: "script,style",

                //force_br_newlines: true,
                force_p_newlines: false,
                forced_root_block: 'div',  // Needed for 3.x <%' if this is being changed TinyMCEPurifyForDB must be updated  %>
                theme_advanced_path: false,
                gecko_spellcheck : true,
                // custom styling
                content_css: "/styles/tinymce/custom_content.css?" + (new Date()).getTime(),
                theme_advanced_font_sizes: "<%=inc_tinyMCEFontSizes %>",
                font_size_style_values: "<%=inc_tinyMCEFontSizes %>",

                setup: tinyMCESetupFunction
            });
		</script>
        <%
    end function

    'This should only be used for text that will be displayed inside a tinymce editor
	function TinyMCEPurifyForDisplay(text)
        text = trim(text & "")

        if text = "" then
            TinyMCEPurifyForDisplay = ""
            exit function
        end if

		dim jsonParams, sanitizedResult
		set jsonParams = JSON.parse("{}")
		jsonParams.set "html", text
		jsonParams.set "removeForbiddenTags", "true"
		set sanitizedResult = JSON.parse(CallMethodWithJSON("mb.Core.Tools.HtmlSanitizer.Sanitizer",jsonParams))
		
		' Return the purified text
		TinyMCEPurifyForDisplay = xssStr(sanitizedResult.Html)
	end function

	function TinyMCEPurifyForDB(text)
        dim t : t = trim(text & "")

        if t = "" then
            TinyMCEPurifyForDB = ""
            exit function
        end if

        ' tinymce sometimes passes the root block with an &nbsp; as the content even when it appears as "empty"
        if t = "<div>&nbsp;</div>" then
            TinyMCEPurifyForDB = ""
            exit function
        end if

		dim jsonParams, sanitizedResult
		set jsonParams = JSON.parse("{}")
		jsonParams.set "html", t
		jsonParams.set "removeForbiddenTags", "true"
		set sanitizedResult = JSON.parse(CallMethodWithJSON("mb.Core.Tools.HtmlSanitizer.Sanitizer",jsonParams))
		
		' Return the purified val
		TinyMCEPurifyForDB = sanitizedResult.Val
	end function

    ' This should only be used for when displaying stored html as html somewhere
    function HtmlPurifyForDisplay(text)
        dim t : t = trim(text & "")

        if t = "" then
            HtmlPurifyForDisplay = ""
            exit function
        end if

        dim jsonParams, sanitizedResult
        set jsonParams = JSON.parse("{}")
        jsonParams.set "html", t
        jsonParams.set "removeForbiddenTags", "true"
        set sanitizedResult = JSON.parse(CallMethodWithJSON("mb.Core.Tools.HtmlSanitizer.Sanitizer",jsonParams))
		
        ' Return the purified text
        HtmlPurifyForDisplay = sanitizedResult.Html
    end function

    'Purifies the HTML then returns only the text nodes from within
    function PurifyAndStripHtml(text)
        dim t : t = trim(text & "")

        if t = "" then
            PurifyAndStripHtml = ""
            exit function
        end if

        dim jsonParams, sanitizedResult
        set jsonParams = JSON.parse("{}")
        jsonParams.set "html", t
        jsonParams.set "removeForbiddenTags", "true"
        set sanitizedResult = JSON.parse(CallMethodWithJSON("mb.Core.Tools.HtmlSanitizer.Sanitizer",jsonParams))

        PurifyAndStripHtml = Replace(sanitizedResult.Text, "&nbsp;", " ")
    end function
		%>

<style type="text/css">
	.userHTML ul {list-style: disc outside none !important;}
	.userHTML ol {list-style: decimal  !important;}
	.userHTML li {margin-left:30px;}
	.userHTML em {font-style: italic;}
	.userHTML blockquote
	{
		display: block;
		margin: 1em 40px;
	}
	.userHTML big 
	{
		font-size: larger;
	}
	.userHTML small
	{
		font-size: smaller;
	}
	.userHTML kbd, .userHTML code
	{
		font-family: monospace;
	}
	.userHTML pre
	{
		margin: 1em 0px;
		display: block;
		white-space: pre;
	}
	.userHTML sup
	{
		vertical-align: super;
		font-size: smaller;
	}
	.userHTML sub
	{
		vertical-align: sub;
		font-size: smaller;
	}
</style>

<script type="text/javascript">
	if (typeof ($) != "undefined") {

		$(function () {
			// Puts focus inside a tinyMCE box when it's corresponding <label> tag is clicked
			$(".for-tinymce").live("click", function (e) {
				if (typeof (tinyMCE) != "undefined") {
					e.preventDefault();
					tinyMCE.get($(this).prop('for')).focus();
				}
			});

			var $userTables = $(".userHTML table");
			
			if ($userTables.attr("border") != "") {
				$userTables.find("td, th").css("border", "1px solid black");
			}
		});
	}
</script>