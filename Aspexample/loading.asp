<!-- #include file="inc_i18n.asp" -->
<html>
<head>
<title><%=Session("StudioName")%> Online</title>
<meta http-equiv="Content-Type" content="text/html">
	<!-- #include file="inc_date_ctrl.asp" -->
</head>
<body onLoad="document.location='<%=request.querystring("id")%>.asp';" >
<div align="left">
<table height="100%" width="<%=strPageWidth%>" cellspacing="0">
  <tr> 
    <td class="center-ch" valign="top" height="100%" width="100%" style="background-color:<%=session("pageColor2")%>;"> <br />
        <table width="80%" cellspacing="0">
          <tr>
            <td class="headText" align="left">&nbsp;</td>
          </tr>
          <tr height="100%"> 
            <td height="100%" class="mainTextBig center-ch"> 
              <table height="100%" class="mainText" width="90%" cellspacing="0">
                <tr height="100%"> 
                  <td height="100%" align="left" valign="top">
                    <table width="100%" cellspacing="0" height="100%">
                      <tr>
                        <td class="center-ch" valign="middle"><br />
                          <br />
                          <br />
                          <br />
                          <br />
                          <br />
                          <img src="<%= contentUrl("/asp/images/loading.gif") %>" width="111" height="43" />
                         </td>
                      </tr>
                    </table>
                  </td>
                </tr>
              </table></td>
          </tr>
        </table>
	  </td>
    </tr>
</table>
</body>
</html>
