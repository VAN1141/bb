<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<%
	gourl= Request.QueryString("url")
	message = Request.QueryString("message")
%>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>DK0HomePage</title>
<meta http-equiv="Refresh" content="2; URL=<%=gourl%>">
</head>

<body topmargin="3" bottommargin="1">
<div align="center" style="top:100px;">
<table bgcolor="#FF9966" width="400px"  style=" border:solid 1px #666666; border-top-style:none ; border-collapse:collapse" cellspacing="0" cellpadding="3">
                  <tr>
                    <td align="center" height="60px" style="color:#FFFFFF; font-family:Arial, Helvetica, sans-serif; font-size:12px; font-weight:bold "><%=message%></td>
                  </tr>
                  <tr>
                    <td align="center" height="20"><img src="Images/css/loadinganimationliferaysd8.gif"  width="70" height="10" border="0"></td>
                  </tr>                 
</table>
</div>
</body>
</html>