<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#Include file="Connection/Connection.asp"-->
<%
	if CheckMod <> 3  and CheckMod <> 2 then
	Response.Redirect "error_info.asp?err=21"
	else
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Administrator Website Nissan Techno Viet Nam</title>
<link rel="stylesheet" type="text/css" href="Images/css/admin.css" />
<script type="text/javascript" src="../Images/java/chrome.js"></script>

</head>

<body bgcolor="#FFFFFF" topmargin="0" bottommargin="0">

<a name="top" id= "top"></a>
<center>
<table width="780" height="100%" border="0" cellspacing="0" cellpadding="0" bgcolor="#B7B1BF" style="border-left:#333333 1px solid;border-right:#333333 1px solid;">
  	<tr>
    	<td width="780" height="90" valign="top" ><img src="../Images/logo_admin.jpg" width="780" height="90" /></td>
  	</tr>
  	<tr>
		<td width="690" bgcolor="#333333" height="25px;"><!--#Include file="Template/MenuNgang.asp"--><script type="text/javascript">cssdropdown.startchrome('menu')</script>	</td>
  	</tr>
 
	<tr>
		<td valign="top" bgcolor="#CCCCCC" style="padding:1px">
		<%
		tinhnang = Request.QueryString("tinhnang")
		Select Case tinhnang
			Case "edit":%>
					<!--#Include file="It_PhanNhom_edit.asp"-->
			<%Case "edits":%>
					<!--#Include file="It_PhanLoai_edit.asp"-->
			<%Case "list":%>
					<!--#Include file="It_PhanNhom_list.asp"-->
			<%Case "lists":%>
					<!--#Include file="It_PhanNhoms_list.asp"-->
			<%Case "add":%>
					<!--#Include file="It_PhanNhom_add.asp"-->
			<%Case "adds":%>
					<!--#Include file="It_PhanLoai_add.asp"-->
			<%Case else:%>
					<!--#Include file="It_PhanNhom_add.asp"-->
		  <%End Select%>	
		</td>
	</tr>
	<tr>
		<td height="23" bgcolor="#333333" style=" border-top:1px solid #FFFFFF;" class="Copyright"><!--#Include file="Template/LienHe.asp"--></td>
	</tr>  
</table>
</center>
</body>
</html>
<%end if%>
<!--#Include file="Connection/EndConnection.asp"-->