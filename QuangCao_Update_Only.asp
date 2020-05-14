<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#Include file="Connection/Connection.asp"-->
<%
	if CheckMod <> 3  and CheckMod <> 2 then
	Response.Redirect "error_info.asp?err=21"
	else
%>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">	
<%
	tinhnangs = Request.QueryString("tinhnang")
	ID = Request.QueryString("ID")
	page = Request.QueryString("page")
	Select Case tinhnangs
		Case "anhien":
			hiding = Request.QueryString("ah")
			set rscommon = Server.CreateObject("ADODB.Recordset")
			sSQL="select * from QuangCao where QuangCao_Id= "&ID
			rscommon.Open sSQL, DK0HP, 3, 3
				rscommon("QuangCao_Lock") = hiding 
				rscommon.Update
				rscommon.ReQuery
			rscommon.Close()
			set rscommon=Nothing
		
		End Select
		end if
		Response.Redirect("QuangCao_add.asp?tinhnang=list&page="&page)	
		%>