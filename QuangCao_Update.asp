<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#Include file="Connection/Connection.asp"-->
<%
	if CheckMod <> 3  and CheckMod <> 2 then
	Response.Redirect "error_info.asp?err=21"
	else
%>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">	
<%
'Xoa thong tin
IDXoa = Request.QueryString("IDXoa")
IF IDXoa<>"" Then
		Set rsCommon = Server.CreateObject("ADODB.Recordset")
		strSQL = "SELECT * FROM QuangCao WHERE QuangCao_Id = "&IDXoa
		rsCommon.Open strSQL,DK0HP,3,3
		If NOT rsCommon.EOF Then
		rsCommon.Delete
		rsCommon.Update
		End If
	rsCommon.Close
   Response.Redirect("QuangCao_add.asp?tinhnang=list&page="&Request.QueryString("page"))
End if

if Request.Form("Action")="SuaTin" then
	ID = Request.Form("IDsua")
	txtQC_ViTri= Request.Form("txtQC_ViTri")
	txtLinkQC = Request.Form("txtLinkQC")
	checkVT = Request.Form("checkVT")
	news_image = Request.Form("news_image")
	txtMoTaQC = Request.Form("txtMoTaQC")
	txtAn = Request.Form("txtAn")
	
	Set rsCommon = Server.CreateObject("ADODB.Recordset")
	sSQL = "Select * from QuangCao WHERE QuangCao_Id = "&ID
	rsCommon.open sSQL,DK0HP,3,3
	If NOT rs.EOF Then
		rsCommon("QuangCao_Link") = txtLinkQC
		rsCommon("QuangCao_Img") = news_image
		rsCommon("QuangCao_Title") = txtMoTaQC
		rsCommon("QuangCao_Location") = checkVT
		
		if txtAn <>"" then
			rsCommon("QuangCao_Lock") = true
		Else
			rsCommon("QuangCao_Lock") = false
		End if
		rsCommon("QuangCao_Order") = txtQC_ViTri
		rsCommon.Update
		rsCommon.ReQuery
	end if
	rsCommon.Close()
	set rsCommon=Nothing
	Response.Redirect("QuangCao_add.asp?tinhnang=list&page="&Request.QueryString("page"))
End if		




' Them Thong Tin
if Request.Form("Action")="VietBai" then
	txtQC_ViTri= Request.Form("txtQC_ViTri")
	txtLinkQC = Request.Form("txtLinkQC")
	checkVT = Request.Form("checkVT")
	news_image = Request.Form("news_image")
	txtMoTaQC = Request.Form("txtMoTaQC")
	txtAn = Request.Form("txtAn")
			
	if txtLinkQC<>"" then
		set rsCommon = Server.CreateObject("ADODB.Recordset")
		sSQL="select * from QuangCao"
		rsCommon.Open sSQL, DK0HP, 3, 3
		rsCommon.Addnew
		rsCommon("QuangCao_Link") = txtLinkQC
		rsCommon("QuangCao_Img") = news_image
		rsCommon("QuangCao_Title") = txtMoTaQC
		rsCommon("QuangCao_Location") = checkVT
		
		if txtAn <>"" then
			rsCommon("QuangCao_Lock") = true
		Else
			rsCommon("QuangCao_Lock") = false
		End if
		rsCommon("QuangCao_Order") = txtQC_ViTri
		
		rsCommon.Update
		rsCommon.Close()
		set rsCommon = nothing
		set rsCommon=Nothing
		Response.Redirect("QuangCao_add.asp?tinhnang=list")
		End if	
	End if	
end if
%>
