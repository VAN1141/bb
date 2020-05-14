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
page = Request.Form("txtpage")
IDXoa = Request.QueryString("IDXoa")
pages = Request.QueryString("page")
IF IDXoa<>"" Then
		Set rsCommon = Server.CreateObject("ADODB.Recordset")
		strSQL = "SELECT * FROM Nhom WHERE Nhom_Id = "&IDXoa
		rsCommon.Open strSQL,DK0HP,3,3
		If NOT rsCommon.EOF Then
		rsCommon.Delete
		rsCommon.Update
		End If
	rsCommon.Close
   Response.Redirect("PhanNhom_add.asp?tinhnang=list&page="&pages)
End if

if Request.Form("Action")="SuaTin" then
	ID= Request.Form("txtMaNhom")
	txtTenVN = Request.Form("txtTenVN")
	txtTenNVJA = Request.Form("txtTenNVJA")
	txtLock = Request.Form("txtLock")
	txtOrder = Request.Form("txtOrder")
	
	Set rsCommon = Server.CreateObject("ADODB.Recordset")
	sSQL = "Select * from Nhom WHERE Nhom_Id = "&ID
	rsCommon.open sSQL,DK0HP,3,3
	If NOT rs.EOF Then
		rsCommon("NhomVN_Name") = txtTenVN
		rsCommon("NhomJA_Name") = txtTenNVJA
		if txtLock <>"" then
			rsCommon("Nhom_Lock") = true
		Else
			rsCommon("Nhom_Lock") = false
		End if
		rsCommon("Nhom_Order") = txtOrder
		rsCommon.Update
		rsCommon.ReQuery
	end if
	rsCommon.Close()
	set rsCommon=Nothing
	Response.Redirect("PhanNhom_add.asp?tinhnang=list&page="&page)
End if		


if Request.Form("Action")="TrangChu" then
For Each strID In Request.Form("idgroup")
		strStatus = Request.Form("Status"&strID)		
		Set rsconn = Server.CreateObject("ADODB.Recordset")
		strSQL = "SELECT * FROM Nhom where Nhom_Id="&strID
		rsconn.Open strSQL,DK0HP,3,3		
		If NOT rsconn.EOF Then
			If strStatus<>"" then
				rsconn("Nhom_Lock") = true
			Else
				rsconn("Nhom_Lock") = false
			End If
			rsconn.Update
		End If
		rsconn.Close
		set rsconn  =nothing
	Next
	For Each strID1 In Request.Form("ordergroup")
		strStatus1 = Request.Form("Statusorder"&strID1)		
		Set rsconn = Server.CreateObject("ADODB.Recordset")
		strSQL1 = "SELECT * FROM Nhom where Nhom_Id="&strID1
		rsconn.Open strSQL1,DK0HP,3,3		
		If NOT rsconn.EOF Then
			If strStatus1 <> "" Then
				rsconn("Nhom_Order") = strStatus1
			End If
			rsconn.Update
		End If
		rsconn.Close
		set rsconn  =nothing
	Next
	Response.Redirect("PhanNhom_add.asp?tinhnang=list")
end if	

' Them Thong Tin
if Request.Form("Action")="VietBai" then
	ID= Request.Form("txtMaNhom")
	txtTenVN = Request.Form("txtTenVN")
	txtTenNVJA = Request.Form("txtTenNVJA")
	txtLock = Request.Form("txtLock")
	txtNV_ViTri = Request.Form("txtNV_ViTri")
		
	if txtTenVN<>"" then
		set rsCommon = Server.CreateObject("ADODB.Recordset")
		sSQL="select * from Nhom"
		rsCommon.Open sSQL, DK0HP, 3, 3
		rsCommon.Addnew
		rsCommon("NhomVN_Name") = txtTenVN
		rsCommon("NhomJA_Name") = txtTenNVJA
		if txtLock <>"" then
			rsCommon("Nhom_Lock") = true
		Else
			rsCommon("Nhom_Lock") = false
		End if
		rsCommon("Nhom_Order") = txtNV_ViTri
		rsCommon.Update
		rsCommon.Close()
		set rsCommon = nothing
		set rsCommon=Nothing
		Response.Redirect("PhanNhom_add.asp?tinhnang=list")
		End if	
	End if	
end if
%>
