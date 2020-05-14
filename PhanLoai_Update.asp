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
		strSQL = "SELECT * FROM Loai WHERE Loai_Id = "&IDXoa
		rsCommon.Open strSQL,DK0HP,3,3
		If NOT rsCommon.EOF Then
		rsCommon.Delete
		rsCommon.Update
		End If
	rsCommon.Close
   Response.Redirect("PhanNhom_add.asp?tinhnang=lists&page="&pages)
End if

if Request.Form("Action")="SuaTin" then
	ID= Request.Form("txtMaNhom")
	IDNhom= Request.Form("txtIDL")
	txtTenVN = Request.Form("txtTenVN")
	txtTenNVJA = Request.Form("txtTenNVJA")
	txtNoiDungVN = Request.Form("txtNoiDungVN")
	txtNoiDungJA = Request.Form("txtNoiDungJA")
	txtLock = Request.Form("txtLock")
	txtOrder = Request.Form("txtOrder")
	
	Set rsCommon = Server.CreateObject("ADODB.Recordset")
	sSQL = "Select * from Loai WHERE Loai_Id = "&ID
	rsCommon.open sSQL,DK0HP,3,3
	If NOT rs.EOF Then
		rsCommon("Nhom_Id") = IDNhom
		rsCommon("LoaiVN_Name") = txtTenVN
		rsCommon("LoaiJA_Name") = txtTenNVJA
		rsCommon("LoaiVN_Body") = txtNoiDungVN
		rsCommon("LoaiJA_Body") = txtNoiDungJA
		if txtLock <>"" then
			rsCommon("Loai_Lock") = true
		Else
			rsCommon("Loai_Lock") = false
		End if
		rsCommon("Loai_Order") = txtOrder
		rsCommon.Update
		rsCommon.ReQuery
	end if
	rsCommon.Close()
	set rsCommon=Nothing
	Response.Redirect("PhanNhom_add.asp?tinhnang=lists&page="&page)
End if		

if Request.Form("Action")="TrangChu" then
For Each strID In Request.Form("idgroup")
		strStatus = Request.Form("Status"&strID)		
		Set rsconn = Server.CreateObject("ADODB.Recordset")
		strSQL = "SELECT * FROM Loai where Loai_Id="&strID
		rsconn.Open strSQL,DK0HP,3,3		
		If NOT rsconn.EOF Then
			If strStatus<>"" then
				rsconn("Loai_Lock") = true
			Else
				rsconn("Loai_Lock") = false
			End If
			rsconn.Update
		End If
		rsconn.Close
		set rsconn  =nothing
	Next
	For Each strID1 In Request.Form("ordergroup")
		strStatus1 = Request.Form("Statusorder"&strID1)		
		Set rsconn = Server.CreateObject("ADODB.Recordset")
		strSQL1 = "SELECT * FROM Loai where Loai_Id="&strID1
		rsconn.Open strSQL1,DK0HP,3,3		
		If NOT rsconn.EOF Then
			If strStatus1 <> "" Then
				rsconn("Loai_Order") = strStatus1
			End If
			rsconn.Update
		End If
		rsconn.Close
		set rsconn  =nothing
	Next
	Response.Redirect("PhanNhom_add.asp?tinhnang=lists")
end if	


' Them Thong Tin
if Request.Form("Action")="VietBai" then
	IDNhom= Request.Form("txtIDL")
	txtTenVN = Request.Form("txtTenVN")
	txtTenNVJA = Request.Form("txtTenNVJA")
	txtNoiDungVN = Request.Form("txtNoiDungVN")
	txtNoiDungJA = Request.Form("txtNoiDungJA")
	txtLock = Request.Form("txtLock")
	txtNV_ViTri = Request.Form("txtNV_ViTri")
		
	if txtTenVN<>"" then
		set rsCommon = Server.CreateObject("ADODB.Recordset")
		sSQL="select * from Loai"
		rsCommon.Open sSQL, DK0HP, 3, 3
		rsCommon.Addnew
		rsCommon("Nhom_Id") = IDNhom
		rsCommon("LoaiVN_Name") = txtTenVN
		rsCommon("LoaiJA_Name") = txtTenNVJA
		rsCommon("LoaiVN_Body") = txtNoiDungVN
		rsCommon("LoaiJA_Body") = txtNoiDungJA
		if txtLock <>"" then
			rsCommon("Loai_Lock") = true
		Else
			rsCommon("Loai_Lock") = false
		End if
		rsCommon("Loai_Order") = txtNV_ViTri
		rsCommon.Update
		rsCommon.Close()
		set rsCommon = nothing
		set rsCommon=Nothing
		Response.Redirect("PhanNhom_add.asp?tinhnang=lists")
		End if	
	End if	
end if
%>
