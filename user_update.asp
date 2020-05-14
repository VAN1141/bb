<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#Include file="Connection/Connection.asp"-->
<!--#include file="inc_sha256.asp"-->
<%
if CheckMod<>3 then
	Response.Redirect "error_info.asp?err=21"
end if
LayID = Request.Form("id")
sUsername = UCase(Trim(Request.Form("Username")))
sUsername = SQLPatch(sUsername)
sAction = Trim(UCase(Request.Form("act")))

if Trim(sUsername)="" then Response.Redirect "error_info.asp?err=90"

sUserPw = Request.Form("UserPw")
sUserPw = SQLPatch(sUserPw)
sConfirmPw = Request.Form("ConfirmPw")
sConfirmPw = SQLPatch(sConfirmPw)

if sUserPw<>sConfirmPw then Response.Redirect "error_info.asp?err=91"
	Set rs = Server.CreateObject("ADODB.Recordset")
	sSQL = ("Select * from Admin WHERE Admin_Id = "&LayID)
	rs.open sSQL,DK0HP,3,3
if sAction="UPDATE" then
	If NOT rs.EOF Then
	rs("User") = sUsername
	if Trim(sUserPw)<>"" then
		rs("Password") = SHA256(sUserPw)
	end if
	rs("Admin_FullName") = Request.Form("UserRealName")
	if(Trim(Request.Form("Gender"))="1") then
		rs("Gender") = true
	else
		rs("Gender") = false
	end if
	rs("Admin_Email") = Request.Form("Address")
	rs("Admin_Phone") = Request.Form("UserDept")
	rs("Decentralization") = CLng(Request.Form("UserLevel"))
	rs.Update
	rs.ReQuery
end if
rs.close
Set rs = Nothing
Response.Redirect "Admin_add.asp?tinhnang=list"
end if
%>