<!--#Include file="Connection/Connection.asp"-->
Nh&oacute;m ph&#7909;:&nbsp; <select  name="slLoaibai" class="inputadmin" style="width:150px ">
<%
IDNC  = Request.QueryString("gr")
if cstr(IDNC)="" then%>
<option value=""  selected > - All - </option>
<%ELSE%>
<option value=""  selected > - All - </option>
<%set rsTemp = server.CreateObject("ADODB.Recordset")
		strSQL = "select * from Loai WHERE Nhom_Id="&IDNC&" AND Loai_Lock=false Order by Loai_Order ASC"
		rsTemp.Open strSQL,DK0HP,3,3
		do while not rsTemp.EOF %>
<option value="<%=rsTemp("Loai_Id")%>"
			<%if CSTR(rsTemp("Loai_Id"))=(cstr(IDNP)) then %>selected <%End if%>> <%=rsTemp("LoaiVN_Name")%></option>
<%rsTemp.Movenext
		 Loop
	rsTemp.Close()
	Set rsTemp = nothing
	eND IF
	%>
</select>
<!--#Include file="Connection/EndConnection.asp"-->