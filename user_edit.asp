<%
if CheckMod<>3 then
	Response.Redirect "error_info.asp?err=21"
end if

sSQL = "SELECT * FROM Admin WHERE Admin_Id=" & Request.QueryString("id")
Set adoRec = Server.CreateObject("ADODB.Recordset")
adoRec.Open sSQL, DK0HP, 3, 3

if adoRec.RecordCount<1 then
	Set adoRec = Nothing
	Response.End
end if

sUsername = adoRec("User")
nUserLevel = adoRec("Decentralization")
sUserRealName = adoRec("Admin_FullName")
sUserDept = adoRec("Admin_Phone")
varUserPw=adoRec("Password")
sGender=adoRec("Gender")
sAddress=adoRec("Admin_Email")

Set adoRec = Nothing
%>
<script language="javascript" src="calender_picker1.js"></script>
<script language="javascript">
function test()
{
if(EditMessage.Username.value=="") 
	{
	alert("Ban chua nhap ten dang nhap");
	EditMessage.Username.focus();
	}
else
	if(EditMessage.UserPw.value=="") 
		{
		alert("Ban chua nhap mat khau");
		EditMessage.UserPw.focus();
		}
	else
		if(EditMessage.UserPw.value!=EditMessage.ConfirmPw.value) 
			{
			alert("Mat khau xac nhan lai khong hop le");
			EditMessage.ConfirmPw.focus();
			}
		else
			if(EditMessage.UserRealName.value=="") 
				{
				alert("Ban chua nhap ho ten");
				EditMessage.UserRealName.focus();
				}
			else
				if(EditMessage.Gender.value=="-1") 
						{
						alert("Ban chua chon gioi tinh");
						EditMessage.Gender.focus();
						}
					else
						if(EditMessage.Address.value=="") 
							{
							alert("Ban chua nhap email");
							EditMessage.Address.focus();
							}
						else
							if(EditMessage.UserDept.value=="") 
								{
								alert("Ban chua nhap so dien thoai");
								EditMessage.UserDept.focus();
								}
							else
								EditMessage.submit();
			
}
</script>
<table border="0" cellpadding="7" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%">
  <tr>
    <td width="100%"><font size="3" color="#4A865A"><b>S&#7917;a thông tin ng&#432;&#7901;i dùng</b></font></td>
  </tr>
  <form name="EditMessage" method="POST" action="user_update.asp">
  <tr>
    <td width="100%">
    <table border="0" cellpadding="3" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%">
      <tr>
        <td width="0%" nowrap colspan="2" bgcolor="#DEDBCE"><b>Thông tin s&#7917; d&#7909;ng</b></td>
      </tr>
      <tr>
        <td width="0%" nowrap>Số ID :</td>
        <td width="100%"><input readonly type="text" name="id" value="<%=Request.QueryString("id")%>" style="font-family: Verdana; font-size: 8pt; width:200px"></td>
      </tr>
      <tr>
        <td width="0%" nowrap>Tên truy c&#7853;p:</td>
        <td width="100%">
        <input type="text" name="Username" size="30" style="font-family: Verdana; font-size: 8pt; width:200px" value="<%=sUsername%>"></td>
      </tr>
      <tr>
        <td width="0%" nowrap>M&#7853;t kh&#7849;u:</td>
        <td width="100%">
        <input type="password" name="UserPw" size="30" style="font-family: Verdana; font-size: 8pt; width:200px" value=""></td>
      </tr>
      <tr>
        <td width="0%" nowrap><font color="#808080">Nh&#7853;p l&#7841;i m&#7853;t kh&#7849;u:</font></td>
        <td width="100%">
        <input type="password" name="ConfirmPw" size="30" style="font-family: Verdana; font-size: 8pt; width:200px" value=""></td>
      </tr>
      <tr>
        <td width="0%" nowrap>Quy&#7873;n:</td>
        <td width="100%"><select size="1" name="UserLevel" style="font-family: Verdana; font-size: 8pt; width:200px">
          <option value="3" <%if nUserLevel=3 then Response.Write "selected=selected" end if%> >Website administrator</option>
          <option value="2" <%if nUserLevel=2 then Response.Write "selected=selected" end if%>>Members only read and post</option>
        </select></td>
      </tr>
    </table>
    </td>
  </tr>
  <tr>
    <td width="100%">
    <table border="0" cellpadding="3" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%">
      <tr>
        <td width="0%" nowrap colspan="2" bgcolor="#DEDBCE"><b>Thông tin cá nhân</b></td>
      </tr>
      <tr>
        <td width="0%" nowrap>&nbsp;</td>
        <td width="100%">&nbsp;
        </td>
      </tr>
      <tr>
        <td width="0%" nowrap>H&#7885; và tên:</td>
        <td width="100%">
        <input type="text" name="UserRealName" size="50" style="font-family: Verdana; font-size: 8pt" value="<%=sUserRealName%>"></td>
      </tr>
	  <tr>
        <td width="0%" nowrap>Gi&#7899;i t&iacute;nh:</td>
        <td width="100%">        
		<select name="Gender" id="Gender">
          <option value="1" <%if(sGender=true) then Response.Write"selected" end if%>  selected="selected">Nam</option>
          <option value="0" <%if(sGender=false) then Response.Write"selected" end if%> >N&#7919;</option>
        </select></td>
      </tr>
	  
	  <tr>
        <td width="0%" nowrap>Email:</td>
        <td width="100%">
        <input name="Address" type="text" id="Address" style="font-family: Verdana; font-size: 8pt" size="50" value="<%=sAddress%>"></td>
      </tr>

	  
	  
      <tr>
        <td width="0%" nowrap>Phone:</td>
        <td width="100%">
        <input type="text" name="UserDept" size="50" style="font-family: Verdana; font-size: 8pt" value="<%=sUserDept%>"></td>
      </tr>
      </table>
    </td>
  </tr>
  <tr>
    <td width="100%">
    <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%" bgcolor="#DEDBCE">
      <tr>
        <td width="100%">
        <table border="0" cellpadding="3" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111">
          <tr>
            <td>
            <button  onClick="JavaScript:test();" name="B2" style="width: 80; height: 22"><b>
            <font face="Verdana" size="1">C&#7853;p nh&#7853;t</font></b></button></td>
            <td>&nbsp;
            </td>
            <td>&nbsp;</td>
            <td>&nbsp;</td>
          </tr>
        </table>
        </td>
      </tr>
    </table>
    </td>
  </tr>
  
  <input type="hidden" name="act" value="UPDATE">
  </form>
</table>