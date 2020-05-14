<table width="220px" border="0" cellspacing="0" cellpadding="0">
      <tr>
         <td height="25" class="MenuB"><%=langin("Nội dung","内容")%></td>
      </tr>
	  <% 
			Set DK0Conn = Server.CreateObject("ADODB.Recordset")
			DK1text = "Select * From Nhom Where Nhom_Id<>11 and Nhom_Lock = False Order by Nhom_Order ASC"
			DK0Conn.open DK1text,DK0HP
			Do while not DK0Conn.EOF
			IDNhomBV=DK0Conn("Nhom_Id")
		%>
      <tr>
        <td height="22" class="MenuHome">&nbsp;<%=langin(DK0Conn("NhomVN_Name"),DK0Conn("NhomJA_Name"))%></td>
      </tr>
	  <tr>
        <td>
			<table width="100%" border="0" cellspacing="0" cellpadding="0">
			<%
				Set DK0Temp = Server.CreateObject("ADODB.Recordset")
				DK2text = "Select * From Loai Where Loai_Lock = False and Nhom_Id="&IDNhomBV&" Order By Loai_Order ASC"
				DK0Temp.open DK2text,DK0HP
				Do while not DK0Temp.EOF
			%>
			  <tr>
				 <td height="22" class="MenuHome1"><a class="MenuHome1" href="Group.asp?ID=<%=DK0Temp("Loai_Id")%>&gr=<%=IDNhomBV%>" target="_self">&nbsp;<%=langin(DK0Temp("LoaiVN_Name"),DK0Temp("LoaiJA_Name"))%></a></td>
			  </tr>
			  <% 
				DK0Temp.MoveNext
				LOOP
				DK0Temp.Close
				Set DK0Temp = nothing	
			%>
			</table>
		 </td>
      </tr>
	 <%
	 	DK0Conn.MoveNext
		LOOP
		DK0Conn.Close
		Set DK0Conn = Nothing
	%>
		
	  <tr>
        <td height="25" class="MenuB"><%=langin("Hình ảnh nổi bật","ハイライト写真")%></td>
      </tr>
	   <tr>
        <td align="center" style="padding:5px;">
		<div id="slides_left">
		<div class="slides_container">
		<%	
				Set DK0Pic = Server.CreateObject("ADODB.Recordset")
				DK3text = "Select * From QuangCao Where QuangCao_Lock = False and QuangCao_Location = 1 Order By QuangCao_Order ASC"
				DK0Pic.open DK3text,DK0HP
				if DK0Pic.recordcount <=0 then
					response.Write("<br>")
				else
				Do while not DK0Pic.EOF
		%>
		<div style="text-align:center;">
		<a href="<%=DK0Pic("QuangCao_Link")%>" target="_blank" title="<%=DK0Pic("QuangCao_Title")%>"><img src="<%=DK0Pic("QuangCao_Img")%>" border="0" alt="<%=DK0Pic("QuangCao_Title")%>"></a></div>
		<% 
				DK0Pic.MoveNext
				LOOP
				end if
				DK0Pic.Close
				Set DK0Pic = nothing	
			%>
			</DIV></div>
		</td>
      </tr>
	  
	  <tr>
			<td class="MenuB" align="left"><%=langin("Bài viết quan tâm","注目記事")%></td>
		</tr>
		<tr>
			<td><!--#Include file="topview.asp"--></td>
		</tr>
	  
      <tr>
         <td height="25" class="MenuB" target="_self"><%=langin("WEB LINK","WEB LINK")%></td>
      </tr>
      <tr>
         <td align="center" style="padding:5px;">
		 <%	
				Set DK0Pic = Server.CreateObject("ADODB.Recordset")
				DK4text = "Select * From QuangCao Where QuangCao_Lock = False and QuangCao_Location = 2 Order By QuangCao_Order ASC"
				DK0Pic.open DK4text,DK0HP
				if DK0Pic.recordcount <=0 then
					response.Write("<br>")
				else
				Do while not DK0Pic.EOF
		%>
		 <a href="<%=DK0Pic("QuangCao_Link")%>" target="_blank" title="<%=DK0Pic("QuangCao_Title")%>"><img src="<%=DK0Pic("QuangCao_Img")%>" width=142 height=76 border="0" alt="<%=DK0Pic("QuangCao_Title")%>" /><br />
		 <% 
				DK0Pic.MoveNext
				LOOP
				end if
				DK0Pic.Close
				Set DK0Pic = nothing	
			%>
		 </td>
      </tr>
    </table>