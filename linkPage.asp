<div class="yqlj_Bg">
	<div class="yqlj">
		<b>”—«È¡¥Ω”£∫</b>
		<%
			iCount = 0
			linkStr = ""
			Set rs = Server.CreateObject("ADODB.Recordset")
			sql = "SELECT top 24 id, linkname, linkwww, classname, linkidorder, lx FROM Link where lx = 1 order by id desc"
			rs.Open sql, conn, 1, 3
			While Not rs.BOF And Not rs.EOF
				iCount = iCount + 1
				
				linknameRs = rs("linkname")
				linkwwwRs = rs("linkwww")
				classnameRs = rs("classname")
				
				If iCount = 1 Then
					linkStr = "<a href=""" & linkwwwRs & """ target=""_blank"">" & linknameRs & "</a>"
				Else
					linkStr = "&nbsp;&nbsp;|&nbsp;&nbsp;<a href=""" & linkwwwRs & """ target=""_blank"">" & linknameRs & "</a>"
				End If
		%>
		<%=linkStr%>
		<%
				rs.MoveNext
			Wend
			rs.Close
			Set rs = Nothing
		%>
	</div>
</div>