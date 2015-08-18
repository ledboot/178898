<div class="jqjs_C">
	<ul>
		<%
			listCount = 0
			pageno = CInt(Request("pageno"))
			If pageno = 0 Or IsEmpty(pageno) Or IsNull(pageno) Then
				pageno = 0
			End If
			Set rs = rstw2(lmname_1)
			RCount = 4
			RRC = CInt(rs.RecordCount)
			If RRC > 0 Then
				if CInt(pageno) = 0 then pageno = 1
				rs.pagesize = RCount
				rs.absolutepage = pageno
				RPC = CInt(rs.pagecount)
				if not rs.bof and not rs.eof then
					for i = 1 to RCount
						listCount = listCount + 1
						
						ArticleIDRs = rs("ArticleID")
						TitleRs = rs("Title")
						indeximg1Rs = rs("indeximg1")
						ContentRs = rs("Content")
						
						ContentRs = Replace(ContentRs, "<P>", "")
						ContentRs = Replace(ContentRs, "</P>", "<br />")
		%>
		<li><a href="<%=showNavPage%>?id=<%=ArticleIDRs%>" target="_blank" title="<%=TitleRs%>"><span><img src="<%=(DoMain & indeximg1Rs)%>" alt="<%=TitleRs%>" width="204" height="143" /></span><p><%=LeftTrue(TitleRs, 64)%></p></a></li>
		<%
						rs.MoveNext
						If rs.EOF Then
							Exit For
						End If
					Next
				End If
			End If
		%>
	</ul>
	<div style="height:10px; width:700px; clear:both"></div>
</div>