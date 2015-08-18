<div class="Selection">
	<%
		if pageno > 1 then
	%>
	<span><a href="<%=NavPage_1%>&pageno=1"><img src="images/Selection_03.jpg" /></a>&nbsp;</span>
	<span><a href="<%=NavPage_1%>&pageno=<%=pageno-1%>"><img src="images/Selection_05.jpg" /></a>&nbsp;</span>
	<%
		Else
	%>
	<span><img src="images/Selection_03.jpg" />&nbsp;</span>
	<span><img src="images/Selection_05.jpg" />&nbsp;</span>
	<%
		End If
	%>
	<span>
		<%
			For i = 1 To RPC
		%>
		<a href="<%=NavPage_1%>&pageno=<%=i%>"><%=i%></a>&nbsp;
		<%
			Next
		%>
	</span>
	<%
		if RPC > pageno then
	%>
	<span><a href="<%=NavPage_1%>&pageno=<%=pageno+1%>"><img src="images/Selection_07.jpg" /></a>&nbsp;</span>
	<span><a href="<%=NavPage_1%>&pageno=<%=RPC%>"><img src="images/Selection_09.jpg" /></a></span>
	<%
		Else
	%>
	<span><img src="images/Selection_07.jpg" />&nbsp;</span>
	<span><img src="images/Selection_09.jpg" /></span>
	<%
		End If
	%>
</div>