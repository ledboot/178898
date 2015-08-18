<div style="height:15px; width:745px; clear:both"></div>
<div class="gsjj_C">
	<%
		if checkIsEmpty(contentSinglePageIndeximg) then
		Else
	%>
	<div class="gsjj_CTp"><a href="<%=(DoMain & contentSinglePageIndeximg)%>" title="<%=contentSinglePageTitle%>" target="_blank"><img src="<%=(DoMain & contentSinglePageIndeximg)%>" alt="<%=contentSinglePageTitle%>" border="0" width="600" /></a></div>
	<%
		End If
	%>
	<div class="gsjj_CMine"><%=contentSinglePageContent%></div>
</div>