<div class="rightContentForLXWM">
	<div class="lxwmInfo">公司地址：<%=ContactUsAddress%><br />联系电话：<%=ContactUsPhoneStr%><br />联系人：<%=ContactUsFax%><br />QQ咨询：<%=YWQQForPicStr%></div>
	<%
		If Not checkIsEmpty(ContactUsindeximgRs) Then
	%>
	<div class="lxwmImg"><a href="<%=(DoMain & ContactUsindeximgRs)%>" target="_blank"><img src="<%=(DoMain & ContactUsindeximgRs)%>" width="550" border="0" /></a></div>
	<%
		End If
	%>
</div>