<div class="rightContentForLXWM">
	<div class="lxwmInfo">��˾��ַ��<%=ContactUsAddress%><br />��ϵ�绰��<%=ContactUsPhoneStr%><br />��ϵ�ˣ�<%=ContactUsFax%><br />QQ��ѯ��<%=YWQQForPicStr%></div>
	<%
		If Not checkIsEmpty(ContactUsindeximgRs) Then
	%>
	<div class="lxwmImg"><a href="<%=(DoMain & ContactUsindeximgRs)%>" target="_blank"><img src="<%=(DoMain & ContactUsindeximgRs)%>" width="550" border="0" /></a></div>
	<%
		End If
	%>
</div>