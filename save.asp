<!--#include file="inc/conn.asp"-->
<%
	if not isnumeric(request.form("passcode")) then
		Error1 "��֤����������֣�����ȷ��д��", " lyb.asp"
	end if
	passcode=Cint(request.form("passcode"))
	if CInt(passcode)<>CInt(Session("VerifyCode")) then
		Error1 "��֤�����", "lyb.asp"
	end if
	name1=HTMLEncode(trim(request("txtname")))
	if name1="" then
		response.write "<script language='javascript'>"
		response.write "alert('�Բ��������������!');"
		response.write "history.go(-1);"
		response.write "</script>"
		response.end
	end if
	i = 1
	len1 = 0
	charinner = mid(name1,1,1)
	while not charinner = ""
		if (charinner="'" or charinner="-" or charinner="%" or charinner="<" or charinner=">" or charinner="#" or charinner="&") then
			response.write "<script language='javascript'>"
			response.write "alert('�Բ������������в��Ϸ����ַ�������������!');"
			response.write "history.go(-1);"
			response.write "</script>"
			response.end
		else
			len1 = len1 + 1
		end if  
		i = i + 1
		charinner = mid(name1,i,1)
	wend
	email1=HTMLEncode(request("txtemail"))
	if email1<>"" then
		if not isValidEmail(email1) then
			response.write "<script language='javascript'>"
			response.write "alert('�����ʼ���ַ��������');"
			response.write "history.go(-1);"
			response.write "</script>"
			response.end
		end if
	end if
	homepage1=HTMLEncode(trim(request("txthomepage")))
	oicq1=HTMLEncode(trim(request("txtoicq")))
	if oicq1<>"" then
		if not isInteger(oicq1) or len(oicq1)<5 then
			response.write "<script language='javascript'>"
			response.write "alert('OICQ������������');"
			response.write "history.go(-1);"
			response.write "</script>"
			response.end
		end if
	end if
	'work1=HTMLEncode(trim(request("work")))
	color1=HTMLEncode(trim(request("color")))
	pic1=HTMLEncode(trim(request("sel_picname")))
	tel1=HTMLEncode(trim(request("tel")))
	if tel1="" then
		response.write "<script language='javascript'>"
		response.write "alert('�Բ�������������ϵ�绰!');"
		response.write "history.go(-1);"
		response.write "</script>"
		response.end
	end if
	content1=HTMLEncode(request("txtContent"))
	if content1="" then
		response.write "<script language='javascript'>"
		response.write "alert('�Բ���������������!');"
		response.write "history.go(-1);"
		response.write "</script>"
		response.end
	end if
	set rs=server.createobject("adodb.recordset")
	sql="select * from book"
	rs.open sql,conn,1,3
	rs.addnew
	rs("name")=name1
	rs("tel")=tel1
	rs("color")=color1
	rs("email")=email1
	rs("homepage")=homepage1
	rs("oicq")=oicq1
	rs("pic")=pic1
	rs("content")=content1
	rs("time")=now()
	rs("ip")=Request.ServerVariables("REMOTE_HOST")
	rs("shenhe") = 1
	rs.update
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
%>
<script language=javascript>
	alert("������Ϣ��ӳɹ�����Ϣ��˺󽫻���ʾ������");
	window.location.href="lyb.asp?lx=2";
</script> 