<!--#include file="inc/Conn.asp" -->
<!--#include file="news/Admin_Config.asp" -->
<%
	stuName = HTMLEncode(Request("stuName"))
	stuTelephone = HTMLEncode(Request("stuTelephone"))
	'Response.Write(stuDeclare_Class_Type_ID & "<br>")
	stuResume_Text = HTMLEncode(Request("stuResume_Text"))
	stuAddress = HTMLEncode(Request("stuAddress"))
	stuQQNumber = HTMLEncode(Request("stuQQNumber"))
	stuEMail = HTMLEncode(Request("stuEMail"))
	
	if not isnumeric(request.form("passcode")) then
		Error1 "��֤����������֣�����ȷ��д��", " zxbm.asp"
	end if
	passcode=Cint(request.form("passcode"))
	if CInt(passcode)<>CInt(Session("VerifyCode")) then
		Error1 "��֤�����", "zxbm.asp"
	end if
	
	Set rs = Server.CreateObject("ADODB.Recordset")
	sql = "SELECT id, stuName, stuSex, stuBirthday, stuI_D_Card, stuCensus_Register_Provinces_ID, stuCensus_Register_City_ID, stuGraduation_School, stuCultural_Degree_ID, stuHome_Address, stuTelephone, stuZip, stuHeight, stuEyesight_Left, stuEyesight_Right, stuDeclare_Class_Type_ID, stuResume_Text, RegisterDatatime, stuAddress, stuQQNumber, stuEMail FROM registrationInfo"
	rs.Open sql, conn, 1, 3
	rs.AddNew
	rs("stuName") = stuName
	rs("stuTelephone") = stuTelephone
	'Response.Write("stuDeclare_Class_Type_ID" & stuDeclare_Class_Type_ID)
	'Response.End()
	rs("stuResume_Text") = stuResume_Text
	rs("RegisterDatatime") = Now()
	rs("stuAddress") = stuAddress
	rs("stuQQNumber") = stuQQNumber
	rs("stuEMail") = stuEMail
	rs.Update
	rs.Close
	Set rs = Nothing
%>
<%
	Call CloseDatabase()
	Call Error1("��ϲ�������ѳɹ������˸�н��ҵ�ĵ�һ�������ǽ���24Сʱ�ھ�������ȡ����ϵ��", "zxbm.asp")
%>