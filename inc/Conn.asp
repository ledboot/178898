<%
	''''--------���岿��------------------ 
'	Dim Fy_Post, Fy_Get, Fy_In, Fy_Inf, Fy_Xh, Fy_db, Fy_dbstr, immitAttack
'	immitAttack = False
'	''''�Զ�����Ҫ���˵��ִ�,�� "��" �ָ� 
'	Fy_In = "''''��'��and��exec��insert��select��delete��update��count��*��chr��mid��master��truncate��char��declare��|"
'	Fy_Inf = split(Fy_In, "��") 
'	If Request.Form <> "" Then 
'		For Each Fy_Post In Request.Form 
'			For Fy_Xh = 0 To Ubound(Fy_Inf) 
'				If Instr(LCase(Request.Form(Fy_Post)), Fy_Inf(Fy_Xh)) <> 0 Then
'					'Set rs = Server.CreateObject("ADODB.Recordset")
''					sql = "select * from immitAttackInfo"
''					rs.Open sql, conn, 1, 3
''					rs.AddNew
''					rs("immitAttackIP") = GetIp()
''					rs("immitAttackTime") = Now()
''					rs("immitAttackAddress") = Request.ServerVariables("URL")
''					rs("immitAttackMethod") = "POST"
''					rs("immitAttackParameter") = Fy_Post
''					rs("immitAttackValue") = Request.Form(Fy_Post)
'					Response.Write(Fy_Post)
'					Response.Write("<br>" & Request.Form(Fy_Post))
'					Response.Write("<br>" & Fy_Inf(Fy_Xh))
'					'Response.End()
'					'rs.Update
'					'rs.Close
'					'Set rs = Nothing
'					immitAttack = True
'				End If 
'			Next 
'		Next 
'	End If
'	If Request.QueryString <> "" Then 
'		For Each Fy_Get In Request.QueryString 
'			For Fy_Xh = 0 To Ubound(Fy_Inf) 
'				If Instr(LCase(Request.QueryString(Fy_Get)), Fy_Inf(Fy_Xh)) <> 0 Then
'					'Set rs = Server.CreateObject("ADODB.Recordset")
''					sql = "select * from immitAttackInfo"
''					rs.Open sql, conn, 1, 3
''					rs.AddNew
''					rs("immitAttackIP") = GetIp()
''					rs("immitAttackTime") = Now()
''					rs("immitAttackAddress") = Request.ServerVariables("URL")
''					rs("immitAttackMethod") = "GET"
''					rs("immitAttackParameter") = Fy_Get
''					rs("immitAttackValue") = Request.QueryString(Fy_Get)
'					Response.Write(Fy_Get)
'					Response.Write("<br>" & Request.Form(Fy_Get))
'					Response.Write("<br>" & Fy_Inf(Fy_Xh))
'					'Response.End()
'					'rs.Update
'					'rs.Close
'					'Set rs = Nothing
'					immitAttack = True
'				End If 
'			Next 
'		Next 
'	End If
'	If immitAttack Then
'		Response.Write("<Script Language=JavaScript>alert(""�벻Ҫ�ڲ����а����Ƿ��ַ�����ע�빥������վ��"");location.href=""../xytj_index.asp"";<
'	/Script>")
'		Response.End()
'	End If
%>
<%
	Const website="��ɳ�澩����"
	Const webAddress = "http://www.178898.com/"
	Const allRightsReservedStr = "��ɳ�澩����"
	meta_des="<meta name=""Description"" content=""�澩����רҵ��չ������ѵ��֤�ķ���,��֤�������Ŀ��,������ҵ�࿼֤����(�繤,����,���ӹ�,���乤,�泵,�г�˾��,���ŵ�,�Ǹ���ҵ,����˾��,���ػ�˾��,����˾�� �ź�,�����������ȵ���ѵ��֤����,������������ѵ��֤����,����ְҵ�ʸ���֤,ʮ����Ϊ��ɳ��������ҵ,����������,��һ�ع�,�����ؿ�,����������,�¹�������������ʮ����ҵ�ṩ���ʷ���,�ܵ�����λ��һ�º���!"">"
	meta_key="<meta name=""Keywords"" content=""�澩������ѵ����,��ɳ�繤������ѵѧУ,��ɳ�泵�г���ѵѧУ,��ɳ����������ѵ"">"
	meta_cop="<meta name=""Copyright"" content=""http://www.178898.com/"">"
	meta_aut="<meta name=""Author"" content=""Http://Www.Hunanidc.Net/"">"
	meta_pub="<meta name=""Publisher"" content=""A�ƻ�,Http://Www.Hunanidc.Net/"">"
	meta_gen="<meta name=""Generator"" content=""A�ƻ�,Http://Www.Hunanidc.Net/"">"
	meta_key_title = "��ɳ�澩����"
	
	'Const rightmouse = " oncontextmenu='return false;'"
	Const rightmouse = ""
	'Const selectmouse = " onselectstart='return false;'"
	Const selectmouse = ""
	Const bgcolor = " bgcolor=""#FFFFFF"""
	'Const bgcolor = " background=images/back.jpg"
	Const Domain=""
	Const DbPath="/Bob/"
	Const UpFiles = "/Upload/"
	Const Cloudy_FSO = "Scripting.FileSystemObject"
	Const Cloudy_RS = "Adodb.RecordSet"
	Const Cloudy_CONN = "Adodb.Connection"
	Const Cloudy_JPEG = "Persits.Jpeg"
	Const Cloudy_Upload = "Persits.Upload"
	
	function GetIp()
		'���������
		getclientip = Request.ServerVariables("HTTP_X_FORWARDED_FOR")
		If getclientip = "" Then
			getclientip = Request.ServerVariables("REMOTE_ADDR")
		end if
		GetIp = getclientip
	end function
	
	On Error Resume Next
	Dim strTemp,hk
	If Trim(Request.QueryString) <> "" Then strTemp =Trim(Request.QueryString)
	strTemp = LCase(strTemp)
	hk=0
	If Instr(strTemp,"%")<>0 then hk=1
	If Instr(strTemp,"%81")<>0 then hk=1
	If Instr(strTemp,"htw(")<>0 then hk=1
	If Instr(strTemp,"count(")<>0 then hk=1
	If Instr(strTemp,"htr(")<>0 then hk=1
	If Instr(strTemp,"asc(")<>0 then hk=1
	If Instr(strTemp,"mid(")<>0 then hk=1
	If Instr(strTemp,"and(")<>0 then hk=1
	If Instr(strTemp,"or(")<>0 then hk=1
	If Instr(strTemp,"char(")<>0 then hk=1
	If Instr(strTemp,"xp_cmdshell")<>0 then hk=1
	If Instr(strTemp,"'")<>0 then hk=1
	if hk=1 then
		Response.Write "�Բ���û�����ݣ�"
		response.end
		hk=0
	End If
	
	on error resume next
	Dim Conn,Connstr,Db
	Db="#Bob_1784fdW7tht898@#]%.mdb"
	Set conn = Server.CreateObject(Cloudy_CONN)
	Connstr="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" &Server.MapPath(Domain&DbPath&Db) & ";Jet OLEDB:Database Password='1788117HYJN898'"
	'User Id=admin;Password=;" 
	Conn.Open Connstr
	if err then
		response.Write(err.description)
		response.write "<div align='center'>���ݿ����������У����Ժ�</div>"
		response.end
	else
		Conn.CommandTimeOut = 999     '����Connection��ʱ 
		Server.ScriptTimeOut = 999    '���ýű���ʱ 
	end if
	
	Function CloseDatabase()
		Conn.close
		Set Conn = NOTHING
	End Function
%>
<!--#include file="Char.asp" -->