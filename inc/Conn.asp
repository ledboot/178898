<%
	''''--------定义部份------------------ 
'	Dim Fy_Post, Fy_Get, Fy_In, Fy_Inf, Fy_Xh, Fy_db, Fy_dbstr, immitAttack
'	immitAttack = False
'	''''自定义需要过滤的字串,用 "防" 分隔 
'	Fy_In = "''''防'防and防exec防insert防select防delete防update防count防*防chr防mid防master防truncate防char防declare防|"
'	Fy_Inf = split(Fy_In, "防") 
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
'		Response.Write("<Script Language=JavaScript>alert(""请不要在参数中包含非法字符尝试注入攻击本网站！"");location.href=""../xytj_index.asp"";<
'	/Script>")
'		Response.End()
'	End If
%>
<%
	Const website="长沙湘京教育"
	Const webAddress = "http://www.178898.com/"
	Const allRightsReservedStr = "长沙湘京教育"
	meta_des="<meta name=""Description"" content=""湘京教育专业开展教育培训考证的服务,考证服务的项目有,特种作业类考证服务(电工,焊工,架子工,制冷工,叉车,行车司机,龙门吊,登高作业,塔吊司机,起重机司机,起重司索 信号,物料提升机等的培训考证服务,建筑工程类培训考证服务,国家职业资格认证,十年来为长沙地区的企业,中铁建集团,三一重工,中联重科,菲亚特汽车,德国大众汽车等数十家企业提供优质服务,受到各单位的一致好评!"">"
	meta_key="<meta name=""Keywords"" content=""湘京教育培训中心,长沙电工焊工培训学校,长沙叉车行车培训学校,长沙建筑工程培训"">"
	meta_cop="<meta name=""Copyright"" content=""http://www.178898.com/"">"
	meta_aut="<meta name=""Author"" content=""Http://Www.Hunanidc.Net/"">"
	meta_pub="<meta name=""Publisher"" content=""A计划,Http://Www.Hunanidc.Net/"">"
	meta_gen="<meta name=""Generator"" content=""A计划,Http://Www.Hunanidc.Net/"">"
	meta_key_title = "长沙湘京教育"
	
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
		'代理服务器
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
		Response.Write "对不起，没有数据！"
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
		response.write "<div align='center'>数据库正在整理中，请稍后！</div>"
		response.end
	else
		Conn.CommandTimeOut = 999     '设置Connection超时 
		Server.ScriptTimeOut = 999    '设置脚本超时 
	end if
	
	Function CloseDatabase()
		Conn.close
		Set Conn = NOTHING
	End Function
%>
<!--#include file="Char.asp" -->