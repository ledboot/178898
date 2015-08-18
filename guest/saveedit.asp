<!--#include file="gsconn.asp"-->
<!--#include file="char.asp"-->
<%
if session("AdminName") = "" then
response.redirect "../login.asp"
end if
%>
<%
name1=HTMLEncode(request.form("name"))
if name1="" then
   response.write "<script language='javascript'>"
   response.write "alert('对不起，请您输入大名!');"
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
   response.write "alert('对不起，您的名字有不合法的字符，请重新输入!');"
   response.write "history.go(-1);"
   response.write "</script>"
   response.end
   else
   len1 = len1 + 1
   end if  
   i = i + 1
   charinner = mid(name1,i,1)
wend
email1=HTMLEncode(request.form("email"))
if email1<>"" then
if not isValidEmail(email1) then
   	   response.write "<script language='javascript'>"
   	   response.write "alert('电子邮件地址输入有误！');"
   	   response.write "history.go(-1);"
   	   response.write "</script>"
   	   response.end
end if
end if
oicq1=HTMLEncode(request.form("oicq"))
if oicq1<>"" then
if not isInteger(oicq1) or len(oicq1)<5 then
   	   response.write "<script language='javascript'>"
   	   response.write "alert('OICQ号码输入有误！');"
   	   response.write "history.go(-1);"
   	   response.write "</script>"
   	   response.end
end if
end if
homepage1=HTMLEncode(request.form("homepage"))
content1=HTMLEncode(request.form("content"))
id=request("id")

set rs=server.createobject("adodb.recordset")
sql="select * from book where id="&id
rs.open sql,conngs,3,3
rs("name")=name1
rs("email")=email1
rs("oicq")=oicq1
rs("homepage")=homepage1
rs("content")=content1
rs.update
rs.close
set rs=noting
conn.close
set conn=nothing

response.redirect "manager.asp"
%>

