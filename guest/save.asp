<!--#include file="gsconn.asp"-->
<!--#include file="char.asp"-->
<%
name1=HTMLEncode(trim(request("txtname")))
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

email1=HTMLEncode(request("txtemail"))
if email1<>"" then
if not isValidEmail(email1) then
   	   response.write "<script language='javascript'>"
   	   response.write "alert('电子邮件地址输入有误！');"
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
   	   response.write "alert('OICQ号码输入有误！');"
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
   response.write "alert('对不起，请您输入联系电话!');"
   response.write "history.go(-1);"
   response.write "</script>"
   response.end
end if
content1=HTMLEncode(request("txtContent"))
if content1="" then
   response.write "<script language='javascript'>"
   response.write "alert('对不起，请您输入留言!');"
   response.write "history.go(-1);"
   response.write "</script>"
   response.end
end if

set rs=server.createobject("adodb.recordset")
sql="select * from book"
rs.open sql,conngs,1,3
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
rs.update
response.redirect "display.asp"
%>
 