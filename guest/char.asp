<%
 function rechar(string)
 ON ERROR RESUME NEXT
 string = Replace(string ,"'","’")
 string = Replace(string ,"&","＆")
 string = Replace(string ,"<","＜") 
 string = Replace(string ,">","＞") 
 string = Replace(string ,"0x","Ox") 
 string = Replace(string ,"0X","Ox")
 string = Replace(string ,"#","＃") 
 rechar=string
 end function 
%>
<%
Rem ==========通用函数=========
Rem 判断数字是否整形
function isInteger(para)
       on error resume next
       dim str
       dim l,i
       if isNUll(para) then 
          isInteger=false
          exit function
       end if
       str=cstr(para)
       if trim(str)="" then
          isInteger=false
          exit function
       end if
       l=len(str)
       for i=1 to l
           if mid(str,i,1)>"9" or mid(str,i,1)<"0" then
              isInteger=false 
              exit function
           end if
       next
       isInteger=true
       if err.number<>0 then err.clear
end function

function IsValidTel(para)
       on error resume next
       dim str
       dim l,i
       if isNUll(para) then 
          IsValidTel=false
          exit function
       end if
       str=cstr(para)
       if len(trim(str))<7 then
          IsValidTel=false
          exit function
       end if
       l=len(str)
       for i=1 to l
           if not (mid(str,i,1)>="0" and mid(str,i,1)<="9" or mid(str,i,1)="-") then
              IsValidTel=false 
              exit function
           end if
       next
       IsValidTel=true
       if err.number<>0 then err.clear
   end function


function IsValidEmail(email)

dim names, name, i, c

'Check for valid syntax in an email address.

IsValidEmail = true
names = Split(email, "@")
if UBound(names) <> 1 then
   IsValidEmail = false
   exit function
end if
for each name in names
   if Len(name) <= 0 then
     IsValidEmail = false
     exit function
   end if
   for i = 1 to Len(name)
     c = Lcase(Mid(name, i, 1))
     if InStr("abcdefghijklmnopqrstuvwxyz_-.", c) <= 0 and not IsNumeric(c) then
       IsValidEmail = false
       exit function
     end if
   next
   if Left(name, 1) = "." or Right(name, 1) = "." then
      IsValidEmail = false
      exit function
   end if
next
if InStr(names(1), ".") <= 0 then
   IsValidEmail = false
   exit function
end if
i = Len(names(1)) - InStrRev(names(1), ".")
if i <> 2 and i <> 3 then
   IsValidEmail = false
   exit function
end if
if InStr(email, "..") > 0 then
   IsValidEmail = false
end if

end function

 function strLength(str)
       ON ERROR RESUME NEXT
       dim WINNT_CHINESE
       WINNT_CHINESE    = (len("同学录")=3)
       if WINNT_CHINESE then
          dim l,t,c
          dim i
          l=len(str)
          t=l
          for i=1 to l
             c=asc(mid(str,i,1))
             if c<0 then c=c+65536
             if c>255 then
                t=t+1
             end if
          next
          strLength=t
       else 
          strLength=len(str)
       end if
       if err.number<>0 then err.clear
   end function 

Rem 错误提示
sub error()
response.write "<br><table cellpadding=0 cellspacing=0 border=0 width="&TableWidth&" bgcolor="&tablebackcolor&" align=center>"&_
		"<tr><td><table cellpadding=3 cellspacing=1 border=0 width=""100%""><tr align=center>"&_
		"<td width=""100%"" bgcolor="&tabletitlecolor&"><font color="&TableFontColor&"><b>论坛错误信息</b></font></td>"&_
		"</tr><tr><td width=""100%"" bgcolor="&tablebodycolor&">"&_
		"<font color="&TableContentColor&"><b>产生错误的可能原因：</b><br><br>"&_
		"<li>您是否仔细阅读了<a href=boardhelp.asp?boardid="&boardid&"><font color="&TableContentColor&">帮助文件</font></a>"&_
		""&errmsg&"</font></td></tr><tr align=center><td width=""100%"" bgcolor="&tabletitlecolor&">"&_
		"<a href=javascript:history.go(-1)><font color="&TableFontColor&"> << 返回上一页</font></a>"&_
		"</td></tr>   </table>   </td></tr></table>"
end sub

function isChinese(para)
       on error resume next
       dim str
       dim i
       if isNUll(para) then 
          isChinese=false
          exit function
       end if
       str=cstr(para)
       if trim(str)="" then
          isChinese=false
          exit function
       end if
       for i=1 to len(str)
		   c=asc(mid(str,i,1))
             if c>=0 then 
			 isChinese=false 
              exit function
           end if
       next
       isChinese=true
       if err.number<>0 then err.clear
   end function

rem 过滤字符
function ChkBadWords(fString)
    dim bwords
    if not(isnull(BadWords) or isnull(fString)) then
    bwords = split(BadWords, "|")
    for i = 0 to ubound(bwords)
        fString = Replace(fString, bwords(i), string(len(bwords(i)),"*")) 
    next
    ChkBadWords = fString
    end if
end function

Rem 过滤HTML代码
function HTMLEncode(fString)
if not isnull(fString) then
    fString = replace(fString, ">", "&gt;")
    fString = replace(fString, "<", "&lt;")

    fString = Replace(fString, CHR(32), "&nbsp;")
    fString = Replace(fString, CHR(9), "&nbsp;")
    fString = Replace(fString, CHR(34), "&quot;")
    fString = Replace(fString, CHR(39), "&#39;")
    fString = Replace(fString, CHR(13), "")
    'fString = Replace(fString, CHR(10) & CHR(10), "</P><P> ")
    fString = Replace(fString, CHR(10), "<BR> ")

    fString=ChkBadWords(fString)
    HTMLEncode = fString
end if
end function

Rem 过滤表单字符
function HTMLcode(fString)
if not isnull(fString) then
    fString = Replace(fString, CHR(13), "")
    fString = Replace(fString, CHR(10) & CHR(10), "</P><P>")
    fString = Replace(fString, CHR(10), "<BR>")
    HTMLcode = fString
end if
end function

'用户IP限制
function LockIP(sip)
	dim str1,str2,str3,str4
	dim num
	LockIP=false
	if isnumeric(left(sip,2)) then
		str1=left(sip,instr(sip,".")-1)
		sip=mid(sip,instr(sip,".")+1)
		str2=left(sip,instr(sip,".")-1)
		sip=mid(sip,instr(sip,".")+1)
		str3=left(sip,instr(sip,".")-1)
		str4=mid(sip,instr(sip,".")+1)
		if isNumeric(str1)=0 or isNumeric(str2)=0 or isNumeric(str3)=0 or isNumeric(str4)=0 then
	
		else
			num=cint(str1)*256*256*256+cint(str2)*256*256+cint(str3)*256+cint(str4)-1
			sql="select count(*) from LockIP where ip1 <="&num&" and ip2 >="&num&""
			set rs=server.createobject("adodb.recordset")
			rs.open sql,conn,1,1
			if rs(0)>0 then 
				LockIP=true
			end if
			rs.close
			set rs=nothing
		end if
	end if
end function

Rem 判断发言是否来自外部
function ChkPost()
	dim server_v1,server_v2
	chkpost=false
	server_v1=Cstr(Request.ServerVariables("HTTP_REFERER"))
	server_v2=Cstr(Request.ServerVariables("SERVER_NAME"))
	if mid(server_v1,8,len(server_v2))<>server_v2 then
		chkpost=false
	else
		chkpost=true
	end if
end function

Rem 过滤SQL非法字符
function checkStr(str)
	if isnull(str) then
		checkStr = ""
		exit function 
	end if
	checkStr=replace(str,"'","''")
end function

Rem 判断用户来源
function address(sip)
	dim str1,str2,str3,str4
	dim num
	dim country,city
	dim irs
	if isnumeric(left(sip,2)) then
	if sip="127.0.0.1" then sip="192.168.0.1"
	str1=left(sip,instr(sip,".")-1)
	sip=mid(sip,instr(sip,".")+1)
	str2=left(sip,instr(sip,".")-1)
	sip=mid(sip,instr(sip,".")+1)
	str3=left(sip,instr(sip,".")-1)
	str4=mid(sip,instr(sip,".")+1)
	if isNumeric(str1)=0 or isNumeric(str2)=0 or isNumeric(str3)=0 or isNumeric(str4)=0 then

	else
		num=cint(str1)*256*256*256+cint(str2)*256*256+cint(str3)*256+cint(str4)-1
		sql="select Top 1 country,city from address where ip1 <="&num&" and ip2 >="&num&""
		set irs=server.createobject("adodb.recordset")
		irs.open sql,conn,1,1
		if irs.eof and irs.bof then 
		country="亚洲"
		city=""
		else
		country=irs(0)
		city=irs(1)
		end if
		irs.close
		set irs=nothing
	end if
	address=country&city
	else
	address="未知"
	end if
end function
%>
 