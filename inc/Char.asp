<%
'������
Sub Error1(Message,url)
	response.write "<script>alert('"&Message&"');location.href='"&url&"';</script>"
	Response.End
end Sub

Sub Error2(Message)
	response.write "<script>alert('"&Message&"');history.back();</script>"
	Response.End
end Sub

Sub error3(Message,url)
	response.write "<script>alert('"&Message&"');top.location.href='"&Domain&"/"&url&"';</script>"
	Response.End
end Sub

Sub error4(Message)
	response.write "<script>alert('"&Message&"');window.close();</script>"
	Response.End
end Sub

Sub error5(Message,url)
	response.write "<script>alert('"&Message&"');window.close();opener.location.href='"&Domain&"/"&url&"';</script>"
	Response.End
end Sub

Sub Error6(Message)
	response.write "<script>alert('"&Message&"');</script>"
	Response.End
end Sub

function selected(req,reqvalue)
	if req=reqvalue then
		selected=" selected"
	else
		selected=""
	end if
end function
function disabled(req,reqvalue)
	if req=reqvalue then
		disabled=" disabled"
	else
		disabled=""
	end if
end function
function checked0(req,reqvalue)
	if req=reqvalue then
		checked0=" checked"
	else
		checked0=""
	end if
end function

'���������
Function get_rand(Z_num)
  dim num1
  dim rndnum
  Randomize
  Do While Len(rndnum)<Z_num
  num1=CStr(Chr((57-48)*rnd+48))
  rndnum=rndnum&num1
  loop
  get_rand=rndnum
End Function

'����ַ���
Public Function Checkstr(Str)
	If Isnull(Str) Then
		CheckStr = ""
		Exit Function 
	End If
	Str = Replace(Str,Chr(0),"")
	CheckStr = Replace(Replace(Replace(Trim(str), "'", ""), Chr(34), ""), ";", "")
	'CheckStr = Replace(Str,"'","''")
End Function

'����ַ������ֵĴ���
Function CheckTheChar(TheChar,TheString)
'TheChar="Ҫ�����ַ���"
'TheString="�������ַ���"
'CheckTheChar(",","9,8,9,7")
	if inStr(TheString,TheChar) then
		for n =1 to Len(TheString)
			if Mid(TheString,n,Len(TheChar))=TheChar then 
				CheckTheChar=CheckTheChar+1
			End if
		Next
	else
		CheckTheChar="0"
	end if
End Function

'6�������
function mkPassword()
	Dim NewPass
	Dim u_per, l_per, int_Counter
	Randomize
	For int_Counter = 1 To 6
	if (int_Counter mod 2)=1 then
		u_per=122
		l_per=97
	else
		u_per=57
		l_per=48
	end if
	NewPass = NewPass & Chr(Int((u_per - l_per + 1) * Rnd + l_per))
	Next
	mkPassword = NewPass
end function

'�������
function makePassword(byVal maxLen)
	Dim strNewPass
	Dim whatsNext, upper, lower, intCounter
	Randomize
	For intCounter = 1 To maxLen
		whatsNext = Int((1 - 0 + 2) * Rnd + 0)
		If whatsNext = 0 Then
			'��д��ĸ
			upper = 90 
			lower = 65 
		Else
			if whatsNext=1 then
				'Сд
				upper=122
				lower=97
			else
				'����
				upper = 57
				lower = 48
			end if
		End If 
		strNewPass = strNewPass & Chr(Int((upper - lower + 1) * Rnd + lower))
	Next
	makePassword = strNewPass
end function

' �ж��Ƿ�ȫ�ַ���
Function IsSafeStr(str)
	Dim s_BadStr, n, i
'	s_BadStr = "' ��&<>?%,;:()`~!@#$^*{}[]|+-=" & Chr(34) & Chr(9) & Chr(32)
	s_BadStr = "' ��&<>?%,;:()`~!@#$^*{}[]|+=" & Chr(34) & Chr(9) & Chr(32)
	n = Len(s_BadStr)
	IsSafeStr = True
	For i = 1 To n
		If Instr(str, Mid(s_BadStr, i, 1)) > 0 Then
			IsSafeStr = False
			Exit Function
		End If
	Next
End Function

'������
Function IsObjInstalled(strClassString)
	On Error Resume Next
	IsObjInstalled = False
	Err = 0
	Dim xTestObj
	Set xTestObj = Server.CreateObject(strClassString)
	If 0 = Err Then IsObjInstalled = True
	Set xTestObj = Nothing
	Err = 0
End Function

Rem ==========ͨ�ú���=========
Rem �ж������Ƿ�����
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

function strlen(str)
dim p_len
p_len=0
strlen=0
if trim(str)<>"" then
p_len=len(trim(str))
for xx=1 to p_len
if asc(mid(str,xx,1))<0 then
strlen=int(strlen) + 2
else
strlen=int(strlen) + 1
end if
next
end if
end function

'�ж�ָ���ַ������ִ���
Function strShowNum(str, searchStr)
	posTemp = 1
	strShowTotal = 0
	while instr(posTemp,str,searchStr)>0 
    	strShowTotal = strShowTotal + 1
   		posTemp = instr(posTemp,str,searchStr) + 1
	wend
	strShowNum = strShowTotal
End Function

function strvalue(str,lennum)
dim p_num
dim i
'str=UnHTMLEncode(str)
if strlen(str)<=lennum then
strvalue=str
else
p_num=0
x=0
do while not p_num > lennum-2
x=x+1
if asc(mid(str,x,1))<0 then
p_num=int(p_num) + 2
else
p_num=int(p_num) + 1
end if
strvalue=left(trim(str),x)&"��"
loop
end if
end function

rem ����Ƿ�ΪӢ��/����/�»���/����
function iszhuce(para)
       on error resume next
       dim str
       dim i
       if isNUll(para) then 
          iszhuce=false
          exit function
       end if
       str=cstr(para)
       if trim(str)="" then
          iszhuce=false
          exit function
       end if
       for i=1 to len(str)
     b = Lcase(Mid(str, i, 1))
     if InStr("abcdefghijklmnopqrstuvwxyz0123456789_", b) <= 0  then
	 		c=asc(mid(str,i,1))
             if c>=0 then 
			 	iszhuce=false 
              	exit function
             end if
       'iszhuce = false
       'exit function
     end if
       next
       iszhuce=true
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

rem �����ַ�
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

Rem ����HTML����
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
    fString = Replace(fString, CHR(10), "<BR>")

    'fString=ChkBadWords(fString)
    HTMLEncode = fString
end if
end function

function UnHTMLEncode(UnfString)
if not isnull(UnfString) then
    UnfString = replace(UnfString, "&gt;", ">")
    UnfString = replace(UnfString, "&lt;", "<")

    UnfString = Replace(UnfString, "&nbsp;", CHR(32))
    UnfString = Replace(UnfString, "&nbsp;", CHR(9))
    UnfString = Replace(UnfString, "&quot;", CHR(34))
    UnfString = Replace(UnfString, "&#39;", CHR(39))
    UnfString = Replace(UnfString, "", CHR(13))
    'fString = Replace(fString, CHR(10) & CHR(10), "</P><P> ")
    UnfString = Replace(UnfString, "<BR>", CHR(10))

    'fString=ChkBadWords(fString)
    UnHTMLEncode = UnfString
end if
end function


Rem �ж��û��Ƿ�����
Function IFConnected()
	If Response.IsClientConnected Then 
		Response.Flush 
	Else 
		Response.End 
	End If 
End Function

Rem �жϷ����Ƿ������ⲿ
function ChkPost()
	dim server_v1,server_v2
	chkpost=false
	server_v1=Cstr(Request.ServerVariables("HTTP_REFERER"))
	server_v2=Cstr(Request.ServerVariables("SERVER_NAME"))
	if mid(server_v1,8,len(server_v2))<>server_v2 then
		chkpost=false
		Response.Write "<script>alert('�Բ�������Ȩ���ʸ�ҳ����ȷ���Ǵӱ�վ�ύ����Ϣ��');history.back();</script>"
		Response.End
	else
		if instr(request.servervariables("http_referer"),"http://"&request.servervariables("host") )<1 then
			chkpost=false
			'response.write "���� URL ʱ�������ϳ���" 
			Response.Write "<script>alert('�Բ�������Ȩ���ʸ�ҳ����ȷ���Ǵӱ�վ�ύ����Ϣ��');history.back();</script>"
			Response.End
		end if 
		chkpost=true
	end if
end function

' ============================================
' ��ʽ��ʱ��(��ʾ)
' ������n_Flag
'	1:"yyyy-mm-dd hh:mm:ss"
'	2:"yyyy-mm-dd"
'	3:"hh:mm:ss"
'	4:"yyyy��mm��dd��"
'	5:"yyyymmdd"
' ============================================
Function Format_Time(s_Time, n_Flag)
	Dim y, m, d, h, mi, s
	Format_Time = ""
	If IsDate(s_Time) = False Then Exit Function
	y = cstr(year(s_Time))
	m = cstr(month(s_Time))
	If len(m) = 1 Then m = "0" & m
	d = cstr(day(s_Time))
	If len(d) = 1 Then d = "0" & d
	h = cstr(hour(s_Time))
	If len(h) = 1 Then h = "0" & h
	mi = cstr(minute(s_Time))
	If len(mi) = 1 Then mi = "0" & mi
	s = cstr(second(s_Time))
	If len(s) = 1 Then s = "0" & s
	Select Case n_Flag
	Case 1
		' yyyy-mm-dd hh:mm:ss
		Format_Time = y & "-" & m & "-" & d & " " & h & ":" & mi & ":" & s
	Case 2
		' yyyy-mm-dd
		Format_Time = y & "-" & m & "-" & d
	Case 3
		' hh:mm:ss
		Format_Time = h & ":" & mi & ":" & s
	Case 4
		' yyyy��mm��dd��
		Format_Time = y & "��" & m & "��" & d & "��"
	Case 5
		' yyyymmdd
		Format_Time = y & m & d
	Case 6
		' yyyy��mm��dd��
		Format_Time = y & "��" & m & "��" & d & "�� " & h & ":" & mi & ":" & s &" "& WeekdayName(Weekday(s_Time))
	Case 7
		' yyyymmddhhmmss
		Format_Time = y & m & d & h & mi & s
	Case 8
		' yyyy-mm-dd
		Format_Time = m & "-" & d
	End Select
End Function

	rem ��ʾ��ߵ�n���ַ�(�Զ�ʶ����)
	Function LeftTrue(str,n)
		If Not checkIsEmpty(str) Then
			If len(str)<=n/2 Then
				LeftTrue=str
			Else
				Dim TStr
				Dim l,t,c
				Dim i
				l=len(str)
				t=l
				TStr=""
				t=0
				for i=1 to l
					c=asc(mid(str,i,1))
					If c<0 then c=c+65536
						If c>255 then
							t=t+2
						Else
							t=t+1
					End If
					If t>n Then exit for
					TStr=TStr&(mid(str,i,1))
				next
				If str = TStr Then
					LeftTrue = TStr
				Else
					LeftTrue = TStr & "��"
				End If
			End If
		Else
			Exit Function
		End If
	End Function

	Private Function getIP()
		Dim strIPAddr  
		If Request.ServerVariables("HTTP_X_FORWARDED_FOR") = "" OR InStr(Request.ServerVariables("HTTP_X_FORWARDED_FOR"), "unknown") > 0 Then  
			strIPAddr = Request.ServerVariables("REMOTE_ADDR")  
		ElseIf InStr(Request.ServerVariables("HTTP_X_FORWARDED_FOR"), ",") > 0 Then  
			strIPAddr = Mid(Request.ServerVariables("HTTP_X_FORWARDED_FOR"), 1, InStr(Request.ServerVariables("HTTP_X_FORWARDED_FOR"), ",")-1)  
		ElseIf InStr(Request.ServerVariables("HTTP_X_FORWARDED_FOR"), ";") > 0 Then  
			strIPAddr = Mid(Request.ServerVariables("HTTP_X_FORWARDED_FOR"), 1, InStr(Request.ServerVariables("HTTP_X_FORWARDED_FOR"), ";")-1)  
		Else  
			strIPAddr = Request.ServerVariables("HTTP_X_FORWARDED_FOR")  
		End If  
		getIP = Trim(Mid(strIPAddr, 1, 30))  
	End Function

	Function checkIsEmpty(checkStr)
		checkStr = Trim(checkStr)

		If checkStr = "" Or IsNull(checkStr) Or IsEmpty(checkStr) Then
			checkIsEmpty = True
		Else
			checkIsEmpty = False
		End If
	End Function
	
	Function RemoveHTML(strText, tempTAGLIST)
		tempStrText = strText
		Dim TAGLIST
		
		If Not checkIsEmpty(tempTAGLIST) Then
			TAGLIST = tempTAGLIST
		Else
			TAGLIST = ";!--;!DOCTYPE;A;ACRONYM;ADDRESS;APPLET;AREA;B;BASE;BASEFONT;" &_ 
					  "BGSOUND;BIG;BLOCKQUOTE;BODY;BR;BUTTON;CAPTION;CENTER;CITE;CODE;" &_ 
					  "COL;COLGROUP;COMMENT;DD;DEL;DFN;DIR;DIV;DL;DT;EM;EMBED;FIELDSET;" &_ 
					  "FONT;FORM;FRAME;FRAMESET;HEAD;H1;H2;H3;H4;H5;H6;HR;HTML;I;IFRAME;IMG;" &_ 
					  "INPUT;INS;ISINDEX;KBD;LABEL;LAYER;LAGEND;LI;LINK;LISTING;MAP;MARQUEE;" &_ 
					  "MENU;META;NOBR;NOFRAMES;NOSCRIPT;OBJECT;OL;OPTION;P;PARAM;PLAINTEXT;" &_ 
					  "PRE;Q;S;SAMP;SCRIPT;SELECT;SMALL;SPAN;STRIKE;STRONG;STYLE;SUB;SUP;" &_ 
					  "TABLE;TBODY;TD;TEXTAREA;TFOOT;TH;THEAD;TITLE;TR;TT;U;UL;VAR;WBR;XMP;"
		End If
		 
		Const BLOCKTAGLIST = ";APPLET;EMBED;FRAMESET;HEAD;NOFRAMES;NOSCRIPT;OBJECT;SCRIPT;STYLE;"      
		Dim nPos1 
		Dim nPos2 
		Dim nPos3 
		Dim strResult 
		Dim strTagName 
		Dim bRemove 
		Dim bSearchForBlock
		
		nPos1 = InStr(tempStrText, "<") 
		Do While nPos1 > 0 
			nPos2 = InStr(nPos1 + 1, tempStrText, ">") 
			If nPos2 > 0 Then 
				strTagName = Mid(tempStrText, nPos1 + 1, nPos2 - nPos1 - 1) 
				strTagName = Replace(Replace(strTagName, vbCr, " "), vbLf, " ") 
	
				nPos3 = InStr(strTagName, " ") 
				If nPos3 > 0 Then 
					strTagName = Left(strTagName, nPos3 - 1) 
				End If 
				If Left(strTagName, 1) = "/" Then 
					strTagName = Mid(strTagName, 2) 
					bSearchForBlock = False 
				Else 
					bSearchForBlock = True 
				End If 
				 
				If InStr(1, TAGLIST, ";" & strTagName & ";", vbTextCompare) > 0 Then 
					bRemove = True 
					If bSearchForBlock Then 
						If InStr(1, BLOCKTAGLIST, ";" & strTagName & ";", vbTextCompare) > 0 Then 
							nPos2 = Len(tempStrText) 
							nPos3 = InStr(nPos1 + 1, tempStrText, "</" & strTagName, vbTextCompare) 
							If nPos3 > 0 Then 
								nPos3 = InStr(nPos3 + 1, tempStrText, ">") 
							End If 
							 
							If nPos3 > 0 Then 
								nPos2 = nPos3 
							End If 
						End If 
					End If 
				Else 
					bRemove = False 
				End If 
				 
				If bRemove Then 
					strResult = strResult & Left(tempStrText, nPos1 - 1) 
					tempStrText = Mid(tempStrText, nPos2 + 1) 
				Else 
					strResult = strResult & Left(tempStrText, nPos1) 
					tempStrText = Mid(tempStrText, nPos1 + 1) 
				End If 
			Else 
				strResult = strResult & tempStrText 
				tempStrText = "" 
			End If 
			 
			nPos1 = InStr(tempStrText, "<") 
		Loop
		
		strResult = strResult & tempStrText
		
		strResult = Replace(strResult, "&nbsp;", "")
		strResult = Replace(strResult, " ", "")
		strResult = Replace(strResult, vbcrlf, "") 
		 
		RemoveHTML = strResult 
	End Function
	
	function nohtml(str)
		tempStr = str
		tempStr = Replace(tempStr, "<BR>", "")
		tempStr = Replace(tempStr, "</P>", "")
		tempStr = Replace(tempStr, "</P>", "")
		tempStr = Replace(tempStr, "&nbsp;", "")
		tempStr = Replace(tempStr, " ", "")
		tempStr = Replace(tempStr, vbcrlf, "")
		dim re  
		Set re=new RegExp  
		re.IgnoreCase =true  
		re.Global=True  
		re.Pattern="(\<.[^\<]*\>)"  
		tempStr=re.replace(tempStr,"")  
		re.Pattern="(\<\/[^\<]*\>)"  
		tempStr=re.replace(tempStr,"")  
		nohtml=tempStr  
		set re=nothing  
	end function
	
	'-----------------------------------------------------------------------------------------
	'�� �� ����FR_CLng(ByVal strl)
	'��    �ã����ַ�תΪ������ֵ
	'��    ����1��
	'          strl--Ҫת�����ַ���
	'�� �� ֵ�����ָ�����ַ����������򷵻���ֵ�����򷵻�0
	'-----------------------------------------------------------------------------------------
	Function FR_CLng(ByVal str1)
		If str1<>"" and IsNumeric(str1) Then
			FR_CLng = Fix(CDbl(str1))
		Else
			FR_CLng = 0
		End If
	End Function
	
	'-----------------------------------------------------------------------------------------
	'�� �� ����FR_Nohtml(ByVal str)
	'��    �ã������ַ����е�HTML���
	'��    ����1��
	'          str--Ҫ���˵��ַ���
	'�� �� ֵ�����ǵ�HTML��Ǻ���ַ���
	'-----------------------------------------------------------------------------------------
	Function FR_Nohtml(ByVal str)
		Set regEx = New RegExp 
		If IsNull(str) Or Trim(str) = "" Then
			FR_Nohtml = ""
			Exit Function
		End If
		regEx.Pattern = "(\<.[^\<]*\>)"
		str = regEx.Replace(str, "")
		regEx.Pattern = "(\<\/[^\<]*\>)"
	'    str = regEx.Replace(str, "")
		'regEx.Pattern = "\[NextPage(.*?)\]" 
		str = regEx.Replace(str, "")    
		str = Replace(str, "'", "''")
		str = Replace(str, chr(39), "''")
	'    str = Replace(str, Chr(34), "")
	'    str = Replace(str, vbCrLf, "")
	'    str = Trim(str)
		FR_Nohtml = str
		set regEx = nothing
	End Function
	
	'**********************
	'��ע�뺯��
	'**********************
	Function SafeRequest(ParaName,ParaType)
		'--- ������� ---
		'ParaName:��������-�ַ���
		'ParaType:��������-������(1��ʾ���ϲ��������֣�0��ʾ���ϲ���Ϊ�ַ�)
		Dim ParaValue
		ParaValue=Request(ParaName)
		If ParaType=1 then
			ParaValue=FR_Clng(ParaValue)
			If not isNumeric(ParaValue) then
				response.write "����" & ParaName & "����Ϊ�����ͣ�"
				response.end
			End if
		Else
			ParaValue=Trim(FR_Nohtml(replace(ParaValue,"'","''")))
			'ParaValue=Trim(replace(ParaValue,"'","''"))
			if isNull(ParaValue) then
				ParaValue=""
			end if
		End if
		SafeRequest=ParaValue
	End function
	
	Sub setRollingContent(lmNameTemp)
%>
		<table width="100%" border="0" cellspacing="0" cellpadding="0">
			<tr>
				<td height="35" class="tdStyle7">&nbsp;</td>
			</tr>
		</table>
		<table width="100%" border="0" cellspacing="0" cellpadding="0">
			<tr>
				<td width="79" class="tdStyle4">&nbsp;</td>
				<td><div id="MarqueeDiv1" class="divStyle10"><%
					Set rs = rstw(10, lmNameTemp)
					While Not rs.BOF And Not rs.EOF
						titleRs = rs("Title")
						Indeximg1Rs = rs("Indeximg1")
				%><img src="<%=(DoMain & Indeximg1Rs)%>" width="136" height="103" border="0" class="imgStyle3" alt="<%=titleRs%>"><%
						rs.MoveNext
					Wend
					rs.Close
					Set rs = Nothing
				%></div></td>
				<td width="82" class="tdStyle5">&nbsp;</td>
			</tr>
		</table>
		<script type="text/javascript">
			/*
				����˵��
				��һ�� ����DIV��ID
				�ڶ��� ����������� 0�������Ϲ��� 2����������� 3�������ҹ���
				������ �����ٶ� ����Խ�� �ٶ�Խ�� �����ʱÿһ֡�ƶ���Ъʱ����ϵ��� ���ΪС�� ���������л���Ч�� ����Խ�� ����Խ��
				���ĸ� ����DIV�ĳ�ʼ���
				����� ����DIV�ĳ�ʼ�߶�
				������ ����ʱÿһ֡�ƶ���Ъʱ�� ����Խ�� ��Ъʱ��Խ�� ������ٶ���ϵ���
				�ڰ˸� ��ʼ����ǰ�ĵȴ�ʱ��
				�ھŸ� ����DIVÿ�ι����ľ���
			*/
			/*********�����ͣ��������ʼ�ȴ�ʱ�估����϶�***************/
			new Marquee("MarqueeDiv1", 2, 2, 549, 103, 20, 0, 3000, -2)	
		</script>
		<table width="100%" border="0" cellspacing="0" cellpadding="0">
			<tr>
				<td height="35" class="tdStyle7">&nbsp;</td>
			</tr>
		</table>
<%
	End Sub
%>
<%
	Sub setRollingLeft(lmNameTemp, infoTop, NavigationLinkTemp)
%>
		<div id="sidebarDiv2">
<%
		Call setLeftLmHeadStyle(lmNameTemp, NavigationLinkTemp, 1)
%>
			<div class="sidebarDivStyle7"></div>
			<div id="MarqueeDiv" class="sidebarDivStyle8">
				<%
					Set rs = rstw(infoTop, lmNameTemp)
					While Not rs.BOF And Not rs.EOF
						TitleRs = rs("Title")
						IndeximgRs = rs("Indeximg")
						Indeximg1Rs = rs("Indeximg1")
				%>
				<div class="sidebarDivStyle9"></div>
				<div class="sidebarDivStyle10"><img src="<%=(DoMain & Indeximg1Rs)%>" width="178" height="130" alt="<%=TitleRs%>" border="0" /></div>
				<div class="sidebarDivStyle9"></div>
				<%
						rs.MoveNext
					Wend
					rs.Close
					Set rs = Nothing
				%>
			</div>
			<script type="text/javascript">
				/*
					����˵��
					��һ�� ����DIV��ID
					�ڶ��� ����������� 0�������Ϲ��� 2����������� 3�������ҹ���
					������ �����ٶ� ����Խ�� �ٶ�Խ�� �����ʱÿһ֡�ƶ���Ъʱ����ϵ��� ���ΪС�� ���������л���Ч�� ����Խ�� ����Խ��
					���ĸ� ����DIV�ĳ�ʼ���
					����� ����DIV�ĳ�ʼ�߶�
					������ ����ʱÿһ֡�ƶ���Ъʱ�� ����Խ�� ��Ъʱ��Խ�� ������ٶ���ϵ���
					�ڰ˸� ��ʼ����ǰ�ĵȴ�ʱ��
					�ھŸ� ����DIVÿ�ι����ľ���
				*/
				/*********���ϼ�Ъ����(����)***************/
				new Marquee("MarqueeDiv", 0, 0.1, 245, 312, 20, 4000, 5000, 156)		//���ϼ�Ъ����(����)(���˴�ʡ�����һ������52����Ĭ��Ϊ��������)
			</script>
			<div class="sidebarDivStyle11"></div>
		</div>
<%
	End Sub
%>