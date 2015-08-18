<%
if session("AdminName") = "" then
response.redirect "../login.asp"
end if
%>
<html>
<head>
<title>管理员回复</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="guestbook.CSS">
</head>

<body bgcolor="#FFFFFF">
<form name="form1" action="re_answer.asp?id=<%=request.QueryString("id")%>
  " method="post" > 
  <div align="center"> <br>
    <table border="1" cellspacing="0" bordercolordark="#0099FF" bordercolorlight="#CCFFFF" width="90%" cellpadding="5">
      <tr> 
        <td colspan="2" height="50"> 
          <div align="center"> 回 复 内 容 <br>
            <br>
            <textarea name="txtanswer" cols="60" rows="6"></textarea>
            <br>
          </div>
        </td>
      </tr>
      <tr> 
        <td colspan="2" height="50"> 
          <div align="center"> 
            <input type="submit" name="Submit" value="送出留言">
            &nbsp;&nbsp; 
            <input type="reset"  name="Submit1" value="重新填写">
			&nbsp;&nbsp; 
            <input type="button" name="Submit2" value="放弃回复" onclick="javascropt:history.back()">
          </div>
        </td>
      </tr>
    </table>
  </div>
</form>
</body>
</html>
