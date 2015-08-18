<!--#include file="gsconn.asp"-->
<!--#include file="char.asp"-->
<%
if session("AdminName") = "" then
response.redirect "../login.asp"
end if
%>
<%
   dim sql
   dim rs
   dim sanswer
   sanswer=HTMLEncode(request.form("txtanswer"))
saveData()

   sub saveData()
       on error resume next
       dim cmdTemp
       dim InsertCursor
       Set cmdTemp = Server.CreateObject("ADODB.Command")
       Set InsertCursor = Server.CreateObject("ADODB.Recordset")
       cmdTemp.CommandText = "SELECT * FROM book where id="&request.QueryString("id")
       cmdTemp.CommandType = 1
       Set cmdTemp.ActiveConnection = conngs
       InsertCursor.Open cmdTemp, , 1, 3
       InsertCursor("answer") =sanswer
       InsertCursor.Update
       InsertCursor.close
       conn.close
       set InsertCursor=nothing
       set conn=nothing
       rs.close
       set rs=nothing  
   end sub 
response.redirect "manager.asp"
%>
