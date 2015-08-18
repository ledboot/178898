<%
   dim conngs   
   dim connstrgs

   on error resume next
   'connstrgs="DBQ="+server.mappath("../Bob/#Bob_yndYFj8wj6ggps@#%.asa")+";DefaultDir=;DRIVER={Microsoft Access Driver (*.mdb)};"
   connstrgs="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" &Server.MapPath("../Bob/#Bob_1784fdW7tht898@#]%.mdb") & ";Jet OLEDB:Database Password='1788117HYJN898'"
     set conngs=server.createobject("ADODB.CONNECTION")
     if err then 
      err.clear
%>
<script language="javascript">
      top.location="error.htm"
  </script>
<%      
     else
        conngs.open connstrgs
        if err then 
          err.clear
       %>
<script language="javascript">
          top.location="error.htm"
       </script>
<%      
        end if
     end if
%>