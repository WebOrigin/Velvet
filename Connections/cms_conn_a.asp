<%
' FileName="Connection_ado_conn_string.htm"
' Type="ADO" 
' DesigntimeType="ADO"
' HTTP="true"
' Catalog=""
' Schema=""
Dim db
db="Library/site_db.mdb"
Dim MM_cms_conn_a_STRING
MM_cms_conn_a_STRING = "Provider = Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(""&db&"")
%>
