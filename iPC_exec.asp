<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<% Server.ScriptTimeOut=5000 %>
<!--#include file="Connections/cms_conn_a.asp" -->

<%
' *** Restrict Access To Page: Grant or deny access to this page
MM_authorizedUsers=""
MM_authFailedURL="iPC_login.asp"
MM_grantAccess=false
If Session("MM_Username") <> "" Then
  If (true Or CStr(Session("MM_UserAuthorization"))="") Or _
         (InStr(1,MM_authorizedUsers,Session("MM_UserAuthorization"))>=1) Then
    MM_grantAccess = true
  End If
End If
If Not MM_grantAccess Then
  MM_qsChar = "?"
  If (InStr(1,MM_authFailedURL,"?") >= 1) Then MM_qsChar = "&"
  MM_referrer = Request.ServerVariables("URL")
  if (Len(Request.QueryString()) > 0) Then MM_referrer = MM_referrer & "?" & Request.QueryString()
  MM_authFailedURL = MM_authFailedURL & MM_qsChar & "accessdenied=" & Server.URLEncode(MM_referrer)
  Response.Redirect(MM_authFailedURL)
End If
%>


<%

text_area_a = Request.Form("editext")
id = Request.Form("id")
Title_A = Request.Form("Title_A")

Dim Update_Cn, StrSQL
Set Update_Cn = Server.CreateObject("ADODB.Connection")
Update_Cn.Open MM_cms_conn_a_STRING
StrSQL = "UPDATE Cont SET Text_Cont='" & text_area_a &"', Text_Title_Cont ='" & Title_A & "' WHERE Text_ID='" & id & "'"
Update_Cn.Execute StrSQL
Update_Cn.close
Set Update_Cn = Nothing

response.Redirect "iPC_admin_main.asp"

%>