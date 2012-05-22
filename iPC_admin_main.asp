<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="Connections/cms_conn_a.asp" -->
<%
' *** Restrict Access To Page: Grant or deny access to this page
MM_authorizedUsers=""
MM_authFailedURL="iPC_login.asp?bad=2"
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
Dim Recordset1__MMColParam
Recordset1__MMColParam = "1"
If (Request("MM_EmptyValue") <> "") Then 
  Recordset1__MMColParam = Request("MM_EmptyValue")
End If
%>
<%
Dim Recordset1
Dim Recordset1_cmd
Dim Recordset1_numRows

Set Recordset1_cmd = Server.CreateObject ("ADODB.Command")
Recordset1_cmd.ActiveConnection = MM_cms_conn_a_STRING
Recordset1_cmd.CommandText = "SELECT Text_Title_Cont FROM Cont WHERE Text_ID = ?" 
Recordset1_cmd.Prepared = true
Recordset1_cmd.Parameters.Append Recordset1_cmd.CreateParameter("param1", 200, 1, 255, Recordset1__MMColParam) ' adVarChar

Set Recordset1 = Recordset1_cmd.Execute
Recordset1_numRows = 0
%>
<%
Dim Recordset2__MMColParam
Recordset2__MMColParam = "2"
If (Request("MM_EmptyValue") <> "") Then 
  Recordset2__MMColParam = Request("MM_EmptyValue")
End If
%>
<%
Dim Recordset2
Dim Recordset2_cmd
Dim Recordset2_numRows

Set Recordset2_cmd = Server.CreateObject ("ADODB.Command")
Recordset2_cmd.ActiveConnection = MM_cms_conn_a_STRING
Recordset2_cmd.CommandText = "SELECT Text_Title_Cont FROM Cont WHERE Text_ID = ?" 
Recordset2_cmd.Prepared = true
Recordset2_cmd.Parameters.Append Recordset2_cmd.CreateParameter("param1", 200, 1, 255, Recordset2__MMColParam) ' adVarChar

Set Recordset2 = Recordset2_cmd.Execute
Recordset2_numRows = 0
%>
<%
Dim Recordset3__MMColParam
Recordset3__MMColParam = "3"
If (Request("MM_EmptyValue") <> "") Then 
  Recordset3__MMColParam = Request("MM_EmptyValue")
End If
%>
<%
Dim Recordset3
Dim Recordset3_cmd
Dim Recordset3_numRows

Set Recordset3_cmd = Server.CreateObject ("ADODB.Command")
Recordset3_cmd.ActiveConnection = MM_cms_conn_a_STRING
Recordset3_cmd.CommandText = "SELECT Text_Title_Cont FROM Cont WHERE Text_ID = ?" 
Recordset3_cmd.Prepared = true
Recordset3_cmd.Parameters.Append Recordset3_cmd.CreateParameter("param1", 200, 1, 255, Recordset3__MMColParam) ' adVarChar

Set Recordset3 = Recordset3_cmd.Execute
Recordset3_numRows = 0
%>
<%
Dim Recordset4__MMColParam
Recordset4__MMColParam = "4"
If (Request("MM_EmptyValue") <> "") Then 
  Recordset4__MMColParam = Request("MM_EmptyValue")
End If
%>
<%
Dim Recordset4
Dim Recordset4_cmd
Dim Recordset4_numRows

Set Recordset4_cmd = Server.CreateObject ("ADODB.Command")
Recordset4_cmd.ActiveConnection = MM_cms_conn_a_STRING
Recordset4_cmd.CommandText = "SELECT Text_Title_Cont FROM Cont WHERE Text_ID = ?" 
Recordset4_cmd.Prepared = true
Recordset4_cmd.Parameters.Append Recordset4_cmd.CreateParameter("param1", 200, 1, 255, Recordset4__MMColParam) ' adVarChar

Set Recordset4 = Recordset4_cmd.Execute
Recordset4_numRows = 0
%>
<%
Dim Recordset5__MMColParam
Recordset5__MMColParam = "5"
If (Request("MM_EmptyValue") <> "") Then 
  Recordset5__MMColParam = Request("MM_EmptyValue")
End If
%>
<%
Dim Recordset5
Dim Recordset5_cmd
Dim Recordset5_numRows

Set Recordset5_cmd = Server.CreateObject ("ADODB.Command")
Recordset5_cmd.ActiveConnection = MM_cms_conn_a_STRING
Recordset5_cmd.CommandText = "SELECT Text_Title_Cont FROM Cont WHERE Text_ID = ?" 
Recordset5_cmd.Prepared = true
Recordset5_cmd.Parameters.Append Recordset5_cmd.CreateParameter("param1", 200, 1, 255, Recordset5__MMColParam) ' adVarChar

Set Recordset5 = Recordset5_cmd.Execute
Recordset5_numRows = 0
%>
<%
Dim Recordset6__MMColParam
Recordset6__MMColParam = "6"
If (Request("MM_EmptyValue") <> "") Then 
  Recordset6__MMColParam = Request("MM_EmptyValue")
End If
%>
<%
Dim Recordset6
Dim Recordset6_cmd
Dim Recordset6_numRows

Set Recordset6_cmd = Server.CreateObject ("ADODB.Command")
Recordset6_cmd.ActiveConnection = MM_cms_conn_a_STRING
Recordset6_cmd.CommandText = "SELECT Text_Title_Cont FROM Cont WHERE Text_ID = ?" 
Recordset6_cmd.Prepared = true
Recordset6_cmd.Parameters.Append Recordset6_cmd.CreateParameter("param1", 200, 1, 255, Recordset6__MMColParam) ' adVarChar

Set Recordset6 = Recordset6_cmd.Execute
Recordset6_numRows = 0
%>
<%
Dim Recordset7__MMColParam
Recordset7__MMColParam = "7"
If (Request("MM_EmptyValue") <> "") Then 
  Recordset7__MMColParam = Request("MM_EmptyValue")
End If
%>
<%
Dim Recordset7
Dim Recordset7_cmd
Dim Recordset7_numRows

Set Recordset7_cmd = Server.CreateObject ("ADODB.Command")
Recordset7_cmd.ActiveConnection = MM_cms_conn_a_STRING
Recordset7_cmd.CommandText = "SELECT Text_Title_Cont FROM Cont WHERE Text_ID = ?" 
Recordset7_cmd.Prepared = true
Recordset7_cmd.Parameters.Append Recordset7_cmd.CreateParameter("param1", 200, 1, 255, Recordset7__MMColParam) ' adVarChar

Set Recordset7 = Recordset7_cmd.Execute
Recordset7_numRows = 0
%>
<%
Dim Recordset8__MMColParam
Recordset8__MMColParam = "8"
If (Request("MM_EmptyValue") <> "") Then 
  Recordset8__MMColParam = Request("MM_EmptyValue")
End If
%>
<%
Dim Recordset8
Dim Recordset8_cmd
Dim Recordset8_numRows

Set Recordset8_cmd = Server.CreateObject ("ADODB.Command")
Recordset8_cmd.ActiveConnection = MM_cms_conn_a_STRING
Recordset8_cmd.CommandText = "SELECT Text_Title_Cont FROM Cont WHERE Text_ID = ?" 
Recordset8_cmd.Prepared = true
Recordset8_cmd.Parameters.Append Recordset8_cmd.CreateParameter("param1", 200, 1, 255, Recordset8__MMColParam) ' adVarChar

Set Recordset8 = Recordset8_cmd.Execute
Recordset8_numRows = 0
%>
<%
Dim Recordset9__MMColParam
Recordset9__MMColParam = "9"
If (Request("MM_EmptyValue") <> "") Then 
  Recordset9__MMColParam = Request("MM_EmptyValue")
End If
%>
<%
Dim Recordset9
Dim Recordset9_cmd
Dim Recordset9_numRows

Set Recordset9_cmd = Server.CreateObject ("ADODB.Command")
Recordset9_cmd.ActiveConnection = MM_cms_conn_a_STRING
Recordset9_cmd.CommandText = "SELECT Text_Title_Cont FROM Cont WHERE Text_ID = ?" 
Recordset9_cmd.Prepared = true
Recordset9_cmd.Parameters.Append Recordset9_cmd.CreateParameter("param1", 200, 1, 255, Recordset9__MMColParam) ' adVarChar

Set Recordset9 = Recordset9_cmd.Execute
Recordset9_numRows = 0
%>
<%
Dim Recordset10__MMColParam
Recordset10__MMColParam = "10"
If (Request("MM_EmptyValue") <> "") Then 
  Recordset10__MMColParam = Request("MM_EmptyValue")
End If
%>
<%
Dim Recordset10
Dim Recordset10_cmd
Dim Recordset10_numRows

Set Recordset10_cmd = Server.CreateObject ("ADODB.Command")
Recordset10_cmd.ActiveConnection = MM_cms_conn_a_STRING
Recordset10_cmd.CommandText = "SELECT Text_Title_Cont FROM Cont WHERE Text_ID = ?" 
Recordset10_cmd.Prepared = true
Recordset10_cmd.Parameters.Append Recordset10_cmd.CreateParameter("param1", 200, 1, 255, Recordset10__MMColParam) ' adVarChar

Set Recordset10 = Recordset10_cmd.Execute
Recordset10_numRows = 0
%>
<%
Dim Recordset11__MMColParam
Recordset11__MMColParam = "11"
If (Request("MM_EmptyValue") <> "") Then 
  Recordset11__MMColParam = Request("MM_EmptyValue")
End If
%>
<%
Dim Recordset11
Dim Recordset11_cmd
Dim Recordset11_numRows

Set Recordset11_cmd = Server.CreateObject ("ADODB.Command")
Recordset11_cmd.ActiveConnection = MM_cms_conn_a_STRING
Recordset11_cmd.CommandText = "SELECT Text_Title_Cont FROM Cont WHERE Text_ID = ?" 
Recordset11_cmd.Prepared = true
Recordset11_cmd.Parameters.Append Recordset11_cmd.CreateParameter("param1", 200, 1, 255, Recordset11__MMColParam) ' adVarChar

Set Recordset11 = Recordset11_cmd.Execute
Recordset11_numRows = 0
%>
<script language="vbscript">

</script>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>iPAGE CMS</title>
<style type="text/css">
<!--
body {
	background-color: #f5f6f7;
}
div,html,body {
	margin: 0;
	padding: 0;
	width:100%;
	height:100%;
}
.STYLE_List_Title {font-family: Arial, Helvetica, sans-serif; font-size: 12px; color: #666666; }
.STYLE_List_Name {font-size: 12px; color: #333333; font-family: Arial, Helvetica, sans-serif;}
.STYLE_Title {font-size: 12px; color: #000000; font-family: Arial, Helvetica, sans-serif; font-weight: bold; }
.COPYTX {
	font-family: Arial, Helvetica, sans-serif;
	font-size: 10px;
	color: #333333;
}

-->
</style>
</head>

<body>
<table width="100%" height="100%" border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td align="center" valign="middle"><table width="780" border="0" cellspacing="0" cellpadding="0">
            <tr>
              <td width="100%"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                <tr>
                  <td colspan="4">&nbsp;</td>
                </tr>
                <tr>
                  <td width="15">&nbsp;</td>
                  <td width="177"><a href="http://www.ipagenz.co.nz/"><img src="iPC_element/ipc_lpgo.png" alt="" width="177" height="30" border="0" /></a></td>
                  <td>&nbsp;</td>
                  <td width="15">&nbsp;</td>
                </tr>
              </table></td>
            </tr>
            <tr>
              <td height="15"></td>
            </tr>
            <tr>
              <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
                <tr>
                  <td width="15" height="15" background="iPC_element/imsbut_48.png"></td>
                  <td background="iPC_element/imsbut_49.png"></td>
                  <td width="15" height="15" background="iPC_element/imsbut_50.png"></td>
                </tr>
                <tr>
                  <td background="iPC_element/imsbut_51.png"></td>
                  <td bgcolor="#FFFFFF">
                  <table width="720" border="0" align="center" cellpadding="0" cellspacing="0">
                    <tr>
                      <td height="35" valign="middle" class="STYLE_Title">PAGE NAME</td>
                      <td valign="middle" class="STYLE_Title">PAGE TITLE</td>
                      <td colspan="2" valign="middle" class="STYLE_Title">ACTION</td>
                    </tr>
                    <tr>
                      <td width="270" height="30"><span class="STYLE_List_Name">Velvet - a definition</span></td>
                      <td width="270"><span class="STYLE_List_Title"><%=(Recordset1.Fields.Item("Text_Title_Cont").Value)%></span></td>
                      <td width="90"><a href="velvetad.asp" target="_blank"><img src="iPC_element/imsbut_28.png" width="72" height="25" border="0" /></a></td>
                      <td width="90"><a href="iPC_edit.asp?id=1"><img src="iPC_element/imsbut_34.png" width="72" height="25" border="0" /></a></td>
                    </tr>
                    <tr>
                      <td height="30"><span class="STYLE_List_Name">Services</span></td>
                      <td><span class="STYLE_List_Title"><%=(Recordset2.Fields.Item("Text_Title_Cont").Value)%></span></td>
                      <td><a href="services.asp" target="_blank"><img src="iPC_element/imsbut_28.png" alt="" width="72" height="25" border="0" /></a></td>
                      <td><a href="iPC_edit.asp?id=2"><img src="iPC_element/imsbut_34.png" alt="" width="72" height="25" border="0" /></a></td>
                    </tr>
                    <tr>
                      <td height="30"><span class="STYLE_List_Name">Services - Roadshow</span></td>
                      <td><span class="STYLE_List_Title"><%=(Recordset3.Fields.Item("Text_Title_Cont").Value)%></span></td>
                      <td><a href="services_R.asp" target="_blank"><img src="iPC_element/imsbut_28.png" alt="" width="72" height="25" border="0" /></a></td>
                      <td><a href="iPC_edit.asp?id=3"><img src="iPC_element/imsbut_34.png" alt="" width="72" height="25" border="0" /></a></td>
                    </tr>
                    <tr>
                      <td height="30"><span class="STYLE_List_Name">Services - Special Events &amp; Sponsorship</span></td>
                      <td><span class="STYLE_List_Title"><%=(Recordset4.Fields.Item("Text_Title_Cont").Value)%></span></td>
                      <td><a href="services_S.asp" target="_blank"><img src="iPC_element/imsbut_28.png" alt="" width="72" height="25" border="0" /></a></td>
                      <td><a href="iPC_edit.asp?id=4"><img src="iPC_element/imsbut_34.png" alt="" width="72" height="25" border="0" /></a></td>
                    </tr>
                    <tr>
                      <td height="30"><span class="STYLE_List_Name">Services - Branding Solution</span></td>
                      <td><span class="STYLE_List_Title"><%=(Recordset5.Fields.Item("Text_Title_Cont").Value)%></span></td>
                      <td><a href="services_B.asp" target="_blank"><img src="iPC_element/imsbut_28.png" alt="" width="72" height="25" border="0" /></a></td>
                      <td><a href="iPC_edit.asp?id=5"><img src="iPC_element/imsbut_34.png" alt="" width="72" height="25" border="0" /></a></td>
                    </tr>
                    <tr>
                      <td height="30"><span class="STYLE_List_Name">Services - Product Launch</span></td>
                      <td><span class="STYLE_List_Title"><%=(Recordset6.Fields.Item("Text_Title_Cont").Value)%></span></td>
                      <td><a href="services_P.asp" target="_blank"><img src="iPC_element/imsbut_28.png" alt="" width="72" height="25" border="0" /></a></td>
                      <td><a href="iPC_edit.asp?id=6"><img src="iPC_element/imsbut_34.png" alt="" width="72" height="25" border="0" /></a></td>
                    </tr>
                    <tr>
                      <td height="30"><span class="STYLE_List_Name">Services - Conference</span></td>
                      <td><span class="STYLE_List_Title"><%=(Recordset7.Fields.Item("Text_Title_Cont").Value)%></span></td>
                      <td><a href="services_C.asp" target="_blank"><img src="iPC_element/imsbut_28.png" alt="" width="72" height="25" border="0" /></a></td>
                      <td><a href="iPC_edit.asp?id=7"><img src="iPC_element/imsbut_34.png" alt="" width="72" height="25" border="0" /></a></td>
                    </tr>
                    <tr>
                      <td height="30"><span class="STYLE_List_Name">Behind Velvet</span></td>
                      <td><span class="STYLE_List_Title"><%=(Recordset8.Fields.Item("Text_Title_Cont").Value)%></span></td>
                      <td><a href="behind.asp" target="_blank"><img src="iPC_element/imsbut_28.png" alt="" width="72" height="25" border="0" /></a></td>
                      <td><a href="iPC_edit.asp?id=8"><img src="iPC_element/imsbut_34.png" alt="" width="72" height="25" border="0" /></a></td>
                    </tr>
                    <tr>
                      <td height="30"><span class="STYLE_List_Name">Testimonial</span></td>
                      <td><span class="STYLE_List_Title"><%=(Recordset9.Fields.Item("Text_Title_Cont").Value)%></span></td>
                      <td><a href="testimonial.asp" target="_blank"><img src="iPC_element/imsbut_28.png" alt="" width="72" height="25" border="0" /></a></td>
                      <td><a href="iPC_edit.asp?id=9"><img src="iPC_element/imsbut_34.png" alt="" width="72" height="25" border="0" /></a></td>
                    </tr>
                    <tr>
                      <td height="30"><span class="STYLE_List_Name">Partnership</span></td>
                      <td><span class="STYLE_List_Title"><%=(Recordset10.Fields.Item("Text_Title_Cont").Value)%></span></td>
                      <td><a href="partnership.asp" target="_blank"><img src="iPC_element/imsbut_28.png" alt="" width="72" height="25" border="0" /></a></td>
                      <td><a href="iPC_edit.asp?id=10"><img src="iPC_element/imsbut_34.png" alt="" width="72" height="25" border="0" /></a></td>
                    </tr>
                    <tr>
                      <td height="30"><span class="STYLE_List_Name">Enquiries</span></td>
                      <td><span class="STYLE_List_Title"><%=(Recordset11.Fields.Item("Text_Title_Cont").Value)%></span></td>
                      <td><a href="enquiries.asp" target="_blank"><img src="iPC_element/imsbut_28.png" alt="" width="72" height="25" border="0" /></a></td>
                      <td><a href="iPC_edit.asp?id=11"><img src="iPC_element/imsbut_34.png" alt="" width="72" height="25" border="0" /></a></td>
                    </tr>
                  </table>                  </td>
                  <td background="iPC_element/imsbut_53.png">&nbsp;</td>
                </tr>
                <tr>
                  <td width="15" height="133" background="iPC_element/imsbut_56.png">&nbsp;</td>
                  <td background="iPC_element/imsbut_57.png">&nbsp;</td>
                  <td background="iPC_element/imsbut_58.png">&nbsp;</td>
                </tr>
              </table></td>
            </tr>
            <tr>
              <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
                <tr>
                  <td width="15">&nbsp;</td>
                  <td><span class="COPYTX">Copyright &copy; 2008 iPAGE New Zealand. All right reserved.</span></td>
                  <td width="150" align="right" class="COPYTX">Version:1.0 Build:08071201</td>
                  <td width="15">&nbsp;</td>
                </tr>
              </table></td>
            </tr>
      </table></td>
    </tr>
  </table>
</body>
</html>
<%
Recordset1.Close()
Set Recordset1 = Nothing
%>
<%
Recordset2.Close()
Set Recordset2 = Nothing
%>
<%
Recordset3.Close()
Set Recordset3 = Nothing
%>
<%
Recordset4.Close()
Set Recordset4 = Nothing
%>
<%
Recordset5.Close()
Set Recordset5 = Nothing
%>
<%
Recordset6.Close()
Set Recordset6 = Nothing
%>
<%
Recordset7.Close()
Set Recordset7 = Nothing
%>
<%
Recordset8.Close()
Set Recordset8 = Nothing
%>
<%
Recordset9.Close()
Set Recordset9 = Nothing
%>
<%
Recordset10.Close()
Set Recordset10 = Nothing
%>
<%
Recordset11.Close()
Set Recordset11 = Nothing
%>
