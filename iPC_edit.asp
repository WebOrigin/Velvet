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
<!-- #INCLUDE file="fckeditor/fckeditor.asp" -->
<%
Dim Recordset1__MMColParam
Recordset1__MMColParam = "1"
If (Request.QueryString("ID") <> "") Then 
  Recordset1__MMColParam = Request.QueryString("ID")
End If
%>
<%
Dim Recordset1
Dim Recordset1_cmd
Dim Recordset1_numRows

Set Recordset1_cmd = Server.CreateObject ("ADODB.Command")
Recordset1_cmd.ActiveConnection = MM_cms_conn_a_STRING
Recordset1_cmd.CommandText = "SELECT * FROM Cont WHERE Text_ID = ?" 
Recordset1_cmd.Prepared = true
Recordset1_cmd.Parameters.Append Recordset1_cmd.CreateParameter("param1", 200, 1, 255, Recordset1__MMColParam) ' adVarChar

Set Recordset1 = Recordset1_cmd.Execute
Recordset1_numRows = 0
%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
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

<script type="text/javascript">
<!--
function MM_validateForm() { //v4.0
  if (document.getElementById){
    var i,p,q,nm,test,num,min,max,errors='',args=MM_validateForm.arguments;
    for (i=0; i<(args.length-2); i+=3) { test=args[i+2]; val=document.getElementById(args[i]);
      if (val) { nm=val.name; if ((val=val.value)!="") {
        if (test.indexOf('isEmail')!=-1) { p=val.indexOf('@');
          if (p<1 || p==(val.length-1)) errors+='- '+nm+' must contain an e-mail address.\n';
        } else if (test!='R') { num = parseFloat(val);
          if (isNaN(val)) errors+='- '+nm+' must contain a number.\n';
          if (test.indexOf('inRange') != -1) { p=test.indexOf(':');
            min=test.substring(8,p); max=test.substring(p+1);
            if (num<min || max<num) errors+='- '+nm+' must contain a number between '+min+' and '+max+'.\n';
      } } } else if (test.charAt(0) == 'R') errors += '- '+nm+' is required.\n'; }
    } if (errors) alert('The following error(s) occurred:\n'+errors);
    document.MM_returnValue = (errors == '');
} }
//-->
</script>
</head>

<body>


<table width="100%" height="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td align="center" valign="middle"><table width="780" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
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
            <td width="15" height="15" background="iPC_element/imsbut_03.png"></td>
            <td background="iPC_element/imsbut_04.png"></td>
            <td width="15" height="15" background="iPC_element/imsbut_06.png"></td>
          </tr>
          <tr>
            <td background="iPC_element/imsbut_08.png">&nbsp;</td>
            <td align="center" bgcolor="#FFFFFF"><form action="iPC_exec.asp" method="post" name="form1" id="form1" onsubmit="MM_validateForm('Title_A','','R');return document.MM_returnValue">
              <label> </label>
              <table width="700" border="0" cellspacing="0" cellpadding="0">
                <tr>
                  <td height="15"></td>
                </tr>
                <tr>
                  <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
                    <tr>
                      <td width="100" class="STYLE_Title">PAGE TITLE:</td>
                      <td><input name="Title_A" type="text" id="Title_A" value="<%=(Recordset1.Fields.Item("Text_Title_Cont").Value)%>" size="49" /></td>
                    </tr>
                  </table></td>
                </tr>
                <tr>
                  <td height="15"><input name="id" type="hidden" id="id" value="<%=Request.QueryString("ID")%>" /></td>
                </tr>
                <tr>
                  <td class="STYLE_Title">PAGE CONTENT:</td>
                </tr>
                <tr>
                  <td height="7" class="STYLE_Title"></td>
                </tr>
                <tr>
                  <td align="center"><%
Dim oFCKeditor
Set oFCKeditor = New FCKeditor
oFCKeditor.BasePath = "fckeditor/"
oFCKeditor.ToolbarSet = "Leask" 
oFCKeditor.Width = "700" 
oFCKeditor.Height = "500" 
oFCKeditor.Value =Recordset1.Fields.Item("Text_Cont").Value
oFCKeditor.Create "editext" 
%></td>
                </tr>
                <tr>
                  <td>&nbsp;</td>
                </tr>
                <tr>
                  <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
                    <tr>
                      <td>&nbsp;</td>
                      <td width="90"><input type="image" src="iPC_element/imsbut_30.png" width="72" height="25" align="absmiddle" border="0" name="Submit" onclick="this.form.submit()" /></td>
                      <td width="90"><input type="image" src="iPC_element/imsbut_35.png" width="72" height="25" align="absmiddle" border="0" name="Reset" onclick="window.location.reload()" /></td>
                      <td width="90"><a href="iPC_admin_main.asp"><img src="iPC_element/imsbut_09.png" width="72" height="25" border="0" /></a></td>
                    </tr>
                  </table></td>
                </tr>
              </table>
              </form></td>
            <td background="iPC_element/imsbut_10.png">&nbsp;</td>
          </tr>
          <tr>
            <td width="15" height="15" background="iPC_element/imsbut_43.png"></td>
            <td background="iPC_element/imsbut_44.png"></td>
            <td width="15" height="15" background="iPC_element/imsbut_45.png"></td>
          </tr>
        </table></td>
      </tr>
      <tr>
        <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td width="15">&nbsp;</td>
            <td><span class="COPYTX">Copyright &copy; 2008 iPAGE New Zealand. All right reserved.</span></td>
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
