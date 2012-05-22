<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="Connections/cms_conn_a.asp" -->
<%
Dim Recordset1__MMColParam
Recordset1__MMColParam = "9"
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
Recordset1_cmd.CommandText = "SELECT * FROM Cont WHERE Text_ID = ?" 
Recordset1_cmd.Prepared = true
Recordset1_cmd.Parameters.Append Recordset1_cmd.CreateParameter("param1", 200, 1, 255, Recordset1__MMColParam) ' adVarChar

Set Recordset1 = Recordset1_cmd.Execute
Recordset1_numRows = 0
%><!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<meta name="description" content="An event management and marketing company based in Auckland specialising in special events management, conference, branding, product launch and road show." />
<meta name="keywords" content="events management, event service, event and marketing, event project management, road show, conference, special event and sponsorship, product launch, branding solution, wella yps, velvet, Claire Johnston, event manager" />
<title>Velvet Event Management and Marketing - <%=(Recordset1.Fields.Item("Text_Title_Cont").Value)%></title>

<script src="Scripts/AC_RunActiveContent.js" type="text/javascript"></script>
<link href="style.css" rel="stylesheet" type="text/css" />
</head>

<body>
<table width="779" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td><table width="779" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td height="93" align="left" valign="top" nowrap="nowrap"><table width="779" border="0" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
          <tr>
            <td width="193" height="93" valign="top" class="Logo_St"><a href="index.html"><img src="element/logo.png" alt="" width="193" height="83" border="0" /></a></td>
            <td width="586" height="93"><script type="text/javascript">
AC_FL_RunContent( 'codebase','http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=9,0,28,0','width','586','height','93','src','element/top_ban','quality','high','pluginspage','http://www.adobe.com/shockwave/download/download.cgi?P1_Prod_Version=ShockwaveFlash','movie','element/top_ban' ); //end AC code
</script><noscript><object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=9,0,28,0" width="586" height="93">
              <param name="movie" value="element/top_ban.swf" />
              <param name="quality" value="high" />
              <embed src="element/top_ban.swf" quality="high" pluginspage="http://www.adobe.com/shockwave/download/download.cgi?P1_Prod_Version=ShockwaveFlash" type="application/x-shockwave-flash" width="586" height="93"></embed>
            </object></noscript></td>
          </tr>
        </table></td>
      </tr>
      <tr>
        <td height="27"><table width="779" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td height="27"><script type="text/javascript">
AC_FL_RunContent( 'codebase','http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=9,0,28,0','width','779','height','27','src','element/menu','quality','high','pluginspage','http://www.adobe.com/shockwave/download/download.cgi?P1_Prod_Version=ShockwaveFlash','movie','element/menu' ); //end AC code
</script><noscript><object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=9,0,28,0" width="779" height="27">
              <param name="movie" value="element/menu.swf" />
              <param name="quality" value="high" />
              <embed src="element/menu.swf" quality="high" pluginspage="http://www.adobe.com/shockwave/download/download.cgi?P1_Prod_Version=ShockwaveFlash" type="application/x-shockwave-flash" width="779" height="27"></embed>
            </object></noscript></td>
            </tr>
        </table></td>
      </tr>
      <tr>
        <td><table width="779" border="0" cellpadding="0"  cellspacing="0" background="element/Velvet-all-pages_02.png" class="table_fm">
          <tr><td height="424" valign="top"><table width="100%" border="0" cellpadding="0" cellspacing="0"><tr valign="top"><td><table width="779" border="0" cellpadding="0"  cellspacing="0" background="element/Velvet-all-pages_02.png" class="table_fm">
          <tr>
            <td height="424" valign="top"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                <tr valign="top">
                  <td width="35" rowspan="3" valign="top">&nbsp;</td>
                  <td colspan="2">&nbsp;</td>
                </tr>
                <tr valign="top">
                  <td colspan="2"><%=(Recordset1.Fields.Item("Text_Cont").Value)%></td>
                </tr>
                
                <tr valign="top">
                  <td colspan="2">&nbsp;</td>
                </tr>
            </table></td>
          </tr>
        </table></td>
        </tr>
        </table></td>
          </tr>
        </table></td>
      </tr>
      <tr>
        <td height="22"><table width="779" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td width="35" height="22" background="element/buttom-bar.png">&nbsp;</td>
            <td width="626" background="element/buttom-bar.png" class="STYLE2">© 2008 Velvet Event &amp; Marketing Ltd. All rights reserved.</td>
            <td width="50" background="element/buttom-bar.png"></td>
            <td width="50" background="element/buttom-bar.png"></td>
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
