<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="Connections/cms_conn_a.asp" -->
<%
dim Text_Err
select case Request.QueryString("bad")
	case 1
		Text_Err="Username and password do not match!"
	case 2
		Text_Err="You must login to view the page!"		
	case else
		Text_Err=""
end select
%>
<%
' *** Validate request to log in to this site.
MM_LoginAction = Request.ServerVariables("URL")
If Request.QueryString <> "" Then MM_LoginAction = MM_LoginAction + "?" + Server.HTMLEncode(Request.QueryString)
MM_valUsername = CStr(Request.Form("user"))
If MM_valUsername <> "" Then
  Dim MM_fldUserAuthorization
  Dim MM_redirectLoginSuccess
  Dim MM_redirectLoginFailed
  Dim MM_loginSQL
  Dim MM_rsUser
  Dim MM_rsUser_cmd
  
  MM_fldUserAuthorization = ""
  MM_redirectLoginSuccess = "iPC_admin_main.asp"
  MM_redirectLoginFailed = "iPC_login.asp?bad=1"

  MM_loginSQL = "SELECT Admin_ID, Admin_Code"
  If MM_fldUserAuthorization <> "" Then MM_loginSQL = MM_loginSQL & "," & MM_fldUserAuthorization
  MM_loginSQL = MM_loginSQL & " FROM [Admin] WHERE Admin_ID = ? AND Admin_Code = ?"
  Set MM_rsUser_cmd = Server.CreateObject ("ADODB.Command")
  MM_rsUser_cmd.ActiveConnection = MM_cms_conn_a_STRING
  MM_rsUser_cmd.CommandText = MM_loginSQL
  MM_rsUser_cmd.Parameters.Append MM_rsUser_cmd.CreateParameter("param1", 200, 1, 255, MM_valUsername) ' adVarChar
  MM_rsUser_cmd.Parameters.Append MM_rsUser_cmd.CreateParameter("param2", 200, 1, 255, Request.Form("pass")) ' adVarChar
  MM_rsUser_cmd.Prepared = true
  Set MM_rsUser = MM_rsUser_cmd.Execute

  If Not MM_rsUser.EOF Or Not MM_rsUser.BOF Then 
    ' username and password match - this is a valid user
    Session("MM_Username") = MM_valUsername
    If (MM_fldUserAuthorization <> "") Then
      Session("MM_UserAuthorization") = CStr(MM_rsUser.Fields.Item(MM_fldUserAuthorization).Value)
    Else
      Session("MM_UserAuthorization") = ""
    End If
    if CStr(Request.QueryString("accessdenied")) <> "" And false Then
      MM_redirectLoginSuccess = Request.QueryString("accessdenied")
    End If
    MM_rsUser.Close
    Response.Redirect(MM_redirectLoginSuccess)
  End If
  MM_rsUser.Close
  Response.Redirect(MM_redirectLoginFailed)
End If
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>iPAGE CMS</title>
<script language="javascript">
var setGradient = (function(){
    //private variables;
    var p_dCanvas = document.createElement('canvas');
    var p_useCanvas =  (typeof(p_dCanvas.getContext) == 'function');
    //
    var p_dCtx = p_useCanvas?p_dCanvas.getContext('2d'):null;
    //cc_on
    var p_isIE = /*@cc_on!@*/false;   
    //test if toDataURL() is supported by Canvas since Safari may not support it
    try{ 
        p_dCtx.canvas.toDataURL();
    }catch(err){
        p_useCanvas = false;
    };
    
    if(p_useCanvas){
        return function (dEl , sColor1 , sColor2 , bRepeatY ){
            if(typeof(dEl) == 'string') 
                dEl =  document.getElementById(dEl);
            if(!dEl) 
                return false;
            var nW = dEl.offsetWidth;
            var nH = dEl.offsetHeight;
            p_dCanvas.width = nW;
            p_dCanvas.height = nH;
            
            var dGradient;
            var sRepeat;
            // Create gradients
            if(bRepeatY){
                dGradient = p_dCtx.createLinearGradient(0,0,nW,0);
                sRepeat = 'repeat-y';
            }else{
                dGradient = p_dCtx.createLinearGradient(0,0,0,nH);
                sRepeat = 'repeat-x';
            }  
    
            dGradient.addColorStop(0,sColor1);
            dGradient.addColorStop(1,sColor2);    
            
            p_dCtx.fillStyle = dGradient;
            p_dCtx.fillRect(0,0,nW,nH);
            var sDataUrl = p_dCtx.canvas.toDataURL('image/png');
            
            with(dEl.style){
                backgroundRepeat = sRepeat;
                backgroundImage = 'url(' + sDataUrl + ')';
                backgroundColor = sColor2;   
            };
        };
    }else if(p_isIE){
        p_dCanvas = p_useCanvas = p_dCtx =  null;  
        return function (dEl , sColor1 , sColor2 , bRepeatY){
            if(typeof(dEl) == 'string') 
                dEl =  document.getElementById(dEl);
            if(!dEl) 
                return false;
            dEl.style.zoom = 1;
            var sF = dEl.currentStyle.filter;
            dEl.style.filter+=' '+['progid:DXImageTransform.Microsoft.gradient(GradientType=',bRepeatY,',enabled=true,startColorstr=',sColor1,',endColorstr=',sColor2,')'].join('');
        };
    }else{
        p_dCanvas = p_useCanvas = p_dCtx =  null;
            return function(dEl, sColor1, sColor2){
                if(typeof(dEl) == 'string') dEl =  document.getElementById(dEl);
                if(!dEl) return false;
                with(dEl.style){
                    backgroundColor = sColor2;
                };
                //alert('your browser does not support gradient effet');
            }
        }
    }
)();
 
</script>

<style type="text/css">
div,html,body {
	margin: 0;
	padding: 0;
	width:100%;
	height:100%;
}
img {
	behavior: url(iPC_element/iepngfix.htc)
}
.inp_sub {
	behavior: url(iPC_element/iepngfix.htc)
}
.INPUT_TA{
	font-family: Arial, Helvetica, sans-serif;
	font-size: 18px;
	color:#363636;
}
.cpr_text {
	font-family: Arial, Helvetica, sans-serif;
	font-size: 11px;
	color: #1b1b1b;
}
</style>
</head>

<body>






<table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0" id="example_a" class="example_a">
  <tr>
    <td><table width="300" border="0" align="center" cellpadding="0" cellspacing="0">
      <tr>
        <td height="400"><table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td height="100">&nbsp;</td>
          </tr>
          <tr>
            <td align="center"><a href="http://www.ipagenz.co.nz/"><img src="iPC_element/ipage_cms_logo_NEW.png" width="272" height="47" border="0"></a></td>
          </tr>
          <tr>
            <td height="50" align="center" valign="middle"><font size="3" face="Arial, Helvetica, sans-serif" color="red"><%=Text_Err%></font></td>
          </tr>
          <tr>
            <td><form name="form1" method="POST" action="<%=MM_LoginAction%>">
              <table width="100%" border="0" cellspacing="0" cellpadding="0">
                <tr>
                  <td><div align="right" class="INPUT_TA">Username:&nbsp;&nbsp;</div></td>
                  <td height="30"><input name="user" type="text" id="user" size="18" /></td>
                </tr>
                <tr>
                  <td><div align="right" class="INPUT_TA">Password:&nbsp;&nbsp;</div></td>
                  <td height="30"><input name="pass" type="password" id="pass" size="20" /></td>
                </tr>
                <tr>
                  <td height="50" colspan="2" align="center" valign="bottom"><img onclick="document.form1.submit()" src="iPC_element/CMS-Login.png" width="72" height="25" />
                    </td>
                  </tr>
              </table>
            </form></td>
          </tr>
          <tr>
            <td height="120">&nbsp;</td>
          </tr>
          <tr>
            <td><div align="center"><span class="cpr_text">Copyright &copy; 2008 iPAGE New Zealand. All rights reserved.</span></div></td>
          </tr>
          <tr>
            <td>&nbsp;</td>
          </tr>
        </table></td>
      </tr>
    </table></td>
  </tr>
</table>

<script language="javascript">
setGradient('example_a','#ffffff','#b3c7ef',0);
</script>








</body>
</html>
