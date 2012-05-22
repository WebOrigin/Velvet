<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<% Server.ScriptTimeOut=5000 %>
<!--#include FILE="upload.inc"--> 
<%
myemail = "info@velvetevents.co.nz"
myname = "Velvet"
smtpserver = "mail.velvetevents.co.nz"
smtpuser = "webenquiries"
smtppwd = "claire"
smtpemail = "webenquiries@velvetevents.co.nz"

dim upload,file,formName,formPath 
set upload=new upload_5xSoft

cufname = upload.form("First Name")
culname = upload.form("Last Name")
cusmail = upload.form("E-mail")
cusmsg = upload.form("Message")
cusnews = upload.form("newsletter")

set jmail=server.CreateObject ("jmail.message")

jmail.From = "webenquiries@velvetevents.co.nz"
jmail.FromName = cufname & " " & culname 
jmail.ReplyTo = cusmail
jmail.Subject = "Web Customer: " & cufname & " " & culname 

jmail.Body = "Name: " & cufname & " " & culname & vbcrlf & vbcrlf & "E-mail: " & cusmail & vbcrlf & vbcrlf & "Newsletter: " & cusnews & vbcrlf & vbcrlf & "Message:" & vbcrlf & cusmsg

jmail.AddRecipient myemail,myname
jmail.MailServerUserName = smtpuser
jmail.MailServerPassWord = smtppwd

isgo=jmail.Send(smtpserver)
if isgo then
Response.Redirect("http://www.velvetevents.co.nz/respond_success.html")
else
Response.Redirect("http://www.velvetevents.co.nz/respond_fail.html")
end if

jmail.Close
set jmail=nothing
%>