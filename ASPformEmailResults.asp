

<%
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'To get the script for work please set the following values:
'Set the credentials for your email account to send the email from
username="craig@craighep.co.uk"     'Insert your email account username between the double quotes            
password="f1f2f3"     'Insert your email account password between the double quotes              

'Set the from and to email addresses
sendFrom = "craig@craighep.co.uk"   'Insert the email address you wish to send from   
sendTo = "crh13@aber.ac.uk"     'Insert the email address to send to in here

'DO NOT CHANGE ANY SCRIPT CODE BELOW THIS LINE.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'This script demonstrates how to send an email using asmtp

'Create a CDO.Configuration object
Set objCdoCfg = Server.CreateObject("CDO.Configuration")

'Configure the settings needed to send an email
objCdoCfg.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") ="intmail.atlas.pipex.net"
objCdoCfg.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
objCdoCfg.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
objCdoCfg.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 0
objCdoCfg.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = username
objCdoCfg.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = password
objCdoCfg.Fields.Update
       	
'Create the email that we are going to send
Set objCdoMessage = Server.CreateObject("CDO.Message")
Set objCdoMessage.Configuration = objCdoCfg
objCdoMessage.From = sendFrom
objCdoMessage.To = sendTo
objCdoMessage.Subject = Request.Form("Subject")

'Add the email body text
objCdoMessage.TextBody = "Sent From: " & Request.Form("Name") & vbCrLf & vbCrLf & Request.Form("Body") & vbCrLf & vbCrLf & "Their Email:" & Request.Form("Email")

	
On Error Resume Next
  	
'Send the email
objCdoMessage.Send

'Check if an exception was thrown		
If Err.Number <> 0 Then
	'Response.Write "<FONT color=""Red"">Error: " & Err.Description & " (" & Err.Number & ")</FONT><br/>"
Else
	Response.Write "<FONT color=""Green"">The email has been sent to " & sendTo & ".</FONT>"
End If
	
'Dispose of the objects after we have used them
Set objCdoMessage = Nothing
Set objCdoCfg = Nothing
Set FSO = nothing
Set TextStream = Nothing
Response.Redirect "http://craighep.co.uk/test/#5"
%>
