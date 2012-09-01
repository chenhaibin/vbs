' define objects (probably cdonts and fso)

'set constants 

' Variables:
' 	sTo - Person receiving the email
'	oMessage - The email object 
'	oAttachments - The multiple attachments
'	oRecipient - Person recieving the email (uses sTo)
'	fso - FileSystemObject

Dim sTo, servername, ofolder, ofiles
Dim oMessage 


set fso = createobject("scripting.filesystemobject")
set omessage= createobject("CDONTS.NewMail")
sDate = Date
set net = createobject("wscript.network")
servername = net.computername

'General email properties
omessage.from = "testuser@halifax.co.uk"	'added today
omessage.to = "daed@mysystem.org"	'added today
oMessage.Subject = "Badmail From " & Servername
oMessage.Body = "This email contains mail from the Badmail Directory on server  " & serverName & "  from the date " & sDate

'Process to add each attachment in a specified folder to the same message

'For each file in the badmail directory, attach it into the message

set ofolder = fso.getfolder("E:\halifax\smtp\badmail")
set ofiles = ofolder.files

For each file in ofiles
	oMessage.Attachfile file.path
	wscript.echo file.path,file.name
Next

'send message

oMessage.send
