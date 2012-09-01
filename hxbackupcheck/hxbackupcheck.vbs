' Name		: HxBackupCheck.vbs - Checks files on repository and makes a report
' Author	: Adrian Farnell
' Date Amended	: 06/09/2000
' Version	: pre-Alpha

' start filesystemobject
set objfso = createobject ("scripting.filesystemobject")

' Open text file for reading
set objfilelist = objfso.opentextfile ("filelist.txt", 1)

' Read text file line by line to set up server arrays
Do while objfilelist.AtEndOfstream <> True

	strline = objfilelist.readline
	strnormalline = strline
	strfindrightline = Left (strline, 13)
	strline = Mid (strline, 14)
	
	'setup array for each set of servers
	if strfindrightline = "<all servers>" then ArrAllServers = Split (strline, ",")
	if strfindrightline = "<web servers>" then ArrWebServers = Split (strline, ",")
	if strfindrightline = "<sql servers>" then ArrsqlServers = Split (strline, ",")
	if strfindrightline = "<mem servers>" then ArrmemServers = Split (strline, ",")
	if strfindrightline = "<aud servers>" then ArraudServers = Split (strline, ",")
	if strfindrightline = "<ses servers>" then ArrsesServers = Split (strline, ",")

	'Checks for Every Hour string creation
	if left(strnormalline, 15) = "<HOURLY CHECKS>" then strstartrecording=1
	if left(strnormalline, 20) = "<00:00-00:59 CHECKS>" then strstartrecording=0
	if strstartrecording=1 then streveryhourtext = streveryhourtext & "," & strnormalline
			
	'Checks 00:00 - 00:59
	if left(strnormalline, 20) = "<00:00-00:59 CHECKS>" then strstartrecording=2
	if left(strnormalline, 20) = "<01:00-01:59 CHECKS>" then strstartrecording=0
	if strstartrecording=2 then strhouronetext = strhouronetext & "," & strnormalline

	'Checks 01:00 - 01:59
	if left(strnormalline, 20) = "<01:00-01:59 CHECKS>" then strstartrecording=3
	if left(strnormalline, 20) = "<02:00-02:59 CHECKS>" then strstartrecording=0
	if strstartrecording=3 then strhourtwotext = strhourtwotext & "," & strnormalline

	'Checks 02:00 - 02:59
	if left(strnormalline, 20) = "<02:00-02:59 CHECKS>" then strstartrecording=4
	if left(strnormalline, 20) = "<03:00-03:59 CHECKS>" then strstartrecording=0
	if strstartrecording=4 then strhourthreetext = strhourthreetext & "," & strnormalline

	'Checks 03:00 - 03:59
	if left(strnormalline, 20) = "<03:00-03:59 CHECKS>" then strstartrecording=5
	if left(strnormalline, 20) = "<04:00-04:59 CHECKS>" then strstartrecording=0
	if strstartrecording=5 then strhourfourtext = strhourfourtext & "," & strnormalline

	'Checks 04:00 - 04:59
	if left(strnormalline, 20) = "<04:00-04:59 CHECKS>" then strstartrecording=6
	if left(strnormalline, 20) = "<05:00-05:59 CHECKS>" then strstartrecording=0
	if strstartrecording=6 then strhourfivetext = strhourfivetext & "," & strnormalline

	'Checks 05:00 - 05:59
	if left(strnormalline, 20) = "<05:00-05:59 CHECKS>" then strstartrecording=7
	if left(strnormalline, 20) = "<06:00-06:59 CHECKS>" then strstartrecording=0
	if strstartrecording=7 then strhoursixtext = strhoursixtext & "," & strnormalline

Loop


' Set up date strings
strrawdateandtime =  date & " " & time
dd = left (strrawdateandtime, 2)
mm = mid (strrawdateandtime, 4, 2)
yyyy = mid (strrawdateandtime, 7, 4)
yy = mid (strrawdateandtime, 9, 2)
hh = hour (time)

wscript.echo "DD: " & dd
wscript.echo "MM: " & mm
wscript.echo "YYYY: " & yyyy
wscript.echo "YY: " & yy
wscript.echo "hh: " & hh

'manipulate every hour into an array so we can process it

'generate list of files to check for every hour

wscript.echo "streveryhourtext: " & streveryhourtext
 
strfilelist = streveryhourtext
			
funcManipulateStrings strfilelist, arreveryhourfiles


Function funcManipulateStrings (strhours, arrhourfiles)

	'replace Date symbols with actual date
	strhours = Mid (strhours,64)
	strhours = replace (strhours, "HHmm", HH & "00")
	strhours = replace (strhours, "DD", DD)
	strhours = replace (strhours, "MM", MM)
	strhours = replace (strhours, "YYYY", YYYY)
	strhours = replace (strhours, "YY", YY)
	
	'get rid of Hour marker
		
	
	'make array of files
	arrhourfiles = Split (strhours, ",")
	
	'Display array
	for each line in arrhourfiles
		i=i+1
		wscript.echo i & " : " &  line
	next
	
		
end function

























wscript.quit


' Hourly Checks
Wscript.echo "Current time is: " & time
Wscript.echo "Checking for files up to: " & hh & ":00"





wscript.echo strfilelist

for each audserver in arraudservers
	for each a in arreveryhourfiles
		a = replace (a, "<aud servers>", audserver)
'		wscript.echo a
	next
next

for each webserver in arrwebservers
	for each a in arreveryhourfiles
		a = replace (a, "<webservers>", webserver)
'		wscript.echo a
	next
next
