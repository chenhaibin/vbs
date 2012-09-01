
' Measure network speed

strTestFile = "testFile.tmp"
strTestFolder = ".\Lots"
strFolderDest = "Test"
strDestination = "\\spleen\e$\"

subOneFile

subonefolder

sub subOneFile

	set oFSO = createobject ("scripting.filesystemobject")

	if oFSO.fileexists ( strDestination & "\" & strTestFile ) = true then
		wscript.echo "file exists at destination" 
		wscript.quit
	else

	end if
	
	wscript.echo " ----------------------------------------------"
	wscript.echo " Starting Upload of 1.2MB File to host"

	strStart = Timer
	wscript.echo " Started Copy at       : " & strStart

	oFSO.copyFile strTestFile, strDestination

	strFinish = Timer
	wscript.echo " Finished Copy at      : " & strFinish
	wscript.echo " Time Taken (/s)       : " & FormatNumber ((strFinish - strStart),2) '& DateDiff ("s", strstart, strfinish) & " secs"

	strTransferBpersec = 1230029 / (strFinish - strStart)'DateDiff ("s", strstart, strfinish)

	wscript.echo " Transfer Speed (MB/s) : " & FormatNumber (strTransferBpersec / 1048576,2) & " MB/s (" &  FormatNumber (strTransferBpersec / 1048576,2) * 1024 & " KB/s)"
	wscript.echo " Transfer Speed (Mb/s) : " & FormatNumber (strTransferBpersec * 8 / 1048576,2) & " Mb/s"

	wscript.echo " ----------------------------------------------"
	wscript.echo " Starting Download of 1.2MB File from host"

	strStart = Timer
	wscript.echo " Started Copy at       : " & strStart

	oFSO.copyFile strDestination & "\" & strTestFile, "c:\temp.tmp"

	strFinish = Timer
	wscript.echo " Finished Copy at      : " & strFinish
	wscript.echo " Time Taken (/s)       : " & FormatNumber ((strFinish - strStart),2) '& DateDiff ("s", strstart, strfinish) & " secs"

	strTransferBpersec = 1230029 / (strFinish - strStart)'DateDiff ("s", strstart, strfinish)

	wscript.echo " Transfer Speed (MB/s) : " & FormatNumber (strTransferBpersec / 1048576,2) & " MB/s (" &  FormatNumber (strTransferBpersec / 1048576,2) * 1024 & " KB/s)"
	wscript.echo " Transfer Speed (Mb/s) : " & FormatNumber (strTransferBpersec * 8 / 1048576,2) & " Mb/s"


	oFSO.deletefile ( "c:\temp.tmp")
	oFSO.deletefile ( strDestination & "\" & strTestFile)

	set oFSO = nothing

end sub

sub subOneFolder

	set oFSO = createobject ("scripting.filesystemobject")
	
	if oFSO.folderexists (strDestination & strFolderDest) = true then
		wscript.echo strDestination & strFolderDest
		wscript.echo "folder exists at destination"
		wscript.quit
	end if

	wscript.echo " ----------------------------------------------"
	wscript.echo " Starting Upload of 1.1MB of small files to host"	

	strStart = Timer
	wscript.echo " Started Copy at       : " & strStart

	oFSO.copyFolder strTestFolder, strDestination & strFolderDest

	strFinish = Timer
	wscript.echo " Finished Copy at      : " & strFinish
	wscript.echo " Time Taken (/s)       : " & FormatNumber ((strFinish - strStart),2) '& DateDiff ("s", strstart, strfinish) & " secs"

	strTransferBpersec = 1079295 / (strFinish - strStart)'DateDiff ("s", strstart, strfinish)

	wscript.echo " Transfer Speed (MB/s) : " & FormatNumber (strTransferBpersec / 1048576,2) & " MB/s (" &  FormatNumber (strTransferBpersec / 1048576,2) * 1024 & " KB/s)"
	wscript.echo " Transfer Speed (Mb/s) : " & FormatNumber (strTransferBpersec * 8 / 1048576,2) & " Mb/s"


	wscript.echo " ----------------------------------------------"
	wscript.echo " Starting Download of 1.1MB of small files from host"	

	strStart = Timer
	wscript.echo " Started Copy at       : " & strStart

	oFSO.copyFolder strDestination & strFolderDest, "c:\testAF"

	strFinish = Timer
	wscript.echo " Finished Copy at      : " & strFinish
	wscript.echo " Time Taken (/s)       : " & FormatNumber ((strFinish - strStart),2) '& DateDiff ("s", strstart, strfinish) & " secs"

	strTransferBpersec = 1079295 / (strFinish - strStart)'DateDiff ("s", strstart, strfinish)

	wscript.echo " Transfer Speed (MB/s) : " & FormatNumber (strTransferBpersec / 1048576,2) & " MB/s (" &  FormatNumber (strTransferBpersec / 1048576,2) * 1024 & " KB/s)"
	wscript.echo " Transfer Speed (Mb/s) : " & FormatNumber (strTransferBpersec * 8 / 1048576,2) & " Mb/s"


	oFSO.deletefolder ("c:\testAF")
	oFSO.deletefolder ( strDestination & strFolderDest)
	wscript.echo " ----------------------------------------------"
	set oFSO = nothing
	
end sub