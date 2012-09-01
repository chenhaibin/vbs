
on error resume next

set oFSO = createobject ("scripting.filesystemobject")

set ofolder = ofso.getfolder ("c:\temp")

set colFiles = ofolder.files

set oConnection = createobject ("ADODB.Connection")

oConnection.Open "Provider=sqloledb;" & _
           	"Data Source=(local);" & _
           	"Initial Catalog=EventLogs;" & _
          	"User ID=test;" & _
           	"Password=test"	

for each file in colfiles

	if (Mid (file.name, InstrRev(file.name, "."))) = ".log" then
	
		wscript.echo file.name
	
		set ofile= ofso.opentextfile (file.path)
		
		y = 0


		Do while oFile.atendofstream <> true

			y = y + 1

			strLine = oFile.Readline

			strLine = Replace (strline, "20030326#", "03-26-2003 ")
			strLine2=strLine
			arrLine = Split (strLine, "#")
			
			arrLine(7) = "ow" & Mid (strLine2, instr(strLine2, "#", 6))
			
			
			
			if Left (strLine, 10) = "03-26-2003" then

				'wscript.echo x & " # " & EvID& " # " &EvType& " # " &strSource& " # " &EvDate& " # " &strComputer& " # " &strAccount& " # " &strDesc
				'
				' Here's the sql I want to perform

				' start a command object
				set oCommand = createObject ("ADODB.Command")
								
				arrLine(7) = Replace (arrLine(7), vbcrlf, "")
				arrLine(7) = Replace (arrLine(7), vbcr, "")
				arrLine(7) = Replace (arrLine(7), vblf, "")
				arrLine(7) = Replace (arrLine(7), "'", "")
				arrLine(7) = Replace (arrLine(7), "--", "")
				arrLine(7) = Replace (arrLine(7), "[", "")
				arrLine(7) = Replace (arrLine(7), "]", "")
				arrLine(7) = Replace (arrLine(7), "=", "")
				
				wscript.echo arrLine (7)

				' set the command text (this is our sql query)
				oCommand.CommandText =   "INSERT INTO " & strLocal & ".EventLogs.dbo.Logs " & _ 
				"(EntryID, EventClass, oDateTime, EventID, Source, Server, oDescription) " & _ 
				"VALUES ('" & y & "', '" &  arrLine (2) & "', '" & arrLine(0) & "', '" & arrLine(4) & "', '" & arrLine(1) &"', '" & arrLine(6) & "', '" & arrLine(7) & "')"

				 ' Use the connection already set up by ADO
				oCommand.ActiveConnection = oConnection
					
				err.clear

				' Execute the sql query, return any text into the results string
				oCommand.Execute 'strResults
				
				if err.number <> 0 then
					wscript.echo Len (arrLine(7))
				end if

				'wscript.echo "Results " & Results & " added to DB."

				' close the command object
				set oCommand = nothing

			else
				'wscript.echo "skipping line '" & strLine & "'"
			end if


		Loop
		
		set ofile = nothing
		
		'wscript.quit
		
	else
		'skip file
	end if

	
next

wscript.quit



