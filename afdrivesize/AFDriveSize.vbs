' AFDrivesize.vbs - Enumerates the current drives (removable, cd-rom, fixed or network) and returns the size
' Amended by - Adrian Farnell
' Version no - v1.2
' Date Amended - 22nd July 2000
	
on error resume next

	wscript.echo ""
	wscript.echo "	USAGE: cscript AFdrivesize.vbs"
	Wscript.echo ""
	Wscript.echo "		Enumerates the drives and gives"
	wscript.echo "		information on type/size."
	wscript.echo ""


wscript.echo "Drive	FreeSpace/%	Freespace/Mb	DriveType"
wscript.echo "--------------------------------------------------------"
wscript.echo showdrivelist

Function ShowDriveList
  Dim fso, d, dc, s, n
  Set fso = CreateObject("Scripting.FileSystemObject")
  Set dc = fso.Drives
  For Each d in dc
    n = ""
    s = s & chr(32) & d.DriveLetter & ": - " & chr(9)
    If d.IsReady Then
	
      'wscript.echo d.freespace & " " & d.totalsize & " " & d.availablespace
      n = (d.freespace / d.totalsize) * 100
      nsize = d.freespace / 1048576
      n = formatnumber (n, 1) & " %" & chr(9) & chr(9)
      nsize = formatnumber (nsize, 1) & chr(9) & chr(9)
  Select Case d.DriveType
    Case 0: t = "Unknown (" & d.filesystem & ")"
    Case 1: t = "Removable (" & d.filesystem & ")"
    Case 2: t = "Fixed (" & d.filesystem & ")"
    Case 3: t = "Network (" & d.filesystem & ")"
    Case 4: t = "CD-ROM (" & d.filesystem & ")"
    Case 5: t = "RAM Disk (" & d.filesystem & ")"
  End Select
   Else
		n = chr(9) & chr(9)
		nsize = chr(9) & chr(9)
	   t = "Drive Not ready"
	End If



  'Select Case d.DriveType
   ' Case 0: t = "Unknown (" & d.filesystem & ")"
    'Case 1: t = "Removable (" & d.filesystem & ")"
    'Case 2: t = "Fixed (" & d.filesystem & ")"
    'Case 3: t = "Network (" & d.filesystem & ")"
    'Case 4: t = "CD-ROM (" & d.filesystem & ")"
    'Case 5: t = "RAM Disk (" & d.filesystem & ")"
  'End Select

   
    
    s = s & n & nsize & t & vbcrlf

  Next
  ShowDriveList = s
End Function

