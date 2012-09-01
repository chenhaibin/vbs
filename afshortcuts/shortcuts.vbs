' Shortcuts.vbs - Makes shortcuts for all standard progs
' Author 	- Adrian Farnell
' Date		- 18 Mar 2001
' Version	- 1.0

set objwshshell = createobject ("wscript.shell")
set objwsharguments = wscript.arguments


if objwsharguments.count <> 1 then

wscript.echo    " !Error! - Not specified any arguments" &  vbcrlf & _
		" USAGE: cscript shortcuts.vbs [Drive on which files are]"
wscript.quit

end if	

strdrive = objwsharguments.item(0)

dim arrexes(30,2)

arrexes(0,1) = "Agent\agent.exe"
	arrexes(0,2) = "ForteAGENT"
arrexes(1,1) = "cooledit\cool96.exe"
	arrexes(1,2) = "CoolEdit '96"
arrexes(2,1) = "corel40\programs\coreldrw.exe"
	arrexes(2,2) = "CorelDRAW 4.0"
arrexes(3,1) = "digiguide\client.exe"
	arrexes(3,2) = "Digiguide"
arrexes(4,1) = "emulate\windows\beeb\beebem.exe"
	arrexes(4,2) = "BeebEM"
arrexes(5,1) = "emulate\windows\doom\doom3d.exe"
	arrexes(5,2) = "WinDoom"
arrexes(6,1) = "emulate\windows\doom2\doom3d.exe"
	arrexes(6,2) = "Win Doom ]["
arrexes(7,1) = "emulate\windows\gameboy\no$gmb.exe"
	arrexes(7,2) = "Gameboy"
arrexes(8,1) = "emulate\windows\genesis\dgen.exe"
	arrexes(8,2) = "Megadrive"
arrexes(9,1) = "emulate\windows\mame\mame32.exe"
	arrexes(9,2) = "Mame 32"
arrexes(10,1) = "emulate\windows\n64\ultra.exe"
	arrexes(10,2) = "N64"
arrexes(11,1) = "emulate\windows\palm\emulator.exe"
	arrexes(11,2) = "Palm OS"
arrexes(12,1) = "emulate\windows\quake\quake.exe"
	arrexes(12,2) = "Quake"
arrexes(13,1) = "emulate\windows\snes\snes9xw.exe"
	arrexes(13,2) = "Snes"
arrexes(14,1) = "emulate\windows\sony\bleem~10.exe"
	arrexes(14,2) = "Bleem!"
arrexes(15,1) = "emulate\windows\sparcade\arcade.exe"
	arrexes(15,2) = "Sparcade"
arrexes(16,1) = "icq\icq.exe"
	arrexes(16,2) = "ICQ"
arrexes(17,1) = "ISIS Draw 2.2.1\Idraw32.exe"
	arrexes(17,2) = "ISIS Draw 2.2.1"
arrexes(18,1) = "Mirc\mirc32.exe"
	arrexes(18,2) = "mIRC32"
arrexes(19,1) = "Paint Shop Pro\psp.exe"
	arrexes(19,2) = "Paint Shop Pro v4.12"
arrexes(20,1) = "prcview\Prcview.exe"
	arrexes(20,2) = "Prcview"
arrexes(21,1) = "putty\putty.exe"
	arrexes(21,2) = "puTTY"
arrexes(22,1) = "Sonique\sonique.exe"
	arrexes(22,2) = "Sonique"
arrexes(23,1) = "Star Office\Program\soffice.exe"
	arrexes(23,2) = "Star Office"
arrexes(24,1) = "Superscan\scanner.exe"
	arrexes(24,2) = "SuperScan"
arrexes(25,1) = "vnc\vncviewer.exe"
	arrexes(25,2) = "VNCViewer"
arrexes(26,1) = "Winamp\winamp.exe"
	arrexes(26,2) = "Winamp"
arrexes(27,1) = "Winimage\winimage.exe"
	arrexes(27,2) = "Winimage"
arrexes(28,1) = "Winzip\winzip32.exe"
	arrexes(28,2) = "Winzip"
arrexes(29,1) = "Xenu\Xenu.exe"
	arrexes(29,2) = "Xenu"


for i = 0 to 29

	'wscript.echo arrexes(i,1) & " : " & arrexes(i,2)
	set objwshshortcut = objwshshell.createshortcut (arrexes(i,2) & ".lnk")
	objwshshortcut.targetpath = strdrive & "\program files\" & arrexes(i,1)
	objwshshortcut.workingdirectory = strdrive & "\Program files\" & left (arrexes(i,1), InstrRev (arrexes (i,1), "\"))
	objwshshortcut.save

next



'set objwshshortcut = objwshshell.createshortcut ("shortcuts.vbs.lnk")

'objwshshortcut.targetpath = "shortcuts.vbs"

'objwshshortcut.save