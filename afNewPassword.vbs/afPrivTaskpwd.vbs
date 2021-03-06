
Dim sbox(255)
Dim key(255)
Dim StrPassword

'* start scheduler object *
set objScheduler = createobject ("Scheduler.SchedulingAgent.1")

'* start shell object *
set objWshShell = createobject ("wscript.shell")

'* enumerate through tasks to find priv.job *
for each objTask in objscheduler.tasks

	'wscript.echo "[" & objtask.jobname & "]"

	if lcase (objtask.jobname) = "priv.job" then
		
		wscript.echo "Priv Job found"
		
		'* get password and decrypt *
		strkey = objwshshell.regread ("HKLM\Software\AF\eky")
		strOldPassword = objwshshell.regread ("HKLM\Software\AF\dwpcnesvc")
		arrOldEncrypted = split ( trim(strOldPassword) )
		for kcount = 0 to ubound (arrOldEncrypted)
			strReconstructString = strReconstructString & Chr (arrOldEncrypted(kcount))
		next
		strOldPassword = FuncEnDecrypt (strReconstructString, strkey)
		
		
		'* set account info *
		objtask.setaccountinformation "afbackup", strOldPassword
		
	else
		wscript.echo "No task found"
	end if

next



function funcCreatePassword

	Dim icount
	Dim strUpperLimit, strLowerLimit, strRnd

	for icount = 1 to 14
		
		randomize
		strUpperLimit = 33
		strLowerLimit = 122
		strRnd = Int ((strUpperLimit - strLowerLimit) * Rnd () + strLowerLimit)
		
		strPassword = strPassword & chr(strRnd)
		strAscii = strAscii & " " & strRnd
	next

	'wscript.echo strPassword
	
end function

Sub RC4Initialize(strPwd)
   ':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
   ':::  This routine called by EnDeCrypt function. Initializes the :::
   ':::  sbox and the key array)                                    :::
   ':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

      dim tempSwap
      dim a
      dim b
      dim intlength
      
      intLength = len(strPwd)
      For a = 0 To 255
         key(a) = asc(mid(strpwd, (a mod intLength)+1, 1))
         sbox(a) = a
      next

      b = 0
      For a = 0 To 255
         b = (b + sbox(a) + key(a)) Mod 256
         tempSwap = sbox(a)
         sbox(a) = sbox(b)
         sbox(b) = tempSwap
      Next
   
End Sub


Function funcEnDeCrypt(plaintxt, psw)
   ':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
   ':::  This routine does all the work. Call it both to ENcrypt    :::
   ':::  and to DEcrypt your data.                                  :::
   ':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

      dim temp
      dim a
      dim i
      dim j
      dim k
      dim cipherby
      dim cipher

      i = 0
      j = 0

      RC4Initialize psw

      For a = 1 To Len(plaintxt)
         i = (i + 1) Mod 256
         j = (j + sbox(i)) Mod 256
         temp = sbox(i)
         sbox(i) = sbox(j)
         sbox(j) = temp
   
         k = sbox((sbox(i) + sbox(j)) Mod 256)

         cipherby = Asc(Mid(plaintxt, a, 1)) Xor k
         cipher = cipher & Chr(cipherby)
      Next

      funcEnDeCrypt = cipher

End Function