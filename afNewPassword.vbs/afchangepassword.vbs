

Dim sbox(255)
Dim key(255)
Dim StrPassword, strascii


'* start wshshell object *
set objWshShell = createobject ("wscript.shell")
set objwshnetwork = createobject ("wscript.network")

'* Create A new password *
funcCreatePassword
strNewPassword = strPassword
strkey = objwshshell.regread ("HKLM\Software\AF\eky")
strlocal = objwshnetwork.computername
wscript.echo "Created new password: " & strpassword
wscript.echo "Computername        : " & strlocal


'* Get old password *
strOldPassword = objwshshell.regread ("HKLM\Software\AF\dwpcnesvc")
arrOldEncrypted = split ( trim(strOldPassword) )
for kcount = 0 to ubound (arrOldEncrypted)
	strReconstructString = strReconstructString & Chr (arrOldEncrypted(kcount))
next
strOldPassword = FuncEnDecrypt (strReconstructString, strkey)
wscript.echo "Old Password        : " & strOldPassword

'* Encrypt password *
strEncryptedPassword = funcEnDeCrypt (strNewPassword, strkey)

'* Turn Encrypted password into string of ascii numbers *
for jcount = 1 to len (strEncryptedPassword)

	strChar = Mid(strEncryptedPassword, Jcount)
	strEncryptedAscii = strEncryptedAscii & " " & Asc (strChar)

next

'* Put encrypted ascii numbers into registry *
wscript.echo "Writing to registry : " & strEncryptedAscii
objwshshell.regwrite "HKLM\software\AF\dwpcnesvc", strEncryptedAscii


'* Now change the password.... *
wscript.echo "Changing Password"
set objUser = getobject ("WinNT://" & strLocal & "/AFBackup")
objuser.changepassword strOldPassword, strNewPassword 
objuser.setinfo

wscript.echo "Log to eventvwr"
objWshShell.logevent 4, " testUser set to " & strEncryptedAscii & " Using Key " & strKey

'arrEncrypted = split (strEncryptedAscii)
'for kcount = 0 to Ubound (arrEncrypted)
'	wscript.echo "!" & arrEncrypted(kcount)
'	strReconstructString = strReconstructString & Chr(arrEncrypted(kcount))
'next

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