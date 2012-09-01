' *******************************************
' Name:		afNewPassword.vbs
' Notes:	Script to create new pwds
'		for test user account
' Version:	0.1
' Author:	Adrian Farnell
' Date:		21/11/2001
' *******************************************

Option Explicit

' Dim Variables
Dim ObjwshShell, objuser, objwshnetwork
Dim strPassword, strEncryptedPwd, strDecryptedPwd, strKey, strLocal, strEncoldpwd, strDecOldPwd, struser
Dim constLengthofPassword
Dim sbox(255)
Dim key(255)

' start objects
set objWshShell = createobject ("wscript.shell")
set objwshnetwork = createobject ("wscript.network")

' Set machine name
strLocal = objwshnetwork.computername

' Which User to rotate
strUser = "backup"

' Set length of password
constLengthofPassword = 14

' Find Key in registry
strkey = objWshShell.regread ("HKLM\Software\Microsoft\dwpcnesvc\strkey")

' Call function to create password
funcCreatePassword
strEncryptedPwd = funcEnDeCrypt (strPassword, strkey)

' Get old encrypted password
strEncOldpwd = objWshShell.regread ("HKLM\Software\Microsoft\dwpcnesvc\strpwd")

' Decrypt it, so that we can use it to set the new password
strDecOldpwd = funcEnDecrypt (strEncOldPwd, strKey)

' Write new encrypted password to registry
objWshShell.regwrite "HKLM\Software\Microsoft\dwpcnesvc\strpwd", strEncryptedPwd

' set testuser users password
set objUser = getobject ("WinNT://" & strLocal & "/" & struser)
objUser.ChangePassword strDecOldPwd, strPassword
objuser.setinfo

' write logfile
objWshShell.logevent 4, " testUser set to {" & strEncryptedPwd & "} on machine " & strLocal

' End execution of script
set objUser = nothing
set objWshShell = nothing
wscript.quit(0)

' ************** functions *******************

function funcCreatePassword

	Dim icount
	Dim strUpperLimit, strLowerLimit, strRnd

	for icount = 1 to constlengthofpassword
		
		randomize
		strUpperLimit = 33
		strLowerLimit = 122
		strRnd = Int ((strUpperLimit - strLowerLimit) * Rnd () + strLowerLimit)
		
		strPassword = strPassword & chr(strRnd)
		
	next

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