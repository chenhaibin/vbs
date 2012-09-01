     Dim sbox(255)
     Dim key(255)
	DIM CIPHER

set wsharguments = wscript.arguments

if wsharguments.count <> 2 then 
	
	wscript.echo "Usage: cscript rc4_inc.vbs [string to endecrypt] [password]" & vbcrlf & vbcrlf & _
		     "       [string to endecrypt] - string to encrypt or decrypt" & vbcrlf & _
		     "       [password] - password to encode by (uses RC4 encryption)"
	wscript.quit
else
	strstring = wsharguments(0)
	strpasswd = wsharguments(1)
	wscript.echo "strstring: " & strstring
	wscript.echo "strpasswd: " & strpasswd

	endecrypt strstring, strpasswd

	wscript.echo "::" & cipher & "::"
end if




   Sub RC4Initialize(strPwd)
   ':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
   ':::  This routine called by EnDeCrypt function. Initializes the :::
   ':::  sbox and the key array)                                    :::
   ':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::



      dim tempSwap
      dim a
      dim b

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
   
   Function EnDeCrypt(plaintxt, psw)
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
     ' dim cipher

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

      EnDeCrypt = cipher

   End Function


