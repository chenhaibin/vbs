
strStart = Timer

wscript.echo funcCreate32

strFinish = Timer

wscript.echo "Took " & strFinish - strStart & " secs"

function funcCreate32

	strOutput = ""
	
	for i = 1 to 14
		
		Randomize
		if rnd > 0.5 then
			Randomize
			strChr = int (10 * Rnd () + 48)
			
			strOutput = strOutput & Chr (strChr)
		else
			Randomize 
			strChr = int (26 * Rnd () + 97)
			if chr (strChr) = "`" then wscript.echo strChr
			strOutput = strOutput & chr (strChr)
		end if
	
	next

	funcCreate32 = strOutput

end function 