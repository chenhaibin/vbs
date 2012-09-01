
arrFirstName = Array ("Harry", "Hermione", "Ron", "Hagrid", "Gary", "Spongebob", "Patrick", "James", "Jean", "John", "Julie", "Julia", _
	    	"Simon", "Emily", "Mary", "Sarah", "Graham", "Malcolm", "William", "Wesley", "Harriet", "Robert", "Phillip")

arrSurName = Array ("Smith", "Jones", "Taylor", "Brown", "Williams", "Wilson", "Johnson", "Davis", "Robinson", "Wright", "Thompson", "Evans", "Walker", "White", "Roberts", "Green", "Hall", "Wood", "Jackson", "Clarke" )

arrDomain = Array ("hotmail.com", "yahoo.com", "google.com", "gmail.com", "mobileme.com", "apple.com", "microsoft.com", "dropbox.com", "zoho.com", "sky.com", "itv.com", "bbc.co.uk", "gmail.co.uk", "freeserve.com")

for i = 0 to 30

	strFirst = arrFirstName (return_RandomInt(23)) 
	strSur = arrSurName (return_RandomInt(20))
	strEmail = strFirst & strSur & "@" & arrDomain(return_RandomInt(14))

	wscript.echo strFirst & " " & strSur & "," & strEmail & "," & return_RegistrantNo

next


function return_RandomInt(intNumber)

	Randomize

	return_RandomInt = int (intNumber * Rnd())

end function


function return_RegistrantNo

	for j = 0 to 10

		strCustomerNo = strCustomerNo & return_RandomInt(9)

	next
	
	return_RegistrantNo = strCustomerNo

end function