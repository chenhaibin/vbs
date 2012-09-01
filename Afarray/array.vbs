
Dim ArrStrSplit
Dim StrSplit

StrSplit = "ECOM0001,ECOM0002,ECOM0011,ECOM0021,ECOM0022,ECOM0023,ECOM0024,ECOM0101,ECOM0103,ECOM0105,ECOM0201,ECOM0203,ECOM0204,ECOM0205"

Wscript.echo StrSplit
Wscript.echo ""

ArrStrSplit = Split (StrSplit, ",")

Wscript.echo "Upper Boundary " & UBound (ArrStrSplit)

StrUpperlimit = UBound(ArrStrSplit)

For server = 0 to StrUpperlimit
	
	Wscript.echo server & ": " & ArrStrSplit (server)

Next