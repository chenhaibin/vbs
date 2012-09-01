
strDate = "16/08/2002 08:34:58"

arrSplit = Split (strDate)

arrDate = Split (arrSplit(0), "/")

strDD = arrDate (0)
strMM = arrDate (1)
strYYYY = arrDate (2)

strReconstruct = strMM & "-" & strDD & "-" & strYYYY & " " & arrSplit(1)

wscript.echo "started as " & strDate
wscript.echo "finished as " & strReconstruct
