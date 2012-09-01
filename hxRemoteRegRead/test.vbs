
'start xml object
Set objXMLDoc = createobject ("MSXml2.DOMDocument.4.0")

objXMLDoc.async = false
objXMLDoc.validateOnParse = true

'load up xml
objXMLDoc.Load ("machines.xml")

'enumerate servers

set ElemList = objXMLDoc.getElementsbyTagName ("SERVER")

for i=0 to (ElemList.length-1)

	wscript.echo elemlist.item(i).xml

next

'figure out which user/pass to use



function funcCheckState

	strState = objXMLDoc.readyState
	select case strState
		
		case 1,2,3,4
			wscript.echo "XML at state: " & strState
		case else
			wscript.echo "Unknown XML State"
	end select


end function

function funcReadReg (strServerName, strUser, strPass)

end function