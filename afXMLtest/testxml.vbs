
on error resume next

err.clear

wscript.echo " - XML 3.0: "

set oXML3 = createobject ("Msxml2.DOMDocument.3.0")

if err.number <> 0 then 

	wscript.echo "    ** ERROR **"

else

	wscript.echo "    OK!"

end if

wscript.echo " - XML 4.0: "

err.clear

set oXML4 = createobject ("Msxml2.DOMDocument.4.0")

if err.number <> 0 then 

	wscript.echo "    ** ERROR **"

else

	wscript.echo "    OK!"

end if

wscript.echo " - XML 5.0: "

err.clear

set oXML5 = createobject ("Msxml2.DOMDocument.5.0")

if err.number <> 0 then 

	wscript.echo "    ** ERROR **"

else

	wscript.echo "    OK!"

end if

wscript.echo " - XML 6.0: "

err.clear

set oXML6 = createobject ("Msxml2.DOMDocument.6.0")

if err.number <> 0 then 

	wscript.echo "    ** ERROR **"

else

	wscript.echo "    OK!"

end if