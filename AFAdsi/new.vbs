
strDomain = "EBUSINESS"
strLocation = "Copley"
strServer = "EBUS0201"
strGroup = "RBIT"
strRole = "Web"

subAddServer2XML strDomain, strLocation, strServer, strGroup, strRole

sub subAddServer2XML (strDomain,strLocation, strServer, strGroup, strRole)

	set objDOM = createObject ("Microsoft.XMLDOM")

	'Turn off asynchronous file loading
	objDOM.async = False

	'loads my xml doc
	blnFileExist = objDOM.load ("New.xml")

	'tested for its existence
	if blnFileExist = False then
		wscript.echo "file doesn't exist"
		set objRoot = objDOM.createElement ( strDomain )
		objDOM.appendChild objRoot
	else
		wscript.echo "file exists"
		set objRoot = objDOM.documentElement
	end if


	Set objNode = objDOM.createElement("customer")
	
	AddNodeAttribute objDOM, objNode, "CustomerID", "XMLP"
	AddNodeAttribute objDOM, objNode, "CompanyName", "XMLPitstop.com"




	objDOM.save ("New.xml")

	set objDOM = nothing

end sub

Sub AddDomElement(objDom, strName, strValue)

	Dim objNode

	'Add new node
	Set objNode = objDom.createElement(strName)
	objNode.Text = strValue
	objDom.documentElement.appendChild objNode

End Sub

Sub AddNodeAttribute(objDom, objNode, strName, strValue)

	Dim objAttrib

	Set objAttrib = objDOM.createAttribute(strName)
	objAttrib.Text =strValue
	objNode.Attributes.setNamedItem objAttrib
	objDOM.documentElement.appendChild objNode
	
End Sub