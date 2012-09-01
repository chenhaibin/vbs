
set ofso = createobject ("scripting.filesystemobject")

set ofile = ofso.OpenTextFile("extract.txt", 1)

set oConnection = createobject ("ADODB.connection")

strUser = "GT32402"
strPassword = "AFQu3ry9023!"
strServer = "10.225.242.23"

oConnection.Open = "Provider=sqloledb;" & _
			"Data Source=" & strServer & ";" & _
			"Initial Catalog=HxSupport;" & _
			"User ID=" & strUser & ";" & _
			"Password=" & strPassword


Do while (ofile.atEndOfStream<>true)

    strline = ofile.readline
    
    arrLine = split (strLine, ",")

    strGUID = getGuid (arrLine(0))

    wscript.echo strGUID    

Loop


function getGUID (userID)

    set objRS = oConnection.Execute("select UserId,HxUniqueID from tbl_AdamExtract_20100628 where UserID = '" & UserID & "'")
    
    Do while not objRS.eof
    
        strResults = objRS("UserID") & "," & objRS("HxUniqueID")
        
        'wscript.echo strResults
    
        objRS.movenext
    
    Loop 
    
    set objRS = nothing

    getGuid = strResults

end function