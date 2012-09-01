
' *********************************************************
' Name:		afDisplayAllNICs.vbs
' Notes:	Enumerates through NICs using WMI display
' 		properties
' Version:	1.0
' Date:		26/11/2001
' Author:	Adrian Farnell
' *********************************************************

Option Explicit

' ** Dim Objects, variables, constants **

Dim objWMI, objNICs,  objNICproperty, objNICproperty2
Dim NIC, property, property2,ip
Dim strAdapterName, stripx, wsharguments, strAdapter, strargument
Dim arrIP()

' ** Start WMI Object **
set objWMI = getobject ("winmgmts:")

' ** Enumerate thru instances of win32_NetworkNic
set objNICs = objWMI.InstancesOf ("win32_NetworkAdapterConfiguration")

' ** Arguments **
set wsharguments = wscript.arguments

if wsharguments.count <> 1 then
	subdisplayusage
end if

if wsharguments(0) = "/?" then
	subDisplayUsage
end if



strArgument = wsharguments(0)

if lcase (strargument) = "/all" then 
	
	subEnumerateAllNics
	wscript.quit
	
end if

if left (lcase (strargument),8) = "/findnic" then

	strAdapter = Mid (wsharguments(0), 10)
	funcFindNic strAdapter
	wscript.quit

end if

wscript.echo " Error! - Invalid Argument Specified"

sub subDisplayUsage

	wscript.echo " USAGE: cscript afDisplayAllNICs.vbs [/all] [/findNIC:name_of_key]" & _
	 vbcrlf & "  [/all] - Used to display all info on all installed NICs" & vbcrlf & _
	 "  [/findNIC:name_of_key] - Used to display only detail on one specific NIC" & vbcrlf & _
	 "                           name_of_key is a unique string identifying NIC" & vbcrlf & _
	 vbcrlf & "  IMPORTANT NOTE: Script can only use one argument, not both at the"  & vbcrlf & _
	 "  same time. If both are specified, script will use the first" & vbcrlf & _
	 vbcrlf & "  eg. cscript afdisplayAllNICs.vbs /findNIC:Intel" & vbcrlf & _
	 "  would find all nics, with Intel in the description field"
	wscript.quit

end sub


sub subEnumerateAllNICs

	wscript.echo "Searching for All NICs......" & vbcrlf & "---------------------------" & vbcrlf
	For each NIC in objNICs
		
		set objNICproperty = NIC.properties_
		
		'For each property in ObjNICproperty
	
			'if lcase (property.name) = "description" then
				'if Instr (property.value,stradaptername) <> 0 then
				
						'wscript.echo vbcrlf & " MADGE Card Found, Properties are: "
						
						'set objNICproperty2 = NIC.properties_
						For each property2 in objNICproperty
							
	
							
							if Vartype (property2.value) = 8204 then
								wscript.echo "  " & property2.name & ": ** VarType is Array **"
							else
								wscript.echo "  " & property2.Name & ": " & property2.value
							end if			
							
							'if lcase (property2.name) = "ipaddress" then
								
							'end if
							
				
						Next
	
				'else
				'	wscript.echo " Message, (" & property.value & ") adapter found."
				'end if
			'else
			'end if	
	
			'Next
			wscript.echo " ---------------------------------"
	Next

End Sub

function funcFindNIC (strAdapterName)

	wscript.echo "Searching for NICs......" & strAdaptername & vbcrlf & "---------------------------" & vbcrlf
	For each NIC in objNICs
		
		set objNICproperty = NIC.properties_
		
		For each property in ObjNICproperty
	
			if lcase (property.name) = "description" then
				if Instr (lcase (property.value), lcase(stradaptername)) <> 0 then
				
						wscript.echo vbcrlf & " " & strAdapterName & " Card Found, Properties are: "
						wscript.echo " --------------------------------------"
						
						set objNICproperty2 = NIC.properties_
						For each property2 in objNICproperty
							
	
							
							if Vartype (property2.value) = 8204 then
								wscript.echo "  " & property2.name & ": ** VarType is Array **"
							else
								wscript.echo "  " & property2.Name & ": " & property2.value
							end if			
							
							'if lcase (property2.name) = "ipaddress" then
								
							'end if
							
				
						Next
						wscript.echo " ------------------------------------"
	
				else
					wscript.echo " Message, (" & property.value & ") adapter found."
				end if
			else
			end if	
	
		Next
		'wscript.echo " ---------------------------------"
	Next

End function
