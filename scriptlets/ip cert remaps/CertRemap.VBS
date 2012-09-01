'/****************************************************************************/
'/*                                                                          */
'/* VBScript                                                                 */
'/*  CerificateReMap.vbs    - Configures IIS Metabase (Certificates)         */
'/*                                                                          */
'/* SYNOPSIS                                                                 */
'/*  Moves certificate details from disabled to an IP address                */
'/*  Expects 2 parameters (Existing Key and New Key)                         */
'/*                                                                          */
'/*  PLATFORM: E-Commerce                                                    */
'/*                                                                          */
'/* AUTHOR                                                                   */
'/*  Simon Turner                                                            */
'/*                                                                          */
'/* DATE                                                                     */
'/*  11th August 2000          - (1.0) Creation                              */
'/*                                                                          */
'/****************************************************************************/

Option Explicit

'/****************************************************************************/
'/*  Declare variables
'/****************************************************************************/
Dim objRoot 
Dim objNewKey
Dim strOldKey
Dim strNewKey

'/****************************************************************************/
'/*  Don't break on error
'/****************************************************************************/
on error resume next

'/****************************************************************************/
'/*  Check the right number of parameters were given
'/****************************************************************************/
If wscript.Arguments.Count <> 2 Then
	DisplayUsage
End If

'/****************************************************************************/
'/*  Assign parameters
'/****************************************************************************/
strOldKey = wscript.Arguments.Item(0)
strNewKey = wscript.Arguments.Item(1) & ":443"

'/****************************************************************************/
'/*  Connect to the SSL root key
'/****************************************************************************/
Set objRoot = GetObject("IIS://LocalHost/W3SVC/SSLKeys")
If err.number <> 0 then
	wscript.echo "Couldn't bind to SSL Root Key"
	wscript.Quit
End If

'/****************************************************************************/
'/*  Move old key to new key
'/****************************************************************************/ 
Set objNewKey = objRoot.MoveHere(strOldKey, strNewKey) 
'If err.number <> 0 then
'	wscript.echo "Couldn't bind to Sub Key - Certificate could not be remapped"
'	wscript.Quit
'End If

'/****************************************************************************/
'/*  Release the objects
'/****************************************************************************/
Set objNewKey = nothing 
Set objRoot = nothing 

'/****************************************************************************/
'/*                                                                          */
'/* Sub DisplayUsage()                                                       */
'/*                                                                          */
'/* SYNOPSIS                                                                 */
'/*  Displays example usage and then quits                                   */
'/*                                                                          */
'/* AUTHOR                                                                   */
'/*  Simon Turner                                                            */
'/*                                                                          */
'/* DATE                                                                     */
'/*  11th August 2000          - (1.0) Creation                              */
'/*                                                                          */
'/****************************************************************************/
Sub DisplayUsage()
	
	wscript.Echo "CertRemap.VBS - Usage"
	wscript.Echo ""
	wscript.Echo "cscript certmap.vbs <Old Key name> <New Key name>"
	wscript.Echo ""
	wscript.Echo "i.e."
	wscript.Echo ""
	wscript.Echo "cscript certmap.vbs disabled1 192.5.1.201"
	wscript.Echo ""
	wscript.Quit

End Sub