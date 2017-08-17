'Description - to get all the Programmatic Identifiers off a Windows OS system
'See : https://msdn.microsoft.com/en-us/library/windows/desktop/cc144152(v=vs.85).aspx

Option Explicit

Const HKEY_CLASSES_ROOT = &H80000000

Dim arrProgID, lstProgID, objReg, strMsg, strProgID, strSubKey, subKey, subKeys()

Set lstProgID = CreateObject( "System.Collections.ArrayList" )
Set objReg    = GetObject( "winmgmts://./root/default:StdRegProv" )

' List all subkeys of HKEY_CLASSES_ROOT\CLSID
objReg.EnumKey HKEY_CLASSES_ROOT, "CLSID", subKeys

' Loop through the list of subkeys
For Each subKey In subKeys
	' Check each subkey for the existence of a ProgID
	strSubKey = "CLSID\" & subKey & "\ProgID"
	objReg.GetStringValue HKEY_CLASSES_ROOT, strSubKey, "", strProgID
	' If a ProgID exists, add it to the list
	If Not IsNull( strProgID ) Then lstProgID.Add strProgID
Next

' Sort the list of ProgIDs
lstProgID.Sort

' Copy the list to an array (this makes displaying it much easier)
arrProgID = lstProgID.ToArray

' Display the entire array
WScript.Echo Join( arrProgID, vbCrLf )