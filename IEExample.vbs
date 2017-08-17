Option Explicit

'    Declaration of variables
Dim ie

'    Declaration of subroutines
'    Sub WaitForLoad waits for webpage to finish loading
'    before proceeding with next line of code
Sub WaitForLoad
    Do while ie.Busy or ie.readystate <> 4
        wscript.sleep 200
    Loop
End Sub

'    Open windows Internet explorer and surf to website
set ie = CreateObject("InternetExplorer.Application")
With ie
    .Navigate "enter url here"
    .Toolbar=0
    .StatusBar=0
    .Height=560
    .Width=1000
    .Top=0
    .Left=0
    .Resizable=0
    WaitForLoad
    .Visible = true
end with

'  Add code below to display in a msgbox the contents of each cell in column 0 (first column) should loop for each row.

dim table, row
Set table = ie.document.getelementsbytagname("table")(0)
for each row in table.rows
    MsgBox "row " & row.rowIndex + 1 & " " & row.cells(0).innerText & " " & row.cells(1).innerText 
next