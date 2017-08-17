Dim i
i = 1
Do
Set WshShell = WScript.CreateObject("WScript.Shell") 
Return = WshShell.Run("iexplore.exe www.msn.com", 1) 
i = i+1

Loop Until i>3 