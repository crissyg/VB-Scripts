
Set WshShell = WScript.CreateObject("WScript.Shell") 
Return = WshShell.Run("iexplore.exe www.msn.com", 1) 
Return = WshShell.Run("chrome.exe www.msn.com", 1) 
Return = WshShell.Run("firefox.exe www.msn.com", 1) 


