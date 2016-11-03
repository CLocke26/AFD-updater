Const ForReading = 1 
Set WshShell = WScript.CreateObject("WScript.Shell") 
Set objFSO = CreateObject("Scripting.FileSystemObject") 
Set objTextFile = objFSO.OpenTextFile _ 
    ("C:\Users\cl21826\Documents\afdbarcodes1.txt", ForReading)

strNumber = 1


Set objOutputFile = objFSO.CreateTextFile("C:\Users\cl21826\Documents\completedbarcodes.txt")
 
Do Until objTextFile.AtEndOfStream 
    strNextLine = objTextFile.Readline 
    arrServiceList = Split(strNextLine , ",") 
    

For count = 0 to 23
	wshShell.SendKeys "{TAB}"
	next
	
WScript.Sleep(2000)
wshShell.SendKeys arrServiceList(0)
WScript.Sleep(1000)
wshShell.SendKeys "{enter}"
WScript.Sleep(1000)
wshShell.SendKeys "{enter}"
WScript.Sleep(10000) 

For count = 0 to 23
	wshShell.SendKeys "{TAB}"
	next

WScript.Sleep(3000)
wshShell.SendKeys "RETURNED TO CONFIG CENTER"
WScript.Sleep(2000)
wshShell.SendKeys "{TAB}"
WScript.Sleep(2000)
wshShell.SendKeys "In stock"
WScript.Sleep(2000)

For count = 0 to 30
	wshShell.SendKeys "{TAB}"
	next

WScript.Sleep(500)
wshShell.SendKeys "{backspace}"
wshShell.SendKeys "{TAB}"
WScript.Sleep(500)
wshShell.SendKeys "/UNITED STATES/X/X/X/S/STOCK CAGE/"
WScript.Sleep(500)											
wshShell.SendKeys "{TAB}"
WScript.Sleep(300)
wshShell.SendKeys "RX00 X S BLDG STOCK CAGE"
WScript.Sleep(500)
wshShell.SendKeys "{TAB}"
WScript.Sleep(300)
wshShell.SendKeys "6164981"
WScript.Sleep(500)

For count = 0 to 2
	wshShell.SendKeys "{TAB}"
	next
	
WScript.Sleep(300)
wshShell.SendKeys "rx00"
WScript.Sleep(1000)
wshShell.SendKeys "{TAB}"
wshShell.SendKeys "^."
WScript.Sleep(1000)
wshShell.SendKeys "{enter}"
WScript.Sleep(1000)
wshShell.SendKeys "{enter}" 
    
    objOutputFile.Write strNumber & ". - " & arrServiceList(0) & vbcrlf


    For i = 1 to Ubound(arrServiceList) 
	
	next
	
WScript.Sleep(27000)

For count = 0 to 3
	wshShell.SendKeys "{TAB}"
	next
	
WScript.Sleep(200)

strNumber = strNumber + 1

Loop 

wscript.echo "All array data has been cycled through."
