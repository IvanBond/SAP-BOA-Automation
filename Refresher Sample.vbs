' sample VBS

'Samples of call VBS from .bat or command prompt
' Refresher.vbs /TargetFilePath:"C:\Temp\Test.xlsm"
' Refresher.vbs /TargetFilePath:"C:\Temp\BOA Automation\Sample with 2 DS.xlsb"

Set NamedArguments = WScript.Arguments.Named

if not NamedArguments.Exists("TargetFilePath") then
	TargetFilePath = "C:\Temp\BOA Automation\Sample with 2 DS.xlsb"
else
	TargetFilePath = NamedArguments("TargetFilePath")
end if

set xlapp = CreateObject("Excel.Application")
xlapp.visible = true

xlapp.workbooks.open TargetFilePath ', true, true

' run macro "Refresh" located in worksheet ControlPanel (ID of worksheet, NOT a name)
if xlapp.Run("ControlPanel.Refresh") = 0 then
	WScript.Echo "Refresh Failed"
end if

xlapp.Quit
' not 100% stable method
' used only for demo purpose
' Better solution can be found here
' https://github.com/IvanBond/Power-Refresh