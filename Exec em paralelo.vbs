Dim ProgramPath, WshShell, ProgramArgs, Process, ScriptEngine, oFS, ScriptPath, i

Set oFS = CreateObject("Scripting.FileSystemObject")


Set WshShell=CreateObject ("WScript.Shell")
ProgramPath = oFS.GetParentFolderName(WScript.ScriptFullName)
ScriptPath = ProgramPath & "\scripts"
ProgramArgs="Hello"
ScriptEngine="CScript.exe"

redim process(oFS.GetFolder(ScriptPath).Files.Count)

For Each file in oFS.GetFolder(ScriptPath).Files
	Set process(i)=WshShell.Exec (ScriptEngine & space (1) & Chr(34) & file & Chr (34) & Space (1) & ProgramArgs)
	i=i+1
Next


for i = 0 to UBound(process) -1
	Do While Process(i).Status=0
		'Currently Waiting on the program to finish execution.
		WScript.Sleep 300
	Loop
next

msgbox "Finalizado"