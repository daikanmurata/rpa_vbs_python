Option Explicit

Dim exeFile,pgmFile,returnMessage
Dim ShellObj,ShellExec

' �����̏���
if WScript.Arguments.Count = 0 then
    returnMessage = -1
    WScript.StdOut.Writeline returnMessage ' Run Script�ɖ߂�l��n��
    WScript.Quit(returnMessage)
end If

exeFile = WScript.Arguments(0) 'ex:python.exe
pgmFile = WScript.Arguments(1) 'ex:hello.py

Set ShellObj = WScript.CreateObject("WScript.Shell")
Set ShellExec = ShellObj.Exec(exeFile + " "+ pgmFile)

Do Until ShellExec.StdOut.AtEndOfStream '�Ō�܂Ń��[�v
    returnMessage = returnMessage & ShellExec.StdOut.ReadLine 'python��print�o�͍s��ǂݎ��
Loop

WScript.StdOut.Writeline returnMessage ' Run Script�ɖ߂�l��n��
WScript.Quit(0)
