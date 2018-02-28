Option Explicit

Dim exeFile,pgmFile,returnMessage
Dim ShellObj,ShellExec

' 引数の処理
if WScript.Arguments.Count = 0 then
    returnMessage = -1
    WScript.StdOut.Writeline returnMessage ' Run Scriptに戻り値を渡す
    WScript.Quit(returnMessage)
end If

exeFile = WScript.Arguments(0) 'ex:python.exe
pgmFile = WScript.Arguments(1) 'ex:hello.py

Set ShellObj = WScript.CreateObject("WScript.Shell")
Set ShellExec = ShellObj.Exec(exeFile + " "+ pgmFile)

Do Until ShellExec.StdOut.AtEndOfStream '最後までループ
    returnMessage = returnMessage & ShellExec.StdOut.ReadLine 'pythonのprint出力行を読み取る
Loop

WScript.StdOut.Writeline returnMessage ' Run Scriptに戻り値を渡す
WScript.Quit(0)
