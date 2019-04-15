REM I'm glad you install it...
REM Good job!
REM Company AbotCode...
REM Enjoy the game :D


REM Script
Set y = CreateObject("Scripting.FileSystemObject")
Set y = Nothing
Do
a = 0
u = 0
Randomize
Number = Int((RND*99)+1)
MsgBox "Я загадал число от 1 до 100, отгадаешь?",64,"Угадаешь?"
Do
a = a + 1
c = InputBox("Угадай число" & vbCrLf & vbCrLf & "Попытка: " & a & vbCrLf & vbCrLf & "Чтобы посмотреть результаты введите число" & vbCrLf & vbCrLf & "Для выхода оставьте поле пустым" & vbCrLf & vbCrLf,"Угадайка")
If c = "!" Then CreateObject("WScript.Shell").Run "notepad C:\Windows\Result.dll",3,True : Exit Do
If c <> "" Then
If IsNumeric(c) = True Then
If CInt(c) < Number Then MsgBox "Нет,это не " & c & ". Я загадал число больше",64,"Угадайка"
If CInt(c) > Number Then MsgBox "Нет,это не " & c & ". Я загадал число меньше",64,"Угадайка"
If CInt(c) = Number Then
Set y = CreateObject("Scripting.FileSystemObject")
MsgBox ("попытки: " & a)
If MsgBox ("Правильно, это было число " & c & ". Начать заного?",36,"Угадайка") = 6 Then Exit Do Else WScript.Quit
End If
Else
MsgBox "Это не число!",16,"Угадайка"
a = a = 1
End If
Else
a = a = 1
l = MsgBox ("Ты ничего не ввел. Выйти?",36,"Угадайка")
If l = 6 Then WScript.Quit
End If
Loop
loop