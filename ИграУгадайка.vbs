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
MsgBox "� ������� ����� �� 1 �� 100, ���������?",64,"��������?"
Do
a = a + 1
c = InputBox("������ �����" & vbCrLf & vbCrLf & "�������: " & a & vbCrLf & vbCrLf & "����� ���������� ���������� ������� �����" & vbCrLf & vbCrLf & "��� ������ �������� ���� ������" & vbCrLf & vbCrLf,"��������")
If c = "!" Then CreateObject("WScript.Shell").Run "notepad C:\Windows\Result.dll",3,True : Exit Do
If c <> "" Then
If IsNumeric(c) = True Then
If CInt(c) < Number Then MsgBox "���,��� �� " & c & ". � ������� ����� ������",64,"��������"
If CInt(c) > Number Then MsgBox "���,��� �� " & c & ". � ������� ����� ������",64,"��������"
If CInt(c) = Number Then
Set y = CreateObject("Scripting.FileSystemObject")
MsgBox ("�������: " & a)
If MsgBox ("���������, ��� ���� ����� " & c & ". ������ ������?",36,"��������") = 6 Then Exit Do Else WScript.Quit
End If
Else
MsgBox "��� �� �����!",16,"��������"
a = a = 1
End If
Else
a = a = 1
l = MsgBox ("�� ������ �� ����. �����?",36,"��������")
If l = 6 Then WScript.Quit
End If
Loop
loop