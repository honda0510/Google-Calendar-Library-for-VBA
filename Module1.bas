Attribute VB_Name = "Module1"
Option Explicit

Sub test()
    Dim GoogleCalendar As GoogleCalendar
    Dim email As String
    Dim password As String
    Dim xml As String
    
    email = Range("A1")
    password = Range("A2")
    xml = Range("A3").Value
    
    ' Google �Ƀ��O�C��
    Set GoogleCalendar = New GoogleCalendar
    GoogleCalendar.login email, password
        
    ' Google �J�����_�[�ɃC�x���g��ǉ�
    GoogleCalendar.add xml
    
    MsgBox "OK"
End Sub
