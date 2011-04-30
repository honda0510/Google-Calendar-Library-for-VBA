Attribute VB_Name = "Module1"
Option Explicit

Sub test()
    Dim GoogleCalendar As GoogleCalendar
    Dim email As String
    Dim password As String
    Dim xml As String
    
    email = Range("A1").Value
    password = Range("A2").Value
    xml = Range("A3").Value
    
    ' Google にログイン
    Set GoogleCalendar = New GoogleCalendar
    GoogleCalendar.login email, password
        
    ' Google カレンダーに予定を追加
    GoogleCalendar.add xml
    
    MsgBox "OK"
End Sub
