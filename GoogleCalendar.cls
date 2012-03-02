VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GoogleCalendar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ----------------------------------------------------------------------
' 参照設定
'
' Microsoft WinHTTP Services, version 5.1
' Microsoft VBScript Regular Expressions 5.5
' ----------------------------------------------------------------------

' HTTP ステータスコード
Private Const HTTP_STATUS_OK As Long = 200
Private Const HTTP_STATUS_CREATED As Long = 201
Private Const HTTP_STATUS_MOVED_TEMPORARILY As Long = 302

Private Const URI_GOOGLE_LOGIN As String = "https://www.google.com/accounts/ClientLogin"
Private Const URI_GOOGLE_CALENDAR As String = "https://www.google.com/calendar/feeds/default/private/full"

Private http As WinHttp.WinHttpRequest
Private reg As VBScript_RegExp_55.RegExp
Private token As String
Private authorization As String

Private Sub Class_Initialize()
    Set http = New WinHttp.WinHttpRequest
    Set reg = New VBScript_RegExp_55.RegExp
End Sub

Private Sub Class_Terminate()
    Set http = Nothing
    Set reg = Nothing
End Sub

' POST /accounts/ClientLogin HTTP/1.1
' Content-Type: application/x-www-form-urlencoded
'
' Email=Email&Passwd=Passwd&service=service&source=source

Public Sub login(email As String, password As String)
    Dim body As String
    Dim pos As Long
    
    body = "Email=" & email & "&Passwd=" & password & "&service=cl&source=Microsoft-VisualBasic-6.5"
    http.Open "POST", URI_GOOGLE_LOGIN, False
    http.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    httpSend body

    pos = InStr(1, http.ResponseText, "Auth=")
    If pos = 0 Then
        ' todo
    Else
        token = trim2(Mid(http.ResponseText, pos + 5))
        authorization = "GoogleLogin auth=" & token
    End If
End Sub

' POST /calendar/feeds/default/private/full HTTP/1.1
' Authorization: [GoogleLogin auth="ClientLogin_token"|AuthSub token="AuthSub_token"]
' Content-Type: application/atom+xml
'
' GData

Public Sub add(xml As String)
    Dim uri As String
    
    http.Open "POST", URI_GOOGLE_CALENDAR, False
    
    ' 自動リダイレクトを無効にする
    http.Option(WinHttpRequestOption_EnableRedirects) = False
    
    http.SetRequestHeader "Authorization", authorization
    http.SetRequestHeader "Content-Type", "application/atom+xml"
    http.Send CVar(xml)

    If http.Status = HTTP_STATUS_CREATED Then
        ' OK
    ElseIf http.Status = HTTP_STATUS_MOVED_TEMPORARILY Then
        uri = redirectURI(http.ResponseText)
        http.Open "POST", uri, False
        http.SetRequestHeader "Authorization", authorization
        http.SetRequestHeader "Content-Type", "application/atom+xml"
        httpSend xml, HTTP_STATUS_CREATED
        ' OK
    Else
        Err.Raise http.Status, , http.StatusText
    End If
End Sub

' <HTML>
' <HEAD>
' <TITLE>Moved Temporarily</TITLE>
' </HEAD>
' <BODY BGCOLOR="#FFFFFF" TEXT="#000000">
' <H1>Moved Temporarily</H1>
' The document has moved <A HREF="http://www.google.com/calendar/feeds/default/private/full?gsessionid=your_gsessionid">here</A>.
' </BODY>
' </HTML>

Private Function redirectURI(html As String) As String
    Dim matches As VBScript_RegExp_55.MatchCollection
    
    reg.Global = False
    reg.IgnoreCase = True
    reg.Pattern = "href=""(.+?)"""
    Set matches = reg.Execute(html)
    redirectURI = matches(0).SubMatches(0)
End Function

Private Function trim2(text As String) As String
    reg.Pattern = "^\s+|\s+$"
    trim2 = reg.Replace(text, "")
End Function

Private Sub httpSend(Optional data, Optional ok_code As Long = HTTP_STATUS_OK)
    If IsMissing(data) Then
        http.Send
    Else
        http.Send data
    End If
    
    If http.Status <> ok_code Then
        Err.Raise http.Status, , http.StatusText
    End If
End Sub
