Attribute VB_Name = "HttpModule"
Option Explicit

Private http As Object
Private cache As Object

Function HttpGet(url As String, Optional memoize As Boolean = True) As String
    
    If http Is Nothing Then Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    If cache Is Nothing Then Set cache = CreateObject("Scripting.Dictionary")
     
    Dim text As String
    
    If memoize And cache.Exists(url) Then
        text = cache(url)
    Else
        http.Open "GET", url, True
        http.send
        http.WaitForResponse
            
        text = http.responsetext
        
        If memoize Then
            cache.Add url, text
        End If
    End If
    
    HttpGet = text

End Function
