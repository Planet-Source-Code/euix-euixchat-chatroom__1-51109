Attribute VB_Name = "modHTML"
Sub ClearHTML(WebControl As WebBrowser)
WebControl.Navigate "about:blank"
    
Do While WebControl.ReadyState <> READYSTATE_COMPLETE
    DoEvents
Loop

WebControl.Document.body.Style.border = "1pt solid black"
End Sub

Sub AddHTML(WebControl As WebBrowser, HTML As String)
WebControl.Document.body.innerHTML = WebControl.Document.body.innerHTML & "<small>" & HTML & "</small><br>"
WebControl.Document.body.scrolltop = CLng(Len(WebControl.Document.body.innerHTML)) * 1000
End Sub


Function ConvertHex(lngColour As Long) As String
Dim strColour As String
strColour = Hex(lngColour)

    Do While Len(strColour) < 6
        strColour = "0" & strColour
    Loop
    
'Reverse the bgr string pairs to rgb
ConvertHex = "#" & Right$(strColour, 2) & Mid$(strColour, 3, 2) & Left$(strColour, 2)
End Function

