Attribute VB_Name = "modGlobals"
Global Username As String 'Chosen Username

Global FontIndex As Integer 'Font Face
Global FontColor As Long 'Font Color

Global PMWindow(1 To 100) As frmPrivateMessage
Global PMCount As Integer

Global Emoticons As String
Global EmoteStatus As Boolean

Global IgnoreDB As New Collection 'Ignore user database

Sub SetEmoticons()
'Sets the emoticons variable
    Emoticons = ":) :-) =) :P :-P =P :D :-D =D :( :-( =( ;) ;-) ;D o_O O_o :O :-O =O"
End Sub

Function EmoteFilename(Emoticon As String)
'Returns the filename for a certain emoticon

Dim Emote As String

'Smiley Face
If Emoticon = ":)" Then Emote = "smile.gif"
If Emoticon = ":-)" Then Emote = "smile.gif"
If Emoticon = "=)" Then Emote = "smile.gif"

'Grinning Face
If Emoticon = ":D" Then Emote = "grin.gif"
If Emoticon = ":-D" Then Emote = "grin.gif"
If Emoticon = "=D" Then Emote = "grin.gif"

'Sad Face
If Emoticon = ":(" Then Emote = "sad.gif"
If Emoticon = ":-(" Then Emote = "sad.gif"
If Emoticon = "=(" Then Emote = "sad.gif"

'Winking Face
If Emoticon = ";)" Then Emote = "wink.gif"
If Emoticon = ";-)" Then Emote = "wink.gif"
If Emoticon = ";D" Then Emote = "wink.gif"

'Tounge Face
If Emoticon = ":P" Then Emote = "tounge.gif"
If Emoticon = ":-P" Then Emote = "tounge.gif"
If Emoticon = "=P" Then Emote = "tounge.gif"

'Oh my.. Face
If Emoticon = ":O" Then Emote = "ohmy.gif"
If Emoticon = ":-O" Then Emote = "ohmy.gif"
If Emoticon = "=O" Then Emote = "ohmy.gif"

'Huh? Face
If Emoticon = "o_O" Then Emote = "huh.gif"
If Emoticon = "O_o" Then Emote = "huh.gif"

'Doh! Face
If Emoticon = ">.<" Then Emote = "doh.gif"

'Mad Face
If Emoticon = ">:(" Then Emote = "mad.gif"
If Emoticon = ">:-(" Then Emote = "mad.gif"

EmoteFilename = "<img src=" & Chr(34) & App.Path & "\Emoticons\" & Emote & Chr(34) & ">"
End Function
Sub NewPM(Username As String)
'Creates a new PM Window
PMCount = PMCount + 1 'Increase the window count

    Set PMWindow(PMCount) = New frmPrivateMessage
    PMWindow(PMCount).PMUser = Username
    PMWindow(PMCount).WindowName = LCase$(Username)
    PMWindow(PMCount).Show
End Sub
