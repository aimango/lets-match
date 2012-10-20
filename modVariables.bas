Attribute VB_Name = "modFinalProj"
Global sUser As String
Global nBestTime, nBestScore As Integer
Global nTime, nTurnsTaken, nScore As Integer
Global nClickCount, n1st, n2nd As Integer

'Allow forms to load in the centre of the screen
Public Sub CenterForm(CurrForm As Form)
    CurrForm.Top = (Screen.Height - CurrForm.Height) / 2
    CurrForm.Left = (Screen.Width - CurrForm.Width) / 2
End Sub

'Randomize card values
Public Function Random(ByVal Low As Long, ByVal High As Long) As Long
    Randomize
    Random = Int((High - Low + 0.999999) * Rnd + Low)
End Function

