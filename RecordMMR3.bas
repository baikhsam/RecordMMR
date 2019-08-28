Attribute VB_Name = "RecordMMR"
Sub RecordMMR()
'Sam Baik
'2/11/2019
'Macro created to record MMR

Application.ScreenUpdating = False

Dim mmr As String
mmr = InputBox("What is your Solo MMR after the match", "Input MMR")
If mmr = "" Then
    Exit Sub
End If
Dim today As Variant
today = Format(Date, "m/dd/yyyy")

Dim i As Integer
i = 2
Do While Cells(i, 1).Value <> ""
    i = i + 1
Loop

Range("A" & i) = today
Range("C" & i) = mmr
Range("D" & i) = "=C" & i & "-C" & (i - 1)
Range("E" & i) = "=(5000-" & "C" & i & ")/25"
Range("F" & i) = "=5000-" & "C" & i
Range("G" & i) = "=(4600-" & "C" & i & ")/25"
Range("H" & i) = "=4600-" & "C" & i

If Range("D" & i) > 0 Then
    Range("D" & i).Font.ColorIndex = 43
End If
Application.ScreenUpdating = True
End Sub
