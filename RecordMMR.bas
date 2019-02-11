Attribute VB_Name = "RecordMMR"
Sub RecordMMR()
'Sam Baik
'2/11/2019
'Macro created to record MMR

Application.ScreenUpdating = False

Dim mmr As Long
mmr = InputBox("What is your Solo MMR after the match", "Input MMR")
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

Application.ScreenUpdating = True
End Sub
