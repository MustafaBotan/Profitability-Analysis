Attribute VB_Name = "Module1"
Option Explicit
Sub startform()

MainForm.Show

End Sub

Sub ApplicationSpeedOptimize()
    Application.ScreenUpdating = False
    'Application.Calculation = xlCalculationManual   ' turn off the automatic calculation
    Application.DisplayStatusBar = False            ' turn off status bar updates
    Application.EnableEvents = False                ' ignore events
    'ActiveSheet.DisplayPageBreaks = False
    Application.DisplayAlerts = False
End Sub

Sub ApplicationRestoreAfterSpeedOptimize()
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.DisplayStatusBar = True
    Application.EnableEvents = True
    'ActiveSheet.DisplayPageBreaks = True
    Application.DisplayAlerts = True
End Sub


Function triangular_inverse(P As Double, L As Double, M As Double, U As Double) As Double
'Given a probability P and lower (L), upper (U), and most common (M) inputs, this
'function calculates the corresponding x value
Dim a As Double, b As Double, c As Double
If P < (M - L) / (U - L) Then
    a = 1
    b = -2 * L
    c = L ^ 2 - P * (M - L) * (U - L)
    triangular_inverse = (-b + Sqr(b ^ 2 - 4 * a * c)) / 2 / a
ElseIf P <= 1 Then
    a = 1
    b = -2 * U
    c = U ^ 2 - (1 - P) * (U - L) * (U - M)
    triangular_inverse = (-b - Sqr(b ^ 2 - 4 * a * c)) / 2 / a
End If
End Function
