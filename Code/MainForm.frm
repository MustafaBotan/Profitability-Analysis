VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MainForm 
   Caption         =   "Monte Carlo Simulation"
   ClientHeight    =   8268.001
   ClientLeft      =   96
   ClientTop       =   408
   ClientWidth     =   9840.001
   OleObjectBlob   =   "MainForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Option Base 1

Private Sub GoButton_Click()

'   THE FOLLOWING DIM STATEMENTS ARE FOR VARIABLES INVOLVED
'   IN THE HISTOGRAM PLOTTING CODE THAT I PROVIDE - DON'T DELETE!

Dim datamin As Double, datamax As Double, datarange As Double
Dim lowbins As Integer, highbins As Integer, nbins As Double
Dim binrangeinit As Double, binrangefinal As Double
Dim bins() As Double, bincenters() As Double, j As Integer
Dim c As Integer, i As Integer, R() As Double
Dim bincounts() As Integer, ChartRange As String, nr As Integer
Dim rn As Double
Dim beta As Double
Dim alpha As Double
Dim posNPV As Integer
'---------------------

'   PLACE THE MAIN BULK OF YOUR CODE HERE.
'   IN ORDER FOR THE HISTOGRAM CODE BELOW (THAT I PROVIDE)
'   TO WORK PROPERLY, THE RESULT OF THIS PART OF THE
'   SUBROUTINE SHOULD CREATE A VECTOR R THAT IS COMPOSED OF
'   THE END RESULT (PROFIT PER COOKIE IN THE EXAMPLE I SHOW)
'   FOR EACH OF THE SIMULATIONS.  FOR EXAMPLE, IF 1,000 SIMULATIONS
'   WERE PERFORMED, THE SIZE OF R WOULD BE 1000 (COLUMN VECTOR).

    Dim DistributionN As Double
    Dim DistributionBetaInv As Double
    Dim DistributionU As Double
    Dim wa As Worksheet

    Set wa = ThisWorkbook.Sheets("Main")
    ReDim R(nsimulations)
Call ApplicationSpeedOptimize
    Randomize
    
For i = 1 To nsimulations
    'Cost of land
    rn = Rnd()
    If rn <= COLP1 / 100 Then
    Cells(3, 2) = 1 * COLV1
    ElseIf rn <= (COLP1 + COLP2) / 100 Then
    Cells(3, 2) = 1 * COLV2
    Else
    Cells(3, 2) = 1 * COLV3
    End If

    'cost of royalties
    alpha = (4 * CORMode + CORH - 5 * CORL) / (CORH - CORL)
    beta = (5 * CORH - CORL - 4 * CORMode) / (CORH - CORL)
    Cells(4, 2) = -WorksheetFunction.Beta_Inv(Rnd, alpha, beta, -1 * CORL, -1 * CORH)
    'tdc
    Cells(5, 2) = WorksheetFunction.Norm_Inv(Rnd(), TDCAve, TDCStd)
    'wc
    Cells(6, 2) = (WCMin) + ((WCMax) - (WCMin)) * Rnd()
    'sc
    Cells(7, 2) = WorksheetFunction.Norm_Inv(Rnd(), SCAve, SCStd)
    'sr
    alpha = (4 * SRMode + SRHigh - 5 * SRLow) / (SRHigh - SRLow)
    beta = (5 * SRHigh - SRLow - 4 * SRMode) / (SRHigh - SRLow)
    Cells(3, 5) = (WorksheetFunction.Beta_Inv(Rnd, alpha, beta, 1 * SRLow, 1 * SRHigh))
    'pc
    Cells(3, 8) = -triangular_inverse(Rnd(), -1 * PCLow, -1 * PCMode, -1 * PCHigh)
    
    'tax
    If Rnd() <= 1 * TaxP1 / 100 Then
    Cells(4, 5) = 1 * TaxV1
    Else
    Cells(4, 5) = 1 * TaxV2
    End If
    ' interest rate
    Cells(4, 8) = (IRMin) + ((IRMax) - (IRMin)) * Rnd()
    R(i) = Range("n24").Value
Next i

For i = 1 To nsimulations
If R(i) > 0 Then
    posNPV = posNPV + 1
End If
Next
MsgBox (posNPV / nsimulations * 100 & "% of the results have a positive NPV")
'----------------------

'   DO NOT MODIFY THE CODE BELOW!  ONCE A VECTOR R OF RESULTS HAS BEEN CREATED,
'   THE CODE BELOW WILL CREATE A HISTOGRAM.  MAKE SURE NOT TO DELETE OR CHANGE
'   THE NAME OF THE "HISTOGRAM DATA" WORKSHEET IN THIS FILE!

'   The code below creates a histogram of the vector R that you should create above
'   R is a vector of the end result of each simulation (profit per cookie in this case)

datamin = WorksheetFunction.Min(R)
datamax = WorksheetFunction.Max(R)
datarange = datamax - datamin
lowbins = Int(WorksheetFunction.Log(nsimulations, 2)) + 1
highbins = Int(Sqr(nsimulations))
nbins = (lowbins + highbins) / 2
binrangeinit = datarange / nbins
ReDim bins(1) As Double
If binrangeinit < 1 Then
    c = 1
    Do
        If 10 * binrangeinit > 1 Then
            binrangefinal = 10 * binrangeinit Mod 10
            Exit Do
        Else
            binrangeinit = 10 * binrangeinit
            c = c + 1
        End If
    Loop
    binrangefinal = binrangefinal / 10 ^ c
ElseIf binrangeinit < 10 Then
    binrangefinal = binrangeinit Mod 10
Else
    c = 1
    Do
        If binrangeinit / 10 < 10 Then
            binrangefinal = binrangeinit / 10 Mod 10
            Exit Do
        Else
            binrangeinit = binrangeinit / 10
            c = c + 1
        End If
    Loop
    binrangefinal = binrangefinal * 10 ^ c
End If
i = 1
bins(1) = (datamin - ((datamin) - (binrangefinal * Fix(datamin / binrangefinal))))
Do
    i = i + 1
    ReDim Preserve bins(i) As Double
    bins(i) = bins(i - 1) + binrangefinal
Loop Until bins(i) > datamax
nbins = i
ReDim Preserve bincounts(nbins - 1) As Integer
ReDim Preserve bincenters(nbins - 1) As Double
For j = 1 To nbins - 1
    c = 0
    For i = 1 To nsimulations
        If R(i) > bins(j) And R(i) <= bins(j + 1) Then
            c = c + 1
        End If
    Next i
    bincounts(j) = c
    bincenters(j) = (bins(j) + bins(j + 1)) / 2
Next j
Sheets("Histogram Data").Select
Cells.Clear
Range("A1").Select
Range("A1:A" & nbins - 1) = WorksheetFunction.Transpose(bincenters)
Range("B1:B" & nbins - 1) = WorksheetFunction.Transpose(bincounts)
MainForm.Hide
Application.ScreenUpdating = False
Charts("Histogram").Delete
ActiveCell.Range("A1:B1").Select
    Range(Selection, Selection.End(xlDown)).Select
    nr = Selection.Rows.Count
    ChartRange = Selection.Address
    ActiveSheet.Shapes.AddChart2(201, xlColumnClustered).Select
    ActiveChart.SetSourceData Source:=Range("'Histogram Data'!" & ChartRange)
    ActiveChart.ChartTitle.Select
    Selection.Delete
    ActiveChart.PlotArea.Select
    Application.CutCopyMode = False
    ActiveChart.FullSeriesCollection(1).Delete
    Application.CutCopyMode = False
    Application.CutCopyMode = False
    ActiveChart.FullSeriesCollection(1).XValues = "='Histogram Data'!" & "$A$1:$A$" & nr
    ActiveChart.Legend.Select
    Selection.Delete
    ActiveChart.SetElement (msoElementPrimaryCategoryAxisTitleAdjacentToAxis)
    ActiveChart.SetElement (msoElementPrimaryValueAxisTitleAdjacentToAxis)
    Selection.Caption = "Count"
    ActiveChart.Axes(xlCategory).AxisTitle.Select
    Selection.Caption = "Bin Center"
    ActiveChart.ChartArea.Select
    ActiveChart.Location Where:=xlLocationAsNewSheet, Name:="Histogram"
    Unload Me
    
'------------------
'   FEEL FREE TO ADD CODE BELOW THIS POINT, E.G. TO OUTPUT A SUMMARY OF THE RESULTS IN MESSAGE BOX(ES)
Call ApplicationRestoreAfterSpeedOptimize
End Sub

Private Sub Label85_Click()

End Sub

Private Sub QuitButton_Click()
Unload Me
End Sub

Private Sub UserForm_Click()

End Sub
