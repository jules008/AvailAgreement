VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit


Private Sub BtnAddWeek_Click()

End Sub

Private Sub BtnClear_Click()
    Selection.ClearContents
End Sub

Private Sub BtnExtraWeek_Click()
    Dim NoWeeks As Integer
    
    'get current no weeks
    NoWeeks = Range("noweeks")
    
    ActiveSheet.Unprotect (Range("Password"))
    
    Select Case NoWeeks
        Case 1
            ExtraWeek Range("A17"), NoWeeks
            Range("AJ27").Formula = "=COUNTIF(C18:AX24,1)/2"
            Range("AN27").Formula = "=COUNTIF(C18:AX24,2)/2"
        Case 2
            ExtraWeek Range("A26"), NoWeeks
            Range("AJ37").Formula = "=COUNTIF(C27:AX33,1)/2"
            Range("AN37").Formula = "=COUNTIF(C27:AX33,2)/2"
        Case 3
             ExtraWeek Range("A35"), NoWeeks
            Range("AJ47").Formula = "=COUNTIF(C36:AX42,1)/2"
            Range("AN47").Formula = "=COUNTIF(C36:AX42,2)/2"
        Case 4
             ExtraWeek Range("A44"), NoWeeks
            Range("AJ57").Formula = "=COUNTIF(C45:AX51,1)/2"
            Range("AN57").Formula = "=COUNTIF(C45:AX51,2)/2"
        Case 5
             ExtraWeek Range("A53"), NoWeeks
            Range("AJ67").Formula = "=COUNTIF(C54:AX60,1)/2"
            Range("AN67").Formula = "=COUNTIF(C54:AX60,2)/2"
        Case 6
             ExtraWeek Range("A62"), NoWeeks
            Range("AJ77").Formula = "=COUNTIF(C63:AX69,1)/2"
            Range("AN77").Formula = "=COUNTIF(C63:AX69,2)/2"
        Case 7
             ExtraWeek Range("A71"), NoWeeks
            Range("AJ87").Formula = "=COUNTIF(C72:AX78,1)/2"
            Range("AN87").Formula = "=COUNTIF(C72:AX78,2)/2"
        Case 8
             ExtraWeek Range("A80"), NoWeeks
            Range("AJ97").Formula = "=COUNTIF(C81:AX87,1)/2"
            Range("AN97").Formula = "=COUNTIF(C81:AX87,2)/2"
        Case 9
             ExtraWeek Range("A89"), NoWeeks
            Range("AJ107").Formula = "=COUNTIF(C90:AX96,1)/2"
            Range("AN107").Formula = "=COUNTIF(C90:AX96,2)/2"
        Case 10
             ExtraWeek Range("A98"), NoWeeks
            Range("AJ117").Formula = "=COUNTIF(C99:AX105,1)/2"
            Range("AN117").Formula = "=COUNTIF(C99:AX105,2)/2"
        Case 11
             ExtraWeek Range("A107"), NoWeeks
            Range("AJ127").Formula = "=COUNTIF(C108:AX114,1)/2"
            Range("AN127").Formula = "=COUNTIF(C108:AX114,2)/2"
    
    End Select
    ActiveSheet.Protect (Range("Password"))


End Sub

Private Sub BtnFill1s_Click()
    Dim cell As Range
    
    For Each cell In Selection
        cell.Value = 1
    Next

End Sub

Private Sub BtnFill2s_Click()
    Dim cell As Range
    
    For Each cell In Selection
        cell.Value = 2
    Next

End Sub

Private Sub BtnResetForm_Click()
    Dim Response As Integer
    Dim NoWeeks As Integer
    
    Response = MsgBox("Are you sure you want to clear the sheet", vbYesNo)
    
    If Response = 6 Then
        ActiveSheet.Unprotect (Range("Password"))
        
        'get current no weeks
        NoWeeks = Range("noweeks")
        
        Select Case NoWeeks
            Case 2
                Range("AJ27:AN27").Formula = ""
                Range("A17:A25").EntireRow.Delete
            Case 3
                Range("AJ36:AN37").Formula = ""
                Range("A17:A34").EntireRow.Delete
            Case 4
                Range("AJ45:AN47").Formula = ""
                Range("A17:A43").EntireRow.Delete
            Case 5
                Range("AJ54:AN57").Formula = ""
                Range("A17:A52").EntireRow.Delete
            Case 6
                Range("AJ63:AN67").Formula = ""
                Range("A17:A61").EntireRow.Delete
            Case 7
                Range("AJ72:AN77").Formula = ""
                Range("A17:A70").EntireRow.Delete
            Case 8
                Range("AJ81:AN87").Formula = ""
                Range("A17:A79").EntireRow.Delete
            Case 9
                Range("AJ90:AN97").Formula = ""
                Range("A17:A88").EntireRow.Delete
            Case 10
                Range("AJ99:AN107").Formula = ""
                Range("A17:A97").EntireRow.Delete
            Case 11
                Range("AJ108:AN117").Formula = ""
                Range("A17:A106").EntireRow.Delete
            Case 12
                Range("AJ117:AN127").Formula = ""
                Range("A17:A115").EntireRow.Delete
        End Select
        
        Range("Clear").ClearContents
        Range("noweeks") = 1
        ActiveSheet.PageSetup.PrintArea = "$A$1:$AX$71"
        ActiveSheet.Range("A31").RowHeight = 35
        ActiveSheet.Protect (Range("Password"))
        ActiveSheet.DisplayPageBreaks = False

    End If

End Sub

Private Sub BtnSave_Click()
    ActiveWorkbook.SaveAs ("\\lincsfire.lincolnshire.gov.uk\folderredir$\Documents\julian.turner\Documents\RDS Project\Availability Agreement\Processed F316s\" & Left(Range("station1"), 2) & " - " & Range("serviceno") & " " & Range("name"))

End Sub

Private Sub OptionButton2_Click()

End Sub
Public Sub ExtraWeek(StartCell As Range, WeekNo As Integer)
    
    Range("A8:AX16").Copy
        
    StartCell.Insert
    
    Set StartCell = StartCell.Offset(-9, 0)
    
    StartCell.Select
    
    StartCell.RowHeight = 35
    StartCell.Offset(14, 0).RowHeight = 15
    StartCell.Offset(23, 0).RowHeight = 35
    
    WeekNo = WeekNo + 1
    
    StartCell.Value = "Week " & WeekNo
    
    Range("noweeks") = WeekNo
    
    Range("A1").Select
    
    ActiveSheet.DisplayPageBreaks = False
        
End Sub


