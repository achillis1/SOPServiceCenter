VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmProgram 
   Caption         =   "Standard Offer Program"
   ClientHeight    =   8640
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   17280
   OleObjectBlob   =   "frmProgram.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmProgram"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CommandButton1_Click()
  Dim intRow As Integer, intCol As Integer
  With MSFlexGrid1
    For intRow = 0 To .Rows - 1
      For intCol = 0 To .Cols - 1
        .TextMatrix(intRow, intCol) = CStr(intRow + intCol)
        
      Next intCol
    Next intRow
  End With 'MSFlexGrid1
End Sub

Private Sub cmdServiceCenter_Click()
    Me.Hide
    frmServiceCenter.Show vbModeless
End Sub

Private Sub MSFlexGrid1_Click()
'TextBox1.Text = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, MSFlexGrid1.Col)
'If MSFlexGrid1.Row = 2 And MSFlexGrid1.Col = 3 Then
'    UserForm2.Show vbModeless
'End If
'
'If MSFlexGrid1.Row = 1 And MSFlexGrid1.Col = 1 Then
'    UserForm2.Hide
'End If
End Sub


Private Sub gd_Click()
    If gd.Row > 0 And gd.Col = 1 Then
        Me.Hide
        frmProject.Show vbModeless
    End If
End Sub

Private Sub UserForm_Activate()
    gd.Rows = 100
    gd.Cols = 16

    For i = 0 To gd.Cols - 1
        gd.ColWidth(i) = 1100
    Next i
    gd.ColWidth(0) = 600
    gd.ColWidth(1) = 400
    gd.ColWidth(3) = 2600
    gd.ColWidth(4) = 2600
    
    
    For i = 0 To gd.Cols - 1
        gd.TextMatrix(0, 0) = "Index"
        gd.TextMatrix(0, 1) = ""
        gd.TextMatrix(0, 2) = "eTrackID"
        gd.TextMatrix(0, 3) = "TrakSmartID"
        gd.TextMatrix(0, 4) = "ProjectName"
        gd.TextMatrix(0, 5) = "SponsorName"
        gd.TextMatrix(0, 6) = "ConstructionType"
        gd.TextMatrix(0, 7) = "MeasureType"
        gd.TextMatrix(0, 8) = "PAEstimated$"
        gd.TextMatrix(0, 9) = "PAApproved$"
        gd.TextMatrix(0, 10) = "kW"
        gd.TextMatrix(0, 11) = "kWh"
    Next i
    
    For i = 1 To gd.Rows - 1
        'gd.TextMatrix(i, 1) = "Edit"
        With gd
            .Row = i
            .Col = 1
            .CellAlignment = flexAlignCenterCenter
'            .CellBackColor = vbBlue
            .CellForeColor = vbBlue
            '.CellFontName = "Courier New"
            '.CellFontSize = 12
            .CellFontBold = True
            '.Font.Underline = True
            .Text = "Edit"
        End With
    Next i
    
    lastrow = Worksheets("Data").Range("A" & Rows.Count).End(xlUp).Row
    lastcolumn = Worksheets("Data").Range("A" & Columns.Count).End(xlToLeft).Column
    For i = 1 To lastrow
        For j = 2 To gd.Cols - 1
            gd.TextMatrix(i, j) = Worksheets("Data").Cells(i + 1, j - 1).Value
        Next j
    Next i
'    Dim intRow As Integer, intCol As Integer
'    With MSFlexGrid1
'      For intRow = 0 To .Rows - 1
'        For intCol = 0 To .Cols - 1
'          .TextMatrix(intRow, intCol) = CStr(intRow + intCol)
'
'        Next intCol
'      Next intRow
'    End With 'MSFlexGrid1

End Sub


