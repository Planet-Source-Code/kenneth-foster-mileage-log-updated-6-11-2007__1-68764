VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form Form1 
   Caption         =   "Daily Business and Personal Mileage Log"
   ClientHeight    =   10065
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14520
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10065
   ScaleWidth      =   14520
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Calculate Miles Per Gallon"
      Height          =   600
      Left            =   4245
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   9255
      Width           =   1515
   End
   Begin VB.Timer Timer1 
      Interval        =   250
      Left            =   4050
      Top             =   8250
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Show Print Form"
      Height          =   600
      Left            =   6630
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   9255
      Width           =   1800
   End
   Begin VB.CommandButton cmdEraseAll 
      BackColor       =   &H00C0FFC0&
      Caption         =   "          Erase all          ( Start New Month)"
      Height          =   585
      Left            =   11400
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   9270
      Width           =   1650
   End
   Begin VB.CommandButton cmdEditUpdate 
      BackColor       =   &H008080FF&
      Caption         =   "Update"
      Height          =   600
      Left            =   9255
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   9255
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.TextBox txtTripMiles 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   7050
      Locked          =   -1  'True
      TabIndex        =   15
      Top             =   8865
      Width           =   1530
   End
   Begin VB.TextBox txtMonYear 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   810
      TabIndex        =   13
      Top             =   9675
      Width           =   1830
   End
   Begin VB.TextBox txtDate 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   0
      TabIndex        =   0
      Top             =   8865
      Width           =   1500
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H000000FF&
      Caption         =   "Delete Selected Item"
      Height          =   600
      Left            =   13170
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   9255
      Width           =   1260
   End
   Begin VB.TextBox txtNotes 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   11610
      TabIndex        =   6
      Top             =   8865
      Width           =   2895
   End
   Begin VB.TextBox txtPM 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   10095
      TabIndex        =   5
      Top             =   8865
      Width           =   1515
   End
   Begin VB.TextBox txtBM 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   8580
      TabIndex        =   4
      Top             =   8865
      Width           =   1515
   End
   Begin VB.TextBox txtStop 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   5505
      TabIndex        =   3
      Top             =   8865
      Width           =   1545
   End
   Begin VB.TextBox txtStart 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   4005
      TabIndex        =   2
      Top             =   8865
      Width           =   1500
   End
   Begin VB.TextBox txtBP 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   1500
      TabIndex        =   1
      Top             =   8865
      Width           =   2505
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   8175
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   14520
      _ExtentX        =   25612
      _ExtentY        =   14420
      _Version        =   393216
      Rows            =   34
      Cols            =   8
      FixedCols       =   0
      BackColorSel    =   16761024
      FocusRect       =   0
      SelectionMode   =   1
      BorderStyle     =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label15 
      Caption         =   "Notes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   11895
      TabIndex        =   28
      Top             =   8565
      Width           =   570
   End
   Begin VB.Label Label14 
      Caption         =   "Personel Miles"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   10140
      TabIndex        =   27
      Top             =   8565
      Width           =   1350
   End
   Begin VB.Label Label13 
      Caption         =   "Business Miles"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   8625
      TabIndex        =   26
      Top             =   8565
      Width           =   1290
   End
   Begin VB.Label Label12 
      Caption         =   "Miles this trip"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   7170
      TabIndex        =   25
      Top             =   8565
      Width           =   1350
   End
   Begin VB.Label Label11 
      Caption         =   "Ending"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   5970
      TabIndex        =   24
      Top             =   8565
      Width           =   705
   End
   Begin VB.Label Label10 
      Caption         =   "Start"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   4560
      TabIndex        =   23
      Top             =   8565
      Width           =   465
   End
   Begin VB.Label Label9 
      Caption         =   "Business Purpose"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1995
      TabIndex        =   22
      Top             =   8565
      Width           =   1590
   End
   Begin VB.Label Label8 
      Caption         =   "Date"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   495
      TabIndex        =   21
      Top             =   8565
      Width           =   495
   End
   Begin VB.Label Label7 
      Caption         =   "Enter data here. ( Use Enter bar only)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   45
      TabIndex        =   18
      Top             =   8220
      Width           =   3870
   End
   Begin VB.Label Label6 
      Caption         =   "Edit Mode :Double click on the line                     you want to edit."
      ForeColor       =   &H000000FF&
      Height          =   405
      Left            =   11895
      TabIndex        =   17
      Top             =   8235
      Width           =   2580
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   420
      Left            =   0
      Top             =   8805
      Width           =   14535
   End
   Begin VB.Label Label1 
      Caption         =   "Current Date...retype to change."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   150
      TabIndex        =   14
      Top             =   9360
      Width           =   3360
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "  Monthly Totals"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   5025
      TabIndex        =   11
      Top             =   8175
      Width           =   1980
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   10020
      TabIndex        =   10
      Top             =   8175
      Width           =   1515
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   8505
      TabIndex        =   9
      Top             =   8175
      Width           =   1515
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   7005
      TabIndex        =   8
      Top             =   8175
      Width           =   1500
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'*******************************************************************
'**                                Mileage Log
'**                               Version 1.0.2
'**                               By Ken Foster
'**                                 June 2007
'**                     Freeware--- no copyrights claimed
'*******************************************************************
'Credits to  Jonas Wolz (jwolzvb@yahoo.de) for the clsTablePrint
'=============================================

Dim gdArray(271)
Dim EditMode As Boolean
Dim fntOld As StdFont
Dim xxx As Integer     ' used in timer1
Dim yyy As Integer     ' used in timer1

Private Sub Command1_Click()
   Form3.Show
End Sub

Private Sub Form_Load()
    GridIni
    LoadGrid
    UpdateGrid
    txtMonYear.Text = Format(Date, "mmmm  yyyy")
    MSFlexGridColors Grid1, 192, 255, 192      'aternate row colors
    Me.Show                                                 'make sure form is visible
    FormatFlexGrid Grid1         'center data in cells
    NextEmptyLine                  'select first empty line and highlight it
    txtDate.SetFocus               'focus on date window
End Sub

Private Sub cmdClear_Click()
    Dim x As Integer
    ClearGridRow
    SaveGrid
    For x = 0 To 271
        gdArray(x) = ""
    Next x
    LoadGrid
    txtDate.SetFocus
End Sub

Private Sub cmdEraseAll_Click()
Dim x As Integer
Dim iresp As String

    iresp = MsgBox("Are you sure?", vbYesNo, "This will erase everything")
    If iresp = vbNo Then Exit Sub
    For x = 0 To 271
       gdArray(x) = ""
    Next x
    Grid1.Clear
    Label2.Caption = ""
    Label3.Caption = ""
    Label4.Caption = ""
    SaveGrid
    GridIni
    LoadGrid
    MSFlexGridColors Grid1, 192, 255, 192      'aternate row colors
    NextEmptyLine
End Sub

Private Sub cmdEditUpdate_Click()
        ' update data to grid
        Grid1.TextMatrix(Grid1.Row, 0) = txtDate.Text
        Grid1.TextMatrix(Grid1.Row, 1) = txtBP.Text
        Grid1.TextMatrix(Grid1.Row, 2) = txtStart.Text
        Grid1.TextMatrix(Grid1.Row, 3) = txtStop.Text
        Grid1.TextMatrix(Grid1.Row, 4) = txtTripMiles.Text
        Grid1.TextMatrix(Grid1.Row, 5) = txtBM.Text
        Grid1.TextMatrix(Grid1.Row, 6) = txtPM.Text
        If txtNotes.Text = "" Then txtNotes.Text = "         "   ' must have something to save here or things get messed up
        Grid1.TextMatrix(Grid1.Row, 7) = txtNotes.Text
        
        'setup for next entry
        ClearTextBoxes
        
        UpdateGrid
        SaveGrid
        LoadGrid
        NextEmptyLine
        'put back into normal mode
        Shape1.FillColor = &HFFC0C0
        Grid1.BackColorSel = &HFFC0C0
        
        cmdEditUpdate.Visible = False
        cmdClear.Enabled = True
        EditMode = False
      
End Sub

Private Sub cmdPrint_Click()
   Form2.Show
   If MsgBox("Is the date correct ?", vbYesNo, "Date Correct ?") = vbNo Then
       NextEmptyLine
       txtMonYear.SelStart = Len(txtMonYear.Text)   ' put blinker at end of text
       txtMonYear.SetFocus
       Unload Form2
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Unload Form2
   Unload Me
End Sub

Private Sub Grid1_Click()
    'put back into normal mode
        Shape1.FillColor = &HFFC0C0
        Grid1.BackColorSel = &HFFC0C0
        cmdEditUpdate.Visible = False
        cmdClear.Enabled = True
        cmdPrint.Enabled = True
        cmdEraseAll.Enabled = True
        EditMode = False
        'clear everything for next entry
        ClearTextBoxes
        txtDate.SetFocus
End Sub

Private Sub Grid1_DblClick()
    If Grid1.TextMatrix(Grid1.Row, 0) = "" Then Exit Sub
    ' you are in edit mode
    cmdEditUpdate.Visible = True
    cmdClear.Enabled = False
    cmdPrint.Enabled = False
    cmdEraseAll.Enabled = False
    EditMode = True
    With Grid1
       Shape1.FillColor = vbRed
       .BackColorSel = vbRed
       
       ' load textboxs for editing
       txtDate.Text = .TextMatrix(.Row, 0)
       txtBP.Text = .TextMatrix(.Row, 1)
       txtStart.Text = .TextMatrix(.Row, 2)
       txtStop.Text = .TextMatrix(.Row, 3)
       txtTripMiles.Text = .TextMatrix(.Row, 4)
       txtBM.Text = .TextMatrix(.Row, 5)
       txtPM.Text = .TextMatrix(.Row, 6)
       txtNotes.Text = .TextMatrix(.Row, 7)
    End With
End Sub

Private Sub Timer1_Timer()
   
      If xxx < 3 Then
         Label7.BackColor = vbRed
         xxx = xxx + 1
      Else
         Label7.BackColor = Form1.BackColor
         yyy = yyy + 1
         xxx = 0
      End If
      If yyy = 3 Then Timer1.Enabled = False
End Sub

Private Sub txtBM_Click()
    If EditMode = True Then
       txtBM.Text = ""
       txtPM.Text = ""
    End If
End Sub

Private Sub txtBM_GotFocus()
   txtBM.BackColor = &HFFF0F0
End Sub

Private Sub txtBM_LostFocus()
   txtBM.BackColor = vbWhite
End Sub

Private Sub txtBM_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
       If txtBM.Text = "" Then
          txtBM.Text = "     "                                                     'blank fill - keeps things neat and orderly
          txtPM.Text = txtTripMiles.Text
          txtNotes.SetFocus
       Else
         txtPM.Text = Val(txtTripMiles.Text) - Val(txtBM.Text)
         If txtBM.Text = txtTripMiles.Text Then txtPM.Text = "      "   'blank fill
         txtNotes.SetFocus
       End If
    End If
    If InStr(1, "8  48 49 50 51 52 53 54 55 56 57", CStr(KeyAscii)) = 0 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtBP_GotFocus()
   txtBP.BackColor = &HFFF0F0
End Sub

Private Sub txtBP_LostFocus()
   txtBP.BackColor = vbWhite
End Sub

Private Sub txtBP_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtStart.SetFocus
End Sub

Private Sub txtDate_GotFocus()
   txtDate.BackColor = &HFFF0F0
End Sub

Private Sub txtDate_LostFocus()
   txtDate.BackColor = vbWhite
End Sub

Private Sub txtDate_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtBP.SetFocus
End Sub

Private Sub txtNotes_GotFocus()
   txtNotes.BackColor = &HFFF0F0
End Sub

Private Sub txtNotes_LostFocus()
   txtNotes.BackColor = vbWhite
End Sub

Private Sub txtNotes_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    'put data in the grid
        Grid1.TextMatrix(Grid1.Row, 0) = txtDate.Text
        Grid1.TextMatrix(Grid1.Row, 1) = txtBP.Text
        Grid1.TextMatrix(Grid1.Row, 2) = txtStart.Text
        Grid1.TextMatrix(Grid1.Row, 3) = txtStop.Text
        Grid1.TextMatrix(Grid1.Row, 4) = txtTripMiles.Text
        Grid1.TextMatrix(Grid1.Row, 5) = txtBM.Text
        Grid1.TextMatrix(Grid1.Row, 6) = txtPM.Text
        If txtNotes.Text = "" Then txtNotes.Text = "         "    ' must have something in here to save or things get messed up
        Grid1.TextMatrix(Grid1.Row, 7) = txtNotes.Text
        
        'setup for next entry
        ClearTextBoxes
        
        UpdateGrid
        SaveGrid
        LoadGrid
        
        ' if last entry was in edit mode then set back to normal
        Shape1.FillColor = &HFFC0C0
        Grid1.BackColorSel = &HFFC0C0
        cmdEditUpdate.Visible = False
        cmdClear.Enabled = True
        NextEmptyLine
    End If
End Sub

Private Sub txtPM_Click()
    If EditMode = True Then
       txtBM.Text = ""
       txtPM.Text = ""
    End If
End Sub

Private Sub txtPM_GotFocus()
   txtPM.BackColor = &HFFF0F0
End Sub

Private Sub txtPM_LostFocus()
   txtPM.BackColor = vbWhite
End Sub

Private Sub txtPM_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtNotes.SetFocus
    If InStr(1, "8  48 49 50 51 52 53 54 55 56 57", CStr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txtStart_Click()
   If EditMode = True Then
       txtStart.Text = ""
       txtStop.Text = ""
       txtTripMiles.Text = ""
       txtBM.Text = ""
       txtPM.Text = ""
    End If
End Sub

Private Sub txtStart_GotFocus()
    txtStart.BackColor = &HFFF0F0
End Sub

Private Sub txtStart_LostFocus()
   txtStart.BackColor = vbWhite
End Sub

Private Sub txtStart_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtStop.SetFocus
    If InStr(1, "8  48 49 50 51 52 53 54 55 56 57", CStr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txtStop_Click()
    If EditMode = True Then
       txtStop.Text = ""
       txtTripMiles.Text = ""
       txtBM.Text = ""
       txtPM.Text = ""
    End If
End Sub

Private Sub txtStop_GotFocus()
   txtStop.BackColor = &HFFF0F0
End Sub

Private Sub txtStop_LostFocus()
   txtStop.BackColor = vbWhite
End Sub

Private Sub txtStop_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       txtTripMiles.Text = Val(txtStop.Text) - Val(txtStart.Text)
       txtBM.SetFocus
    End If
    If InStr(1, "8  48 49 50 51 52 53 54 55 56 57", CStr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub LoadGrid()
    Dim x As Integer
    Dim Y As Integer
    Dim z As Integer
    Dim File As Integer
    Dim instring As String
    Dim FileName As String
    Dim P As Integer
    
    P = 0
    FileName = App.Path & "\" & "DataList.txt"
    On Error GoTo Clnup
    File = FreeFile
    Open FileName For Input As #File ' file opened for reading
    Line Input #File, instring
    'load the array
    gdArray(0) = instring
    While Not EOF(File)
        Line Input #File, instring
        P = P + 1
        gdArray(P) = instring
        Wend
        Close #File
        'load the grid
        For x = 1 To 33
            For Y = 0 To 7
                Grid1.TextMatrix(x, Y) = gdArray(z)
                z = z + 1
            Next Y
        Next x

Clnup:
        Close #File
    End Sub

Private Sub SaveGrid()
    Dim x As Integer
    Dim Y As Integer
    Dim Strg As String
    Dim FileName As String
    Dim ff As Integer
    
    For x = 1 To 33
        For Y = 0 To 7
            If Grid1.TextMatrix(x, Y) = "" Then GoTo here          'don't save empty lines
            Strg = Strg & Grid1.TextMatrix(x, Y) & vbNewLine    ' create the string to save
here:
        Next Y
    Next x
    On Error GoTo Handle
    
    ff = FreeFile
    FileName = App.Path & "\" & "DataList.txt"
    Open FileName For Output As #ff
    Print #ff, Strg
    Close #ff
    Exit Sub
Handle:
    
    MsgBox "Error " & Err.Number & vbCrLf & Err.Description, vbCritical, "Error"
End Sub

Private Sub UpdateGrid()
    Dim x As Integer
    'make sure labels are empty before loading them up
    Label2.Caption = ""
    Label3.Caption = ""
    Label4.Caption = ""
    For x = 1 To 33
        Label2.Caption = Val(Label2.Caption) + Val(Grid1.TextMatrix(x, 4))
        Label3.Caption = Val(Label3.Caption) + Val(Grid1.TextMatrix(x, 5))
        Label4.Caption = Val(Label4.Caption) + Val(Grid1.TextMatrix(x, 6))
    Next x
    If Label2.Caption = "0" Then Label2.Caption = ""
    If Label3.Caption = "0" Then Label3.Caption = ""
    If Label4.Caption = "0" Then Label4.Caption = ""
    
End Sub

Private Sub ClearGridRow()
    Dim x As Integer
    For x = 0 To 7                                      'clear the selected row
        Grid1.TextMatrix(Grid1.Row, x) = ""
    Next x
    UpdateGrid
    SaveGrid
    'clear everything for next entry
    ClearTextBoxes
End Sub

Private Sub MSFlexGridColors(ColorGrid As MSFlexGrid, R As Integer, G As Integer, B As Integer)
' alternates row colors
Dim j As Integer
Dim i As Integer
    For j = 0 To ColorGrid.Cols - 1
        For i = 1 To ColorGrid.Rows - 1
            If i / 2 <> Int(i / 2) Then
                ColorGrid.Col = j
                ColorGrid.Row = i
                ColorGrid.CellBackColor = RGB(R, G, B)
            End If
        Next i
    Next j
End Sub

Public Sub NextEmptyLine()
'looks for first empty line and selects it
Dim x As Integer
    For x = 1 To 33
       If Grid1.TextMatrix(x, 0) = "" Then
         Grid1.Row = x
         Grid1.Col = 0
         Grid1.SetFocus
         Grid1.ColSel = 7
         txtDate.SetFocus
         Exit Sub
       End If
    Next x
End Sub

Public Function FormatFlexGrid(FlexGrid As MSFlexGrid)
'centers data in cells
Dim i As Integer
    For i = 0 To FlexGrid.Cols - 1
        FlexGrid.ColAlignment(i) = 4   ' 4 = flexaligncenter 8 = right 0 = Left
    Next
End Function

Private Sub ClearTextBoxes()
        txtDate.Text = ""
        txtBP.Text = ""
        txtStart.Text = ""
        txtStop.Text = ""
        txtTripMiles.Text = ""
        txtBM.Text = ""
        txtPM.Text = ""
        txtNotes.Text = ""
End Sub

Private Sub GridIni()
    Grid1.ColWidth(0) = 1500
    txtDate.Width = Grid1.ColWidth(0)
    'txtDate.Left = Grid1.ColPos(0)
    Grid1.TextMatrix(0, 0) = " Date "
    Grid1.ColWidth(1) = 2500
    txtBP.Width = Grid1.ColWidth(1)
    txtBP.Left = Grid1.ColPos(1) + 7
    Grid1.TextMatrix(0, 1) = "Business Purpose"
    Grid1.ColWidth(2) = 1500
    txtStart.Width = Grid1.ColWidth(2)
    txtStart.Left = Grid1.ColPos(2)
    Grid1.TextMatrix(0, 2) = "Start"
    Grid1.ColWidth(3) = 1500
    txtStop.Width = Grid1.ColWidth(3)
    txtStop.Left = Grid1.ColPos(3)
    Grid1.TextMatrix(0, 3) = "Ending"
    Grid1.ColWidth(4) = 1500
    txtTripMiles.Width = Grid1.ColWidth(4)
    txtTripMiles.Left = Grid1.ColPos(4)
    Grid1.TextMatrix(0, 4) = "Miles this Trip"
    Grid1.ColWidth(5) = 1500
    txtBM.Width = Grid1.ColWidth(5)
    txtBM.Left = Grid1.ColPos(5)
    Grid1.TextMatrix(0, 5) = "Business Miles"
    Grid1.ColWidth(6) = 1500
    txtPM.Width = Grid1.ColWidth(6)
    txtPM.Left = Grid1.ColPos(6)
    Grid1.TextMatrix(0, 6) = "Personal Miles"
    Grid1.ColWidth(7) = 3000
    txtNotes.Width = Grid1.ColWidth(7)
    txtNotes.Left = Grid1.ColPos(7)
    Grid1.TextMatrix(0, 7) = "Notes"
    Shape1.Width = Form1.Width
    Label2.Left = Grid1.ColPos(4)
    Label2.Width = Grid1.ColWidth(4)
    Label3.Left = Grid1.ColPos(5)
    Label3.Width = Grid1.ColWidth(5)
    Label4.Left = Grid1.ColPos(6)
    Label4.Width = Grid1.ColWidth(6)
    Label5.Left = Label2.Left - Label5.Width
End Sub

Public Sub ImportFlexGrid(clsTP As clsTablePrint, flxGrd As MSFlexGrid, Optional ByVal sngDesiredWidth As Single = -1)
    Dim lRow As Long, lCol As Long
    Dim sngFXGGesWidth As Single
    
    clsTP.Rows = flxGrd.Rows - flxGrd.FixedRows
    clsTP.Cols = flxGrd.Cols
    clsTP.HeaderRows = flxGrd.FixedRows
    clsTP.HasFooter = False
    clsTP.LineThickness = flxGrd.GridLineWidth
    'Use double line width
    clsTP.HeaderLineThickness = 2 * clsTP.LineThickness

    'Set the row height
    clsTP.RowHeightMin = flxGrd.RowHeightMin
    clsTP.FooterRowHeightMin = clsTP.RowHeightMin
    clsTP.HeaderRowHeightMin = clsTP.RowHeightMin
    
    'Use some reasonable default values:
    clsTP.CellXOffset = 10
    clsTP.CellYOffset = 30
    clsTP.CenterMergedHeader = False
    clsTP.ResizeCellsToPicHeight = True
    clsTP.PrintHeaderOnEveryPage = True
    
    Set fntOld = New StdFont
    With flxGrd
        sngFXGGesWidth = 0
        For lRow = 0 To .FixedRows - 1
            For lCol = 0 To .Cols - 1
                .Col = lCol
                .Row = lRow '+ .FixedRows
                Set clsTP.HeaderFont(lRow, lCol) = GetGridFont(flxGrd)
                If (lRow = 0) Then
                    Select Case .FixedAlignment(lCol) '.CellAlignment
                    Case flexAlignLeftTop, flexAlignLeftBottom, flexAlignLeftCenter
                        clsTP.ColAlignment(lCol) = eLeft
                    Case flexAlignRightTop, flexAlignRightBottom, flexAlignRightCenter
                        clsTP.ColAlignment(lCol) = eRight
                    Case flexAlignCenterTop, flexAlignCenterBottom, flexAlignCenterCenter
                        clsTP.ColAlignment(lCol) = eCenter
                    Case flexAlignGeneral 'Always Left here
                        clsTP.ColAlignment(lCol) = eLeft
                    End Select
                    sngFXGGesWidth = sngFXGGesWidth + .ColWidth(lCol)
                End If
                clsTP.HeaderText(lRow, lCol) = .Text
            Next
            clsTP.MergeHeaderRow(lRow) = .MergeRow(lRow)
        Next
        For lCol = 0 To .Cols - 1
            For lRow = 0 To .Rows - .FixedRows - 1
                .Col = lCol
                .Row = lRow + .FixedRows
                Set clsTP.FontMatrix(lRow, lCol) = GetGridFont(flxGrd)
                If Not (.CellPicture Is Nothing) Then
                    If .CellPicture.Handle <> 0 Then
                        Set clsTP.PictureMatrix(lRow, lCol) = .CellPicture
                    End If
                End If
                clsTP.TextMatrix(lRow, lCol) = .Text
                If (lCol = 0) Then
                    clsTP.MergeRow(lRow) = .MergeRow(lRow)
                End If
            Next
            If sngDesiredWidth > 0 Then
                clsTP.ColWidth(lCol) = (.ColWidth(lCol) / sngFXGGesWidth) * sngDesiredWidth
            Else
                clsTP.ColWidth(lCol) = .ColWidth(lCol)
            End If
            clsTP.MergeCol(lCol) = .MergeCol(lCol)
            clsTP.MergeHeaderCol(lCol) = .MergeCol(lCol)
        Next
    End With
End Sub

'Helper Function for ImportFlexGrid()
Private Function GetGridFont(flxGrd As MSFlexGrid) As StdFont
    Dim bDiff As Boolean
    
    If fntOld Is Nothing Then bDiff = True: GoTo DiffCheck
    'Font styles:
    bDiff = bDiff Or (flxGrd.CellFontBold <> fntOld.Bold) Or _
            (flxGrd.CellFontItalic <> fntOld.Italic) Or (flxGrd.CellFontUnderline <> fntOld.Underline) Or _
            (flxGrd.CellFontStrikeThrough <> fntOld.Strikethrough)
    'Name:
    bDiff = bDiff Or (flxGrd.CellFontName <> fntOld.Name)
    'Size:
    bDiff = bDiff Or (flxGrd.CellFontSize <> fntOld.Size)
DiffCheck:
    If bDiff Then
        Set fntOld = New StdFont
        fntOld.Name = flxGrd.CellFontName
        fntOld.Size = flxGrd.CellFontSize
        fntOld.Bold = flxGrd.CellFontBold
        fntOld.Italic = flxGrd.CellFontItalic
        fntOld.Underline = flxGrd.CellFontUnderline
        fntOld.Strikethrough = flxGrd.CellFontStrikeThrough
    End If
    Set GetGridFont = fntOld
End Function
