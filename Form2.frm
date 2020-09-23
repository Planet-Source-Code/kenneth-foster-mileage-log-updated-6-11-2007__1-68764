VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Print Preview"
   ClientHeight    =   9090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12615
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   ScaleHeight     =   9090
   ScaleWidth      =   12615
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Cancel"
      Height          =   390
      Left            =   6495
      TabIndex        =   5
      Top             =   8595
      Width           =   1785
   End
   Begin VB.PictureBox picScroll 
      Height          =   8415
      Left            =   105
      ScaleHeight     =   8355
      ScaleWidth      =   12180
      TabIndex        =   2
      Top             =   105
      Width           =   12240
      Begin VB.VScrollBar vscScroll 
         Height          =   3765
         Left            =   11940
         Max             =   32000
         TabIndex        =   4
         Top             =   15
         Width           =   240
      End
      Begin VB.PictureBox picTarget 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   8355
         Left            =   0
         ScaleHeight     =   8325
         ScaleWidth      =   11865
         TabIndex        =   3
         Top             =   0
         Width           =   11895
      End
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Print "
      Height          =   405
      Left            =   4140
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   8580
      Width           =   1965
   End
   Begin VB.CheckBox chkColWidth 
      Caption         =   "Resize Col widths to fill page"
      Height          =   240
      Left            =   150
      TabIndex        =   0
      Top             =   8670
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   2520
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'# Author: Jonas Wolz (jwolzvb@yahoo.de)    #

Option Explicit

'The dimensions of the DIN A4 paper size in Twips:
Const A4Height = 16840, A4Width = 11907

'To get the scroll width:
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Const SM_CYHSCROLL = 3
Private Const SM_CXVSCROLL = 2

'Declared Private WithEvents to get NewPage event:
Private WithEvents cTP As clsTablePrint
Attribute cTP.VB_VarHelpID = -1

Private Sub InitializePictureBox()
    Dim sngVSCWidth As Single, sngHSCHeight As Single
    'Set the size to the DIN A4 width:
    picTarget.Width = A4Width
    picTarget.Height = A4Height
    'Resize the scrollbars:
    sngVSCWidth = GetSystemMetrics(SM_CXVSCROLL) * Screen.TwipsPerPixelX
    sngHSCHeight = GetSystemMetrics(SM_CYHSCROLL) * Screen.TwipsPerPixelY
    vscScroll.Move picScroll.ScaleWidth - sngVSCWidth, 0, sngVSCWidth, picScroll.ScaleHeight
    
    SetScrollBars
End Sub

Private Sub SetScrollBars()
    vscScroll.Max = (picTarget.Height - picScroll.ScaleHeight + 3000) / 120 + 1
End Sub


Private Sub cmdPrint_Click()
    
    If MsgBox("The application will now print the grid on the default printer.", vbInformation + vbOKCancel, "Print") = vbCancel Then Exit Sub
    
    'Simply initialize the printer:
    Printer.Print
    
    'Read the FlexGrid:
    'Set the wanted width of the table to -1 to get the exact widths of the FlexGrid,
    ' to ScaleWidth - [the left and right margins] to get a fitting table !
    Form1.ImportFlexGrid cTP, Form1.Grid1, IIf((chkColWidth.Value = vbChecked), Printer.ScaleWidth - 2 * 550, -1)
    
    'Set margins (not needed, but looks better !):
    cTP.MarginBottom = 567 '567 equals to 1 cm
    cTP.MarginLeft = 530
    cTP.MarginTop = 2000
    
    'Print the Title and Date
    Printer.CurrentX = 3200
    Printer.FontSize = 12
    Printer.FontBold = True
    Printer.Print "Daily Business and Personel Mileage Log"
    Printer.Print
    Printer.CurrentX = 3200
    Printer.Print "___________________________________"
    Printer.CurrentX = 5000
    Printer.CurrentY = Printer.CurrentY - 300
    Printer.Print Form1.txtMonYear.Text
    Printer.FontBold = False
    Printer.FontSize = 10
    Printer.CurrentX = 5000
    Printer.Print "Month and Year"
    Printer.FontSize = 10
    
    
    'Class begins drawing at CurrentY !
    Printer.CurrentY = cTP.MarginTop
    
    'Finally draw the Grid !
    cTP.DrawTable Printer
    'Done with drawing !
    Printer.CurrentX = 4000: Printer.CurrentY = Printer.CurrentY + 100
    Printer.Print "Monthly Totals:"
    Printer.CurrentX = 5600: Printer.CurrentY = Printer.CurrentY - 230
    Printer.Print Form1.Label2.Caption
    Printer.CurrentX = 6700: Printer.CurrentY = Printer.CurrentY - 200
    Printer.Print Form1.Label3.Caption
    Printer.CurrentX = 8000: Printer.CurrentY = Printer.CurrentY - 220
    Printer.Print Form1.Label4.Caption
   
    
    'Say VB it should finally send it:
    Printer.EndDoc
End Sub

Private Sub cmdRefresh_Click()
    
    'Read the FlexGrid:
    'Set the wanted width of the table to -1 to get the exact widths of the FlexGrid,
    ' to ScaleWidth - [the left and right margins] to get a fitting table !
    Form1.ImportFlexGrid cTP, Form1.Grid1, IIf((chkColWidth.Value = vbChecked), picTarget.ScaleWidth - 2 * 530, -1)
    
    'Set margins (not needed, but looks better !):
    cTP.MarginBottom = 567 '567 equals to 1 cm
    cTP.MarginLeft = 530
    cTP.MarginTop = 2000
    
    'Clear the box:
    picTarget.Cls
    
    'Class begins drawing at CurrentY !
    picTarget.CurrentY = cTP.MarginTop
    
    'Finally draw the Grid !
    cTP.DrawTable picTarget
    'Done with drawing !
End Sub

Private Sub Command1_Click()
   Form1.NextEmptyLine
   'Form2.Hide
   Unload Me
End Sub

Private Sub cTP_NewPage(objOutput As Object, TopMarginAlreadySet As Boolean, bCancel As Boolean, ByVal lLastPrintedRow As Long)
    
    'The class wants a new page, look what to do
    If TypeOf objOutput Is Printer Then
        Printer.NewPage
    Else 'We are printing on the PictureBox !
        objOutput.CurrentY = objOutput.ScaleHeight
        'Simply increase the height of the PicBox here
        ' (very simple, but looks bad in "real" applications)
        objOutput.Height = objOutput.Height + A4Height
        'Draw a line to show the new page:
        objOutput.Line (0, objOutput.CurrentY)-(objOutput.ScaleWidth, objOutput.CurrentY), &H808080
        
        'Set the CurrentY to the position the class should continie with drawing and...
        objOutput.CurrentY = objOutput.CurrentY + cTP.MarginTop
        '... tell it to do so:
        TopMarginAlreadySet = True
        
        'Set the ScrollBar Max properties:
        SetScrollBars
    End If
End Sub

Private Sub Form_Load()
    InitializePictureBox
    'FillFlexGrid
    Set cTP = New clsTablePrint
    cmdRefresh_Click
End Sub

Private Sub vscScroll_Change()
    picTarget.Top = -CSng(vscScroll.Value) * 120
End Sub

Private Sub vscScroll_Scroll()
    vscScroll_Change
End Sub

