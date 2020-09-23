VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "Miles Per Gallon"
   ClientHeight    =   2340
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form3"
   ScaleHeight     =   2340
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   330
      Left            =   3465
      TabIndex        =   10
      Top             =   1725
      Width           =   1110
   End
   Begin VB.TextBox txtMilesTraveled 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   2550
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   990
      Width           =   1575
   End
   Begin VB.TextBox txtMPG 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   2550
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   1740
      Width           =   795
   End
   Begin VB.TextBox txtGalNow 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   2550
      TabIndex        =   2
      Top             =   1365
      Width           =   795
   End
   Begin VB.TextBox txtORNow 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   2550
      TabIndex        =   1
      Top             =   600
      Width           =   1575
   End
   Begin VB.TextBox txtORLast 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   2535
      TabIndex        =   0
      Top             =   135
      Width           =   1605
   End
   Begin VB.Label Label6 
      Caption         =   "Use Enter Bar for all entries, even to Clear."
      ForeColor       =   &H000040C0&
      Height          =   195
      Left            =   120
      TabIndex        =   11
      Top             =   2085
      Width           =   3150
   End
   Begin VB.Label Label5 
      Caption         =   "Miles traveled"
      Height          =   240
      Left            =   1470
      TabIndex        =   8
      Top             =   1035
      Width           =   1080
   End
   Begin VB.Label Label4 
      Caption         =   "Miles Per Gallon"
      Height          =   240
      Left            =   1305
      TabIndex        =   7
      Top             =   1770
      Width           =   1245
   End
   Begin VB.Label Label3 
      Caption         =   "Number of gallons (this fill-up)"
      Height          =   300
      Left            =   390
      TabIndex        =   5
      Top             =   1395
      Width           =   2145
   End
   Begin VB.Label Label2 
      Caption         =   "Odometer reading ( this fill-up)"
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   645
      Width           =   2235
   End
   Begin VB.Label Label1 
      Caption         =   "Odometer reading (last fill-up)"
      Height          =   255
      Left            =   375
      TabIndex        =   3
      Top             =   180
      Width           =   2115
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim NotEmpty As Boolean

Private Sub Cal()
   On Error Resume Next
   txtMPG.Text = Format(txtMilesTraveled.Text / txtGalNow.Text, "###.0")
   cmdClear.SetFocus
   NotEmpty = True
End Sub

Private Sub cmdClear_Click()
   txtORLast.Text = ""
   txtORNow.Text = ""
   txtMilesTraveled.Text = ""
   txtGalNow.Text = ""
   txtMPG.Text = ""
   NotEmpty = False
   txtORLast.SetFocus
End Sub

Private Sub cmdClear_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then cmdClear_Click
End Sub

Private Sub txtGalNow_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Cal
End Sub

Private Sub txtORLast_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 And NotEmpty = False Then txtORNow.SetFocus
   
   If KeyAscii = 13 And NotEmpty = True Then cmdClear_Click
  
End Sub

Private Sub txtORNow_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      txtMilesTraveled.Text = (Val(txtORNow.Text) - Val(txtORLast.Text))
      txtGalNow.SetFocus
   End If
End Sub
