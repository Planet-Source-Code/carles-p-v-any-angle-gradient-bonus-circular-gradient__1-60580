VERSION 5.00
Begin VB.Form fTest 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "mGradient (any angle) test"
   ClientHeight    =   6375
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8655
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   425
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   577
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picScroll 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   6660
      ScaleHeight     =   17
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   110
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   4770
      Width           =   1680
      Begin VB.Line lnScroll 
         BorderColor     =   &H000000FF&
         X1              =   2
         X2              =   2
         Y1              =   16
         Y2              =   -1
      End
   End
   Begin VB.TextBox txtAngle 
      Height          =   315
      Left            =   7605
      MaxLength       =   3
      TabIndex        =   3
      Text            =   "0"
      Top             =   570
      Width           =   735
   End
   Begin VB.CommandButton cmdPaint 
      Caption         =   "&Paint"
      Default         =   -1  'True
      Height          =   495
      Left            =   6660
      TabIndex        =   4
      Top             =   1065
      Width           =   1680
   End
   Begin VB.TextBox txtIterations 
      Height          =   315
      Left            =   7605
      MaxLength       =   3
      TabIndex        =   1
      Text            =   "1"
      Top             =   150
      Width           =   735
   End
   Begin VB.Line lnAngle 
      BorderColor     =   &H000000FF&
      X1              =   507
      X2              =   507
      Y1              =   276
      Y2              =   198
   End
   Begin VB.Label lblAngle 
      Caption         =   "Angle"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6645
      TabIndex        =   2
      Top             =   615
      Width           =   1020
   End
   Begin VB.Label lblIterations 
      Caption         =   "Iterations"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6645
      TabIndex        =   0
      Top             =   195
      Width           =   1020
   End
   Begin VB.Label lblTiming 
      Height          =   675
      Left            =   6675
      TabIndex        =   5
      Top             =   1740
      Width           =   1590
   End
End
Attribute VB_Name = "fTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const PI     As Single = 3.14159265358979
Private Const TO_RAD As Single = PI / 180
Private m_oTiming    As New cTiming



Private Sub Form_Load()

    If (App.LogMode <> 1) Then
        Call MsgBox("Absolutely recommended: compile first...")
    End If

    Set Me.Icon = Nothing
    Call Me.Show
    Call VBA.DoEvents
    
    Call picScroll_MouseMove(1, 0, 0, 0)
End Sub

Private Sub Form_Paint()
    
 Const PI As Single = 3.14159265358979
   
   Me.ScaleLeft = -500
   Me.ScaleTop = -250
   Me.Circle (0, 0), 50, vbBlack
   Me.Line (-60, 0)-(60, 0), vbWhite
   Me.CurrentX = Me.CurrentX - 6
   Me.CurrentY = Me.CurrentY - 14
   Me.Print "0º"
   Me.Line (0, -60)-(0, 60), vbWhite
End Sub



Private Sub cmdPaint_Click()
  
  Dim i  As Long
  Dim it As Long
    
    With txtIterations
        If (Not IsNumeric(.Text)) Then
            Call MsgBox("Please, enter a valid 'Iterations' number")
            Call .SetFocus
            .SelStart = 0
            .SelLength = Len(.Text)
            Exit Sub
        End If
        it = Val(.Text)
    End With
    
    Call m_oTiming.Reset
    For i = 1 To it
        Call mGradient.PaintGradient(Me.hDC, 10, 10, 100, 100, RGB(255, 0, 0), RGB(0, 0, 255), Val(txtAngle))
        Call mGradient.PaintGradient(Me.hDC, 115, 10, 300, 100, RGB(255, 0, 0), RGB(0, 0, 255), Val(txtAngle))
        Call mGradient.PaintGradient(Me.hDC, 10, 115, 100, 300, RGB(255, 0, 0), RGB(0, 0, 255), Val(txtAngle))
        Call mGradient.PaintGradient(Me.hDC, 115, 115, 300, 300, RGB(255, 0, 0), RGB(0, 0, 255), Val(txtAngle))
    Next i
    lblTiming = it * 4 & " gradients at " & Val(txtAngle) & "º rendered in " & Format$(m_oTiming.Elapsed / 1000, "0.0000 s") & vbCrLf & vbCrLf
End Sub



Private Sub picScroll_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call picScroll_MouseMove(Button, Shift, x, y)
End Sub

Private Sub picScroll_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  
  Dim lAngle As Long
    
    If (Button) Then
        If (x < 0) Then x = 0
        If (x > picScroll.ScaleWidth - 1) Then x = picScroll.ScaleWidth - 1
        With lnScroll
            .X1 = x
            .X2 = x
        End With
        
        lAngle = (x * 364) \ picScroll.ScaleWidth '364?: only for rounding
        With lnAngle
            .X1 = 0
            .Y1 = 0
            .X2 = .X1 + 60 * Cos((360 - lAngle) * TO_RAD)
            .Y2 = .Y1 + 60 * Sin((360 - lAngle) * TO_RAD)
            Call .Refresh
        End With
    
        If (picScroll.Tag = vbNullString) Then txtAngle.Text = lAngle
        Call cmdPaint_Click
    End If
End Sub

Private Sub txtAngle_Change()
    
  Dim lAngle As Long
    
    If (IsNumeric(txtAngle.Text)) Then
        lAngle = Val(txtAngle.Text)
        lAngle = lAngle Mod 360
        If (lAngle < 0) Then lAngle = 360 + lAngle
        
        picScroll.Tag = "!"
        Call picScroll_MouseMove(1, 0, (lAngle / 364) * picScroll.ScaleWidth, 0) '364?: only for rounding
        picScroll.Tag = vbNullString
    End If
End Sub
