VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form fDemo 
   Caption         =   "Gradient (any angle) Demo"
   ClientHeight    =   6735
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5550
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   449
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   370
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Frame fraOptions 
      Caption         =   "Options"
      Height          =   1935
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   5295
      Begin VB.OptionButton optMethod 
         Caption         =   "API method"
         Height          =   255
         Index           =   1
         Left            =   1440
         TabIndex        =   12
         Top             =   720
         Value           =   -1  'True
         Width           =   1575
      End
      Begin VB.OptionButton optMethod 
         Caption         =   "VB-only method"
         Height          =   255
         Index           =   0
         Left            =   1440
         TabIndex        =   11
         Top             =   360
         Width           =   1695
      End
      Begin VB.CommandButton cmdGradient 
         Caption         =   "one-time gradient"
         Height          =   495
         Left            =   120
         TabIndex        =   4
         Top             =   1320
         Width           =   1815
      End
      Begin VB.TextBox txtAngle 
         Height          =   315
         Left            =   4200
         TabIndex        =   3
         Text            =   "72"
         Top             =   360
         Width           =   615
      End
      Begin VB.CommandButton cmdLoop 
         Caption         =   "continuous loop"
         Height          =   495
         Left            =   3360
         TabIndex        =   2
         Top             =   1320
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "- OR -"
         Height          =   255
         Left            =   2400
         TabIndex        =   13
         Top             =   1440
         Width           =   495
      End
      Begin VB.Line Line2 
         X1              =   3240
         X2              =   3240
         Y1              =   240
         Y2              =   1080
      End
      Begin VB.Line Line1 
         X1              =   1320
         X2              =   1320
         Y1              =   240
         Y2              =   1080
      End
      Begin VB.Label lblCap 
         Alignment       =   1  'Right Justify
         Caption         =   "Angle:"
         Height          =   255
         Index           =   0
         Left            =   3360
         TabIndex        =   5
         Top             =   360
         Width           =   615
      End
      Begin VB.Label lblTime 
         Caption         =   "Time:"
         Height          =   255
         Left            =   3480
         TabIndex        =   10
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label lblCap 
         Alignment       =   1  'Right Justify
         Caption         =   "Color 2: "
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   9
         Top             =   720
         Width           =   735
      End
      Begin VB.Label lblColor 
         BackColor       =   &H00FF80FF&
         BorderStyle     =   1  'Fixed Single
         Height          =   270
         Index           =   1
         Left            =   840
         TabIndex        =   8
         Top             =   720
         Width           =   360
      End
      Begin VB.Label lblCap 
         Alignment       =   1  'Right Justify
         Caption         =   "Color 1: "
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   735
      End
      Begin VB.Label lblColor 
         BackColor       =   &H0080FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   270
         Index           =   0
         Left            =   840
         TabIndex        =   6
         Top             =   360
         Width           =   360
      End
   End
   Begin VB.PictureBox pCanvas 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   6
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4455
      Left            =   120
      ScaleHeight     =   293
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   349
      TabIndex        =   0
      Top             =   2160
      Width           =   5295
   End
   Begin MSComDlg.CommonDialog cdlColor 
      Left            =   3480
      Top             =   360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "fDemo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' pGradientDemo.vbp
' 2005 February 19
' redbird77@earthlink.net
' http://home.earthlink.net/~redbird77

' Ho hum, another gradient... but it's pretty fast and supports angles.
' I've got some major snippage of a version that supports multiple colors
' (each with their own position and transparency).  I'm in the process of
' cleaning it all up.  It's pretty fast too since it uses cDIB from Carles PV
' and the nifty Bresenham line drawing algorithm.
'
' If you want any snippage, please email me.

Option Explicit

' Not needed for gradient, just for continuous gradient demo.
Private m_bRun              As Boolean

' Not needed for gradient, just for sizeable picturebox demo.
Private Const GWL_STYLE     As Long = -16
Private Const SWP_DRAWFRAME As Long = &H20
Private Const SWP_NOMOVE    As Long = &H2
Private Const SWP_NOSIZE    As Long = &H1
Private Const SWP_NOZORDER  As Long = &H4
Private Const WS_THICKFRAME As Long = &H40000

Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Sub SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal Y As Long, ByVal CX As Long, ByVal CY As Long, ByVal wFlags As Long)

Private Sub cmdGradient_Click()

' Draw a single gradient.

Dim t As Single

    t = Timer
    
    If optMethod(0).Value Then
    
        DrawGradientVB Me, lblColor(0).BackColor, lblColor(1).BackColor, CSng(txtAngle.Text)
        
    Else
    
        DrawGradient Me.hDC, Me.ScaleWidth, Me.ScaleHeight, _
                     lblColor(0).BackColor, lblColor(1).BackColor, CSng(txtAngle.Text)
                 
    End If
       
    lblTime.Caption = "Time: " & Format$(Timer - t, "0.000")
    
End Sub

Private Sub cmdLoop_Click()

' Continuously draw gradients at incrementing angles.

Dim t As Single

    m_bRun = Not m_bRun
    
    If cmdLoop.Caption = "stop" Then
        cmdLoop.Caption = "continuous loop"
    Else
        cmdLoop.Caption = "stop"
    End If

    Do While m_bRun
    
        t = Timer
        
        ' Yes, this demo could be improved by caching the color values, width,
        ' height, method, etc.
        
        If optMethod(0).Value Then
            DrawGradientVB pCanvas, lblColor(0).BackColor, lblColor(1).BackColor, CSng(txtAngle.Text)
        Else
            DrawGradient pCanvas.hDC, pCanvas.ScaleWidth, pCanvas.ScaleHeight, _
                     lblColor(0).BackColor, lblColor(1).BackColor, CSng(txtAngle.Text)
        End If
        
        txtAngle.Text = (CSng(txtAngle.Text) + 1) Mod 360
        
        lblTime.Caption = "Time: " & Format$(Timer - t, "0.000")
        
        DoEvents
        
    Loop
    
End Sub

Private Sub Form_Load()

Dim lOld As Long

    ' Quick and easy way to handle user clicking "Cancel" in the color dialog.
    cdlColor.CancelError = True
    
    ' Make picturebox canvas sizeable.
    With pCanvas
        lOld = GetWindowLong(.hWnd, GWL_STYLE)
        lOld = SetWindowLong(.hWnd, GWL_STYLE, lOld Or WS_THICKFRAME)
        SetWindowPos .hWnd, fDemo.hWnd, 0, 0, 0, 0, SWP_NOZORDER Or SWP_NOSIZE Or SWP_NOMOVE Or SWP_DRAWFRAME
    End With
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    m_bRun = False
    
End Sub

Private Sub lblColor_Click(Index As Integer)

' If user clicks cancel, an error will occur, program flow will jump to
' the ErrHandler label thus bypassing setting the label's backcolor.

    On Error GoTo ErrHandler

    cdlColor.ShowColor

    lblColor(Index).BackColor = cdlColor.Color

ErrHandler:

End Sub
