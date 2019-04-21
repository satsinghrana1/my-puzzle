VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7500
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9450
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form1.frx":14872
   ScaleHeight     =   7500
   ScaleWidth      =   9450
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   2090
      Left            =   5160
      Picture         =   "Form1.frx":267DF
      ScaleHeight     =   2085
      ScaleWidth      =   4170
      TabIndex        =   32
      Top             =   2925
      Visible         =   0   'False
      Width           =   4175
   End
   Begin VB.Timer Timer3 
      Interval        =   1
      Left            =   360
      Top             =   6960
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   525
      Left            =   1440
      TabIndex        =   31
      Text            =   "filename"
      Top             =   120
      Visible         =   0   'False
      Width           =   6615
   End
   Begin VB.Timer tclose 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   8160
      Top             =   480
   End
   Begin VB.PictureBox Picture1 
      Height          =   5310
      Left            =   3180
      ScaleHeight     =   5250
      ScaleWidth      =   1920
      TabIndex        =   21
      Top             =   1420
      Visible         =   0   'False
      Width           =   1980
      Begin VB.Image Image11 
         Height          =   5250
         Left            =   0
         Top             =   0
         Width           =   1920
      End
   End
   Begin VB.PictureBox won 
      BorderStyle     =   0  'None
      Height          =   2090
      Left            =   1710
      Picture         =   "Form1.frx":2EF55
      ScaleHeight     =   2085
      ScaleWidth      =   4830
      TabIndex        =   22
      Top             =   2640
      Visible         =   0   'False
      Width           =   4835
      Begin VB.CommandButton exit 
         Height          =   375
         Left            =   1200
         Picture         =   "Form1.frx":398F5
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   1680
         Width           =   735
      End
      Begin VB.CommandButton playagain 
         Height          =   735
         Left            =   120
         Picture         =   "Form1.frx":3C11A
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Moves :"
         BeginProperty Font 
            Name            =   "Segoe Print"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   2040
         TabIndex        =   24
         Top             =   1560
         Width           =   3015
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Time :"
         BeginProperty Font 
            Name            =   "Segoe Print"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   2040
         TabIndex        =   23
         Top             =   1200
         Width           =   3015
      End
   End
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   360
      Top             =   3240
   End
   Begin VB.CommandButton Command4 
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "Lucida Sans Typewriter"
         Size            =   8.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   9120
      TabIndex        =   3
      Top             =   7200
      Width           =   255
   End
   Begin VB.Timer Timer5 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   8640
      Top             =   7200
   End
   Begin VB.Timer Timer4 
      Interval        =   1
      Left            =   600
      Top             =   1920
   End
   Begin VB.TextBox stage 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   0
      TabIndex        =   20
      Text            =   "1"
      Top             =   1920
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton start 
      Height          =   495
      Left            =   3600
      Picture         =   "Form1.frx":3F75B
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6920
      Width           =   1095
   End
   Begin VB.Timer gametime 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   8880
      Top             =   4080
   End
   Begin VB.TextBox moves 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8400
      TabIndex        =   19
      Text            =   "0"
      Top             =   5040
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox time 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8400
      TabIndex        =   18
      Text            =   "0"
      Top             =   4080
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.CommandButton done 
      Enabled         =   0   'False
      Height          =   550
      Left            =   7680
      Picture         =   "Form1.frx":42400
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6240
      Width           =   550
   End
   Begin VB.TextBox t4 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   11280
      TabIndex        =   16
      Top             =   720
      Width           =   1455
   End
   Begin VB.TextBox t3 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   9840
      TabIndex        =   15
      Top             =   720
      Width           =   1455
   End
   Begin VB.TextBox t2 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   11280
      TabIndex        =   14
      Top             =   120
      Width           =   1455
   End
   Begin VB.TextBox t1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   9840
      TabIndex        =   13
      Top             =   120
      Width           =   1455
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   8280
      Top             =   6240
   End
   Begin VB.CommandButton loadpic 
      Height          =   375
      Left            =   7420
      Picture         =   "Form1.frx":44ABD
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3360
      Width           =   975
   End
   Begin VB.PictureBox free 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   1800
      Left            =   5040
      ScaleHeight     =   1770
      ScaleWidth      =   1770
      TabIndex        =   12
      Top             =   5040
      Width           =   1800
   End
   Begin VB.PictureBox p8 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   1800
      Left            =   3240
      ScaleHeight     =   1770
      ScaleWidth      =   1770
      TabIndex        =   11
      Top             =   5040
      Width           =   1800
      Begin VB.Image Image8 
         Height          =   5400
         Left            =   -1800
         Picture         =   "Form1.frx":47384
         Stretch         =   -1  'True
         Top             =   -3600
         Width           =   5400
      End
   End
   Begin VB.PictureBox p7 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   1800
      Left            =   1440
      ScaleHeight     =   1770
      ScaleWidth      =   1770
      TabIndex        =   10
      Top             =   5040
      Width           =   1800
      Begin VB.Image Image7 
         Height          =   5400
         Left            =   0
         Picture         =   "Form1.frx":573B1
         Stretch         =   -1  'True
         Top             =   -3600
         Width           =   5400
      End
   End
   Begin VB.PictureBox p6 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   1800
      Left            =   5040
      ScaleHeight     =   1770
      ScaleWidth      =   1770
      TabIndex        =   9
      Top             =   3240
      Width           =   1800
      Begin VB.Image Image6 
         Height          =   5400
         Left            =   -3600
         Picture         =   "Form1.frx":673DE
         Stretch         =   -1  'True
         Top             =   -1800
         Width           =   5400
      End
   End
   Begin VB.PictureBox p5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   1800
      Left            =   3240
      ScaleHeight     =   1770
      ScaleWidth      =   1770
      TabIndex        =   8
      Top             =   3240
      Width           =   1800
      Begin VB.Image Image5 
         Height          =   5400
         Left            =   -1800
         Picture         =   "Form1.frx":7740B
         Stretch         =   -1  'True
         Top             =   -1800
         Width           =   5400
      End
   End
   Begin VB.PictureBox p4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   1800
      Left            =   1440
      ScaleHeight     =   1770
      ScaleWidth      =   1770
      TabIndex        =   7
      Top             =   3240
      Width           =   1800
      Begin VB.Image Image4 
         Height          =   5400
         Left            =   0
         Picture         =   "Form1.frx":87438
         Stretch         =   -1  'True
         Top             =   -1800
         Width           =   5400
      End
   End
   Begin VB.PictureBox p3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   1800
      Left            =   5040
      ScaleHeight     =   1770
      ScaleWidth      =   1770
      TabIndex        =   6
      Top             =   1440
      Width           =   1800
      Begin VB.Image Image3 
         Height          =   5400
         Left            =   -3600
         Picture         =   "Form1.frx":97465
         Stretch         =   -1  'True
         Top             =   0
         Width           =   5400
      End
   End
   Begin VB.PictureBox p2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   1800
      Left            =   3240
      ScaleHeight     =   1770
      ScaleWidth      =   1770
      TabIndex        =   5
      Top             =   1440
      Width           =   1800
      Begin VB.Image Image2 
         Height          =   5400
         Left            =   -1800
         Picture         =   "Form1.frx":A7492
         Stretch         =   -1  'True
         Top             =   0
         Width           =   5400
      End
   End
   Begin VB.PictureBox p1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   1800
      Left            =   1440
      ScaleHeight     =   1770
      ScaleWidth      =   1770
      TabIndex        =   4
      Top             =   1440
      Width           =   1800
      Begin VB.Image Image1 
         Height          =   5400
         Left            =   0
         Picture         =   "Form1.frx":B74BF
         Stretch         =   -1  'True
         Top             =   0
         Width           =   5400
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   8760
      Top             =   840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DefaultExt      =   ".jpg"
      DialogTitle     =   "load image"
      Filter          =   "JPEG( *.jpg ;*.jpeg;*.jpe;*.jfif) GIF(*.gif)"
      InitDir         =   "pictures"
   End
   Begin VB.Image Image12 
      Height          =   1290
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Visible         =   0   'False
      Width           =   1320
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Moves"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Left            =   7455
      TabIndex        =   30
      Top             =   5160
      Width           =   1005
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Time"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Left            =   7515
      TabIndex        =   29
      Top             =   4200
      Width           =   795
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   345
      Left            =   7815
      TabIndex        =   28
      Top             =   5520
      Width           =   225
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   345
      Left            =   7815
      TabIndex        =   27
      Top             =   4560
      Width           =   225
   End
   Begin VB.Image Image10 
      Height          =   240
      Left            =   3240
      Picture         =   "Form1.frx":C74EC
      Top             =   1120
      Width           =   1860
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H000000C0&
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   9140
      TabIndex        =   17
      Top             =   60
      Width           =   255
   End
   Begin VB.Image Image9 
      Height          =   1800
      Left            =   6960
      Picture         =   "Form1.frx":CA820
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   1800
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim r As Integer
Dim l As Integer
Dim u As Integer
Dim d As Integer






Private Sub Command4_Click()
Picture1.Visible = True
Picture2.Visible = True
Timer5.Enabled = True
End Sub

Private Sub done_Click()
p1.Enabled = False
p2.Enabled = False
p3.Enabled = False
p4.Enabled = False
p5.Enabled = False
p6.Enabled = False
p7.Enabled = False
p8.Enabled = False
free.Enabled = False
Timer1.Enabled = True

End Sub


Private Sub exit_Click()
tclose.Enabled = True

End Sub

Private Sub gametime_Timer()
time.Text = time.Text + 1
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
t1.Text = free.Left
t2.Text = free.Top
t3.Text = p1.Left
t4.Text = p1.Top
End Sub
Private Sub Image2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
t1.Text = free.Left
t2.Text = free.Top
t3.Text = p2.Left
t4.Text = p2.Top
End Sub
Private Sub Image3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
t1.Text = free.Left
t2.Text = free.Top
t3.Text = p3.Left
t4.Text = p3.Top
End Sub
Private Sub Image4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
t1.Text = free.Left
t2.Text = free.Top
t3.Text = p4.Left
t4.Text = p4.Top
End Sub
Private Sub Image5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
t1.Text = free.Left
t2.Text = free.Top
t3.Text = p5.Left
t4.Text = p5.Top
End Sub
Private Sub Image6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
t1.Text = free.Left
t2.Text = free.Top
t3.Text = p6.Left
t4.Text = p6.Top
End Sub
Private Sub Image7_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
t1.Text = free.Left
t2.Text = free.Top
t3.Text = p7.Left
t4.Text = p7.Top
End Sub

Private Sub Image8_Click()
''''''''right''''''''
If t1.Text = 3240 Then
If t2.Text = 1440 Then
If t3.Text = 1440 Then
If t4.Text = 1440 Then
p8.Left = p8.Left + 1800
free.Left = free.Left - 1800
moves.Text = moves.Text + 1
End If
End If
End If
End If
If t1.Text = 3240 Then
If t2.Text = 3240 Then
If t3.Text = 1440 Then
If t4.Text = 3240 Then
p8.Left = p8.Left + 1800
free.Left = free.Left - 1800
moves.Text = moves.Text + 1
End If
End If
End If
End If
If t1.Text = 3240 Then
If t2.Text = 5040 Then
If t3.Text = 1440 Then
If t4.Text = 5040 Then
p8.Left = p8.Left + 1800
free.Left = free.Left - 1800
moves.Text = moves.Text + 1
End If
End If
End If
End If
If t1.Text = 5040 Then
If t2.Text = 1440 Then
If t3.Text = 3240 Then
If t4.Text = 1440 Then
p8.Left = p8.Left + 1800
free.Left = free.Left - 1800
moves.Text = moves.Text + 1
End If
End If
End If
End If
If t1.Text = 5040 Then
If t2.Text = 3240 Then
If t3.Text = 3240 Then
If t4.Text = 3240 Then
p8.Left = p8.Left + 1800
free.Left = free.Left - 1800
moves.Text = moves.Text + 1
End If
End If
End If
End If
If t1.Text = 5040 Then
If t2.Text = 5040 Then
If t3.Text = 3240 Then
If t4.Text = 5040 Then
p8.Left = p8.Left + 1800
free.Left = free.Left - 1800
moves.Text = moves.Text + 1
End If
End If
End If
End If
''''''''''right''''''''''

''''''''''left'''''''''''
If t1.Text = 3240 Then
If t2.Text = 5040 Then
If t3.Text = 5040 Then
If t4.Text = 5040 Then
p8.Left = p8.Left - 1800
free.Left = free.Left + 1800
moves.Text = moves.Text + 1
End If
End If
End If
End If
If t1.Text = 3240 Then
If t2.Text = 3240 Then
If t3.Text = 5040 Then
If t4.Text = 3240 Then
p8.Left = p8.Left - 1800
free.Left = free.Left + 1800
moves.Text = moves.Text + 1
End If
End If
End If
End If
If t1.Text = 3240 Then
If t2.Text = 1440 Then
If t3.Text = 5040 Then
If t4.Text = 1440 Then
p8.Left = p8.Left - 1800
free.Left = free.Left + 1800
moves.Text = moves.Text + 1
End If
End If
End If
End If
If t1.Text = 1440 Then
If t2.Text = 5040 Then
If t3.Text = 3240 Then
If t4.Text = 5040 Then
p8.Left = p8.Left - 1800
free.Left = free.Left + 1800
moves.Text = moves.Text + 1
End If
End If
End If
End If
If t1.Text = 1440 Then
If t2.Text = 3240 Then
If t3.Text = 3240 Then
If t4.Text = 3240 Then
p8.Left = p8.Left - 1800
free.Left = free.Left + 1800
moves.Text = moves.Text + 1
End If
End If
End If
End If
If t1.Text = 1440 Then
If t2.Text = 1440 Then
If t3.Text = 3240 Then
If t4.Text = 1440 Then
p8.Left = p8.Left - 1800
free.Left = free.Left + 1800
moves.Text = moves.Text + 1
End If
End If
End If
End If
''''''''''left'''''''''

''''''''''up''''''''''
If t1.Text = 5040 Then
If t2.Text = 3240 Then
If t3.Text = 5040 Then
If t4.Text = 5040 Then
p8.Top = p8.Top - 1800
free.Top = free.Top + 1800
moves.Text = moves.Text + 1
End If
End If
End If
End If
If t1.Text = 3240 Then
If t2.Text = 3240 Then
If t3.Text = 3240 Then
If t4.Text = 5040 Then
p8.Top = p8.Top - 1800
free.Top = free.Top + 1800
moves.Text = moves.Text + 1
End If
End If
End If
End If
If t1.Text = 1440 Then
If t2.Text = 3240 Then
If t3.Text = 1440 Then
If t4.Text = 5040 Then
p8.Top = p8.Top - 1800
free.Top = free.Top + 1800
moves.Text = moves.Text + 1
End If
End If
End If
End If
If t1.Text = 1440 Then
If t2.Text = 1440 Then
If t3.Text = 1440 Then
If t4.Text = 3240 Then
p8.Top = p8.Top - 1800
free.Top = free.Top + 1800
moves.Text = moves.Text + 1
End If
End If
End If
End If
If t1.Text = 3240 Then
If t2.Text = 1440 Then
If t3.Text = 3240 Then
If t4.Text = 3240 Then
p8.Top = p8.Top - 1800
free.Top = free.Top + 1800
moves.Text = moves.Text + 1
End If
End If
End If
End If
If t1.Text = 5040 Then
If t2.Text = 1440 Then
If t3.Text = 5040 Then
If t4.Text = 3240 Then
p8.Top = p8.Top - 1800
free.Top = free.Top + 1800
moves.Text = moves.Text + 1
End If
End If
End If
End If
''''''''''''up'''''''''

'''''''''down'''''''''
If t1.Text = 1440 Then
If t2.Text = 3240 Then
If t3.Text = 1440 Then
If t4.Text = 1440 Then
p8.Top = p8.Top + 1800
free.Top = free.Top - 1800
moves.Text = moves.Text + 1
End If
End If
End If
End If
If t1.Text = 3240 Then
If t2.Text = 3240 Then
If t3.Text = 3240 Then
If t4.Text = 1440 Then
p8.Top = p8.Top + 1800
free.Top = free.Top - 1800
moves.Text = moves.Text + 1
End If
End If
End If
End If
If t1.Text = 5040 Then
If t2.Text = 3240 Then
If t3.Text = 5040 Then
If t4.Text = 1440 Then
p8.Top = p8.Top + 1800
free.Top = free.Top - 1800
moves.Text = moves.Text + 1
End If
End If
End If
End If
If t1.Text = 1440 Then
If t2.Text = 5040 Then
If t3.Text = 1440 Then
If t4.Text = 3240 Then
p8.Top = p8.Top + 1800
free.Top = free.Top - 1800
moves.Text = moves.Text + 1
End If
End If
End If
End If
If t1.Text = 3240 Then
If t2.Text = 5040 Then
If t3.Text = 3240 Then
If t4.Text = 3240 Then
p8.Top = p8.Top + 1800
free.Top = free.Top - 1800
moves.Text = moves.Text + 1
End If
End If
End If
End If
If t1.Text = 5040 Then
If t2.Text = 5040 Then
If t3.Text = 5040 Then
If t4.Text = 3240 Then
p8.Top = p8.Top + 1800
free.Top = free.Top - 1800
moves.Text = moves.Text + 1
End If
End If
End If
End If
End Sub

Private Sub Image8_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
t1.Text = free.Left
t2.Text = free.Top
t3.Text = p8.Left
t4.Text = p8.Top
End Sub
Private Sub Image6_Click()
''''''''right''''''''
If t1.Text = 3240 Then
If t2.Text = 1440 Then
If t3.Text = 1440 Then
If t4.Text = 1440 Then
p6.Left = p6.Left + 1800
free.Left = free.Left - 1800
moves.Text = moves.Text + 1
End If
End If
End If
End If
If t1.Text = 3240 Then
If t2.Text = 3240 Then
If t3.Text = 1440 Then
If t4.Text = 3240 Then
p6.Left = p6.Left + 1800
free.Left = free.Left - 1800
moves.Text = moves.Text + 1
End If
End If
End If
End If
If t1.Text = 3240 Then
If t2.Text = 5040 Then
If t3.Text = 1440 Then
If t4.Text = 5040 Then
p6.Left = p6.Left + 1800
free.Left = free.Left - 1800
moves.Text = moves.Text + 1
End If
End If
End If
End If
If t1.Text = 5040 Then
If t2.Text = 1440 Then
If t3.Text = 3240 Then
If t4.Text = 1440 Then
p6.Left = p6.Left + 1800
free.Left = free.Left - 1800
moves.Text = moves.Text + 1
End If
End If
End If
End If
If t1.Text = 5040 Then
If t2.Text = 3240 Then
If t3.Text = 3240 Then
If t4.Text = 3240 Then
p6.Left = p6.Left + 1800
free.Left = free.Left - 1800
moves.Text = moves.Text + 1
End If
End If
End If
End If
If t1.Text = 5040 Then
If t2.Text = 5040 Then
If t3.Text = 3240 Then
If t4.Text = 5040 Then
p6.Left = p6.Left + 1800
free.Left = free.Left - 1800
moves.Text = moves.Text + 1
End If
End If
End If
End If
''''''''''right''''''''''

''''''''''left'''''''''''
If t1.Text = 3240 Then
If t2.Text = 5040 Then
If t3.Text = 5040 Then
If t4.Text = 5040 Then
p6.Left = p6.Left - 1800
free.Left = free.Left + 1800
moves.Text = moves.Text + 1
End If
End If
End If
End If
If t1.Text = 3240 Then
If t2.Text = 3240 Then
If t3.Text = 5040 Then
If t4.Text = 3240 Then
p6.Left = p6.Left - 1800
free.Left = free.Left + 1800
moves.Text = moves.Text + 1
End If
End If
End If
End If
If t1.Text = 3240 Then
If t2.Text = 1440 Then
If t3.Text = 5040 Then
If t4.Text = 1440 Then
p6.Left = p6.Left - 1800
free.Left = free.Left + 1800
moves.Text = moves.Text + 1
End If
End If
End If
End If
If t1.Text = 1440 Then
If t2.Text = 5040 Then
If t3.Text = 3240 Then
If t4.Text = 5040 Then
p6.Left = p6.Left - 1800
free.Left = free.Left + 1800
moves.Text = moves.Text + 1
End If
End If
End If
End If
If t1.Text = 1440 Then
If t2.Text = 3240 Then
If t3.Text = 3240 Then
If t4.Text = 3240 Then
p6.Left = p6.Left - 1800
free.Left = free.Left + 1800
moves.Text = moves.Text + 1
End If
End If
End If
End If
If t1.Text = 1440 Then
If t2.Text = 1440 Then
If t3.Text = 3240 Then
If t4.Text = 1440 Then
p6.Left = p6.Left - 1800
free.Left = free.Left + 1800
moves.Text = moves.Text + 1
End If
End If
End If
End If
''''''''''left'''''''''

''''''''''up''''''''''
If t1.Text = 5040 Then
If t2.Text = 3240 Then
If t3.Text = 5040 Then
If t4.Text = 5040 Then
p6.Top = p6.Top - 1800
free.Top = free.Top + 1800
moves.Text = moves.Text + 1
End If
End If
End If
End If
If t1.Text = 3240 Then
If t2.Text = 3240 Then
If t3.Text = 3240 Then
If t4.Text = 5040 Then
p6.Top = p6.Top - 1800
free.Top = free.Top + 1800
moves.Text = moves.Text + 1
End If
End If
End If
End If
If t1.Text = 1440 Then
If t2.Text = 3240 Then
If t3.Text = 1440 Then
If t4.Text = 5040 Then
p6.Top = p6.Top - 1800
free.Top = free.Top + 1800
moves.Text = moves.Text + 1
End If
End If
End If
End If
If t1.Text = 1440 Then
If t2.Text = 1440 Then
If t3.Text = 1440 Then
If t4.Text = 3240 Then
p6.Top = p6.Top - 1800
free.Top = free.Top + 1800
moves.Text = moves.Text + 1
End If
End If
End If
End If
If t1.Text = 3240 Then
If t2.Text = 1440 Then
If t3.Text = 3240 Then
If t4.Text = 3240 Then
p6.Top = p6.Top - 1800
free.Top = free.Top + 1800
moves.Text = moves.Text + 1
End If
End If
End If
End If
If t1.Text = 5040 Then
If t2.Text = 1440 Then
If t3.Text = 5040 Then
If t4.Text = 3240 Then
p6.Top = p6.Top - 1800
free.Top = free.Top + 1800
moves.Text = moves.Text + 1
End If
End If
End If
End If
''''''''''''up'''''''''

'''''''''down'''''''''
If t1.Text = 1440 Then
If t2.Text = 3240 Then
If t3.Text = 1440 Then
If t4.Text = 1440 Then
p6.Top = p6.Top + 1800
free.Top = free.Top - 1800
moves.Text = moves.Text + 1
End If
End If
End If
End If
If t1.Text = 3240 Then
If t2.Text = 3240 Then
If t3.Text = 3240 Then
If t4.Text = 1440 Then
p6.Top = p6.Top + 1800
free.Top = free.Top - 1800
moves.Text = moves.Text + 1
End If
End If
End If
End If
If t1.Text = 5040 Then
If t2.Text = 3240 Then
If t3.Text = 5040 Then
If t4.Text = 1440 Then
p6.Top = p6.Top + 1800
free.Top = free.Top - 1800
moves.Text = moves.Text + 1
End If
End If
End If
End If
If t1.Text = 1440 Then
If t2.Text = 5040 Then
If t3.Text = 1440 Then
If t4.Text = 3240 Then
p6.Top = p6.Top + 1800
free.Top = free.Top - 1800
moves.Text = moves.Text + 1
End If
End If
End If
End If
If t1.Text = 3240 Then
If t2.Text = 5040 Then
If t3.Text = 3240 Then
If t4.Text = 3240 Then
p6.Top = p6.Top + 1800
free.Top = free.Top - 1800
moves.Text = moves.Text + 1
End If
End If
End If
End If
If t1.Text = 5040 Then
If t2.Text = 5040 Then
If t3.Text = 5040 Then
If t4.Text = 3240 Then
p6.Top = p6.Top + 1800
free.Top = free.Top - 1800
moves.Text = moves.Text + 1
End If
End If
End If
End If
End Sub
Private Sub Image5_Click()
''''''''right''''''''
If t1.Text = 3240 Then
If t2.Text = 1440 Then
If t3.Text = 1440 Then
If t4.Text = 1440 Then
p5.Left = p5.Left + 1800
free.Left = free.Left - 1800
moves.Text = moves.Text + 1

End If
End If
End If
End If
If t1.Text = 3240 Then
If t2.Text = 3240 Then
If t3.Text = 1440 Then
If t4.Text = 3240 Then
p5.Left = p5.Left + 1800
free.Left = free.Left - 1800
moves.Text = moves.Text + 1
End If
End If
End If
End If
If t1.Text = 3240 Then
If t2.Text = 5040 Then
If t3.Text = 1440 Then
If t4.Text = 5040 Then
p5.Left = p5.Left + 1800
free.Left = free.Left - 1800
moves.Text = moves.Text + 1
End If
End If
End If
End If
If t1.Text = 5040 Then
If t2.Text = 1440 Then
If t3.Text = 3240 Then
If t4.Text = 1440 Then
p5.Left = p5.Left + 1800
free.Left = free.Left - 1800
moves.Text = moves.Text + 1
End If
End If
End If
End If
If t1.Text = 5040 Then
If t2.Text = 3240 Then
If t3.Text = 3240 Then
If t4.Text = 3240 Then
p5.Left = p5.Left + 1800
free.Left = free.Left - 1800
moves.Text = moves.Text + 1
End If
End If
End If
End If
If t1.Text = 5040 Then
If t2.Text = 5040 Then
If t3.Text = 3240 Then
If t4.Text = 5040 Then
p5.Left = p5.Left + 1800
free.Left = free.Left - 1800
moves.Text = moves.Text + 1
End If
End If
End If
End If
''''''''''right''''''''''

''''''''''left'''''''''''
If t1.Text = 3240 Then
If t2.Text = 5040 Then
If t3.Text = 5040 Then
If t4.Text = 5040 Then
p5.Left = p5.Left - 1800
free.Left = free.Left + 1800
moves.Text = moves.Text + 1
End If
End If
End If
End If
If t1.Text = 3240 Then
If t2.Text = 3240 Then
If t3.Text = 5040 Then
If t4.Text = 3240 Then
p5.Left = p5.Left - 1800
free.Left = free.Left + 1800
moves.Text = moves.Text + 1
End If
End If
End If
End If
If t1.Text = 3240 Then
If t2.Text = 1440 Then
If t3.Text = 5040 Then
If t4.Text = 1440 Then
p5.Left = p5.Left - 1800
free.Left = free.Left + 1800
moves.Text = moves.Text + 1
End If
End If
End If
End If
If t1.Text = 1440 Then
If t2.Text = 5040 Then
If t3.Text = 3240 Then
If t4.Text = 5040 Then
p5.Left = p5.Left - 1800
free.Left = free.Left + 1800
moves.Text = moves.Text + 1
End If
End If
End If
End If
If t1.Text = 1440 Then
If t2.Text = 3240 Then
If t3.Text = 3240 Then
If t4.Text = 3240 Then
p5.Left = p5.Left - 1800
free.Left = free.Left + 1800
moves.Text = moves.Text + 1
End If
End If
End If
End If
If t1.Text = 1440 Then
If t2.Text = 1440 Then
If t3.Text = 3240 Then
If t4.Text = 1440 Then
p5.Left = p5.Left - 1800
free.Left = free.Left + 1800
moves.Text = moves.Text + 1
End If
End If
End If
End If
''''''''''left'''''''''

''''''''''up''''''''''
If t1.Text = 5040 Then
If t2.Text = 3240 Then
If t3.Text = 5040 Then
If t4.Text = 5040 Then
p5.Top = p5.Top - 1800
free.Top = free.Top + 1800
moves.Text = moves.Text + 1
End If
End If
End If
End If
If t1.Text = 3240 Then
If t2.Text = 3240 Then
If t3.Text = 3240 Then
If t4.Text = 5040 Then
p5.Top = p5.Top - 1800
free.Top = free.Top + 1800
moves.Text = moves.Text + 1
End If
End If
End If
End If
If t1.Text = 1440 Then
If t2.Text = 3240 Then
If t3.Text = 1440 Then
If t4.Text = 5040 Then
p5.Top = p5.Top - 1800
free.Top = free.Top + 1800
moves.Text = moves.Text + 1
End If
End If
End If
End If
If t1.Text = 1440 Then
If t2.Text = 1440 Then
If t3.Text = 1440 Then
If t4.Text = 3240 Then
p5.Top = p5.Top - 1800
free.Top = free.Top + 1800
moves.Text = moves.Text + 1
End If
End If
End If
End If
If t1.Text = 3240 Then
If t2.Text = 1440 Then
If t3.Text = 3240 Then
If t4.Text = 3240 Then
p5.Top = p5.Top - 1800
free.Top = free.Top + 1800
moves.Text = moves.Text + 1
End If
End If
End If
End If
If t1.Text = 5040 Then
If t2.Text = 1440 Then
If t3.Text = 5040 Then
If t4.Text = 3240 Then
p5.Top = p5.Top - 1800
free.Top = free.Top + 1800
moves.Text = moves.Text + 1
End If
End If
End If
End If
''''''''''''up'''''''''

'''''''''down'''''''''
If t1.Text = 1440 Then
If t2.Text = 3240 Then
If t3.Text = 1440 Then
If t4.Text = 1440 Then
p5.Top = p5.Top + 1800
free.Top = free.Top - 1800
moves.Text = moves.Text + 1
End If
End If
End If
End If
If t1.Text = 3240 Then
If t2.Text = 3240 Then
If t3.Text = 3240 Then
If t4.Text = 1440 Then
p5.Top = p5.Top + 1800
free.Top = free.Top - 1800
moves.Text = moves.Text + 1
End If
End If
End If
End If
If t1.Text = 5040 Then
If t2.Text = 3240 Then
If t3.Text = 5040 Then
If t4.Text = 1440 Then
p5.Top = p5.Top + 1800
free.Top = free.Top - 1800
moves.Text = moves.Text + 1
End If
End If
End If
End If
If t1.Text = 1440 Then
If t2.Text = 5040 Then
If t3.Text = 1440 Then
If t4.Text = 3240 Then
p5.Top = p5.Top + 1800
free.Top = free.Top - 1800
moves.Text = moves.Text + 1
End If
End If
End If
End If
If t1.Text = 3240 Then
If t2.Text = 5040 Then
If t3.Text = 3240 Then
If t4.Text = 3240 Then
p5.Top = p5.Top + 1800
free.Top = free.Top - 1800
moves.Text = moves.Text + 1
End If
End If
End If
End If
If t1.Text = 5040 Then
If t2.Text = 5040 Then
If t3.Text = 5040 Then
If t4.Text = 3240 Then
p5.Top = p5.Top + 1800
free.Top = free.Top - 1800
moves.Text = moves.Text + 1
End If
End If
End If
End If
End Sub
Private Sub Image4_Click()
''''''''right''''''''
If t1.Text = 3240 Then
If t2.Text = 1440 Then
If t3.Text = 1440 Then
If t4.Text = 1440 Then
p4.Left = p4.Left + 1800
free.Left = free.Left - 1800
moves.Text = moves.Text + 1
End If
End If
End If
End If
If t1.Text = 3240 Then
If t2.Text = 3240 Then
If t3.Text = 1440 Then
If t4.Text = 3240 Then
p4.Left = p4.Left + 1800
free.Left = free.Left - 1800
moves.Text = moves.Text + 1
End If
End If
End If
End If
If t1.Text = 3240 Then
If t2.Text = 5040 Then
If t3.Text = 1440 Then
If t4.Text = 5040 Then
p4.Left = p4.Left + 1800
free.Left = free.Left - 1800
moves.Text = moves.Text + 1
End If
End If
End If
End If
If t1.Text = 5040 Then
If t2.Text = 1440 Then
If t3.Text = 3240 Then
If t4.Text = 1440 Then
p4.Left = p4.Left + 1800
free.Left = free.Left - 1800
moves.Text = moves.Text + 1
End If
End If
End If
End If
If t1.Text = 5040 Then
If t2.Text = 3240 Then
If t3.Text = 3240 Then
If t4.Text = 3240 Then
p4.Left = p4.Left + 1800
free.Left = free.Left - 1800
moves.Text = moves.Text + 1
End If
End If
End If
End If
If t1.Text = 5040 Then
If t2.Text = 5040 Then
If t3.Text = 3240 Then
If t4.Text = 5040 Then
p4.Left = p4.Left + 1800
free.Left = free.Left - 1800
moves.Text = moves.Text + 1
End If
End If
End If
End If
''''''''''right''''''''''

''''''''''left'''''''''''
If t1.Text = 3240 Then
If t2.Text = 5040 Then
If t3.Text = 5040 Then
If t4.Text = 5040 Then
p4.Left = p4.Left - 1800
free.Left = free.Left + 1800
moves.Text = moves.Text + 1
End If
End If
End If
End If
If t1.Text = 3240 Then
If t2.Text = 3240 Then
If t3.Text = 5040 Then
If t4.Text = 3240 Then
p4.Left = p4.Left - 1800
free.Left = free.Left + 1800
moves.Text = moves.Text + 1
End If
End If
End If
End If
If t1.Text = 3240 Then
If t2.Text = 1440 Then
If t3.Text = 5040 Then
If t4.Text = 1440 Then
p4.Left = p4.Left - 1800
free.Left = free.Left + 1800
moves.Text = moves.Text + 1
End If
End If
End If
End If
If t1.Text = 1440 Then
If t2.Text = 5040 Then
If t3.Text = 3240 Then
If t4.Text = 5040 Then
p4.Left = p4.Left - 1800
free.Left = free.Left + 1800
moves.Text = moves.Text + 1
End If
End If
End If
End If
If t1.Text = 1440 Then
If t2.Text = 3240 Then
If t3.Text = 3240 Then
If t4.Text = 3240 Then
p4.Left = p4.Left - 1800
free.Left = free.Left + 1800
moves.Text = moves.Text + 1
End If
End If
End If
End If
If t1.Text = 1440 Then
If t2.Text = 1440 Then
If t3.Text = 3240 Then
If t4.Text = 1440 Then
p4.Left = p4.Left - 1800
free.Left = free.Left + 1800
moves.Text = moves.Text + 1
End If
End If
End If
End If
''''''''''left'''''''''

''''''''''up''''''''''
If t1.Text = 5040 Then
If t2.Text = 3240 Then
If t3.Text = 5040 Then
If t4.Text = 5040 Then
p4.Top = p4.Top - 1800
free.Top = free.Top + 1800
moves.Text = moves.Text + 1
End If
End If
End If
End If
If t1.Text = 3240 Then
If t2.Text = 3240 Then
If t3.Text = 3240 Then
If t4.Text = 5040 Then
p4.Top = p4.Top - 1800
free.Top = free.Top + 1800
moves.Text = moves.Text + 1
End If
End If
End If
End If
If t1.Text = 1440 Then
If t2.Text = 3240 Then
If t3.Text = 1440 Then
If t4.Text = 5040 Then
p4.Top = p4.Top - 1800
free.Top = free.Top + 1800
moves.Text = moves.Text + 1
End If
End If
End If
End If
If t1.Text = 1440 Then
If t2.Text = 1440 Then
If t3.Text = 1440 Then
If t4.Text = 3240 Then
p4.Top = p4.Top - 1800
free.Top = free.Top + 1800
moves.Text = moves.Text + 1
End If
End If
End If
End If
If t1.Text = 3240 Then
If t2.Text = 1440 Then
If t3.Text = 3240 Then
If t4.Text = 3240 Then
p4.Top = p4.Top - 1800
free.Top = free.Top + 1800
moves.Text = moves.Text + 1
End If
End If
End If
End If
If t1.Text = 5040 Then
If t2.Text = 1440 Then
If t3.Text = 5040 Then
If t4.Text = 3240 Then
p4.Top = p4.Top - 1800
free.Top = free.Top + 1800
moves.Text = moves.Text + 1
End If
End If
End If
End If
''''''''''''up'''''''''

'''''''''down'''''''''
If t1.Text = 1440 Then
If t2.Text = 3240 Then
If t3.Text = 1440 Then
If t4.Text = 1440 Then
p4.Top = p4.Top + 1800
free.Top = free.Top - 1800
moves.Text = moves.Text + 1
End If
End If
End If
End If
If t1.Text = 3240 Then
If t2.Text = 3240 Then
If t3.Text = 3240 Then
If t4.Text = 1440 Then
p4.Top = p4.Top + 1800
free.Top = free.Top - 1800
moves.Text = moves.Text + 1
End If
End If
End If
End If
If t1.Text = 5040 Then
If t2.Text = 3240 Then
If t3.Text = 5040 Then
If t4.Text = 1440 Then
p4.Top = p4.Top + 1800
free.Top = free.Top - 1800
moves.Text = moves.Text + 1
End If
End If
End If
End If
If t1.Text = 1440 Then
If t2.Text = 5040 Then
If t3.Text = 1440 Then
If t4.Text = 3240 Then
p4.Top = p4.Top + 1800
free.Top = free.Top - 1800
moves.Text = moves.Text + 1
End If
End If
End If
End If
If t1.Text = 3240 Then
If t2.Text = 5040 Then
If t3.Text = 3240 Then
If t4.Text = 3240 Then
p4.Top = p4.Top + 1800
free.Top = free.Top - 1800
moves.Text = moves.Text + 1
End If
End If
End If
End If
If t1.Text = 5040 Then
If t2.Text = 5040 Then
If t3.Text = 5040 Then
If t4.Text = 3240 Then
p4.Top = p4.Top + 1800
free.Top = free.Top - 1800
moves.Text = moves.Text + 1
End If
End If
End If
End If
End Sub
Private Sub Image7_Click()
''''''''right''''''''
If t1.Text = 3240 Then
If t2.Text = 1440 Then
If t3.Text = 1440 Then
If t4.Text = 1440 Then
p7.Left = p7.Left + 1800
free.Left = free.Left - 1800
moves.Text = moves.Text + 1
End If
End If
End If
End If
If t1.Text = 3240 Then
If t2.Text = 3240 Then
If t3.Text = 1440 Then
If t4.Text = 3240 Then
p7.Left = p7.Left + 1800
free.Left = free.Left - 1800
moves.Text = moves.Text + 1
End If
End If
End If
End If
If t1.Text = 3240 Then
If t2.Text = 5040 Then
If t3.Text = 1440 Then
If t4.Text = 5040 Then
p7.Left = p7.Left + 1800
free.Left = free.Left - 1800
moves.Text = moves.Text + 1
End If
End If
End If
End If
If t1.Text = 5040 Then
If t2.Text = 1440 Then
If t3.Text = 3240 Then
If t4.Text = 1440 Then
p7.Left = p7.Left + 1800
free.Left = free.Left - 1800
moves.Text = moves.Text + 1
End If
End If
End If
End If
If t1.Text = 5040 Then
If t2.Text = 3240 Then
If t3.Text = 3240 Then
If t4.Text = 3240 Then
p7.Left = p7.Left + 1800
free.Left = free.Left - 1800
moves.Text = moves.Text + 1
End If
End If
End If
End If
If t1.Text = 5040 Then
If t2.Text = 5040 Then
If t3.Text = 3240 Then
If t4.Text = 5040 Then
p7.Left = p7.Left + 1800
free.Left = free.Left - 1800
moves.Text = moves.Text + 1
End If
End If
End If
End If
''''''''''right''''''''''

''''''''''left'''''''''''
If t1.Text = 3240 Then
If t2.Text = 5040 Then
If t3.Text = 5040 Then
If t4.Text = 5040 Then
p7.Left = p7.Left - 1800
free.Left = free.Left + 1800
moves.Text = moves.Text + 1
End If
End If
End If
End If
If t1.Text = 3240 Then
If t2.Text = 3240 Then
If t3.Text = 5040 Then
If t4.Text = 3240 Then
p7.Left = p7.Left - 1800
free.Left = free.Left + 1800
moves.Text = moves.Text + 1
End If
End If
End If
End If
If t1.Text = 3240 Then
If t2.Text = 1440 Then
If t3.Text = 5040 Then
If t4.Text = 1440 Then
p7.Left = p7.Left - 1800
free.Left = free.Left + 1800
moves.Text = moves.Text + 1
End If
End If
End If
End If
If t1.Text = 1440 Then
If t2.Text = 5040 Then
If t3.Text = 3240 Then
If t4.Text = 5040 Then
p7.Left = p7.Left - 1800
free.Left = free.Left + 1800
moves.Text = moves.Text + 1
End If
End If
End If
End If
If t1.Text = 1440 Then
If t2.Text = 3240 Then
If t3.Text = 3240 Then
If t4.Text = 3240 Then
p7.Left = p7.Left - 1800
free.Left = free.Left + 1800
moves.Text = moves.Text + 1
End If
End If
End If
End If
If t1.Text = 1440 Then
If t2.Text = 1440 Then
If t3.Text = 3240 Then
If t4.Text = 1440 Then
p7.Left = p7.Left - 1800
free.Left = free.Left + 1800
moves.Text = moves.Text + 1
End If
End If
End If
End If
''''''''''left'''''''''

''''''''''up''''''''''
If t1.Text = 5040 Then
If t2.Text = 3240 Then
If t3.Text = 5040 Then
If t4.Text = 5040 Then
p7.Top = p7.Top - 1800
free.Top = free.Top + 1800
moves.Text = moves.Text + 1
End If
End If
End If
End If
If t1.Text = 3240 Then
If t2.Text = 3240 Then
If t3.Text = 3240 Then
If t4.Text = 5040 Then
p7.Top = p7.Top - 1800
free.Top = free.Top + 1800
moves.Text = moves.Text + 1
End If
End If
End If
End If
If t1.Text = 1440 Then
If t2.Text = 3240 Then
If t3.Text = 1440 Then
If t4.Text = 5040 Then
p7.Top = p7.Top - 1800
free.Top = free.Top + 1800
moves.Text = moves.Text + 1
End If
End If
End If
End If
If t1.Text = 1440 Then
If t2.Text = 1440 Then
If t3.Text = 1440 Then
If t4.Text = 3240 Then
p7.Top = p7.Top - 1800
free.Top = free.Top + 1800
moves.Text = moves.Text + 1
End If
End If
End If
End If
If t1.Text = 3240 Then
If t2.Text = 1440 Then
If t3.Text = 3240 Then
If t4.Text = 3240 Then
p7.Top = p7.Top - 1800
free.Top = free.Top + 1800
moves.Text = moves.Text + 1
End If
End If
End If
End If
If t1.Text = 5040 Then
If t2.Text = 1440 Then
If t3.Text = 5040 Then
If t4.Text = 3240 Then
p7.Top = p7.Top - 1800
free.Top = free.Top + 1800
moves.Text = moves.Text + 1
End If
End If
End If
End If
''''''''''''up'''''''''

'''''''''down'''''''''
If t1.Text = 1440 Then
If t2.Text = 3240 Then
If t3.Text = 1440 Then
If t4.Text = 1440 Then
p7.Top = p7.Top + 1800
free.Top = free.Top - 1800
moves.Text = moves.Text + 1
End If
End If
End If
End If
If t1.Text = 3240 Then
If t2.Text = 3240 Then
If t3.Text = 3240 Then
If t4.Text = 1440 Then
p7.Top = p7.Top + 1800
free.Top = free.Top - 1800
moves.Text = moves.Text + 1
End If
End If
End If
End If
If t1.Text = 5040 Then
If t2.Text = 3240 Then
If t3.Text = 5040 Then
If t4.Text = 1440 Then
p7.Top = p7.Top + 1800
free.Top = free.Top - 1800
moves.Text = moves.Text + 1
End If
End If
End If
End If
If t1.Text = 1440 Then
If t2.Text = 5040 Then
If t3.Text = 1440 Then
If t4.Text = 3240 Then
p7.Top = p7.Top + 1800
free.Top = free.Top - 1800
moves.Text = moves.Text + 1
End If
End If
End If
End If
If t1.Text = 3240 Then
If t2.Text = 5040 Then
If t3.Text = 3240 Then
If t4.Text = 3240 Then
p7.Top = p7.Top + 1800
free.Top = free.Top - 1800
moves.Text = moves.Text + 1
End If
End If
End If
End If
If t1.Text = 5040 Then
If t2.Text = 5040 Then
If t3.Text = 5040 Then
If t4.Text = 3240 Then
p7.Top = p7.Top + 1800
free.Top = free.Top - 1800
moves.Text = moves.Text + 1
End If
End If
End If
End If
End Sub
Private Sub Image1_Click()
''''''''right''''''''
If t1.Text = 3240 Then
If t2.Text = 1440 Then
If t3.Text = 1440 Then
If t4.Text = 1440 Then
p1.Left = p1.Left + 1800
free.Left = free.Left - 1800
moves.Text = moves.Text + 1
End If
End If
End If
End If
If t1.Text = 3240 Then
If t2.Text = 3240 Then
If t3.Text = 1440 Then
If t4.Text = 3240 Then
p1.Left = p1.Left + 1800
free.Left = free.Left - 1800
moves.Text = moves.Text + 1
End If
End If
End If
End If
If t1.Text = 3240 Then
If t2.Text = 5040 Then
If t3.Text = 1440 Then
If t4.Text = 5040 Then
p1.Left = p1.Left + 1800
free.Left = free.Left - 1800
moves.Text = moves.Text + 1
End If
End If
End If
End If
If t1.Text = 5040 Then
If t2.Text = 1440 Then
If t3.Text = 3240 Then
If t4.Text = 1440 Then
p1.Left = p1.Left + 1800
free.Left = free.Left - 1800
moves.Text = moves.Text + 1
End If
End If
End If
End If
If t1.Text = 5040 Then
If t2.Text = 3240 Then
If t3.Text = 3240 Then
If t4.Text = 3240 Then
p1.Left = p1.Left + 1800
free.Left = free.Left - 1800
moves.Text = moves.Text + 1
End If
End If
End If
End If
If t1.Text = 5040 Then
If t2.Text = 5040 Then
If t3.Text = 3240 Then
If t4.Text = 5040 Then
p1.Left = p1.Left + 1800
free.Left = free.Left - 1800
moves.Text = moves.Text + 1
End If
End If
End If
End If
''''''''''right''''''''''

''''''''''left'''''''''''
If t1.Text = 3240 Then
If t2.Text = 5040 Then
If t3.Text = 5040 Then
If t4.Text = 5040 Then
p1.Left = p1.Left - 1800
free.Left = free.Left + 1800
moves.Text = moves.Text + 1
End If
End If
End If
End If
If t1.Text = 3240 Then
If t2.Text = 3240 Then
If t3.Text = 5040 Then
If t4.Text = 3240 Then
p1.Left = p1.Left - 1800
free.Left = free.Left + 1800
moves.Text = moves.Text + 1
End If
End If
End If
End If
If t1.Text = 3240 Then
If t2.Text = 1440 Then
If t3.Text = 5040 Then
If t4.Text = 1440 Then
p1.Left = p1.Left - 1800
free.Left = free.Left + 1800
moves.Text = moves.Text + 1
End If
End If
End If
End If
If t1.Text = 1440 Then
If t2.Text = 5040 Then
If t3.Text = 3240 Then
If t4.Text = 5040 Then
p1.Left = p1.Left - 1800
free.Left = free.Left + 1800
moves.Text = moves.Text + 1
End If
End If
End If
End If
If t1.Text = 1440 Then
If t2.Text = 3240 Then
If t3.Text = 3240 Then
If t4.Text = 3240 Then
p1.Left = p1.Left - 1800
free.Left = free.Left + 1800
moves.Text = moves.Text + 1
End If
End If
End If
End If
If t1.Text = 1440 Then
If t2.Text = 1440 Then
If t3.Text = 3240 Then
If t4.Text = 1440 Then
p1.Left = p1.Left - 1800
free.Left = free.Left + 1800
moves.Text = moves.Text + 1
End If
End If
End If
End If
''''''''''left'''''''''

''''''''''up''''''''''
If t1.Text = 5040 Then
If t2.Text = 3240 Then
If t3.Text = 5040 Then
If t4.Text = 5040 Then
p1.Top = p1.Top - 1800
free.Top = free.Top + 1800
moves.Text = moves.Text + 1
End If
End If
End If
End If
If t1.Text = 3240 Then
If t2.Text = 3240 Then
If t3.Text = 3240 Then
If t4.Text = 5040 Then
p1.Top = p1.Top - 1800
free.Top = free.Top + 1800
moves.Text = moves.Text + 1
End If
End If
End If
End If
If t1.Text = 1440 Then
If t2.Text = 3240 Then
If t3.Text = 1440 Then
If t4.Text = 5040 Then
p1.Top = p1.Top - 1800
free.Top = free.Top + 1800
moves.Text = moves.Text + 1
End If
End If
End If
End If
If t1.Text = 1440 Then
If t2.Text = 1440 Then
If t3.Text = 1440 Then
If t4.Text = 3240 Then
p1.Top = p1.Top - 1800
free.Top = free.Top + 1800
moves.Text = moves.Text + 1
End If
End If
End If
End If
If t1.Text = 3240 Then
If t2.Text = 1440 Then
If t3.Text = 3240 Then
If t4.Text = 3240 Then
p1.Top = p1.Top - 1800
free.Top = free.Top + 1800
moves.Text = moves.Text + 1
End If
End If
End If
End If
If t1.Text = 5040 Then
If t2.Text = 1440 Then
If t3.Text = 5040 Then
If t4.Text = 3240 Then
p1.Top = p1.Top - 1800
free.Top = free.Top + 1800
moves.Text = moves.Text + 1
End If
End If
End If
End If
''''''''''''up'''''''''

'''''''''down'''''''''
If t1.Text = 1440 Then
If t2.Text = 3240 Then
If t3.Text = 1440 Then
If t4.Text = 1440 Then
p1.Top = p1.Top + 1800
free.Top = free.Top - 1800
moves.Text = moves.Text + 1
End If
End If
End If
End If
If t1.Text = 3240 Then
If t2.Text = 3240 Then
If t3.Text = 3240 Then
If t4.Text = 1440 Then
p1.Top = p1.Top + 1800
free.Top = free.Top - 1800
moves.Text = moves.Text + 1
End If
End If
End If
End If
If t1.Text = 5040 Then
If t2.Text = 3240 Then
If t3.Text = 5040 Then
If t4.Text = 1440 Then
p1.Top = p1.Top + 1800
free.Top = free.Top - 1800
moves.Text = moves.Text + 1
End If
End If
End If
End If
If t1.Text = 1440 Then
If t2.Text = 5040 Then
If t3.Text = 1440 Then
If t4.Text = 3240 Then
p1.Top = p1.Top + 1800
free.Top = free.Top - 1800
moves.Text = moves.Text + 1
End If
End If
End If
End If
If t1.Text = 3240 Then
If t2.Text = 5040 Then
If t3.Text = 3240 Then
If t4.Text = 3240 Then
p1.Top = p1.Top + 1800
free.Top = free.Top - 1800
moves.Text = moves.Text + 1
End If
End If
End If
End If
If t1.Text = 5040 Then
If t2.Text = 5040 Then
If t3.Text = 5040 Then
If t4.Text = 3240 Then
p1.Top = p1.Top + 1800
free.Top = free.Top - 1800
moves.Text = moves.Text + 1
End If
End If
End If
End If
End Sub
Private Sub Image2_Click()
''''''''right''''''''
If t1.Text = 3240 Then
If t2.Text = 1440 Then
If t3.Text = 1440 Then
If t4.Text = 1440 Then
p2.Left = p2.Left + 1800
free.Left = free.Left - 1800
moves.Text = moves.Text + 1
End If
End If
End If
End If
If t1.Text = 3240 Then
If t2.Text = 3240 Then
If t3.Text = 1440 Then
If t4.Text = 3240 Then
p2.Left = p2.Left + 1800
free.Left = free.Left - 1800
moves.Text = moves.Text + 1
End If
End If
End If
End If
If t1.Text = 3240 Then
If t2.Text = 5040 Then
If t3.Text = 1440 Then
If t4.Text = 5040 Then
p2.Left = p2.Left + 1800
free.Left = free.Left - 1800
moves.Text = moves.Text + 1
End If
End If
End If
End If
If t1.Text = 5040 Then
If t2.Text = 1440 Then
If t3.Text = 3240 Then
If t4.Text = 1440 Then
p2.Left = p2.Left + 1800
free.Left = free.Left - 1800
moves.Text = moves.Text + 1
End If
End If
End If
End If
If t1.Text = 5040 Then
If t2.Text = 3240 Then
If t3.Text = 3240 Then
If t4.Text = 3240 Then
p2.Left = p2.Left + 1800
free.Left = free.Left - 1800
moves.Text = moves.Text + 1
End If
End If
End If
End If
If t1.Text = 5040 Then
If t2.Text = 5040 Then
If t3.Text = 3240 Then
If t4.Text = 5040 Then
p2.Left = p2.Left + 1800
free.Left = free.Left - 1800
moves.Text = moves.Text + 1
End If
End If
End If
End If
''''''''''right''''''''''

''''''''''left'''''''''''
If t1.Text = 3240 Then
If t2.Text = 5040 Then
If t3.Text = 5040 Then
If t4.Text = 5040 Then
p2.Left = p2.Left - 1800
free.Left = free.Left + 1800
moves.Text = moves.Text + 1
End If
End If
End If
End If
If t1.Text = 3240 Then
If t2.Text = 3240 Then
If t3.Text = 5040 Then
If t4.Text = 3240 Then
p2.Left = p2.Left - 1800
free.Left = free.Left + 1800
moves.Text = moves.Text + 1
End If
End If
End If
End If
If t1.Text = 3240 Then
If t2.Text = 1440 Then
If t3.Text = 5040 Then
If t4.Text = 1440 Then
p2.Left = p2.Left - 1800
free.Left = free.Left + 1800
moves.Text = moves.Text + 1
End If
End If
End If
End If
If t1.Text = 1440 Then
If t2.Text = 5040 Then
If t3.Text = 3240 Then
If t4.Text = 5040 Then
p2.Left = p2.Left - 1800
free.Left = free.Left + 1800
moves.Text = moves.Text + 1
End If
End If
End If
End If
If t1.Text = 1440 Then
If t2.Text = 3240 Then
If t3.Text = 3240 Then
If t4.Text = 3240 Then
p2.Left = p2.Left - 1800
free.Left = free.Left + 1800
moves.Text = moves.Text + 1
End If
End If
End If
End If
If t1.Text = 1440 Then
If t2.Text = 1440 Then
If t3.Text = 3240 Then
If t4.Text = 1440 Then
p2.Left = p2.Left - 1800
free.Left = free.Left + 1800
moves.Text = moves.Text + 1
End If
End If
End If
End If
''''''''''left'''''''''

''''''''''up''''''''''
If t1.Text = 5040 Then
If t2.Text = 3240 Then
If t3.Text = 5040 Then
If t4.Text = 5040 Then
p2.Top = p2.Top - 1800
free.Top = free.Top + 1800
moves.Text = moves.Text + 1
End If
End If
End If
End If
If t1.Text = 3240 Then
If t2.Text = 3240 Then
If t3.Text = 3240 Then
If t4.Text = 5040 Then
p2.Top = p2.Top - 1800
free.Top = free.Top + 1800
moves.Text = moves.Text + 1
End If
End If
End If
End If
If t1.Text = 1440 Then
If t2.Text = 3240 Then
If t3.Text = 1440 Then
If t4.Text = 5040 Then
p2.Top = p2.Top - 1800
free.Top = free.Top + 1800
moves.Text = moves.Text + 1
End If
End If
End If
End If
If t1.Text = 1440 Then
If t2.Text = 1440 Then
If t3.Text = 1440 Then
If t4.Text = 3240 Then
p2.Top = p2.Top - 1800
free.Top = free.Top + 1800
moves.Text = moves.Text + 1
End If
End If
End If
End If
If t1.Text = 3240 Then
If t2.Text = 1440 Then
If t3.Text = 3240 Then
If t4.Text = 3240 Then
p2.Top = p2.Top - 1800
free.Top = free.Top + 1800
moves.Text = moves.Text + 1
End If
End If
End If
End If
If t1.Text = 5040 Then
If t2.Text = 1440 Then
If t3.Text = 5040 Then
If t4.Text = 3240 Then
p2.Top = p2.Top - 1800
free.Top = free.Top + 1800
moves.Text = moves.Text + 1
End If
End If
End If
End If
''''''''''''up'''''''''

'''''''''down'''''''''
If t1.Text = 1440 Then
If t2.Text = 3240 Then
If t3.Text = 1440 Then
If t4.Text = 1440 Then
p2.Top = p2.Top + 1800
free.Top = free.Top - 1800
moves.Text = moves.Text + 1
End If
End If
End If
End If
If t1.Text = 3240 Then
If t2.Text = 3240 Then
If t3.Text = 3240 Then
If t4.Text = 1440 Then
p2.Top = p2.Top + 1800
free.Top = free.Top - 1800
moves.Text = moves.Text + 1
End If
End If
End If
End If
If t1.Text = 5040 Then
If t2.Text = 3240 Then
If t3.Text = 5040 Then
If t4.Text = 1440 Then
p2.Top = p2.Top + 1800
free.Top = free.Top - 1800
moves.Text = moves.Text + 1
End If
End If
End If
End If
If t1.Text = 1440 Then
If t2.Text = 5040 Then
If t3.Text = 1440 Then
If t4.Text = 3240 Then
p2.Top = p2.Top + 1800
free.Top = free.Top - 1800
moves.Text = moves.Text + 1
End If
End If
End If
End If
If t1.Text = 3240 Then
If t2.Text = 5040 Then
If t3.Text = 3240 Then
If t4.Text = 3240 Then
p2.Top = p2.Top + 1800
free.Top = free.Top - 1800
moves.Text = moves.Text + 1
End If
End If
End If
End If
If t1.Text = 5040 Then
If t2.Text = 5040 Then
If t3.Text = 5040 Then
If t4.Text = 3240 Then
p2.Top = p2.Top + 1800
free.Top = free.Top - 1800
moves.Text = moves.Text + 1
End If
End If
End If
End If
End Sub
Private Sub Image3_Click()
''''''''right''''''''
If t1.Text = 3240 Then
If t2.Text = 1440 Then
If t3.Text = 1440 Then
If t4.Text = 1440 Then
p3.Left = p3.Left + 1800
free.Left = free.Left - 1800
moves.Text = moves.Text + 1
End If
End If
End If
End If
If t1.Text = 3240 Then
If t2.Text = 3240 Then
If t3.Text = 1440 Then
If t4.Text = 3240 Then
p3.Left = p3.Left + 1800
free.Left = free.Left - 1800
moves.Text = moves.Text + 1
End If
End If
End If
End If
If t1.Text = 3240 Then
If t2.Text = 5040 Then
If t3.Text = 1440 Then
If t4.Text = 5040 Then
p3.Left = p3.Left + 1800
free.Left = free.Left - 1800
moves.Text = moves.Text + 1
End If
End If
End If
End If
If t1.Text = 5040 Then
If t2.Text = 1440 Then
If t3.Text = 3240 Then
If t4.Text = 1440 Then
p3.Left = p3.Left + 1800
free.Left = free.Left - 1800
moves.Text = moves.Text + 1
End If
End If
End If
End If
If t1.Text = 5040 Then
If t2.Text = 3240 Then
If t3.Text = 3240 Then
If t4.Text = 3240 Then
p3.Left = p3.Left + 1800
free.Left = free.Left - 1800
moves.Text = moves.Text + 1
End If
End If
End If
End If
If t1.Text = 5040 Then
If t2.Text = 5040 Then
If t3.Text = 3240 Then
If t4.Text = 5040 Then
p3.Left = p3.Left + 1800
free.Left = free.Left - 1800
moves.Text = moves.Text + 1
End If
End If
End If
End If
''''''''''right''''''''''

''''''''''left'''''''''''
If t1.Text = 3240 Then
If t2.Text = 5040 Then
If t3.Text = 5040 Then
If t4.Text = 5040 Then
p3.Left = p3.Left - 1800
free.Left = free.Left + 1800
moves.Text = moves.Text + 1
End If
End If
End If
End If
If t1.Text = 3240 Then
If t2.Text = 3240 Then
If t3.Text = 5040 Then
If t4.Text = 3240 Then
p3.Left = p3.Left - 1800
free.Left = free.Left + 1800
moves.Text = moves.Text + 1
End If
End If
End If
End If
If t1.Text = 3240 Then
If t2.Text = 1440 Then
If t3.Text = 5040 Then
If t4.Text = 1440 Then
p3.Left = p3.Left - 1800
free.Left = free.Left + 1800
moves.Text = moves.Text + 1
End If
End If
End If
End If
If t1.Text = 1440 Then
If t2.Text = 5040 Then
If t3.Text = 3240 Then
If t4.Text = 5040 Then
p3.Left = p3.Left - 1800
free.Left = free.Left + 1800
moves.Text = moves.Text + 1
End If
End If
End If
End If
If t1.Text = 1440 Then
If t2.Text = 3240 Then
If t3.Text = 3240 Then
If t4.Text = 3240 Then
p3.Left = p3.Left - 1800
free.Left = free.Left + 1800
moves.Text = moves.Text + 1
End If
End If
End If
End If
If t1.Text = 1440 Then
If t2.Text = 1440 Then
If t3.Text = 3240 Then
If t4.Text = 1440 Then
p3.Left = p3.Left - 1800
free.Left = free.Left + 1800
moves.Text = moves.Text + 1
End If
End If
End If
End If
''''''''''left'''''''''

''''''''''up''''''''''
If t1.Text = 5040 Then
If t2.Text = 3240 Then
If t3.Text = 5040 Then
If t4.Text = 5040 Then
p3.Top = p3.Top - 1800
free.Top = free.Top + 1800
moves.Text = moves.Text + 1
End If
End If
End If
End If
If t1.Text = 3240 Then
If t2.Text = 3240 Then
If t3.Text = 3240 Then
If t4.Text = 5040 Then
p3.Top = p3.Top - 1800
free.Top = free.Top + 1800
moves.Text = moves.Text + 1
End If
End If
End If
End If
If t1.Text = 1440 Then
If t2.Text = 3240 Then
If t3.Text = 1440 Then
If t4.Text = 5040 Then
p3.Top = p3.Top - 1800
free.Top = free.Top + 1800
moves.Text = moves.Text + 1
End If
End If
End If
End If
If t1.Text = 1440 Then
If t2.Text = 1440 Then
If t3.Text = 1440 Then
If t4.Text = 3240 Then
p3.Top = p3.Top - 1800
free.Top = free.Top + 1800
moves.Text = moves.Text + 1
End If
End If
End If
End If
If t1.Text = 3240 Then
If t2.Text = 1440 Then
If t3.Text = 3240 Then
If t4.Text = 3240 Then
p3.Top = p3.Top - 1800
free.Top = free.Top + 1800
moves.Text = moves.Text + 1
End If
End If
End If
End If
If t1.Text = 5040 Then
If t2.Text = 1440 Then
If t3.Text = 5040 Then
If t4.Text = 3240 Then
p3.Top = p3.Top - 1800
free.Top = free.Top + 1800
moves.Text = moves.Text + 1
End If
End If
End If
End If
''''''''''''up'''''''''

'''''''''down'''''''''
If t1.Text = 1440 Then
If t2.Text = 3240 Then
If t3.Text = 1440 Then
If t4.Text = 1440 Then
p3.Top = p3.Top + 1800
free.Top = free.Top - 1800
moves.Text = moves.Text + 1
End If
End If
End If
End If
If t1.Text = 3240 Then
If t2.Text = 3240 Then
If t3.Text = 3240 Then
If t4.Text = 1440 Then
p3.Top = p3.Top + 1800
free.Top = free.Top - 1800
moves.Text = moves.Text + 1
End If
End If
End If
End If
If t1.Text = 5040 Then
If t2.Text = 3240 Then
If t3.Text = 5040 Then
If t4.Text = 1440 Then
p3.Top = p3.Top + 1800
free.Top = free.Top - 1800
moves.Text = moves.Text + 1
End If
End If
End If
End If
If t1.Text = 1440 Then
If t2.Text = 5040 Then
If t3.Text = 1440 Then
If t4.Text = 3240 Then
p3.Top = p3.Top + 1800
free.Top = free.Top - 1800
moves.Text = moves.Text + 1
End If
End If
End If
End If
If t1.Text = 3240 Then
If t2.Text = 5040 Then
If t3.Text = 3240 Then
If t4.Text = 3240 Then
p3.Top = p3.Top + 1800
free.Top = free.Top - 1800
moves.Text = moves.Text + 1
End If
End If
End If
End If
If t1.Text = 5040 Then
If t2.Text = 5040 Then
If t3.Text = 5040 Then
If t4.Text = 3240 Then
p3.Top = p3.Top + 1800
free.Top = free.Top - 1800
moves.Text = moves.Text + 1
End If
End If
End If
End If
End Sub







Private Sub Label1_Click()
tclose.Enabled = True
End Sub


Private Sub loadpic_Click()
Image12.Picture = Image9.Picture

CommonDialog1.ShowOpen
Text1.Text = CommonDialog1.FileName
Image9.Picture = LoadPicture(Text1.Text)
Image1.Picture = LoadPicture(Text1.Text)
Image2.Picture = LoadPicture(Text1.Text)
Image3.Picture = LoadPicture(Text1.Text)
Image4.Picture = LoadPicture(Text1.Text)
Image5.Picture = LoadPicture(Text1.Text)
Image6.Picture = LoadPicture(Text1.Text)
Image7.Picture = LoadPicture(Text1.Text)
Image8.Picture = LoadPicture(Text1.Text)
End Sub

Private Sub moves_Change()
Label5.Caption = moves.Text
End Sub

Private Sub playagain_Click()
Timer4.Enabled = True
start.Enabled = True
loadpic.Enabled = True
won.Visible = False
End Sub

Private Sub start_Click()
p1.Enabled = True
p2.Enabled = True
p3.Enabled = True
p4.Enabled = True
p5.Enabled = True
p6.Enabled = True
p7.Enabled = True
p8.Enabled = True
free.Enabled = True
gametime.Enabled = True
If stage.Text = "1" Then
p1.Left = 3240
p1.Top = 3240
p2.Left = 5040
p2.Top = 3240
p3.Left = 1440
p3.Top = 5040
p4.Left = 3240
p4.Top = 5040
p5.Left = 1440
p5.Top = 1440
p6.Left = 3240
p6.Top = 1440
p7.Left = 5040
p7.Top = 1440
p8.Left = 1440
p8.Top = 3240
free.Left = 5040
free.Top = 5040
End If
If stage.Text = "2" Then
p1.Left = 5040
p1.Top = 3240
p2.Left = 1440
p2.Top = 5040
p3.Left = 3240
p3.Top = 5040
p4.Left = 3240
p4.Top = 3240
p5.Left = 1440
p5.Top = 3240
p6.Left = 1440
p6.Top = 1440
p7.Left = 3240
p7.Top = 1440
p8.Left = 5040
p8.Top = 1440
free.Left = 5040
free.Top = 5040
End If
If stage.Text = "3" Then
p1.Left = 3240
p1.Top = 5040

p2.Left = 3240
p2.Top = 3240

p3.Left = 5040
p3.Top = 3240

p4.Left = 1440
p4.Top = 5040

p5.Left = 3240
p5.Top = 1440

p6.Left = 5040
p6.Top = 1440

p7.Left = 1440
p7.Top = 3240

p8.Left = 1440
p8.Top = 1440

free.Left = 5040
free.Top = 5040
End If
If stage.Text = "4" Then
p1.Left = 3240
p1.Top = 1440

p2.Left = 1440
p2.Top = 1440

p3.Left = 1440
p3.Top = 3240

p4.Left = 5040
p4.Top = 1440

p5.Left = 5040
p5.Top = 3240

p6.Left = 3240
p6.Top = 3240

p7.Left = 3240
p7.Top = 5040

p8.Left = 1440
p8.Top = 5040

free.Left = 5040
free.Top = 5040
End If
If stage.Text = "5" Then
p1.Left = 3240
p1.Top = 1440

p2.Left = 1440
p2.Top = 1440

p3.Left = 1440
p3.Top = 3240

p4.Left = 5040
p4.Top = 1440

p5.Left = 3240
p5.Top = 5040

p6.Left = 1440
p6.Top = 5040

p7.Left = 5040
p7.Top = 3240

p8.Left = 3240
p8.Top = 3240

free.Left = 5040
free.Top = 5040
End If
If stage.Text = "6" Then
p1.Left = 5040
p1.Top = 1440

p2.Left = 1440
p2.Top = 3240

p3.Left = 1440
p3.Top = 1440

p4.Left = 3240
p4.Top = 1440

p5.Left = 1440
p5.Top = 5040

p6.Left = 3240
p6.Top = 5040

p7.Left = 3240
p7.Top = 3240

p8.Left = 5040
p8.Top = 3240

free.Left = 5040
free.Top = 5040
End If
If stage.Text = "7" Then
p1.Left = 1440
p1.Top = 3240

p2.Left = 5040
p2.Top = 1440

p3.Left = 3240
p3.Top = 1440

p4.Left = 1440
p4.Top = 1440

p5.Left = 5040
p5.Top = 3240

p6.Left = 3240
p6.Top = 3240

p7.Left = 3240
p7.Top = 5040

p8.Left = 1440
p8.Top = 5040

free.Left = 5040
free.Top = 5040
End If
If stage.Text = "8" Then
p1.Left = 1440
p1.Top = 5040

p2.Left = 3240
p2.Top = 5040

p3.Left = 3240
p3.Top = 3240

p4.Left = 5040
p4.Top = 3240

p5.Left = 5040
p5.Top = 1440

p6.Left = 1440
p6.Top = 3240

p7.Left = 1440
p7.Top = 1440

p8.Left = 3240
p8.Top = 1440

free.Left = 5040
free.Top = 5040
End If
start.Enabled = False
Timer4.Enabled = False
loadpic.Enabled = False
done.Enabled = True
End Sub


Private Sub tclose_Timer()
Me.Top = Me.Top + 120
Me.Height = Me.Height - 240
If Me.Height < 240 Then
End
End If

End Sub

Private Sub time_Change()
Label2.Caption = time.Text
End Sub

Private Sub Timer1_Timer()
If p1.Left > 1440 Then
p1.Left = p1.Left - 120
End If
If p1.Top > 1440 Then
p1.Top = p1.Top - 120
End If

If p2.Left < 3240 Then
p2.Left = p2.Left + 120
End If
If p2.Left > 3240 Then
p2.Left = p2.Left - 120
End If
If p2.Top > 1440 Then
p2.Top = p2.Top - 120
End If

If p3.Left < 5040 Then
p3.Left = p3.Left + 120
End If
If p3.Top > 1440 Then
p3.Top = p3.Top - 120
End If

If p4.Top < 3240 Then
p4.Top = p4.Top + 120
End If
If p4.Top > 3240 Then
p4.Top = p4.Top - 120
End If
If p4.Left > 1440 Then
p4.Left = p4.Left - 120
End If

If p5.Left > 3240 Then
p5.Left = p5.Left - 120
End If
If p5.Left < 3240 Then
p5.Left = p5.Left + 120
End If
If p5.Top > 3240 Then
p5.Top = p5.Top - 120
End If
If p5.Top < 3240 Then
p5.Top = p5.Top + 120
End If

If p6.Left < 5040 Then
p6.Left = p6.Left + 120
End If
If p6.Top > 3240 Then
p6.Top = p6.Top - 120
End If
If p6.Top < 3240 Then
p6.Top = p6.Top + 120
End If

If p7.Top < 5040 Then
p7.Top = p7.Top + 120
End If
If p7.Left > 1440 Then
p7.Left = p7.Left - 120
End If

If p8.Left < 3240 Then
p8.Left = p8.Left + 120
End If
If p8.Left > 3240 Then
p8.Left = p8.Left - 120
End If
If p8.Top < 5040 Then
p8.Top = p8.Top + 120
End If

If free.Left < 5040 Then
free.Left = free.Left + 120
End If
If free.Top < 5040 Then
free.Top = free.Top + 120
End If
If p1.Top = 1440 Then
If p1.Left = 1440 Then
If p2.Top = 1440 Then
If p2.Left = 3240 Then
If p3.Top = 1440 Then
If p3.Left = 5040 Then
If p4.Top = 3240 Then
If p4.Left = 1440 Then
If p5.Top = 3240 Then
If p5.Left = 3240 Then
If p6.Top = 3240 Then
If p6.Left = 5040 Then
If p7.Top = 5040 Then
If p7.Left = 1440 Then
If p8.Top = 5040 Then
If p8.Left = 3240 Then
If free.Top = 5040 Then
If free.Left = 5040 Then
Timer1.Enabled = False
start.Enabled = True
loadpic.Enabled = True
done.Enabled = False
gametime.Enabled = False
time.Text = "0"
moves.Text = "0"
Timer4.Enabled = True
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End Sub


Private Sub Timer2_Timer()
If p1.Top = 1440 Then
If p1.Left = 1440 Then
If p2.Top = 1440 Then
If p2.Left = 3240 Then
If p3.Top = 1440 Then
If p3.Left = 5040 Then
If p4.Top = 3240 Then
If p4.Left = 1440 Then
If p5.Top = 3240 Then
If p5.Left = 3240 Then
If p6.Top = 3240 Then
If p6.Left = 5040 Then
If p7.Top = 5040 Then
If p7.Left = 1440 Then
If p8.Top = 5040 Then
If p8.Left = 3240 Then
If free.Top = 5040 Then
If free.Left = 5040 Then
If done.Enabled = True Then
won.Visible = True
p1.Enabled = False
p2.Enabled = False
p3.Enabled = False
p4.Enabled = False
p5.Enabled = False
p6.Enabled = False
p7.Enabled = False
p8.Enabled = False
free.Enabled = False
gametime.Enabled = False
Label3.Caption = "Time : " + time.Text + " Sec"
Label4.Caption = "Moves : " + moves.Text
done.Enabled = False
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End Sub

Private Sub Timer3_Timer()
If Text1.Text = "" Then
Image1.Picture = Image12.Picture
Image2.Picture = Image12.Picture
Image3.Picture = Image12.Picture
Image4.Picture = Image12.Picture
Image5.Picture = Image12.Picture
Image6.Picture = Image12.Picture
Image7.Picture = Image12.Picture
Image8.Picture = Image12.Picture
Image9.Picture = Image12.Picture

End If
End Sub

Private Sub Timer4_Timer()
If stage.Text < 9 Then
stage.Text = stage.Text + 1
If stage.Text = 9 Then
stage.Text = 1
End If
End If
End Sub

Private Sub Timer5_Timer()
Picture1.Visible = False
Picture2.Visible = False
Timer5.Enabled = False
End Sub
