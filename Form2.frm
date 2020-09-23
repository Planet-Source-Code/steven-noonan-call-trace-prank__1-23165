VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form2 
   Caption         =   "Trace"
   ClientHeight    =   3360
   ClientLeft      =   2985
   ClientTop       =   1215
   ClientWidth     =   9165
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   3360
   ScaleWidth      =   9165
   Begin MSComctlLib.ProgressBar ProgressBar2 
      Height          =   135
      Left            =   4575
      TabIndex        =   14
      Top             =   2640
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   238
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   8160
      Top             =   2040
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   2500
      Left            =   1920
      Top             =   120
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   2160
      Top             =   1200
   End
   Begin VB.TextBox Text12 
      Height          =   1935
      Left            =   135
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   13
      Text            =   "Form2.frx":0000
      Top             =   600
      Width           =   4095
   End
   Begin VB.TextBox n10 
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   3600
      TabIndex        =   9
      Top             =   120
      Width           =   315
   End
   Begin VB.TextBox n9 
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   3240
      TabIndex        =   8
      Top             =   120
      Width           =   315
   End
   Begin VB.TextBox n8 
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   2880
      TabIndex        =   7
      Top             =   120
      Width           =   315
   End
   Begin VB.TextBox n7 
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   2520
      TabIndex        =   6
      Top             =   120
      Width           =   315
   End
   Begin VB.TextBox n6 
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   2160
      TabIndex        =   5
      Top             =   120
      Width           =   315
   End
   Begin VB.TextBox n5 
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   1800
      TabIndex        =   4
      Top             =   120
      Width           =   315
   End
   Begin VB.TextBox n4 
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   1440
      TabIndex        =   3
      Top             =   120
      Width           =   315
   End
   Begin VB.TextBox n3 
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   1080
      TabIndex        =   2
      Top             =   120
      Width           =   315
   End
   Begin VB.TextBox n2 
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   720
      TabIndex        =   1
      Top             =   120
      Width           =   315
   End
   Begin VB.TextBox n1 
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   315
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Trace"
      Height          =   255
      Left            =   1995
      TabIndex        =   12
      Top             =   3000
      Width           =   5175
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   2415
      Left            =   4335
      TabIndex        =   10
      Top             =   120
      Width           =   135
      _ExtentX        =   238
      _ExtentY        =   4260
      _Version        =   393216
      Appearance      =   1
      Max             =   10
      Orientation     =   1
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   4575
      MultiLine       =   -1  'True
      TabIndex        =   11
      Top             =   120
      Visible         =   0   'False
      Width           =   4455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "0%"
      Height          =   255
      Left            =   135
      TabIndex        =   15
      Top             =   2640
      Width           =   4095
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a As String
Dim b As String
Dim c As String
Dim d As String
Dim e As String
Dim f As String
Dim g As String
Dim h As String
Dim i As String
Dim j As String

Private Sub Command1_Click()
Timer1.Enabled = True
Timer2.Enabled = True
End Sub

Private Sub Form_Load()
a = False
b = False
c = False
d = False
e = False
f = True
g = False
h = False
i = False
j = False
Text1.Text = Form1.Text1.Text & vbCrLf & Form1.Text2.Text & vbCrLf & Form1.Text3.Text & ", " & Form1.Text5.Text & " " & Form1.Text6.Text & vbCrLf & "(" & Form1.t1.Text & ") " & Form1.t2.Text & "-" & Form1.t3.Text
Text12.Text = "101010100101010010101001010100101011001010100010101001010010101001010101100110101001010100101010101010010101010100101010010101010110100010010010101001010100010101001010101001010001101011101010010110010100101001100101001010100100101010001010100101010010101001001010101111010001001100100110010100101010010101001001010010010101010100011010010010101010110111101010100010101010010100101010010101010100"
End Sub

Private Sub Form_Terminate()
Unload
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub n7_Change()
Timer1.Enabled = False
Timer2.Enabled = False
Timer3.Enabled = True
End Sub

Private Sub Timer1_Timer()
If Text12.Text = "010101010010101001010100101010010110100101010001010100101001010100101011010101010100101010010101010101001010101010010101001010101101010001001001010100101010001010100101010100101001010111010101001101001010010101010010100101010010010101000101010010101001010100100101011110101000101010010101001010010101001010100100101001001010101010010101001001010101101111010101010001010101001010010101001010101010" Then
Text12.Text = "101010100101010010101001010100101011001010100010101001010010101001010101100110101001010100101010101010010101010100101010010101010110100010010010101001010100010101001010101001010001101011101010010110010100101001100101001010100100101010001010100101010010101001001010101111010001001100100110010100101010010101001001010010010101010100011010010010101010110111101010100010101010010100101010010101010100"
Exit Sub
End If
If Text12.Text = "101010100101010010101001010100101011001010100010101001010010101001010101100110101001010100101010101010010101010100101010010101010110100010010010101001010100010101001010101001010001101011101010010110010100101001100101001010100100101010001010100101010010101001001010101111010001001100100110010100101010010101001001010010010101010100011010010010101010110111101010100010101010010100101010010101010100" Then
Text12.Text = "110111011011101101110110111011011100010111010101110110110101110110110110010101011010110101011101110110101110111011011101101110111000100101101101110110111010101110110111011010110101001101011101011000101101011101001011010111011011011101010111011011101101110110110110110011101010110001011000101101011101101110101010110101101110111011010101101101110111000111110111010101110110101101011101101110111010"
Exit Sub
End If
If Text12.Text = "110111011011101101110110111011011100010111010101110110110101110110110110010101011010110101011101110110101110111011011101101110111000100101101101110110111010101110110111011010110101001101011101011000101101011101001011010111011011011101010111011011101101110110110110110011101010110001011000101101011101101110101010110101101110111011010101101101110111000111110111010101110110101101011101101110111010" Then
Text12.Text = "010101010010101001010100101010010110100101010001010100101001010100101011010101010100101010010101010101001010101010010101001010101101010001001001010100101010001010100101010100101001010111010101001101001010010101010010100101010010010101000101010010101001010100100101011110101000101010010101001010010101001010100100101001001010101010010101001001010101101111010101010001010101001010010101001010101010"
Exit Sub
End If
End Sub

Private Sub Timer2_Timer()
ProgressBar1.Value = ProgressBar1.Value + 1
Label1.Caption = Int((ProgressBar1.Value / ProgressBar1.Max) * 100) & "%"
If f = True Then
    n6.Text = Mid(Form1.t2, 3, 1)
    a = True
    b = False
    c = False
    d = False
    e = False
    f = False
    g = False
    h = False
    i = False
    j = False
    Exit Sub
End If
If a = True Then
    n1.Text = Mid(Form1.t1, 1, 1)
    a = False
    b = False
    c = True
    d = False
    e = False
    f = False
    g = False
    h = False
    i = False
    j = False
    Exit Sub
End If
If c = True Then
    n3.Text = Mid(Form1.t1, 3, 1)
    a = False
    b = False
    c = False
    d = False
    e = False
    f = False
    g = False
    h = False
    i = True
    j = False
    Exit Sub
End If
If i = True Then
    n9.Text = Mid(Form1.t3, 3, 1)
    a = False
    b = False
    c = False
    d = True
    e = False
    f = False
    g = False
    h = False
    i = False
    j = False
    Exit Sub
End If
If d = True Then
    n4.Text = Mid(Form1.t2, 1, 1)
    a = False
    b = True
    c = False
    d = False
    e = False
    f = False
    g = False
    h = False
    i = False
    j = False
    Exit Sub
End If
If b = True Then
    n2.Text = Mid(Form1.t1, 2, 1)
    a = False
    b = False
    c = False
    d = False
    e = False
    f = False
    g = False
    h = True
    i = False
    j = False
    Exit Sub
End If
If h = True Then
    n8.Text = Mid(Form1.t3, 2, 1)
    a = False
    b = False
    c = False
    d = False
    e = False
    f = False
    g = False
    h = False
    i = False
    j = True
    Exit Sub
End If
If j = True Then
    n10.Text = Mid(Form1.t3, 4, 1)
    a = False
    b = False
    c = False
    d = False
    e = True
    f = False
    g = False
    h = False
    i = False
    j = False
    Exit Sub
End If
If e = True Then
    n5.Text = Mid(Form1.t2, 2, 1)
    a = False
    b = False
    c = False
    d = False
    e = False
    f = False
    g = True
    h = False
    i = False
    j = False
    Exit Sub
End If
If g = True Then
    n7.Text = Mid(Form1.t3, 1, 1)
    a = False
    b = False
    c = False
    d = False
    e = False
    f = False
    g = False
    h = False
    i = False
    j = False
    Exit Sub
End If
End Sub

Private Sub Timer3_Timer()
ProgressBar2.Value = ProgressBar2.Value + 1
If ProgressBar2.Value = ProgressBar2.Max Then
    Timer3.Enabled = False
    Text1.Visible = True
    If Form1.Check1.Value = 1 Then
        frmBrowser.Show
        Command1.Enabled = False
    End If
End If
End Sub
