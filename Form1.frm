VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Configuration"
   ClientHeight    =   2940
   ClientLeft      =   6075
   ClientTop       =   4125
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2940
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox t3 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Left            =   2760
      MaxLength       =   4
      TabIndex        =   7
      Top             =   1935
      Width           =   495
   End
   Begin VB.TextBox t2 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Left            =   2280
      MaxLength       =   3
      TabIndex        =   6
      Top             =   1935
      Width           =   375
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Connected to the internet?"
      Height          =   255
      Left            =   53
      TabIndex        =   8
      Top             =   2243
      Width           =   4575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   255
      Left            =   1186
      TabIndex        =   9
      Top             =   2603
      Width           =   2415
   End
   Begin VB.TextBox Text6 
      Height          =   285
      Left            =   1800
      TabIndex        =   4
      Top             =   1560
      Width           =   2775
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   1800
      MaxLength       =   2
      TabIndex        =   3
      Top             =   1200
      Width           =   2775
   End
   Begin VB.TextBox t1 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Left            =   1800
      MaxLength       =   3
      TabIndex        =   5
      Top             =   1935
      Width           =   375
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   1800
      TabIndex        =   2
      Top             =   840
      Width           =   2775
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1800
      TabIndex        =   1
      Top             =   480
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1800
      TabIndex        =   0
      Top             =   120
      Width           =   2775
   End
   Begin VB.Label Label6 
      Caption         =   "Phone"
      Height          =   255
      Left            =   226
      TabIndex        =   15
      Top             =   1883
      Width           =   1815
   End
   Begin VB.Label Label5 
      Caption         =   "Zip"
      Height          =   255
      Left            =   226
      TabIndex        =   14
      Top             =   1523
      Width           =   1815
   End
   Begin VB.Label Label4 
      Caption         =   "State (e.g. VA)"
      Height          =   255
      Left            =   226
      TabIndex        =   13
      Top             =   1163
      Width           =   1815
   End
   Begin VB.Label Label3 
      Caption         =   "City"
      Height          =   255
      Left            =   226
      TabIndex        =   12
      Top             =   803
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "Address"
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   480
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Name"
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Text1.Text <> "" And Text2.Text <> "" And Text3.Text <> "" And Text5.Text <> "" And Text6.Text <> "" And t1.Text <> "" And t2.Text <> "" And t3.Text <> "" Then
    Me.Hide
    Form2.Show
    Else:
    MsgBox "Fill in ALL the boxes!", vbCritical, "Hold on there, Bronco!"
End If
End Sub

Private Sub Form_Load()
Text1.Text = GetSetting(App.Title, "CFG", "Name")
Text2.Text = GetSetting(App.Title, "CFG", "Address")
Text3.Text = GetSetting(App.Title, "CFG", "City")
Text5.Text = GetSetting(App.Title, "CFG", "State")
Text6.Text = GetSetting(App.Title, "CFG", "Zip")
t1.Text = GetSetting(App.Title, "CFG", "phone1")
t2.Text = GetSetting(App.Title, "CFG", "phone2")
t3.Text = GetSetting(App.Title, "CFG", "phone3")
End Sub

Private Sub t1_Change()
SaveSetting App.Title, "CFG", "phone1", t1.Text
End Sub

Private Sub t2_Change()
SaveSetting App.Title, "CFG", "phone2", t2.Text
End Sub

Private Sub t3_Change()
SaveSetting App.Title, "CFG", "phone3", t3.Text
End Sub

Private Sub Text1_Change()
SaveSetting App.Title, "CFG", "Name", Text1.Text
End Sub

Private Sub Text2_Change()
SaveSetting App.Title, "CFG", "Address", Text2.Text
End Sub

Private Sub Text3_Change()
SaveSetting App.Title, "CFG", "City", Text3.Text
End Sub

Private Sub Text5_Change()
SaveSetting App.Title, "CFG", "State", Text5.Text
End Sub

Private Sub Text6_Change()
SaveSetting App.Title, "CFG", "Zip", Text6.Text
End Sub
