VERSION 5.00
Begin VB.Form frmTest 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "GenSoft MySQL Provider Library v2.0"
   ClientHeight    =   3525
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5310
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3525
   ScaleWidth      =   5310
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text6 
      Height          =   285
      Left            =   2700
      TabIndex        =   15
      Text            =   "3306"
      Top             =   480
      Width           =   2535
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00BCC7C5&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   630
      Left            =   120
      ScaleHeight     =   630
      ScaleWidth      =   5145
      TabIndex        =   12
      Top             =   2835
      Width           =   5145
      Begin VB.CommandButton cmdTest2 
         Caption         =   "Use MyReader"
         Enabled         =   0   'False
         Height          =   375
         Left            =   3000
         TabIndex        =   14
         Top             =   120
         Width           =   2055
      End
      Begin VB.CommandButton cmdTest1 
         Caption         =   "Convert ADO Recordset"
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   120
         Width           =   2055
      End
   End
   Begin VB.CommandButton cmdDisconnect 
      Caption         =   "Diconnect"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4275
      TabIndex        =   11
      Top             =   2385
      Width           =   975
   End
   Begin VB.CommandButton cmdConnect 
      Caption         =   "Connect"
      Height          =   375
      Left            =   3195
      TabIndex        =   10
      Top             =   2385
      Width           =   975
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   2700
      TabIndex        =   8
      Top             =   2040
      Width           =   2535
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   2700
      TabIndex        =   6
      Text            =   "gurevin_test"
      Top             =   1680
      Width           =   2535
   End
   Begin VB.TextBox Text3 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   2700
      PasswordChar    =   "*"
      TabIndex        =   4
      Text            =   "test"
      Top             =   1320
      Width           =   2535
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   2700
      TabIndex        =   2
      Text            =   "gurevin_test"
      Top             =   960
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   2700
      TabIndex        =   1
      Text            =   "www.gurevin.net"
      Top             =   120
      Width           =   2535
   End
   Begin VB.Label Label6 
      BackColor       =   &H009AA5AB&
      Caption         =   " MySQL Server Port Number"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   120
      TabIndex        =   16
      Top             =   480
      Width           =   2535
   End
   Begin VB.Label Label5 
      BackColor       =   &H009AA5AB&
      Caption         =   " Character Set (Optionale)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   120
      TabIndex        =   9
      Top             =   2040
      Width           =   2535
   End
   Begin VB.Label Label4 
      BackColor       =   &H009AA5AB&
      Caption         =   " Database (Optionale)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   120
      TabIndex        =   7
      Top             =   1680
      Width           =   2535
   End
   Begin VB.Label Label3 
      BackColor       =   &H009AA5AB&
      Caption         =   " MySQL Password"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   120
      TabIndex        =   5
      Top             =   1320
      Width           =   2535
   End
   Begin VB.Label Label2 
      BackColor       =   &H009AA5AB&
      Caption         =   " MySQL Username"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   2535
   End
   Begin VB.Label Label1 
      BackColor       =   &H009AA5AB&
      Caption         =   " MySQL Server IP or Name"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2535
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Written by Mehmet Gürevin :)

Option Explicit

Private WithEvents MyCNN                        As MyConnection
Attribute MyCNN.VB_VarHelpID = -1

Private Sub cmdConnect_Click()
    If (MyCNN.Connect(Text1.Text, Text2.Text, Text3.Text, Text4.Text, Val(Text6.Text)) = False) Then
        Call MsgBox("Error: Connection Failed!")
    End If
    If (Text5.Text <> "") Then
        MyCNN.Charset = Text5
    End If
End Sub

Private Sub cmdDisconnect_Click()
    Call MyCNN.Disconnect
End Sub

Private Sub cmdTest1_Click()
    Dim mRS                     As ADOR.Recordset
    
    Set mRS = New ADOR.Recordset
        Call MyCNN.ConvertADORS(MyCNN.Execute("SELECT * FROM test;"), mRS)
        
        Do While Not mRS.EOF
            Call MsgBox(mRS(0))
            Call mRS.MoveNext
        Loop
    Set mRS = Nothing
End Sub

Private Sub Command1_Click()
    Dim mReader                 As MyReader
    
    Set mReader = New MyReader
    
    Set mReader = MyCNN.ExecuteReader("SELECT * FROM test;")
        Do While mReader.Read
            MsgBox mReader.GetValue(0).Value
        Loop
    Set mReader = Nothing
End Sub

Private Sub Form_Load()
    Set MyCNN = New MyConnection
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set MyCNN = Nothing
End Sub

Private Sub MyCNN_Connected(ByVal lAPIHandle As Long, ByVal sHost As String, ByVal sUser As String, ByVal sPass As String, ByVal sDatabase As String, ByVal lPort As String, ByVal sUnixSocket As String)
    Call MsgBox("Connected OK.")
    cmdConnect.Enabled = False
    cmdDisconnect.Enabled = True
    cmdTest1.Enabled = True
    cmdTest2.Enabled = True
    Text1.Enabled = False
    Text2.Enabled = False
    Text3.Enabled = False
    Text4.Enabled = False
    Text5.Enabled = False
    Text6.Enabled = False
End Sub

Private Sub MyCNN_Disconnected(ByVal lAPIHandle As Long, ByVal sHost As String)
    Call MsgBox("Disconnected OK.")
    cmdConnect.Enabled = True
    cmdDisconnect.Enabled = False
    cmdTest1.Enabled = False
    cmdTest2.Enabled = False
    Text1.Enabled = True
    Text2.Enabled = True
    Text3.Enabled = True
    Text4.Enabled = True
    Text5.Enabled = True
    Text6.Enabled = True
End Sub
