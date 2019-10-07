VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm_project_add 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Add Project"
   ClientHeight    =   9675
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7635
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frm_project_add.frx":0000
   ScaleHeight     =   9675
   ScaleWidth      =   7635
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   1920
      Top             =   7920
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=pms;Data Source=Admin"
      OLEDBString     =   "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=pms;Data Source=Admin"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSComCtl2.DTPicker project_deadline 
      Height          =   375
      Left            =   1440
      TabIndex        =   18
      Top             =   5400
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd-mm-yyyy"
      Format          =   63832065
      CurrentDate     =   43287
   End
   Begin VB.PictureBox correct 
      BorderStyle     =   0  'None
      Height          =   375
      Index           =   3
      Left            =   6000
      Picture         =   "frm_project_add.frx":234CF
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   17
      Top             =   5400
      Width           =   375
   End
   Begin VB.PictureBox correct 
      BorderStyle     =   0  'None
      Height          =   375
      Index           =   2
      Left            =   6000
      Picture         =   "frm_project_add.frx":26CCF
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   16
      Top             =   4080
      Width           =   375
   End
   Begin VB.PictureBox correct 
      BorderStyle     =   0  'None
      Height          =   375
      Index           =   1
      Left            =   6000
      Picture         =   "frm_project_add.frx":2A4CF
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   15
      Top             =   2640
      Width           =   375
   End
   Begin VB.PictureBox correct 
      BorderStyle     =   0  'None
      Height          =   375
      Index           =   0
      Left            =   6000
      Picture         =   "frm_project_add.frx":2DCCF
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   14
      Top             =   1200
      Width           =   375
   End
   Begin VB.PictureBox Caution 
      BorderStyle     =   0  'None
      Height          =   375
      Index           =   3
      Left            =   6000
      Picture         =   "frm_project_add.frx":314CF
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   13
      ToolTipText     =   "Invalid Date"
      Top             =   5400
      Width           =   375
   End
   Begin VB.PictureBox Caution 
      BorderStyle     =   0  'None
      Height          =   375
      Index           =   2
      Left            =   6000
      Picture         =   "frm_project_add.frx":34EB9
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   12
      Top             =   4080
      Width           =   375
   End
   Begin VB.PictureBox Caution 
      BorderStyle     =   0  'None
      Height          =   375
      Index           =   1
      Left            =   6000
      Picture         =   "frm_project_add.frx":388A3
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   11
      Top             =   2640
      Width           =   375
   End
   Begin VB.PictureBox Caution 
      BorderStyle     =   0  'None
      Height          =   375
      Index           =   0
      Left            =   6000
      Picture         =   "frm_project_add.frx":3C28D
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   10
      ToolTipText     =   "Field Should Contain a Value."
      Top             =   1200
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Clear Fields"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   475
      Left            =   4080
      TabIndex        =   4
      Top             =   6480
      Width           =   2175
   End
   Begin VB.CommandButton cmd_submit 
      Appearance      =   0  'Flat
      BackColor       =   &H000040C0&
      Caption         =   "Create Project"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   475
      Left            =   1440
      MaskColor       =   &H00808080&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6480
      UseMaskColor    =   -1  'True
      Width           =   2535
   End
   Begin VB.TextBox txt_client_name 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1320
      TabIndex        =   2
      ToolTipText     =   "The Client For Whom The pRoject Is Being Developed"
      Top             =   4080
      Width           =   4500
   End
   Begin VB.TextBox txt_project_head 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1320
      TabIndex        =   1
      ToolTipText     =   "The Leading Faculty Of The Project  "
      Top             =   2640
      Width           =   4500
   End
   Begin VB.TextBox txt_project_title 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1320
      TabIndex        =   0
      ToolTipText     =   "The Title Of The New Project"
      Top             =   1200
      Width           =   4500
   End
   Begin VB.Label close 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000009&
      Height          =   375
      Left            =   7320
      TabIndex        =   9
      Top             =   0
      Width           =   495
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Deadline"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   255
      Left            =   1440
      TabIndex        =   8
      Top             =   4920
      Width           =   1095
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Client Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   255
      Left            =   1320
      TabIndex        =   7
      Top             =   3600
      Width           =   2415
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Project Head"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   255
      Left            =   1320
      TabIndex        =   6
      Top             =   2160
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Project Title"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   255
      Left            =   1320
      TabIndex        =   5
      Top             =   720
      Width           =   1575
   End
End
Attribute VB_Name = "frm_project_add"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection

Private Sub close_Click()
x = MsgBox("Are you sure You Want To quit Form?", vbYesNo)
If x = 6 Then
frm_project_add.Hide
End If

End Sub

Private Sub cmd_submit_Click()
Dim f As Boolean
f = False
For i = 0 To 3
f = Caution(i).Visible
If f = True Then
Exit For
End If
Next i

If f = True Or Len(txt_project_title) = 0 Or Len(txt_client_name) = 0 Or Len(txt_project_head) = 0 Then
x = MsgBox("Ivalid or Empty DataFields Encountered", vbInformation)
Else
a = MsgBox("Create New Project?", vbYesNo)
If a = 6 Then
Dim cmd As String
cmd = "insert into projects values('" & txt_project_title.Text & "','" & txt_project_head.Text & "','" & txt_client_name.Text & "','" & project_deadline.Value & "')"
MsgBox (cmd)
con.ConnectionString = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=pms;Data Source=Admin"
con.ConnectionString = cmd
con.Open
con.Execute cmd
End If
End If


End Sub

Private Sub Command1_Click()
x = MsgBox("Are You Sure That You want to Clear the fields? ", vbYesNo)
If x = 6 Then
For i = 0 To 3
Caution(i).Visible = False
correct(i).Visible = False
Next i
txt_project_title.Text = Empty
txt_project_head.Text = Empty
txt_client_name.Text = Empty
End If


End Sub

Private Sub Form_Load()
For i = 0 To 3
Caution(i).Visible = False
correct(i).Visible = False
Next i





End Sub

Private Sub txt_client_name_LostFocus()
If Len(txt_client_name.Text) = 0 Then
Caution(2).Visible = True
correct(2).Visible = False
Else
correct(2).Visible = True
Caution(2).Visible = False
End If
End Sub

Private Sub txt_project_head_LostFocus()
If Len(txt_project_head.Text) = 0 Then
Caution(1).Visible = True
correct(1).Visible = False
Else
correct(1).Visible = True
Caution(1).Visible = False
End If

End Sub

Private Sub txt_project_title_LostFocus()
If Len(txt_project_title.Text) = 0 Then
Caution(0).Visible = True
correct(0).Visible = False
Else
correct(0).Visible = True
Caution(0).Visible = False
End If







End Sub
