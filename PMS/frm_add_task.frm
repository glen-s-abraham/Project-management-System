VERSION 5.00
Begin VB.Form frm_add_task 
   Caption         =   "Add New Task"
   ClientHeight    =   9735
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7635
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "frm_add_task.frx":0000
   ScaleHeight     =   9735
   ScaleWidth      =   7635
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame 
      BackColor       =   &H00808000&
      BorderStyle     =   0  'None
      Height          =   7335
      Left            =   8280
      TabIndex        =   1
      Top             =   2160
      Width           =   6855
      Begin VB.TextBox txt_tsk_title 
         BackColor       =   &H00C0C0C0&
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
         ForeColor       =   &H80000005&
         Height          =   375
         Left            =   600
         TabIndex        =   7
         Top             =   1320
         Width           =   4980
      End
      Begin VB.TextBox txt_tsk_description 
         BackColor       =   &H00C0C0C0&
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
         ForeColor       =   &H80000005&
         Height          =   2655
         Left            =   600
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   6
         Top             =   2880
         Width           =   5655
      End
      Begin VB.CommandButton cmd_add 
         BackColor       =   &H00808080&
         Caption         =   "Add Task To Project"
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
         Left            =   600
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   6240
         Width           =   3135
      End
      Begin VB.CommandButton cmd_clear 
         BackColor       =   &H00808080&
         Caption         =   "Clear Fileds"
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
         Left            =   3840
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   6240
         Width           =   2535
      End
      Begin VB.PictureBox Caution 
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   0
         Left            =   5760
         Picture         =   "frm_add_task.frx":24A41
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   3
         ToolTipText     =   "Field Should Contain a Value."
         Top             =   1320
         Width           =   375
      End
      Begin VB.PictureBox correct 
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   0
         Left            =   5760
         Picture         =   "frm_add_task.frx":2842B
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   2
         Top             =   1320
         Width           =   375
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Task Title"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   600
         TabIndex        =   9
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Task Description"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   600
         TabIndex        =   8
         Top             =   2280
         Width           =   2175
      End
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   " Add New Task"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5175
   End
End
Attribute VB_Name = "frm_add_task"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim pid As String

Private Sub close_Click()
X = MsgBox("Are you sure You Want To quit Form?", vbYesNo)
If X = 6 Then
Unload frm_add_task
End If


End Sub

Private Sub cmd_add_Click()
If Len(txt_tsk_title) = 0 Or Len(txt_tsk_description) = 0 Then
X = MsgBox("Ivalid or Empty DataFields Encountered", vbInformation)
Else
a = MsgBox("Add New task To Project?", vbYesNo)
If a = 6 Then
Dim cmd As String


cmd = "insert into tasks values("
cmd = cmd + pid
cmd = cmd + ",'"
cmd = cmd + txt_tsk_title.Text
cmd = cmd + "','"
cmd = cmd + txt_tsk_description.Text
cmd = cmd + "','PENDING')"


db_con.con.Execute UCase((cmd))
Unload Me




End If
End If

End Sub

Private Sub cmd_clear_Click()
X = MsgBox("Are You Sure That You want to Clear the fields? ", vbYesNo)
If X = 6 Then

Caution(0).Visible = False
correct(0).Visible = False

txt_tsk_title.Text = Empty
txt_tsk_description.Text = Empty

End If
End Sub

Private Sub Form_Load()
center Frame

pid = frm_task_manage.txt_pid.Text


Caution(0).Visible = False
correct(0).Visible = False

End Sub

Private Sub Text1_Change()

End Sub

Private Sub Form_Unload(Cancel As Integer)
frm_task_manage.taskfetch
End Sub

Private Sub txt_tsk_title_KeyPress(KeyAscii As Integer)
k = KeyAscii
If (k >= 48 And k <= 57) Or (k >= 96 And k <= 123) Or (k >= 65 And k <= 91) Or (k = 8) Or (k = 32) Then
txt_tsk_title.Locked = False
Else
txt_tsk_title.Locked = True
End If
End Sub

Private Sub txt_tsk_title_Validate(Cancel As Boolean)
If Len(txt_tsk_title) = 0 Then
Caution(0).Visible = True
correct(0).Visible = False
Else
Caution(0).Visible = False
correct(0).Visible = True
End If

End Sub
