VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm_update_project 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Update Project"
   ClientHeight    =   7710
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5925
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   Picture         =   "frm_update_project.frx":0000
   ScaleHeight     =   7710
   ScaleWidth      =   5925
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame 
      BackColor       =   &H00808000&
      BorderStyle     =   0  'None
      Height          =   7815
      Left            =   8040
      TabIndex        =   1
      Top             =   1680
      Width           =   6255
      Begin VB.ComboBox cmb_status 
         BackColor       =   &H00C0C0C0&
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
         Height          =   420
         ItemData        =   "frm_update_project.frx":24A41
         Left            =   960
         List            =   "frm_update_project.frx":24A4B
         TabIndex        =   19
         Top             =   6240
         Width           =   4500
      End
      Begin VB.PictureBox correct 
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   3
         Left            =   5520
         Picture         =   "frm_update_project.frx":24A62
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   18
         Top             =   4920
         Width           =   375
      End
      Begin VB.PictureBox correct 
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   2
         Left            =   5520
         Picture         =   "frm_update_project.frx":28262
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   17
         Top             =   3600
         Width           =   375
      End
      Begin VB.PictureBox correct 
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   1
         Left            =   5520
         Picture         =   "frm_update_project.frx":2BA62
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   16
         Top             =   2160
         Width           =   375
      End
      Begin VB.PictureBox correct 
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   0
         Left            =   5520
         Picture         =   "frm_update_project.frx":2F262
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   15
         Top             =   720
         Width           =   375
      End
      Begin VB.PictureBox Caution 
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   3
         Left            =   5520
         Picture         =   "frm_update_project.frx":32A62
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   14
         ToolTipText     =   "Invalid Date"
         Top             =   4920
         Width           =   375
      End
      Begin VB.PictureBox Caution 
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   2
         Left            =   5520
         Picture         =   "frm_update_project.frx":3644C
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   13
         Top             =   3600
         Width           =   375
      End
      Begin VB.PictureBox Caution 
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   1
         Left            =   5520
         Picture         =   "frm_update_project.frx":39E36
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   12
         Top             =   2160
         Width           =   375
      End
      Begin VB.PictureBox Caution 
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   0
         Left            =   5520
         Picture         =   "frm_update_project.frx":3D820
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   11
         ToolTipText     =   "Field Should Contain a Value."
         Top             =   720
         Width           =   375
      End
      Begin VB.CommandButton cmd_submit 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         Caption         =   "Update Project"
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
         Left            =   960
         MaskColor       =   &H00808080&
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   6960
         UseMaskColor    =   -1  'True
         Width           =   2535
      End
      Begin VB.TextBox txt_client_name 
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
         ForeColor       =   &H80000009&
         Height          =   375
         Left            =   840
         TabIndex        =   4
         ToolTipText     =   "The Client For Whom The pRoject Is Being Developed"
         Top             =   3600
         Width           =   4500
      End
      Begin VB.TextBox txt_project_head 
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
         ForeColor       =   &H80000009&
         Height          =   375
         Left            =   840
         TabIndex        =   3
         ToolTipText     =   "The Leading Faculty Of The Project  "
         Top             =   2160
         Width           =   4500
      End
      Begin VB.TextBox txt_project_title 
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
         ForeColor       =   &H80000009&
         Height          =   375
         Left            =   840
         TabIndex        =   2
         ToolTipText     =   "The Title Of The New Project"
         Top             =   720
         Width           =   4500
      End
      Begin MSComCtl2.DTPicker project_deadline 
         Height          =   375
         Left            =   960
         TabIndex        =   5
         Top             =   4920
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
         CalendarBackColor=   12632256
         CalendarTitleBackColor=   12632256
         CustomFormat    =   "dd-mm-yyyy"
         Format          =   16515073
         CurrentDate     =   43287
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Status"
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
         Height          =   255
         Left            =   960
         TabIndex        =   20
         Top             =   5640
         Width           =   1095
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
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   960
         TabIndex        =   9
         Top             =   4440
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
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   840
         TabIndex        =   8
         Top             =   3120
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
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   840
         TabIndex        =   7
         Top             =   1680
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
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   840
         TabIndex        =   6
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   " Update Project Details"
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
Attribute VB_Name = "frm_update_project"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cur_rec As New ADODB.Recordset

Private Sub close_Click()
Unload Me
End Sub

Private Sub cmd_submit_Click()


If Len(txt_project_title) = 0 Or Len(txt_client_name) = 0 Or Len(txt_project_head) = 0 Then
X = MsgBox("Ivalid or Empty DataFields Encountered", vbInformation)
Else
a = MsgBox("Update Project?", vbYesNo)
If a = 6 Then
cmd = "update projects   set pro_name= '" & Trim$(txt_project_title.Text) & "',pro_head= '" & Trim$(txt_project_head.Text) & "',client_name= '" & Trim$(txt_client_name.Text) & "',deadline= '" & Trim$(project_deadline.Value) & "' " & ",cur_status='" & Trim$(cmb_status.Text) & "'"
cmd = cmd + " where pro_id=" + frm_edit_project.txt_pid.Text

db_con.con.Execute (UCase(cmd))
MsgBox ("Project Updated")
Else
pro_title = (cur_rec!pro_name)
head = (cur_rec!pro_head)
client = (cur_rec!client_name)
txt_project_title.Text = Replace(pro_title, " ", "")
txt_project_head.Text = RTrim(head)
txt_client_name.Text = RTrim(client)
project_deadline.Value = RTrim(cur_rec!deadline)

End If
End If
End Sub

Private Sub Form_Load()
For i = 0 To 3
Caution(i).Visible = False
correct(i).Visible = False
Next i
cmd = "select * from projects where pro_id="
cmd = cmd + frm_edit_project.txt_pid.Text
Set cur_rec = db_con.con.Execute(cmd)


If cur_rec.EOF = False And cur_rec.BOF = False Then
txt_project_title.Text = cur_rec!pro_name
txt_project_head.Text = cur_rec!pro_head
txt_client_name.Text = cur_rec!client_name
project_deadline.Value = cur_rec!deadline
cmb_status.Text = cur_rec!cur_status
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
frm_edit_project.selectall



End Sub










Private Sub txt_client_name_KeyPress(KeyAscii As Integer)
k = KeyAscii
If (k >= 96 And k <= 123) Or (k >= 65 And k <= 91) Or (k = 8) Or (k = 32) Then
txt_client_name.Locked = False
Else
txt_client_name.Locked = True
End If
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

Private Sub txt_client_name_Validate(Cancel As Boolean)
If Len(txt_project_head.Text) = 0 Then
Caution(1).Visible = True
correct(1).Visible = False
Else
correct(1).Visible = True
Caution(1).Visible = False
End If
End Sub

Private Sub txt_project_head_KeyPress(KeyAscii As Integer)
k = KeyAscii
If (k >= 96 And k <= 123) Or (k >= 65 And k <= 91) Or (k = 8) Or (k = 32) Then
txt_project_head.Locked = False
Else
txt_project_head.Locked = True
End If



End Sub

Private Sub txt_project_head_Validate(Cancel As Boolean)
If Len(txt_project_head.Text) = 0 Then
Caution(1).Visible = True
correct(1).Visible = False

Else
correct(1).Visible = True
Caution(1).Visible = False
End If
End Sub

Private Sub txt_project_title_KeyPress(KeyAscii As Integer)
k = KeyAscii
If (k >= 48 And k <= 57) Or (k >= 96 And k <= 123) Or (k >= 65 And k <= 91) Or (k = 8) Or (k = 32) Then
txt_project_title.Locked = False
Else
txt_project_title.Locked = True
End If

End Sub

Private Sub txt_project_title_Validate(Cancel As Boolean)
If Len(txt_project_title.Text) = 0 Or (txt_project_title.Text = " ") Then
Caution(0).Visible = True
correct(0).Visible = False
Else
correct(0).Visible = True
Caution(0).Visible = False
End If
End Sub

