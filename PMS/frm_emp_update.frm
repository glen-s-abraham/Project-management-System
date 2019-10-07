VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frm_emp_update 
   Caption         =   "Edit Employee Details"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6300
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   Picture         =   "frm_emp_update.frx":0000
   ScaleHeight     =   8595
   ScaleWidth      =   6300
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame 
      BackColor       =   &H00808000&
      BorderStyle     =   0  'None
      Height          =   8175
      Left            =   8160
      TabIndex        =   1
      Top             =   1680
      Width           =   6255
      Begin VB.CommandButton cmd_add 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         Caption         =   "Update Details"
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
         Left            =   720
         MaskColor       =   &H00800000&
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   7320
         UseMaskColor    =   -1  'True
         Width           =   2415
      End
      Begin VB.TextBox emp_id 
         Height          =   405
         Left            =   3240
         TabIndex        =   22
         Text            =   "Text1"
         Top             =   7320
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.PictureBox Caution 
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   4
         Left            =   5280
         Picture         =   "frm_emp_update.frx":24A41
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   15
         ToolTipText     =   "Field Should Contain a Value."
         Top             =   6240
         Width           =   375
      End
      Begin VB.PictureBox correct 
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   4
         Left            =   5280
         Picture         =   "frm_emp_update.frx":2842B
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   14
         Top             =   6240
         Width           =   375
      End
      Begin VB.PictureBox Caution 
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   3
         Left            =   5280
         Picture         =   "frm_emp_update.frx":2BC2B
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   13
         ToolTipText     =   "Field Should Contain a Value."
         Top             =   4800
         Width           =   375
      End
      Begin VB.PictureBox correct 
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   3
         Left            =   5280
         Picture         =   "frm_emp_update.frx":2F615
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   12
         Top             =   4800
         Width           =   375
      End
      Begin VB.PictureBox Caution 
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   2
         Left            =   5280
         Picture         =   "frm_emp_update.frx":32E15
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   11
         ToolTipText     =   "Field Should Contain a Value."
         Top             =   3360
         Width           =   375
      End
      Begin VB.PictureBox correct 
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   2
         Left            =   5280
         Picture         =   "frm_emp_update.frx":367FF
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   10
         Top             =   3360
         Width           =   375
      End
      Begin VB.PictureBox Caution 
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   1
         Left            =   5280
         Picture         =   "frm_emp_update.frx":39FFF
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   9
         ToolTipText     =   "Field Should Contain a Value."
         Top             =   2040
         Width           =   375
      End
      Begin VB.PictureBox correct 
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   1
         Left            =   5280
         Picture         =   "frm_emp_update.frx":3D9E9
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   8
         Top             =   2040
         Width           =   375
      End
      Begin VB.PictureBox Caution 
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   0
         Left            =   5280
         Picture         =   "frm_emp_update.frx":411E9
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   7
         ToolTipText     =   "Ivalid Adhaar Number"
         Top             =   600
         Width           =   375
      End
      Begin VB.PictureBox correct 
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   0
         Left            =   5280
         Picture         =   "frm_emp_update.frx":44BD3
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   6
         Top             =   600
         Width           =   375
      End
      Begin VB.ComboBox cmb_dept 
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
         ForeColor       =   &H00000000&
         Height          =   420
         ItemData        =   "frm_emp_update.frx":483D3
         Left            =   600
         List            =   "frm_emp_update.frx":483E3
         TabIndex        =   5
         Text            =   "Choose Department"
         Top             =   6240
         Width           =   4500
      End
      Begin VB.TextBox txt_mail 
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
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   600
         TabIndex        =   4
         Top             =   4800
         Width           =   4500
      End
      Begin VB.TextBox txt_contact 
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
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   600
         MaxLength       =   10
         TabIndex        =   3
         Top             =   3360
         Width           =   4500
      End
      Begin VB.TextBox txt_name 
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
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   600
         TabIndex        =   2
         Top             =   2040
         Width           =   4500
      End
      Begin MSMask.MaskEdBox txt_adhaar 
         Height          =   375
         Left            =   600
         TabIndex        =   16
         Top             =   600
         Width           =   4500
         _ExtentX        =   7938
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   0
         BackColor       =   12632256
         ForeColor       =   -2147483642
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         MaxLength       =   14
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "####-####-####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Employee Department"
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
         Left            =   480
         TabIndex        =   21
         Top             =   5520
         Width           =   5295
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Employee E-mail ID"
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
         Left            =   480
         TabIndex        =   20
         Top             =   4200
         Width           =   5175
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Emplooyee Mobile Number "
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
         Left            =   480
         TabIndex        =   19
         Top             =   2760
         Width           =   5175
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Employee Name"
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
         Left            =   480
         TabIndex        =   18
         Top             =   1440
         Width           =   3255
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   " Adhaar Number"
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
         Left            =   480
         TabIndex        =   17
         Top             =   120
         Width           =   5175
      End
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   " Update Employee Details"
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
Attribute VB_Name = "frm_emp_update"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim erec As ADODB.Recordset

Private Sub close_Click()
Unload Me
End Sub

Private Sub cmd_add_Click()
Dim f As Boolean
f = False
For i = 0 To 3
f = Caution(i).Visible
If f = True Then
Exit For
End If
Next i

If f = True Or Len(txt_name) = 0 Or Len(txt_contact) = 0 Or Len(txt_mail) = 0 Or cmb_dept.Text = "Choose Department" Or Len(cmb_dept) = 0 Then
X = MsgBox("Ivalid or Empty DataFields Encountered", vbInformation)
Else
a = MsgBox("Are You Sure That You Want To Update The Conten", vbYesNo)
If a = 6 Then
cmd = "update employees set emp_adhaar='"
cmd = cmd + txt_adhaar.Text + "',"
cmd = cmd + "emp_name='"
cmd = cmd + UCase(txt_name.Text) + "',"
cmd = cmd + "emp_mobile='"
cmd = cmd + txt_contact.Text + "',"
cmd = cmd + "emp_mail='"
cmd = cmd + txt_mail.Text + "',"
cmd = cmd + "emp_dep='"
cmd = cmd + UCase(cmb_dept.Text) + "'"
cmd = cmd + "where emp_id = "
cmd = cmd + emp_id.Text

'MsgBox (cmd)
db_con.con.Execute (cmd)
MsgBox ("Data Updated")
End If

End If

End Sub

Private Sub Form_Load()
emp_id.Text = frm_emp_edit.txt_eid
cmd = "select * from employees where emp_id="
cmd = cmd + emp_id.Text
Set erec = db_con.con.Execute(cmd)
txt_adhaar.Text = erec!emp_adhaar
txt_name.Text = erec!emp_name
txt_contact.Text = erec!emp_mobile
txt_mail.Text = erec!emp_mail
cmb_dept.Text = erec!emp_dep
For i = 0 To 4
Caution(i).Visible = False
correct(i).Visible = False
Next i

dot = 0
at = 0

center Frame
End Sub

Private Sub Form_Unload(Cancel As Integer)
frm_emp_edit.selectall

End Sub

Private Sub cmb_dept_Validate(Cancel As Boolean)
If cmb_dept.Text = "Choose Department" Or Len(cmb_dept) = 0 Then
Caution(4).Visible = True
correct(4).Visible = False
Else
Caution(4).Visible = False
correct(4).Visible = True
End If
End Sub


Private Sub Command1_Click()
X = MsgBox("Are You Sure That You want to Clear the fields? ", vbYesNo)
If X = 6 Then
For i = 0 To 3
Caution(i).Visible = False
correct(i).Visible = False
Next i
For i = 0 To 2
txt_adhaar(i).Text = Empty
Next i

txt_name.Text = Empty
txt_contact.Text = Empty
txt_mail.Text = Empty
cmb_dept.Text = "Choose Department"

End If
End Sub







Private Sub txt_emp_mobile_Change()

End Sub

Private Sub txt_adhaar_Validate(Cancel As Boolean)
l = txt_adhaar.Text

If Len(l) < 12 Then
Caution(0).Visible = True
correct(0).Visible = False
Else
Caution(0).Visible = False
correct(0).Visible = True
End If
End Sub

Private Sub txt_contact_KeyPress(KeyAscii As Integer)
k = KeyAscii
If (k >= 48 And k <= 57) Or (k = 8) Or (k = 32) Then
txt_contact.Locked = False
Else
txt_contact.Locked = True
End If
End Sub



Private Sub txt_contact_Validate(Cancel As Boolean)
If Len(txt_contact) > 0 And Len(txt_contact) = 10 Then
Caution(2).Visible = False
correct(2).Visible = True
Else
Caution(2).Visible = True
correct(2).Visible = False
End If
End Sub

Private Sub txt_mail_KeyPress(KeyAscii As Integer)
k = KeyAscii
If (k >= 48 And k <= 57) Or (k >= 96 And k <= 123) Or (k >= 64 And k <= 91) Or (k = 8) Or (k = 32) Or (k = 46) Then
txt_mail.Locked = False
Else
txt_mail.Locked = True
End If

If k = 46 Then dot = dot + 1
If k = 64 Then at = at + 1

End Sub

Private Sub txt_mail_Validate(Cancel As Boolean)

If Len(txt_mail) = 0 Then

Caution(3).Visible = True
correct(3).Visible = False
Else
Caution(3).Visible = False
correct(3).Visible = True
End If
End Sub

Private Sub txt_name_KeyPress(KeyAscii As Integer)
k = KeyAscii
If (k >= 96 And k <= 123) Or (k >= 65 And k <= 91) Or (k = 8) Or (k = 32) Then
txt_name.Locked = False
Else
txt_name.Locked = True
End If
End Sub

Private Sub txt_name_Validate(Cancel As Boolean)
If Len(txt_name) = 0 Then
Caution(1).Visible = True
correct(1).Visible = False
Else
Caution(1).Visible = False
correct(1).Visible = True
End If
End Sub

