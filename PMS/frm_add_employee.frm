VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frm_add_employee 
   Caption         =   "Add New Worker"
   ClientHeight    =   9570
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7650
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "frm_add_employee.frx":0000
   ScaleHeight     =   9570
   ScaleWidth      =   7650
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame 
      BackColor       =   &H00808000&
      BorderStyle     =   0  'None
      Height          =   8175
      Left            =   8040
      TabIndex        =   1
      Top             =   840
      Width           =   6255
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
         TabIndex        =   18
         Top             =   2040
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
         TabIndex        =   17
         Top             =   3360
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
         TabIndex        =   16
         Top             =   4800
         Width           =   4500
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
         ItemData        =   "frm_add_employee.frx":24A41
         Left            =   600
         List            =   "frm_add_employee.frx":24A51
         TabIndex        =   15
         Text            =   "Choose Department"
         Top             =   6240
         Width           =   4500
      End
      Begin VB.CommandButton cmd_add 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         Caption         =   "Add To Database"
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
         MaskColor       =   &H00800000&
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   7320
         UseMaskColor    =   -1  'True
         Width           =   2415
      End
      Begin VB.CommandButton Command1 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         Caption         =   "Reset Fields"
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
         Left            =   3120
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   7320
         Width           =   2535
      End
      Begin VB.PictureBox correct 
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   0
         Left            =   5280
         Picture         =   "frm_add_employee.frx":24A76
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   12
         Top             =   480
         Width           =   375
      End
      Begin VB.PictureBox Caution 
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   0
         Left            =   5280
         Picture         =   "frm_add_employee.frx":28276
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   11
         ToolTipText     =   "Ivalid Adhaar Number"
         Top             =   480
         Width           =   375
      End
      Begin VB.PictureBox correct 
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   1
         Left            =   5280
         Picture         =   "frm_add_employee.frx":2BC60
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   10
         Top             =   1920
         Width           =   375
      End
      Begin VB.PictureBox Caution 
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   1
         Left            =   5280
         Picture         =   "frm_add_employee.frx":2F460
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   9
         ToolTipText     =   "Field Should Contain a Value."
         Top             =   1920
         Width           =   375
      End
      Begin VB.PictureBox correct 
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   2
         Left            =   5280
         Picture         =   "frm_add_employee.frx":32E4A
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   8
         Top             =   3240
         Width           =   375
      End
      Begin VB.PictureBox Caution 
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   2
         Left            =   5280
         Picture         =   "frm_add_employee.frx":3664A
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   7
         ToolTipText     =   "Field Should Contain a Value."
         Top             =   3240
         Width           =   375
      End
      Begin VB.PictureBox correct 
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   3
         Left            =   5280
         Picture         =   "frm_add_employee.frx":3A034
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   6
         Top             =   4680
         Width           =   375
      End
      Begin VB.PictureBox Caution 
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   3
         Left            =   5280
         Picture         =   "frm_add_employee.frx":3D834
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   5
         ToolTipText     =   "Field Should Contain a Value."
         Top             =   4680
         Width           =   375
      End
      Begin VB.PictureBox correct 
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   4
         Left            =   5280
         Picture         =   "frm_add_employee.frx":4121E
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   4
         Top             =   6120
         Width           =   375
      End
      Begin VB.PictureBox Caution 
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   4
         Left            =   5280
         Picture         =   "frm_add_employee.frx":44A1E
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   3
         ToolTipText     =   "Field Should Contain a Value."
         Top             =   6120
         Width           =   375
      End
      Begin MSMask.MaskEdBox txt_adhaar 
         Height          =   375
         Left            =   600
         TabIndex        =   2
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
         TabIndex        =   23
         Top             =   120
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
         TabIndex        =   22
         Top             =   1440
         Width           =   3255
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
         TabIndex        =   21
         Top             =   2760
         Width           =   5175
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
         TabIndex        =   19
         Top             =   5520
         Width           =   5295
      End
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Add New Employee"
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
Attribute VB_Name = "frm_add_employee"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim dot As Integer
Dim at As Integer

Private Sub close_Click()
X = MsgBox("Are you sure You Want To quit Form?", vbYesNo)
If X = 6 Then
Unload frm_add_employee
End If
End Sub

Private Sub Combo1_Change()

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
a = MsgBox("Add New Employee To DataBase?", vbYesNo)
If a = 6 Then

Dim cmd As String
cmd = "insert into employees values(0,'"
cmd = cmd + txt_adhaar.Text
cmd = cmd + "','"
cmd = cmd + UCase(txt_name.Text)
cmd = cmd + "','"
cmd = cmd + txt_contact.Text
cmd = cmd + "','"
cmd = cmd + txt_mail.Text
cmd = cmd + "','"
cmd = cmd + UCase(cmb_dept.Text)
cmd = cmd + "')"
'MsgBox (cmd)


db_con.con.Execute (cmd)
For i = 0 To 3
Caution(i).Visible = False
correct(i).Visible = False
Next i

txt_adhaar.Text = Empty
txt_name.Text = Empty
txt_contact.Text = Empty
txt_mail.Text = Empty
cmb_dept.Text = "Choose Department"
correct(4).Visible = False



End If
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

Private Sub Form_Load()
For i = 0 To 4
Caution(i).Visible = False
correct(i).Visible = False
Next i

dot = 0
at = 0

center Frame


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
