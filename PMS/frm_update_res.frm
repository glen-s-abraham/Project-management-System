VERSION 5.00
Begin VB.Form frm_update_res 
   Caption         =   "Update Resource"
   ClientHeight    =   7200
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6195
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   Picture         =   "frm_update_res.frx":0000
   ScaleHeight     =   7200
   ScaleWidth      =   6195
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame 
      BackColor       =   &H00808000&
      BorderStyle     =   0  'None
      Height          =   6375
      Left            =   6600
      TabIndex        =   1
      Top             =   2160
      Width           =   6015
      Begin VB.TextBox rid 
         Height          =   495
         Left            =   360
         TabIndex        =   11
         Text            =   "Text1"
         Top             =   120
         Visible         =   0   'False
         Width           =   4215
      End
      Begin VB.PictureBox correct 
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   1
         Left            =   5400
         Picture         =   "frm_update_res.frx":24A41
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   8
         Top             =   2760
         Width           =   375
      End
      Begin VB.PictureBox Caution 
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   1
         Left            =   5400
         Picture         =   "frm_update_res.frx":28241
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   7
         ToolTipText     =   "Field Should Contain a Valid Numeric Value."
         Top             =   2760
         Width           =   375
      End
      Begin VB.PictureBox Caution 
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   0
         Left            =   5400
         Picture         =   "frm_update_res.frx":2BC2B
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   6
         ToolTipText     =   "Field Should Contain a Value."
         Top             =   1320
         Width           =   375
      End
      Begin VB.PictureBox correct 
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   0
         Left            =   5400
         Picture         =   "frm_update_res.frx":2F615
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   5
         Top             =   1320
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00808080&
         Caption         =   "Update Resource"
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
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   4680
         Width           =   2415
      End
      Begin VB.TextBox txt_res_qty 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0;(0)"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
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
         Left            =   360
         MaxLength       =   3
         TabIndex        =   3
         Top             =   2760
         Width           =   2295
      End
      Begin VB.TextBox txt_res_name 
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
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   360
         TabIndex        =   2
         Top             =   1320
         Width           =   4815
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Resource Quantity"
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
         Left            =   360
         TabIndex        =   10
         Top             =   2160
         Width           =   2295
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Resource Name"
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
         Left            =   360
         TabIndex        =   9
         Top             =   840
         Width           =   2295
      End
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   " Update Resource Details"
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
Attribute VB_Name = "frm_update_res"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim res As ADODB.Recordset

Private Sub close_Click()
Unload Me
frm_edit_res.Show

End Sub

Private Sub Command1_Click()
If f = True Or Len(txt_res_name) = 0 Or Len(txt_res_qty) = 0 Or IsNumeric(txt_res_qty) = False Then
X = MsgBox("Ivalid or Empty DataFields Encountered", vbInformation)
Else
a = MsgBox("Add New Resource To DataBase?", vbYesNo)
If a = 6 Then

cmd = cmd + "update company_res set res_name='"
cmd = cmd + UCase(txt_res_name.Text)
cmd = cmd + "',tot_qty="
cmd = cmd + txt_res_qty.Text
cmd = cmd + "where res_id="
cmd = cmd + rid.Text
db_con.con.Execute (cmd)
End If
End If

End Sub

Private Sub Form_Load()
center Frame

For i = 0 To 1
Caution(i).Visible = False
correct(i).Visible = False
Next i

rid.Text = frm_edit_res.txt_pid.Text
cmd = "select * from company_res where res_id="
cmd = cmd + rid.Text
Set rec = db_con.con.Execute(cmd)
If rec.EOF = False And rec.BOF = False Then
txt_res_name.Text = rec!res_name
txt_res_qty.Text = rec!tot_qty
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
frm_edit_res.fetchall

End Sub


Private Sub txt_res_name_KeyPress(KeyAscii As Integer)
k = KeyAscii
If (k >= 96 And k <= 123) Or (k >= 65 And k <= 91) Or (k = 8) Or (k = 32) Then
txt_res_name.Locked = False
Else
txt_res_name.Locked = True
End If
End Sub

Private Sub txt_res_name_Validate(Cancel As Boolean)
If Len(txt_res_name) = 0 Then
Caution(0).Visible = True
correct(0).Visible = False
Else
Caution(0).Visible = False
correct(0).Visible = True
End If
End Sub

Private Sub txt_res_qty_KeyPress(KeyAscii As Integer)
k = KeyAscii
If (k >= 48 And k <= 57) Or (k = 8) Or (k = 32) Then
txt_res_qty.Locked = False
Else
txt_res_qty.Locked = True
End If
End Sub

Private Sub txt_res_qty_Validate(Cancel As Boolean)
If Len(txt_res_qty) = 0 Then
Caution(1).Visible = True
correct(1).Visible = False
Else
Caution(1).Visible = False
correct(1).Visible = True
End If
End Sub

