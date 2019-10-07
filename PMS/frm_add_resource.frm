VERSION 5.00
Begin VB.Form frm_add_resource 
   Caption         =   "Resource"
   ClientHeight    =   9720
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7650
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Moveable        =   0   'False
   Picture         =   "frm_add_resource.frx":0000
   ScaleHeight     =   9720
   ScaleWidth      =   7650
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame 
      BackColor       =   &H00808000&
      BorderStyle     =   0  'None
      Height          =   6375
      Left            =   8760
      TabIndex        =   1
      Top             =   1200
      Width           =   6015
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
         TabIndex        =   9
         Top             =   1320
         Width           =   4815
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
         TabIndex        =   8
         Top             =   2760
         Width           =   2295
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00808080&
         Caption         =   "Add to Database"
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
         TabIndex        =   7
         Top             =   4080
         Width           =   2415
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00808080&
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
         Left            =   2880
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   4080
         Width           =   2295
      End
      Begin VB.PictureBox correct 
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   0
         Left            =   5400
         Picture         =   "frm_add_resource.frx":24A41
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   5
         Top             =   1320
         Width           =   375
      End
      Begin VB.PictureBox Caution 
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   0
         Left            =   5400
         Picture         =   "frm_add_resource.frx":28241
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   4
         ToolTipText     =   "Field Should Contain a Value."
         Top             =   1320
         Width           =   375
      End
      Begin VB.PictureBox Caution 
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   1
         Left            =   5400
         Picture         =   "frm_add_resource.frx":2BC2B
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   3
         ToolTipText     =   "Field Should Contain a Valid Numeric Value."
         Top             =   2760
         Width           =   375
      End
      Begin VB.PictureBox correct 
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   1
         Left            =   5400
         Picture         =   "frm_add_resource.frx":2F615
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   2
         Top             =   2760
         Width           =   375
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
         TabIndex        =   11
         Top             =   840
         Width           =   2295
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
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   " Add New Resource"
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
Attribute VB_Name = "frm_add_resource"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub close_Click()

X = MsgBox("Are you sure You Want To quit Form?", vbYesNo)
If X = 6 Then
Unload frm_add_resource
End If
End Sub

Private Sub Command1_Click()
Dim f As Boolean
f = False
For i = 0 To 1
f = Caution(i).Visible
If f = True Then
Exit For
End If
Next i

If f = True Or Len(txt_res_name) = 0 Or Len(txt_res_qty) = 0 Or IsNumeric(txt_res_qty) = False Then
X = MsgBox("Ivalid or Empty DataFields Encountered", vbInformation)
Else
a = MsgBox("Add New Resource To DataBase?", vbYesNo)
If a = 6 Then
Dim cmd As String
cmd = cmd + "insert into company_res values('"
cmd = cmd + UCase(txt_res_name.Text)
cmd = cmd + "',"
cmd = cmd + txt_res_qty.Text
cmd = cmd + ",0)"
'MsgBox (cmd)

db_con.con.Execute (cmd)
MsgBox ("New Resources Added")

For i = 0 To 1
Caution(i).Visible = False
correct(i).Visible = False
Next i
txt_res_name.Text = Empty
txt_res_qty.Text = Empty
End If
End If
End Sub

Private Sub Command2_Click()
X = MsgBox("Are You Sure That You want to Clear the fields? ", vbYesNo)
If X = 6 Then
For i = 0 To 1
Caution(i).Visible = False
correct(i).Visible = False
Next i
txt_res_name.Text = Empty
txt_res_qty.Text = Empty
End If
End Sub

Private Sub Form_Load()
center Frame

For i = 0 To 1
Caution(i).Visible = False
correct(i).Visible = False
Next i
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
