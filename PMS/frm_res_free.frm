VERSION 5.00
Begin VB.Form frm_res_free 
   Caption         =   "Free Resources"
   ClientHeight    =   9735
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7665
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   Picture         =   "frm_res_free.frx":0000
   ScaleHeight     =   9735
   ScaleWidth      =   7665
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame 
      BackColor       =   &H00808000&
      BorderStyle     =   0  'None
      Height          =   5895
      Left            =   7320
      TabIndex        =   1
      Top             =   2880
      Width           =   6375
      Begin VB.CommandButton cmd_add 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         Caption         =   "Free Resource"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   720
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   4320
         Width           =   2055
      End
      Begin VB.ComboBox cmb_qty 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   720
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   3120
         Width           =   4575
      End
      Begin VB.TextBox cmb_res 
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
         ForeColor       =   &H80000007&
         Height          =   375
         Left            =   720
         TabIndex        =   2
         Top             =   1440
         Width           =   4620
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Choose Quantity"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   495
         Left            =   480
         TabIndex        =   6
         Top             =   2400
         Width           =   2175
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   " Resource "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   495
         Left            =   480
         TabIndex        =   5
         Top             =   840
         Width           =   2175
      End
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   " Free Resources"
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
Attribute VB_Name = "frm_res_free"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim res As ADODB.Recordset
Dim cur_res As ADODB.Recordset
Dim pro_res As ADODB.Recordset
Dim tot_qty  As Integer
Dim avl_qty As Integer
Dim qty_use As Integer

Private Sub fetchres()
cmb_qty.Clear

cmd1 = "select * from resources where res_project="
cmd1 = cmd1 + frm_manage_res.txt_pid.Text + "and"
cmd1 = cmd1 + " res_name='" + cmb_res.Text + "'"
Set cur_res = db_con.con.Execute(cmd1)

If Not cur_res.RecordCount = 0 Then
qty_use = cur_res!qty_inuse

For i = 1 To qty_use
    cmb_qty.AddItem i
Next i
    cmb_qty.ListIndex = 0
Else
MsgBox ("no entity")

End If
End Sub

Private Sub close_Click()
Unload Me
frm_manage_res.resfetch

End Sub

Private Sub cmd_add_Click()
Dim cmd2 As String

cmd2 = "update resources set qty_inuse=qty_inuse - "
cmd2 = cmd2 + cmb_qty.Text + " where res_project="
cmd2 = cmd2 + frm_manage_res.txt_pid.Text + "and"
cmd2 = cmd2 + " res_name='" + cmb_res.Text + "'"
db_con.con.Execute (cmd2)
cmd2 = "update company_res set qty_inuse = qty_inuse -"
cmd2 = cmd2 + cmb_qty.Text + " where res_name='"
cmd2 = cmd2 + cmb_res.Text + "'"
db_con.con.Execute (cmd2)

cmd2 = "delete from resources where qty_inuse=0 and res_project="
cmd2 = cmd2 + frm_manage_res.txt_pid.Text + "and"
cmd2 = cmd2 + " res_name='" + cmb_res.Text + "'"
db_con.con.Execute (cmd2)

frm_manage_res.resfetch

Unload Me



End Sub

Private Sub Form_Load()
center Frame

cmb_res.Text = frm_manage_res.rname.Text
fetchres
End Sub

Private Sub Form_Unload(Cancel As Integer)
frm_manage_res.resfetch
frm_manage_res.rname.Text = "nil"

End Sub

