VERSION 5.00
Begin VB.Form frm_res_allocate 
   Caption         =   "Allocate Resource"
   ClientHeight    =   9765
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7650
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   Picture         =   "frm_res_allocate.frx":0000
   ScaleHeight     =   9765
   ScaleWidth      =   7650
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame 
      BackColor       =   &H00808000&
      BorderStyle     =   0  'None
      Height          =   6015
      Left            =   6720
      TabIndex        =   1
      Top             =   3000
      Width           =   5895
      Begin VB.ComboBox cmb_res 
         BackColor       =   &H00808080&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFC0&
         Height          =   420
         Left            =   600
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1680
         Width           =   4215
      End
      Begin VB.ComboBox cmb_qty 
         BackColor       =   &H00808080&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFC0&
         Height          =   420
         Left            =   600
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   3240
         Width           =   4215
      End
      Begin VB.CommandButton cmd_add 
         BackColor       =   &H00808080&
         Caption         =   "Add To Project"
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
         Left            =   600
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   4440
         Width           =   2055
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Choose Resource "
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
         Height          =   495
         Left            =   360
         TabIndex        =   6
         Top             =   960
         Width           =   2175
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
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   480
         TabIndex        =   5
         Top             =   2520
         Width           =   2175
      End
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   " Allocate Resources"
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
Attribute VB_Name = "frm_res_allocate"
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

Private Sub close_Click()
Unload Me
frm_manage_res.resfetch


End Sub

Private Sub cmb_res_Click()
cmb_qty.Clear

cmd1 = "select * from company_res where res_name='"
cmd1 = cmd1 + cmb_res.Text + "'"
Set cur_res = db_con.con.Execute(cmd1)
tot_qty = cur_res!tot_qty
qty_use = cur_res!qty_inuse
avl_qty = tot_qty - qty_use
For i = 1 To avl_qty
    cmb_qty.AddItem i
Next i
    
    

End Sub

Private Sub Combo1_Change()

End Sub

Private Sub cmd_add_Click()
If Not Len(cmb_res.Text) = 0 And Not Len(cmb_qty.Text) = 0 Then
Dim cmd3 As String
Dim cur_qty As Integer
cur_qty = cmb_qty.Text
cur_qty = cur_qty + qty_inuse


If cur_qty <= tot_qty Then

Set pro_res = db_con.con.Execute("select * from resources where res_name='" & cmb_res.Text & "' and res_project='" & frm_manage_res.txt_pid.Text & "'")

If pro_res.RecordCount = 0 Then
    cmd3 = "insert into resources values("
    cmd3 = cmd3 + frm_manage_res.txt_pid.Text
    cmd3 = cmd3 + ",'"
    cmd3 = cmd3 + cmb_res.Text
    cmd3 = cmd3 + "',"
    cmd3 = cmd3 + cmb_qty.Text
    cmd3 = cmd3 + ")"
    db_con.con.Execute (cmd3)
Else
    cmd3 = "update resources set qty_inuse=qty_inuse+"
    cmd3 = cmd3 + cmb_qty.Text
    cmd3 = cmd3 + " where res_name='"
    cmd3 = cmd3 + cmb_res.Text + "'"
    db_con.con.Execute (cmd3)
End If

cmd3 = "update company_res set qty_inuse=qty_inuse+"
cmd3 = cmd3 + cmb_qty.Text
cmd3 = cmd3 + " where res_name='"
cmd3 = cmd3 + cmb_res.Text + "'"
db_con.con.Execute (cmd3)
cmd = "select res_name from company_res"
Set res = db_con.con.Execute(cmd)

Else
MsgBox ("Resource outnumbered")
End If
Else
MsgBox ("please choose a resource and quantity")
End If
End Sub

Private Sub Form_Load()
center Frame

Dim cmd As String
cmd = "select res_name from company_res"
Set res = db_con.con.Execute(cmd)
Do Until res.EOF
    cmb_res.AddItem res!res_name
    res.MoveNext
Loop



End Sub

Private Sub Form_Unload(Cancel As Integer)
frm_manage_res.resfetch
End Sub
