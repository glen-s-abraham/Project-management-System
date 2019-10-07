VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frm_edit_project 
   ClientHeight    =   6345
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12390
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   Picture         =   "frm_edit_project.frx":0000
   ScaleHeight     =   6345
   ScaleWidth      =   12390
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame 
      BackColor       =   &H00808000&
      BorderStyle     =   0  'None
      Height          =   6975
      Left            =   2280
      TabIndex        =   1
      Top             =   840
      Width           =   13095
      Begin VB.TextBox txt_pid 
         DataSource      =   "db_con"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   8160
         TabIndex        =   6
         Text            =   "nil"
         Top             =   720
         Visible         =   0   'False
         Width           =   3375
      End
      Begin VB.CommandButton cmd_add 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
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
         Height          =   615
         Left            =   840
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   5400
         Width           =   2295
      End
      Begin VB.CommandButton cmd_remove 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         Caption         =   "Remove Project"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3600
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   5400
         Width           =   2175
      End
      Begin VB.TextBox search 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         ForeColor       =   &H80000005&
         Height          =   375
         Left            =   3120
         TabIndex        =   3
         Top             =   720
         Width           =   4215
      End
      Begin MSDataGridLib.DataGrid pro_grid 
         Bindings        =   "frm_edit_project.frx":24A41
         Height          =   3615
         Left            =   840
         TabIndex        =   2
         Top             =   1440
         Width           =   11295
         _ExtentX        =   19923
         _ExtentY        =   6376
         _Version        =   393216
         AllowUpdate     =   0   'False
         Appearance      =   0
         BackColor       =   16776960
         BorderStyle     =   0
         ForeColor       =   8421504
         HeadLines       =   1
         RowHeight       =   24
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Search a Project"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000004&
         Height          =   495
         Left            =   840
         TabIndex        =   7
         Top             =   720
         Width           =   2655
      End
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   " Project Settings"
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
Attribute VB_Name = "frm_edit_project"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim pid As New ADODB.Recordset
Dim rec As New ADODB.Recordset
Dim resqty As New ADODB.Recordset

Dim cur_rec As New ADODB.Recordset

Public Sub selectall()
cmd2 = "select * from projects"
Set rec = db_con.con.Execute(cmd2)
Set pro_grid.DataSource = rec
End Sub

Public Sub fetchproject()
Dim cmd As String
cmd = "select pro_id from projects where pro_name='"
cmd = cmd + cmb_proj.Text
cmd = cmd + "'"
Set pid = db_con.con.Execute(cmd)
txt_pid.Text = pid!pro_id
Dim cmd2 As String
cmd2 = "select * from projects where pro_id="
cmd2 = cmd2 + txt_pid.Text
Set rec = db_con.con.Execute(cmd2)
Set pro_grid.DataSource = rec
End Sub

Private Sub cmb_proj_Click()
If cmb_proj.Text = "---Select All---" Then
selectall
Else
fetchproject
End If
End Sub

Private Sub close_Click()
Unload Me
End Sub

Private Sub cmd_add_Click()
If txt_pid.Text = "nil" Or txt_pid.Text = " " Or Not IsNumeric(txt_pid.Text) Then
    MsgBox ("select a project")
Else
    frm_update_project.Show
End If

End Sub

Private Sub cmd_remove_Click()
If txt_pid.Text = "nil" Or txt_pid.Text = " " Or Not IsNumeric(txt_pid.Text) Then
    MsgBox ("select a project")
Else

        
a = MsgBox("Are You Sure You Want To Delete employee : " & pro_grid.Columns(1), vbYesNo)

If a = 6 Then
cmd2 = "delete from tasks where task_project="
        cmd2 = cmd2 + txt_pid.Text
        db_con.con.Execute (cmd2)

cmd3 = " update employees set emp_project=0 where emp_project="
cmd3 = cmd3 + txt_pid.Text
db_con.con.Execute (cmd3)

cmd3 = "select res_name,qty_inuse from resources where res_project="
cmd3 = cmd3 + txt_pid.Text
Set resqty = db_con.con.Execute(cmd3)
If Not resqty.BOF = True Or Not resqty.EOF = True Then
qty = resqty!qty_inuse
rname = resqty!res_name
cmd4 = "delete from resources where res_project= "
cmd4 = cmd4 + txt_pid.Text
db_con.con.Execute (cmd4)

cmd4 = "update company_res set qty_inuse = qty_inuse-" + CStr(qty)
cmd4 = cmd4 + " where res_name='"
cmd4 = cmd4 + rname + "'"
db_con.con.Execute (cmd4)
End If

cmd = "delete from projects where pro_id="
cmd = cmd + txt_pid.Text
db_con.con.Execute (cmd)


selectall

End If
txt_pid.Text = "nil"
End If

End Sub

Private Sub Form_Load()
center Frame
selectall

End Sub

Private Sub Frame1_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

Private Sub pro_grid_RowColChange(LastRow As Variant, ByVal LastCol As Integer)

    If Not pro_grid.ApproxCount = 0 Then
        pro_grid.MarqueeStyle = dbgHighlightRow

        txt_pid.Text = pro_grid.Columns(0)
    Else
        txt_pid.Text = "nil"
    End If



End Sub

Private Sub search_Change()
If Len(search) > 0 Then
cmd2 = "select * from projects where pro_id like '"
cmd2 = cmd2 + search.Text + "%'"
cmd2 = cmd2 + " or pro_name like '"
cmd2 = cmd2 + search.Text + "%'"
cmd2 = cmd2 + " or pro_head like '"
cmd2 = cmd2 + search.Text + "%'"
cmd2 = cmd2 + " or client_name like '"
cmd2 = cmd2 + search.Text + "%'"



Set rec = db_con.con.Execute(cmd2)
Set pro_grid.DataSource = rec
Else
selectall
End If
End Sub
