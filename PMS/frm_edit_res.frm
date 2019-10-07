VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frm_edit_res 
   Caption         =   "Resource Settings"
   ClientHeight    =   6630
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12855
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   Picture         =   "frm_edit_res.frx":0000
   ScaleHeight     =   6630
   ScaleWidth      =   12855
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame 
      BackColor       =   &H00808000&
      BorderStyle     =   0  'None
      Height          =   6255
      Left            =   4440
      TabIndex        =   1
      Top             =   3120
      Width           =   12615
      Begin VB.TextBox search 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         ForeColor       =   &H80000005&
         Height          =   375
         Left            =   3000
         TabIndex        =   5
         Top             =   480
         Width           =   4215
      End
      Begin VB.CommandButton cmd_remove 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         Caption         =   "Remove Resource"
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
         Left            =   3360
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   5160
         Width           =   2415
      End
      Begin VB.CommandButton cmd_add 
         Appearance      =   0  'Flat
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
         Height          =   615
         Left            =   600
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   5160
         Width           =   2295
      End
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
         Left            =   7920
         TabIndex        =   2
         Text            =   "nil"
         Top             =   480
         Visible         =   0   'False
         Width           =   3375
      End
      Begin MSDataGridLib.DataGrid res_grid 
         Bindings        =   "frm_edit_res.frx":24A41
         Height          =   3615
         Left            =   600
         TabIndex        =   6
         Top             =   1200
         Width           =   11295
         _ExtentX        =   19923
         _ExtentY        =   6376
         _Version        =   393216
         AllowUpdate     =   0   'False
         AllowArrows     =   0   'False
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
         Caption         =   "Search a Resource"
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
         Left            =   600
         TabIndex        =   7
         Top             =   480
         Width           =   2655
      End
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   " Resource Settings"
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
Attribute VB_Name = "frm_edit_res"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim res As New ADODB.Recordset
Dim rname As String

Public Sub fetchall()
cmd = "select res_id,res_name,tot_qty from company_res"
Set res = db_con.con.Execute(cmd)
Set res_grid.DataSource = res

End Sub

Private Sub close_Click()
Unload Me

End Sub

Private Sub cmd_add_Click()
If txt_pid.Text = "nil" Or txt_pid.Text = " " Or Not IsNumeric(txt_pid.Text) Then
    MsgBox ("select a Resource")
Else
    frm_update_res.Show
End If
End Sub

Private Sub cmd_remove_Click()
If txt_pid.Text = "nil" Or txt_pid.Text = " " Or Not IsNumeric(txt_pid.Text) Then
    MsgBox ("select a Resource")
Else
cmd = "delete from company_res where res_id="
cmd = cmd + txt_pid.Text
db_con.con.Execute (cmd)

cmd = "delete from resources where res_name='" + rname + "'"
db_con.con.Execute (cmd)


fetchall
txt_pid.Text = "nil"
End If
End Sub

Private Sub Form_Load()
center Frame
fetchall
End Sub

Private Sub res_grid_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
If Not res_grid.ApproxCount = 0 Then
        res_grid.MarqueeStyle = dbgHighlightRow
        txt_pid.Text = res_grid.Columns(0)
        rname = res_grid.Columns(1)
    Else
        txt_pid.Text = "nil"
        rname = "nil"

End If
End Sub

Private Sub search_Change()
If Len(search) > 0 Then
cmd2 = "select res_id,res_name,tot_qty from company_res where res_name like '"
cmd2 = cmd2 + search.Text + "%'"
cmd2 = cmd2 + " or res_id like '"
cmd2 = cmd2 + search.Text + "%'"

Set res = db_con.con.Execute(cmd2)
Set res_grid.DataSource = res
Else
fetchall
End If

End Sub
