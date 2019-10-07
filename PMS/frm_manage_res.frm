VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frm_manage_res 
   Caption         =   "Resource Management"
   ClientHeight    =   6330
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12345
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   Picture         =   "frm_manage_res.frx":0000
   ScaleHeight     =   6330
   ScaleWidth      =   12345
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame 
      BackColor       =   &H00808000&
      BorderStyle     =   0  'None
      Height          =   6255
      Left            =   4440
      TabIndex        =   1
      Top             =   2520
      Width           =   12495
      Begin VB.ComboBox cmb_proj 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
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
         ForeColor       =   &H80000009&
         Height          =   420
         Left            =   3000
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   480
         Width           =   4215
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
         Left            =   9360
         TabIndex        =   5
         Text            =   "nil"
         Top             =   480
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.CommandButton cmd_add 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         Caption         =   "Add New Resource"
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
         Left            =   720
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   5160
         Width           =   2655
      End
      Begin VB.CommandButton cmd_remove 
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
         Height          =   615
         Left            =   3480
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   5160
         Width           =   2175
      End
      Begin VB.TextBox rname 
         Height          =   495
         Left            =   8280
         TabIndex        =   2
         Text            =   "nil"
         Top             =   5280
         Visible         =   0   'False
         Width           =   3135
      End
      Begin MSDataGridLib.DataGrid res_grid 
         Bindings        =   "frm_manage_res.frx":24A41
         Height          =   3615
         Left            =   600
         TabIndex        =   7
         Top             =   1200
         Width           =   11295
         _ExtentX        =   19923
         _ExtentY        =   6376
         _Version        =   393216
         BackColor       =   16776960
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
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Select a Project"
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
         Height          =   495
         Left            =   600
         TabIndex        =   8
         Top             =   480
         Width           =   1935
      End
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   " Manage Resource"
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
Attribute VB_Name = "frm_manage_res"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim pid As New ADODB.Recordset
Dim rec As New ADODB.Recordset
Dim res_rec As New ADODB.Recordset
Public Sub resfetch()
Dim cmd As String
cmd = "select pro_id from projects where pro_name='"
cmd = cmd + cmb_proj.Text
cmd = cmd + "'"
Set pid = db_con.con.Execute(cmd)
txt_pid.Text = pid!pro_id
Dim cmd2 As String
cmd2 = "select res_name,qty_inuse from resources where res_project="
cmd2 = cmd2 + txt_pid.Text
Set res_rec = db_con.con.Execute(cmd2)
Set res_grid.DataSource = res_rec

End Sub

Private Sub close_Click()
Unload Me

End Sub

Private Sub cmb_proj_Click()
resfetch
rname.Text = "nil"

End Sub

Private Sub cmd_add_Click()
If Not Len(cmb_proj.Text) = 0 Then
    frm_res_allocate.Show

Else
    MsgBox ("Choose A Valid Project")
End If
End Sub

Private Sub cmd_remove_Click()
If Not Not Len(cmb_proj.Text) = 0 Or txt_pid.Text = "nil" Or rname.Text = "nil" Then
    MsgBox ("select a Resource")
Else
    frm_res_free.Show


End If
End Sub

Private Sub Form_Load()
center Frame



Set rec = db_con.con.Execute("select pro_name from projects")


Do Until rec.EOF
    cmb_proj.AddItem rec!pro_name
    rec.MoveNext
Loop

End Sub

Private Sub res_grid_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
If Not Len(cmb_proj.Text) = 0 Then
    
    If Not res_grid.ApproxCount = 0 Then
        res_grid.MarqueeStyle = dbgHighlightRow
        rname.Text = res_grid.Columns(0)
    Else
        rname.Text = "nil"
    End If
Else
    MsgBox ("Choose A Valid Project")
End If

End Sub
