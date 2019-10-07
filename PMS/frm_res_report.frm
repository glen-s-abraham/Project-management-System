VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frm_res_report 
   Caption         =   "Resource Report"
   ClientHeight    =   6840
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13305
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "frm_res_report.frx":0000
   ScaleHeight     =   6840
   ScaleWidth      =   13305
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame 
      BackColor       =   &H00808000&
      BorderStyle     =   0  'None
      Height          =   7815
      Left            =   1440
      TabIndex        =   1
      Top             =   960
      Width           =   13815
      Begin VB.TextBox txt_title 
         Height          =   615
         Left            =   8160
         TabIndex        =   6
         Text            =   "Text1"
         Top             =   6240
         Visible         =   0   'False
         Width           =   2535
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
         Left            =   8520
         TabIndex        =   5
         Text            =   "nil"
         Top             =   840
         Visible         =   0   'False
         Width           =   3375
      End
      Begin VB.TextBox search 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         ForeColor       =   &H80000005&
         Height          =   375
         Left            =   3480
         TabIndex        =   4
         Top             =   840
         Width           =   4215
      End
      Begin VB.CommandButton cmd_report 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         Caption         =   "Generate Report"
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
         Left            =   1200
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   5640
         Width           =   2295
      End
      Begin VB.CommandButton Command1 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         Caption         =   "Generate All  Resource Report"
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
         Left            =   4080
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   5640
         Width           =   2295
      End
      Begin MSDataGridLib.DataGrid pro_grid 
         Bindings        =   "frm_res_report.frx":24A41
         Height          =   3615
         Left            =   1200
         TabIndex        =   7
         Top             =   1560
         Width           =   11175
         _ExtentX        =   19711
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
         Left            =   1200
         TabIndex        =   8
         Top             =   840
         Width           =   2655
      End
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Resource Report"
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
Attribute VB_Name = "frm_res_report"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim prec As New ADODB.Recordset
Dim rec As New ADODB.Recordset
Public Sub selectall()
cmd = "select * from resources where res_project=" + txt_pid.Text
Set rec = ReportEnv.rep_con.Execute(cmd)
End Sub
Public Sub fetchproject()
cmd = "select * from projects"
Set prec = ReportEnv.rep_con.Execute(cmd)
Set pro_grid.DataSource = prec

End Sub

Private Sub close_Click()
Unload Me

End Sub

Private Sub cmd_report_Click()
If Not txt_pid.Text = "nil" Or IsNumeric(txt_pid.Text) Then
selectall
Set Resource.DataSource = rec
Resource.Sections("Section4").Controls("lbl_project").Caption = txt_title.Text
Resource.Show
Else
MsgBox ("Choose a valid project")
End If

End Sub

Private Sub Command1_Click()
cmd = "select res_name,qty_inuse,tot_qty from company_res"
Set rec = ReportEnv.rep_con.Execute(cmd)
Set cmpny_res.DataSource = rec
cmpny_res.Show
End Sub

Private Sub Form_Load()
center Frame

fetchproject


End Sub

Private Sub pro_grid_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
pro_grid.MarqueeStyle = dbgHighlightRow

txt_pid.Text = pro_grid.Columns(0)
txt_title.Text = pro_grid.Columns(1)
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



Set rec = ReportEnv.rep_con.Execute(cmd2)
Set pro_grid.DataSource = rec
Else
fetchproject
End If
End Sub
