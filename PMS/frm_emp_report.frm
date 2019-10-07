VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frm_emp_report 
   Caption         =   "Employee Report"
   ClientHeight    =   7140
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13035
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "frm_emp_report.frx":0000
   ScaleHeight     =   7140
   ScaleWidth      =   13035
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame 
      BackColor       =   &H00808000&
      BorderStyle     =   0  'None
      Height          =   8175
      Left            =   2640
      TabIndex        =   1
      Top             =   1680
      Width           =   13815
      Begin VB.TextBox txt_title 
         Height          =   615
         Left            =   8520
         TabIndex        =   9
         Text            =   "Text1"
         Top             =   6480
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
         Left            =   8880
         TabIndex        =   8
         Text            =   "nil"
         Top             =   1080
         Visible         =   0   'False
         Width           =   3375
      End
      Begin VB.TextBox search 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         ForeColor       =   &H80000005&
         Height          =   375
         Left            =   3840
         TabIndex        =   7
         Top             =   1080
         Visible         =   0   'False
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
         Left            =   1560
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   6600
         Width           =   2295
      End
      Begin VB.OptionButton opt_all 
         Caption         =   "Fetch All Employees"
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
         Left            =   1560
         MaskColor       =   &H000000FF&
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   5400
         Value           =   -1  'True
         Width           =   2415
      End
      Begin VB.OptionButton opt_available 
         Caption         =   "Available Employees"
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
         Left            =   4320
         MaskColor       =   &H000000FF&
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   5400
         Width           =   2535
      End
      Begin VB.OptionButton opt_busy 
         Caption         =   "Unavailable Employees"
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
         Left            =   7320
         MaskColor       =   &H000000FF&
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   5400
         Width           =   2535
      End
      Begin VB.OptionButton opt_projects 
         Caption         =   "Project Employees"
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
         Left            =   10320
         MaskColor       =   &H000000FF&
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   5400
         Width           =   2295
      End
      Begin MSDataGridLib.DataGrid pro_grid 
         Bindings        =   "frm_emp_report.frx":24A41
         Height          =   3615
         Left            =   1560
         TabIndex        =   10
         Top             =   1680
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
      Begin VB.Label lbl_search 
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
         Left            =   1560
         TabIndex        =   11
         Top             =   1080
         Visible         =   0   'False
         Width           =   2655
      End
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Employee Report"
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
Attribute VB_Name = "frm_emp_report"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rec As New ADODB.Recordset
Dim pro_rec As New ADODB.Recordset

Public Sub pro_fetch()
cmd = "select * from projects"
Set pro_rec = ReportEnv.rep_con.Execute(cmd)
Set pro_grid.DataSource = pro_rec


End Sub
Public Sub fetchall()
cmd = "select emp_adhaar,emp_name,emp_mail,emp_mobile,emp_dep from employees"
Set rec = ReportEnv.rep_con.Execute(cmd)
End Sub
Public Sub fetchavailable()
cmd = "select emp_adhaar,emp_name,emp_mail,emp_mobile,emp_dep from employees where emp_project=0"
Set rec = ReportEnv.rep_con.Execute(cmd)

End Sub
Public Sub fetchbusy()
cmd = "select emp_adhaar,emp_name,emp_mail,emp_mobile,emp_dep from employees where emp_project!=0"
Set rec = ReportEnv.rep_con.Execute(cmd)
End Sub
Public Sub fetchproject()
cmd = "select emp_adhaar,emp_name,emp_mail,emp_mobile,emp_dep from employees where emp_project=" + txt_pid.Text
Set rec = ReportEnv.rep_con.Execute(cmd)
End Sub


Private Sub close_Click()
Unload Me
End Sub

Private Sub cmd_report_Click()
If opt_all.Value = True Then
fetchall

ElseIf opt_available.Value = True Then
fetchavailable

ElseIf opt_busy.Value = True Then
fetchbusy

ElseIf opt_projects.Value = True Then
If Not txt_pid.Text = "nil" Or IsNumeric(txt_pid.Text) Then
fetchproject

Employee.Sections("Section4").Controls("lbl_title").Caption = "Employees Report for " + txt_title.Text

Else
MsgBox ("choose a valid project")
End If

End If
Set Employee.DataSource = rec
Employee.Show

End Sub

Private Sub Form_Load()
center Frame
End Sub

Private Sub opt_all_Click()
fetchall
search.Visible = False
lbl_search.Visible = False
Set pro_grid.DataSource = rec

End Sub

Private Sub opt_available_Click()
fetchavailable
search.Visible = False
lbl_search.Visible = False
Set pro_grid.DataSource = rec
End Sub

Private Sub opt_busy_Click()
fetchbusy
search.Visible = False
lbl_search.Visible = False
Set pro_grid.DataSource = rec
End Sub

Private Sub opt_projects_Click()
pro_fetch
search.Visible = True
lbl_search.Visible = True
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
pro_fetch

End If
End Sub
