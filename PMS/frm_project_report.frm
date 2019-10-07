VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frm_project_report 
   Caption         =   "Project Report"
   ClientHeight    =   8670
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12390
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "frm_project_report.frx":0000
   ScaleHeight     =   8670
   ScaleWidth      =   12390
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame 
      BackColor       =   &H00808000&
      BorderStyle     =   0  'None
      Height          =   8415
      Left            =   3360
      TabIndex        =   1
      Top             =   1680
      Width           =   12015
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
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   7080
         Width           =   2295
      End
      Begin VB.Frame frm_date 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   1575
         Left            =   360
         TabIndex        =   8
         Top             =   4920
         Width           =   11175
         Begin VB.CommandButton cmd_fetch 
            Appearance      =   0  'Flat
            BackColor       =   &H00808080&
            Caption         =   "Fetch Entries"
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
            Left            =   9360
            Style           =   1  'Graphical
            TabIndex        =   9
            Top             =   480
            Width           =   1575
         End
         Begin MSComCtl2.DTPicker dt_start 
            Height          =   375
            Left            =   1800
            TabIndex        =   10
            Top             =   600
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   661
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CalendarBackColor=   12632256
            CalendarTitleBackColor=   12632256
            CustomFormat    =   "mm-dd-yyyy"
            Format          =   16449537
            CurrentDate     =   43287
         End
         Begin MSComCtl2.DTPicker dt_end 
            Height          =   375
            Left            =   6480
            TabIndex        =   11
            Top             =   600
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   661
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CalendarBackColor=   12632256
            CalendarTitleBackColor=   12632256
            CustomFormat    =   "mm-dd-yyyy"
            Format          =   16449537
            CurrentDate     =   43287
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Start Date"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   360
            TabIndex        =   13
            Top             =   600
            Width           =   1455
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "End Date"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   5040
            TabIndex        =   12
            Top             =   600
            Width           =   1455
         End
      End
      Begin VB.OptionButton opt_date 
         Caption         =   "Date Range"
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
         Left            =   2640
         MaskColor       =   &H000000FF&
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   4200
         Width           =   1815
      End
      Begin VB.OptionButton opt_all 
         Caption         =   "Fetch All"
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
         Left            =   360
         MaskColor       =   &H000000FF&
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   4200
         Value           =   -1  'True
         Width           =   1575
      End
      Begin VB.OptionButton opt_due 
         Caption         =   "Due Projects"
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
         Left            =   5040
         MaskColor       =   &H000000FF&
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   4200
         Width           =   1815
      End
      Begin VB.OptionButton opt_complete 
         Caption         =   "Completed"
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
         Left            =   7440
         MaskColor       =   &H000000FF&
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   4200
         Width           =   1695
      End
      Begin VB.OptionButton opt_pendin 
         Caption         =   "Pending"
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
         Left            =   9720
         MaskColor       =   &H000000FF&
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   4200
         Width           =   1815
      End
      Begin VB.TextBox txt_date 
         Height          =   495
         Left            =   8640
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   7320
         Visible         =   0   'False
         Width           =   2175
      End
      Begin MSDataGridLib.DataGrid pro_grid 
         Bindings        =   "frm_project_report.frx":24A41
         Height          =   3615
         Left            =   360
         TabIndex        =   15
         Top             =   480
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
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   " Project Report"
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
Attribute VB_Name = "frm_project_report"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rec As New ADODB.Recordset
Public Sub pendingfetch()
cmd = "select * from projects where cur_status='PENDING'"
Set rec = ReportEnv.rep_con.Execute(cmd)
Set pro_grid.DataSource = rec
End Sub
Public Sub completefetch()
cmd = "select * from projects where cur_status='COMPLETE'"
Set rec = ReportEnv.rep_con.Execute(cmd)
Set pro_grid.DataSource = rec
End Sub
Public Sub duefetch()
cmd = "select * from projects where deadline < '" + txt_date.Text + "'"
Set rec = ReportEnv.rep_con.Execute(cmd)
Set pro_grid.DataSource = rec



End Sub
Public Sub datefetch()
cmd = "select * from projects where deadline between '"
cmd = cmd + CStr(dt_start.Value) + "' and '" + CStr(dt_end.Value) + "'"

Set rec = ReportEnv.rep_con.Execute(cmd)
Set pro_grid.DataSource = rec

End Sub
Public Sub fetchall()
cmd = "select * from projects"
Set rec = ReportEnv.rep_con.Execute(cmd)
Set pro_grid.DataSource = rec



End Sub

Private Sub Command1_Click()
projects.Show
End Sub

Private Sub close_Click()
Unload Me
End Sub

Private Sub cmd_remove_Click()

End Sub

Private Sub cmd_fetch_Click()
If opt_date.Value = True Then
datefetch
End If










End Sub

Private Sub cmd_report_Click()

Set projects.DataSource = rec
projects.DataMember = rec.DataMember

projects.Show



projects.Show


End Sub

Private Sub dt_end_Click()
If opt_date.Value = True Then
datefetch
End If
End Sub

Private Sub dt_start_Click()
If opt_date.Value = True Then
datefetch
End If
End Sub

Private Sub Form_Load()
center Frame
fetchall
txt_date.Text = Date

End Sub

Private Sub Option2_Click()

End Sub

Private Sub opt_all_Click()
If opt_all.Value = True Then
fetchall
End If
End Sub

Private Sub opt_complete_Click()
If opt_complete.Value = True Then
completefetch
End If
End Sub

Private Sub opt_due_Click()
If opt_due.Value = True Then
duefetch
End If
End Sub

Private Sub opt_pendin_Click()
If opt_pendin.Value = True Then
pendingfetch
End If
End Sub
