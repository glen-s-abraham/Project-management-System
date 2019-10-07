VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frm_emp_add 
   Caption         =   "Add Employee To Project"
   ClientHeight    =   9735
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7680
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   Picture         =   "frm_emp_add.frx":0000
   ScaleHeight     =   9735
   ScaleWidth      =   7680
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame 
      BackColor       =   &H00808000&
      BorderStyle     =   0  'None
      Height          =   8655
      Left            =   6000
      TabIndex        =   1
      Top             =   1200
      Width           =   9615
      Begin VB.TextBox search 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   2760
         TabIndex        =   5
         Top             =   360
         Width           =   4215
      End
      Begin VB.TextBox txt_emp_id 
         Height          =   375
         Left            =   1200
         TabIndex        =   3
         Top             =   7080
         Visible         =   0   'False
         Width           =   3375
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00808080&
         Caption         =   "Add"
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
         Left            =   1320
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   7800
         Width           =   1455
      End
      Begin MSDataGridLib.DataGrid emp_grid 
         Bindings        =   "frm_emp_add.frx":24A41
         Height          =   5895
         Left            =   720
         TabIndex        =   4
         Top             =   1080
         Width           =   8175
         _ExtentX        =   14420
         _ExtentY        =   10398
         _Version        =   393216
         AllowUpdate     =   0   'False
         AllowArrows     =   0   'False
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
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Search Employee"
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
         TabIndex        =   6
         Top             =   360
         Width           =   2415
      End
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   " Assign Employee"
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
Attribute VB_Name = "frm_emp_add"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim pid As String
Dim emprec As ADODB.Recordset
Dim cmd As String

Public Sub empfetch()
cmd = "select emp_id,emp_name,emp_dep from employees where emp_project=0"
Set emprec = db_con.con.Execute(cmd)

End Sub

Private Sub close_Click()

Unload frm_emp_add

frm_manage_employee.empfetch

End Sub

Private Sub Command2_Click()

If Not txt_emp_id.Text = "nil" Then
    empid = txt_emp_id.Text
    cmd2 = "update employees set emp_project="
    cmd2 = cmd2 + pid
    cmd2 = cmd2 + " where emp_id="
    cmd2 = cmd2 + empid
    db_con.con.Execute (cmd2)
    
    emp_grid.ClearFields
    
    
    
    
Else
    MsgBox ("choose an employee")
End If




empfetch


End Sub

Private Sub emp_grid_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If Not emp_grid.ApproxCount = 0 Then
        emp_grid.MarqueeStyle = dbgHighlightRow
        txt_emp_id.Text = emp_grid.Columns(0)

    Else
        txt_emp_id.Text = "nil"
        MsgBox ("No employees available")
    End If




End Sub

Private Sub Form_Load()
center Frame
pid = frm_manage_employee.txt_pid.Text
txt_emp_id.Text = "nil"
empfetch
Set emp_grid.DataSource = emprec



End Sub

Private Sub Form_Unload(Cancel As Integer)
frm_manage_employee.empfetch
End Sub

Private Sub search_Change()
If Len(search) > 0 Then
cmd2 = "select emp_id,emp_name,emp_dep from employees where emp_id like '"
cmd2 = cmd2 + search.Text + "%'"
cmd2 = cmd2 + " or emp_name like '"
cmd2 = cmd2 + search.Text + "%'"
cmd2 = cmd2 + " or emp_adhaar like '"
cmd2 = cmd2 + search.Text + "%'"
cmd2 = cmd2 + " or emp_dep like '"
cmd2 = cmd2 + search.Text + "%'"



Set emprec = db_con.con.Execute(cmd2)
Set emp_grid.DataSource = emprec
Else
empfetch
Set emp_grid.DataSource = emprec

End If
End Sub
