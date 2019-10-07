VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frm_manage_employee 
   Caption         =   "Employee management"
   ClientHeight    =   6660
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12525
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   Picture         =   "frm_manage_employee.frx":0000
   ScaleHeight     =   6660
   ScaleWidth      =   12525
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame 
      BackColor       =   &H00808000&
      BorderStyle     =   0  'None
      Height          =   7095
      Left            =   1800
      TabIndex        =   1
      Top             =   840
      Width           =   13095
      Begin VB.CommandButton cmd_remove 
         BackColor       =   &H00808080&
         Caption         =   "Remove Employee"
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
         Left            =   3840
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   5880
         Width           =   2535
      End
      Begin VB.CommandButton cmd_add 
         BackColor       =   &H00808080&
         Caption         =   "Add New Employee"
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
         Left            =   1080
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   5880
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
         Left            =   9720
         TabIndex        =   4
         Top             =   840
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.ComboBox cmb_proj 
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
         ForeColor       =   &H00FFFFFF&
         Height          =   420
         Left            =   3360
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   840
         Width           =   4215
      End
      Begin VB.TextBox emp_id 
         Height          =   615
         Left            =   7560
         TabIndex        =   2
         Text            =   "nil"
         Top             =   5760
         Visible         =   0   'False
         Width           =   1695
      End
      Begin MSDataGridLib.DataGrid emp_grid 
         Bindings        =   "frm_manage_employee.frx":24A41
         Height          =   3615
         Left            =   960
         TabIndex        =   7
         Top             =   1560
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
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   960
         TabIndex        =   8
         Top             =   840
         Width           =   1935
      End
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Manage  Employee "
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
Attribute VB_Name = "frm_manage_employee"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim pid As New ADODB.Recordset
Dim employees As New ADODB.Recordset
Dim rec As New ADODB.Recordset

Public Sub empfetch()
Dim cmd As String
cmd = "select pro_id from projects where pro_name='"
cmd = cmd + cmb_proj.Text
cmd = cmd + "'"
Set pid = db_con.con.Execute(cmd)
txt_pid.Text = pid!pro_id
Dim cmd2 As String
cmd2 = "select emp_id,emp_name,emp_mobile,emp_mail,emp_dep from employees where emp_project="
cmd2 = cmd2 + txt_pid.Text
Set employees = db_con.con.Execute(cmd2)
Set emp_grid.DataSource = employees
End Sub
Private Sub cmd_refresh_Click()
End Sub

Private Sub close_Click()
Unload Me

End Sub

Private Sub cmb_proj_Click()
empfetch
emp_id.Text = "nil"

End Sub

Private Sub cmd_add_Click()
If Not Len(cmb_proj.Text) = 0 Then
    frm_emp_add.Show
Else
    MsgBox ("Choose A Valid Project")
End If

End Sub

Private Sub cmd_remove_Click()
If Not Len(cmb_proj.Text) = 0 Then
    If Not emp_id.Text = "nil" Then
        
        a = MsgBox("Are You Sure You Want To Delete Employee : " & emp_grid.Columns(1), vbYesNo)
        If a = 6 Then
        Dim cmd3 As String
        cmd3 = " update employees set emp_project=0 where emp_id="
        cmd3 = cmd3 + emp_id.Text
        db_con.con.Execute (cmd3)
        emp_id.Text = "nil"
        empfetch
        Else
        emp_grid.ReBind
        
        
End If

    Else
        MsgBox ("Choose an Employee")
    End If
Else
    MsgBox ("Choose A Valid Project")
End If
 

End Sub

Private Sub emp_grid_RowColChange(LastRow As Variant, ByVal LastCol As Integer)

If Not Len(cmb_proj.Text) = 0 Then
    
    If Not emp_grid.ApproxCount = 0 Then
    
    emp_grid.MarqueeStyle = dbgHighlightRow
        emp_id.Text = emp_grid.Columns(0)
    Else
        emp_id.Text = "nil"
    End If
Else
    MsgBox ("Choose A Valid Project")
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

