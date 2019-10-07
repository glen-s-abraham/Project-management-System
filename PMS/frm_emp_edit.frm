VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frm_emp_edit 
   Caption         =   "Edit Employee Details"
   ClientHeight    =   6630
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11745
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   Picture         =   "frm_emp_edit.frx":0000
   ScaleHeight     =   6630
   ScaleWidth      =   11745
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame 
      BackColor       =   &H00808000&
      BorderStyle     =   0  'None
      Height          =   6615
      Left            =   1800
      TabIndex        =   1
      Top             =   2280
      Width           =   12735
      Begin VB.TextBox search 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   3360
         TabIndex        =   6
         Top             =   720
         Width           =   4215
      End
      Begin VB.CommandButton cmd_remove 
         Appearance      =   0  'Flat
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
         Left            =   3960
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   5760
         Width           =   2415
      End
      Begin VB.CommandButton cmd_add 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         Caption         =   "Update Details"
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
         TabIndex        =   4
         Top             =   5760
         Width           =   2295
      End
      Begin VB.TextBox txt_eid 
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
         Left            =   8400
         TabIndex        =   3
         Text            =   "nil"
         Top             =   720
         Visible         =   0   'False
         Width           =   3375
      End
      Begin MSDataGridLib.DataGrid emp_grid 
         Bindings        =   "frm_emp_edit.frx":24A41
         Height          =   3615
         Left            =   1080
         TabIndex        =   2
         Top             =   1440
         Width           =   10695
         _ExtentX        =   18865
         _ExtentY        =   6376
         _Version        =   393216
         AllowUpdate     =   0   'False
         AllowArrows     =   -1  'True
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
         Left            =   1080
         TabIndex        =   7
         Top             =   720
         Width           =   2415
      End
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Employee Settings"
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
Attribute VB_Name = "frm_emp_edit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim erec As ADODB.Recordset

Private Sub close_Click()
Unload Me

End Sub

Private Sub cmd_add_Click()
If txt_eid.Text = "nil" Or txt_eid.Text = " " Or Not IsNumeric(txt_eid.Text) Then
    MsgBox ("select a employee")
Else
      frm_emp_update.Show
End If

End Sub

Private Sub cmd_remove_Click()
If txt_eid.Text = "nil" Or txt_eid.Text = " " Or Not IsNumeric(txt_eid.Text) Then
    MsgBox ("select a employee")
Else

        
        a = MsgBox("Are You Sure You Want To Delete employee : " & emp_grid.Columns(2), vbYesNo)
        If a = 6 Then
        Dim cmd3 As String
        cmd = "delete from employees where emp_id="
        cmd = cmd + txt_eid.Text
        db_con.con.Execute (cmd)
        txt_eid.Text = "nil"
        selectall
        Else
        emp_grid.ReBind
        End If
        txt_eid.Text = "nil"
        
        

End If

End Sub

Private Sub emp_grid_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
If Not emp_grid.ApproxCount = 0 Then
        emp_grid.MarqueeStyle = dbgHighlightRow

        txt_eid.Text = emp_grid.Columns(0)
    Else
        txt_eid.Text = "nil"
    End If



End Sub

Private Sub Form_Load()
center Frame
selectall

End Sub
Public Sub selectall()

cmd = "select emp_id,emp_adhaar,emp_name,emp_mobile,emp_mail,emp_dep from employees"
Set erec = db_con.con.Execute(cmd)
Set emp_grid.DataSource = erec
End Sub




Private Sub search_Change()
If Len(search) > 0 Then
cmd2 = "select emp_id,emp_adhaar,emp_name,emp_mobile,emp_mail,emp_dep from employees where emp_id like '"
cmd2 = cmd2 + search.Text + "%'"
cmd2 = cmd2 + " or emp_name like '"
cmd2 = cmd2 + search.Text + "%'"
cmd2 = cmd2 + " or emp_adhaar like '"
cmd2 = cmd2 + search.Text + "%'"
cmd2 = cmd2 + " or emp_dep like '"
cmd2 = cmd2 + search.Text + "%'"



Set erec = db_con.con.Execute(cmd2)
Set emp_grid.DataSource = erec
Else
selectall
End If
End Sub
