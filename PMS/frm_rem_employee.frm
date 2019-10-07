VERSION 5.00
Begin VB.Form frm_rem_employee 
   BorderStyle     =   0  'None
   Caption         =   "Select Employee"
   ClientHeight    =   9750
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7680
   LinkTopic       =   "Form3"
   Picture         =   "frm_rem_employee.frx":0000
   ScaleHeight     =   9750
   ScaleWidth      =   7680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Remove"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1560
      TabIndex        =   4
      Top             =   4080
      Width           =   1335
   End
   Begin VB.ComboBox cmb_name 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   1560
      TabIndex        =   1
      Text            =   "Emp_name"
      Top             =   2880
      Width           =   3735
   End
   Begin VB.ComboBox cmb_id 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   1560
      TabIndex        =   0
      Text            =   "Emp_id"
      Top             =   1560
      Width           =   3855
   End
   Begin VB.Label close 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000009&
      Height          =   375
      Left            =   7320
      TabIndex        =   5
      Top             =   0
      Width           =   495
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Employee Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1440
      TabIndex        =   3
      Top             =   2280
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "Employee ID"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1440
      TabIndex        =   2
      Top             =   1080
      Width           =   1935
   End
End
Attribute VB_Name = "frm_rem_employee"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim emprec As ADODB.Recordset
Dim ename As String
Dim e_id As String


Dim pid As String

Private Sub empfetch()
cmd = "select emp_id,emp_name from employees where emp_project="
cmd = cmd + pid
Set emprec = db_con.con.Execute(cmd)
End Sub

Private Sub close_Click()
Unload frm_rem_employee


frm_manage_employee.empfetch

End Sub

Private Sub cmb_id_Click()
e_id = cmb_id.Text
cmb_name.ListIndex = cmb_id.ListIndex


End Sub

Private Sub cmb_name_Click()
ename = cmb_name.Text
cmb_id.ListIndex = cmb_name.ListIndex

End Sub

Private Sub Command1_Click()
If cmb_id.Text = " " Or cmb_id.Text = "Emp_id" Or cmb_name.Text = " " Or cmb_name.Text = "Emp_name" Then
MsgBox ("Choose valid id")
Else
cmd3 = " update employees set emp_project=0 where emp_id="
cmd3 = cmd3 + e_id
db_con.con.Execute (cmd3)
cmb_id.RemoveItem (cmb_id.ListIndex)
cmb_name.RemoveItem (cmb_name.ListIndex)
End If
cmb_id.Text = "Emp_id"
cmb_name.Text = "Emp_name"
End Sub



Private Sub Form_Load()
pid = frm_manage_employee.txt_pid.Text
empfetch


Do Until emprec.EOF
    cmb_id.AddItem emprec!emp_id
    cmb_name.AddItem emprec!emp_name
    emprec.MoveNext
Loop

End Sub
