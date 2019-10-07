Attribute VB_Name = "form_center"
Public Sub center(child As Object)
On Error Resume Next

child.Move (Screen.Width - child.Width) / 2, (Screen.Height - child.Height) / 2

End Sub
