Attribute VB_Name = "Module1"
Public rs As New ADODB.Recordset

Public rsd As New ADODB.Recordset
Public LocEx As String

Public strSearc As String

Public Sub display()
If Not rs.EOF Then
 txtNo.Text = Str(rsd(0)) & " " 'Space is appended to avoid error when value is null
 txtTitle.Text = rsd(1) & " "
 txtAuthor.Text = rsd(2) & " "
 txtTranslator.Text = rsd(3) & " " 'as salary is numeric it is converted to string
 Cmb.Text = rsd(4) & " "
 txtLoc.Text = rsd(5) & " "
 txtDate.Text = Str(rsd(6)) & " "
End If
End Sub


