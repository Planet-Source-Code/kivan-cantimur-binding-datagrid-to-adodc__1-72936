VERSION 5.00
Begin VB.Form Form5 
   Caption         =   "Subject Search"
   ClientHeight    =   1050
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3225
   LinkTopic       =   "Form5"
   ScaleHeight     =   1050
   ScaleWidth      =   3225
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdFindNext 
      Caption         =   "Find Next"
      Height          =   375
      Left            =   1680
      TabIndex        =   1
      Top             =   360
      Width           =   1095
   End
   Begin VB.TextBox txtSubject 
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   1215
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdFindNext_Click()
On Error GoTo Errlab

    txtSubject.Text = strSearc
     rsd.MoveNext
     rsd.Find "Subject Like '" & txtSubject.Text & "'"
        Call display
        
Errlab:
    Exit Sub

End Sub

Private Sub Form_Load()
txtSubject.Text = strSearc
txtSubject.Enabled = False
End Sub
