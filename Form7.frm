VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form7 
   Caption         =   "Form7"
   ClientHeight    =   6975
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4875
   LinkTopic       =   "Form7"
   ScaleHeight     =   6975
   ScaleWidth      =   4875
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ListView ListView1 
      Height          =   4695
      Left            =   720
      TabIndex        =   2
      Top             =   2160
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   8281
      View            =   2
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
      FlatScrollBar   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1560
      TabIndex        =   1
      Top             =   720
      Width           =   1815
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Results:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   600
      TabIndex        =   3
      Top             =   1800
      Width           =   840
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Subject:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   600
      TabIndex        =   0
      Top             =   720
      Width           =   810
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db As New ADODB.Connection
Dim rs As ADODB.Recordset

Private Sub Combo1_Click()
Set rs = New ADODB.Recordset

ListView1.ListItems.clear

rs.Open "Select * from tablo", db, adOpenDynamic, adLockOptimistic
 
Dim litem As ListItem

While rs.EOF = False

    If rs![Subject] = Combo1.Text Then
        
        Set litem = ListView1.ListItems.Add(, , rs![Title])
                                  
    End If
   rs.MoveNext
    
Wend
rs.Close
End Sub

Private Sub Form_Load()
On Error Resume Next
Form4.WindowState = vbNormal
Set rs = New ADODB.Recordset
    
    db.ConnectionString = "DRIVER={Microsoft Access Driver (*.mdb)};DBQ=" & App.Path & "\kitap.mdb"
    db.Open

rs.Open "Select * from subjct", db, adOpenDynamic, adLockOptimistic

While rs.EOF = False
Combo1.AddItem rs![Subject]
rs.MoveNext
Wend
rs.Close
End Sub

Private Sub Form_Unload(Cancel As Integer)
db.Close
Set db = Nothing
End Sub

Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
ListView1.SortKey = ColumnHeader.Index - 1
End Sub
