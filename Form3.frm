VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "DataBase"
   ClientHeight    =   6930
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4530
   LinkTopic       =   "Form3"
   ScaleHeight     =   6930
   ScaleWidth      =   4530
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      Height          =   1335
      Left            =   240
      TabIndex        =   19
      Top             =   5280
      Width           =   2535
      Begin VB.TextBox txtfld 
         Height          =   285
         Index           =   10
         Left            =   120
         TabIndex        =   21
         Text            =   "Author"
         Top             =   720
         Width           =   975
      End
      Begin VB.TextBox txtAuttbl 
         Height          =   285
         Left            =   600
         TabIndex        =   20
         Text            =   "author"
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.CommandButton cmdTableCreate 
      Caption         =   "Create tables"
      Height          =   375
      Left            =   3000
      TabIndex        =   18
      Top             =   5160
      Width           =   1335
   End
   Begin VB.Frame Frame2 
      Height          =   1455
      Left            =   240
      TabIndex        =   15
      Top             =   3720
      Width           =   2535
      Begin VB.TextBox txtfld 
         Enabled         =   0   'False
         Height          =   285
         Index           =   8
         Left            =   120
         TabIndex        =   17
         Text            =   "Subject"
         Top             =   720
         Width           =   855
      End
      Begin VB.TextBox txtTbl 
         Enabled         =   0   'False
         Height          =   285
         Left            =   600
         TabIndex        =   16
         Text            =   "subjct"
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.TextBox txtDB 
      Height          =   285
      Left            =   1440
      TabIndex        =   12
      Top             =   480
      Width           =   1455
   End
   Begin VB.TextBox txtfld 
      Height          =   285
      Index           =   4
      Left            =   360
      TabIndex        =   9
      Top             =   2520
      Width           =   1215
   End
   Begin VB.TextBox txtfld 
      Height          =   285
      Index           =   5
      Left            =   1680
      TabIndex        =   8
      Top             =   2520
      Width           =   1215
   End
   Begin VB.TextBox txtfld 
      Height          =   285
      Index           =   6
      Left            =   360
      TabIndex        =   7
      Top             =   2880
      Width           =   1215
   End
   Begin VB.TextBox txtfld 
      Enabled         =   0   'False
      Height          =   285
      Index           =   0
      Left            =   360
      TabIndex        =   6
      Text            =   "ID"
      Top             =   1800
      Width           =   1215
   End
   Begin VB.TextBox txtfld 
      Enabled         =   0   'False
      Height          =   285
      Index           =   2
      Left            =   360
      TabIndex        =   5
      Text            =   "Author"
      Top             =   2160
      Width           =   1215
   End
   Begin VB.TextBox txtfld 
      Enabled         =   0   'False
      Height          =   285
      Index           =   1
      Left            =   1680
      TabIndex        =   4
      Text            =   "Title"
      Top             =   1800
      Width           =   1215
   End
   Begin VB.TextBox txtfld 
      Enabled         =   0   'False
      Height          =   285
      Index           =   3
      Left            =   1680
      TabIndex        =   3
      Text            =   "Subject"
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   2175
      Left            =   240
      TabIndex        =   2
      Top             =   1200
      Width           =   3735
      Begin VB.TextBox txtTable 
         Enabled         =   0   'False
         Height          =   285
         Left            =   600
         TabIndex        =   13
         Text            =   "tablo"
         Top             =   240
         Width           =   1695
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Add"
         Height          =   375
         Left            =   2760
         TabIndex        =   11
         Top             =   1560
         Width           =   855
      End
      Begin VB.TextBox txtfld 
         Height          =   285
         Index           =   7
         Left            =   1440
         TabIndex        =   10
         Top             =   1680
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdCreate 
      Caption         =   "Create"
      Height          =   375
      Left            =   3360
      TabIndex        =   1
      Top             =   480
      Width           =   855
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Database name:"
      Height          =   195
      Left            =   240
      TabIndex        =   14
      Top             =   480
      Width           =   1170
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   ".mdb"
      Height          =   195
      Left            =   2880
      TabIndex        =   0
      Top             =   600
      Width           =   345
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim X As Integer
Dim catDB As ADOX.Catalog
Dim tblNew, tblNw, tblAuthor As ADOX.Table
Dim CnnString As String

Private Sub cmdAdd_Click()
On Error Resume Next

Set tblNew = New ADOX.Table
tblNew.Name = txtTable.Text

    X = 0
    While X <= 7
        
            tblNew.Columns.Append txtfld.Item(X).Text, adVarWChar
        
        X = X + 1
    Wend

    catDB.Tables.Append tblNew
    
MsgBox "Table " & txtTable.Text & " was successfully added into " & txtDB.Text & ".mdb", vbOKOnly, "ADO Extension"
End Sub

Private Sub cmdCreate_Click()
    CnnString = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" _
                & App.Path & "\" & txtDB.Text & ".mdb"
    
    Set catDB = New ADOX.Catalog
    catDB.Create CnnString
    catDB.ActiveConnection = CnnString
End Sub

Private Sub cmdTableCreate_Click()
On Error Resume Next

Set tblNew = New ADOX.Table
tblNew.Name = txtTable.Text

Set tblNw = New ADOX.Table
tblNw.Name = txtTbl.Text

Set tblAuthor = New ADOX.Table
tblAuthor.Name = txtAuttbl.Text

    X = 0
    While X <= 7
        
            tblNew.Columns.Append txtfld.Item(X).Text, adVarWChar
        
        X = X + 1
    Wend
    
        tblNw.Columns.Append txtfld(8).Text, adVarWChar
        
        tblAuthor.Columns.Append txtfld(10).Text, adVarWChar
        
        
    catDB.Tables.Append tblNew
    catDB.Tables.Append tblNw
    catDB.Tables.Append tblAuthor

MsgBox "Table " & txtTable.Text & " was successfully added into " & txtDB.Text & ".mdb", vbOKOnly, "ADO Extension"
End Sub

