VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form1 
   Caption         =   " "
   ClientHeight    =   8925
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12660
   LinkTopic       =   "Form1"
   ScaleHeight     =   8925
   ScaleWidth      =   12660
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command6 
      Caption         =   "List by subject"
      Height          =   375
      Left            =   2160
      TabIndex        =   43
      Top             =   7680
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Show Report"
      Height          =   375
      Left            =   10920
      TabIndex        =   42
      Top             =   7680
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Create an Excell Report"
      Height          =   375
      Left            =   8760
      TabIndex        =   41
      Top             =   7680
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Online Search"
      Height          =   375
      Left            =   3480
      TabIndex        =   40
      Top             =   7680
      Width           =   1455
   End
   Begin VB.CommandButton cmdAutBook 
      Caption         =   "Author's works"
      Height          =   375
      Left            =   720
      TabIndex        =   32
      Top             =   7680
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Count"
      Height          =   375
      Left            =   9000
      TabIndex        =   28
      Top             =   2040
      Width           =   975
   End
   Begin VB.CommandButton cmdTitleNext 
      Caption         =   "Find next"
      Height          =   255
      Left            =   2520
      TabIndex        =   27
      Top             =   3120
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   3120
      TabIndex        =   26
      Top             =   2640
      Width           =   1335
   End
   Begin VB.ComboBox Cmb1 
      Height          =   315
      Left            =   6600
      TabIndex        =   25
      Top             =   3000
      Width           =   1215
   End
   Begin VB.CommandButton cmdSubnext 
      Caption         =   "Find next"
      Height          =   255
      Left            =   9360
      TabIndex        =   24
      Top             =   3120
      Width           =   975
   End
   Begin VB.CommandButton Subnext 
      Caption         =   "Search by Subject"
      Height          =   495
      Left            =   7920
      TabIndex        =   23
      Top             =   2880
      Width           =   1335
   End
   Begin VB.ComboBox Cmb 
      Height          =   315
      Left            =   1080
      TabIndex        =   5
      Top             =   1800
      Width           =   1815
   End
   Begin VB.CommandButton cmdautnext 
      Caption         =   "Find next"
      Height          =   255
      Left            =   5400
      TabIndex        =   22
      Top             =   3120
      Width           =   975
   End
   Begin VB.CommandButton cmdSearAut 
      Caption         =   "Search by Author"
      Height          =   375
      Left            =   3840
      TabIndex        =   21
      Top             =   3000
      Width           =   1455
   End
   Begin VB.TextBox txtNo 
      Height          =   285
      Left            =   1080
      TabIndex        =   1
      Top             =   360
      Width           =   735
   End
   Begin VB.TextBox txtDate 
      Height          =   285
      Left            =   1560
      TabIndex        =   7
      Top             =   2520
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog ComDlg 
      Left            =   8280
      Top             =   1080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Open Datebase"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8640
      TabIndex        =   20
      Top             =   480
      Width           =   1335
   End
   Begin VB.CommandButton cmdCreate 
      Caption         =   "Create Database"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6960
      TabIndex        =   19
      Top             =   480
      Width           =   1575
   End
   Begin VB.CommandButton cmdDeleteNumber 
      Caption         =   "Delete from ID"
      Height          =   375
      Left            =   5160
      TabIndex        =   18
      Top             =   1320
      Width           =   1455
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   375
      Left            =   3120
      TabIndex        =   17
      Top             =   840
      Width           =   975
   End
   Begin VB.CommandButton cmdUptBatch 
      Caption         =   "Update batch"
      Height          =   375
      Left            =   5280
      TabIndex        =   16
      Top             =   840
      Width           =   1335
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "Search by Title"
      Height          =   375
      Left            =   1080
      TabIndex        =   15
      Top             =   3000
      Width           =   1335
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "Update"
      Height          =   375
      Left            =   5280
      TabIndex        =   14
      Top             =   360
      Width           =   975
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4200
      TabIndex        =   13
      Top             =   360
      Width           =   975
   End
   Begin VB.CommandButton cmdAddNew 
      Caption         =   "Add new"
      Height          =   375
      Left            =   3120
      TabIndex        =   12
      Top             =   360
      Width           =   975
   End
   Begin VB.CommandButton cmdMoveLast 
      Caption         =   ">|"
      Height          =   255
      Left            =   6480
      TabIndex        =   11
      Top             =   2400
      Width           =   615
   End
   Begin VB.CommandButton cmdMoveFirst 
      Caption         =   "|<"
      Height          =   255
      Left            =   4680
      TabIndex        =   10
      Top             =   2400
      Width           =   615
   End
   Begin VB.CommandButton cmdMovePrevious 
      Caption         =   "<"
      Height          =   255
      Left            =   5280
      TabIndex        =   9
      Top             =   2400
      Width           =   615
   End
   Begin VB.CommandButton cmdMoveNext 
      Caption         =   ">"
      Height          =   255
      Left            =   5880
      TabIndex        =   8
      Top             =   2400
      Width           =   615
   End
   Begin VB.TextBox txtTranslator 
      Height          =   285
      Left            =   1080
      TabIndex        =   4
      Top             =   1440
      Width           =   1815
   End
   Begin VB.TextBox txtAuthor 
      Height          =   285
      Left            =   1080
      TabIndex        =   3
      Top             =   1080
      Width           =   1815
   End
   Begin VB.TextBox txtLoc 
      Height          =   285
      Left            =   1080
      TabIndex        =   6
      Top             =   2160
      Width           =   1815
   End
   Begin VB.TextBox txtTitle 
      Height          =   285
      Left            =   1080
      TabIndex        =   2
      Top             =   720
      Width           =   1815
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   4095
      Left            =   120
      TabIndex        =   0
      Top             =   3480
      Width           =   12135
      _ExtentX        =   21405
      _ExtentY        =   7223
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
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
            LCID            =   1055
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
            LCID            =   1055
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
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "Publishing Date:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   39
      Top             =   2520
      Width           =   1410
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "Publisher:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   38
      Top             =   2160
      Width           =   855
   End
   Begin VB.Label Label8 
      Caption         =   "Subject:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   37
      Top             =   1800
      Width           =   735
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Translator:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   36
      Top             =   1440
      Width           =   930
   End
   Begin VB.Label Label6 
      Caption         =   "Author:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   35
      Top             =   1080
      Width           =   615
   End
   Begin VB.Label Label5 
      Caption         =   "Title:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   34
      Top             =   720
      Width           =   615
   End
   Begin VB.Label Label4 
      Caption         =   "ID:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   33
      Top             =   360
      Width           =   615
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   8400
      TabIndex        =   31
      Top             =   2040
      Width           =   45
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Number of records:"
      Height          =   195
      Left            =   6840
      TabIndex        =   30
      Top             =   2040
      Width           =   1350
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   8400
      TabIndex        =   29
      Top             =   1800
      Width           =   75
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cn As New ADODB.Connection
'Dim rs As New ADODB.Recordset
'Dim rsd As New ADODB.Recordset
Dim rscb As New ADODB.Recordset
Dim rcb As New ADODB.Recordset

'Yazar tablosu icin
Dim rsAuthor As New ADODB.Recordset
'-----
Dim rsSubject As New ADODB.Recordset

Dim rcount As New ADODB.Recordset

Dim rdlt As New ADODB.Recordset
Dim rsupd As New ADODB.Recordset
Dim rsbat As New ADODB.Recordset
'report to excell
Dim rst As New ADODB.Recordset
Dim iCols As Integer
Dim oApp As Excel.Application
Dim oWB As Excel.Workbook
Dim strSearch As String

Private Sub Cmb_Click()
'txtSubject.Text = Cmb.Text
'Cmb.Text = ""
End Sub

Private Sub cmdAddNew_Click()
On Error Resume Next
cmdSave.Enabled = True
    Call clear

    rs.AddNew
    
    rsd.MoveLast
  
If Not rs.EOF Then
   txtNo.Text = Str(rsd(0)) & " " 'Space is appended to avoid error when value is null
End If
    txtNo.Text = txtNo.Text + 1
    txtTitle.SetFocus
    
If txtNo.Text = "" Then
    txtNo.Text = "1"
    txtTitle.SetFocus
End If
End Sub

Private Sub clear()
    txtNo.Text = ""
    txtTitle.Text = ""
    txtAuthor.Text = ""
    txtTranslator.Text = ""
    Cmb.Text = ""
    txtLoc.Text = ""
    txtDate.Text = ""
End Sub

Private Sub cmdAutBook_Click()
Form4.Show
Load Form4
End Sub

Private Sub cmdautnext_Click()
On Error GoTo Errlab
DataGrid1.Scroll 0, 1
       
   Text1.Text = strSearch
    strSearch = strSearch
    rsd.MoveNext
    rsd.Find "Author Like '" & Text1.Text & "'"
       Call display
Errlab:
    Exit Sub
End Sub

Private Sub cmdCreate_Click()
Form3.Show
Load Form3
End Sub

Private Sub cmdDelete_Click()
On Error Resume Next
rsd.Open "Select * from tablo", cn, adOpenDynamic, adLockOptimistic
If MsgBox("Are you sure do you want to delete the record?", vbOKCancel + vbExclamation, "Delete") = vbOK Then
    rsd.Delete
    rsd.MoveNext
        If rsd.EOF Then
            rsd.MoveLast
            MsgBox "Last record..."
        End If
    Call displayrdlt
    rsd.Requery
End If
End Sub

Private Sub displayrdlt()
rdlt(0) = txtNo.Text
rdlt(1) = txtTitle.Text
rdlt(2) = txtAuthor.Text
rdlt(3) = txtTranslator.Text
rdlt(4) = Cmb.Text
rdlt(5) = txtLoc.Text
rdlt(6) = txtDate.Text
End Sub

Private Sub cmdDeleteNumber_Click()
On Error Resume Next
Dim strNo
       
rsd.Open "Select * from tablo", cn, adOpenDynamic, adLockOptimistic

rsd.MoveFirst

strNo = InputBox("Enter the ID of the record you want to delete:", "------")
If MsgBox("Are you sure?", vbOKCancel + vbExclamation, "Delete") = vbOK Then

Do While Not rsd.EOF
    strNo = strNo
    
    rsd.Find "No Like '" & strNo & "'"

rsd.Delete
rsd.MoveNext
        If rsd.EOF Then
            rsd.MoveLast
            MsgBox "Last record."
            Exit Sub
        End If
rdlt(0) = txtNo.Text
rdlt(1) = txtTitle.Text
rdlt(2) = txtAuthor.Text
rdlt(3) = txtTranslator.Text
rdlt(4) = Cmb.Text
rdlt(5) = txtLoc.Text
rdlt(6) = txtDate.Text

Loop
End If
    rsd.Requery
End Sub

Private Sub cmdMoveFirst_Click()
On Error Resume Next
    rsd.MoveFirst
    Call display
End Sub

Private Sub cmdMoveLast_Click()
On Error Resume Next
    rsd.MoveLast
    Call display
End Sub

Private Sub cmdMoveNext_Click()
On Error Resume Next
    rsd.MoveNext
    If rsd.EOF Then
       rsd.MoveLast
       MsgBox "You are at the last record", vbInformation
    Else
       Call display
    End If
End Sub

Private Sub cmdMovePrevious_Click()
On Error Resume Next
    rsd.MovePrevious
    If rsd.BOF Then
       rsd.MoveFirst
       MsgBox "You are at the first record", vbInformation
    Else
       Call display
    End If
End Sub

Private Sub cmdRequery_Click()
On Error Resume Next
    rsd.Requery
End Sub

Private Sub cmdSave_Click()
On Error Resume Next
Dim sSQL, rsSQL As String

sSQL = "Select Author from yazar WHERE Author = '" & Replace(txtAuthor.Text, "'", "''") & "'"
rsSQL = "Select Subject from subjct WHERE Subject = '" & Replace(Cmb.Text, "'", "''") & "'"

rsAuthor.Open sSQL, cn, adOpenDynamic, adLockOptimistic
rsSubject.Open rsSQL, cn, adOpenDynamic, adLockOptimistic

If Cmb.Text = "" Then
    MsgBox "Select a subject or enter a new subject."
    Cmb.SetFocus
        Exit Sub
End If
    
  If rsAuthor.EOF Then ' Author not exists, add it
    rsAuthor.AddNew
    rsAuthor.Fields("Author") = txtAuthor.Text
    rsAuthor.Update
  End If
    
 If rsSubject.EOF Then
    rsSubject.AddNew
    rsSubject.Fields("Subject") = Cmb.Text
    rsSubject.Update
    Cmb1.AddItem Cmb.Text
    Cmb.AddItem Cmb.Text
End If
    
    Call assignsave
    rs.Save
    rsd.Requery
    
  MsgBox "Record has been saved", vbInformation, "----"
  
  rsAuthor.Close
  Set rsAuthor = Nothing
    rsSubject.Close
    Set rsSubject = Nothing
    
cmdSave.Enabled = False
'+

'------
    
     'Do Until Adodc1.Recordset.EOF
     '   If Adodc1.Recordset.Fields("Name").Value = sFile Then
            'MsgBox ("This Picture Exists!"), vbOKOnly, "Error"
            'Exit Sub
    '    End If
        'Adodc1.Recordset.MoveNext
    'Loop
    
   ' "SELECT iKonu FROM konu WHERE iKonu = '" & txtSubject.Text & "'"
   
   ' cn.Execute ("INSERT INTO konu(iKonu)VALUES('" & txtSubject.Text & "')")
   ' MsgBox "Record Saved", vbInformation, "Record is Saved"
End Sub

Private Sub assignsave()
rs(0) = txtNo.Text
rs(1) = txtTitle.Text
rs(2) = txtAuthor.Text
rs(3) = txtTranslator.Text
rs(4) = Cmb.Text
rs(5) = txtLoc.Text
rs(6) = txtDate.Text
End Sub
Private Sub assign()
rsd(0) = txtNo.Text
rsd(1) = txtTitle.Text
rsd(2) = txtAuthor.Text
rsd(3) = txtTranslator.Text
rsd(4) = Cmb.Text
rsd(5) = txtLoc.Text
rsd(6) = txtDate.Text
End Sub

Private Sub cmdSearAut_Click()
On Error GoTo Errlab

    rsd.MoveFirst
    
    strSearch = InputBox("", "Search author")
    strSearch = strSearch & "%"
    
    rsd.Find "Author Like '" & strSearch & "'"
    
    Text1.Text = strSearch
    Text1.Enabled = False
        Call display
Errlab:
    Exit Sub
End Sub

Private Sub cmdSearch_Click()
On Error GoTo Errlab
    
    rsd.MoveFirst
    
    strSearch = InputBox("", "Search title")
    strSearch = strSearch & "%"
    
    rsd.Find "Title Like '" & strSearch & "'"
    
    Text1.Text = strSearch
    Text1.Enabled = False
    
        Call display
Errlab:
    Exit Sub
End Sub

Private Sub cmdSen_Click()
On Error GoTo Errlab
    
    strSearch = InputBox("", "Search author")
    strSearch = strSearch & "%"
    rsd.MoveNext
    rsd.Find "Author Like '" & strSearch & "'"
        Call display
Errlab:
    Exit Sub
End Sub

Private Sub cmdSubnext_Click()
Form5.Show
Load Form5
End Sub

Private Sub cmdTitleNext_Click()
On Error GoTo Errlab
DataGrid1.Scroll 0, 1
       
   Text1.Text = strSearch
    strSearch = strSearch
    rsd.MoveNext
    rsd.Find "Title Like '" & Text1.Text & "'"
       Call display
Errlab:
    Exit Sub
End Sub

Private Sub cmdUpdate_Click()
On Error Resume Next
rsd.Open "Select * from tablo", cn, adOpenDynamic, adLockOptimistic
    rsd.Update
    rsAuthor.Update
    Call assign
    rsd.Requery
End Sub

Private Sub cmdUptBatch_Click()
On Error Resume Next
rsbat.Open "Select * from tablo", cn, adOpenStatic, adLockBatchOptimistic

    Call assignbatch
    rsbat.UpdateBatch
    rsd.Requery
    'Set rsd = Nothing
End Sub

Private Sub assignbatch()
rsbat.AddNew
rsbat(0) = txtNo.Text
rsbat(1) = txtTitle.Text
rsbat(2) = txtAuthor.Text
rsbat(3) = txtTranslator.Text
rsbat(4) = Cmb.Text
rsbat(5) = txtLoc.Text
rsbat(6) = txtDate.Text
End Sub

Private Sub Command1_Click()
Form2.Show
Load Form2
End Sub

Private Sub Command2_Click()
On Error Resume Next
Dim sql As String
'adoConn.ConnectionString = "DRIVER={Microsoft Access Driver (*.mdb)};DBQ=" & MdbFile   'App.Path & "\Examples.mdb"
rs.Close
rsd.Close
cn.Close
ComDlg.DialogTitle = "Open"
ComDlg.ShowOpen
cn.ConnectionString = "DRIVER={Microsoft Access Driver (*.mdb)};DBQ=" & ComDlg.FileName
    
    'cn.ConnectionString = "DRIVER={Microsoft Access Driver (*.mdb)};DBQ=" & App.Path & "\emp.mdb"
    cn.Open
    Debug.Print "Connection Object Created"

    rs.Open "Select * from tablo", cn, adOpenDynamic, adLockOptimistic
    
    sql = "select * from tablo order by No ASC"
    rsd.Open sql, cn, adOpenStatic, adLockOptimistic, adCmdText
    Debug.Print "Recordset Object Created"
    
With DataGrid1
   Set .DataSource = rsd
   .AllowUpdate = False
End With

Call display
End Sub

Private Sub Command3_Click()
Dim i As Integer
Dim start_time As Single

  For i = 0 To 10
        rcount.Open "SELECT * FROM tablo", cn, adOpenStatic, , adCmdText
        Label3.Caption = Format$(rcount.RecordCount)
        rcount.Close
  Next i
End Sub

Private Sub Command4_Click()
On Error Resume Next

rst.Open "SELECT * FROM tablo order by ID ASC", cn, adOpenDynamic, adLockOptimistic
    Set oApp = New Excel.Application
    Set oWB = oApp.Workbooks.Add
    
    For iCols = 0 To rst.Fields.Count - 1
        oWB.Sheets(1).Cells(1, iCols + 1).Value = rst.Fields(iCols).Name
    Next
    
    oWB.Sheets(1).Range(oWB.Sheets(1).Cells(1, 1), oWB.Sheets(1).Cells(1, rst.Fields.Count)).Font.Bold = True
    'List all the oRs data...
    oWB.Sheets(1).Range("A2").CopyFromRecordset rst
     
     ComDlg.Filter = "Excel Files(*.xls)|*.xls"
     ComDlg.ShowSave
     If ComDlg.FileName <> "" Then
        
        oWB.SaveAs ComDlg.FileName
        oWB.Saved = True
    End If
    'you have to close the objects to compelete the process
    oWB.Close
    Set oWB = Nothing
    oApp.Quit
    Set oApp = Nothing
    
    rst.Close
    Set rst = Nothing
    
    LocEx = ComDlg.FileName
End Sub

Private Sub Command5_Click()
Form6.Show
Load Form6
End Sub

Private Sub Command6_Click()
Form7.Show
Load Form7
End Sub

Private Sub DataGrid1_DblClick()
On Error Resume Next
rsd.Open "Select * from tablo", cn, adOpenDynamic, adLockOptimistic
With rsd
   txtNo.Text = .Fields(0)
   txtTitle.Text = .Fields(1)
   txtAuthor.Text = .Fields(2)
   txtTranslator.Text = .Fields(3)
   Cmb.Text = .Fields(4)
   txtLoc.Text = .Fields(5)
   txtDate = .Fields(6)
End With
End Sub

Private Sub Form_Load()
On Error Resume Next
Dim sql As String
'adoConn.ConnectionString = "DRIVER={Microsoft Access Driver (*.mdb)};DBQ=" & MdbFile   'App.Path & "\Examples.mdb"
'cd.ShowOpen
'cn.ConnectionString = "DRIVER={Microsoft Access Driver (*.mdb)};DBQ=" & CD.FileName
    
    cn.ConnectionString = "DRIVER={Microsoft Access Driver (*.mdb)};DBQ=" & App.Path & "\kitap.mdb"
    cn.Open
    Debug.Print "Connection Object Created"

    rs.Open "Select * from tablo", cn, adOpenDynamic, adLockOptimistic
    
    sql = "select * from tablo"
    rsd.Open sql, cn, adOpenStatic, adLockOptimistic, adCmdText
    Debug.Print "Recordset Object Created"
    
With DataGrid1
   Set .DataSource = rsd
   .AllowUpdate = False
End With
    
    rscb.Open "Select * from subjct", cn, adOpenDynamic, adLockOptimistic
    rscb.MoveFirst
    
    Do Until rscb.EOF
        Cmb.AddItem rscb(0)
        rscb.MoveNext
    Loop
    
    rcb.Open "Select * from subjct", cn, adOpenDynamic, adLockOptimistic
    rcb.MoveFirst
    
    Do Until rcb.EOF
        Cmb1.AddItem rcb(0)
        rcb.MoveNext
    Loop

Call display

    'Kaydetme bölümlerinde:
    'rsd.Close
    'Set rs = Nothing
    'Set cn = Nothing
End Sub
Private Sub display()
On Error Resume Next
If Not rs.EOF Then
   txtNo.Text = Str(rsd(0)) & " " 'Space is appended to avoid error when value is null
   txtTitle.Text = rsd(1) & " "
   txtAuthor.Text = rsd(2) & " "
   txtTranslator.Text = rsd(3) & " " 'as salary is numeric it is converted to string
   Cmb.Text = rsd(4) & " "
   txtLoc.Text = rsd(5) & " "
   txtDate.Text = rsd(6) & " "
End If
End Sub

Private Sub Subnext_Click()
On Error GoTo Errlab
    
    rsd.MoveFirst
    
    strSearc = Cmb1.Text
    strSearc = strSearc
    
    rsd.Find "Subject Like '" & strSearc & "'"
        Call display
Errlab:
    Exit Sub
End Sub
