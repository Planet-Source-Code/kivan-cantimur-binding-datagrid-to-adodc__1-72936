VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "IEFRAME.dll"
Begin VB.Form Form2 
   Caption         =   "Online Search"
   ClientHeight    =   9660
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11820
   LinkTopic       =   "Form2"
   ScaleHeight     =   9660
   ScaleWidth      =   11820
   StartUpPosition =   3  'Windows Default
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   8655
      Left            =   240
      TabIndex        =   3
      Top             =   840
      Width           =   11415
      ExtentX         =   20135
      ExtentY         =   15266
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Search"
      Height          =   375
      Left            =   7680
      TabIndex        =   2
      Top             =   360
      Width           =   855
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   5760
      TabIndex        =   1
      Top             =   360
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1080
      TabIndex        =   0
      Top             =   360
      Width           =   4455
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim longestName As Integer
Dim isAuto As Boolean

Private Sub Combo1_GotFocus()
    Combo1.Text = ""
End Sub

Sub proccessKeyUp(KeyCode As Integer, Combo1 As ComboBox)

    Dim currLen As Integer
    Dim currText As String
    Dim X As Integer

    If Len(Combo1.Text) > longestName Then Exit Sub
    
    If KeyCode = vbKeySpace Then GoTo Allow
    If KeyCode > 47 Or KeyCode < 1 Then

Allow:
        currLen = Len(Combo1.Text)
        For X = 0 To Combo1.ListCount - 1
            If Left(LCase(Combo1.List(X)), currLen) = LCase(Combo1.Text) Then
                Combo1.ListIndex = X
                Combo1.SelStart = currLen
                Combo1.SelLength = Len(Combo1.Text) - currLen
                Exit For
            End If
        Next X
    End If
End Sub


Private Sub Combo1_KeyUp(KeyCode As Integer, Shift As Integer)
If isAuto Then proccessKeyUp KeyCode, Combo1
End Sub

Private Sub Command1_Click()
'This is where we get the search criteria
'then we replace spaces with + signs
Let thesearch = Text1.Text
Dim i As Integer
Let i = 1
While i <= Len(thesearch)
 If Mid(thesearch, i, 1) = " " Then
  Mid(thesearch, i, 1) = "+"
 End If
 i = i + 1
Wend
'Send Information to Webbrowsers and load criteria

If Combo1.Text = "Amazon" Then
    WebBrowser1.Navigate "http://www.amazon.com/s/ref=nb_sb_noss?url=search-alias%3Daps&field-keywords=" & thesearch
End If
If Combo1.Text = "Powell's Books" Then
    WebBrowser1.Navigate "http://www.powells.com/s?header=Search+Form&kw=" & thesearch
End If
If Combo1.Text = "Google Books" Then
    WebBrowser1.Navigate "http://books.google.com.tr/books?q=" & thesearch
End If
If Combo1.Text = "Biblio.com" Then
    WebBrowser1.Navigate "http://www.biblio.com/search.php?stage=1&keyisbn=" & thesearch
End If
If Combo1.Text = "Books A Million" Then
    WebBrowser1.Navigate "http://www.booksamillion.com/search?id=4691071814830&query=" & thesearch & "&where=All"
End If
End Sub

Private Sub Form_Load()
Combo1.AddItem "Amazon"
Combo1.AddItem "Powell's Books"
Combo1.AddItem "Google Books"
Combo1.AddItem "Biblio.com"
Combo1.AddItem "Books A Million"
        longestName = 1
        isAuto = True
End Sub

