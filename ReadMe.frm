VERSION 5.00
Begin VB.Form ReadMe 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Read Me"
   ClientHeight    =   5565
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   7140
   Icon            =   "ReadMe.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5565
   ScaleWidth      =   7140
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox hlpText 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5055
      Left            =   105
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   45
      Width           =   6975
   End
   Begin VB.CommandButton hlpClose 
      Caption         =   "&Close"
      Height          =   375
      Left            =   90
      TabIndex        =   0
      Top             =   5175
      Width           =   1215
   End
End
Attribute VB_Name = "ReadMe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This Function is used to load a text file into a text box
Function GetTextFromFile(txtFile, txtopen As TextBox)
    Dim sfile As String
    Dim nfile As Integer
    On Error Resume Next
    
    nfile = FreeFile
    sfile = txtFile
    Open sfile For Input As nfile
    txtopen = Input(LOF(nfile), nfile)
    Close nfile
End Function

Private Sub Form_Load()
  hlpText.Locked = False
 'Load Readme.txt into the text box
  Call GetTextFromFile(App.Path & "\README.txt", hlpText)
  hlpText.Locked = True
  hlpText.Enabled = True
  Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
End Sub

Private Sub hlpClose_Click()
  Unload Me
End Sub

