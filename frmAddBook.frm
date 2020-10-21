VERSION 5.00
Begin VB.Form frmAddBook 
   BackColor       =   &H00404040&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Scorpion Address Book"
   ClientHeight    =   4830
   ClientLeft      =   3015
   ClientTop       =   3585
   ClientWidth     =   8670
   Icon            =   "frmAddBook.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4830
   ScaleWidth      =   8670
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Display Panel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4200
      Left            =   120
      TabIndex        =   18
      Top             =   135
      Width           =   8415
      Begin VB.Frame Frame3 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Status Window"
         Height          =   900
         Left            =   5565
         TabIndex        =   30
         Top             =   3240
         Width           =   2805
         Begin VB.Label lblStatus 
            BackColor       =   &H00C0C0C0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   600
            Left            =   75
            TabIndex        =   31
            ToolTipText     =   "Status area"
            Top             =   195
            Width           =   2670
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Control Panel"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2895
         Left            =   5760
         MousePointer    =   99  'Custom
         TabIndex        =   29
         Top             =   360
         Width           =   2535
         Begin VB.CommandButton cmdPrev 
            Caption         =   "<<"
            Height          =   315
            Left            =   1350
            TabIndex        =   12
            ToolTipText     =   "Move to previous record"
            Top             =   945
            Width           =   960
         End
         Begin VB.CommandButton cmdNext 
            Caption         =   ">>"
            Height          =   300
            Left            =   255
            TabIndex        =   11
            ToolTipText     =   "Move to next record"
            Top             =   945
            Width           =   975
         End
         Begin VB.CommandButton cmdemail 
            Caption         =   "&Send E-mail"
            Height          =   300
            Left            =   525
            TabIndex        =   17
            ToolTipText     =   "Send e-mail"
            Top             =   2385
            Width           =   1425
         End
         Begin VB.CommandButton cmdAbout 
            Caption         =   "&About"
            Height          =   300
            Left            =   1350
            TabIndex        =   16
            ToolTipText     =   "About Abbress Book"
            Top             =   1875
            Width           =   975
         End
         Begin VB.CommandButton cmdClose 
            Caption         =   "&Close"
            Height          =   300
            Left            =   270
            TabIndex        =   15
            ToolTipText     =   "Close Address Book"
            Top             =   1875
            Width           =   975
         End
         Begin VB.CommandButton cmdFind 
            Caption         =   "S&earch"
            Height          =   300
            Left            =   270
            TabIndex        =   13
            ToolTipText     =   "Search for record"
            Top             =   1410
            Width           =   975
         End
         Begin VB.CommandButton cmdUpdate 
            Caption         =   "&Save"
            Height          =   300
            Left            =   1350
            TabIndex        =   14
            ToolTipText     =   "Save changes to record"
            Top             =   1410
            Width           =   975
         End
         Begin VB.CommandButton cmdAdd 
            Caption         =   "&Add"
            Height          =   300
            Left            =   240
            TabIndex        =   9
            ToolTipText     =   "Add new record"
            Top             =   480
            Width           =   975
         End
         Begin VB.CommandButton cmdDelete 
            Caption         =   "&Delete"
            Height          =   300
            Left            =   1320
            TabIndex        =   10
            ToolTipText     =   "Delete current record"
            Top             =   480
            Width           =   975
         End
      End
      Begin VB.TextBox txtFields 
         DataField       =   "Lname"
         DataSource      =   "Data1"
         Height          =   285
         Index           =   0
         Left            =   2160
         MaxLength       =   35
         TabIndex        =   0
         Top             =   360
         Width           =   3375
      End
      Begin VB.TextBox txtFields 
         DataField       =   "Fname"
         DataSource      =   "Data1"
         Height          =   285
         Index           =   1
         Left            =   2160
         MaxLength       =   15
         TabIndex        =   1
         Top             =   675
         Width           =   3375
      End
      Begin VB.TextBox txtFields 
         DataField       =   "City"
         DataSource      =   "Data1"
         Height          =   285
         Index           =   2
         Left            =   2160
         MaxLength       =   20
         TabIndex        =   2
         Top             =   990
         Width           =   3375
      End
      Begin VB.TextBox txtFields 
         DataField       =   "Area"
         DataSource      =   "Data1"
         Height          =   285
         Index           =   3
         Left            =   2265
         MaxLength       =   3
         TabIndex        =   3
         Top             =   1320
         Width           =   375
      End
      Begin VB.TextBox txtFields 
         DataField       =   "Phone1"
         DataSource      =   "Data1"
         Height          =   285
         Index           =   4
         Left            =   2880
         MaxLength       =   3
         TabIndex        =   4
         Top             =   1320
         Width           =   375
      End
      Begin VB.TextBox txtFields 
         DataField       =   "Phone2"
         DataSource      =   "Data1"
         Height          =   285
         Index           =   5
         Left            =   3360
         MaxLength       =   4
         TabIndex        =   5
         Top             =   1320
         Width           =   495
      End
      Begin VB.TextBox txtFields 
         DataField       =   "Email"
         DataSource      =   "Data1"
         Height          =   285
         Index           =   6
         Left            =   2160
         MaxLength       =   90
         TabIndex        =   6
         Top             =   1680
         Width           =   3375
      End
      Begin VB.TextBox txtFields 
         DataField       =   "ICQ"
         DataSource      =   "Data1"
         Height          =   285
         Index           =   7
         Left            =   2160
         MaxLength       =   9
         TabIndex        =   7
         Top             =   1995
         Width           =   975
      End
      Begin VB.TextBox txtFields 
         DataField       =   "Memo"
         DataSource      =   "Data1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Index           =   8
         Left            =   1470
         MaxLength       =   250
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   8
         Top             =   2310
         Width           =   4065
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   3240
         TabIndex        =   21
         TabStop         =   0   'False
         Text            =   "-"
         Top             =   1320
         Width           =   150
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   2160
         TabIndex        =   20
         TabStop         =   0   'False
         Text            =   "("
         Top             =   1320
         Width           =   150
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   2
         Left            =   2640
         TabIndex        =   19
         TabStop         =   0   'False
         Text            =   ")"
         Top             =   1320
         Width           =   150
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Last Name:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   28
         Top             =   375
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "First Name:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   27
         Top             =   690
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "City:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   26
         Top             =   1020
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Phone Number:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   25
         Top             =   1335
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "E-mail:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   6
         Left            =   240
         TabIndex        =   24
         Top             =   1695
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "ICQ Number:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   7
         Left            =   240
         TabIndex        =   23
         Top             =   2010
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Comments:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   8
         Left            =   255
         TabIndex        =   22
         Top             =   2340
         Width           =   1815
      End
   End
   Begin VB.Data Data1 
      Align           =   2  'Align Bottom
      Connect         =   "Access"
      DatabaseName    =   "AB.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "idxAB"
      Top             =   4485
      Visible         =   0   'False
      Width           =   8670
   End
   Begin VB.Label lblDate 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "MMMM d, yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   3
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   315
      Left            =   7560
      TabIndex        =   33
      ToolTipText     =   "Today's date"
      Top             =   4470
      Width           =   975
   End
   Begin VB.Label lblBar 
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   315
      Left            =   105
      TabIndex        =   32
      ToolTipText     =   "Records bar"
      Top             =   4470
      Width           =   8430
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuReport 
         Caption         =   "&View Report"
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuData 
      Caption         =   "&Data"
      Begin VB.Menu mnuPrev 
         Caption         =   "&Previous"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuNext 
         Caption         =   "&Next"
         Shortcut        =   ^N
      End
      Begin VB.Menu sep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "&Delete"
         Shortcut        =   ^D
      End
      Begin VB.Menu mnuSave 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu sep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuemail 
         Caption         =   "e-&mail"
         Shortcut        =   ^M
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      NegotiatePosition=   3  'Right
      Begin VB.Menu mnuReadMe 
         Caption         =   "&Read Me"
         Shortcut        =   ^R
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
         Shortcut        =   ^A
      End
   End
End
Attribute VB_Name = "frmAddBook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim TR As Long
Dim ws As Workspace
Dim db As Database
Dim rs As Recordset
Dim datDate As Date



Private Sub Form_Load()
datDate = CDate(Format(Now(), "MMMM,D,YYYY"))
lblDate = datDate
lblStatus = "Program started successfully"
Provider = "Microsoft.Jet.OLEDB.4.0"
Set ws = DBEngine.Workspaces(0)
Set db = ws.OpenDatabase(App.Path & "\AB.mdb")
    With Data1
        .DatabaseName = App.Path & "\AB.mdb"
        .RecordSource = "idxAB"
        .Refresh
    End With
    'Check to see if recordcount is 0
On Error Resume Next
Data1.Recordset.MoveFirst
TR = Data1.Recordset.RecordCount
Screen.MousePointer = vbDefault
    If TR > 0 Then
        txtFields(0).Enabled = True
        txtFields(1).Enabled = True
        txtFields(2).Enabled = True
        txtFields(3).Enabled = True
        txtFields(4).Enabled = True
        txtFields(5).Enabled = True
        txtFields(6).Enabled = True
        txtFields(7).Enabled = True
        txtFields(8).Enabled = True
        cmdDelete.Enabled = True
        cmdUpdate.Enabled = True
        cmdNext.Enabled = True
        cmdPrev.Enabled = True
        cmdFind.Enabled = True
        cmdemail.Enabled = True
        mnuReport.Enabled = True
        mnuData.Enabled = True
'       Data1.Recordset.MoveFirst
        lblStatus = "Program started successfully"
        lblBar.Caption = "Record: " & (Data1.Recordset.AbsolutePosition + 1 & " of " & (Data1.Recordset.RecordCount))
    End If
    
   'Check if copy of program is already running
   If App.PrevInstance Then
      MsgBox "Address Book is already running in memory", vbOKOnly, "Address Book Running"
      ActivatePrevInstance
   End If
   Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
End Sub

Private Sub cmdAbout_Click()
'Open About form
frmAbout.Show
End Sub

Private Sub cmdAdd_Click()
'Add record
On Error GoTo ErrHandle
If TR = 0 Then
    TR = TR + 1
    txtFields(0).Enabled = True
    txtFields(1).Enabled = True
    txtFields(2).Enabled = True
    txtFields(3).Enabled = True
    txtFields(4).Enabled = True
    txtFields(5).Enabled = True
    txtFields(6).Enabled = True
    txtFields(7).Enabled = True
    txtFields(8).Enabled = True
    cmdAdd.Enabled = True
    cmdDelete.Enabled = True
    cmdNext.Enabled = True
    cmdPrev.Enabled = True
    cmdFind.Enabled = True
    cmdUpdate.Enabled = True
    cmdemail.Enabled = True
    mnuReport.Enabled = True
    mnuData.Enabled = True
End If
    txtFields(0).SetFocus
    Data1.Recordset.AddNew
    lblStatus = "Adding new record"

End_of_Proc:
Exit Sub

ErrHandle:
lblStatus = "Error number " & Err.Number & " encountered."
Select Case Err.Number
    Case 3426
        MsgBox ("Each record requires a Last & First name !"), vbOKOnly
        If txtFields(0) = "" Then
            txtFields(0).SetFocus
        Else
            txtFields(1).SetFocus
        End If
        Screen.MousePointer = vbDefault
        Resume Next
    Case Else
        MsgBox "Unknown error has been encountered Saving record!" _
        & Space(1) & "Note the Error number" & Space(1) & Err.Number, vbOKOnly
        Screen.MousePointer = vbDefault
        Resume Next
End Select
Err.Number = 0
End Sub

Private Sub cmdDelete_Click()
'Delete record
  txtFields(0).SetFocus
  TR = Data1.Recordset.RecordCount
  Data1.Recordset.Delete
  lblStatus = "Deleted Record" & Space(1) & txtFields(1) & Space(1) & txtFields(0)
  Data1.Refresh

If TR = 1 And txtFields(0).Text = "" Then
    TR = 0
    Data1.Refresh
    MsgBox "Last record removed from database!", 48
Else
    TR = TR + 1
    Data1.Refresh
End If
End Sub

Private Sub cmdemail_Click()
'Send e-mail to user listed
lblStatus = "Sending e-mail"
SendTo = txtFields(6).Text

    If SendTo = "" Then
        MsgBox "There is no email address entered!", 48, "Error sending e-mail"
        lblStatus = lblStatus & " failed."
    Else
        SendTo = "mailto:" & SendTo
        ShellExecute hwnd, "open", SendTo, vbNullString, vbNullString, SW_SHOWDEFAULT
    End If
End Sub

Private Sub cmdFind_Click()
'Search for first record matching users request
sstr = InputBox("Enter Last Name to Search")
If sstr = "" Then
    Exit Sub
Else
lblStatus = "Search Results for " & sstr
Data1.Recordset.FindFirst "Lname='" & sstr & "'"
        If Data1.Recordset.NoMatch Then
           MsgBox UCase(sstr) & " was not found in the database, check your spelling!", 48, "Search failed"
           lblStatus = lblStatus & " failed."
        End If
End If
End Sub

Private Sub cmdNext_Click()
'Move to next record if not EOF
If Data1.Recordset.BOF = True Or Data1.Recordset.EOF = True Then
    If Data1.Recordset.EOF = True Then
        MsgBox "End of file reached", 48, "Record Warning"
    End If
Else
    Data1.Recordset.MoveNext
    If txtFields(0).Text = "" Then
            MsgBox "End of file reached", 48, "Record Warning"
            Data1.Recordset.MoveLast
    End If
End If
End Sub

Private Sub cmdPrev_Click()
'Move to previous record if not at BOF
If Data1.Recordset.BOF = True Or Data1.Recordset.EOF = True Then
    If Data1.Recordset.BOF = True Then
        MsgBox "Beginning of file reached", 48, "Record Warning"
    End If
Else
    Data1.Recordset.MovePrevious
    If txtFields(0).Text = "" Then
            MsgBox "Beginning of file reached", 48, "Record Warning"
            Data1.Recordset.MoveFirst
    End If
End If
End Sub

Private Sub cmdUpdate_Click()
'Save changes to database & check for errors
On Error GoTo ErrHandle
  cmdUpdate.SetFocus
  Data1.UpdateRecord
  Data1.Recordset.Bookmark = Data1.Recordset.LastModified
  lblStatus = "Saved record" & Space(1) & txtFields(1) & Space(1) & txtFields(0)

End_of_Proc:
  Exit Sub

ErrHandle:
lblStatus = "Error number " & Err.Number & " encountered."
Select Case Err.Number
    Case 3058
        MsgBox ("Each record requires a Last & First name !"), vbOKOnly
        Resume End_of_Proc
    Case 524
        If txtFields(0) = "" Then
            MsgBox ("Last name must be filled in!"), vbOKOnly
            txtFields(0).SetFocus
        Else
            MsgBox ("First name must be filled in!"), vbOKOnly
            txtFields(1).SetFocus
        End If
    Resume End_of_Proc:
    Case 0
        Resume Next
    Case Else
        MsgBox "An error has been encountered Saving record!" _
        & Space(1) & "Note the Error number" & Space(1) & Err.Number, vbOKOnly
        
    Resume End_of_Proc:
End Select
End Sub

Private Sub cmdClose_Click()
'Close program
  Unload Me
End Sub

Private Sub Data1_Reposition()
'Update lblBar with records info
 On Error Resume Next
Screen.MousePointer = vbDefault
If TR = 0 Then
    lblStatus.Caption = "Click Add to Start"
    txtFields(0).Enabled = False
    txtFields(1).Enabled = False
    txtFields(2).Enabled = False
    txtFields(3).Enabled = False
    txtFields(4).Enabled = False
    txtFields(5).Enabled = False
    txtFields(6).Enabled = False
    txtFields(7).Enabled = False
    txtFields(8).Enabled = False
    cmdDelete.Enabled = False
    cmdUpdate.Enabled = False
    cmdNext.Enabled = False
    cmdPrev.Enabled = False
    cmdFind.Enabled = False
    cmdemail.Enabled = False
    mnuReport.Enabled = False
    mnuData.Enabled = False
    lblBar.Caption = "Database is empty"
Else
    lblBar.Caption = "Record: " & (Data1.Recordset.AbsolutePosition + 1 & " of " & (Data1.Recordset.RecordCount))
End If
End Sub

Private Sub Data1_Validate(Action As Integer, Save As Integer)
'Check for what action was taken
Select Case Action
    Case vbdataActionMaximixe
    Case vbDataActionMoveFirst
    Case vbDataActionMovePrevious
    Case vbDataActionMoveNext
    Case vbDataActionMoveLast
    Case vbDataActionAddNew
    Case vbDataActionUpdate
    Case vbDataActionDelete
    Case vbDataActionFind
    Case vbDataActionBookmark
    Case vbDataActionClose

  End Select
  Screen.MousePointer = vbHourglass
End Sub

Private Sub mnuAbout_Click()
frmAbout.Show
End Sub

Private Sub mnuDelete_Click()
'Delete record
  txtFields(0).SetFocus
  TR = Data1.Recordset.RecordCount
  Data1.Recordset.Delete
  lblStatus = "Deleted Record" & Space(1) & txtFields(1) & Space(1) & txtFields(0)
  Data1.Refresh

If TR = 1 And txtFields(0) = "" Then
    TR = 0
    Data1.Refresh
    MsgBox "Last record removed from database!", 48
Else
    TR = TR + 1
    Data1.Refresh
End If
End Sub

Private Sub mnuemail_Click()
'Send e-mail to user listed
lblStatus = "Sending e-mail"
SendTo = txtFields(6).Text

    If SendTo = "" Then
        MsgBox "There is no email address entered!", 48, "Error sending e-mail"
        lblStatus = lblStatus & " failed."
    Else
        SendTo = "mailto:" & SendTo
        ShellExecute hwnd, "open", SendTo, vbNullString, vbNullString, SW_SHOWDEFAULT
    End If
End Sub

Private Sub mnuExit_Click()
Unload Me
End Sub


Private Sub mnuNext_Click()
'Move to next record if not EOF
If Data1.Recordset.BOF = True Or Data1.Recordset.EOF = True Then
    If Data1.Recordset.EOF = True Then
        MsgBox "End of file reached", 48, "Record Warning"
    End If
Else
    Data1.Recordset.MoveNext
    If txtFields(0).Text = "" Then
            MsgBox "End of file reached", 48, "Record Warning"
            Data1.Recordset.MoveLast
    End If
End If
End Sub

Private Sub mnuPrev_Click()
'Move to previous record if not at BOF
If Data1.Recordset.BOF = True Or Data1.Recordset.EOF = True Then
    If Data1.Recordset.BOF = True Then
        MsgBox "Beginning of file reached", 48, "Record Warning"
    End If
Else
    Data1.Recordset.MovePrevious
    If txtFields(0).Text = "" Then
            MsgBox "Beginning of file reached", 48, "Record Warning"
            Data1.Recordset.MoveFirst
    End If
End If
End Sub

Private Sub mnuReadMe_Click()
ReadMe.Show
End Sub

Private Sub mnuReport_Click()
On Error GoTo ErrRpt
DataRpt.Show


End_of_Proc:
Exit Sub

ErrRpt:
Select Case Err.Number
    Case 713
        MsgBox "Missing required file Msdbrptr.dll to run the report feature", 16, "Vew Report  Critical Error"
        Resume Next
    Case Else
        MsgBox "An unknown error has halted the View Report" & Err.Number, 16, "View Report Critical Error"
        Resume Next
End Select
End Sub

Private Sub mnuSave_Click()
'Save changes to database & check for errors
On Error GoTo ErrHandle
  cmdUpdate.SetFocus
  Data1.UpdateRecord
  Data1.Recordset.Bookmark = Data1.Recordset.LastModified
  lblStatus = "Saved record" & Space(1) & txtFields(1) & Space(1) & txtFields(0)

End_of_Proc:
  Exit Sub

ErrHandle:
lblStatus = "Error number " & Err.Number & " encountered."
Select Case Err.Number
    Case 3058
        MsgBox ("Each record requires a Last & First name !"), vbOKOnly
        Resume End_of_Proc
    Case 524
        If txtFields(0) = "" Then
            MsgBox ("Last name must be filled in!"), vbOKOnly
            txtFields(0).SetFocus
        Else
            MsgBox ("First name must be filled in!"), vbOKOnly
            txtFields(1).SetFocus
        End If
    Resume End_of_Proc:
    Case 0
        Resume Next
    Case Else
        MsgBox "An error has been encountered Saving record!" _
        & Space(1) & "Note the Error number" & Space(1) & Err.Number, vbOKOnly
        
    Resume End_of_Proc:
End Select
End Sub
