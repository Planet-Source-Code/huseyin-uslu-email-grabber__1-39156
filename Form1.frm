VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Email Grabber by Hüseyin Uslu"
   ClientHeight    =   8925
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5940
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8925
   ScaleWidth      =   5940
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      Caption         =   "Grabbed emails"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3735
      Left            =   120
      TabIndex        =   5
      Top             =   4800
      Width           =   5655
      Begin VB.CommandButton Command2 
         Caption         =   "Thanks! Save to disk "
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3360
         TabIndex        =   7
         Top             =   3120
         Width           =   2175
      End
      Begin VB.ListBox List1 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2595
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   5415
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Process..."
      Height          =   1455
      Left            =   120
      TabIndex        =   4
      Top             =   3240
      Width           =   5655
      Begin MSComctlLib.ProgressBar ProgressCurrent 
         Height          =   375
         Left            =   1680
         TabIndex        =   8
         Top             =   360
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   661
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Min             =   1e-4
      End
      Begin MSComctlLib.ProgressBar ProgressOverall 
         Height          =   375
         Left            =   1680
         TabIndex        =   9
         Top             =   840
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   661
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Min             =   1e-4
      End
      Begin VB.Label Label3 
         Caption         =   "Overall progress:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "File progress:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   480
         Width           =   1215
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Grab the email now!! Go go go!"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   2640
      Width           =   5775
   End
   Begin VB.Frame Frame1 
      Caption         =   "HTML files are in this directory"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5775
      Begin VB.DirListBox Dir1 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1440
         Left            =   120
         TabIndex        =   2
         Top             =   840
         Width           =   5535
      End
      Begin VB.DriveListBox Drive1 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   5535
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Please vote me in PSCode !"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   120
      MousePointer    =   2  'Cross
      TabIndex        =   12
      Top             =   8640
      Width           =   5655
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib _
     "shell32.dll" Alias "ShellExecuteA" _
     (ByVal hWnd As Long, ByVal lpOperation _
     As String, ByVal lpFile As String, ByVal _
     lpParameters As String, ByVal lpDirectory _
     As String, ByVal nShowCmd As Long) As Long
Private Sub Command1_Click()
Me.Command2.Enabled = False
Me.ProgressCurrent.Value = 0.1
Me.ProgressOverall.Value = 0.1
Me.List1.Clear
Dim spath As String
spath = Dir1.Path
Command1.Enabled = False

Set fso = CreateObject("Scripting.FileSystemObject")
Set folder = fso.GetFolder(spath)
Set Files = folder.Files

Dim i As Integer
'count total files
For Each file In Files
i = i + 1
Next
Me.ProgressOverall.Tag = 100 / i

'dosyalarý isle
For Each file In Files
DoEvents
If Me.ProgressOverall.Value + Me.ProgressOverall.Tag >= 100 Then
    Me.ProgressOverall.Value = 99.9
Else
    Me.ProgressOverall.Value = Me.ProgressOverall.Value + Me.ProgressOverall.Tag
End If

'dosyayi aç
Me.ProgressCurrent.Value = 0.1
Dim filepath As String

filepath = spath & "\" & file.Name

Dim dotloc As Integer
Dim extension As String
dotloc = 1
Do While (InStr(dotloc + 1, filepath, ".") <> 0)
    dotloc = InStr(dotloc + 1, filepath, ".")
Loop
extension = Mid(filepath, dotloc + 1)
DoEvents

If LCase(extension) = "html" Or LCase(extension) = "htm" Or LCase(extension) = "asp" Or LCase(extension) = "php" Or LCase(extension) = "php3" Or LCase(extension) = "aspx" Then
Me.ProgressCurrent.Value = 10
Set textStreamObject = fso.OpenTextFile(filepath, 1, False, 0)
DoEvents

Dim content As String
Dim loc As Long
loc = 1

content = textStreamObject.ReadAll

'@ bul
If InStr(loc, content, "@") <> 0 Then
loc = InStr(loc, content, "@")
'find the space before @ and after @
Dim l As Long
Dim r As Long
Dim foundleft As Boolean
Dim foundright As Boolean
foundleft = False
foundright = False
l = loc
r = loc

Do While (foundleft = False)
    If Mid(content, l, 1) = " " Then
        foundleft = True
    Else
        l = l - 1
    End If
Loop

Do While (foundright = False)
    If Mid(content, r, 1) = " " Then
        foundright = True
    Else
        r = r + 1
    End If
Loop
Me.ProgressCurrent.Value = 50

'email formats
Dim semail As String
semail = Mid(content, l + 1, r - l)
If Left(semail, 4) = "href" Then
Dim sl As Long
Dim sr As Long
sl = InStr(1, semail, "<")
sr = InStr(1, semail, ">")
semail = Mid(semail, sr + 1, sl - sr - 1)
End If

DoEvents

'search for duplicates
Dim sduplicate As Boolean
sduplicate = False
Dim o As Integer

For o = 0 To Me.List1.ListCount - 1
    If Trim(LCase(semail)) = Trim(LCase(Me.List1.List(o))) Then
        sduplicate = True
    End If
    DoEvents
Next o

Me.ProgressCurrent.Value = 70
If sduplicate = False Then List1.AddItem semail


End If
Set textStreamObject = Nothing


End If

Me.ProgressCurrent.Value = 100
DoEvents

Next

Me.ProgressOverall.Value = 100


Set Files = folder.Files
Set file = Nothing
Set fso = Nothing


Command1.Enabled = True


MsgBox "Finished Grabbing! Found a total of " & Me.List1.ListCount & " email adresses!", vbInformation, "Finished..."
Me.Command2.Enabled = True
End Sub

Private Sub Command2_Click()
Dim sfree As Integer
sfree = FreeFile

Open App.Path & "\emails.txt" For Output As #sfree
    Dim g As Integer
    For g = 0 To Me.List1.ListCount - 1
        Print #1, Me.List1.List(g)
    Next g
Close #sfree
MsgBox "File written as " & App.Path & "\emails.txt successfully!", vbInformation, "File written!"
End Sub

Private Sub Drive1_Change()
On Error GoTo err:
Dir1.Path = Drive1.List(Drive1.ListIndex)

Exit Sub
err:
MsgBox err.Number
MsgBox err.Description
End Sub

Private Sub Form_Load()
Me.ProgressCurrent.Value = 0.1
Me.ProgressOverall.Value = 0.1
End Sub

Private Sub Label1_Click()
ShellExecute 0&, vbNullString, "http://www.pscode.com/vb/scripts/BrowseCategoryOrSearchResults.asp?optSort=Alphabetical&lngWId=1&B1=Quick+Search&blnWorldDropDownUsed=TRUE&txtMaxNumberOfEntriesPerPage=10&blnResetAllVariables=TRUE&txtCriteria=huseyin+or+h%FCseyin", vbNullString, _
      vbNullString, SW_SHOWNORMAL
End Sub
