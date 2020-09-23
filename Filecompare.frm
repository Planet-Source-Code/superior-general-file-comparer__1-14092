VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Filecompare 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CuckSoft© File Comparer"
   ClientHeight    =   4485
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   4680
   Icon            =   "Filecompare.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4485
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton of 
      Caption         =   "Open File 4"
      Height          =   375
      Index           =   3
      Left            =   2400
      TabIndex        =   5
      Top             =   600
      Width           =   2175
   End
   Begin VB.CommandButton of 
      Caption         =   "Open File 3"
      Height          =   375
      Index           =   2
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   2055
   End
   Begin VB.ListBox dumpbox 
      Height          =   2595
      ItemData        =   "Filecompare.frx":00D2
      Left            =   120
      List            =   "Filecompare.frx":00D4
      TabIndex        =   3
      Top             =   1560
      Width           =   4455
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   2160
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      Filter          =   "All Files|*.*"
   End
   Begin VB.CommandButton compare 
      Caption         =   "Compare Files"
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   4455
   End
   Begin VB.CommandButton of 
      Caption         =   "Open File 2"
      Height          =   375
      Index           =   1
      Left            =   2400
      TabIndex        =   1
      Top             =   120
      Width           =   2175
   End
   Begin VB.CommandButton of 
      Caption         =   "Open File 1"
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label Label1 
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   4200
      Width           =   4455
   End
   Begin VB.Menu mnumode 
      Caption         =   "Modes"
      Begin VB.Menu mnudiff 
         Caption         =   "Difference Finder"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnusim 
         Caption         =   "Similarity Finder"
      End
      Begin VB.Menu mnubreak1 
         Caption         =   "-"
      End
      Begin VB.Menu mnufil 
         Caption         =   "2 Files"
         Checked         =   -1  'True
         Index           =   0
      End
      Begin VB.Menu mnufil 
         Caption         =   "3 Files"
         Index           =   1
      End
      Begin VB.Menu mnufil 
         Caption         =   "4 Files"
         Index           =   2
      End
      Begin VB.Menu mnubreak2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuzero 
         Caption         =   "Ignore zeros in similarities"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuhex 
         Caption         =   "Display Hexidecimal Values"
         Checked         =   -1  'True
      End
   End
End
Attribute VB_Name = "Filecompare"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'File Names
Dim file(0 To 3) As String
'Stores whether the user has pressed the stop button or not
Dim stopcom As Boolean

'Other variables
Dim dif As Boolean, numofiles As Integer

Private Sub compare_Click()
Dim x As Long, fl1 As Long, fl2 As Long, shrtfil As Long
Dim h1 As Byte, h2 As Byte, h3 As Byte, h4 As Byte, count As Long
On Error GoTo errfix
'Makes sure that it hasn't started comparing
If compare.Caption <> "Stop" Then
 stopcom = False
 'Makes sure that two files have been selected
 If file(0) <> "" And file(1) <> "" Then
  compare.Caption = "Stop"
  'Open the files for comparing...
  Open file(0) For Binary As #1
  Open file(1) For Binary As #2
  'Clear the listbox...
  dumpbox.Clear
  'Find the file lengths...
  fl1 = LOF(1)
  fl2 = LOF(2)
  'Display them...
  dumpbox.AddItem "Length of file 1: " & fl1
  dumpbox.AddItem "Length of file 2: " & fl2
  'Give control back to windows for a sec...
  DoEvents
  count = 0
  'Find the shorter file
  If fl1 > fl2 Then
   shrtfil = fl2
  Else
   shrtfil = fl1
  End If
  'Depending on how many files are selected, do this...
  Select Case numofiles
  Case 2  '2 files
   'if searching for differences...
   If dif = True Then
   'Loop through the files...
   For x = 1 To shrtfil
    'Get the byte...
    Get #1, x, h1
    Get #2, x, h2
    'Compare them...
    If h1 <> h2 Then
     'Display the differences in either hex or decimal
     If mnuhex.Checked Then
      dumpbox.AddItem x & ":    " & Hex(h1) & ", " & Hex(h2)
     Else
      dumpbox.AddItem x & ":    " & h1 & ", " & h2
     End If
     count = count + 1
    End If
    'Windows...
    Label1.Caption = "Pos: " & x & ", Found: " & count
    DoEvents
    'If the user pressed stop...
    If stopcom = True Then Error (155)
   Next x
  Else
   'Searching for similarities...
   For x = 1 To shrtfil
    'Get the two bytes...
    Get #1, x, h1
    Get #2, x, h2
    'Depending on whether the user wants to see zeros...
    If mnuzero.Checked And h1 <> 0 Or Not (mnuzero.Checked) Then
     'If they are the same...
     If h1 = h2 Then
      'Display in either hex or decimal...
      If mnuhex.Checked Then
       dumpbox.AddItem x & ":    " & Hex(h1)
      Else
       dumpbox.AddItem x & ":    " & h1
      End If
      count = count + 1
     End If
    End If
    'Windows...
    Label1.Caption = "Pos: " & x & ", Found: " & count
    DoEvents
    'If user pressed stop...
    If stopcom = True Then Error (155)
   Next x
  End If
 Case 3
  'and so on for 3 files...
  Open file(2) For Binary As #3
  dumpbox.AddItem "Length of file 3: " & LOF(3)
  If LOF(3) < shrtfil Then shrtfil = LOF(3)
  If dif = False Then
   For x = 1 To shrtfil
    Get #1, x, h1
    Get #2, x, h2
    Get #3, x, h3
    If mnuzero.Checked And h1 <> 0 Or Not (mnuzero.Checked) Then
     If h1 = h2 And h1 = h3 Then
      If mnuhex.Checked Then
       dumpbox.AddItem x & ":    " & Hex(h1)
      Else
       dumpbox.AddItem x & ":    " & h1
      End If
      count = count + 1
     End If
    End If
    Label1.Caption = "Pos: " & x & ", Found: " & count
    DoEvents
    If stopcom = True Then Error (155)
   Next x
  Else
   MsgBox "You can only search for similarities with this many files"
  End If
 Case 4
  'And 4 files...
  Open file(2) For Binary As #3
  Open file(3) For Binary As #4
  dumpbox.AddItem "Length of file 3: " & LOF(3)
  dumpbox.AddItem "Length of file 4: " & LOF(4)
  If LOF(3) < shrtfil Then shrtfil = LOF(3)
  If LOF(4) < shrtfil Then shrtfil = LOF(4)
   If dif = False Then
    For x = 1 To shrtfil
     Get #1, x, h1
     Get #2, x, h2
     Get #3, x, h3
     Get #4, x, h4
     If mnuzero.Checked And h1 <> 0 Or Not (mnuzero.Checked) Then
      If h1 = h2 And h1 = h3 And h1 = h4 Then
       If mnuhex.Checked Then
        dumpbox.AddItem x & ":    " & Hex(h1)
       Else
        dumpbox.AddItem x & ":    " & h1
       End If
       count = count + 1
      End If
     End If
     Label1.Caption = "Pos: " & x & ", Found: " & count
     DoEvents
     If stopcom = True Then Error (155)
    Next x
   Else
    MsgBox "You can only search for similarities with this many files"
   End If
  End Select
  dumpbox.AddItem "Done"
  'Display the number of similarities or differences found...
  If dif = True Then
   dumpbox.AddItem "Total differences: " & count
  Else
   dumpbox.AddItem "Total similarities: " & count
  End If
  'Close all the files...
  Close
  compare.Caption = "Compare Files"
 Else
  'If files weren't selected...
  MsgBox "You must choose 2 files"
 End If
Else
 'The user pressed the stop button...
 compare.Caption = "Compare Files"
 stopcom = True
End If
Exit Sub

errfix:
If Err = 155 Then
 dumpbox.AddItem "Aborted"
 Close
 stopcom = False
Else
 MsgBox "Unexpected error"
 Close
End If
End Sub

Private Sub Form_Load()
 'Set the starting variables...
 dif = True
 numofiles = 2
End Sub

Private Sub mnudiff_Click()
 'the user clicked on the difference menu item
 'set the checks...
 mnusim.Checked = False
 mnudiff.Checked = True
 'set the variable...
 dif = True
End Sub

Private Sub mnufil_Click(Index As Integer)
 'Indexing controls that do similar things is so much easier
 Dim x As Integer
 'uncheck all the members of the array
 For x = 0 To 2
  mnufil(x).Checked = False
 Next x
 'check the one that was clicked
 mnufil(Index).Checked = True
 'set the number of files based on the one that was clicked
 numofiles = Index + 2
End Sub

Private Sub mnuhex_Click()
 'if the user clicked on the hex menu item
 'toggle the check...
 mnuhex.Checked = Not (mnuhex.Checked)
End Sub

Private Sub mnusim_Click()
 'The user wants to search for similarities...
 mnudiff.Checked = False
 mnusim.Checked = True
 dif = False
End Sub

Private Sub mnuzero_Click()
 'Toggle whether or not zeros are displayed...
 mnuzero.Checked = Not (mnuzero.Checked)
End Sub

Private Sub of_Click(Index As Integer)
 'User clicked one of the open file buttons
 On Error GoTo errfix
 'If a file hasn't yet been opened...
 If InStr(of(Index).Caption, "Open") Then
  'show the common dialog...
  CD.ShowOpen
  'set the file name variable...
  file(Index) = CD.FileName
  of(Index).Caption = "Close File" & Index + 1
 Else
  'empty the variable
  file(Index) = ""
  of(Index).Caption = "Open File " & Index + 1
 End If
 
errfix:
'the user clicked cancel
End Sub
