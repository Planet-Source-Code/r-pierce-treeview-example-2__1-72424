VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmTVex 
   Caption         =   "                   Music File Database                                                 (another TreeView example by:  R. Pierce)"
   ClientHeight    =   7185
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   11580
   LinkTopic       =   "Form1"
   ScaleHeight     =   7185
   ScaleWidth      =   11580
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Printer Fonts Available"
      Height          =   495
      Left            =   10260
      TabIndex        =   18
      Top             =   4500
      Width           =   1125
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Insert Before Selected Node"
      Height          =   255
      Left            =   7440
      TabIndex        =   17
      Top             =   2250
      Width           =   2355
   End
   Begin VB.TextBox Text4 
      Height          =   3405
      Left            =   5280
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   16
      Top             =   2610
      Width           =   4875
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Search is case sensitive"
      Height          =   255
      Left            =   8070
      TabIndex        =   15
      Top             =   6480
      Width           =   2025
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Search"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   10260
      TabIndex        =   14
      Top             =   6450
      Width           =   945
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   6390
      TabIndex        =   12
      Top             =   6090
      Width           =   4965
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Parent Nodes"
      Height          =   345
      Left            =   10260
      TabIndex        =   11
      ToolTipText     =   "The Parent Node of the selected node"
      Top             =   3600
      Width           =   1125
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Root Node"
      Height          =   345
      Left            =   10260
      TabIndex        =   10
      ToolTipText     =   "The Root Node  (there's only one)"
      Top             =   3180
      Width           =   1125
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Child Nodes"
      Height          =   345
      Left            =   10260
      TabIndex        =   9
      ToolTipText     =   "Child Nodes of the selected node"
      Top             =   4020
      Width           =   1125
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Add/Insert  Node"
      Height          =   315
      Left            =   9840
      TabIndex        =   8
      ToolTipText     =   "HighLight A  Node, This Will Add A Child To Selected Parent  - or -  Insert A Node Before or After Child Selected"
      Top             =   2220
      Width           =   1545
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Add Node"
      Height          =   315
      Left            =   10320
      TabIndex        =   5
      ToolTipText     =   "This is the Parent Name. (1st node is also the root)"
      Top             =   960
      Width           =   1065
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   5280
      TabIndex        =   2
      Top             =   1710
      Width           =   6075
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   5310
      TabIndex        =   1
      Top             =   450
      Width           =   6045
   End
   Begin MSComctlLib.TreeView TV 
      Height          =   6945
      Left            =   60
      TabIndex        =   0
      Top             =   120
      Width           =   5025
      _ExtentX        =   8864
      _ExtentY        =   12250
      _Version        =   393217
      HideSelection   =   0   'False
      LineStyle       =   1
      Style           =   7
      Appearance      =   1
   End
   Begin MSComDlg.CommonDialog CDlg 
      Left            =   2310
      Top             =   5370
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label6 
      Caption         =   "Search Titles:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5250
      TabIndex        =   13
      Top             =   6120
      Width           =   1125
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "The Song Titles/Artists Contained In The Music File"
      Height          =   225
      Left            =   5310
      TabIndex        =   7
      Top             =   1410
      Width           =   6045
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "The Music Filename"
      Height          =   225
      Left            =   5340
      TabIndex        =   6
      Top             =   150
      Width           =   6015
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   225
      Left            =   10050
      TabIndex        =   4
      Top             =   6840
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   225
      Left            =   5310
      TabIndex        =   3
      Top             =   6840
      Width           =   4575
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuOpen 
         Caption         =   "Open"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "Save"
      End
      Begin VB.Menu mnuBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPrint 
         Caption         =   "Print"
      End
      Begin VB.Menu mnuBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuRClick 
      Caption         =   "RClick"
      Visible         =   0   'False
      Begin VB.Menu mnuDelete 
         Caption         =   "Delete"
      End
   End
End
Attribute VB_Name = "frmTVex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' This is a TreeView example which there are so many of I can understand it!
' Feel free to update/upgrade this and re-post it.  R. Pierce 7-28-2009
' Uploaded to VBFreeCode 8-21-09
' This is the final update unless I do recursive read/write to save any configuration you came up with...
'
Const BC = &H80000005
Const BC_Err = &HC0C0FF
Dim entry As Node
Dim song As Node
Dim w As Long
Dim x As Long
Dim y As Long
Dim z As Long
Dim nam As String
Dim fnam As String
Dim KCode As Integer

Private Sub Command1_Click()
Text4 = "": Dim flg As Byte: flg = 0
For w = 1 To TV.Nodes.Count ' step thru all of the nodes
    If TV.Nodes(w).Parent Is Nothing Then 'this is the top node
        For z = 1 To TV.Nodes(w).Children
            If z = 1 Then
                x = TV.Nodes(w).Child.Index
            Else
                x = TV.Nodes(x).Next.Index
            End If
            If Check1.Value = 0 Then
                If InStrRev(TV.Nodes(x).Text, Text3.Text, , 1) > 0 Then 'search the children node's Titles
                    Text4 = Text4 + TV.Nodes(w).Text + "  -  " + TV.Nodes(x).Text + Format$(x, " 00") + vbCrLf
                    flg = 1
                End If
            Else
                If InStrRev(TV.Nodes(x).Text, Text3.Text, , 0) > 0 Then 'search the children node's Titles
                    Text4 = Text4 + TV.Nodes(w).Text + "  -  " + TV.Nodes(x).Text + Format$(x, " 00") + vbCrLf
                    flg = 1
                End If
            End If
        Next z
    End If
Next w
If flg = 0 Then Text4 = "Search String Not Found"
End Sub

Private Sub Command2_Click()
'show available printer fonts
Text4.Text = vbNull
For x = 0 To Printer.FontCount - 1
        Text4.Text = Text4.Text & Printer.Fonts(x) & vbCrLf
    Next x


End Sub

Private Sub Command5_Click()
On Error GoTo errmsg
If Text1.Text = "" Then
    Text4 = "Field Is Empty, Can Not Be Added To Database"
    Exit Sub
Else ' User has type an entry for the TreeView
    'Lets make sure the new filename does not already exist in our database (song titles are not searched)
    If TV.Nodes.Count > 0 Then
        For x = 1 To TV.Nodes.Count ' (1 is the root node of the tree)
            If TV.Nodes(x).Parent Is Nothing Then 'were only searching the top nodes which have no parents
                If Text1.Text = TV.Nodes(x).Text Then
                    Text1.BackColor = BC_Err
                    Text4 = "Name is already in the database..."
                    Exit Sub
                End If
            End If
        Next x
    End If

    Text1.BackColor = BC
    x = TV.Nodes.Count + 1

    If y = 0 Then
        Set entry = TV.Nodes.Add
    Else ' you highlighted a node, insert the node following the highlighted one (top nodes only)
        Set entry = TV.Nodes.Add(y)
    End If
    'On Error GoTo errmsg
    entry.Text = Text1.Text
    'de-highlight the former selected node that you just inserted another one after
    If y > 0 Then
        TV.DropHighlight = Nothing
        TV.SelectedItem = Nothing ' delete this and the former selected item can still be seen
    End If
    Label2 = "Node Count: " & TV.Nodes.Count
    y = 0 'reset y after an entry has been made
End If
Exit Sub
errmsg: 'just for the impossible ):
ret = MsgBox("Error " & Err.Number & " reported." & vbCrLf & Err.Description, vbCritical, "Error Reported")

End Sub

Private Sub Command6_Click()
If Text2 = "" Then
    Text4 = "Field is empty, type a name for the child, then save"
    Exit Sub
End If

If y > 0 Then   'a node has been selected with y holding the selections index.
                '
    If TV.SelectedItem.Parent Is Nothing Then   'were just adding another entry to the selected parent's nodes
                                                'collection which is added in the next available position
        Set entry = TV.Nodes.Add(y, tvwChild, , Text2.Text)
        Text4 = ""
    Else 'were inserting a node preceding or following the selected node as user desires
        If Check2.Value = 0 Then
            Set entry = TV.Nodes.Add(y, tvwNext, , Text2.Text) 'following
        Else
            Set entry = TV.Nodes.Add(y, tvwPrevious, , Text2.Text) 'preceding
        End If
    End If
    Label2 = "Node Count: " & TV.Nodes.Count
Else
    Text4 = "Select Which Node To Store Titles/Artists In"
End If
End Sub

Private Sub Command7_Click()
If y > 0 Then 'Display the child nodes and indexes
Text4 = ""
    For z = 1 To TV.Nodes(y).Children
      If z = 1 Then
          x = TV.Nodes(y).Child.Index
      Else
          x = TV.Nodes(x).Next.Index
      End If
    Text4 = Text4 & TV.Nodes(x).Text & " - " & Str(x) & vbCrLf
    Next z
Else
    Text4 = "Nothing selected to process"
End If
End Sub

Private Sub Command8_Click()
' the 1st node is the root node... All nodes testify to this!
' A tree has ONLY ONE root node - whose parent property is always nothing
' just like all of the other top nodes, except they are parents not the one and only root node
' Whose a bastard without any parents -- Cold World
If y > 0 Then 'Show selected root node
On Error GoTo what_error
    'Text4 = entry.Root.Text + vbCrLf 'try this as soon as you load something (comment out the line below)
    Text4 = TV.SelectedItem.Root
Else
    Text4 = "Nothing selected to process"
End If
Exit Sub

what_error:
Text4 = ""
rtrn = MsgBox("Error : " & Str(Err.Number) + vbCrLf & Err.Description, vbCritical, "ERROR RETURNED")

End Sub

Private Sub Command9_Click()
If y > 0 Then
     If TV.SelectedItem.Parent Is Nothing Then
        Text4 = "This is a parent - " & TV.SelectedItem.Text
     Else
        Text4 = TV.SelectedItem.Parent
     End If
Else
    Text4 = "Nothing selected to process"
End If
End Sub

Private Sub Form_Load()
Label1 = "No Database Loaded"
Label2 = "Node Count: " & TV.Nodes.Count
frmTVex.Caption = " SCM's  TreeView Example      Version " & App.Major & "." & App.Minor & "   build " & App.Revision & "        " & fnam
End Sub

Private Sub Form_Resize()
modResize.Resize Me

Select Case frmTVex.WindowState
Case 1:
frmTVex.Caption = "SCM's TV Example"
Case Else
frmTVex.Caption = " SCM's  TreeView Example   Version " & App.Major & "." & App.Minor & "   build " & App.Revision & "        " & fnam
End Select

End Sub

Private Sub mnuDelete_Click()
            TV.Nodes.Remove (y)
            TV.DropHighlight = Nothing
            TV.SelectedItem = Nothing
            Label2 = "Node Count: " & TV.Nodes.Count
            Label1.Caption = Label1.Caption + "   <DELETED>"
End Sub

Private Sub mnuExit_Click()
Unload Me
End Sub

Private Sub mnuOpen_Click()
'Open (load) Database File Menu
On Error GoTo Cancel ' if user cancels just bypass load routine
Dim filnum As Integer:  Dim tmp$: Dim tmp2$
TV.Nodes.Clear ' start with an empty TreeView
CDlg.InitDir = App.Path
CDlg.CancelError = True
CDlg.Flags = &H4 Or &H800 Or &H1000 ' file must exist|no read only checkbox|path must exist
CDlg.InitDir = App.Path
CDlg.DialogTitle = "Load Music Database"
CDlg.Filter = "Music Database Files (*.mdb)|*.mdb"
CDlg.ShowOpen 'user selected a file to load
filnum = FreeFile ' obtain a free file number
Open CDlg.FileName For Input As #filnum ' open selected file
'load the TreeView from file
While Not EOF(filnum)
    Set entry = TV.Nodes.Add 'set-up a parent node to add
    Line Input #filnum, tmp$ 'get the parent node text
    entry.Text = tmp$ 'add the node to the TreeView
    Line Input #filnum, tmp$ 'get the number of children this node has
    For x = 1 To Val(tmp$) ' loop thru the child nodes
        Line Input #filnum, tmp2$ 'get the child name
        Set song = TV.Nodes.Add(entry, tvwChild, , tmp2$) 'add the child node
    Next x
Wend
Close #filnum 'were done
Label2 = "Node Count: " & TV.Nodes.Count
Label1 = "[ " + CDlg.FileTitle + " ]" & "      >LOADED<"
Text1 = "": Text2 = "": Text4 = ""
fnam = CDlg.FileName
frmTVex.Caption = " SCM's  TreeView Example   Version " & App.Major & "." & App.Minor & "   build " & App.Revision & "        " & fnam
Cancel:
End Sub

Private Sub mnuPrint_Click()
' Print Menu Option from the File menu selection
If TV.Nodes.Count = 0 Then
    Text4 = "Database Is Empty - Nothing To Print"
    Exit Sub 'Nothing to save
End If
' Print the filepath
Printer.Font = "Courier New": Printer.FontSize = 14: Printer.FontBold = True: _
Printer.Print: Printer.Print
If InStr(fnam, ".") > 0 Then
    Printer.Print Left$(fnam, InStrRev(fnam, ".") - 1)
Else
    Printer.Print fnam
End If
' Print the date/time stamp
Printer.Font = "Georgia": Printer.FontSize = 8: Printer.FontBold = False: Printer.Print Date$ & " / " & Time$
Printer.Print
Printer.Print: Printer.Print
' Print the treeview control
For w = 1 To TV.Nodes.Count ' step thru all of the nodes
    If TV.Nodes(w).Parent Is Nothing Then 'this is a top node
        Printer.Font = "Arial": Printer.FontSize = 13: Printer.FontBold = False
        Printer.Print TV.Nodes(w).Text  'print parent text (filename that is used to save the below songs)
        Printer.Font = "Times New Roman": Printer.FontSize = 12
        For z = 1 To TV.Nodes(w).Children
            If z = 1 Then ' the control points us to the next index
               x = TV.Nodes(w).Child.Index
            Else
                x = TV.Nodes(x).Next.Index
            End If
            Printer.Print , TV.Nodes(x).Text 'print the child text (song and artist)
        Next z
        Printer.Print ' line between parent nodes
    End If
Next w
' Done printing
Printer.EndDoc

End Sub

Private Sub mnuSave_Click()
'Save Database To File Menu
Dim filnum As Integer
On Error GoTo Cancel
'Check to make sure the database isn't empty (The TreeView)
If TV.Nodes.Count = 0 Then
    Text4 = "Database Is Empty - No File Saved"
    Exit Sub 'Nothing to save
End If
CDlg.CancelError = True
CDlg.Flags = &H4 Or &H800 Or &H1000 ' file must exist|no read only checkbox|path must exist
CDlg.InitDir = App.Path
CDlg.DialogTitle = "Save Music Database to File"
CDlg.Filter = "Music Database Files (*.mdb)|*.mdb"
CDlg.ShowSave

filnum = FreeFile ' obtain a free file number
Open CDlg.FileName For Output As #filnum

For w = 1 To TV.Nodes.Count ' step thru all of the nodes
    If TV.Nodes(w).Parent Is Nothing Then 'this is a top node
        Print #filnum, TV.Nodes(w).Text 'print parent text (filename that is used to save the below songs)
        Print #filnum, TV.Nodes(w).Children 'print the number of children were saving (songs and artists | 1 thru z)
        For z = 1 To TV.Nodes(w).Children
            If z = 1 Then
                x = TV.Nodes(w).Child.Index
            Else
                x = TV.Nodes(x).Next.Index
            End If
            Print #filnum, TV.Nodes(x).Text 'save the child text to file (song and artist)
        Next z
    End If
Next w
Close #filnum

Text4 = CDlg.FileName & "   >saved to disk<"
Exit Sub
Cancel:

End Sub

Private Sub Text1_Change()
Text1.BackColor = BC
If Len(Text4) > 0 Then
Text4 = ""
End If
End Sub

Private Sub Text2_Change()
If Len(Text4) > 0 Then
Text4 = ""
End If
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    KeyAscii = 0
    Command1_Click
End If
End Sub

Private Sub TV_KeyDown(KeyCode As Integer, Shift As Integer)
Dim rtrn As Byte
If KeyCode = 46 Then
    If y > 0 Then
        rtrn = MsgBox("", vbYesNo, "Delete Highlighted Node")
        If rtrn = 6 Then
            TV.Nodes.Remove (y)
            TV.DropHighlight = Nothing
            TV.SelectedItem = Nothing
            Label2 = "Node Count: " & TV.Nodes.Count
            Label1.Caption = Label1.Caption + "   <DELETED>"
        End If
    End If
End If
End Sub

Private Sub TV_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then
PopupMenu mnuRClick
End If
End Sub

Private Sub TV_NodeClick(ByVal Node As MSComctlLib.Node)
If TV.Nodes.Count > 0 Then 'prevents an error if TV is empty
y = TV.SelectedItem.Index ' when a node has been clicked, we save the nodes's index (position) in y
'Label1.Caption = Str(y)
nam = TV.SelectedItem.Text
Label1.Caption = nam & " - " & Str(y)
Set TV.DropHighlight = TV.SelectedItem
End If
End Sub
