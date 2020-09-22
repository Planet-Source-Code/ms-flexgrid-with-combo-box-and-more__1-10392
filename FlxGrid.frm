VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Begin VB.Form FlxGrid 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Flex Grid Demo"
   ClientHeight    =   3360
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   5475
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3360
   ScaleWidth      =   5475
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Check1"
      Height          =   255
      Left            =   3840
      TabIndex        =   6
      Top             =   1200
      Width           =   1095
   End
   Begin VB.ListBox List1 
      Height          =   645
      ItemData        =   "FlxGrid.frx":0000
      Left            =   1440
      List            =   "FlxGrid.frx":0002
      TabIndex        =   5
      Top             =   720
      Width           =   1215
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   2880
      TabIndex        =   4
      Text            =   "Combo1"
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton cmdDelRow 
      Caption         =   "Del Row"
      Height          =   375
      Left            =   1200
      TabIndex        =   3
      Top             =   2880
      Width           =   975
   End
   Begin VB.CommandButton cmdAddRow 
      Caption         =   "Add Row"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   2880
      Width           =   975
   End
   Begin VB.CommandButton End 
      Caption         =   "End"
      Height          =   375
      Left            =   4320
      TabIndex        =   1
      Top             =   2880
      Width           =   975
   End
   Begin MSFlexGridLib.MSFlexGrid FlxGd 
      Height          =   2655
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   4683
      _Version        =   393216
      Rows            =   10
      Cols            =   5
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
   End
   Begin VB.Menu MnuFGridRows 
      Caption         =   "Rows"
      Visible         =   0   'False
      Begin VB.Menu MnuFGridAddRow 
         Caption         =   "Add a Row"
      End
      Begin VB.Menu MnuFGridSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDeleteGridRow 
         Caption         =   "Delete a Row"
      End
   End
End
Attribute VB_Name = "FlxGrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Dim x As Integer
    With FlxGd
        .ColAlignment(-1) = 1       'all Left alligned
        For x = 1 To .Cols - 1
           .TextMatrix(0, x) = "Col " + Str(x)
        Next
        For x = 1 To FlxGd.Rows - 1
           .TextMatrix(x, 0) = "Row " + Str(x)
        Next
        .Row = 1
        .Col = 1
        .CellBackColor = &HC0FFFF   'lt. yellow
    End With
    Combo1_Load
    List1_Load
    Check1_Load
End Sub

Private Sub cmdAddRow_Click()
   AddGridRow
End Sub

Private Sub cmdDelRow_Click()
   DeleteGridRow
End Sub

Private Sub FlxGd_EnterCell()
    FlxGd.CellBackColor = &HC0FFFF    'lt. yellow
    FlxGd.Tag = ""                    'clear temp storage
End Sub

Private Sub FlxGd_LeaveCell()
    If FlxGd.Col = 2 Then
      FlxGd = Format$(FlxGd, "#")      'alpha-number format
    Else
      If FlxGd.Col = 3 Then            'this is for Checkboxes
        If Check1.Value = 0 Then
          FlxGd.Text = "No"
        End If
      End If
      FlxGd = Format$(FlxGd, "0.00")
    End If
    FlxGd.CellBackColor = &H80000005
End Sub

Private Sub FlxGd_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case 46                 '<Del>, clear cell
        FlxGd.Tag = FlxGd   'assign to temp storage
        FlxGd = ""
  End Select
End Sub

Private Sub FlxGd_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 13            'ENTER key
            Advance_Cell   'advance new cell
        Case 8             'Backspace
            If Len(FlxGd) Then
              FlxGd = Left$(FlxGd, Len(FlxGd) - 1)
            End If
        Case 27                      'ESC
            If FlxGd.Tag > "" Then   'only if not NULL
              FlxGd = FlxGd.Tag      'restore original text
            End If
        Case Else
            FlxGd = FlxGd + Chr(KeyAscii)
    End Select
End Sub

Private Sub FlxGd_Click()
    If Combo1.Visible = True Then
      Combo1.Visible = False
      FlxGd.CellBackColor = &H80000005  'white
    Else
      If List1.Visible = True Then
        List1.Visible = False
        FlxGd.CellBackColor = &H80000005
      Else
        If Check1.Visible = True Then
          Check1.Visible = False
          FlxGd.CellBackColor = &H80000005
        End If
      End If
    End If
    If FlxGd.Col = 1 Then        ' Position and size the ListBox, then show it.
      List1.Width = FlxGd.CellWidth
      List1.Left = FlxGd.CellLeft + FlxGd.Left
      List1.Top = FlxGd.CellTop + FlxGd.Top
      List1.Text = FlxGd.Text
      List1.Visible = True
    Else
      If FlxGd.Col = 2 Then      ' Position and size the ComboBox, then show it.
        Combo1.Width = FlxGd.CellWidth
        Combo1.Left = FlxGd.CellLeft + FlxGd.Left
        Combo1.Top = FlxGd.CellTop + FlxGd.Top
        Combo1.Text = FlxGd.Text
        Combo1.Visible = True
      Else
        If FlxGd.Col = 3 Then    ' Position and size the CheckBox, then show it.
          Check1.Width = FlxGd.CellWidth
          Check1.Left = FlxGd.CellLeft + FlxGd.Left
          Check1.Top = FlxGd.CellTop + FlxGd.Top
          If FlxGd.Text = "Yes" Then
            Check1.Value = 1
          Else
            If FlxGd.Text = "No" Then
              Check1.Value = 0
            End If
          End If
          Check1.Visible = True
        End If
      End If
    End If
End Sub

Private Sub Check1_Click()   ' Place the selected Yes/No into the Cell and hide the CheckBox.
    If FlxGd.Col = 3 Then
      If Check1.Value = 1 Then
        FlxGd.Text = "Yes"
      Else
        If Check1.Value = 0 Then
          FlxGd.Text = "No"
        End If
      End If
      Check1.Visible = False
    End If
End Sub

Private Sub Check1_Load()    ' Load the Checkbox.
    Check1.Visible = False
    Check1.Value = False
    Check1.Caption = "Test?"
    Check1.Width = FlxGd.CellWidth
End Sub

Private Sub List1_Load()    ' Load the ListBox's list.
    List1.Visible = False
    List1.Width = FlxGd.CellWidth
    List1.AddItem ""
    List1.AddItem "Dog"
    List1.AddItem "Cat"
    List1.AddItem "Fish"
    List1.ListIndex = 0
End Sub

Private Sub List1_Click()   ' Place the selected item into the Cell and hide the ListBox.
    If FlxGd.Col = 1 Then
      FlxGd.Text = List1.Text
      List1.Visible = False
    End If
End Sub

Private Sub Combo1_Click()  ' Place the selected item into the Cell and hide the ComboBox.
    If FlxGd.Col = 2 Then
      FlxGd.Text = Combo1.Text
      Combo1.Visible = False
    End If
End Sub

Private Sub Combo1_Load()   ' Load the ComboBox's list.
    FlxGd.RowHeightMin = Combo1.Height
    Combo1.Visible = False
    Combo1.Width = FlxGd.CellWidth
    Combo1.AddItem "1"
    Combo1.AddItem "2"
    Combo1.AddItem "3"
End Sub
   
Private Sub FlxGd_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim Row As Integer, Col As Integer
    Row = FlxGd.MouseRow
    Col = FlxGd.MouseCol
    If Button = 2 And (Col = 0 Or Row = 0) Then
      FlxGd.Col = IIf(Col = 0, 1, Col)
      FlxGd.Row = IIf(Row = 0, 1, Row)
      PopupMenu MnuFGridRows
    End If
End Sub

Private Sub MnuFGridAddRow_Click()
   AddGridRow
End Sub

Private Sub AddGridRow()
    With FlxGd
        .Rows = .Rows + 1
        .Row = .Rows - 1
        .TextMatrix(.Row, 0) = "Row " + Str(.Row)
    End With
End Sub

Private Sub mnuDeleteGridRow_Click()
   DeleteGridRow
End Sub

Private Sub DeleteGridRow()
    Dim Row As Integer, n As Integer, x As Integer
    With FlxGd
        If .Rows > 2 Then        'make sure we don't del a row
          Row = .Row
          For n = 1 To .Cols - 1
             If .TextMatrix(Row, n) > "" Then
               x = 1
               Exit For
             End If
          Next
          If x Then
            n = MsgBox("Data in Row" + Str$(Row) + ".  Delete anyway?", vbYesNo, "Delete Row...")
          End If
          If x = 0 Or n = 6 Then           'no exist. data or YES
            For n = .Row To .Rows - 2      'move exist data up 1 row
               For x = 1 To FlxGd.Cols - 1
                  .TextMatrix(n, x) = .TextMatrix(n + 1, x)
               Next
            Next
            If Row = .Rows - 1 Then     'set new cursor row
              .Row = .Rows - 2
            End If
            .Rows = .Rows - 1           'delete last row
          End If
        End If
    End With
End Sub

Private Sub End_Click()
    End
End Sub

Private Sub Advance_Cell()                  'advance to next cell
    With FlxGd
        .HighLight = flexHighlightNever     'turn off hi-lite
        If .Col < .Cols - 1 Then
          .Col = .Col + 1
        Else
          If .Row < .Rows - 1 Then
            .Row = .Row + 1                 'down 1 row
            .Col = 1                        'first column
          Else
            .Row = 1
            .Col = 1
          End If
        End If
        If .CellTop + .CellHeight > .Top + .Height Then
          .TopRow = .TopRow + 1             'make sure row is visible
        End If
        .HighLight = flexHighlightAlways    'turn on hi-lite
    End With
End Sub

