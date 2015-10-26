VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   9975
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   17700
   LinkTopic       =   "Form1"
   ScaleHeight     =   9975
   ScaleWidth      =   17700
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command18 
      Caption         =   "numgrid"
      Height          =   375
      Left            =   14640
      TabIndex        =   22
      Top             =   2640
      Width           =   1095
   End
   Begin VB.ListBox List8 
      Height          =   840
      Left            =   15600
      TabIndex        =   21
      Top             =   7560
      Width           =   1095
   End
   Begin VB.CommandButton Command17 
      Caption         =   "BackTrace"
      Height          =   375
      Left            =   13560
      TabIndex        =   20
      Top             =   2040
      Width           =   975
   End
   Begin VB.CommandButton Command16 
      Caption         =   "Search 3"
      Height          =   495
      Left            =   16080
      TabIndex        =   19
      Top             =   2640
      Width           =   855
   End
   Begin VB.CommandButton Command15 
      Caption         =   "Search 2"
      Height          =   375
      Left            =   16080
      TabIndex        =   18
      Top             =   2160
      Width           =   855
   End
   Begin VB.CommandButton Command14 
      Caption         =   "auto >>"
      Height          =   375
      Left            =   14760
      TabIndex        =   17
      Top             =   2040
      Width           =   855
   End
   Begin VB.CommandButton Command13 
      Caption         =   "Search 1"
      Height          =   375
      Left            =   13560
      TabIndex        =   16
      Top             =   1440
      Width           =   855
   End
   Begin VB.CommandButton Command12 
      Caption         =   "step rev 2"
      Height          =   375
      Left            =   15960
      TabIndex        =   14
      Top             =   1320
      Width           =   1095
   End
   Begin VB.CommandButton Command11 
      Caption         =   "step rev 1"
      Height          =   375
      Left            =   14760
      TabIndex        =   13
      Top             =   840
      Width           =   855
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   150
      Left            =   16560
      Top             =   3240
   End
   Begin VB.CommandButton Command10 
      Caption         =   "auto"
      Height          =   375
      Left            =   14760
      TabIndex        =   12
      Top             =   1440
      Width           =   855
   End
   Begin VB.CommandButton Command9 
      Caption         =   "setmouse"
      Height          =   375
      Left            =   13560
      TabIndex        =   11
      Top             =   840
      Width           =   855
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   13440
      Top             =   7200
   End
   Begin VB.FileListBox File1 
      Height          =   2820
      Left            =   13560
      Pattern         =   "*.grd"
      TabIndex        =   9
      Top             =   7920
      Width           =   1455
   End
   Begin VB.CommandButton Command8 
      Caption         =   "load"
      Height          =   615
      Left            =   13800
      TabIndex        =   8
      Top             =   7200
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton Command7 
      Caption         =   "save"
      Height          =   495
      Left            =   13680
      TabIndex        =   7
      Top             =   7320
      Width           =   1335
   End
   Begin VB.CommandButton Command6 
      Caption         =   "step"
      Height          =   375
      Left            =   15960
      TabIndex        =   6
      Top             =   1680
      Width           =   1095
   End
   Begin VB.CommandButton Command5 
      Caption         =   "erase"
      Height          =   375
      Left            =   13560
      TabIndex        =   5
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton Command4 
      Caption         =   "wall"
      Height          =   375
      Left            =   14760
      TabIndex        =   4
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton Command3 
      Caption         =   "end"
      Height          =   375
      Left            =   15960
      TabIndex        =   3
      Top             =   960
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Start"
      Height          =   375
      Left            =   15960
      TabIndex        =   2
      Top             =   600
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "draw grid"
      Height          =   375
      Left            =   13560
      TabIndex        =   1
      Top             =   2640
      Width           =   975
   End
   Begin MSFlexGridLib.MSFlexGrid G 
      Height          =   10935
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   13455
      _ExtentX        =   23733
      _ExtentY        =   19288
      _Version        =   393216
      Rows            =   1
      Cols            =   1
   End
   Begin VB.Line Line2 
      X1              =   13560
      X2              =   15720
      Y1              =   2520
      Y2              =   2520
   End
   Begin VB.Line Line1 
      X1              =   15720
      X2              =   15720
      Y1              =   2520
      Y2              =   0
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   495
      Left            =   16320
      TabIndex        =   15
      Top             =   3840
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   375
      Left            =   13920
      TabIndex        =   10
      Top             =   6600
      Width           =   495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim ds   As Integer, de As Integer
Dim cx As Integer, cy As Integer
Dim corncx As Integer, corncy As Integer
Dim endcx As Integer, endcy As Integer
Dim startcy As Integer, startcx As Integer

Dim popbackindex As Integer

Dim matds(32, 32) As Integer, matde(32, 32) As Integer
Dim list(10, 256) As Integer
Dim ptr(10) As Integer


Dim isended As Boolean
Dim traceback As Boolean
Dim issearching As Boolean
Dim ispopBack As Boolean
Dim searchinyellow As Boolean
Dim isapproved As Boolean
Dim isnomorevaluetopop  As Boolean
Dim isendofexplore As Boolean


Dim matsize As Integer
Dim pen As String
Dim OpenFile As String
Dim prevX As Integer, prevY As Integer

Dim isnumgrid As Boolean


Dim isfinished As Boolean
Private Sub Command1_Click()
matsize = 36
G.Rows = matsize
G.Cols = matsize

Dim i As Integer
Dim j As Integer

For i = 0 To matsize - 1
If i Mod 2 = 0 Then
    G.ColWidth(i) = G.Width / matsize + 180
    G.RowHeight(i) = G.Height / matsize + 180
Else
    G.ColWidth(i) = G.Width / matsize - 200
    G.RowHeight(i) = G.Height / matsize - 200
End If

    G.TextMatrix(0, i) = i & "C"
    G.TextMatrix(i, 0) = i & "R"
    
    G.ColAlignment(i) = 4
    
    
Next

cx = corncx
cy = corncy



End Sub

Private Sub Command10_Click()
If Timer2.Enabled = True Then
    Timer2.Enabled = False
Else
    Timer2.Enabled = True
End If
    
End Sub

Private Sub Command11_Click()
getNextCell_rev
End Sub



Private Sub Command13_Click()
cx = 16
cy = 16
endcx = 2
endcy = 2

getshortestpath

Exit Sub
issearching = True
isended = False
traceback = False
de = 0
ds = 1
cx = InputBox("cx") ' 2
cy = InputBox("cy") ' 2
endcx = InputBox("endcx") ' 16
endcy = InputBox("endcy") '  16

G.TextMatrix(cy, cx) = 0
G.Col = cx
G.Row = cy
G.CellBackColor = vbGreen

ptr(6) = 0
ptr(7) = 0
ptr(8) = 0
ptr(9) = 0
ptr(10) = 0

'List6.Clear
'List7.Clear
'List8.Clear
'List9.Clear
'List10.Clear


End Sub

Private Sub Command14_Click()
While (isfinished = False)
    Command11_Click
Wend
End Sub






Private Sub traceback_init()
    traceback = True
    issearching = False
    G.Col = cx
    G.Row = cy
    G.CellBackColor = vbCyan
    isended = False
    ptr(6) = 0
    ptr(7) = 0
    ptr(8) = 0
    ptr(9) = 0
    ptr(10) = 0
End Sub

Private Sub Command17_Click()
traceback_init
End Sub


Private Sub Command2_Click()
pen = "s"

End Sub

Private Sub Command3_Click()
pen = "e"

End Sub

Private Sub Command4_Click()
pen = "w"
End Sub


Private Sub Command5_Click()
pen = ""
End Sub

Private Sub Command7_Click()
Dim allCells As String
Dim fnum As Integer
Dim curRow, curCol As Integer

    curRow = G.Row
    curCol = G.Col
    
    'CommonDialog1.DefaultExt = "GRD"
    'CommonDialog1.Action = 2
    'If CommonDialog1.FileName = "" Then Exit Sub
    EditSelect_Click
    allCells = G.Clip
    fnum = FreeFile
    Open App.Path & "\" & File1.ListCount + 1 & ".grd" For Output As #fnum
    Write #fnum, allCells
    Close #fnum
    
    G.Row = curRow
    G.Col = curCol
    G.RowSel = G.Row
    G.ColSel = G.Col
    
    File1.Refresh
End Sub

Private Sub Command8_Click()
Command1_Click
Dim allCells As String
Dim fnum As Integer
Dim curRow, curCol As Integer

'On Error GoTo NoFileSelected
    OpenFile = App.Path & "\" & File1.FileName
    fnum = FreeFile
    Open OpenFile For Input As #fnum
    Input #fnum, allCells
    EditSelect_Click
    G.Clip = allCells
    Close #fnum

    G.Row = 1
    G.Col = 1
    G.RowSel = G.Row
    G.ColSel = G.Col
G.Visible = False
Dim i As Integer, j As Integer
For i = 0 To matsize - 2
    For j = 0 To matsize - 2
    G.Col = j
    G.Row = i
    If G.Text = "w" Then
        G.CellBackColor = vbRed
    Else
        G.CellBackColor = vbWhite
    End If
    Next

Next

G.Visible = True
    Exit Sub

NoFileSelected:
    Exit Sub
    
End Sub

Private Sub EditSelect_Click()
    G.Row = 1
    G.Col = 1
    G.RowSel = G.Rows - 1
    G.ColSel = G.Cols - 1
End Sub

Private Sub Command9_Click()
pen = "setmouse"
traceback = False
issearching = False
isended = False
ptr(1) = 0
ptr(2) = 0
ptr(3) = 0
ptr(4) = 0
ptr(5) = 0

de = 0
ds = 1

End Sub

Private Sub File1_Click()
Command8_Click
cx = corncx
cy = corncy
G.Row = cy
G.Col = cx
G.CellBackColor = vbYellow
End Sub


Private Sub Form_Load()
File1.Path = App.Path
corncx = 2
corncy = 32
ds = -1

cx = corncx
cy = corncy
End Sub

Private Sub G_KeyPress(KeyAscii As Integer)
Timer1.Enabled = False
G.Text = G.Text & KeyAscii - 48
End Sub

Private Sub G_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next
G.ToolTipText = "dS = " & matds(G.MouseRow, G.MouseCol) & " dE = " & matde(G.MouseRow, G.MouseCol)

End Sub


Private Sub G_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If IsNumeric(G.Text) = True And (pen = "w") Then
ElseIf pen = "setmouse" Then
    cx = G.Col
    cy = G.Row
    corncx = cx
    corncy = cy
Else
    G.Text = pen
    If G.Text = "w" Then
        G.CellBackColor = vbRed
    Else
        G.CellBackColor = vbWhite
    End If
End If
End Sub


Private Sub Timer1_Timer()
If Not pen = "setmouse" Then Exit Sub
On Error Resume Next
Dim tcx As Integer, tcy As Integer
tcx = G.Col
tcy = G.Row

G.Col = cx
G.Row = cy

If G.CellFontBold = True Then
    G.CellFontBold = False
Else
    G.CellFontBold = True
End If

If ds = 1 Then
    Label1 = "V"
ElseIf ds = -1 Then
    Label1 = "^"
ElseIf de = 1 Then
    Label1 = ">"
ElseIf de = True Then
    Label1 = "<"
End If

Label2 = ptr(1)

End Sub

Public Sub getNextCell_rev()
Dim i As Integer
Dim tde As Integer, tds As Integer, tcx As Integer, tcy As Integer, fde As Integer, fds As Integer
Dim t1 As Integer
Dim success As Boolean, followWhite As Boolean

up0:
If ispopBack = True And isapproved = True And traceback = False And issearching = False Then GoTo s3

If isendofexplore = True And cx = corncx And cy = corncy And ispopBack = False And traceback = False And isfinished = False Then
    endcx = 16
    endcy = 16
    isapproved = True
    ispopBack = True
    getshortestpath
    traceback_init
    isfinished = True
    Exit Sub
End If
                

If isfinished = True And cx = 16 And cy = 16 And ispopBack = False And traceback = False Then
    changecolor vbCyan, vbYellow
    changecolor vbGreen, vbYellow
    startcx = 16
    startcy = 16
    endcx = corncx
    endcy = corncy
    getshortestpath
    traceback_init
    Exit Sub
End If
    
If traceback = True Then
    If endcx = cx And endcy = cy Then
        Timer2.Enabled = False
        isended = True
        If ispopBack = True Then
            ispopBack = False
            traceback = False
            changecolor vbCyan, vbYellow
            changecolor vbGreen, vbYellow
        End If
        Exit Sub
    End If
End If


tde = de
tds = ds

CheckCase:


If Not G.TextMatrix(cy + tds, cx + tde) = "w" Then
    G.Col = cx + tde * 2
    G.Row = cy + tds * 2
    If traceback = True Then
        If G.CellBackColor = vbGreen Then
            GoTo s2
        Else
            GoTo SkipIt
        End If
    End If

    If Not G.CellBackColor = vbYellow Then
s2:

        If issearching = True Then GoTo SkipIt
        
        If CInt(G.TextMatrix(cy + tds * 2, cx + tde * 2)) < CInt(G.TextMatrix(cy, cx)) And followWhite = False Then
            If success = False And followWhite = False Then
            If ispopBack = True And isapproved = False Then
                i = 0
                isapproved = True
                GoTo up0:
            End If
                matde(cy, cx) = tde
                matds(cy, cx) = tds
                tcy = cy + tds * 2
                tcx = cx + tde * 2
                fde = tde
                fds = tds
                G.Col = tcx
                G.Row = tcy
                If traceback = True Then
                    G.CellBackColor = vbCyan
                Else
                    G.CellBackColor = vbYellow
                End If
                success = True
            Else
                listadd 1, cy
                listadd 2, cx
                listadd 4, de
                listadd 5, ds
                listadd 3, ptr(1) - 1
            End If
        End If
        
        If followWhite = True Then
            If success = False Then
                If ispopBack = True And isapproved = False Then
                    isapproved = True
                    i = 0
                    GoTo up0:
                End If
                    matde(cy, cx) = tde
                    matds(cy, cx) = tds
                    tcy = cy + tds * 2
                    tcx = cx + tde * 2
                    fde = tde
                    fds = tds
                    G.Col = tcx
                    G.Row = tcy
                    G.CellBackColor = vbYellow
                    success = True
            ElseIf Not CInt(G.TextMatrix(cy + tds * 2, cx + tde * 2)) < CInt(G.TextMatrix(cy, cx)) Then
                listadd 1, cy
                listadd 2, cx
                listadd 4, de
                listadd 5, ds
                listadd 3, ptr(1) - 1
            End If
        End If
    ElseIf issearching = True Then
        G.TextMatrix(cy + 2 * tds, cx + 2 * tde) = CInt(G.TextMatrix(cy, cx)) + 1
        listadd 6, cy + 2 * tds
        listadd 7, cx + 2 * tde
        listadd 9, de
        listadd 10, ds
        G.CellBackColor = vbGreen
        listadd 8, ptr(6) - 2
        If cy + 2 * tds = endcy And cx + 2 * tde = endcx Then
           isended = True
        End If
    End If
End If

SkipIt:
i = i + 1
up1:
    Select Case i
    Case 1
    If de = 0 Then
        tde = 1
        tds = 0
    Else
        i = 2
        GoTo up1
    End If
    Case 2
    If de = 0 Then
        tde = -1
        tds = 0
    Else
        i = 3
        GoTo up1
    End If
    Case 3
    If ds = 0 Then
        tds = 1
        tde = 0
    Else
        i = 4
        GoTo up1
    End If
    Case 4
    If ds = 0 Then
        tds = -1
        tde = 0
    Else
        i = 5
        GoTo up1
    End If
    Case 5
    If de = 0 Then
        tds = -ds
        tde = 0
    Else
        i = 6
        GoTo up1
    End If
    Case 6
    If ds = 0 Then
        tde = -de
        tds = 0
    Else
        i = 7
        GoTo up1
    End If
    Case 7
    If issearching = False Then
        If followWhite = False Then
                followWhite = True
                tde = de
                tds = ds
                i = 0
        ElseIf success = False Then
            t1 = pop(3)
            If isnomorevaluetopop = True Then
                endcx = corncx
                endcy = corncy
                isapproved = True
                ispopBack = True
                isendofexplore = True
                Exit Sub
            End If
            If t1 = -1 Then
                Exit Sub
            End If
            popbackindex = t1
            If ispopBack = False Then
                startcx = cx
                startcy = cy
            End If
            endcy = pop(1)
            endcx = pop(2)
            de = pop(4)
            ds = pop(5)
            cy = endcy
            cx = endcx
            isapproved = False
            ispopBack = True
            i = 0
            GoTo up0
        Else
            cx = tcx
            cy = tcy
            de = fde
            ds = fds
            Exit Sub
        End If
    Else
        Exit Sub
    End If
    End Select
    GoTo CheckCase
s3:
                
cx = startcx
cy = startcy
getshortestpath
traceback_init
GoTo up0

End Sub

Private Sub Timer2_Timer()
Command11_Click
End Sub





Public Function pop(id As Integer) As Integer
If ptr(id) = 0 Then
    MsgBox "No More value to pop !"
    pop = 0
    Timer2.Enabled = False
    isended = True
    ispopBack = False
    isnomorevaluetopop = True
    Exit Function
End If
    pop = list(id, ptr(id))
    ptr(id) = ptr(id) - 1
End Function

Public Sub changecolor(c1 As Long, c2 As Long)
G.Visible = False
    Dim i As Integer, j As Integer
    For i = 2 To G.Cols - 4
        For j = 2 To G.Rows - 4
            G.Col = i
            G.Row = j
            If G.CellBackColor = c1 Then
                G.CellBackColor = c2
            End If
        Next
    Next
G.Visible = True
End Sub

Public Sub getshortestpath()
Dim t1 As Integer
Dim tcx As Integer, tcy As Integer, tde As Integer, tds As Integer, tendcx As Integer, tendcy As Integer
tcx = cx
tcy = cy
tde = de
tds = ds

cx = endcx
cy = endcy
tendcx = endcx
tendcy = endcy
endcx = tcx
endcy = tcy


issearching = True
isended = False
traceback = False

G.TextMatrix(cy, cx) = 0
G.Col = cx
G.Row = cy
G.CellBackColor = vbGreen

ptr(6) = 0
ptr(7) = 0
ptr(8) = 0
ptr(9) = 0
ptr(10) = 0


getNextCell_rev

DoEvents
up1:
If ptr(8) > 0 And isended = False Then
            t1 = popinv(8) + 1
            cy = popinv(6)
            cx = popinv(7)
            ds = popinv(9)
            de = popinv(10)
            getNextCell_rev
            GoTo up1
End If

ptr(6) = 0
ptr(7) = 0
ptr(8) = 0
ptr(9) = 0
ptr(10) = 0


cx = tcx
cy = tcy
de = tde
ds = tds

endcx = tendcx
endcy = tendcy
End Sub



Public Sub listadd(id As Integer, value As Integer)
ptr(id) = ptr(id) + 1
list(id, ptr(id)) = value
End Sub

Public Function popinv(id As Integer)
Dim t1 As Integer
Dim j As Integer

t1 = list(id, 1)
For j = 0 To ptr(id) - 1
    list(id, j) = list(id, j + 1)
Next
popinv = t1
ptr(id) = ptr(id) - 1

End Function
