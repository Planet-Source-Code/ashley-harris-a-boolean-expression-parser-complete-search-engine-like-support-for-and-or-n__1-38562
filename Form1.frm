VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5250
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9135
   LinkTopic       =   "Form1"
   ScaleHeight     =   5250
   ScaleWidth      =   9135
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Height          =   345
      Left            =   1560
      TabIndex        =   2
      Top             =   4410
      Width           =   6285
   End
   Begin VB.TextBox Text1 
      Height          =   4065
      Left            =   90
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   285
      Width           =   8715
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      Height          =   315
      Left            =   1560
      TabIndex        =   4
      Top             =   4875
      Width           =   6330
   End
   Begin VB.Label Label2 
      Caption         =   "QUERY STRING:"
      Height          =   300
      Left            =   135
      TabIndex        =   3
      Top             =   4485
      Width           =   1305
   End
   Begin VB.Label Label1 
      Caption         =   "results:"
      Height          =   270
      Left            =   105
      TabIndex        =   1
      Top             =   15
      Width           =   585
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'See the SearchMod for more information, this is just a 5 minute implementation of it.

Dim fso As New FileSystemObject
Dim data(1 To 231) As String

Private Const recordcount As Long = 231

Private Sub Form_Load()
    Caption = "Boolean Search Example - VB6"
    Dim ts As TextStream
    Set ts = fso.OpenTextFile(fso.BuildPath(App.Path, "data.txt"))
    For a = 1 To recordcount
        data(a) = ts.ReadLine
    Next a
    ts.Close
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Dim t As Single
        Text1 = ""
        search
        Text2.SelStart = 0
        Text2.SelLength = Len(Text2) + 1
    End If
End Sub

Public Sub search()
    t = Timer
    For a = 1 To recordcount
        If checkstring(Text2, data(a)) Then
            b = b & data(a) & vbNewLine & vbNewLine
            c = c + 1
        End If
    Next a
    t = Timer - t
    Label3.Caption = c & " found. " & recordcount & " records searched in " & t & " seconds (" & Format(t / recordcount, ".000000") & " per record)"
    On Error GoTo out
    Text1 = b
    Exit Sub
out:
    Text1 = "Display Error: " & Err.Description
End Sub

