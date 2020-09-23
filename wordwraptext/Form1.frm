VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6930
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8445
   LinkTopic       =   "Form1"
   ScaleHeight     =   6930
   ScaleWidth      =   8445
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Height          =   4350
      Left            =   690
      TabIndex        =   2
      Top             =   1170
      Width           =   4125
   End
   Begin VB.TextBox Text1 
      Height          =   525
      Left            =   810
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   420
      Width           =   7395
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   315
      Left            =   660
      TabIndex        =   0
      Top             =   6420
      Width           =   1395
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    Dim m_cls As New WordWrapText
    Dim mLineContent As New Collection
    Dim i As Long
    m_cls.StringSource = Text1.Text
    m_cls.AddDismember " "
    m_cls.AddDismember ","
    m_cls.AddDismember ":"
    m_cls.AddDismember "."
    m_cls.AddDismember "?"
    m_cls.AddDismember "("
    m_cls.AddDismember ")"
    m_cls.bSplitWord = True
    m_cls.hdc = Me.hdc
    m_cls.Width = Text1.Width / 50
    
    
    Set List1.Font = Me.Font
    If m_cls.CalcucateIt = 0 Then
      Text1.Text = m_cls.LineCount & " " & m_cls.StringHeight
       Set mLineContent = m_cls.LineContent
        For i = 1 To m_cls.LineCount
           List1.AddItem mLineContent(i)
        Next
    End If
    m_cls.ClearDismember
    m_cls.bSplitWord = False
    If m_cls.CalcucateIt = 0 Then
      Text1.Text = m_cls.LineCount & " " & m_cls.StringHeight
       Set mLineContent = m_cls.LineContent
        For i = 1 To m_cls.LineCount
           List1.AddItem mLineContent(i)
        Next
    End If
    
    
    
   Set m_cls = Nothing
End Sub

