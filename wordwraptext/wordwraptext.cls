VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WordWrapText"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit
'****************************************************************************
' *  Type:   Class
' *  Name:   WordWrapText
' *  Author: shangdeming(http://www.blanksoft.com)
' *  Desc.:  Automatically  word wrap a single line text to multi-line  with fixed width
' *
' *
' *  Modifications
' *
' *    Name          Date      Reason
' * -----------    ---------  -------------------------------------------------
' * shangdeming    2004/04/07  NEW
' *
' ****************************************************************************
Private Type TEXTMETRIC
        tmHeight As Long
        tmAscent As Long
        tmDescent As Long
        tmInternalLeading As Long
        tmExternalLeading As Long
        tmAveCharWidth As Long
        tmMaxCharWidth As Long
        tmWeight As Long
        tmOverhang As Long
        tmDigitizedAspectX As Long
        tmDigitizedAspectY As Long
        tmFirstChar As Byte
        tmLastChar As Byte
        tmDefaultChar As Byte
        tmBreakChar As Byte
        tmItalic As Byte
        tmUnderlined As Byte
        tmStruckOut As Byte
        tmPitchAndFamily As Byte
        tmCharSet As Byte
End Type
'Declare API Call
Private Declare Function GetTextMetrics Lib "gdi32" Alias "GetTextMetricsA" (ByVal hdc As Long, lpMetrics As TEXTMETRIC) As Long
Private m_Dism As Collection
Private m_LineCount As Long
Private m_StringHeight As Long
Private m_strSource As String
Private m_Width As Long
Private m_isSplitWord As Boolean
Private m_hdc As Long
Private colLineContent As Collection


Public Function CalcucateIt() As Long
   Dim Textmetrics As TEXTMETRIC
   Dim i_strAveWidth As Long
   Dim i_strHeight As Long
   Dim i_Len As Long, i As Long
   Dim strtmp As String, strWord As String, strtmpLine As String
   Dim i_charCountLine As Long
   i_Len = Len(m_strSource)
   Set colLineContent = New Collection
   CalcucateIt = 0
   If GetTextMetrics(m_hdc, Textmetrics) <> 0 Then 'success
             i_strAveWidth = Textmetrics.tmAveCharWidth
             i_strHeight = Textmetrics.tmHeight
             i_charCountLine = Int(m_Width / i_strAveWidth)
    Else
        CalcucateIt = -1
        Exit Function
    End If
   
   If m_isSplitWord = False Then
        For i = 1 To i_Len
            strtmp = Mid(m_strSource, i, 1)
            If Len(strtmpLine) = i_charCountLine Then
                colLineContent.Add strtmpLine
                strtmpLine = strtmp
            Else
                strtmpLine = strtmpLine & strtmp
                If i = i_Len Then colLineContent.Add strtmpLine
            End If
        Next
   Else
        strtmpLine = ""
        strWord = ""
        For i = 1 To i_Len
           strtmp = Mid(m_strSource, i, 1)
            If Checkdismember(strtmp) = False Then
              strWord = strWord & strtmp
              If i = i_Len Then colLineContent.Add strWord
            Else
                strWord = strWord & strtmp
                
                If Len(strWord) < i_charCountLine - Len(strtmpLine) Then
                    
                    strtmpLine = strtmpLine & strWord
                    strWord = ""
                    If i = i_Len Then colLineContent.Add strtmpLine
                Else
                   colLineContent.Add strtmpLine
                   strtmpLine = strWord
                   strWord = ""
                End If
            End If
        
        Next
   End If
   m_LineCount = colLineContent.Count
   m_StringHeight = m_LineCount * i_strHeight
End Function

Public Function AddDismember(ByVal strDism As String)
    m_Dism.Add strDism
End Function

Private Sub Class_Initialize()
    Set m_Dism = New Collection
    initValue
End Sub

Public Function ClearDismember()
'    Dim i As Long
'    For i = 1 To m_Dism.Count
'      m_Dism.Remove (i)
'    Next
    Set m_Dism = Nothing
    Set m_Dism = New Collection
End Function

Private Sub Class_Terminate()
    Set m_Dism = Nothing
End Sub

Public Property Get LineCount() As Long
    LineCount = m_LineCount
End Property



Public Property Get StringHeight() As Long
    StringHeight = m_StringHeight
End Property

Public Property Get StringSource() As String
    StringSource = m_strSource
End Property

Public Property Let StringSource(ByVal strNewValue As String)
    m_strSource = strNewValue
End Property

Public Property Get Width() As Long
   Width = m_Width
End Property

Public Property Let Width(ByVal iNewValue As Long)
    If iNewValue <= 0 Then m_Width = 1
    m_Width = iNewValue
End Property

Public Property Get bSplitWord() As Boolean
   bSplitWord = m_isSplitWord
End Property

Public Property Let bSplitWord(ByVal bNewValue As Boolean)
    m_isSplitWord = bNewValue
End Property

Private Sub initValue()
 m_LineCount = 0
 m_StringHeight = 0
 m_strSource = ""
 m_Width = 0
 m_isSplitWord = False
 hdc = 0
End Sub
Private Function Checkdismember(str As String) As Boolean
    Dim i As Long
    Checkdismember = False
    For i = 1 To m_Dism.Count
        If str = m_Dism(i) Then
            Checkdismember = True
            Exit Function
        End If
    Next
End Function

Public Property Get hdc() As Long
    hdc = m_hdc
End Property

Public Property Let hdc(ByVal iNewValue As Long)
    m_hdc = iNewValue
End Property
Public Property Get LineContent() As Collection
    If Not IsNull(colLineContent) Then Set LineContent = colLineContent
End Property


