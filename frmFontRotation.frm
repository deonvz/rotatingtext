VERSION 5.00
Begin VB.Form frmFontRotation 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000007&
   Caption         =   "Form1"
   ClientHeight    =   5235
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6735
   ForeColor       =   &H000000FF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   349
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   449
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmFontRotation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Rotating Text Sample
'Author: Deon van Zyl


Option Explicit
'API's used in this sample
Private Declare Function CreateFont Lib "gdi32" Alias "CreateFontA" (ByVal H As Long, ByVal W As Long, ByVal E As Long, ByVal O As Long, ByVal W As Long, ByVal I As Long, ByVal U As Long, ByVal S As Long, ByVal C As Long, ByVal OP As Long, ByVal CP As Long, ByVal Q As Long, ByVal PAF As Long, ByVal F As String) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetTickCount Lib "kernel32" () As Long

'Constant text to draw
Const TEXTOUTPUT As String = "Your text here"
Const PI As Single = 3.141593

'API constants
Const ANSI_CHARSET As Long = 0
Const FF_DONTCARE As Long = 0
Const CLIP_LH_ANGLES As Long = &H10
Const CLIP_DEFAULT_PRECIS As Long = 0
Const OUT_TT_ONLY_PRECIS As Long = 7
Const PROOF_QUALITY As Long = 2
Const TRUETYPE_FONTTYPE As Long = &H4
Const p_WIDTH As Long = 12
Const p_HEIGHT As Long = 12


'Center coordinates
Dim pXCenter As Long
Dim pYCenter As Long

'LookUp table with relative coordinates
Dim LookUp(1 To 2, 1 To 36) As Long
Dim pRadius As Long
'ending flag
Dim TimeToEnd As Boolean

'Main animation procedure
Private Sub RunMain()
Const FrameInterval As Long = 35
Dim LastFrameTime As Long
Dim Angle As Long

'Show the form
Me.Show

Angle = 1800
Do
    'check to see if we have to end
    If TimeToEnd Then Exit Do
    
        
        If GetTickCount() - LastFrameTime > FrameInterval Then  'Time to update
            
            'update angle
            Angle = (Angle Mod 3600) - 100
            'clear the form
            Me.Cls
            
            DrawRotatedText Angle
            
            LastFrameTime = GetTickCount()
                        
        End If
        
    DoEvents

Loop


End Sub
'Draws the rotated text
Private Sub DrawRotatedText(Angle As Long)
Dim NewFont As Long
Dim OldFont As Long
Static I As Long

'creat the font
NewFont = CreateFont(p_HEIGHT, p_WIDTH, Angle, 0, FF_DONTCARE, 0, 0, 0, ANSI_CHARSET, OUT_TT_ONLY_PRECIS, CLIP_LH_ANGLES Or CLIP_DEFAULT_PRECIS, PROOF_QUALITY, TRUETYPE_FONTTYPE, "Arial")

'set the new font
OldFont = SelectObject(Me.hdc, NewFont)

I = (I Mod 36) + 1

CurrentX = pXCenter + LookUp(1, I)
CurrentY = pYCenter + LookUp(2, I)

Print TEXTOUTPUT

'set the old font back
NewFont = SelectObject(Me.hdc, OldFont)

'Clean up
DeleteObject NewFont

End Sub

Private Sub Form_Load()

pRadius = ((Len(TEXTOUTPUT) * p_WIDTH) / 2)

BuildLookupTable
RunMain

End Sub

Private Sub Form_Resize()
'calculate center
pXCenter = Me.ScaleWidth / 2
pYCenter = Me.ScaleHeight / 2
End Sub

Private Sub Form_Unload(Cancel As Integer)
'flag the end
TimeToEnd = True
End Sub

'Builds the lookup table with the circle coordinates
Private Sub BuildLookupTable()
Dim I As Long
Dim Angle As Long
Const XIndex As Long = 1
Const YIndex As Long = 2

For I = LBound(LookUp, 2) To UBound(LookUp, 2)
    LookUp(XIndex, I) = CLng(Cos((Angle * PI / 180)) * pRadius)
    LookUp(YIndex, I) = CLng(Sin((Angle * PI / 180)) * pRadius)
    Angle = (Angle Mod 360) + 10
Next I

End Sub
