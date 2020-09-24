VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   Caption         =   "Polygon Area Calculator"
   ClientHeight    =   8040
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   12165
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   14.182
   ScaleMode       =   7  'Centimeter
   ScaleWidth      =   21.458
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4815
      Left            =   7320
      ScaleHeight     =   4785
      ScaleWidth      =   4185
      TabIndex        =   3
      Top             =   2400
      Visible         =   0   'False
      Width           =   4215
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4815
      Left            =   5280
      ScaleHeight     =   4785
      ScaleWidth      =   4185
      TabIndex        =   2
      Top             =   1320
      Visible         =   0   'False
      Width           =   4215
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   2160
      Top             =   1680
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DialogTitle     =   "Open"
      Filter          =   "Text File (*.txt)|*.txt"
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   495
      Left            =   0
      TabIndex        =   1
      Top             =   480
      Visible         =   0   'False
      Width           =   12135
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   12135
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      Visible         =   0   'False
      X1              =   19.05
      X2              =   17.78
      Y1              =   1.27
      Y2              =   5.08
   End
   Begin VB.Menu MnuFile 
      Caption         =   "File"
      Begin VB.Menu MnuNew 
         Caption         =   "New"
         Shortcut        =   {F2}
      End
      Begin VB.Menu MnuOpen 
         Caption         =   "Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu MnuSave 
         Caption         =   "Save As"
         Shortcut        =   ^S
      End
      Begin VB.Menu Mnusep 
         Caption         =   "-"
      End
      Begin VB.Menu MnuImage 
         Caption         =   "Open Image File"
         Shortcut        =   ^I
      End
      Begin VB.Menu MnuRemove 
         Caption         =   "Remove Image"
         Shortcut        =   ^R
      End
      Begin VB.Menu MuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu MnuSnap 
         Caption         =   "Snap to Grid"
         Shortcut        =   ^G
      End
   End
   Begin VB.Menu MnuOptions 
      Caption         =   "Options"
      Begin VB.Menu MnuColor 
         Caption         =   "Drawing Color"
      End
      Begin VB.Menu MnuGrid 
         Caption         =   "Grid Size"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PolyGonVertices() As Single
Dim StartDrawing  As Boolean, VertixNum As Integer, VertixCount As Integer
Dim ImageLoaded As Boolean, GridStep As Single, PolySaved As Boolean
Private Sub Form_Activate()



DrawGrid

Label1.Left = 0
Label1.Width = Form1.Width
Label2.Left = 0
Label2.Width = Form1.Width






End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)






If KeyAscii = 27 And StartDrawing Then
    MnuNew_Click
  
End If











End Sub


Private Sub Form_Load()


GridStep = 1





End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)


If VertixCount <> 0 Then Exit Sub


If StartDrawing = False Then StartDrawing = True


If MnuSnap.Checked Then x = SnapToGrid(x): y = SnapToGrid(y)
PolySaved = False

VertixNum = VertixNum + 1
If VertixNum > 1 Then
  Line (x, y)-(Line1.X1, Line1.Y1)
End If

Line1.X1 = x
Line1.Y1 = y
Line1.X2 = x
Line1.Y2 = y
Line1.Visible = True


    ReDim Preserve PolyGonVertices(1 To 2, 1 To VertixNum)
    PolyGonVertices(1, VertixNum) = x
    PolyGonVertices(2, VertixNum) = y

If VertixNum > 1 Then
   If PolyGonVertices(1, VertixNum - 1) = x And PolyGonVertices(2, VertixNum - 1) = y Then
          ReDim Preserve PolyGonVertices(1 To 2, 1 To VertixNum - 1)
          VertixNum = VertixNum - 1
   End If

End If

If Button = 2 Then

  If VertixNum < 3 Then
    MnuNew_Click
    Exit Sub
  End If

  VertixCount = VertixNum
  VertixNum = 0

  StartDrawing = False
  Line (PolyGonVertices(1, 1), PolyGonVertices(2, 1))-(Line1.X2, Line1.Y2)
 
  ReDim Poly(1 To VertixCount) As Integer
  Line1.Visible = False
  For m% = 1 To VertixCount
    Poly(m) = m
  Next m

  
area = GetPolyArea(Poly, 0)

End If



End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)



If MnuSnap.Checked Then x = SnapToGrid(x): y = SnapToGrid(y)
Form1.Caption = "Polygon Area Calculator ( " + CStr(x) + " , " + CStr(y) + " )"


If StartDrawing Then
 Line1.X2 = x
 Line1.Y2 = y
End If









End Sub



Public Function GetPolyArea(Poly() As Integer, Level As Integer) As Single

Level = Level + 1

RemoveExtraPoints Poly
If UBound(Poly) < 3 Then area! = 0: GoTo zxc


ReDim tempPoly(1 To UBound(Poly)) As Integer
tempPoly = Poly
Dim PolyOnSide() As Integer


If UBound(tempPoly) = 3 Then
   area! = GetTriangleArea(tempPoly(1), tempPoly(2), tempPoly(3))
   GoTo zxc
End If



CreateConvexPoly tempPoly
area! = GetAreaPerfect(tempPoly)

For h% = 1 To UBound(tempPoly)

  a% = h
  If a = UBound(tempPoly) Then
     b% = 1
  Else
     b = h + 1
  End If

  y% = 0
uio:
  y = y + 1
    If Poly(y) = tempPoly(a) Then GoTo hjk
  GoTo uio
hjk:

    ReDim PolyOnSide(1 To 1)
    PolyOnSide(1) = tempPoly(a)
ert:
    y = y + 1
    If y > UBound(Poly) Then y = 1
    ReDim Preserve PolyOnSide(1 To UBound(PolyOnSide) + 1)
    PolyOnSide(UBound(PolyOnSide)) = Poly(y)
 
 
   If Poly(y) <> tempPoly(b) Then GoTo ert

   SortArray PolyOnSide



   If UBound(PolyOnSide) > 2 Then
       area = area - GetPolyArea(PolyOnSide, Level)
   End If

 
Next h

zxc:
Level = Level - 1
GetPolyArea = area

If Level = 0 Then
   Label1.Visible = True
   Label1.Caption = "The area of this Polygon using the recursive method = " + CStr(GetPolyArea) + " area unit"

   area = 0
   For x% = 1 To UBound(Poly) - 1
       area = area + PolyGonVertices(1, Poly(x + 1)) * PolyGonVertices(2, Poly(x)) - PolyGonVertices(1, Poly(x)) * PolyGonVertices(2, Poly(x + 1))
   Next x
   area = Abs((area + PolyGonVertices(1, Poly(1)) * PolyGonVertices(2, Poly(x)) - PolyGonVertices(1, Poly(x)) * PolyGonVertices(2, Poly(1))) / 2)
   Label2.Visible = True
   Label2.Caption = "The area of this Polygon using the trapezoidal method = " + CStr(area) + " area unit"
End If






End Function

Public Sub CreateConvexPoly(OldPoly() As Integer)

' This subroutine forms the convex polygon which consists of some _
  vertices of the input polygon




ReDim NewPoly(1 To UBound(OldPoly)) As Integer
NewPoly = OldPoly           ' get a copy of the input polygon
ReDim OldPoly(1 To 1) As Integer

' Get the points of highest and lowest value of the Y component                                  '*
                                                                                                 '*
Dim TopX As Integer, BottomX As Integer                                                          '*
                                                                                                 '*
TopX = NewPoly(1)                                                                                '*
BottomX = NewPoly(1)                                                                             '*
                                                                                                 '*
                                                                                                 '*
For x% = 2 To UBound(NewPoly)                                                                    '*
   If PolyGonVertices(2, NewPoly(x)) <= PolyGonVertices(2, TopX) Then TopX = NewPoly(x)          '*
   If PolyGonVertices(2, NewPoly(x)) >= PolyGonVertices(2, BottomX) Then BottomX = NewPoly(x)    '*
Next x                                                                                           '*



'***************From Top to right direction ****************************************
tempTopXLeft% = TopX
tempTopXRight% = TopX
R% = TopX
L% = TopX
OldPoly(1) = TopX

aaa:
angright! = 90
For x = 1 To UBound(NewPoly)
  If R <> NewPoly(x) Then
      If PolyGonVertices(1, NewPoly(x)) > PolyGonVertices(1, R) And PolyGonVertices(2, NewPoly(x)) = PolyGonVertices(2, R) Then
        tempTopXRight% = NewPoly(x)
        Changed = True
        Exit For
      End If
      angTemp! = GetAngle(NewPoly(x), R)
      If (angTemp <= angright And angTemp > 0 And PolyGonVertices(2, NewPoly(x)) > PolyGonVertices(2, R)) Then
         Changed = True
         angright = angTemp
         tempTopXRight% = NewPoly(x)
      End If
  End If
Next x

If Changed Then
  Changed = False
  R = tempTopXRight%
  AddToLastPoly OldPoly, R
  GoTo aaa
End If

'***************From Top to left direction ****************************************

cvb:
angleft! = -90
For x = 1 To UBound(NewPoly)
  If L <> NewPoly(x) Then
      If PolyGonVertices(1, NewPoly(x)) < PolyGonVertices(1, L) And PolyGonVertices(2, NewPoly(x)) = PolyGonVertices(2, L) Then
        tempTopXLeft% = NewPoly(x)
        Changed = True
        Exit For
      End If
      angTemp! = GetAngle(NewPoly(x), L)
      If (angTemp >= angleft And angTemp < 0 And PolyGonVertices(2, NewPoly(x)) > PolyGonVertices(2, L)) Or (angleft! = -90 And angTemp = 90 And PolyGonVertices(2, NewPoly(x)) > PolyGonVertices(2, L) And IsASide(NewPoly, L, NewPoly(x))) Then
         Changed = True
         angleft = angTemp
         If angTemp = 90 Then angleft = -90
         tempTopXLeft% = NewPoly(x)
      End If
  End If
Next x

If Changed Then
  Changed = False
  L = tempTopXLeft%
  AddToLastPoly OldPoly, L
  GoTo cvb
End If



'***************************From Bottom to right direction *****************************************

tempTopXLeft% = BottomX
tempTopXRight% = BottomX
R% = BottomX
L% = BottomX
AddToLastPoly OldPoly, BottomX

bbb:
angright! = -90
For x = 1 To UBound(NewPoly)
  If R <> NewPoly(x) Then
      If PolyGonVertices(1, NewPoly(x)) > PolyGonVertices(1, R) And PolyGonVertices(2, NewPoly(x)) = PolyGonVertices(2, R) Then
        tempTopXRight% = NewPoly(x)
        Changed = True
        Exit For
      End If
      angTemp! = GetAngle(NewPoly(x), R)
      If (angTemp >= angright And angTemp < 0 And PolyGonVertices(2, NewPoly(x)) < PolyGonVertices(2, R)) Or (angright! = -90 And angTemp = 90 And PolyGonVertices(2, NewPoly(x)) < PolyGonVertices(2, R) And IsASide(NewPoly, R, NewPoly(x))) Then
         Changed = True
         angright = angTemp
         If angTemp = 90 Then angright = -90
         tempTopXRight% = NewPoly(x)
      End If
  End If
Next x

If Changed Then
  Changed = False
  R = tempTopXRight%
  AddToLastPoly OldPoly, R
  GoTo bbb
End If

'***************************From Bottom to left direction *****************************************
azs:
angleft! = 90
For x = 1 To UBound(NewPoly)
  If L <> NewPoly(x) Then
      If PolyGonVertices(1, NewPoly(x)) < PolyGonVertices(1, L) And PolyGonVertices(2, NewPoly(x)) = PolyGonVertices(2, L) Then
        tempTopXLeft% = NewPoly(x)
        Changed = True
        Exit For
      End If
      angTemp! = GetAngle(NewPoly(x), L)
      If angTemp <= angleft And angTemp > 0 And PolyGonVertices(2, NewPoly(x)) < PolyGonVertices(2, L) Then
         Changed = True
         angleft = angTemp
         tempTopXLeft% = NewPoly(x)
      End If
  End If
Next x

If Changed Then
  Changed = False
  L = tempTopXLeft%
  AddToLastPoly OldPoly, L
  GoTo azs
End If

'"""""""""""""""""Sort Last Poly ******************


SortArray OldPoly




End Sub

Public Function GetAreaPerfect(Poly() As Integer)



For x% = 2 To UBound(Poly) - 1
  GetAreaPerfect = GetAreaPerfect + GetTriangleArea(Poly(1), Poly(x), Poly(x + 1))
Next x



End Function

Public Function GetTriangleArea(P1 As Integer, P2 As Integer, P3 As Integer) As Single


a! = GetLineLength(P1, P2)
b! = GetLineLength(P3, P2)
C! = GetLineLength(P1, P3)
s! = (a + b + C) / 2



GetTriangleArea = Sqr(s * (s - a) * (s - b) * (s - C))



End Function

Public Function GetLineLength(P1 As Integer, P2 As Integer) As Single




GetLineLength = Sqr((PolyGonVertices(1, P1) - PolyGonVertices(1, P2)) ^ 2 + (PolyGonVertices(2, P1) - PolyGonVertices(2, P2)) ^ 2)





End Function

Public Sub AddToLastPoly(Poly() As Integer, P1 As Integer)


'This code add the value of P1 to the array Poly if it doesn't exist.



For x% = 1 To UBound(Poly)
  If Poly(x) = P1 Then
     GoTo www
  End If
Next x


If x = UBound(Poly) + 1 Then
      ReDim Preserve Poly(1 To UBound(Poly) + 1)
      Poly(UBound(Poly)) = P1
End If


www:
End Sub

Public Function GetAngle(P1 As Integer, P2 As Integer) As Single


If PolyGonVertices(1, P1) = PolyGonVertices(1, P2) Then GetAngle = 90: Exit Function

GetAngle! = 180 * 7 / 22 * Atn((PolyGonVertices(2, P1) - PolyGonVertices(2, P2)) _
                                / (PolyGonVertices(1, P1) - PolyGonVertices(1, P2)))




End Function

Public Sub SortArray(Poly() As Integer)



fff:

For x% = 1 To UBound(Poly) - 1
    If Poly(x) > Poly(x + 1) Then
      a% = Poly(x)
      Poly(x) = Poly(x + 1)
      Poly(x + 1) = a
      GoTo fff
      
    End If


Next x








End Sub

Private Sub Form_Resize()

Form_Activate
DrawGrid








End Sub

Private Sub Form_Unload(Cancel As Integer)



Cancel = Exiting("exit")








End Sub

Private Sub MnuColor_Click()

On Error GoTo sdf

CD1.ShowColor

Form1.ForeColor = CD1.Color

If VertixCount > 0 Then DrawPolygon
If VertixNum > 0 Then DrawPath


Exit Sub


sdf:

If Err.Number <> 32755 Then
  MsgBox Err.Description, vbCritical + vbOKOnly, "Error"
  MnuNew_Click
End If




End Sub

Private Sub MnuGrid_Click()



temp = CStr(GridStep)


asdf:
temp = InputBox("Enter the grid size value ." + vbCrLf + vbCrLf + _
                "This value must range from 0.1 to 1.0 with step 0.1 ." + vbCrLf + vbCrLf + _
                "Empty value will affect nothing .", "Grid size", temp)


If temp = "" Then Exit Sub

If Len(temp) <> Len(Trim(temp)) Then GoTo asdf
If Not IsNumeric(temp) Then GoTo asdf
If CSng(temp) < 0.1 Or CSng(temp) > 1 Then GoTo asdf
If Len(CStr(CSng(temp))) <> 3 and CSng(temp) <> 1 Then GoTo asdf



GridStep = CSng(temp)


Form1.Cls

DrawGrid


If ImageLoaded = True Then
    ImagePaint
End If



If VertixCount > 0 Then DrawPolygon
 

  
If VertixNum > 0 Then DrawPath





End Sub

Private Sub MnuImage_Click()





On Error GoTo cvb




CD1.InitDir = App.Path
CD1.FileName = ""
CD1.Filter = "(*.bmp)|*.bmp|(*.jpg)|*.jpg|(*.gif)|*.gif|All Files|*.*"
CD1.FilterIndex = 4
CD1.ShowOpen

Picture1.Picture = LoadPicture()

Form1.ScaleMode = 1
Picture1.Picture = LoadPicture(CD1.FileName)
If Picture1.Width > Form1.Width * 0.9 Then MsgBox "Image is too wide to open .  ", vbCritical + vbOKOnly, "Error": Exit Sub
If Picture1.Height > Form1.Height * 0.9 Then MsgBox "Image is too tall to open .  ", vbCritical + vbOKOnly, "Error": Exit Sub
Form1.Cls
Picture2.Picture = Picture1.Picture
ImageLoaded = True
MnuSnap.Checked = False


ImagePaint
DrawGrid
If VertixCount > 0 Then DrawPolygon
If VertixNum > 0 Then DrawPath


  
Exit Sub



cvb:


If Err.Number <> 32755 Then
  MsgBox Err.Description, vbCritical + vbOKOnly, "Error"
  MnuNew_Click
End If

















End Sub

Private Sub MnuNew_Click()



If Exiting("draw a new polygon") = 1 Then Exit Sub



    VertixNum = 0
    VertixCount = 0
    Cls




If ImageLoaded = True Then
  ImagePaint
End If

    PolySaved = True
    StartDrawing = False
    DrawGrid
    Erase PolyGonVertices
    Line1.Visible = False
    Label1.Visible = False
    Label2.Visible = False







End Sub

Private Sub MnuOpen_Click()


On Error GoTo wer

If Exiting("open a saved file") = 1 Then Exit Sub



CD1.InitDir = App.Path
CD1.FileName = ""
CD1.Filter = "(Text File)|*.txt"
CD1.ShowOpen

Open CD1.FileName For Input As #1
ReDim PolyGonVertices(1 To 2, 1 To 1)
Line1.Visible = False



y% = 1
Do
    If y = 1 Then
        Input #1, PolyGonVertices(1, y), PolyGonVertices(2, y)
    Else
        Input #1, PolyGonVertices(1, y), PolyGonVertices(2, y)
        PolyGonVertices(1, y) = PolyGonVertices(1, y) + PolyGonVertices(1, 1)
        PolyGonVertices(2, y) = PolyGonVertices(2, y) + PolyGonVertices(2, 1)
        If PolyGonVertices(2, y) = PolyGonVertices(2, y - 1) And PolyGonVertices(1, y) = PolyGonVertices(1, y - 1) Then
          GoTo ghj
        End If
        
    End If
    
    ReDim Preserve PolyGonVertices(1 To 2, 1 To y + 1)
    y = y + 1
ghj:
Loop Until EOF(1)
ReDim Preserve PolyGonVertices(1 To 2, 1 To y - 1)
y = y - 1
If PolyGonVertices(2, y) = PolyGonVertices(2, 1) And PolyGonVertices(1, y) = PolyGonVertices(1, 1) Then
    ReDim Preserve PolyGonVertices(1 To 2, 1 To y - 1)
    y = y - 1
End If
  


Close #1

PolySaved = True

Cls


If ImageLoaded = True Then
 ImagePaint
End If


DrawGrid


VertixCount = y
  
DrawPolygon
  
  
   ReDim Poly(1 To VertixCount) As Integer
  
  For m% = 1 To VertixCount
    Poly(m) = m
  Next m
 
  area = GetPolyArea(Poly, 0)
   
  
Exit Sub



wer:
Close #1
Line1.Visible = False
  
If Err.Number <> 32755 Then
  MsgBox Err.Description, vbCritical + vbOKOnly, "Error"
  MnuNew_Click
End If




End Sub



Public Sub RemoveExtraPoints(Poly() As Integer)

ReDim tempPoly(1 To UBound(Poly)) As Integer

tempPoly = Poly

For x% = 1 To UBound(Poly) - 2

  If RemoveMidPoint(tempPoly(x), tempPoly(x + 1), tempPoly(x + 2)) = 1 Then
    Poly(x + 1) = 0
  End If
Next x

If RemoveMidPoint(tempPoly(UBound(Poly) - 1), tempPoly(UBound(Poly)), tempPoly(1)) = 1 Then Poly(UBound(Poly)) = 0
If RemoveMidPoint(tempPoly(UBound(Poly)), tempPoly(1), tempPoly(2)) = 1 Then Poly(1) = 0




ReDim NewPoly(1 To 1) As Integer



For x% = 1 To UBound(Poly)

  If Poly(x) <> 0 Then
    NewPoly(UBound(NewPoly)) = Poly(x)
    ReDim Preserve NewPoly(1 To UBound(NewPoly) + 1)
  End If
Next x

If UBound(NewPoly) > 1 Then
   ReDim Preserve NewPoly(1 To UBound(NewPoly) - 1)
End If

ReDim Poly(1 To UBound(NewPoly))

Poly = NewPoly


End Sub

Public Function RemoveMidPoint(P1 As Integer, P2 As Integer, P3 As Integer) As Integer



RemoveMidPoint = 0

If (PolyGonVertices(2, P1) - PolyGonVertices(2, P3)) = 0 Then
   If (PolyGonVertices(2, P2) - PolyGonVertices(2, P3)) = 0 Then
     RemoveMidPoint = 1
     Exit Function
   End If
Else
     If (PolyGonVertices(2, P2) - PolyGonVertices(2, P3)) = 0 Then
        Exit Function
     Else
        If (PolyGonVertices(1, P1) - PolyGonVertices(1, P3)) / (PolyGonVertices(2, P1) - PolyGonVertices(2, P3)) = _
           (PolyGonVertices(1, P2) - PolyGonVertices(1, P3)) / (PolyGonVertices(2, P2) - PolyGonVertices(2, P3)) Then RemoveMidPoint = 1
     End If
End If

   
   
  



End Function

Public Sub DrawGrid()



Form1.ScaleMode = 7
TempColor = ForeColor
ForeColor = vbWhite
DrawWidth = 1
For x = 0 To ScaleWidth + GridStep Step GridStep
    For y = 0 To ScaleHeight + GridStep Step GridStep
        PSet (x, y)
    Next y
Next x
DrawWidth = 1
ForeColor = TempColor


End Sub

Private Sub MnuRemove_Click()


ImageLoaded = False

Cls

DrawGrid
If VertixCount > 0 Then DrawPolygon
If VertixNum > 0 Then DrawPath









End Sub

Private Sub MnuSave_Click()

On Error GoTo hjkl


If Not (StartDrawing = False And VertixCount > 2) Then MsgBox "Nothing to save  .    ", vbCritical + vbOKOnly, "Polygon Area Calculator": Exit Sub


CD1.FileName = ""
CD1.Filter = "(Text File)|*.txt"
CD1.InitDir = App.Path
CD1.ShowSave

Open CD1.FileName For Output As #1

Print #1, CStr(PolyGonVertices(1, 1)) + "," + CStr(PolyGonVertices(2, 1))
For x% = 2 To VertixCount

  Print #1, CStr(PolyGonVertices(1, x) - PolyGonVertices(1, 1)) + "," + CStr(PolyGonVertices(2, x) - PolyGonVertices(2, 1))
   
Next x



Close #1


PolySaved = True

Exit Sub




hjkl:
Close #1

If Err.Number <> 32755 Then
  MsgBox Err.Description, vbCritical + vbOKOnly, "Error"
  MnuNew_Click
End If



End Sub


Private Sub MnuSnap_Click()


MnuSnap.Checked = Not (MnuSnap.Checked)











End Sub



Public Function IsASide(Poly() As Integer, P1 As Integer, P2 As Integer) As Boolean


For x% = 1 To UBound(Poly) - 1
 If (Poly(x) = P1 And Poly(x + 1) = P2) Or (Poly(x) = P1 And Poly(x + 1) = P2) Then IsASide = True: Exit Function
 

 
Next x

 If (Poly(1) = P1 And Poly(UBound(Poly)) = P2) Or (Poly(UBound(Poly)) = P1 And Poly(1) = P2) Then IsASide = True



End Function

Public Sub DrawPath()



 For x% = 1 To VertixNum - 1
     Line (PolyGonVertices(1, x), PolyGonVertices(2, x))-(PolyGonVertices(1, x + 1), PolyGonVertices(2, x + 1))
 Next x
 

End Sub

Public Sub DrawPolygon()



 For x% = 1 To VertixCount - 1
     Line (PolyGonVertices(1, x), PolyGonVertices(2, x))-(PolyGonVertices(1, x + 1), PolyGonVertices(2, x + 1))
 Next x
 Line (PolyGonVertices(1, x), PolyGonVertices(2, x))-(PolyGonVertices(1, 1), PolyGonVertices(2, 1))




End Sub


Public Function Exiting(Action As String) As Integer



temp = UCase(Left(Action, 1)) + Right(Action, Len(Action) - 1)

If VertixNum > 0 Then
   ddd% = MsgBox("If you don't need to complete the current polygon, press OK to " + Action + " otherwise press Cancel then go to complete it .", vbQuestion + vbOKCancel + vbDefaultButton2, temp)

   If ddd = vbCancel Then Exiting = 1
   Exit Function
End If






If VertixCount > 0 And PolySaved = False Then
   ddd% = MsgBox("If you don't need to save the drawn polygon, press OK to " + Action + " otherwise press Cancel then go to save it .", vbQuestion + vbOKCancel + vbDefaultButton2, temp)

   If ddd = vbCancel Then Exiting = 1
End If







End Function

Public Function SnapToGrid(P As Single) As Single


temp! = P / GridStep


If temp - Int(temp) >= 0.5 Then
   P = (Int(temp) + 1) * GridStep
Else
   P = Int(temp) * GridStep
End If



SnapToGrid = Round(P, 1)

End Function

Public Sub ImagePaint()


   Form1.ScaleMode = 1
   Form1.PaintPicture Picture2.Picture, -(Picture2.Width - Form1.Width) / 2, -(Picture2.Height - Form1.Height) / 2
   Form1.ScaleMode = 7





End Sub
