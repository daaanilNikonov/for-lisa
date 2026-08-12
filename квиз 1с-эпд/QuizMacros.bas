Attribute VB_Name = "QuizMacros"
Option Explicit

' Импорт: PowerPoint → Вид → Макросы → редактор VBA → File → Import
' Затем сохраните презентацию как .pptm и назначьте макросы на ячейки
' (или используйте pptx: там номера открытых вопросов краснеют без макросов).

Private Sub PaintRed(oShp As Shape)
    On Error Resume Next
    oShp.Fill.Visible = msoTrue
    oShp.Fill.Solid
    oShp.Fill.ForeColor.RGB = RGB(192, 57, 43)  ' #C0392B
    oShp.TextFrame.TextRange.Font.Color.RGB = RGB(255, 255, 255)
End Sub

Sub BackToBoard()
    SlideShowWindows(1).View.GotoSlide 2
End Sub

Sub OpenCell_0_0(oShp As Shape)
    PaintRed oShp
    SlideShowWindows(1).View.GotoSlide 3
End Sub

Sub OpenCell_0_1(oShp As Shape)
    PaintRed oShp
    SlideShowWindows(1).View.GotoSlide 4
End Sub

Sub OpenCell_0_2(oShp As Shape)
    PaintRed oShp
    SlideShowWindows(1).View.GotoSlide 5
End Sub

Sub OpenCell_0_3(oShp As Shape)
    PaintRed oShp
    SlideShowWindows(1).View.GotoSlide 6
End Sub

Sub OpenCell_0_4(oShp As Shape)
    PaintRed oShp
    SlideShowWindows(1).View.GotoSlide 7
End Sub

Sub OpenCell_1_0(oShp As Shape)
    PaintRed oShp
    SlideShowWindows(1).View.GotoSlide 8
End Sub

Sub OpenCell_1_1(oShp As Shape)
    PaintRed oShp
    SlideShowWindows(1).View.GotoSlide 9
End Sub

Sub OpenCell_1_2(oShp As Shape)
    PaintRed oShp
    SlideShowWindows(1).View.GotoSlide 10
End Sub

Sub OpenCell_1_3(oShp As Shape)
    PaintRed oShp
    SlideShowWindows(1).View.GotoSlide 11
End Sub

Sub OpenCell_1_4(oShp As Shape)
    PaintRed oShp
    SlideShowWindows(1).View.GotoSlide 12
End Sub

Sub OpenCell_2_0(oShp As Shape)
    PaintRed oShp
    SlideShowWindows(1).View.GotoSlide 13
End Sub

Sub OpenCell_2_1(oShp As Shape)
    PaintRed oShp
    SlideShowWindows(1).View.GotoSlide 14
End Sub

Sub OpenCell_2_2(oShp As Shape)
    PaintRed oShp
    SlideShowWindows(1).View.GotoSlide 15
End Sub

Sub OpenCell_2_3(oShp As Shape)
    PaintRed oShp
    SlideShowWindows(1).View.GotoSlide 16
End Sub

Sub OpenCell_2_4(oShp As Shape)
    PaintRed oShp
    SlideShowWindows(1).View.GotoSlide 17
End Sub

Sub OpenCell_3_0(oShp As Shape)
    PaintRed oShp
    SlideShowWindows(1).View.GotoSlide 18
End Sub

Sub OpenCell_3_1(oShp As Shape)
    PaintRed oShp
    SlideShowWindows(1).View.GotoSlide 19
End Sub

Sub OpenCell_3_2(oShp As Shape)
    PaintRed oShp
    SlideShowWindows(1).View.GotoSlide 20
End Sub

Sub OpenCell_3_3(oShp As Shape)
    PaintRed oShp
    SlideShowWindows(1).View.GotoSlide 21
End Sub

Sub OpenCell_3_4(oShp As Shape)
    PaintRed oShp
    SlideShowWindows(1).View.GotoSlide 22
End Sub

Sub OpenCell_4_0(oShp As Shape)
    PaintRed oShp
    SlideShowWindows(1).View.GotoSlide 23
End Sub

Sub OpenCell_4_1(oShp As Shape)
    PaintRed oShp
    SlideShowWindows(1).View.GotoSlide 24
End Sub

Sub OpenCell_4_2(oShp As Shape)
    PaintRed oShp
    SlideShowWindows(1).View.GotoSlide 25
End Sub

Sub OpenCell_4_3(oShp As Shape)
    PaintRed oShp
    SlideShowWindows(1).View.GotoSlide 26
End Sub

Sub OpenCell_4_4(oShp As Shape)
    PaintRed oShp
    SlideShowWindows(1).View.GotoSlide 27
End Sub

