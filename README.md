<div align="center">

## CHAMELEON BUTTON 2\.0\.6 Memory Leak FIX


</div>

### Description

When using the chameleon button user control I noticed that eventually the application will begin to run out of resources. I was extremely curious to what was causing it, and decided to identify the problem. Using the vbAccelerator GUI Resource Tracer (www.vbaccelerator.com), I was able to find the cause of the problem. In the procedure DrawFrame(...) the memory leak occurs. I implemented DeleteObject calls for the hPen and hObject. Originally the code would only call DeleteObject on the value of hObject and not hPen. Hope anyone that uses the chameleon button user control finds this helpful. This user control is hosted on this website incase you are unfamiliar with it.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Cody Rutkowski](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/cody-rutkowski.md)
**Level**          |Advanced
**User Rating**    |4.8 (29 globes from 6 users)
**Compatibility**  |VB 6\.0
**Category**       |[Graphics](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/graphics__1-46.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/cody-rutkowski-chameleon-button-2-0-6-memory-leak-fix__1-70509/archive/master.zip)

### API Declarations

```
' Replace the existing procedure with the procedure provided below.
```


### Source Code

```
Private Sub DrawFrame(ByVal ColHigh As Long, ByVal ColDark As Long, ByVal ColLight As Long, ByVal ColShadow As Long, ByVal ExtraOffset As Boolean, Optional ByVal Flat As Boolean = False)
'a very fast way to draw windows-like frames
Dim pt As POINTAPI
Dim frHe As Long, frWi As Long, frXtra As Long
  frHe = He - 1 + ExtraOffset: frWi = Wi - 1 + ExtraOffset: frXtra = Abs(ExtraOffset)
  With UserControl
    Dim hObject As Long
    Dim hPen As Long
    '=============================
    hPen = CreatePen(PS_SOLID, 1, ColHigh)
    hObject = SelectObject(.hDC, hPen)
    MoveToEx .hDC, frXtra, frHe, pt
    LineTo .hDC, frXtra, frXtra
    LineTo .hDC, frWi, frXtra
    Call DeleteObject(hObject)
    Call DeleteObject(hPen)
    '=============================
    hPen = CreatePen(PS_SOLID, 1, ColDark)
    hObject = SelectObject(.hDC, hPen)
    LineTo .hDC, frWi, frHe
    LineTo .hDC, frXtra - 1, frHe
    MoveToEx .hDC, frXtra + 1, frHe - 1, pt
    Call DeleteObject(hObject)
    Call DeleteObject(hPen)
    If Flat Then Exit Sub
    '=============================
    hPen = CreatePen(PS_SOLID, 1, ColLight)
    hObject = SelectObject(.hDC, hPen)
    LineTo .hDC, frXtra + 1, frXtra + 1
    LineTo .hDC, frWi - 1, frXtra + 1
    Call DeleteObject(hObject)
    Call DeleteObject(hPen)
    '=============================
    hPen = CreatePen(PS_SOLID, 1, ColShadow)
    hObject = SelectObject(.hDC, hPen)
    LineTo .hDC, frWi - 1, frHe - 1
    LineTo .hDC, frXtra, frHe - 1
    Call DeleteObject(hObject)
    Call DeleteObject(hPen)
  End With
End Sub
```

