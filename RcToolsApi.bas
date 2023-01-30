Attribute VB_Name = "RcToolsApi"
Option Explicit

'Vari·veis
Public hWndLightBox As Long
Public lngLightBox As Long
Public MeuForm As Long
Public Style_Form As Long
Public hIcone As Long
Public Const SM_CXSCREEN = 0
Public Const SM_CYSCREEN = 1
Public Const LOG_PIXELS_X = 88
Public Const POINTS_PER_INCH As Long = 72
Public Const lngLightBoxColour As Long = vbButtonShadow
Public Const GWL_STYLE As Long = (-16)
Public Const GWL_EXSTYLE As Long = (-20)
Public Const ESTILO_ATUAL As Long = (-16)
Public Const WS_EX_APPWINDOW As Long = &H40000
Public Const WS_EX_LAYERED = &H80000
Public Const LWA_ALPHA = &H2
Public Const LWA_COLORKEY = &H1
Public Const WS_CAPTION As Long = &HC00000
Public Const WS_EX_DLGMODALFRAME As Long = &H1
Public Const WS_MINIMIZEBOX As Long = &H20000
Public Const WS_MAXIMIZEBOX As Long = &H10000

'// Contantes do Õçcone
Public Const FOCO_ICONE = &H80
Public Const ICONE = 0&
Public Const GRANDE_ICONE = 1&
Public Const IDC_HAND = 32649&


'API's
#If VBA7 Then
    Public Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    Public Declare PtrSafe Function MoveWindow Lib "user32.dll" (ByVal hwnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
    Public Declare PtrSafe Function GetWindowLong Lib "user32.dll" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
    Public Declare PtrSafe Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
    Public Declare PtrSafe Function SetLayeredWindowAttributes Lib "user32.dll" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
    Public Declare PtrSafe Function GetSystemMetrics Lib "user32.dll" (ByVal nIndex As Long) As Long
    Public Declare PtrSafe Function DrawMenuBar Lib "user32" (ByVal hwnd As Long) As Long
    Public Declare PtrSafe Function FindWindowA Lib "user32" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    Public Declare PtrSafe Function IconApp Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
    Public Declare PtrSafe Function LoadCursorBynum Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long
    Public Declare PtrSafe Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long
#Else
    Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    Public Declare Function MoveWindow Lib "user32.dll" (ByVal hwnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
    Public Declare Function GetWindowLong Lib "user32.dll" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
    Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
    Public Declare Function SetLayeredWindowAttributes Lib "user32.dll" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
    Public Declare Function GetSystemMetrics Lib "user32.dll" (ByVal nIndex As Long) As Long
    Public Declare Function DrawMenuBar Lib "user32" (ByVal hwnd As Long) As Long
    Public Declare Function FindWindowA Lib "user32" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    Public Declare Function IconApp Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
    Public Declare Function LoadCursorBynum Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long
    Public Declare Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long
#End If

Public coll As New Collection


Public Function MouseCursor(CursorType As Long)
  Dim lngRet As Long
  lngRet = LoadCursorBynum(0&, CursorType)
  lngRet = SetCursor(lngRet)
End Function

Public Function MouseMoveIcon()
    Call MouseCursor(IDC_HAND)
End Function

Function HideTitleBarAndBordar(frm As Object)
    Dim lngWindow As Long
    Dim lFrmHdl As Long
    lFrmHdl = FindWindow(vbNullString, frm.Caption)
    lngWindow = GetWindowLong(lFrmHdl, GWL_STYLE)
    lngWindow = lngWindow And (Not WS_CAPTION)
    SetWindowLong lFrmHdl, GWL_STYLE, lngWindow
    lngWindow = GetWindowLong(lFrmHdl, GWL_EXSTYLE)
    lngWindow = lngWindow And Not WS_EX_DLGMODALFRAME
    SetWindowLong lFrmHdl, GWL_EXSTYLE, lngWindow
    DrawMenuBar lFrmHdl
End Function

Function MakeUserformTransparent(frm As Object, Optional color As Variant)
    Dim FormHandle As Long
    Dim bytOpacity As Byte
    
    FormHandle = FindWindow(vbNullString, frm.Caption)
    If IsMissing(color) Then color = &H8000&        '//rgbWhite
    bytOpacity = 1
    
    SetWindowLong FormHandle, GWL_EXSTYLE, GetWindowLong(FormHandle, GWL_EXSTYLE) Or WS_EX_LAYERED
    
    frm.BackColor = color
    SetLayeredWindowAttributes FormHandle, color, bytOpacity, LWA_COLORKEY

End Function

Public Sub removeCaption(objForm As Object, Optional color As Variant)
    On Error Resume Next
    Dim lngMyHandle As Long, lngCurrentStyle As Long, lngNewStyle As Long
    Dim ico As MSForms.Image
    
    
    If objForm.Caption <> "" Then: objForm.Caption: Else objForm.Caption = " "
    
    If Val(Application.Version) < 9 Then
        lngMyHandle = FindWindowA("ThunderXFrame", objForm.Caption)
    Else
        lngMyHandle = FindWindowA("ThunderDFrame", objForm.Caption)
    End If

    MeuForm = FindWindowA(vbNullString, objForm.Caption)
    Style_Form = Style_Form Or &HCCCCC0 '&HDC47BE '&HDCCEBE
    SetWindowLong MeuForm, ESTILO_ATUAL, (Style_Form)
    
    lngCurrentStyle = GetWindowLong(lngMyHandle, GWL_STYLE)
    lngNewStyle = lngCurrentStyle Or WS_MINIMIZEBOX Or WS_MAXIMIZEBOX
    SetWindowLong lngMyHandle, GWL_STYLE, lngNewStyle
    
    Set ico = objForm.Controls.Add("Forms.Image.1", "ico", False)
    
    ico.Picture = LoadPicture(ThisWorkbook.Path & "/ico.ico")
    
    
    hIcone = objForm.ico.Picture.Handle
    Call IconApp(MeuForm, FOCO_ICONE, ICONE, ByVal hIcone)
    
End Sub

Public Function ScreenWidth() As Long
    On Error Resume Next
    ScreenWidth = GetSystemMetrics(SM_CXSCREEN)
End Function

Public Function ScreenHeight() As Long
    On Error Resume Next
    ScreenHeight = GetSystemMetrics(SM_CYSCREEN)
End Function

Public Sub PopupForm(FrmPopup As Object, FrmIndex As Object)
    On Error GoTo Catch

    hWndLightBox = FindWindowA("ThunderDFrame", FrmPopup.Caption)
    MoveWindow hWndLightBox, ScreenWidth, ScreenHeight, ScreenWidth, ScreenHeight, 1
    FrmPopup.BackColor = lngLightBoxColour
    lngLightBox = GetWindowLong(hWndLightBox, GWL_STYLE)
    lngLightBox = lngLightBox And Not &HC00000
    SetWindowLong hWndLightBox, GWL_STYLE, lngLightBox
    SetWindowLong hWndLightBox, GWL_EXSTYLE, WS_EX_APPWINDOW
    SetWindowLong hWndLightBox, GWL_EXSTYLE, WS_EX_LAYERED
    SetLayeredWindowAttributes hWndLightBox, 0, 80, LWA_ALPHA
    
    With FrmPopup
        .Width = FrmIndex.Width - 12
        .Height = FrmIndex.Height - 6
        .Top = FrmIndex.Top
        .Left = FrmIndex.Left + 6
    End With
    ' // exit
    Exit Sub

    ' // log error
Catch:
    Debug.Print "LightBox-UserForm_Initialize(): " & Err.Number & ", " & Err.Description

End Sub

Public Sub AlignChildFormOnCenter(Child As Object)
    With Child
        .Left = Popup.Left + (Popup.Width / 2) - (.Width / 2) + 24
        .Top = Popup.Top + (Popup.Height / 2) - (.Height / 2) + 24
    End With
End Sub

