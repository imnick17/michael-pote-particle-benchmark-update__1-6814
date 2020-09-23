Attribute VB_Name = "Module1"
Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Const SRCCOPY = &HCC0020 ' (DWORD) dest = source
Public Const SRCAND = &H8800C6  ' (DWORD) dest = source AND dest
Public Const SRCPAINT = &HEE0086    ' dest = source OR dest
Public Const SRCERASE = &H440328        ' (DWORD) dest = source AND (NOT dest )
Public Const SRCINVERT = &H660046       ' (DWORD) dest = source XOR dest
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Type OPENFILENAME
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
    End Type
    Public Const OFN_READONLY = &H1
    Public Const OFN_OVERWRITEPROMPT = &H2
    Public Const OFN_HIDEREADONLY = &H4
    Public Const OFN_NOCHANGEDIR = &H8
    Public Const OFN_SHOWHELP = &H10
    Public Const OFN_ENABLEHOOK = &H20
    Public Const OFN_ENABLETEMPLATE = &H40
    Public Const OFN_ENABLETEMPLATEHANDLE = &H80
    Public Const OFN_NOVALIDATE = &H100
    Public Const OFN_ALLOWMULTISELECT = &H200
    Public Const OFN_EXTENSIONDIFFERENT = &H400
    Public Const OFN_PATHMUSTEXIST = &H800
    Public Const OFN_FILEMUSTEXIST = &H1000
    Public Const OFN_CREATEPROMPT = &H2000
    Public Const OFN_SHAREAWARE = &H4000
    Public Const OFN_NOREADONLYRETURN = &H8000
    Public Const OFN_NOTESTFILECREATE = &H10000
    Public Const OFN_NONETWORKBUTTON = &H20000
    Public Const OFN_NOLONGNAMES = &H40000 ' force no long names for 4.x modules
    Public Const OFN_EXPLORER = &H80000 ' new look commdlg
    Public Const OFN_NODEREFERENCELINKS = &H100000
    Public Const OFN_LONGNAMES = &H200000 ' force long names for 3.x modules
    Public Const OFN_SHAREFALLTHROUGH = 2
    Public Const OFN_SHARENOWARN = 1
    Public Const OFN_SHAREWARN = 0


Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Public Type BrowseInfo
     hwndOwner As Long
     pIDLRoot As Long
     pszDisplayName As Long
     lpszTitle As Long
     ulFlags As Long
     lpfnCallback As Long
     lParam As Long
     iImage As Long
End Type

Public Const BIF_RETURNONLYFSDIRS = 1
Public Const MAX_PATH = 260

Public Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Long)
Public Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Public Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
Public Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long
Function OpenDialog(Form1 As Form, Filter As String, Title As String, InitDir As String) As String

    
    Dim ofn As OPENFILENAME
    Dim a As Long
    ofn.lStructSize = Len(ofn)
    ofn.hwndOwner = Form1.hwnd
    ofn.hInstance = App.hInstance
    If Right$(Filter, 1) <> "|" Then Filter = Filter + "|"


    For a = 1 To Len(Filter)
        If Mid$(Filter, a, 1) = "|" Then Mid$(Filter, a, 1) = Chr$(0)
    Next

    ofn.lpstrFilter = Filter
    ofn.lpstrFile = Space$(254)
    ofn.nMaxFile = 255
    ofn.lpstrFileTitle = Space$(254)
    ofn.nMaxFileTitle = 255
    ofn.lpstrInitialDir = InitDir
    ofn.lpstrTitle = Title
    ofn.flags = OFN_HIDEREADONLY Or OFN_FILEMUSTEXIST
    a = GetOpenFileName(ofn)


    If (a) Then
        OpenDialog = Trim$(ofn.lpstrFile)
    Else
        OpenDialog = ""
    End If

End Function

Sub Trans(Pic As PictureBox, picM As PictureBox, X, Y, destHDC, Size As Integer)
StretchBlt destHDC, X, Y, Pic.ScaleWidth - Size, Pic.ScaleHeight - Size, picM.hdc, 0, 0, Pic.ScaleWidth, Pic.ScaleHeight, SRCAND
StretchBlt destHDC, X, Y, Pic.ScaleWidth - Size, Pic.ScaleHeight - Size, Pic.hdc, 0, 0, Pic.ScaleWidth, Pic.ScaleHeight, SRCPAINT
End Sub
