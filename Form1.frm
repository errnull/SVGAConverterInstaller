VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H8000000E&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7230
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4515
   LinkTopic       =   "Form1"
   ScaleHeight     =   7230
   ScaleWidth      =   4515
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.Image Image3 
      Height          =   345
      Left            =   4080
      Picture         =   "Form1.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Height          =   615
      Left            =   1320
      TabIndex        =   0
      Top             =   6000
      Width           =   1815
   End
   Begin VB.Image Image1 
      Height          =   345
      Left            =   4080
      Picture         =   "Form1.frx":0D99
      Stretch         =   -1  'True
      Top             =   0
      Width           =   420
   End
   Begin VB.Image Image2 
      Height          =   7470
      Left            =   0
      Picture         =   "Form1.frx":1B76
      Top             =   -240
      Width           =   4530
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Const WS_EX_LAYERED = &H80000
Private Const GWL_EXSTYLE = (-20)
Private Const LWA_ALPHA = &H2
Private Const LWA_COLORKEY = &H1

Option Explicit
Dim BolIsMove As Boolean, MousX As Long, MousY As Long
 
Dim xCursor As Long, yCursor As Long

'---------------------------------------------------------------------------------------
' Module    : ModuleFile
' Author    : ROVAST
' Date      : 2014-4-22
' Purpose   : 文件相关操作模块
' Function  : 1、选取文件夹
'--------------------------------------------------------------------------------------
  
Private Type BrowseInfo
    hWndOwner As Long
    pIDLRoot As Long
    pszDisplayName As Long
    lpszTitle As Long
    ulFlags As Long
    lpfnCallback As Long
    lParam As Long
    iImage As Long
End Type
Const BIF_RETURNONLYFSDIRS = 1
Const BIF_NEWDIALOGSTYLE = &H40
Const BIF_EDITBOX = &H10
Const BIF_USENEWUI = BIF_NEWDIALOGSTYLE Or BIF_EDITBOX
Const MAX_PATH = 260
Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Long)
Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Private Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
  
  
'---------------------------------------------------------------------------------------
' Procedure : BrowseForFolder
' Author    : ROVAST
' Date      : 2014-4-22
' Purpose   : 选取文件夹（不含新建文件夹指令） 返回BrowseForFolder
'---------------------------------------------------------------------------------------
'
Public Function BrowseForFolder(Optional sTitle As String = "请选择您的 Animate CC 安装目录") As String
    Dim iNull As Integer, lpIDList As Long, lResult As Long
    Dim sPath As String, udtBI As BrowseInfo
  
    With udtBI
        .hWndOwner = 0 ' Me.hWnd
        .lpszTitle = lstrcat(sTitle, "")
        .ulFlags = BIF_RETURNONLYFSDIRS Or BIF_USENEWUI
    End With
    lpIDList = SHBrowseForFolder(udtBI)
    If lpIDList Then
       sPath = String$(MAX_PATH, 0)
        SHGetPathFromIDList lpIDList, sPath
        CoTaskMemFree lpIDList
       iNull = InStr(sPath, vbNullChar)
        If iNull Then
          sPath = Left$(sPath, iNull - 1)
        End If
    End If
  
    BrowseForFolder = sPath
End Function

Private Sub Image2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button And vbLeftButton Then
        xCursor = X: yCursor = Y
    End If
End Sub
Private Sub Image2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image3.Visible = False
    If Button And vbLeftButton Then
        Me.Move Me.Left - xCursor + X, Me.Top - yCursor + Y
    End If
End Sub


Private Sub Form_Load()
   Me.BackColor = &HFF0000
   Dim rtn As Long
   Dim BorderStyler
   BorderStyler = 0
   rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
   rtn = rtn Or WS_EX_LAYERED
   SetWindowLong hwnd, GWL_EXSTYLE, rtn
   SetLayeredWindowAttributes hwnd, &HFF0000, 0, LWA_COLORKEY
End Sub




Private Sub Image1_Click()

    End

End Sub


Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Image3.Visible = True
    
End Sub

Private Sub Image3_Click()

    End

End Sub

Private Sub Label1_Click()

    Shell App.path & "\install.bat", vbHide
    
    If Dir("C:\Program Files\Adobe\Adobe Animate CC 2017\Animate.exe") = "" Then
    
        If Dir("C:\Program Files (x86)\Adobe Animate CC 2017\Animate.exe") = "" Then
    
            MsgBox "检测不到您的 Animate CC 安装目录，请手动选择。"
    
           Dim path As String
           path = BrowseForFolder
           
           Dim SourceFile2 As String
            Dim DestinationFile2 As String
        
            SourceFile2 = App.path & "\others\IMSLib.dll"
            DestinationFile2 = path & "\IMSLib.dll"
        
            FileCopy SourceFile2, DestinationFile2
        
            MsgBox "安装成功", vbOKOnly + vbInformation
            End
    
        Else
    
            Dim SourceFile As String
            Dim DestinationFile As String
        
            SourceFile = App.path & "\others\IMSLib.dll"
            DestinationFile = "C:\Program Files (x86)\Adobe Animate CC 2017\IMSLib.dll"
        
            FileCopy SourceFile, DestinationFile
        
            MsgBox "安装成功", vbOKOnly + vbInformation
            End
    
        End If
        
    Else
    
        Dim SourceFile1 As String
        Dim DestinationFile1 As String
        
        SourceFile1 = App.path & "\others\IMSLib.dll"
        DestinationFile1 = "C:\Program Files\Adobe\Adobe Animate CC 2017\IMSLib.dll"
        
        FileCopy SourceFile1, DestinationFile1
        
        MsgBox "安装成功", vbOKOnly + vbInformation
        End
    
    End If

End Sub

