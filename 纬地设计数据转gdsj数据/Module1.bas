Attribute VB_Name = "Module1"
Public Type Point
    zb(0 To 2) As Double
    zj As Double
    lc As Double
    dh As String
    jj As Double
    xs As Integer
  End Type
  
Public Type dmx
    lc As Double
    bg As Double
  End Type

Public Type zdm
    lc As Double
    bg As Double
    xs As Integer
End Type



Public imax As Integer          '定义平面数据点数目
Public dmxnum As Double         '定义地面线数据点数目
Public zdmnum As Double         '定义纵断面变坡点数目

Public jiaodian() As Point      '定义数据点结构体数组

Public dmxd() As dmx            '定义地面线点数组
Public zdmd() As zdm            '定义纵断面变坡点数据数组
Public savedir As String        '定义地面线分图数据保存路径
Public startdmx As Double, enddmx As Double         '定义地面线数据分图起止桩号

Public wanguanR As Double           '定义弯管半径
Public wantouR As Double            '定义弯头半径
Public pingtanR As Double           '定义平弹半径


















'此段为选择文件夹程序


Public numm As String
Public num As Integer
Option Explicit
   
Declare Function SendMessage Lib "user32" Alias "SendMessageA" ( _
          ByVal hWnd As Long, _
          ByVal wMsg As Long, _
          ByVal wParam As Long, _
          ByVal lParam As String) As Long
Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" ( _
          ByVal pidl As Long, _
          ByVal pszPath As String) As Long
Declare Function SHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolderA" ( _
          lpBrowseInfo As BROWSEINFO) As Long
Type BROWSEINFO
          hOwner   As Long
          pidlRoot   As Long
          pszDisplayName   As String
          lpszTitle   As String
          ulFlags   As Long
          lpfnCallback   As Long
          lParam   As Long
          iImage   As Long
End Type
Dim xStartPath     As String
   
Function SelectDir(Optional StartPath As String, Optional Titel As String) As String
          Dim iBROWSEINFO     As BROWSEINFO
          With iBROWSEINFO
                  .lpszTitle = IIf(Len(Titel), Titel, "【请选择文件夹】")
                  .ulFlags = 7
                  If Len(StartPath) Then
                  xStartPath = StartPath & vbNullChar
                  .lpfnCallback = GetAddressOf(AddressOf CallBack)
                  End If
          End With
          Dim xPath     As String, NoErr       As Long:     xPath = Space$(512)
          NoErr = SHGetPathFromIDList(SHBrowseForFolder(iBROWSEINFO), xPath)
          SelectDir = IIf(NoErr, Left$(xPath, InStr(xPath, Chr(0)) - 1), "")
End Function
   
Function GetAddressOf(Address As Long) As Long
          GetAddressOf = Address
End Function
   
Function CallBack(ByVal hWnd As Long, _
                                      ByVal Msg As Long, _
                                      ByVal pidl As Long, _
                                      ByVal pData As Long) As Long
          Select Case Msg
                  Case 1
                          Call SendMessage(hWnd, 1126, 1, xStartPath)
                  Case 2
                          Dim sDir     As String * 64, tmp           As Long
                          tmp = SHGetPathFromIDList(pidl, sDir)
                          If tmp = 1 Then SendMessage hWnd, 1124, 0, sDir
          End Select
End Function

'此段为选择文件夹程序


