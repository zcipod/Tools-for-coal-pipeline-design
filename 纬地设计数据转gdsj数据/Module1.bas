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



Public imax As Integer          '����ƽ�����ݵ���Ŀ
Public dmxnum As Double         '������������ݵ���Ŀ
Public zdmnum As Double         '�����ݶ�����µ���Ŀ

Public jiaodian() As Point      '�������ݵ�ṹ������

Public dmxd() As dmx            '��������ߵ�����
Public zdmd() As zdm            '�����ݶ�����µ���������
Public savedir As String        '��������߷�ͼ���ݱ���·��
Public startdmx As Double, enddmx As Double         '������������ݷ�ͼ��ֹ׮��

Public wanguanR As Double           '������ܰ뾶
Public wantouR As Double            '������ͷ�뾶
Public pingtanR As Double           '����ƽ���뾶


















'�˶�Ϊѡ���ļ��г���


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
                  .lpszTitle = IIf(Len(Titel), Titel, "����ѡ���ļ��С�")
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

'�˶�Ϊѡ���ļ��г���


