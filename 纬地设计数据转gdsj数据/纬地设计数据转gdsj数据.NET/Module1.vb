Option Strict Off
Option Explicit On
Module Module1
	Public Structure Point
		<VBFixedArray(2)> Dim zb() As Double
		Dim zj As Double
		Dim lc As Double
		Dim dh As String
		Dim jj As Double
		Dim xs As Short
		
		'UPGRADE_TODO: ������á�Initialize������ʼ���˽ṹ��ʵ���� �����Ի�ø�����Ϣ:��ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="B4BFF9E0-8631-45CF-910E-62AB3970F27B"��
		Public Sub Initialize()
			ReDim zb(2)
		End Sub
	End Structure
	
	Public Structure dmx
		Dim lc As Double
		Dim bg As Double
	End Structure
	
	Public Structure zdm
		Dim lc As Double
		Dim bg As Double
		Dim xs As Short
	End Structure
	
	
	
	Public imax As Short '����ƽ�����ݵ���Ŀ
	Public dmxnum As Double '������������ݵ���Ŀ
	Public zdmnum As Double '�����ݶ�����µ���Ŀ
	
	Public jiaodian() As Point '�������ݵ�ṹ������
	
	Public dmxd() As dmx '��������ߵ�����
	Public zdmd() As zdm '�����ݶ�����µ���������
	Public savedir As String '��������߷�ͼ���ݱ���·��
	Public startdmx, enddmx As Double '������������ݷ�ͼ��ֹ׮��
	
	Public wanguanR As Double '������ܰ뾶
	Public wantouR As Double '������ͷ�뾶
	Public pingtanR As Double '����ƽ���뾶
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	'�˶�Ϊѡ���ļ��г���
	
	
	Public numm As String
	Public num As Short
	
	Declare Function SendMessage Lib "user32"  Alias "SendMessageA"(ByVal hWnd As Integer, ByVal wMsg As Integer, ByVal wParam As Integer, ByVal lParam As String) As Integer
	Declare Function SHGetPathFromIDList Lib "shell32.dll"  Alias "SHGetPathFromIDListA"(ByVal pidl As Integer, ByVal pszPath As String) As Integer
	'UPGRADE_WARNING: �ṹ BROWSEINFO ����Ҫ����ʹ���������Ϊ�� Declare ����еĲ������ݡ� �����Ի�ø�����Ϣ:��ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"��
	Declare Function SHBrowseForFolder Lib "shell32.dll"  Alias "SHBrowseForFolderA"(ByRef lpBrowseInfo As BROWSEINFO) As Integer
	Structure BROWSEINFO
		Dim hOwner As Integer
		Dim pidlRoot As Integer
		Dim pszDisplayName As String
		Dim lpszTitle As String
		Dim ulFlags As Integer
		Dim lpfnCallback As Integer
		Dim lParam As Integer
		Dim iImage As Integer
	End Structure
	Dim xStartPath As String
	
	Function SelectDir(Optional ByRef StartPath As String = "", Optional ByRef Titel As String = "") As String
		Dim iBROWSEINFO As BROWSEINFO
		With iBROWSEINFO
			.lpszTitle = IIf(Len(Titel), Titel, "����ѡ���ļ��С�")
			.ulFlags = 7
			If Len(StartPath) Then
				xStartPath = StartPath & vbNullChar
				'UPGRADE_WARNING: Ϊ AddressOf CallBack ���ί�� �����Ի�ø�����Ϣ:��ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="E9E157F7-EF0C-4016-87B7-7D7FBBC6EE08"��
				.lpfnCallback = GetAddressOf(AddressOf CallBack)
			End If
		End With
		Dim xPath As String
		Dim NoErr As Integer : xPath = Space(512)
		NoErr = SHGetPathFromIDList(SHBrowseForFolder(iBROWSEINFO), xPath)
		SelectDir = IIf(NoErr, Left(xPath, InStr(xPath, Chr(0)) - 1), "")
	End Function
	
	Function GetAddressOf(ByRef Address As Integer) As Integer
		GetAddressOf = Address
	End Function
	
	Function CallBack(ByVal hWnd As Integer, ByVal Msg As Integer, ByVal pidl As Integer, ByVal pData As Integer) As Integer
		Dim sDir As New VB6.FixedLengthString(64)
		Dim tmp As Integer
		Select Case Msg
			Case 1
				Call SendMessage(hWnd, 1126, 1, xStartPath)
			Case 2
				tmp = SHGetPathFromIDList(pidl, sDir.Value)
				If tmp = 1 Then SendMessage(hWnd, 1124, 0, sDir.Value)
		End Select
	End Function
	
	'�˶�Ϊѡ���ļ��г���
End Module