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
		
		'UPGRADE_TODO: 必须调用“Initialize”来初始化此结构的实例。 单击以获得更多信息:“ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="B4BFF9E0-8631-45CF-910E-62AB3970F27B"”
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
	
	
	
	Public imax As Short '定义平面数据点数目
	Public dmxnum As Double '定义地面线数据点数目
	Public zdmnum As Double '定义纵断面变坡点数目
	
	Public jiaodian() As Point '定义数据点结构体数组
	
	Public dmxd() As dmx '定义地面线点数组
	Public zdmd() As zdm '定义纵断面变坡点数据数组
	Public savedir As String '定义地面线分图数据保存路径
	Public startdmx, enddmx As Double '定义地面线数据分图起止桩号
	
	Public wanguanR As Double '定义弯管半径
	Public wantouR As Double '定义弯头半径
	Public pingtanR As Double '定义平弹半径
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	'此段为选择文件夹程序
	
	
	Public numm As String
	Public num As Short
	
	Declare Function SendMessage Lib "user32"  Alias "SendMessageA"(ByVal hWnd As Integer, ByVal wMsg As Integer, ByVal wParam As Integer, ByVal lParam As String) As Integer
	Declare Function SHGetPathFromIDList Lib "shell32.dll"  Alias "SHGetPathFromIDListA"(ByVal pidl As Integer, ByVal pszPath As String) As Integer
	'UPGRADE_WARNING: 结构 BROWSEINFO 可能要求封送处理属性作为此 Declare 语句中的参数传递。 单击以获得更多信息:“ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"”
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
			.lpszTitle = IIf(Len(Titel), Titel, "【请选择文件夹】")
			.ulFlags = 7
			If Len(StartPath) Then
				xStartPath = StartPath & vbNullChar
				'UPGRADE_WARNING: 为 AddressOf CallBack 添加委托 单击以获得更多信息:“ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="E9E157F7-EF0C-4016-87B7-7D7FBBC6EE08"”
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
	
	'此段为选择文件夹程序
End Module