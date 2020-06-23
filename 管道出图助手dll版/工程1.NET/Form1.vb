Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Friend Class frmTest
	Inherits System.Windows.Forms.Form
	Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Integer, ByVal HwndNewparent As Integer) As Integer
	Private Declare Function GetParent Lib "user32" (ByVal hwnd As Integer) As Integer
	Private m_oapp As Object
	
	Public WriteOnly Property application() As Object
		Set(ByVal Value As Object)
			m_oapp = Value
		End Set
	End Property
	Private Sub frmTest_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		'UPGRADE_WARNING: 未能解析对象 m_oapp.ActiveDocument 的默认属性。 单击以获得更多信息:“ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"”
		SetParent(Me.Handle.ToInt32, GetParent(GetParent(m_oapp.ActiveDocument.hwnd)))
	End Sub
	
	
	
	Private Sub Command1_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command1.Click
		Dim sFile As String
		Textbox1.Text = "请选择已经标注好的文件（源文件）"
		If Textbox1.Text <> "请选择已经标注好的文件（源文件）" And Textbox1.Text <> "" Then
			sFile = Textbox1.Text
		Else
			'UPGRADE_WARNING: CommonDialog 变量未升级 单击以获得更多信息:“ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="671167DC-EA81-475D-B690-7A40C7BF4A23"”
			With dlgCommonDialog
				.Title = "请选择已经标注好的文件（源文件）"
				'UPGRADE_WARNING: 在 Visual Basic .NET 中不支持 CommonDialog CancelError 属性。 单击以获得更多信息:“ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="8B377936-3DF7-4745-AA26-DD00FA5B9BE1"”
				.CancelError = False
				.FileName = ""
				'ToDo: 设置 common dialog 控件的标志和属性
				'UPGRADE_WARNING: Filter 有新行为。 单击以获得更多信息:“ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"”
				.Filter = "CAD文件 (*.dwg)|*.dwg|所有文件(*.*)|*.*"
				.ShowDialog()
				If Len(.FileName) = 0 Then
					Exit Sub
				End If
				sFile = .FileName
			End With
		End If
		'ToDo: 添加处理打开的文件的代码
		If sFile = "" Then
		Else
			Textbox1.Text = sFile
			sfile1 = sFile
		End If
		
	End Sub
	
	Private Sub Command2_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command2.Click
		Dim sFile As String
		Textbox2.Text = "请选择复制目标文件"
		If Textbox2.Text <> "请选择复制目标文件" And Textbox2.Text <> "" Then
			sFile = Textbox2.Text
		Else
			'UPGRADE_WARNING: CommonDialog 变量未升级 单击以获得更多信息:“ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="671167DC-EA81-475D-B690-7A40C7BF4A23"”
			With dlgCommonDialog
				.Title = "请选择复制目标文件"
				'UPGRADE_WARNING: 在 Visual Basic .NET 中不支持 CommonDialog CancelError 属性。 单击以获得更多信息:“ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="8B377936-3DF7-4745-AA26-DD00FA5B9BE1"”
				.CancelError = False
				.FileName = ""
				'ToDo: 设置 common dialog 控件的标志和属性
				'UPGRADE_WARNING: Filter 有新行为。 单击以获得更多信息:“ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"”
				.Filter = "CAD文件 (*.dwg)|*.dwg|所有文件(*.*)|*.*"
				.ShowDialog()
				If Len(.FileName) = 0 Then
					Exit Sub
				End If
				sFile = .FileName
			End With
		End If
		'ToDo: 添加处理打开的文件的代码
		If sFile = "" Then
		Else
			Textbox2.Text = sFile
			sfile2 = sFile
		End If
		
	End Sub
	
	
	
	Private Sub Command3_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command3.Click
		Dim i As Object
		Dim accadPs1 As Object
		Dim j As Object
		Dim ll As Object
		
		Dim Acadapp As Autodesk.AutoCAD.Interop.AcadApplication
		Dim Acaddoc1 As Autodesk.AutoCAD.Interop.AcadDocument
		Dim Acaddoc2 As Autodesk.AutoCAD.Interop.AcadDocument
		Dim AcadPs1 As Autodesk.AutoCAD.Interop.AcadLayout
		Dim AcadPs2 As Autodesk.AutoCAD.Interop.AcadLayout
		Dim SSet As Autodesk.AutoCAD.Interop.AcadSelectionSet
		Dim timest, timeend As Double
		timest = VB.Timer()
		
		On Error Resume Next
		Acadapp = GetObject( , "AutoCAD.Application.18")
		'Acadapp.Visible = False
		
		Dim Pt1(2) As Double
		Dim Pt2(2) As Double
		Pt1(0) = -5000
		Pt1(1) = -5000
		Pt1(2) = 0
		Pt2(0) = 5000
		Pt2(1) = 5000
		Pt2(2) = 0
		
		Text4.Text = "正在打开目标文件" & sfile2
		Acaddoc1 = Acadapp.Documents.Open(sfile1)
		Text4.Text = "正在打开源文件"
		Acaddoc2 = Acadapp.Documents.Open(sfile2)
		
		'UPGRADE_WARNING: 检测到使用了 Null/IsNull()。 单击以获得更多信息:“ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"”
		If IsDbNull(Acaddoc1) Then
			MsgBox("源文件未选择！")
			Exit Sub
		End If
		Text4.Text = "正在打开CAD窗口"
		'UPGRADE_WARNING: 检测到使用了 Null/IsNull()。 单击以获得更多信息:“ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"”
		If IsDbNull(Acaddoc2) Then
			MsgBox("源文件未选择！")
			Exit Sub
		End If
		
		
		'此处开始循环
		If Check1.CheckState = True Then MsgBox("文件载入成功，开始复制！")
		Dim Ft(0) As Short
		Dim Fd(0) As Object
		Dim objs() As Autodesk.AutoCAD.Interop.AcadEntity
		For ll = 0 To Acaddoc1.Layouts.Count - 1
			If Acaddoc1.Layouts.Item(ll).Name = "Model" Or Acaddoc1.Layouts.Item(ll).Name = "布局1" Then GoTo 123
			AcadPs1 = Acaddoc1.Layouts.Item(ll)
			
			'   MsgBox AcadPs1.Name
			
			For j = 0 To Acaddoc2.Layouts.Count - 1
				If Acaddoc2.Layouts.Item(j).Name = AcadPs1.Name Then Exit For
			Next 
			AcadPs2 = Acaddoc2.Layouts.Item(j)
			
			
			'    MsgBox AcadPs2.Name
			'UPGRADE_WARNING: 未能解析对象 accadPs1.Name 的默认属性。 单击以获得更多信息:“ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"”
			If Check1.CheckState = True Then MsgBox("正在复制：" & accadPs1.Name)
			'UPGRADE_WARNING: 未能解析对象 accadPs1.Name 的默认属性。 单击以获得更多信息:“ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"”
			Text4.Text = "正在复制：" & accadPs1.Name
			Acaddoc1.ActiveLayout = AcadPs1
			Acadapp.ZoomAll()
			Acaddoc2.ActiveLayout = AcadPs2
			'此处开始复制
			On Error Resume Next
			'UPGRADE_WARNING: 检测到使用了 Null/IsNull()。 单击以获得更多信息:“ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"”
			If Not IsDbNull(Acaddoc1.SelectionSets.Item("dd")) Then
				SSet = Acaddoc1.SelectionSets.Item("dd")
				SSet.Delete()
			End If
			SSet = Acaddoc1.SelectionSets.Add("dd")
			Ft(0) = 8
			'UPGRADE_WARNING: 未能解析对象 Fd(0) 的默认属性。 单击以获得更多信息:“ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"”
			Fd(0) = Textbox3.Text
			'SSet.Select acSelectionSetAll, , , Ft, Fd
			Acadapp.ZoomAll()
			SSet.Select(Autodesk.AutoCAD.Interop.AcSelect.acSelectionSetCrossing, Pt1, Pt2, Ft, Fd)
			
			'MsgBox SSet.Count
			ReDim objs(SSet.Count - 1)
			For i = 0 To SSet.Count - 1
				'UPGRADE_WARNING: 未能解析对象 i 的默认属性。 单击以获得更多信息:“ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"”
				objs(i) = SSet.Item(i)
			Next 
			Acaddoc1.CopyObjects(objs, Acaddoc2.PaperSpace)
			'此处结束复制
123: 
		Next 
		
		'此处结束循环
		
		'Acadapp.Visible = True
		Text4.Text = "复制已完成"
		'UPGRADE_NOTE: 在对对象 Acaddoc1 进行垃圾回收前，不可以将其销毁。 单击以获得更多信息:“ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"”
		Acaddoc1 = Nothing
		'UPGRADE_NOTE: 在对对象 Acaddoc2 进行垃圾回收前，不可以将其销毁。 单击以获得更多信息:“ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"”
		Acaddoc2 = Nothing
		'acadapp.Quit
		'UPGRADE_NOTE: 在对对象 Acadapp 进行垃圾回收前，不可以将其销毁。 单击以获得更多信息:“ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"”
		Acadapp = Nothing
		timeend = VB.Timer()
		MsgBox("一共用去" & System.Math.Round(timeend - timest, 0) & "秒")
	End Sub
End Class