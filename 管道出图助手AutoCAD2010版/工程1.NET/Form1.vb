Option Strict Off
Option Explicit On
Friend Class Form1
	Inherits System.Windows.Forms.Form
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
		Dim j As Object
		Dim ll As Object
		
		Dim Acadapp As New Autodesk.AutoCAD.Interop.AcadApplication
		Dim Acaddoc1 As New Autodesk.AutoCAD.Interop.AcadDocument
		Dim Acaddoc2 As New Autodesk.AutoCAD.Interop.AcadDocument
		Dim AcadPs1 As Autodesk.AutoCAD.Interop.AcadLayout
		Dim AcadPs2 As Autodesk.AutoCAD.Interop.AcadLayout
		Dim SSet As Autodesk.AutoCAD.Interop.AcadSelectionSet
		Dim Pt1(2) As Double
		Dim Pt2(2) As Double
		Pt1(0) = -5000
		Pt1(1) = -5000
		Pt1(2) = 0
		Pt2(0) = 5000
		Pt2(1) = 5000
		Pt2(2) = 0
		
		
		Acaddoc1 = Acadapp.Documents.Open(sfile1)
		Acaddoc2 = Acadapp.Documents.Open(sfile2)
		
		'UPGRADE_WARNING: 检测到使用了 Null/IsNull()。 单击以获得更多信息:“ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"”
		If IsDbNull(Acaddoc1) Then
			MsgBox("源文件未选择！")
			Exit Sub
		End If
		
		'UPGRADE_WARNING: 检测到使用了 Null/IsNull()。 单击以获得更多信息:“ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"”
		If IsDbNull(Acaddoc2) Then
			MsgBox("源文件未选择！")
			Exit Sub
		End If
		
		'Acadapp.Visible = True
		'此处开始循环
		
		Dim Ft(0) As Short
		Dim Fd(0) As Object
		Dim objs() As Autodesk.AutoCAD.Interop.AcadEntity
		For ll = 0 To Acaddoc1.Layouts.Count - 1
			If Acaddoc1.Layouts.Item(ll).Name = "Model" Then GoTo 123
			AcadPs1 = Acaddoc1.Layouts.Item(ll)
			
			'   MsgBox AcadPs1.Name
			
			For j = 0 To Acaddoc2.Layouts.Count - 1
				If Acaddoc2.Layouts.Item(j).Name = AcadPs1.Name Then Exit For
			Next 
			AcadPs2 = Acaddoc2.Layouts.Item(j)
			
			
			'    MsgBox AcadPs2.Name
			Acaddoc1.ActiveLayout = AcadPs1
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
		
		Acadapp.Visible = True
		'UPGRADE_NOTE: 在对对象 Acaddoc1 进行垃圾回收前，不可以将其销毁。 单击以获得更多信息:“ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"”
		Acaddoc1 = Nothing
		'UPGRADE_NOTE: 在对对象 Acaddoc2 进行垃圾回收前，不可以将其销毁。 单击以获得更多信息:“ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"”
		Acaddoc2 = Nothing
		'acadapp.Quit
		'UPGRADE_NOTE: 在对对象 Acadapp 进行垃圾回收前，不可以将其销毁。 单击以获得更多信息:“ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"”
		Acadapp = Nothing
	End Sub
	
	Private Sub Form1_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		
	End Sub
End Class