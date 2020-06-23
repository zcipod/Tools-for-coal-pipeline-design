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
		'UPGRADE_WARNING: δ�ܽ������� m_oapp.ActiveDocument ��Ĭ�����ԡ� �����Ի�ø�����Ϣ:��ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"��
		SetParent(Me.Handle.ToInt32, GetParent(GetParent(m_oapp.ActiveDocument.hwnd)))
	End Sub
	
	
	
	Private Sub Command1_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command1.Click
		Dim sFile As String
		Textbox1.Text = "��ѡ���Ѿ���ע�õ��ļ���Դ�ļ���"
		If Textbox1.Text <> "��ѡ���Ѿ���ע�õ��ļ���Դ�ļ���" And Textbox1.Text <> "" Then
			sFile = Textbox1.Text
		Else
			'UPGRADE_WARNING: CommonDialog ����δ���� �����Ի�ø�����Ϣ:��ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="671167DC-EA81-475D-B690-7A40C7BF4A23"��
			With dlgCommonDialog
				.Title = "��ѡ���Ѿ���ע�õ��ļ���Դ�ļ���"
				'UPGRADE_WARNING: �� Visual Basic .NET �в�֧�� CommonDialog CancelError ���ԡ� �����Ի�ø�����Ϣ:��ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="8B377936-3DF7-4745-AA26-DD00FA5B9BE1"��
				.CancelError = False
				.FileName = ""
				'ToDo: ���� common dialog �ؼ��ı�־������
				'UPGRADE_WARNING: Filter ������Ϊ�� �����Ի�ø�����Ϣ:��ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"��
				.Filter = "CAD�ļ� (*.dwg)|*.dwg|�����ļ�(*.*)|*.*"
				.ShowDialog()
				If Len(.FileName) = 0 Then
					Exit Sub
				End If
				sFile = .FileName
			End With
		End If
		'ToDo: ��Ӵ���򿪵��ļ��Ĵ���
		If sFile = "" Then
		Else
			Textbox1.Text = sFile
			sfile1 = sFile
		End If
		
	End Sub
	
	Private Sub Command2_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command2.Click
		Dim sFile As String
		Textbox2.Text = "��ѡ����Ŀ���ļ�"
		If Textbox2.Text <> "��ѡ����Ŀ���ļ�" And Textbox2.Text <> "" Then
			sFile = Textbox2.Text
		Else
			'UPGRADE_WARNING: CommonDialog ����δ���� �����Ի�ø�����Ϣ:��ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="671167DC-EA81-475D-B690-7A40C7BF4A23"��
			With dlgCommonDialog
				.Title = "��ѡ����Ŀ���ļ�"
				'UPGRADE_WARNING: �� Visual Basic .NET �в�֧�� CommonDialog CancelError ���ԡ� �����Ի�ø�����Ϣ:��ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="8B377936-3DF7-4745-AA26-DD00FA5B9BE1"��
				.CancelError = False
				.FileName = ""
				'ToDo: ���� common dialog �ؼ��ı�־������
				'UPGRADE_WARNING: Filter ������Ϊ�� �����Ի�ø�����Ϣ:��ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"��
				.Filter = "CAD�ļ� (*.dwg)|*.dwg|�����ļ�(*.*)|*.*"
				.ShowDialog()
				If Len(.FileName) = 0 Then
					Exit Sub
				End If
				sFile = .FileName
			End With
		End If
		'ToDo: ��Ӵ���򿪵��ļ��Ĵ���
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
		
		Text4.Text = "���ڴ�Ŀ���ļ�" & sfile2
		Acaddoc1 = Acadapp.Documents.Open(sfile1)
		Text4.Text = "���ڴ�Դ�ļ�"
		Acaddoc2 = Acadapp.Documents.Open(sfile2)
		
		'UPGRADE_WARNING: ��⵽ʹ���� Null/IsNull()�� �����Ի�ø�����Ϣ:��ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"��
		If IsDbNull(Acaddoc1) Then
			MsgBox("Դ�ļ�δѡ��")
			Exit Sub
		End If
		Text4.Text = "���ڴ�CAD����"
		'UPGRADE_WARNING: ��⵽ʹ���� Null/IsNull()�� �����Ի�ø�����Ϣ:��ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"��
		If IsDbNull(Acaddoc2) Then
			MsgBox("Դ�ļ�δѡ��")
			Exit Sub
		End If
		
		
		'�˴���ʼѭ��
		If Check1.CheckState = True Then MsgBox("�ļ�����ɹ�����ʼ���ƣ�")
		Dim Ft(0) As Short
		Dim Fd(0) As Object
		Dim objs() As Autodesk.AutoCAD.Interop.AcadEntity
		For ll = 0 To Acaddoc1.Layouts.Count - 1
			If Acaddoc1.Layouts.Item(ll).Name = "Model" Or Acaddoc1.Layouts.Item(ll).Name = "����1" Then GoTo 123
			AcadPs1 = Acaddoc1.Layouts.Item(ll)
			
			'   MsgBox AcadPs1.Name
			
			For j = 0 To Acaddoc2.Layouts.Count - 1
				If Acaddoc2.Layouts.Item(j).Name = AcadPs1.Name Then Exit For
			Next 
			AcadPs2 = Acaddoc2.Layouts.Item(j)
			
			
			'    MsgBox AcadPs2.Name
			'UPGRADE_WARNING: δ�ܽ������� accadPs1.Name ��Ĭ�����ԡ� �����Ի�ø�����Ϣ:��ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"��
			If Check1.CheckState = True Then MsgBox("���ڸ��ƣ�" & accadPs1.Name)
			'UPGRADE_WARNING: δ�ܽ������� accadPs1.Name ��Ĭ�����ԡ� �����Ի�ø�����Ϣ:��ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"��
			Text4.Text = "���ڸ��ƣ�" & accadPs1.Name
			Acaddoc1.ActiveLayout = AcadPs1
			Acadapp.ZoomAll()
			Acaddoc2.ActiveLayout = AcadPs2
			'�˴���ʼ����
			On Error Resume Next
			'UPGRADE_WARNING: ��⵽ʹ���� Null/IsNull()�� �����Ի�ø�����Ϣ:��ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"��
			If Not IsDbNull(Acaddoc1.SelectionSets.Item("dd")) Then
				SSet = Acaddoc1.SelectionSets.Item("dd")
				SSet.Delete()
			End If
			SSet = Acaddoc1.SelectionSets.Add("dd")
			Ft(0) = 8
			'UPGRADE_WARNING: δ�ܽ������� Fd(0) ��Ĭ�����ԡ� �����Ի�ø�����Ϣ:��ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"��
			Fd(0) = Textbox3.Text
			'SSet.Select acSelectionSetAll, , , Ft, Fd
			Acadapp.ZoomAll()
			SSet.Select(Autodesk.AutoCAD.Interop.AcSelect.acSelectionSetCrossing, Pt1, Pt2, Ft, Fd)
			
			'MsgBox SSet.Count
			ReDim objs(SSet.Count - 1)
			For i = 0 To SSet.Count - 1
				'UPGRADE_WARNING: δ�ܽ������� i ��Ĭ�����ԡ� �����Ի�ø�����Ϣ:��ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"��
				objs(i) = SSet.Item(i)
			Next 
			Acaddoc1.CopyObjects(objs, Acaddoc2.PaperSpace)
			'�˴���������
123: 
		Next 
		
		'�˴�����ѭ��
		
		'Acadapp.Visible = True
		Text4.Text = "���������"
		'UPGRADE_NOTE: �ڶԶ��� Acaddoc1 ������������ǰ�������Խ������١� �����Ի�ø�����Ϣ:��ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"��
		Acaddoc1 = Nothing
		'UPGRADE_NOTE: �ڶԶ��� Acaddoc2 ������������ǰ�������Խ������١� �����Ի�ø�����Ϣ:��ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"��
		Acaddoc2 = Nothing
		'acadapp.Quit
		'UPGRADE_NOTE: �ڶԶ��� Acadapp ������������ǰ�������Խ������١� �����Ի�ø�����Ϣ:��ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"��
		Acadapp = Nothing
		timeend = VB.Timer()
		MsgBox("һ����ȥ" & System.Math.Round(timeend - timest, 0) & "��")
	End Sub
End Class