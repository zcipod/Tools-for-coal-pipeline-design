VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmTest 
   Caption         =   "管道出图助手——by Dream"
   ClientHeight    =   7170
   ClientLeft      =   5970
   ClientTop       =   4905
   ClientWidth     =   6990
   LinkTopic       =   "Form1"
   ScaleHeight     =   7170
   ScaleWidth      =   6990
   Begin VB.TextBox Text4 
      Height          =   270
      Left            =   1080
      MultiLine       =   -1  'True
      TabIndex        =   18
      Top             =   6720
      Width           =   5655
   End
   Begin VB.CommandButton Command4 
      Caption         =   "帮助"
      Height          =   375
      Left            =   5400
      TabIndex        =   10
      Top             =   720
      Width           =   975
   End
   Begin VB.Frame Frame2 
      Caption         =   "弯头弯管标注工具"
      Height          =   1935
      Left            =   120
      TabIndex        =   9
      Top             =   3840
      Width           =   6735
      Begin VB.CommandButton Command7 
         Caption         =   "开始标注"
         Height          =   375
         Left            =   4920
         TabIndex        =   17
         Top             =   1320
         Width           =   1695
      End
      Begin VB.TextBox Text3 
         Height          =   375
         Left            =   2040
         TabIndex        =   16
         Text            =   "请输入起始桩号"
         Top             =   1320
         Width           =   2535
      End
      Begin VB.CommandButton Command6 
         Caption         =   "选择要标注的文件"
         Height          =   375
         Left            =   4920
         TabIndex        =   14
         Top             =   840
         Width           =   1695
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Text            =   "请选择要标注的CAD文件"
         Top             =   840
         Width           =   4455
      End
      Begin VB.CommandButton Command5 
         Caption         =   "选择壁厚表文件"
         Height          =   375
         Left            =   4920
         TabIndex        =   12
         Top             =   360
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Text            =   "请选择壁厚表文件"
         Top             =   360
         Width           =   4455
      End
      Begin VB.Label Label3 
         Caption         =   "输入CAD文件起始桩号"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   1440
         Width           =   1935
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "复制工具"
      Height          =   2175
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   6735
      Begin VB.CheckBox Check1 
         Caption         =   "Check1"
         Height          =   255
         Left            =   1920
         TabIndex        =   21
         Top             =   1800
         Width           =   255
      End
      Begin VB.CommandButton Command3 
         Caption         =   "开始复制"
         Height          =   375
         Left            =   4920
         TabIndex        =   7
         Top             =   1320
         Width           =   1695
      End
      Begin VB.ComboBox Textbox3 
         Height          =   300
         ItemData        =   "Form1.frx":0000
         Left            =   1920
         List            =   "Form1.frx":000D
         TabIndex        =   6
         Text            =   "请选择或输入要复制的图层"
         Top             =   1440
         Width           =   2655
      End
      Begin VB.CommandButton Command2 
         Caption         =   "选择目标文件"
         Height          =   375
         Left            =   4920
         TabIndex        =   5
         Top             =   840
         Width           =   1695
      End
      Begin VB.TextBox Textbox2 
         Height          =   375
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   4
         TabStop         =   0   'False
         Text            =   "Form1.frx":002D
         Top             =   840
         Width           =   4455
      End
      Begin VB.CommandButton Command1 
         Caption         =   "选择源文件"
         Height          =   375
         Left            =   4920
         TabIndex        =   3
         Top             =   360
         Width           =   1695
      End
      Begin VB.TextBox Textbox1 
         Height          =   375
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   2
         TabStop         =   0   'False
         Text            =   "Form1.frx":0040
         Top             =   360
         Width           =   4455
      End
      Begin VB.Label Label5 
         Caption         =   "是否提示复制信息"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   1800
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "输入要复制的图层"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   1440
         Width           =   1695
      End
   End
   Begin MSComDlg.CommonDialog dlgCommonDialog 
      Left            =   360
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label4 
      Caption         =   "复制状态"
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   6720
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "管道出图助手"
      BeginProperty Font 
         Name            =   "华文新魏"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      TabIndex        =   0
      Top             =   240
      Width           =   2295
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal HwndNewparent As Long) As Long
Private Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Private m_oapp As Object

Public Property Set application(ByVal vnewvalue As Object)
    Set m_oapp = vnewvalue
End Property
Private Sub Form_Load()
    SetParent Me.hwnd, GetParent(GetParent(m_oapp.ActiveDocument.hwnd))
End Sub



Private Sub Command1_Click()
    Dim sFile As String
    Textbox1.Text = "请选择已经标注好的文件（源文件）"
If Textbox1.Text <> "请选择已经标注好的文件（源文件）" And Textbox1.Text <> "" Then
    sFile = Textbox1.Text
Else
    With dlgCommonDialog
        .DialogTitle = "请选择已经标注好的文件（源文件）"
        .CancelError = False
        .FileName = ""
        'ToDo: 设置 common dialog 控件的标志和属性
        .Filter = "CAD文件 (*.dwg)|*.dwg|所有文件(*.*)|*.*"
        .ShowOpen
        If Len(.FileName) = 0 Then
            Exit Sub
        End If
        sFile = .FileName
    End With
End If
    'ToDo: 添加处理打开的文件的代码
If sFile = "" Then
Else:
    Textbox1.Text = sFile
sfile1 = sFile
End If

End Sub

Private Sub Command2_Click()
    Dim sFile As String
    Textbox2.Text = "请选择复制目标文件"
If Textbox2.Text <> "请选择复制目标文件" And Textbox2.Text <> "" Then
    sFile = Textbox2.Text
Else
    With dlgCommonDialog
        .DialogTitle = "请选择复制目标文件"
        .CancelError = False
        .FileName = ""
        'ToDo: 设置 common dialog 控件的标志和属性
        .Filter = "CAD文件 (*.dwg)|*.dwg|所有文件(*.*)|*.*"
        .ShowOpen
        If Len(.FileName) = 0 Then
            Exit Sub
        End If
        sFile = .FileName
    End With
End If
    'ToDo: 添加处理打开的文件的代码
If sFile = "" Then
Else:
    Textbox2.Text = sFile
sfile2 = sFile
End If

End Sub



Private Sub Command3_Click()

Dim Acadapp As AcadApplication
Dim Acaddoc1 As AcadDocument
Dim Acaddoc2 As AcadDocument
Dim AcadPs1 As AcadLayout
Dim AcadPs2 As AcadLayout
Dim SSet As AcadSelectionSet
Dim timest As Double, timeend As Double
timest = Timer

On Error Resume Next
Set Acadapp = GetObject(, "AutoCAD.Application.18")
'Acadapp.Visible = False

Dim Pt1(0 To 2) As Double, Pt2(0 To 2) As Double
Pt1(0) = -5000
Pt1(1) = -5000
Pt1(2) = 0
Pt2(0) = 5000
Pt2(1) = 5000
Pt2(2) = 0

Text4.Text = "正在打开目标文件" & sfile2
Set Acaddoc1 = Acadapp.Documents.Open(sfile1)
Text4.Text = "正在打开源文件"
Set Acaddoc2 = Acadapp.Documents.Open(sfile2)

If IsNull(Acaddoc1) Then
    MsgBox "源文件未选择！"
    Exit Sub
End If
Text4.Text = "正在打开CAD窗口"
If IsNull(Acaddoc2) Then
    MsgBox "源文件未选择！"
    Exit Sub
End If


'此处开始循环
If Check1.Value = True Then MsgBox "文件载入成功，开始复制！"
For ll = 0 To Acaddoc1.Layouts.Count - 1
    If Acaddoc1.Layouts.Item(ll).Name = "Model" Or Acaddoc1.Layouts.Item(ll).Name = "布局1" Then GoTo 123
    Set AcadPs1 = Acaddoc1.Layouts.Item(ll)
    
 '   MsgBox AcadPs1.Name
    
    For j = 0 To Acaddoc2.Layouts.Count - 1
        If Acaddoc2.Layouts.Item(j).Name = AcadPs1.Name Then Exit For
    Next
    Set AcadPs2 = Acaddoc2.Layouts.Item(j)
    
    
'    MsgBox AcadPs2.Name
If Check1.Value = True Then MsgBox "正在复制：" & accadPs1.Name
Text4.Text = "正在复制：" & accadPs1.Name
Acaddoc1.ActiveLayout = AcadPs1
Acadapp.ZoomAll
Acaddoc2.ActiveLayout = AcadPs2
'此处开始复制
On Error Resume Next
If Not IsNull(Acaddoc1.SelectionSets.Item("dd")) Then
    Set SSet = Acaddoc1.SelectionSets.Item("dd")
    SSet.Delete
End If
Set SSet = Acaddoc1.SelectionSets.Add("dd")
Dim Ft(0) As Integer, Fd(0)
Ft(0) = 8: Fd(0) = Textbox3.Text
'SSet.Select acSelectionSetAll, , , Ft, Fd
Acadapp.ZoomAll
SSet.Select acSelectionSetCrossing, Pt1, Pt2, Ft, Fd

'MsgBox SSet.Count
Dim objs() As AcadEntity
ReDim objs(0 To SSet.Count - 1)
For i = 0 To SSet.Count - 1
    Set objs(i) = SSet.Item(i)
Next
    Acaddoc1.CopyObjects objs, Acaddoc2.PaperSpace
'此处结束复制
123:
Next

'此处结束循环

'Acadapp.Visible = True
Text4.Text = "复制已完成"
Set Acaddoc1 = Nothing
Set Acaddoc2 = Nothing
'acadapp.Quit
Set Acadapp = Nothing
timeend = Timer
MsgBox "一共用去" & Round(timeend - timest, 0) & "秒"
End Sub

