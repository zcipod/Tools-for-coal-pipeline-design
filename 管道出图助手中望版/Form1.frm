VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "�ܵ���ͼ���������桪��by Dream"
   ClientHeight    =   7485
   ClientLeft      =   5970
   ClientTop       =   4905
   ClientWidth     =   6990
   LinkTopic       =   "Form1"
   ScaleHeight     =   7485
   ScaleWidth      =   6990
   Begin VB.TextBox Text4 
      Height          =   270
      Left            =   1080
      MultiLine       =   -1  'True
      TabIndex        =   18
      Top             =   7080
      Width           =   5655
   End
   Begin VB.CommandButton Command4 
      Caption         =   "����"
      Height          =   375
      Left            =   5400
      TabIndex        =   10
      Top             =   720
      Width           =   975
   End
   Begin VB.Frame Frame2 
      Caption         =   "��ͷ��ܱ�ע����"
      Height          =   1935
      Left            =   120
      TabIndex        =   9
      Top             =   3360
      Width           =   6735
      Begin VB.CommandButton Command7 
         Caption         =   "��ʼ��ע"
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
         Text            =   "��������ʼ׮��"
         Top             =   1320
         Width           =   2535
      End
      Begin VB.CommandButton Command6 
         Caption         =   "ѡ��Ҫ��ע���ļ�"
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
         Text            =   "��ѡ��Ҫ��ע��CAD�ļ�"
         Top             =   840
         Width           =   4455
      End
      Begin VB.CommandButton Command5 
         Caption         =   "ѡ��ں���ļ�"
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
         Text            =   "��ѡ��ں���ļ�"
         Top             =   360
         Width           =   4455
      End
      Begin VB.Label Label3 
         Caption         =   "����CAD�ļ���ʼ׮��"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   1440
         Width           =   1935
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "���ƹ���"
      Height          =   2055
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   6735
      Begin VB.CommandButton Command3 
         Caption         =   "��ʼ����"
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
         Text            =   "��ѡ�������Ҫ���Ƶ�ͼ��"
         Top             =   1440
         Width           =   2655
      End
      Begin VB.CommandButton Command2 
         Caption         =   "ѡ��Ŀ���ļ�"
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
         Caption         =   "ѡ��Դ�ļ�"
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
      Begin VB.Label Label1 
         Caption         =   "����Ҫ���Ƶ�ͼ��"
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
      Caption         =   "����״̬"
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   7080
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "�ܵ���ͼ����"
      BeginProperty Font 
         Name            =   "������κ"
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
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Dim sFile As String
    Textbox1.Text = "��ѡ���Ѿ���ע�õ��ļ���Դ�ļ���"
If Textbox1.Text <> "��ѡ���Ѿ���ע�õ��ļ���Դ�ļ���" And Textbox1.Text <> "" Then
    sFile = Textbox1.Text
Else
    With dlgCommonDialog
        .DialogTitle = "��ѡ���Ѿ���ע�õ��ļ���Դ�ļ���"
        .CancelError = False
        .FileName = ""
        'ToDo: ���� common dialog �ؼ��ı�־������
        .Filter = "CAD�ļ� (*.dwg)|*.dwg|�����ļ�(*.*)|*.*"
        .ShowOpen
        If Len(.FileName) = 0 Then
            Exit Sub
        End If
        sFile = .FileName
    End With
End If
    'ToDo: ��Ӵ���򿪵��ļ��Ĵ���
If sFile = "" Then
Else:
    Textbox1.Text = sFile
sfile1 = sFile
End If

End Sub

Private Sub Command2_Click()
    Dim sFile As String
    Textbox2.Text = "��ѡ����Ŀ���ļ�"
If Textbox2.Text <> "��ѡ����Ŀ���ļ�" And Textbox2.Text <> "" Then
    sFile = Textbox2.Text
Else
    With dlgCommonDialog
        .DialogTitle = "��ѡ����Ŀ���ļ�"
        .CancelError = False
        .FileName = ""
        'ToDo: ���� common dialog �ؼ��ı�־������
        .Filter = "CAD�ļ� (*.dwg)|*.dwg|�����ļ�(*.*)|*.*"
        .ShowOpen
        If Len(.FileName) = 0 Then
            Exit Sub
        End If
        sFile = .FileName
    End With
End If
    'ToDo: ��Ӵ���򿪵��ļ��Ĵ���
If sFile = "" Then
Else:
    Textbox2.Text = sFile
sfile2 = sFile
End If

End Sub



Private Sub Command3_Click()

Dim zwcadapp As New ZwcadApplication
Dim zwcaddoc1 As New ZwcadDocument
Dim zwcaddoc2 As New ZwcadDocument
Dim zwcadPs1 As ZwcadLayout
Dim zwcadPs2 As ZwcadLayout
Dim SSet As ZwcadSelectionSet
Dim Pt1(0 To 2) As Double, Pt2(0 To 2) As Double
Dim retObjects As ZwcadSelectionSet
Dim retObjects1 As ZwcadSelectionSet
Dim pspace As ZwcadPaperSpace
Dim timest As Double, timeend As Double
timest = Timer
Pt1(0) = -5000
Pt1(1) = -5000
Pt1(2) = 0
Pt2(0) = 5000
Pt2(1) = 5000
Pt2(2) = 0
Text4.Text = "���ڴ�Ŀ���ļ�" & sfile2
Set zwcaddoc2 = zwcadapp.Documents.Open(sfile2)
Text4.Text = "���ڴ�Դ�ļ�"
Set zwcaddoc1 = zwcadapp.Documents.Open(sfile1)

If IsNull(zwcaddoc1) Then
    MsgBox "Դ�ļ�δѡ��"
    Exit Sub
End If

If IsNull(zwcaddoc2) Then
    MsgBox "Դ�ļ�δѡ��"
    Exit Sub
End If
Text4.Text = "���ڴ�CAD����"
'zwcadapp.Visible = True
'�˴���ʼѭ��
For ll = 0 To zwcaddoc1.Layouts.Count - 1

    If zwcaddoc1.Layouts.Item(ll).Name = "Model" Or zwcaddoc1.Layouts.Item(ll).Name = "����1" Then GoTo 123
    Set zwcadPs1 = zwcaddoc1.Layouts.Item(ll)
    
'    MsgBox zwcadPs1.Name
    
    For j = 0 To zwcaddoc2.Layouts.Count - 1
        If zwcaddoc2.Layouts.Item(j).Name = zwcadPs1.Name Then Exit For
    Next
    Set zwcadPs2 = zwcaddoc2.Layouts.Item(j)
    
    
'    MsgBox zwcadPs2.Name
Text4.Text = "���ڸ��ƣ�" & zwcadPs1.Name
zwcaddoc2.ActiveLayout = zwcadPs2
zwcaddoc1.ActiveLayout = zwcadPs1

'�˴���ʼ����
On Error Resume Next
If Not IsNull(zwcaddoc1.SelectionSets.Item("dd")) Then
    Set SSet = zwcaddoc1.SelectionSets.Item("dd")
    SSet.Delete
End If
Set SSet = zwcaddoc1.SelectionSets.Add("dd")
Dim Ft(0) As Integer, Fd(0) As Variant
Ft(0) = 8: Fd(0) = Textbox3.Text
'SSet.Select zcSelectionSetAll
zwcaddoc1.Activate
SSet.Select zcSelectionSetCrossingWindow, Pt1, Pt2, Ft, Fd

'MsgBox SSet.Count
'Dim objs() As ZwcadEntity
'ReDim objs(0 To SSet.Count - 1)
'For i = 0 To SSet.Count - 1
'    Set objs(i) = SSet.Item(i)
'Next
zwcaddoc2.Activate
'zwcaddoc2.ActiveLayout = zwcaddoc2.ModelSpace
Set retObjects = zwcaddoc2.CopyObjects(SSet)
zwcaddoc2.ActiveLayout = zwcadPs2
Set retObjects1 = zwcaddoc2.CopyObjects(retObjects)
retObjects.Erase
    
'�˴���������
123:
Next

'�˴�����ѭ��

zwcadapp.Visible = True
Text4.Text = "���������"
Set zwcaddoc1 = Nothing
Set zwcaddoc2 = Nothing
'zwcadapp.Quit
Set zwcadapp = Nothing
timeend = Timer
MsgBox "һ����ȥ" & Round(timeend - timest, 0) & "��"
End Sub

Private Sub Form_Load()

End Sub

Private Sub Text4_Change()

End Sub
