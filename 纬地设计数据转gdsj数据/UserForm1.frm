VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "管道出图助手-By：Dream"
   ClientHeight    =   4860
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7080
   OleObjectBlob   =   "UserForm1.dsx":0000
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function OSWinHelp% Lib "user32" Alias "WinHelpA" (ByVal hWnd&, ByVal HelpFile$, ByVal wCommand%, dwData As Any)


Private Sub CommandButton10_Click()
Dim a(0 To 2) As Double






   Dim PtPick As Variant
    UserForm1.Hide
        PtPick = ThisDrawing.Utility.GetPoint(, "选择点")
    a(0) = PtPick(0): a(1) = PtPick(1)
        
'a(0) = 7445
'a(1) = -2900
a(2) = 0
Dim n As Integer
For n = 285 To 351

    ThisDrawing.ModelSpace.AddText n, a, 80
    a(0) = a(0) + 500
Next
UserForm1.Show
End Sub

Private Sub CommandButton2_Click()
    Dim sFile As String
If TextBox1.Value <> "请选择要打开的平面交点数据文件" And TextBox1.Value <> "" Then
    sFile = TextBox1.Value
Else
    With dlgCommonDialog
        .DialogTitle = "打开"
        .CancelError = False
        .FileName = ""
        'ToDo: 设置 common dialog 控件的标志和属性
        .Filter = "Excel文件 (*.xls;*.xlsx)|*.xls;*.xlsx|所有文件(*.*)|*.*"
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
    TextBox1.Value = sFile
    Dim xlApp As New Excel.Application
    Dim xlBook As New Excel.Workbook
    Dim xlSheet As New Excel.Worksheet
    
    On Error Resume Next
       '后台进程运行excel程序，并得到该工作簿
       Set xlBook = xlApp.Workbooks.Open(sFile)
       xlApp.Visible = False
       '获得该工作簿的“sheet1”表
       Set xlSheet = xlBook.Sheets("sheet1")
        
       '读取excel单元数据
    '    xlSheet.Cells(1, 1).Value = Text1.Text
    '    xlSheet.Cells(2, 1).Value = Text2.Text
        Dim x  As Integer
        x = 1
    Do
         x = x + 1
         If xlSheet.Cells(x, 1).Value = "" Then
         x = x - 1
         imax = x - 1
         Exit Do
         Else:
         
         ReDim Preserve jiaodian(x - 1)
         

        jiaodian(x - 2).dh = xlSheet.Cells(x, 1).Value
        jiaodian(x - 2).zb(0) = Round(xlSheet.Cells(x, 3).Value, 3)
        jiaodian(x - 2).zb(1) = Round(xlSheet.Cells(x, 2).Value, 3)
        jiaodian(x - 2).zb(2) = 0
        jiaodian(x - 2).lc = xlSheet.Cells(x, 4).Value
        jiaodian(x - 2).jj = xlSheet.Cells(x, 5).Value
        jiaodian(x - 2).zj = xlSheet.Cells(x, 6).Value
        jiaodian(x - 2).xs = xlSheet.Cells(x, 7).Value
        
         End If
         
        Loop
End If
       Set xlSheet = Nothing
       Set xlBook = Nothing
       xlApp.Visible = Ture
       xlApp.Quit
       Set xlApp = Nothing
MsgBox "平面数据读取成功"
End Sub

'函数
'pi值计算
Public Function pi()
    pi = 2 * (Atn(0) + 2 * Atn(1))
End Function

'旋转角度计算
Public Function fangweijiao(zb1() As Double, zb2() As Double)
    Dim sinvalue As Double
    Dim jiaodu As Double
    sinvalue = (zb2(1) - zb1(1)) / Sqr((zb2(1) - zb1(1)) * (zb2(1) - zb1(1)) + (zb2(0) - zb1(0)) * (zb2(0) - zb1(0)))
'    MsgBox sinvalue
    jiaodu = Atn(sinvalue / Sqr(-sinvalue * sinvalue + 1))
'    MsgBox pi()
'    MsgBox jiaodu
    If zb2(0) - zb1(0) > 0 Then
        fangweijiao = jiaodu
    Else
        If zb2(1) - zb1(1) > 0 Then
            fangweijiao = pi - jiaodu
        Else
            fangweijiao = -pi - jiaodu
        End If
    End If
'    MsgBox fangweijiao
End Function

Private Sub CommandButton3_Click()
    Unload UserForm2
    Unload Me
    
End Sub

Private Sub CommandButton4_Click()
    Dim sFile As String
    With dlgCommonDialog
        .DialogTitle = "打开地面线数据"
        .CancelError = False
        .FileName = ""
        'ToDo: 设置 common dialog 控件的标志和属性
        .Filter = "纬地地面线文件 (*.dmx)|*.DMX|所有文件(*.*)|*.*"
        .ShowOpen
        If Len(.FileName) = 0 Then
            Exit Sub
        End If
        sFile = .FileName
    End With
    'ToDo: 添加处理打开的文件的代码
If sFile = "" Then
Else:
    Open sFile For Input As #10
    Dim temp As String
    Dim i As Integer
    i = 0
    Input #10, temp
    Do While Not EOF(10)
        ReDim Preserve dmxd(i + 1)
        Input #10, dmxd(i).lc, dmxd(i).bg
'        MsgBox dmxd(i).lc
'        MsgBox dmxd(i).bg
        If dmxd(i).lc = 0 And dmxd(i).bg = 0 Then Exit Do
        i = i + 1
    Loop
    Close #10
    dmxnum = i
    MsgBox "读取成功"
End If




End Sub

Private Sub CommandButton5_Click()
    savedir = SelectDir("c:\", "选择目标文件夹")
    UserForm1.Hide
    UserForm2.Show
End Sub


Private Sub CommandButton6_Click()
    Dim numtemp As Integer
    Dim filenam As String
    Dim qdtemp As Double, zdtemp As Double
    Dim i As Integer
    Dim dianshu As Integer '统计每张图里的点数
    Dim l As Double         '记录地面线点序号
    Dim fnam As Integer
    Dim R As String
    Dim pointss As String
    Dim sum As Double       '记录总和，用于计算平均值
    
    
    
    
    
    
    '分图，确定分图数量
    numtemp = Fix((enddmx - startdmx) / 1000) + 1
    qdtemp = startdmx
    zdtemp = qdtemp + 1000
    j = 1       '用j记录第几张图
    
    l = 0
    Do While dmxd(l).lc < startdmx        '此段为丢弃多余的地面线数据
        l = l + 1
    Loop
    
    For j = 1 To numtemp
    
    Open savedir & "temp" For Output As #10
    
   '此段为写入分图的起点的数据，内插法进行
            Print #10, "Point:"
            Write #10, qdtemp
            Write #10, (dmxd(l).bg - dmxd(l - 1).bg) / (dmxd(l).lc - dmxd(l - 1).lc) * (qdtemp - dmxd(l - 1).lc) + dmxd(l - 1).bg
            sum = (dmxd(l).bg - dmxd(l - 1).bg) / (dmxd(l).lc - dmxd(l - 1).lc) * (qdtemp - dmxd(l - 1).lc) + dmxd(l - 1).bg
            dianshu = 1
    '此段为写入分图的起点的数据，内插法进行
    
    '正式开始写地面线数据
        Do
        Print #10, "Point:"
        Write #10, dmxd(l).lc
        Write #10, dmxd(l).bg
        sum = sum + dmxd(l).bg
        dianshu = dianshu + 1
        If l + 1 = dmxnum Then GoTo hereok      '如果数据全部读完了，就跳出循环，直接完成最后一个数据文件
        l = l + 1
        Loop Until Not dmxd(l).lc < zdtemp
    
    '写最后一个数据，即每张图的终点
            Print #10, "Point:"
            Write #10, zdtemp
            Write #10, (dmxd(l).bg - dmxd(l - 1).bg) / (dmxd(l).lc - dmxd(l - 1).lc) * (zdtemp - dmxd(l - 1).lc) + dmxd(l - 1).bg
            sum = sum + (dmxd(l).bg - dmxd(l - 1).bg) / (dmxd(l).lc - dmxd(l - 1).lc) * (zdtemp - dmxd(l - 1).lc) + dmxd(l - 1).bg
            dianshu = dianshu + 1
    '写最后一个数据，即每张图的终点
        
        Close #10
        Open savedir & "temp" For Input As #20
        Open savedir & "\C" & j For Output As #11
            Write #11, 1
            Write #11, 200
            Write #11, 2000
            Write #11, Fix(sum / dianshu / 2) * 2 - 6
            Write #11, dianshu
            
            Do While Not EOF(20)        '将数据从临时文件里写入到单个文件
                Line Input #20, R
                Print #11, R
            Loop
            Close #11
            Close #20
            
            qdtemp = qdtemp + 1000      '推进窗格
            If zdtemp + 1000 < enddmx Then
            zdtemp = zdtemp + 1000
            Else
            zdtemp = enddmx
            End If
     Next
        
        
hereok:                                         '如果数据不够而退出循环后，完成最后一个数据文件

        If Not j > numtemp Then
        

    '写最后一张图的最后一个数据
            Print #10, "Point:"
            Write #10, dmxd(l).lc
            Write #10, dmxd(l).bg
            dianshu = dianshu + 1
    '写最后一张图的最后一个数据
        
        Close #10
        Open savedir & "temp" For Input As #20
        Open savedir & "\C" & j For Output As #11
            Write #11, 1
            Write #11, 200
            Write #11, 2000
            Write #11, Fix(sum / dianshu / 2) * 2 - 6
            Write #11, dianshu
            
            Do While Not EOF(20)        '将数据从临时文件里写入到单个文件
                Line Input #20, R
                Print #11, R
            Loop
            Close #11
            Close #20
            
        Else
        End If
        

Kill savedir & "temp"
MsgBox "文件生成成功"
End Sub

Private Sub CommandButton7_Click()
    CommandButton2_Click
End Sub

Private Sub CommandButton8_Click()
    Dim sFile As String
If TextBox2.Value <> "请选择要打开的纬地纵断面设计数据" And TextBox2.Value <> "" Then
    sFile = TextBox2.Value
Else
    With dlgCommonDialog
        .DialogTitle = "打开"
        .CancelError = False
        .FileName = ""
        'ToDo: 设置 common dialog 控件的标志和属性
        .Filter = "纬地纵断面设计数据文件 (*.zdm)|*.zdm|所有文件(*.*)|*.*"
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
    TextBox2.Value = sFile
    
    Dim strtemp As String
    Dim i As Integer
    Dim noway As String
    Open sFile For Input As #20
    Input #20, strtemp
    Input #20, zdmnum
    ReDim Preserve zdmd(zdmnum)
    For i = 0 To zdmnum - 1
'    For i = 0 To 5
        Input #20, zdmd(i).lc, zdmd(i).bg, zdmd(i).xs, noway
'        MsgBox zdmd(i).lc
'        MsgBox zdmd(i).bg
'        MsgBox zdmd(i).xs
    Next
End If
Close #20
MsgBox "纵断面数据读取成功！"
End Sub

Private Sub CommandButton9_Click()
   Dim sFile As String
If TextBox3.Value <> "请选择要保存的文件路径与名称" And TextBox3.Value <> "" Then
    sFile = TextBox3.Value
Else
    With dlgCommonDialog
        .DialogTitle = "打开"
        .CancelError = False
        .FileName = "S1"
        'ToDo: 设置 common dialog 控件的标志和属性
        .Filter = "数据文档 (*.*)|*.*"
        .ShowOpen
        If Len(.FileName) = 0 Then
            Exit Sub
        End If
        sFile = .FileName
    End With
End If
    'ToDo: 添加处理打开的文件的代码
If sFile = "S1" Then sFile = ""
If sFile = "" Then
Else:
    TextBox3.Value = sFile
    wanguanR = TextBox4.Value
        
    wantouR = TextBox5.Value
    pingtanR = 650
    bianpodianshuju (sFile)
End If

End Sub

Private Function toshit(shit As Double)
'    Dim toshit As String
    Dim temp1 As Integer, temp2 As Double
    temp1 = Fix(shit)
    temp2 = Round(60 * (shit - temp1), 4)
    If temp2 < 10 Then
        toshit = temp1 & ".0" & temp2 * 10000
    Else
        toshit = temp1 & "." & temp2 * 10000
    End If
End Function


Private Sub bianpodianshuju(filepa As String)
    Dim p As Double, z As Double
    Dim bgtemp As Double
    p = 0
    z = 0
    
    On Error Resume Next
    
    
    Do While jiaodian(p).lc < zdmd(z).lc        '此段为丢弃多余的平面交点数据
        If Not p + 1 = imax Then p = p + 1
    Loop
    
    '从这里开始正式开始写数据
    Open filepa For Output As #10
    
    Do
        If jiaodian(p).lc > zdmd(z).lc - 2 And jiaodian(p).lc < zdmd(z).lc + 2 Then        '纵断面与平面点重合
            Select Case jiaodian(p).xs
                Case 0
                If p <> 0 Then
                    MsgBox jiaodian(p).lc & "此处平弹与纵向变坡点发生了组合,将以纵断面形式为准！"
                    jiaodian(p).xs = zdmd(z).xs
'                    Exit Sub
                Else
                    Print #10, "ST"
                    Write #10, zdmd(z).lc
                    Write #10, zdmd(z).bg
                    Print #10, 0
                    z = z + 1
                    If Not p + 1 = imax Then p = p + 1
                End If
                
                Case 1                                   '1为弯管
                
                Print #10, "W"
                Write #10, jiaodian(p).lc
                Write #10, zdmd(z).bg
                Print #10, wanguanR
                Print #10, toshit(jiaodian(p).zj)
                If Abs((zdmd(z).bg - zdmd(z - 1).bg) / (jiaodian(p).lc - zdmd(z - 1).lc)) > 0.14 Then MsgBox jiaodian(p).lc & "处平纵合并之后前坡大于14%！"
                If Abs((zdmd(z).bg - zdmd(z + 1).bg) / (jiaodian(p).lc - zdmd(z + 1).lc)) > 0.14 Then MsgBox jiaodian(p).lc & "处平纵合并之后后坡大于14%！"
                
                z = z + 1
                If Not p + 1 = imax Then p = p + 1
                
                Case 2                                    '2为弯头，暂时将两个都按照一样的处理
                Print #10, "W"
                Write #10, jiaodian(p).lc
                Write #10, zdmd(z).bg
                Print #10, wantouR
                Print #10, toshit(jiaodian(p).zj)
                If Abs((zdmd(z).bg - zdmd(z - 1).bg) / (jiaodian(p).lc - zdmd(z - 1).lc)) > 0.14 Then MsgBox jiaodian(p).lc & "处平纵合并之后前坡大于14%！"
                If Abs((zdmd(z).bg - zdmd(z + 1).bg) / (jiaodian(p).lc - zdmd(z + 1).lc)) > 0.14 Then MsgBox jiaodian(p).lc & "处平纵合并之后后坡大于14%！"
                z = z + 1
                If Not p + 1 = imax Then p = p + 1
            End Select
        
        
        ElseIf jiaodian(p).lc < zdmd(z).lc - 2 Then        '平面点小于纵断面点，计算并写入平面点
            bgtemp = (zdmd(z).bg - zdmd(z - 1).bg) / (zdmd(z).lc - zdmd(z - 1).lc) * (jiaodian(p).lc - zdmd(z - 1).lc) + zdmd(z - 1).bg         '计算平面点的设计标高
            
            
            Select Case jiaodian(p).xs
                Case 0                                   '0为平弹
                    Print #10, "PT"
                    Write #10, jiaodian(p).lc
                    Write #10, bgtemp
                    Write #10, pingtanR
                    Print #10, toshit(jiaodian(p).zj)
                    If Not p + 1 = imax Then p = p + 1
                
                Case 1                                   '1为弯管
                    Print #10, "W"
                    Write #10, jiaodian(p).lc
                    Write #10, bgtemp
                    Print #10, wanguanR
                    Print #10, toshit(jiaodian(p).zj)
                    If Not p + 1 = imax Then p = p + 1
                
                
                
                Case 2                                   '2为弯头
                    Print #10, "W"
                    Write #10, jiaodian(p).lc
                    Write #10, bgtemp
                    Print #10, wantouR
                    Print #10, toshit(jiaodian(p).zj)
                    If Not p + 1 = imax Then p = p + 1
                
            End Select
            
            
        Else                                             '平面点大于纵断面点，写纵断面点
            Select Case zdmd(z).xs                       '按照纵断面形式选择写入内容
            Case 0                                       '0为竖弹
                Print #10, "ST"
                Write #10, zdmd(z).lc
                Write #10, zdmd(z).bg
                Print #10, 0

            Case 1                                       '1为弯管
                Print #10, "W"
                Write #10, zdmd(z).lc
                Write #10, zdmd(z).bg
                Print #10, wanguanR
                Print #10, 180

                
                
            Case 2                                       '2为弯头
                Print #10, "W"
                Write #10, zdmd(z).lc
                Write #10, zdmd(z).bg
                Print #10, wantouR
                Print #10, 180
                
            End Select
            If z > 1 Then
            If Abs((zdmd(z).bg - zdmd(z - 1).bg) / (zdmd(z).lc - zdmd(z - 1).lc)) > 0.14 Then MsgBox zdmd(z).lc & "处前坡大于14%！"
'            If Abs((zdmd(z).bg - zdmd(z + 1).bg) / (zdmd(z).lc - zdmd(z + 1).lc)) > 0.14 Then MsgBox jiaodian(p).lc & "处后坡大于14%！"
            End If
            z = z + 1
        End If
        
        
        If p + 1 = imax Then jiaodian(p).lc = 1000000
        
    Loop Until z = zdmnum
    
    
    
Close #10
MsgBox "文件" & filepa & "已生成！"
    
    
End Sub

Private Sub TextBox6_Change()
    Select Case TextBox6.Value
    Case 610
        TextBox4.Value = 24.4
        TextBox5.Value = 3.7
    Case 559
        TextBox4.Value = 22.4
        TextBox5.Value = 3.4
    Case 323
        TextBox4.Value = 13
        TextBox5.Value = 2
    Case 273
        TextBox4.Value = 13
        TextBox5.Value = 2
 '   Default
    End Select
    
End Sub
