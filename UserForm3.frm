VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm3 
   Caption         =   "UserForm3"
   ClientHeight    =   3088
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7500
   OleObjectBlob   =   "UserForm3.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "UserForm3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
#If VBA7 And Win64 Then
    Private Declare PtrSafe Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
    Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    Private Declare PtrSafe Function ClipCursor Lib "user32" (lpRect As Any) As Long
    Private Declare PtrSafe Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
    Private Declare PtrSafe Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" ( _
        ByVal pCaller As Long, _
        ByVal szURL As String, _
        ByVal szFileName As String, _
        ByVal dwReserved As Long, _
        ByVal lpfnCB As Long _
        ) As Long
    Private Declare PtrSafe Function DeleteUrlCacheEntry Lib "wininet.dll" Alias "DeleteUrlCacheEntryA" (ByVal lpszUrlName As String) As Long
#Else
    Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
    Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    Private Declare Function ClipCursor Lib "user32" (lpRect As Any) As Long
    Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
    Private Declare Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" ( _
        ByVal pCaller As Long, _
        ByVal szURL As String, _
        ByVal szFileName As String, _
        ByVal dwReserved As Long, _
        ByVal lpfnCB As Long _
        ) As Long
    Private Declare Function DeleteUrlCacheEntry Lib "wininet.dll" Alias "DeleteUrlCacheEntryA" (ByVal lpszUrlName As String) As Long
#End If
Public kk As Integer
Public sr1 As String
Public zong1 As Integer 'zscr这次输入，zong1为旧值

Sub shuzu()
    Dim zan(3, 2) As Variant
    Dim ming As String
    Dim i As Integer
    
    i = 0
    zan(1, 1) = 1
    zan(1, 2) = wo
    zan(2, 1) = 1000
    zan(2, 2) = "s"
    zan(3, 1) = 201
    zan(3, 2) = er
    For X = 1 To 3
        If zan(X, 1) > i Then
            i = zan(X, 1)
            ix = X
            ming = zan(ix, 2)
        End If
    Next
    Debug.Print i
    Debug.Print ix
    Debug.Print zan(ix, 2)
    
    'c = max(zan(i, 1))
End Sub
Sub duqu()       '读取最大值
    Dim ming, fen() As String
    Dim i, zhong, xu, zan() As Integer
    Dim h As Long
    
    Dim mingz(), str() As Variant
    
    zhong = 0
    Open "C:\bbdata\dz.txt" For Input As #1
    Do While Not EOF(1)
        h = h + 1       '行数
        ReDim Preserve str(h)
        
        Line Input #1, str(h)  '读取原始行
    Loop
    Close #1
    ReDim zan(h)
    ReDim mingz(h)
    
    For i = 1 To h
        fen = Split(str(i), " ")
        zan(i) = fen(0)
        mingz(i) = fen(1)
        'Debug.Print zan(i)
        If zan(i) > zhong Then
            'Debug.Print zhong
            zhong = zan(i)
            xu = i
        End If
    Next
    
    ming = mingz(xu)
    Debug.Print ming & zhong & xu
    
    
    
End Sub
Sub chushi() '读取自己的值，或者创建空值，找最大值
    
End Sub
Sub butgm() '改名字
    
End Sub

Sub dz()
    
    
End Sub

Private Sub CommandButton1_Click() '换名字
    Dim sr, xr As String   'sr是输入名字,xr是写入text  'UserForm3.textnc.text昵称 UserForm3.labelzan.caption赞数标签
    
    Dim i, zhong, xu, xuda, zan() As Integer
    Dim h As Long
    Dim mingz(), str(), cp() As Variant
    
    '读取，查找是否有名字
    zhong = 0
    sr = UserForm3.Textnc.Text   '读取输入名字
    Open "C:\bbdata\dz.txt" For Input As #1
    Do While Not EOF(1)
        h = h + 1       '行数
        ReDim Preserve str(h)
        Line Input #1, str(h)  '读取原始行
    Loop
    Close #1
    ReDim zan(h)
    ReDim mingz(h)
    ReDim cp(h)
    For i = 1 To h
        zan(i) = Split(str(i), " ", 3)(0)
        mingz(i) = Split(str(i), " ", 3)(1)
        cp(i) = Split(str(i), " ", 3)(2)
        'Debug.Print zan(i)
        If cp(i) = Environ("computername") Then
            '找旧名字
            mingz(i) = sr    '替换
            you = 1
            ' zhong = zan(i)  '该名字的赞数
            xu = i            '序号
            'UserForm3.Labelzan.Caption = zhong '显示已有赞数
            'Exit For
        End If
        
        
    Next
    
    For i = 1 To h
        If i = 1 Then
            Open "C:\bbdata\dz.txt" For Output As #1
            xr = zan(i) & " " & mingz(i) & " " & cp(i)                                '输出
            Print #1, xr
            Close #1
        Else
            
            Open "C:\bbdata\dz.txt" For Append As #1
            xr = zan(i) & " " & mingz(i) & " " & cp(i)                                '输出
            Print #1, xr
            Close #1
        End If
    Next
    
    Open "C:\bbdata\ncg.txt" For Output As #1
    xr = UserForm3.Textnc.Text                                 '输出nc
    Print #1, xr
    Close #1
    NewMacros.duqu
    
End Sub

Private Sub CommandButton2_Click()
    '读取最大值和总数，差值
    Dim ming As String
    Dim i, zhong, zong, xu, zan() As Integer
    Dim h As Long
    
    Dim mingz(), str(), fen As Variant
    
    
    lngReturn = URLDownloadToFile(0, "http://39.103.174.58/sever/05/dz.txt", "C:\bbdata\dz.txt", 0, 0) '查看最新版本
    DeleteUrlCacheEntry "http://39.103.174.58/sever/05/dz.txt" '清除缓存
    'If lngReturn = 0 Then
    
    '  Else
    
    '      MsgBox "服务器连接失败"
    '      Exit Sub
    ' End If
    zhong = 0
    Open "C:\bbdata\dz.txt" For Input As #1
    Do While Not EOF(1)
        h = h + 1       '行数
        ReDim Preserve str(h)
        
        Line Input #1, str(h)  '读取原始行
    Loop
    Close #1
    ReDim zan(h)
    ReDim mingz(h)
    ReDim cp(h)
    For i = 1 To h
        zan(i) = Split(str(i), " ", 3)(0)
        mingz(i) = Split(str(i), " ", 3)(1)
        cp(i) = Split(str(i), " ", 3)(2)
        'Debug.Print zan(i)
        If zan(i) > zhong Then
            'Debug.Print zhong
            zhong = zan(i)
            xu = i
        End If
        zong = zong + zan(i)
    Next
    
    ming = mingz(xu)
    UserForm3.Labelming.Caption = "    " & mingz(xu)
    UserForm3.Labelda.Caption = "同学总共点了" & zhong & "个赞！！"
    UserForm3.Labelzong.Caption = zong
    UserForm2.Label4.Caption = zong
    
    'Debug.Print zong
    
    
    
    
    
End Sub

Private Sub Label5_Click()  'dz
    Dim ming, sr, xr As String  'sr是输入名字,xr是写入dztext  'UserForm3.textnc.text UserForm3.labelzan.caption
    
    Dim i, zhong, xu, zan() As Integer
    Dim h As Long
    Dim mingz(), str(), cp(), fen As Variant
    
    UserForm3.Labelzan.Caption = UserForm3.Labelzan.Caption + 1
    UserForm2.zcsr = UserForm2.zcsr + 1
    
    '读取，查找是否有名字
    sr = UserForm3.Textnc.Text   '读取输入名字
    
    Open "C:\bbdata\dz.txt" For Input As #1
    Do While Not EOF(1)
        h = h + 1       '行数
        ReDim Preserve str(h)
        
        Line Input #1, str(h)  '读取原始行
    Loop
    Close #1
    ReDim zan(h)
    ReDim mingz(h)
    ReDim cp(h)
    For i = 1 To h
        zan(i) = Split(str(i), " ", 3)(0)
        mingz(i) = Split(str(i), " ", 3)(1)
        cp(i) = Split(str(i), " ", 3)(2)
        'Debug.Print zan(i)
        If mingz(i) = sr Then
            '找到名字啦
            zhong = zan(i)  '该名字的赞数
            xu = i            '序号
        End If
    Next
    zan(xu) = UserForm3.Labelzan.Caption
    For i = 1 To h
        If i = 1 Then
            Open "C:\bbdata\dz.txt" For Output As #1
            xr = zan(i) & " " & mingz(i) & " " & cp(i)                                '输出
            Print #1, xr
            Close #1
        Else
            
            Open "C:\bbdata\dz.txt" For Append As #1
            xr = zan(i) & " " & mingz(i) & " " & cp(i)                               '输出
            Print #1, xr
            Close #1
        End If
    Next
    If UserForm2.zcsr = 10 Then
        MsgBox "啊！我太幸福了！", vbOKOnly, "来自张同学"
    ElseIf UserForm2.zcsr = 12 Then
        MsgBox "我此刻就像是二十个大海的主人，", vbOKOnly, "来自张同学"
    ElseIf UserForm2.zcsr = 15 Then
        MsgBox "它的每一粒泥沙都是珠玉，每一滴海水都是天上的琼浆。", vbOKOnly, "来自张同学"
    ElseIf UserForm2.zcsr = 22 Then
        MsgBox "最芬芳的花蕾中有蜜蜂，", vbOKOnly, "来自张同学"
    ElseIf UserForm2.zcsr = 25 Then
        MsgBox "最美丽的人的心里，才会有如此令人难以承受的赞赏", vbOKOnly, "来自张同学"
    ElseIf UserForm2.zcsr = 30 Then
        MsgBox "照耀万物的太阳，", vbOKOnly, "啊！我太幸福了！"
    ElseIf UserForm2.zcsr = 32 Then
        MsgBox "自有天地以来也不曾见过一个可以和你媲美的人！", vbOKOnly, "来自张同学"
    ElseIf UserForm2.zcsr = 35 Then
        MsgBox "容我直言，", vbOKOnly, "来自张同学"
    ElseIf UserForm2.zcsr = 38 Then
        MsgBox "谁见了天仙一般的你，不会像一个野蛮的印度人，", vbOKOnly, "来自张同学"
    ElseIf UserForm2.zcsr = 39 Then
        MsgBox "当东方的朝阳开始呈现他的绮丽，俯首拜服，用他虔诚的胸膛贴敷土地？", vbOKOnly, "来自张同学"
    ElseIf UserForm2.zcsr = 40 Then
        MsgBox "哪一道鹰隼般威凌闪闪的眼光，不会眩耀于你的美丽，敢仰望你眉宇间的天堂？", vbOKOnly, "来自张同学"
    ElseIf UserForm2.zcsr = 50 Then
        MsgBox "说真的，", vbOKOnly, "来自张同学"
    ElseIf UserForm2.zcsr = 53 Then
        MsgBox "倘不是为了你，白昼都要失去他的光亮", vbOKOnly, "来自张同学"
    ElseIf UserForm2.zcsr = 68 Then
        MsgBox "但是说到底，", vbOKOnly, "来自张同学"
    ElseIf UserForm2.zcsr = 70 Then
        MsgBox "任何赞美，都比不上你自身的美妙", vbOKOnly, "来自张同学"
    ElseIf UserForm2.zcsr = 76 Then
        MsgBox "不要再让我夸你啦!", vbOKOnly, "来自张同学"
    ElseIf UserForm2.zcsr = 79 Then
        MsgBox "出售单身舍友（各种款式），有需要务必联系我哈哈哈", vbOKOnly, "来自张同学"
    End If
    UserForm2.Label7.Caption = "+1"
    UserForm3.Label7.Caption = "+1"
    
    
    NewMacros.duqu
    
End Sub


Private Sub UserForm_Activate()
    
    Dim bb As String
    
    kk = 1
    
    Labelming.WordWrap = False
    
    bb = Labelming
    
    
    
    Do
        
        If kk = 2 Then Exit Sub
        
        '      bb = Right(bb, 1) & Left(bb, Len(bb) - 1)
        
        '        bb = Right(bb, Len(bb) - 1) & Left(bb, 1)
        
        bb = Right(bb, Len(bb) - 1) & Left(bb, 1)
        
        Labelming = bb
        
        
        
        vv = Timer
        
        Do While Timer < vv + 0.2
            
            DoEvents
            
        Loop
        
        
        
    Loop
    
    
    
End Sub

Private Sub UserForm_Initialize()
    Dim sr, xr As String  'sr是输入名字,xr是写入text  'UserForm3.textnc.text昵称 UserForm3.labelzan.caption赞数标签
    
    Dim i, zhong, zong1, xu, xuda, you As Integer
    Dim zan(), da, h As Long
    Dim mingz(), str(), cp(), fen As Variant
    If Dir("c:\bbdata\ncg.txt") = "" Then   '有没有改过名称
        Open "C:\bbdata\ncg.txt" For Output As #1
        xr = Environ("computername")                                 '输出
        Print #1, xr
        Close #1
        Textnc.Text = Environ("computername")
        sr1 = Environ("computername")
    Else
        Open "C:\bbdata\ncg.txt" For Input As #1
        Do While Not EOF(1)
            Line Input #1, sr1  '读取原始行
        Loop
        Close #1
        Textnc.Text = sr1
    End If
    
    
    '读取，查找是否有名字
    sr = UserForm3.Textnc.Text   '读取输入名字
    
    you = 0    'you为1即为已经有名字
    da = 0     '最大值
    Open "C:\bbdata\dz.txt" For Input As #1
    Do While Not EOF(1)
        h = h + 1       '行数
        ReDim Preserve str(h)
        Line Input #1, str(h)  '读取原始行
    Loop
    Close #1
    ReDim zan(h)
    ReDim mingz(h)
    ReDim cp(h)
    For i = 1 To h
        zan(i) = Split(str(i), " ", 3)(0)
        mingz(i) = Split(str(i), " ", 3)(1)
        cp(i) = Split(str(i), " ", 3)(2)
        'Debug.Print zan(i)
        If mingz(i) = sr Then
            '找名字
            you = 1
            zhong = zan(i)  '该名字的赞数
            xu = i            '序号
            UserForm3.Labelzan.Caption = zhong '显示已有赞数
            Debug.Print zhong   'zhong是已有值
            'Exit For
        End If
        ' Debug.Print da
        'Debug.Print zan(i)
        
        'If zan(i) > da Then     '找最大值
        '     'Debug.Print zhong
        '     da = zan(i)
        '     xuda = i
        ' End If
        
    Next
    If you <> 1 Then     '如果没有找到名字
        UserForm3.Labelzan.Caption = 0
        Open "C:\bbdata\dz.txt" For Append As #1
        xr = 0 & " " & sr & " " & Environ("computername")                                    '输出
        Print #1, xr
        Close #1
    End If
    NewMacros.duqu
    Label7.Caption = "+" & Labelzong.Caption - UserForm2.zong1
End Sub


Private Sub UserForm_Terminate()
    kk = 2
    'Shell ("cmd /c C:\bbdata\dz.bat"), vbHide
End Sub
