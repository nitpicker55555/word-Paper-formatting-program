VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm2 
   Caption         =   "版本"
   ClientHeight    =   2824
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   5670
   OleObjectBlob   =   "UserForm2.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "UserForm2"
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
Dim gs, tzz As String
Public kk As Integer
Public zong1, zcsr As Integer

Private Sub CommandButton1_Click()
    
    UserForm4.Show
    Open "C:\bbdata\" & UserForm1.Label21.Caption & "\cishu.txt" For Append As #1
    gs = "留言"                                       'wuyiyi
    Print #1, gs
    Close #1
    
End Sub

Private Sub gengxin_Click()
    
End Sub

Private Sub CommandButton2_Click()
    Open "C:\bbdata\" & UserForm1.Label21.Caption & "\cishu.txt" For Append As #1
    gs = "更新"                                       'wuyiyi
    Print #1, gs
    Close #1
    
    lngReturn = URLDownloadToFile(0, "http://39.103.174.58/sever/05/banben.txt", "C:\bbdata\bb.txt", 0, 0)
    DeleteUrlCacheEntry "http://39.103.174.58/sever/05/banben.txt" '清除缓存
    
    
    ' 注意：URLDownloadToFile函数返回0表示文件下载成功
    
    '判断返回的结果是否为0,则返回True，否则返回False
    
    If lngReturn = 0 Then
        
    Else
        
        MsgBox "服务器连接失败"
        Exit Sub
    End If
    
    Open "C:\bbdata\bb.txt" For Input As #1
    Do While Not EOF(1)
        Line Input #1, all  '从已打开的顺序文件中读出一行并将它分配给 String 变量
        'Line Input # 语句一次只从文件中读出一个字符，直到遇到回车符 (Chr(13))
        '或回车–换行符 (Chr(13) + Chr(10)) 为止。回车–换行符将被跳过，而不会被附加到字符串上
        Debug.Print all
    Loop
    Close #1
    If all > UserForm1.Label20.Caption Then                                         '目前的版本号
        lngReturn = URLDownloadToFile(0, "http://39.103.174.58/sever/05/bbn.txt", "C:\bbdata\bbn.txt", 0, 0)
        DeleteUrlCacheEntry "http://39.103.174.58/sever/05/bbn.txt" '清除缓存
        
        Open "C:\bbdata\bbn.txt" For Input As #1
        Do While Not EOF(1)
            Line Input #1, bbn  '从已打开的顺序文件中读出一行并将它分配给 String 变量
            'Line Input # 语句一次只从文件中读出一个字符，直到遇到回车符 (Chr(13))
            '或回车–换行符 (Chr(13) + Chr(10)) 为止。回车–换行符将被跳过，而不会被附加到字符串上
            Debug.Print bbn
        Loop
        Close #1
        If MsgBox("" & bbn, vbOKCancel, "版本" & all & "的程序更新") = vbOK Then
            lngReturn2 = URLDownloadToFile(0, "http://39.103.174.58/sever/05/pb.docm", "" & ThisDocument.Path & "\建大自动排版" & all & ".docm", 0, 0)
            MsgBox "版本为" & all & "的新的程序已经下载至同一文件夹中！旧程序不会自动删除。", vbOKOnly, "来自张同学"
        End If
    Else
        MsgBox "你的版本是最新的哦", vbOKOnly, "来自张同学"
        
    End If
    
    If lngReturn2 = 0 Then
        
    Else
        
        MsgBox "服务器连接失败"
        
    End If
    
End Sub



Private Sub CommandButton3_Click()
    
    lngReturn = URLDownloadToFile(0, "http://39.103.174.58/sever/05/tzz.txt", "C:\bbdata\tzz.txt", 0, 0)
    DeleteUrlCacheEntry "http://39.103.174.58/sever/05/tzz.txt" '清除缓存
    
    Open "C:\bbdata\tzz.txt" For Input As #1
    Do While Not EOF(1)
        Line Input #1, tzz
        Debug.Print tzz
    Loop
    Close #1
    Open "C:\bbdata\" & UserForm1.Label21.Caption & "\cishu.txt" For Append As #1
    gs = tzz & "_查看消息"                                       'wuyiyi
    Print #1, gs
    Close #1
    
    MsgBox "" & tzz, vbOKOnly, "来自张同学的消息"
    
End Sub

Private Sub CommandButton4_Click()
    UserForm3.Show
    Open "C:\bbdata\" & UserForm1.Label21.Caption & "\cishu.txt" For Append As #1
    gs = "小彩蛋"                                       'wuyiyi
    Print #1, gs
    Close #1
End Sub

Private Sub Image1_BeforeDragOver(ByVal Cancel As MSForms.ReturnBoolean, ByVal Data As MSForms.DataObject, ByVal X As Single, ByVal Y As Single, ByVal DragState As MSForms.fmDragState, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)
    
End Sub

Private Sub Image1_click()
    
End Sub

Private Sub Label3_Click()
    
End Sub

Private Sub Label5_Click()
    Dim ming, sr, xr As String  'sr是输入名字,xr是写入dztext  'UserForm3.textnc.text UserForm3.labelzan.caption
    
    Dim i, zhong, xu, zan() As Integer
    Dim h As Long
    Dim mingz(), str(), cp(), fen As Variant
    
    UserForm3.Labelzan.Caption = UserForm3.Labelzan.Caption + 1
    zcsr = zcsr + 1
    
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
    If zcsr = 10 Then
        MsgBox "啊！我太幸福了！", vbOKOnly, "来自张同学"
    ElseIf zcsr = 12 Then
        MsgBox "我此刻就像是二十个大海的主人，", vbOKOnly, "来自张同学"
    ElseIf zcsr = 15 Then
        MsgBox "它的每一粒泥沙都是珠玉，每一滴海水都是天上的琼浆。", vbOKOnly, "来自张同学"
    ElseIf zcsr = 22 Then
        MsgBox "最芬芳的花蕾中有蜜蜂，", vbOKOnly, "来自张同学"
    ElseIf zcsr = 25 Then
        MsgBox "最美丽的人的心里，才会有如此令人难以承受的赞赏", vbOKOnly, "来自张同学"
    ElseIf zcsr = 30 Then
        MsgBox "照耀万物的太阳，", vbOKOnly, "啊！我太幸福了！"
    ElseIf zcsr = 32 Then
        MsgBox "自有天地以来也不曾见过一个可以和你媲美的人！", vbOKOnly, "来自张同学"
    ElseIf zcsr = 35 Then
        MsgBox "容我直言，", vbOKOnly, "来自张同学"
    ElseIf zcsr = 38 Then
        MsgBox "谁见了天仙一般的你，不会像一个野蛮的印度人，", vbOKOnly, "来自张同学"
    ElseIf zcsr = 39 Then
        MsgBox "当东方的朝阳开始呈现他的绮丽，俯首拜服，用他虔诚的胸膛贴敷土地？", vbOKOnly, "来自张同学"
    ElseIf zcsr = 40 Then
        MsgBox "哪一道鹰隼般威凌闪闪的眼光，不会眩耀于你的美丽，敢仰望你眉宇间的天堂？", vbOKOnly, "来自张同学"
    ElseIf zcsr = 50 Then
        MsgBox "说真的，", vbOKOnly, "来自张同学"
    ElseIf zcsr = 53 Then
        MsgBox "倘不是为了你，白昼都要失去他的光亮", vbOKOnly, "来自张同学"
    ElseIf zcsr = 68 Then
        MsgBox "但是说到底，", vbOKOnly, "来自张同学"
    ElseIf zcsr = 70 Then
        MsgBox "任何赞美，都比不上你自身的美妙", vbOKOnly, "来自张同学"
    ElseIf zcsr = 76 Then
        MsgBox "不要再让我夸你啦!", vbOKOnly, "来自张同学"
    ElseIf zcsr = 79 Then
        MsgBox "出售单身舍友（各种款式），有需要务必联系我哈哈哈", vbOKOnly, "来自张同学"
    End If
    UserForm2.Label7.Caption = "+1"
    UserForm3.Label7.Caption = "+1"
    
    
    NewMacros.duqu
    
    
    
End Sub


Private Sub Label8_Click()
    UserForm3.Show
    
End Sub

Private Sub UserForm_Click()
    
End Sub
Private Sub UserForm_Activate()
    
    Dim bb As String
    
    kk = 1
    
    Label8.WordWrap = True
    
    bb = Label8
    
    
    
    Do
        
        If kk = 2 Then Exit Sub
        
        '      bb = Right(bb, 1) & Left(bb, Len(bb) - 1)
        
        '        bb = Right(bb, Len(bb) - 1) & Left(bb, 1)
        
        bb = Right(bb, Len(bb) - 1) & Left(bb, 1)
        
        Label8 = bb
        Label8.Width = CommandButton4.Width
        
        
        
        vv = Timer
        
        Do While Timer < vv + 0.2
            
            DoEvents
            
        Loop
        
        
        
    Loop
    
    
    
End Sub


Private Sub UserForm_Error(ByVal Number As Integer, ByVal Description As MSForms.ReturnString, ByVal SCode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, ByVal CancelDisplay As MSForms.ReturnBoolean)
    
End Sub

Private Sub UserForm_Initialize()
    On Error Resume Next
    Dim all, t, tzz, ttz, dztz, dzztz, bat As String
    Dim tz, xr As String
    Dim ming As String
    Dim i, zhong, xu, zan() As Integer
    Dim h As Long
    
    Dim mingz(), str(), fen As Variant
    zcsr = 0
    Label8.Top = CommandButton4.Height / 2 + 10
    Label8.Width = CommandButton4.Width
    Label8.Left = CommandButton4.Left
    lngReturn = URLDownloadToFile(0, "http://39.103.174.58/sever/05/dz.txt", "C:\bbdata\dz.txt", 0, 0)
    DeleteUrlCacheEntry "http://39.103.174.58/sever/05/dz.txt" '清除缓存
    
    If Dir("C:\bbdata\dz.txt") = "" Then     '点赞文档
        Open "C:\bbdata\dz.txt" For Output As #1
        xr = ""                                '输出
        Print #1, xr
        Close #1
    End If
    Dim z As Integer
    ' On Error Resume Next
    Dim bb As Integer
    UserForm2.Caption = "版本：" & UserForm1.Label20.Caption
    If Dir("C:\bbdata\dz.bat") = "" Then   'dz上传
        Open "C:\bbdata\dz.bat" For Output As #1
        bat = "echo y | C:\bbdata\pscp.exe -l root -pw ""9aS0dF()"" C:\bbdata\dz.txt" & " " & "root@39.103.174.58:/sever/05/dz.txt"
        Print #1, bat
        Close #1
    End If
    If Dir("C:\bbdata\bb.txt") = "" Then
        MkDir "C:\bbdata" '创建文件夹
        Open "C:\bbdata\bb.txt" For Output As #1
        bb = UserForm1.Label20.Caption                                        'wuyiyi
        Write #1, bb
        Close #1
    End If
    If Dir("C:\bbdata\pscp.exe") = "" Then  '第一次huichuan
        lngReturn = URLDownloadToFile(0, "http://39.103.174.58/sever/05/231.21", "C:\bbdata\pscp.txt", 0, 0)
        Shell "cmd /c rename C:\bbdata\pscp.txt pscp.exe", vbHide
        MkDir "C:\bbdata\" & UserForm1.Label21.Caption '创建文件夹
    End If
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
        zong1 = zong1 + zan(i)       '旧值
    Next
    
    NewMacros.duqu
    UserForm2.Label7.Caption = "+" & UserForm2.Label4.Caption - zong1
    
End Sub

Private Sub UserForm_Terminate()
    kk = 2
    Shell ("cmd /c C:\bbdata\dz.bat"), vbHide
    
End Sub
