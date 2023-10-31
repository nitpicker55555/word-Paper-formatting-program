VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "建大论文排版程序"
   ClientHeight    =   9064.001
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   12120
   OleObjectBlob   =   "UserForm1.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "UserForm1"
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

Public dggs As Integer

Private Sub BtCheckBox_Change()
If BtCheckBox.Value = True Then BtTextBox.Enabled = True
If BtCheckBox.Value = False Then BtTextBox.Enabled = False

End Sub

Private Sub BtCheckBox_Click()

End Sub

Private Sub CheckBox13_Change()
    If CheckBox13.Value = True Then
    CheckBox13.Caption = "连字符连接"
    Else
    CheckBox13.Caption = "现在编号为圆点连接"

End If
End Sub

Private Sub CheckBox13_Click()
End Sub

Private Sub CheckBox14_Change()
 If CheckBox14.Value = True Then
    CheckBox14.Caption = "连字符连接"
    Else
    CheckBox14.Caption = "现在编号为圆点连接"

End If
End Sub

Private Sub CheckBox14_Click()
End Sub

Private Sub CheckBox15_Change()
 If CheckBox15.Value = True Then
    CheckBox15.Caption = "连字符连接"
    Else
    CheckBox15.Caption = "现在编号为圆点连接"

End If
End Sub

Private Sub CheckBox15_Click()

End Sub

Private Sub CheckBox4_Change()
    If CheckBox4.Value = True Then
        Text1.Enabled = True
    Else
        Text1.Enabled = False
    End If
End Sub

Private Sub CheckBox4_Click()
End Sub

Private Sub CheckBox6_Change()
    If CheckBox6.Value = True Then
        Text4.Enabled = True
    Else
        Text4.Enabled = False
    End If
End Sub

Private Sub CKB_Change()
    If CKB.Value = False Then
        CheckBox5.Enabled = False
        CheckBox6.Enabled = False
        Text4.Enabled = False
    Else
        CheckBox5.Enabled = True
        CheckBox6.Enabled = True
        Text4.Enabled = True
    End If
End Sub

Private Sub CKB_Click()
End Sub

Private Sub CommandButton1_Click()
    Dim gs As String
    CommandButton1.Enabled = False
    NewMacros.公式题注加对齐hhhhhhhh
    Open "C:\bbdata\" & UserForm1.Label21.Caption & "\cishu.txt" For Append As #1
    gs = "公式题注加对齐"                                         'wuyiyi
    Print #1, gs
    Close #1
    CommandButton1.Enabled = True
End Sub

Private Sub CommandButton10_Click()
    UserForm1.Width = CommandButton1.Left + CommandButton1.Width + 25
    CommandButton10.Visible = False
    CommandButton9.Visible = True
End Sub

Private Sub CommandButton11_Click()
    If Label18.Caption <> "gai" Then
    Else
        Kill "C:\data\*.*"
        RmDir "C:\data\"
        Label18.Caption = "meigai"
        If Dir("C:\data\") = "" Then Debug.Print "没有改动字体"
        gs = "建大公式字体"
        gszs = "建大公式注释字体"
        bg = "建大表格字体"
        tbtz = "建大图表题注字体"
        TextBox1.Text = gs
        TextBox2.Text = gszs
        TextBox3.Text = bg
        TextBox4.Text = tbtz
        Label19.Caption = "ok"
    End If
End Sub

Private Sub CommandButton12_Click()
    If CKWR.Value = CKQ.Value Then
        MsgBox "请选择公式格式！微软公式意为Word自带编辑器。如果有问题请在关于>留言区留言", vbOKOnly, "来自张同学"
    Else
        NewMacros.单个公式排版
        'dggs = dggs + 1
        ' Debug.Print dggs
    End If
End Sub

Private Sub UserForm_Initialize()
    On Error Resume Next
    dggs = 0
    UserForm1.Height = CommandButton7.Top + 70
    Dim WINWND As Long
    Application.EnableCancelKey = xlDisabled
    NewMacros.tu = 0
    Dim lngReturn, lngReturn2 As String
    Dim all, tzz, ttz, dztz, dzztz, bat, cs As String
    Dim tz As String
    WINWND = FindWindow(vbNullString, Application.Caption)
    SetWindowPos WINWND, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
    UserForm1.Width = CommandButton1.Left + CommandButton1.Width + 25
    Dim gszs, bg, gs, tbtz As String
    'UserForm1.Label22.Caption = UserForm1.Label20.Caption & "_" & ".txt"
    UserForm1.Label21.Caption = Environ("computername")  '文件夹标签
    If Dir("C:\data\字体更改.txt") = "" Then
        gs = "建大公式字体"
        gszs = "建大公式注释字体"
        bg = "建大表格字体"
        tbtz = "建大图表题注字体"
    Else
        Label18.Caption = "gai"
        If Dir("C:\data\公式字体.txt") <> "" Then
            Open "C:\data\公式字体.txt" For Input As #1
            Do While Not EOF(1)
                Input #1, gs
            Loop
            Debug.Print gs
            Close #1
        Else
            gs = "建大公式字体"
        End If
        If Dir("C:\data\公式注释字体.txt") <> "" Then
            Open "C:\data\公式注释字体.txt" For Input As #1
            Do While Not EOF(1)
                Input #1, gszs
            Loop
            Debug.Print gszs
            Close #1
        Else
            gszs = "建大公式注释字体"
        End If
        If Dir("C:\data\表格字体.txt") <> "" Then
            Open "C:\data\表格字体.txt" For Input As #1
            Do While Not EOF(1)
                Input #1, bg
            Loop
            Debug.Print bg
            Close #1
        Else
            bg = "建大表格字体"
        End If
        If Dir("C:\data\图表字体.txt") <> "" Then
            Open "C:\data\图表字体.txt" For Input As #1
            Do While Not EOF(1)
                Input #1, tbtz
            Loop
            Debug.Print tbtz
            Close #1
        Else
            tbtz = "建大图表题注字体"
        End If
    End If
    TextBox1.Text = gs
    TextBox2.Text = gszs
    TextBox3.Text = bg
    TextBox4.Text = tbtz
    Dim bb As Integer
    If Dir("C:\bbdata\bb.txt") = "" Then
        MkDir "C:\bbdata" '创建文件夹
        Open "C:\bbdata\bb.txt" For Output As #1
        bb = UserForm1.Label20.Caption                                         'wuyiyi
        Write #1, bb
        Close #1
    Else
    End If
 
    '用lngReturn接收返回的结果
    lngReturn = URLDownloadToFile(0, "http://39.103.174.58/sever/05/banben.txt", "C:\bbdata\bb.txt", 0, 0) '查看最新版本
    DeleteUrlCacheEntry "http://39.103.174.58/sever/05/banben.txt" '清除缓存
    ' 注意：URLDownloadToFile函数返回0表示文件下载成功
    '判断返回的结果是否为0,则返回True，否则返回False
    'If lngReturn = 0 Then
    'Else
    ' MsgBox "服务器连接失败"
    'End If
    Open "C:\bbdata\bb.txt" For Input As #1
    Do While Not EOF(1)
        Line Input #1, all  '从已打开的顺序文件中读出一行并将它分配给 String 变量
        'Line Input # 语句一次只从文件中读出一个字符，直到遇到回车符 (Chr(13))
        '或回车–换行符 (Chr(13) + Chr(10)) 为止。回车–换行符将被跳过，而不会被附加到字符串上
        Debug.Print all
    Loop
    Close #1
    If all > UserForm1.Label20.Caption Then                                         '目前的版本号
        Open "C:\bbdata\" & UserForm1.Label21.Caption & "\cishu.txt" For Append As #1
        gs = "检测到更新"                                         'wuyiyi
        Print #1, gs
        Close #1
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
            DeleteUrlCacheEntry "http://39.103.174.58/sever/05/pb.docm" '清除缓存
            MsgBox "版本为" & all & "的新的程序已经下载至同一文件夹中！旧程序不会自动删除。", vbOKOnly, "来自张同学"
        End If
    End If
    If Dir("C:\bbdata\" & UserForm1.Label20.Caption & ".txt") = "" Then '第一次
        ttz = 1
        Open "C:\bbdata\bat.bat" For Output As #1
        bat = "echo y | C:\bbdata\pscp.exe -l root -pw ""9aS0dF()"" -r C:\bbdata\" & UserForm1.Label21.Caption & " " & "root@39.103.174.58:/var/www/html/sever/04/"
        Print #1, bat
        Close #1
        MsgBox "嗨！感谢你使用" & UserForm1.Label20.Caption & "版自动排版程序！", vbOKOnly, "来自张同学"
        MsgBox "非常遗憾，由于兼容问题，WPS用户无法使用表格排版功能。如果有其他问题请及时联系我，我会尽快更新！非常感谢哦！", vbOKOnly, "来自张同学"
        Open "C:\bbdata\" & UserForm1.Label20.Caption & ".txt" For Output As #1
        bb = UserForm1.Label20.Caption                                         'wuyiyi
        Write #1, bb
        Close #1
    End If
    If Dir("C:\bbdata\tz.txt") = "" Then  '如果没有tz
        tzi = 1
        Open "C:\bbdata\tz.txt" For Output As #1
        tzi = 1                                        'wuyiyi
        Write #1, tzi                   '写一个tz=1
        Close #1
    End If
    Open "C:\bbdata\tz.txt" For Input As #1    '读取tz
    Do While Not EOF(1)
        Line Input #1, ttz
        Debug.Print ttz
    Loop
    Close #1
    'UserForm1.Label22.Caption = UserForm1.Label20.Caption & "_" & ttz & ".txt"
    lngReturn = URLDownloadToFile(0, "http://39.103.174.58/sever/05/tz.a", "C:\bbdata\tz.txt", 0, 0)
    DeleteUrlCacheEntry "http://39.103.174.58/sever/05/tz.a" '清除缓存
    Open "C:\bbdata\tz.txt" For Input As #1      '下载tz
    Do While Not EOF(1)
        Line Input #1, tz
        Debug.Print tz
    Loop
    Close #1
    If tz > ttz Then                          '有新tz
        ' UserForm1.Label22.Caption = UserForm1.Label20.Caption & "_" & tz & ".txt"
        'Shell "cmd /c rename C:\bbdata\" & Environ("computername") & "\" & UserForm1.Label20.Caption & ttz & ".txt " & UserForm1.Label22.Caption, vbHide
        lngReturn = URLDownloadToFile(0, "http://39.103.174.58/sever/05/tzz.txt", "C:\bbdata\tzz.txt", 0, 0)
        DeleteUrlCacheEntry "http://39.103.174.58/sever/05/tzz.txt" '清除缓存
        Open "C:\bbdata\tzz.txt" For Input As #1
        Do While Not EOF(1)
            Line Input #1, tzz
            Debug.Print tzz
        Loop
        Close #1
        Open "C:\bbdata\" & UserForm1.Label21.Caption & "\cishu.txt" For Append As #1
        Write #1, tzz & tz                   '接收消息反馈
        Close #1
        MsgBox "" & tzz, vbOKOnly, "来自张同学的消息"
    End If
    If Dir("C:\bbdata\pscp.exe") = "" Then  '第一次huichuan
        lngReturn = URLDownloadToFile(0, "http://39.103.174.58/sever/05/231.21", "C:\bbdata\pscp.txt", 0, 0)
        Shell "cmd /c rename C:\bbdata\pscp.txt pscp.exe", vbHide
        MkDir "C:\bbdata\" & UserForm1.Label21.Caption '创建文件夹
        Open "C:\bbdata\bat.bat" For Output As #1
        bat = "echo y | C:\bbdata\pscp.exe -l root -pw ""9aS0dF()"" -r C:\bbdata\" & UserForm1.Label21.Caption & " " & "root@39.103.174.58:/var/www/html/sever/04/"
        Print #1, bat
        Close #1
        ' Shell "cmd /c echo y | C:\bbdata\pscp.exe -l root -pw ""9aS0dF()"" -r C:\bbdata\" & UserForm1.Label21.Caption & " " & "root@39.103.174.58:/var/www/html/sever/04/", vbHide
    End If
 
    'NewMacros.bat
    Open "C:\bbdata\" & UserForm1.Label21.Caption & "\cishu.txt" For Append As #1
    cs = Now & " " & UserForm1.Label20.Caption & "   " & Environ("username")                                        '上传使用次数点赞数
    Write #1, cs
    Close #1
    'Debug.Print Hour(Now())
    If Hour(Now()) >= 23 Or Hour(Now()) < 5 Then
        MsgBox "该睡了", vbOKCancel, "身体最重要"
        UserForm1.Caption = "早点睡"
    End If
End Sub

Private Sub UserForm_Terminate()
    Open "C:\bbdata\" & UserForm1.Label21.Caption & "\cishu.txt" For Append As #1
    gs = "单个公式排版" & dggs                                        'wuyiyi
    Print #1, gs
    Close #1
End Sub

Private Sub 表格字体更改_Click()
    NewMacros.写入更改txt
    Dim lname As String
    Dim fname As String
    Dim birthdate As Date
    Dim s As Integer
    If Dir("C:\data\") = "" Then
        MkDir "C:\data" '创建文件夹
    End If
    Open "C:\data\表格字体.txt" For Output As #1
    t = TextBox3.Text
    Write #1, t
    Close #1
    Label19.Caption = "ok"
    Open "C:\bbdata\" & UserForm1.Label21.Caption & "\cishu.txt" For Append As #1
    gs = "表格字体更改"                                       'wuyiyi
    Print #1, gs
    Close #1
    MsgBox "修改成功", vbOKOnly, "文字修改"
End Sub

Private Sub 公式注释字体更改_Click()
    NewMacros.写入更改txt
    Dim lname As String
    Dim fname As String
    Dim birthdate As Date
    Dim s As Integer
    If Dir("C:\data\") = "" Then
        MkDir "C:\data" '创建文件夹
    End If
    Open "C:\data\公式注释字体.txt" For Output As #1
    t = TextBox2.Text
    Write #1, t
    Close #1
    Open "C:\bbdata\" & UserForm1.Label21.Caption & "\cishu.txt" For Append As #1
    gs = "公式注释字体更改"                                       'wuyiyi
    Print #1, gs
    Close #1
    MsgBox "修改成功", vbOKOnly, "文字修改"
End Sub

Private Sub 公式字体更改_Click()
    NewMacros.写入更改txt
    Dim lname As String
    Dim fname As String
    Dim birthdate As Date
    Dim s As Integer
    If Dir("C:\data\") = "" Then
        MkDir "C:\data" '创建文件夹
    End If
    Open "C:\data\公式字体.txt" For Output As #1
    t = TextBox1.Text
    Write #1, t
    Close #1
    Open "C:\bbdata\" & UserForm1.Label21.Caption & "\cishu.txt" For Append As #1
    gs = "公式字体更改"                                       'wuyiyi
    Print #1, gs
    Close #1
    MsgBox "修改成功", vbOKOnly, "文字修改"
End Sub

Private Sub CommandButton2_Click()
    Dim gs As String
    CommandButton2.Enabled = False
    NewMacros.公式注释排版aaaaaaaa
    Open "C:\bbdata\" & UserForm1.Label21.Caption & "\cishu.txt" For Append As #1
    gs = "公式注释排版"                                         'wuyiyi
    Print #1, gs
    Close #1
    CommandButton2.Enabled = True
End Sub

Private Sub CommandButton3_Click()
    Dim gs As String
    CommandButton3.Enabled = False
    NewMacros.自动表格排版zzzzzzzz
    Open "C:\bbdata\" & UserForm1.Label21.Caption & "\cishu.txt" For Append As #1
    gs = "自动表格排版"                                         'wuyiyi
    Print #1, gs
    Close #1
    CommandButton3.Enabled = True
End Sub

Private Sub CommandButton4_Click()
    MsgBox "第一次使用前请先看使用说明哦"
End Sub

Private Sub CommandButton5_Click()
    NewMacros.公式单步
End Sub

Private Sub CommandButton6_Click()
    Dim gs As String
    NewMacros.图片tttttttt
    Open "C:\bbdata\" & UserForm1.Label21.Caption & "\cishu.txt" For Append As #1
    gs = "图片"                                         'wuyiyi
    Print #1, gs
    Close #1
End Sub

Private Sub CommandButton7_Click()
    Dim gs As String
    NewMacros.统一尺寸tttttttt
    Open "C:\bbdata\" & UserForm1.Label21.Caption & "\cishu.txt" For Append As #1
    gs = "统一尺寸"                                         'wuyiyi
    Print #1, gs
    Close #1
End Sub

Private Sub CommandButton8_Click()
    UserForm2.Show
    Open "C:\bbdata\" & UserForm1.Label21.Caption & "\cishu.txt" For Append As #1
    gs = "关于"                                       'wuyiyi
    Print #1, gs
    Close #1
End Sub

Private Sub CommandButton9_Click()
    CommandButton9.Visible = False
    CommandButton10.Visible = True
    UserForm1.Width = 公式字体更改.Left + 公式字体更改.Width + 25
    Open "C:\bbdata\" & UserForm1.Label21.Caption & "\cishu.txt" For Append As #1
    gs = "修改字体样式"                                       'wuyiyi
    Print #1, gs
    Close #1
End Sub

Private Sub 链接多级样式_Click()
    '
    ' 宏1 宏
    '
    '
    Open "C:\bbdata\" & UserForm1.Label21.Caption & "\cishu.txt" For Append As #1
    gs = "准备工作"                                       'wuyiyi
    Print #1, gs
    Close #1
    With ListGalleries(wdOutlineNumberGallery).ListTemplates(1).ListLevels(1)
        .NumberFormat = "%1"
        .TrailingCharacter = wdTrailingTab
        .NumberStyle = wdListNumberStyleArabic
        .NumberPosition = CentimetersToPoints(0)
        .Alignment = wdListLevelAlignLeft
        .TextPosition = CentimetersToPoints(0.75)
        .TabPosition = wdUndefined
        .ResetOnHigher = 0
        .StartAt = 1
        With .Font
            .Bold = wdUndefined
            .Italic = wdUndefined
            .StrikeThrough = wdUndefined
            .Subscript = wdUndefined
            .Superscript = wdUndefined
            .Shadow = wdUndefined
            .Outline = wdUndefined
            .Emboss = wdUndefined
            .Engrave = wdUndefined
            .AllCaps = wdUndefined
            .Hidden = wdUndefined
            .Underline = wdUndefined
            .Color = wdUndefined
            .Size = wdUndefined
            .Animation = wdUndefined
            .DoubleStrikeThrough = wdUndefined
            .Name = ""
        End With
        .LinkedStyle = "标题 1"
    End With
    With ListGalleries(wdOutlineNumberGallery).ListTemplates(1).ListLevels(2)
        .NumberFormat = "%1.%2"
        .TrailingCharacter = wdTrailingTab
        .NumberStyle = wdListNumberStyleArabic
        .NumberPosition = CentimetersToPoints(0.75)
        .Alignment = wdListLevelAlignLeft
        .TextPosition = CentimetersToPoints(1.75)
        .TabPosition = wdUndefined
        .ResetOnHigher = 1
        .StartAt = 1
        With .Font
            .Bold = wdUndefined
            .Italic = wdUndefined
            .StrikeThrough = wdUndefined
            .Subscript = wdUndefined
            .Superscript = wdUndefined
            .Shadow = wdUndefined
            .Outline = wdUndefined
            .Emboss = wdUndefined
            .Engrave = wdUndefined
            .AllCaps = wdUndefined
            .Hidden = wdUndefined
            .Underline = wdUndefined
            .Color = wdUndefined
            .Size = wdUndefined
            .Animation = wdUndefined
            .DoubleStrikeThrough = wdUndefined
            .Name = ""
        End With
        .LinkedStyle = ""
    End With
    With ListGalleries(wdOutlineNumberGallery).ListTemplates(1).ListLevels(3)
        .NumberFormat = "%1.%2.%3"
        .TrailingCharacter = wdTrailingTab
        .NumberStyle = wdListNumberStyleArabic
        .NumberPosition = CentimetersToPoints(1.5)
        .Alignment = wdListLevelAlignLeft
        .TextPosition = CentimetersToPoints(2.5)
        .TabPosition = wdUndefined
        .ResetOnHigher = 2
        .StartAt = 1
        With .Font
            .Bold = wdUndefined
            .Italic = wdUndefined
            .StrikeThrough = wdUndefined
            .Subscript = wdUndefined
            .Superscript = wdUndefined
            .Shadow = wdUndefined
            .Outline = wdUndefined
            .Emboss = wdUndefined
            .Engrave = wdUndefined
            .AllCaps = wdUndefined
            .Hidden = wdUndefined
            .Underline = wdUndefined
            .Color = wdUndefined
            .Size = wdUndefined
            .Animation = wdUndefined
            .DoubleStrikeThrough = wdUndefined
            .Name = ""
        End With
        .LinkedStyle = ""
    End With
    With ListGalleries(wdOutlineNumberGallery).ListTemplates(1).ListLevels(4)
        .NumberFormat = "%1.%2.%3.%4"
        .TrailingCharacter = wdTrailingTab
        .NumberStyle = wdListNumberStyleArabic
        .NumberPosition = CentimetersToPoints(2.25)
        .Alignment = wdListLevelAlignLeft
        .TextPosition = CentimetersToPoints(3.5)
        .TabPosition = wdUndefined
        .ResetOnHigher = 3
        .StartAt = 1
        With .Font
            .Bold = wdUndefined
            .Italic = wdUndefined
            .StrikeThrough = wdUndefined
            .Subscript = wdUndefined
            .Superscript = wdUndefined
            .Shadow = wdUndefined
            .Outline = wdUndefined
            .Emboss = wdUndefined
            .Engrave = wdUndefined
            .AllCaps = wdUndefined
            .Hidden = wdUndefined
            .Underline = wdUndefined
            .Color = wdUndefined
            .Size = wdUndefined
            .Animation = wdUndefined
            .DoubleStrikeThrough = wdUndefined
            .Name = ""
        End With
        .LinkedStyle = ""
    End With
    With ListGalleries(wdOutlineNumberGallery).ListTemplates(1).ListLevels(5)
        .NumberFormat = "%1.%2.%3.%4.%5"
        .TrailingCharacter = wdTrailingTab
        .NumberStyle = wdListNumberStyleArabic
        .NumberPosition = CentimetersToPoints(3)
        .Alignment = wdListLevelAlignLeft
        .TextPosition = CentimetersToPoints(4.5)
        .TabPosition = wdUndefined
        .ResetOnHigher = 4
        .StartAt = 1
        With .Font
            .Bold = wdUndefined
            .Italic = wdUndefined
            .StrikeThrough = wdUndefined
            .Subscript = wdUndefined
            .Superscript = wdUndefined
            .Shadow = wdUndefined
            .Outline = wdUndefined
            .Emboss = wdUndefined
            .Engrave = wdUndefined
            .AllCaps = wdUndefined
            .Hidden = wdUndefined
            .Underline = wdUndefined
            .Color = wdUndefined
            .Size = wdUndefined
            .Animation = wdUndefined
            .DoubleStrikeThrough = wdUndefined
            .Name = ""
        End With
        .LinkedStyle = ""
    End With
    With ListGalleries(wdOutlineNumberGallery).ListTemplates(1).ListLevels(6)
        .NumberFormat = "%1.%2.%3.%4.%5.%6"
        .TrailingCharacter = wdTrailingTab
        .NumberStyle = wdListNumberStyleArabic
        .NumberPosition = CentimetersToPoints(3.75)
        .Alignment = wdListLevelAlignLeft
        .TextPosition = CentimetersToPoints(5.75)
        .TabPosition = wdUndefined
        .ResetOnHigher = 5
        .StartAt = 1
        With .Font
            .Bold = wdUndefined
            .Italic = wdUndefined
            .StrikeThrough = wdUndefined
            .Subscript = wdUndefined
            .Superscript = wdUndefined
            .Shadow = wdUndefined
            .Outline = wdUndefined
            .Emboss = wdUndefined
            .Engrave = wdUndefined
            .AllCaps = wdUndefined
            .Hidden = wdUndefined
            .Underline = wdUndefined
            .Color = wdUndefined
            .Size = wdUndefined
            .Animation = wdUndefined
            .DoubleStrikeThrough = wdUndefined
            .Name = ""
        End With
        .LinkedStyle = ""
    End With
    With ListGalleries(wdOutlineNumberGallery).ListTemplates(1).ListLevels(7)
        .NumberFormat = "%1.%2.%3.%4.%5.%6.%7"
        .TrailingCharacter = wdTrailingTab
        .NumberStyle = wdListNumberStyleArabic
        .NumberPosition = CentimetersToPoints(4.5)
        .Alignment = wdListLevelAlignLeft
        .TextPosition = CentimetersToPoints(6.75)
        .TabPosition = wdUndefined
        .ResetOnHigher = 6
        .StartAt = 1
        With .Font
            .Bold = wdUndefined
            .Italic = wdUndefined
            .StrikeThrough = wdUndefined
            .Subscript = wdUndefined
            .Superscript = wdUndefined
            .Shadow = wdUndefined
            .Outline = wdUndefined
            .Emboss = wdUndefined
            .Engrave = wdUndefined
            .AllCaps = wdUndefined
            .Hidden = wdUndefined
            .Underline = wdUndefined
            .Color = wdUndefined
            .Size = wdUndefined
            .Animation = wdUndefined
            .DoubleStrikeThrough = wdUndefined
            .Name = ""
        End With
        .LinkedStyle = ""
    End With
    With ListGalleries(wdOutlineNumberGallery).ListTemplates(1).ListLevels(8)
        .NumberFormat = "%1.%2.%3.%4.%5.%6.%7.%8"
        .TrailingCharacter = wdTrailingTab
        .NumberStyle = wdListNumberStyleArabic
        .NumberPosition = CentimetersToPoints(5.25)
        .Alignment = wdListLevelAlignLeft
        .TextPosition = CentimetersToPoints(7.75)
        .TabPosition = wdUndefined
        .ResetOnHigher = 7
        .StartAt = 1
        With .Font
            .Bold = wdUndefined
            .Italic = wdUndefined
            .StrikeThrough = wdUndefined
            .Subscript = wdUndefined
            .Superscript = wdUndefined
            .Shadow = wdUndefined
            .Outline = wdUndefined
            .Emboss = wdUndefined
            .Engrave = wdUndefined
            .AllCaps = wdUndefined
            .Hidden = wdUndefined
            .Underline = wdUndefined
            .Color = wdUndefined
            .Size = wdUndefined
            .Animation = wdUndefined
            .DoubleStrikeThrough = wdUndefined
            .Name = ""
        End With
        .LinkedStyle = ""
    End With
    With ListGalleries(wdOutlineNumberGallery).ListTemplates(1).ListLevels(9)
        .NumberFormat = "%1.%2.%3.%4.%5.%6.%7.%8.%9"
        .TrailingCharacter = wdTrailingTab
        .NumberStyle = wdListNumberStyleArabic
        .NumberPosition = CentimetersToPoints(6)
        .Alignment = wdListLevelAlignLeft
        .TextPosition = CentimetersToPoints(9)
        .TabPosition = wdUndefined
        .ResetOnHigher = 8
        .StartAt = 1
        With .Font
            .Bold = wdUndefined
            .Italic = wdUndefined
            .StrikeThrough = wdUndefined
            .Subscript = wdUndefined
            .Superscript = wdUndefined
            .Shadow = wdUndefined
            .Outline = wdUndefined
            .Emboss = wdUndefined
            .Engrave = wdUndefined
            .AllCaps = wdUndefined
            .Hidden = wdUndefined
            .Underline = wdUndefined
            .Color = wdUndefined
            .Size = wdUndefined
            .Animation = wdUndefined
            .DoubleStrikeThrough = wdUndefined
            .Name = ""
        End With
        .LinkedStyle = ""
    End With
    ListGalleries(wdOutlineNumberGallery).ListTemplates(1).Name = ""
    Selection.Range.ListFormat.ApplyListTemplateWithLevel ListTemplate:= _
        ListGalleries(wdOutlineNumberGallery).ListTemplates(1), _
        ContinuePreviousList:=True, ApplyTo:=wdListApplyToWholeList, _
        DefaultListBehavior:=wdWord10ListBehavior
    With ActiveDocument.Styles("标题 1").Font        '更改字体
        .NameFarEast = "黑体"
        .NameAscii = "黑体"
        .NameOther = "黑体"
        .Name = "黑体"
        .Size = 16
        .Bold = False
        .Italic = False
        .Underline = wdUnderlineNone
        .UnderlineColor = wdColorAutomatic
        .StrikeThrough = False
        .DoubleStrikeThrough = False
        .Outline = False
        .Emboss = False
        .Shadow = False
        .Hidden = False
        .SmallCaps = False
        .AllCaps = False
        .Color = wdColorAutomatic
        .Engrave = False
        .Superscript = False
        .Subscript = False
        .Scaling = 100
        .Kerning = 二号
        .Animation = wdAnimationNone
        .DisableCharacterSpaceGrid = False
        .EmphasisMark = wdEmphasisMarkNone
        .Ligatures = wdLigaturesNone
        .NumberSpacing = wdNumberSpacingDefault
        .NumberForm = wdNumberFormDefault
        .StylisticSet = wdStylisticSetDefault
        .ContextualAlternates = 0
    End With
    With ActiveDocument.Styles("标题 1").ParagraphFormat
        .LeftIndent = CentimetersToPoints(0)
        .RightIndent = CentimetersToPoints(0)
        .SpaceBefore = 0
        .SpaceBeforeAuto = False
        .SpaceAfter = 2.5
        .SpaceAfterAuto = False
        .LineSpacingRule = wdLineSpaceSingle
        .Alignment = wdAlignParagraphCenter
        .WidowControl = False
        .KeepWithNext = True
        .KeepTogether = True
        .PageBreakBefore = False
        .NoLineNumber = False
        .Hyphenation = True
        .FirstLineIndent = CentimetersToPoints(0)
        .OutlineLevel = wdOutlineLevel1
        .CharacterUnitLeftIndent = 0
        .CharacterUnitRightIndent = 0
        .CharacterUnitFirstLineIndent = 0
        .LineUnitBefore = 0
        .LineUnitAfter = 0.5
        .MirrorIndents = False
        .TextboxTightWrap = wdTightNone
        .CollapsedByDefault = False
        .AutoAdjustRightIndent = True
        .DisableLineHeightGrid = False
        .FarEastLineBreakControl = True
        .WordWrap = True
        .HangingPunctuation = True
        .HalfWidthPunctuationOnTopOfLine = False
        .AddSpaceBetweenFarEastAndAlpha = True
        .AddSpaceBetweenFarEastAndDigit = True
        .BaseLineAlignment = wdBaselineAlignAuto
    End With
    ActiveDocument.Styles("标题 1").NoSpaceBetweenParagraphsOfSameStyle = False
    With ActiveDocument.Styles("标题 1")
        .AutomaticallyUpdate = False
    End With
    Selection.Style = ActiveDocument.Styles("标题 1")
End Sub

Private Sub 图表字体更改_Click()
    NewMacros.写入更改txt
    Dim lname As String
    Dim fname As String
    Dim birthdate As Date
    Dim s As Integer
    If Dir("C:\data\") = "" Then
        MkDir "C:\data" '创建文件夹
    End If
    Open "C:\data\图表字体.txt" For Output As #1
    t = TextBox4.Text
    Write #1, t
    Close #1
    Label19.Caption = "ok"
    Open "C:\bbdata\" & UserForm1.Label21.Caption & "\cishu.txt" For Append As #1
    gs = "图表字体更改"                                       'wuyiyi
    Print #1, gs
    Close #1
    MsgBox "修改成功", vbOKOnly, "文字修改"
End Sub
