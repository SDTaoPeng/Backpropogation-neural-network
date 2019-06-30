Attribute VB_Name = "FrmBP2C"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit                       
Dim xb As Double, Yb As Double        
Dim Px As Integer, Py As Integer     
Dim Px0 As Integer, Py0 As Integer, Px1 As Integer, Py1 As Integer
Dim xx As Double, yy As Double
Dim Kx As Double, Ky As Double, KC As Double        
Const xmax = 0.99: Const xb1 = 0.05
Const ymax = 0.99: Const yb1 = 0.05                  
Dim Qc As Double
Dim t1 As Double
Dim StudyNumber As Long


'================
Dim Xishumax As Integer, Xishu(0 To 10) As Double
Dim X2() As Double, Y2() As Double

Dim RX() As Double, RY() As Double
'================

'=======================================================
Dim wc1() As Double  '测量点
Dim wc2() As Double  '测量值
Dim wcp As Double    '测量值平均
Dim R2 As Double     '相关系数的平方
'--------------------------------------------------------
Dim I1 As Long, I2 As Long, i3 As Long  '临时使用
Dim d1 As Double
Dim s1 As String
'---------------------------------------------------------
Const pi = 3.141592

Dim N As Long                  '输入层单元个数:n
Dim i As Long                  '输入层变量:i=1 to n

Dim p As Long                  '隐含层(中间层)单元个数:p
Dim j As Long                  '隐含层(中间层）变量:j=1 to p

Dim q As Long                  '输出层单元个数:q
Dim t As Long                  '输出层变量:t=1 to q

Dim m As Long                  '学习模式对数:m
Dim k As Long                  '学习模式变量:k=1 to m

Dim A() As Double              
Dim y() As Double              

Dim w() As Double              '输入层至中间层连接权数组:w(i,j)
Dim v() As Double              '中间层至输出层连接权数组:v(j,t)
Dim O() As Double              '中间层各单元输出阈值:O(j)
Dim r() As Double              '输出层各单元输出阈值:r(t)

Dim s() As Double    '中间层各单元的输入s(k,j)
Dim B() As Double    '中间层各单元的输出b(k,j)
Dim l() As Double    '输出层各单元的输入l(k,t)
Dim c() As Double    '输出层各单元的输出c(k,t)
Dim d() As Double    '输出层各单元的一般化误差d(k,t)
Dim e() As Double    '中间层各单元的一般化误差e(k,j)

Dim LN As Long       
Dim RF As Double     
Dim BT As Double      
Dim EE As Double      

Dim ss() As Double    
Dim aa() As Double    
Dim bb() As Double    
Dim ll() As Double    
Dim cc() As Double    


Private Sub Form_Load()

   File1.Path = "..\Tdata"       '数据文件路径  
   Learn_Pause = "ON"
   Pause.Caption = "暂停学习"

End Sub


Private Sub File1_Click()        '选择数据文件
    TDataFileName = File1.Path & "\" & File1.FileName
    HSDataFileName = "..\HS\" & File1.FileName
    GenDataFileName = "..\gencurve\" & File1.FileName

    Text1.Text = "已经选择的数据文件为:" & "   ..\Tdata\" & File1.FileName
End Sub


Private Sub CmdPara_Click()      '“设定确认”按钮
   p = Val(Text6.Text)           '可调的隐含层(中间层)单元个数:p
   StudyNumber = Val(Text5.Text) '学习次数（可获取不同学习次数的最终结果值）
   Text5.BackColor = &HFFC0FF
   Text6.BackColor = &HFFC0FF
End Sub
   
Private Sub CmdStudySatrt_Click()    '单击"开始学习"按钮
    Call PraSub
    '(1)给各连接权w(i,j)、v(j,t)、及阈值O(j)、r(t)赋予(-0.1,+0.1)间的随机值
    
    For i = 1 To N
       For j = 1 To p: w(i, j) = SgnRnd(Rnd()) * 0.1 * Rnd(): Next j
    Next i
    
    For j = 1 To p
       For t = 1 To q: v(j, t) = SgnRnd(Rnd()) * 0.1 * Rnd(): Next t
    Next j
    
    For j = 1 To p: O(j) = SgnRnd(Rnd()) * 0.1 * Rnd(): Next j
    
    For t = 1 To q: r(t) = SgnRnd(Rnd()) * 0.1 * Rnd(): Next t
    
    '(2)随机选取一个学习模式对提供给网络,即k选取1到m之间的任意值(整数)
    'k = Int(m * (Rnd())) + 1
   t1 = Timer()
   For LN = 1 To StudyNumber     '学习次数循环
        DoEvents
       
        If Learn_Pause = "OFF" Then
           Do
             DoEvents
           Loop Until Learn_Pause <> "OFF"
        End If
       For k = 1 To m         '学习模式对循环
            '(3)用a(k,i)(i=1 to n),连接权w(i,j),阈值O(j)计算中间各单元的输入s(k,j),通过
            For j = 1 To p
                s(k, j) = 0
                For i = 1 To N: s(k, j) = s(k, j) + w(i, j) * A(k, i): Next i
                s(k, j) = s(k, j) - O(j)                                      '(4.32)
                B(k, j) = fs(s(k, j))                                         '(4.33)
            Next j
            '(4)用中间层各单元的输出b(k,j)(j=1 to p),连接权v(j,t),阈值r(j)计算输出层各单元的输入l(k,t),通过
            'S函数计算输出层各单元的输出响应c(k,t)
            For t = 1 To q
                l(k, t) = 0
                For j = 1 To p: l(k, t) = l(k, t) + v(j, t) * B(k, j): Next j
                l(k, t) = l(k, t) - r(t)
                c(k, t) = fs(l(k, t))
            Next t
            '(5)用希望输出模式y(k,t) (t=1 to q),网络实际输出c(k,t),计算输出层的各单元的一般化误差d(k,t)
            For t = 1 To q
                d(k, t) = (y(k, t) - c(k, t)) * c(k, t) * (1 - c(k, t))     '(4.36)
            Next t
            '(6)用连接权v(j,t),输出层一般化误差d(k,t),中间层的输出b(k,j),计算中间层各单元的一般化误差e(k,j)
            For j = 1 To p
                d1 = 0: For t = 1 To q: d1 = d1 + d(k, t) * v(j, t): Next t
                e(k, j) = d1 * B(k, j) * (1 - B(k, j))                      '(4.37)
            Next j
            '(7)用输出层一般化误差d(k,t),中间层的输出b(k,j)修正连接权v(j,t)和阈值r(t)
            For j = 1 To p
                For t = 1 To q
                    v(j, t) = v(j, t) + RF * d(k, t) * B(k, j)  '(4.38)
                Next t
            Next j
            For t = 1 To q
                r(t) = r(t) - RF * d(k, t)                      '(4.39)
            Next t
            '(8)用中间层一般化误差e(k,j),输入层的a(k,i)修正连接权w(i,j)和阈值O(j)
            For i = 1 To N
                For j = 1 To p
                    w(i, j) = w(i, j) + BT * e(k, j) * A(k, i)  '(4.40)
                Next j
            Next i
            For j = 1 To p
                O(j) = O(j) - BT * e(k, j)                      '(4.41)
            Next j
       Next k
       If (Int(LN / 1000) * 1000 = LN) Or (LN = StudyNumber) Then
          '均方差EE
          EE = 0#
           For k = 1 To m
             For t = 1 To q
			   EE = (y(k, t) - c(k, t)) * (y(k, t) - c(k, t)) / 2         '求均方差
			   'EE =  EE + (y(k, t) - c(k, t)) * (y(k, t) - c(k, t)) / 2  '求全局误差
            Next t
          Next k
          TxtEE.Text = Format$(EE, "0.00000000000000"): TxtEE.Refresh    '显示
          TxtLearnNumber.Text = Format$(LN, "############0"): TxtLearnNumber.Refresh
          TxtTimer.Text = Format$(Timer() - t1, "#######0") & "秒": TxtTimer.Refresh
          '显示连接权w(i,j)
          s1 = ""
          For i = 1 To N
              For j = 1 To p
              s1 = s1 & "ω(" & Format$(i, "0") & "," & Format$(j, "0") & ")=" & Format$(w(i, j), "###0.00000000") & Chr(&HD) & Chr(&HA)
              Next j
          Next i
          TxtW.Text = s1
           '显示阈值O(j)
          s1 = ""
          For j = 1 To p
              s1 = s1 & "θ(" & Format$(j, "0") & ")=" & Format$(O(j), "###0.00000000") & Chr(&HD) & Chr(&HA)
          Next j
          TxtO.Text = s1
          '显示连接权v(j,t)
          s1 = ""
          For j = 1 To p
              For t = 1 To q
              s1 = s1 & "υ(" & Format$(j, "0") & "," & Format$(t, "0") & ")=" & Format$(v(j, t), "###0.00000000") & Chr(&HD) & Chr(&HA)
              Next t
          Next j
          Txtv.Text = s1
          '显示阈值γ
          s1 = ""
          For t = 1 To q
              s1 = s1 & "γ(" & Format$(t, "0") & ")=" & Format$(r(t), "###0.00000000") & Chr(&HD) & Chr(&HA)
          Next t
          Txtr.Text = s1
      End If
   Next LN     '对学习次数循环
End Sub

Private Sub PraSub()             '初始化
   N = 1                      
   q = 2                        
   p = Val(Text6.Text)          
   '
   Call ReadFile(TDataFileName) 
   '
   m = UBound(Tfx)               
   StudyNumber = Val(Text5.Text) 
   '-----------------------------------------------------------
   ReDim w(1 To N, 1 To p) As Double    
   ReDim v(1 To p, 1 To q) As Double    
   ReDim O(1 To p) As Double           
   ReDim r(1 To p) As Double            
   ReDim A(1 To m, 1 To N) As Double    
   ReDim y(1 To m, 1 To q) As Double    
   ReDim s(1 To m, 1 To p) As Double    
   ReDim B(1 To m, 1 To p) As Double    
   ReDim l(1 To m, 1 To q) As Double    
   ReDim c(1 To m, 1 To q) As Double    
   ReDim d(1 To m, 1 To q) As Double    
   ReDim e(1 To m, 1 To p) As Double    
   ReDim ss(1 To p) As Double   
   ReDim bb(1 To p) As Double  
   ReDim aa(1 To N) As Double 
   ReDim ll(1 To q) As Double   
   ReDim cc(1 To q) As Double  
   '----------回归问题学习模式对-----------------------------
   
'   For i = LBound(Tfx) To UBound(Tfx)
'       a(i, 1) = Tfx(i)
'       y(i, 1) = 0.5 + 0.4 * Sin(a(i, 1) * 2 * pi) + SgnRnd(Rnd()) * 0.15 * Rnd()  '
'       y(i, 2) = 0.5 + 0.4 * Cos(a(i, 1) * 2 * pi) + SgnRnd(Rnd()) * 0.15 * Rnd() '
'   Next i
   
    For i = LBound(Tfx) To UBound(Tfx)
       A(i, 1) = Tfx(i)
       y(i, 1) = Fxy(i, 1)
       y(i, 2) = Fxy(i, 2)
   Next i   
   
   '---------------------------------------------------------
   ReDim wc1(1 To m) As Double
   ReDim wc2(1 To m, 1 To q) As Double
   
   wcp = 0
   For i = 1 To m
      wc1(i) = A(i, 1): wc2(i, 1) = y(i, 1): wc2(i, 2) = y(i, 2):
      wcp = wcp + wc2(i, 1) + wc2(i, 2)
   Next i
   wcp = wcp / m
   '------------------------
    RF = 0.5: BT = 0.5     '学习率1,2
End Sub


Private Function SgnRnd(x As Double) As Double  '定义随机符号函数
   If x < 0.5 Then
      SgnRnd = -1
   Else
      SgnRnd = 1
   End If
End Function


Private Function fs(x As Double) As Double     '定义S函数
   fs = 1 / (1 + Exp(-x))
End Function


Private Sub CmdDataPoint_Click()      '画数据点
   Call DrawZoBiao    '画坐标
      '画数据点
   For I1 = LBound(Tfx) To UBound(Tfx) - 2 Step 1
   'For i1 = 0 To 0.9999
       'Call DrawLargePoint(y(I1, 1), y(I1, 2), vbBlack)
       Call DrawLargePoint(y(I1, 1), y(I1, 2), &H808080)
       
   Next I1
    'Call ReadFile("PCA_Ti_n.TXT")
End Sub



Private Sub Cndstart_Click()     '"回想分析"按钮---统一的轮廓数学表达式为此处
   Dim I1 As Integer
   Dim d1 As Double
   Dim N2 As Integer, M2 As Integer
       Dim Ir As Integer, Jr As Integer, Nr As Integer
    Dim Sar() As String
    
    '----------------------------------------------------
    Open TDataFileName For Input As #1          '有正确的文件名,打开文件
    Ir = 0                                 '文件总行数初值=0
    Do Until EOF(1)
       Ir = Ir + 1
       ReDim Preserve Sar(1 To Ir) As String '重新定义字符串数组的最大下标
       Line Input #1, Sar(Ir)      '读一行—>最新行
    Loop
    Close #1                               '关闭文件
    '
    '重新定义x()、y()数组,给它们赋值
    ReDim X2(1 To Ir - 1) As Double                 '重新定义最大下标
    ReDim Y2(1 To Ir - 1) As Double
    For Jr = 1 To Ir - 1
        Sar(Jr) = LTrim$(RTrim$(Sar(Jr)))
        X2(Jr) = Mid$(Sar(Jr), 14, 13)  'x (第一分量)
        Y2(Jr) = Mid$(Sar(Jr), 28, 13)  'y (第二分量)
    Next Jr
    
    '----------------------------------------------------
    '
    M2 = 20
    N2 = UBound(Y2) / M2 - 1
       ReDim RX(0 To N2)
    ReDim RY(0 To N2)

    Call DrawZoBiao    '画坐标
    
    For I1 = 0 To N2
        d1 = I1 * M2 + 1
        RX(I1) = (X2(d1) + X2(d1 + 1) + X2(d1 + 2) + X2(d1 + 3)) / 4
        RY(I1) = (Y2(d1) + Y2(d1 + 1) + Y2(d1 + 2) + Y2(d1 + 3)) / 4
        'Call DrawZPoint(RX(I1), RY(I1), vbRed)
    Next I1
    
    For I1 = 1 To UBound(X2)   '对数据点循环
       Call DrawLargePoint(X2(I1), Y2(I1), &H808080)
    Next I1
   
   Call DrawZoBiao    '画坐标
  
   '画回想曲线
   For d1 = Tfx(1) To Tfx(m) Step 0.001
   'For d1 = 0.001 To 0.999 Step 0.002
       Call hx(d1)
       Call DrawXYPiont2(cc(1), cc(2), vbBlue)
   Next d1
   
            For I1 = 0 To N2
             Call DrawZPoint(RX(I1), RY(I1), vbRed)
            'Call DrawZPoint(RX(I1), RY(I1), vbRed)
         Next I1
   '========神经网络回归=====计算R
End Sub


Private Sub hx(xb)     '"回想-曲线回画"
    aa(1) = xb
    For j = 1 To p
         ss(j) = 0
         ss(j) = ss(j) + w(1, j) * aa(1)
         ss(j) = ss(j) - O(j)
         bb(j) = fs(ss(j))
    Next j
     For t = 1 To q
         ll(t) = 0: For j = 1 To p: ll(t) = ll(t) + v(j, t) * bb(j): Next j
         ll(t) = ll(t) - r(t)
         cc(t) = fs(ll(t))
     Next t
End Sub



Private Sub DrawXYPiont(xx, yy, Color)
    Px = Kx * xx + Px0: Py = Ky * yy + Py0
    PicC_Qc.Line (Px - 50, Py - 50)-(Px + 50, Py + 50), Color
    PicC_Qc.Line (Px - 50, Py + 50)-(Px + 50, Py - 50), Color
End Sub


Private Sub DrawXYPiont1(xx, yy, Color)
    Px = Kx * xx + Px0 - 1: Py = Ky * yy + Py0 - 1: Px1 = Px + 2: Py1 = Py + 2
    PicC_Qc.Line (Px, Py)-(Px1, Py1), Color, B
End Sub

Private Sub DrawXYPiont2(xx, yy, Color)
    Px = Kx * xx + Px0: Py = Ky * yy + Py0
    PicC_Qc.PSet (Px, Py), Color
    PicC_Qc.DrawWidth = 7
End Sub


'-----------------------------------------------------------------------------------------------
Private Sub ReadFile(FileName As String)
    Dim Ir As Integer, Jr As Integer, Nr As Integer
    Dim Sar() As String
    '
    If Len(Trim(LTrim(Text1.Text))) <= 10 Then
       MsgBox ("未选择数据文件")
       Exit Sub
    End If
    Open FileName For Input As #1          '有正确的文件名,打开文件
    Ir = 0                                 '文件总行数初值=0
    Do Until EOF(1)
       Ir = Ir + 1
       ReDim Preserve Sar(1 To Ir) As String 
       Line Input #1, Sar(Ir)      '读一行—>最新行
    Loop
    Close #1                               '关闭文件
    '
    '重新定义x()、y()数组,给它们赋值
    ReDim Tfx(1 To Ir - 1) As Double                 
    ReDim Fxy(1 To Ir - 1, 1 To 2) As Double         
    For Jr = 1 To Ir - 1
        Sar(Jr) = LTrim$(RTrim$(Sar(Jr)))
        Tfx(Jr) = Mid$(Sar(Jr), 1, 12)   
        Fxy(Jr, 1) = Mid$(Sar(Jr), 14, 13)  
        Fxy(Jr, 2) = Mid$(Sar(Jr), 28, 13)  
    Next Jr
    '---------------------------------------------------
    xymax1 = Mid$(LTrim$(RTrim$(Sar(Ir))), 1, 12)  '
    Sumx1 = Mid$(Sar(Ir), 14, 13)                  '
    Sumy1 = Mid$(Sar(Ir), 28, 13)                  '
End Sub



Private Sub CmdDistance_Click()   '求总距离 deta f
    Dim i As Integer, I1 As Integer, I2 As Integer
    Dim j As Integer, J1 As Integer, j2 As Integer
    Dim t1 As Double
    Dim d1 As Double, d2 As Double, Dz As Double
    Dim t10 As Double, t11 As Double
    '求PCA-BP曲线点->CurcvsPoint
   For J1 = 0 To 9999
       t1 = J1 / 10000
       Call hx(t1)
       CurcvsPoint(J1).x = cc(1): CurcvsPoint(J1).y = cc(2)
   Next J1
   '求Dz
   Dz = 0#
   For I1 = 1 To m    '对数据点循环
       d1 = (Fxy(I1, 1) - CurcvsPoint(0).x) ^ 2 + (Fxy(I1, 2) - CurcvsPoint(0).y) ^ 2
       For j2 = 1 To 9999
          d2 = (Fxy(I1, 1) - CurcvsPoint(j2).x) ^ 2 + (Fxy(I1, 2) - CurcvsPoint(j2).y) ^ 2
          If d2 <= d1 Then d1 = d2
       Next j2
       Dz = Dz + d1
   Next I1
   '显示Dz
   TxtPCABP.Text = Dz: TxtPCABP.Refresh
   '
    Dim N As Byte
   Dim xymax As Double, SumX As Double, SumY As Double
   Dim TLine As Integer
   Dim TextLine() As String
    Dim s1 As String
    Dim x() As Double, y() As Double, sqrxy() As Double
    Dim A1 As Double, a2 As Double
   'FileName = "..\data\gencurve.TXT"
  
    '(1)从文件中读到字符串数组TextLine中
    TLine = 0
    On Error GoTo WUHS
    Open HSDataFileName For Input As #1    
    TLine = 0                              
    Do Until EOF(1)
       TLine = TLine + 1
       ReDim Preserve TextLine(1 To TLine) 
       Line Input #1, TextLine(TLine)     
    Loop
    Close #1                             
    i = 0
    Do
       i = i + 1
    Loop Until (Len(Trim(TextLine(i))) <= 2 Or i = TLine)
    If i < TLine Then TLine = i - 1
    ReDim x(1 To TLine)                   
    ReDim y(1 To TLine)                    
    ReDim sqrxy(1 To TLine)
    '===================================================================
    'Sumx = 0#: Sumy = 0#
    For i = 1 To TLine
        TextLine(i) = LTrim$(RTrim$(TextLine(i)))  '
        N = InStr(TextLine(i), " ")
        x(i) = Left$(TextLine(i), N)   ': Sumx = Sumx + x(i)
        y(i) = Right$(TextLine(i), Len(TextLine(i)) - N)  ': Sumy = Sumy + y(i)
    Next i
    '平移使E?=0,并求sqrxy(i)
    
    'Sumx = Sumx / TLine: Sumy = Sumy / TLine
     For i = 1 To TLine
         x(i) = x(i) - Sumx1
         y(i) = y(i) - Sumy1
         'sqrxy(i) = Sqr(x(i) * x(i) + y(i) * (y(i)))
     Next i
    For i = 1 To TLine: x(i) = x(i) / xymax1: y(i) = y(i) / xymax1: Next i
    '
    
    '限制在0-1之间
    For i = 1 To TLine
      x(i) = (x(i) + 1) / 2: y(i) = (y(i) + 1) / 2
      CurcvsPoint(i).x = x(i): CurcvsPoint(i).y = y(i)
    Next i
   '求Dz
   Dz = 0#
   For I1 = 1 To m    '对数据点循环
       d1 = (Fxy(I1, 1) - CurcvsPoint(0).x) ^ 2 + (Fxy(I1, 2) - CurcvsPoint(0).y) ^ 2
       For j2 = 1 To TLine
          d2 = (Fxy(I1, 1) - CurcvsPoint(j2).x) ^ 2 + (Fxy(I1, 2) - CurcvsPoint(j2).y) ^ 2
          If d2 <= d1 Then d1 = d2
       Next j2
       Dz = Dz + d1
   Next I1
   '显示Dz
   TxtHS.Text = Dz: TxtHS.Refresh
   GoTo Sub_EXIT:
WUHS:
   TxtHS.Text = ""
Sub_EXIT:
End Sub


