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
Dim wc1() As Double  '������
Dim wc2() As Double  '����ֵ
Dim wcp As Double    '����ֵƽ��
Dim R2 As Double     '���ϵ����ƽ��
'--------------------------------------------------------
Dim I1 As Long, I2 As Long, i3 As Long  '��ʱʹ��
Dim d1 As Double
Dim s1 As String
'---------------------------------------------------------
Const pi = 3.141592

Dim N As Long                  '����㵥Ԫ����:n
Dim i As Long                  '��������:i=1 to n

Dim p As Long                  '������(�м��)��Ԫ����:p
Dim j As Long                  '������(�м�㣩����:j=1 to p

Dim q As Long                  '����㵥Ԫ����:q
Dim t As Long                  '��������:t=1 to q

Dim m As Long                  'ѧϰģʽ����:m
Dim k As Long                  'ѧϰģʽ����:k=1 to m

Dim A() As Double              
Dim y() As Double              

Dim w() As Double              '��������м������Ȩ����:w(i,j)
Dim v() As Double              '�м�������������Ȩ����:v(j,t)
Dim O() As Double              '�м�����Ԫ�����ֵ:O(j)
Dim r() As Double              '��������Ԫ�����ֵ:r(t)

Dim s() As Double    '�м�����Ԫ������s(k,j)
Dim B() As Double    '�м�����Ԫ�����b(k,j)
Dim l() As Double    '��������Ԫ������l(k,t)
Dim c() As Double    '��������Ԫ�����c(k,t)
Dim d() As Double    '��������Ԫ��һ�㻯���d(k,t)
Dim e() As Double    '�м�����Ԫ��һ�㻯���e(k,j)

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

   File1.Path = "..\Tdata"       '�����ļ�·��  
   Learn_Pause = "ON"
   Pause.Caption = "��ͣѧϰ"

End Sub


Private Sub File1_Click()        'ѡ�������ļ�
    TDataFileName = File1.Path & "\" & File1.FileName
    HSDataFileName = "..\HS\" & File1.FileName
    GenDataFileName = "..\gencurve\" & File1.FileName

    Text1.Text = "�Ѿ�ѡ��������ļ�Ϊ:" & "   ..\Tdata\" & File1.FileName
End Sub


Private Sub CmdPara_Click()      '���趨ȷ�ϡ���ť
   p = Val(Text6.Text)           '�ɵ���������(�м��)��Ԫ����:p
   StudyNumber = Val(Text5.Text) 'ѧϰ�������ɻ�ȡ��ͬѧϰ���������ս��ֵ��
   Text5.BackColor = &HFFC0FF
   Text6.BackColor = &HFFC0FF
End Sub
   
Private Sub CmdStudySatrt_Click()    '����"��ʼѧϰ"��ť
    Call PraSub
    '(1)��������Ȩw(i,j)��v(j,t)������ֵO(j)��r(t)����(-0.1,+0.1)������ֵ
    
    For i = 1 To N
       For j = 1 To p: w(i, j) = SgnRnd(Rnd()) * 0.1 * Rnd(): Next j
    Next i
    
    For j = 1 To p
       For t = 1 To q: v(j, t) = SgnRnd(Rnd()) * 0.1 * Rnd(): Next t
    Next j
    
    For j = 1 To p: O(j) = SgnRnd(Rnd()) * 0.1 * Rnd(): Next j
    
    For t = 1 To q: r(t) = SgnRnd(Rnd()) * 0.1 * Rnd(): Next t
    
    '(2)���ѡȡһ��ѧϰģʽ���ṩ������,��kѡȡ1��m֮�������ֵ(����)
    'k = Int(m * (Rnd())) + 1
   t1 = Timer()
   For LN = 1 To StudyNumber     'ѧϰ����ѭ��
        DoEvents
       
        If Learn_Pause = "OFF" Then
           Do
             DoEvents
           Loop Until Learn_Pause <> "OFF"
        End If
       For k = 1 To m         'ѧϰģʽ��ѭ��
            '(3)��a(k,i)(i=1 to n),����Ȩw(i,j),��ֵO(j)�����м����Ԫ������s(k,j),ͨ��
            For j = 1 To p
                s(k, j) = 0
                For i = 1 To N: s(k, j) = s(k, j) + w(i, j) * A(k, i): Next i
                s(k, j) = s(k, j) - O(j)                                      '(4.32)
                B(k, j) = fs(s(k, j))                                         '(4.33)
            Next j
            '(4)���м�����Ԫ�����b(k,j)(j=1 to p),����Ȩv(j,t),��ֵr(j)������������Ԫ������l(k,t),ͨ��
            'S����������������Ԫ�������Ӧc(k,t)
            For t = 1 To q
                l(k, t) = 0
                For j = 1 To p: l(k, t) = l(k, t) + v(j, t) * B(k, j): Next j
                l(k, t) = l(k, t) - r(t)
                c(k, t) = fs(l(k, t))
            Next t
            '(5)��ϣ�����ģʽy(k,t) (t=1 to q),����ʵ�����c(k,t),���������ĸ���Ԫ��һ�㻯���d(k,t)
            For t = 1 To q
                d(k, t) = (y(k, t) - c(k, t)) * c(k, t) * (1 - c(k, t))     '(4.36)
            Next t
            '(6)������Ȩv(j,t),�����һ�㻯���d(k,t),�м������b(k,j),�����м�����Ԫ��һ�㻯���e(k,j)
            For j = 1 To p
                d1 = 0: For t = 1 To q: d1 = d1 + d(k, t) * v(j, t): Next t
                e(k, j) = d1 * B(k, j) * (1 - B(k, j))                      '(4.37)
            Next j
            '(7)�������һ�㻯���d(k,t),�м������b(k,j)��������Ȩv(j,t)����ֵr(t)
            For j = 1 To p
                For t = 1 To q
                    v(j, t) = v(j, t) + RF * d(k, t) * B(k, j)  '(4.38)
                Next t
            Next j
            For t = 1 To q
                r(t) = r(t) - RF * d(k, t)                      '(4.39)
            Next t
            '(8)���м��һ�㻯���e(k,j),������a(k,i)��������Ȩw(i,j)����ֵO(j)
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
          '������EE
          EE = 0#
           For k = 1 To m
             For t = 1 To q
			   EE = (y(k, t) - c(k, t)) * (y(k, t) - c(k, t)) / 2         '�������
			   'EE =  EE + (y(k, t) - c(k, t)) * (y(k, t) - c(k, t)) / 2  '��ȫ�����
            Next t
          Next k
          TxtEE.Text = Format$(EE, "0.00000000000000"): TxtEE.Refresh    '��ʾ
          TxtLearnNumber.Text = Format$(LN, "############0"): TxtLearnNumber.Refresh
          TxtTimer.Text = Format$(Timer() - t1, "#######0") & "��": TxtTimer.Refresh
          '��ʾ����Ȩw(i,j)
          s1 = ""
          For i = 1 To N
              For j = 1 To p
              s1 = s1 & "��(" & Format$(i, "0") & "," & Format$(j, "0") & ")=" & Format$(w(i, j), "###0.00000000") & Chr(&HD) & Chr(&HA)
              Next j
          Next i
          TxtW.Text = s1
           '��ʾ��ֵO(j)
          s1 = ""
          For j = 1 To p
              s1 = s1 & "��(" & Format$(j, "0") & ")=" & Format$(O(j), "###0.00000000") & Chr(&HD) & Chr(&HA)
          Next j
          TxtO.Text = s1
          '��ʾ����Ȩv(j,t)
          s1 = ""
          For j = 1 To p
              For t = 1 To q
              s1 = s1 & "��(" & Format$(j, "0") & "," & Format$(t, "0") & ")=" & Format$(v(j, t), "###0.00000000") & Chr(&HD) & Chr(&HA)
              Next t
          Next j
          Txtv.Text = s1
          '��ʾ��ֵ��
          s1 = ""
          For t = 1 To q
              s1 = s1 & "��(" & Format$(t, "0") & ")=" & Format$(r(t), "###0.00000000") & Chr(&HD) & Chr(&HA)
          Next t
          Txtr.Text = s1
      End If
   Next LN     '��ѧϰ����ѭ��
End Sub

Private Sub PraSub()             '��ʼ��
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
   '----------�ع�����ѧϰģʽ��-----------------------------
   
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
    RF = 0.5: BT = 0.5     'ѧϰ��1,2
End Sub


Private Function SgnRnd(x As Double) As Double  '����������ź���
   If x < 0.5 Then
      SgnRnd = -1
   Else
      SgnRnd = 1
   End If
End Function


Private Function fs(x As Double) As Double     '����S����
   fs = 1 / (1 + Exp(-x))
End Function


Private Sub CmdDataPoint_Click()      '�����ݵ�
   Call DrawZoBiao    '������
      '�����ݵ�
   For I1 = LBound(Tfx) To UBound(Tfx) - 2 Step 1
   'For i1 = 0 To 0.9999
       'Call DrawLargePoint(y(I1, 1), y(I1, 2), vbBlack)
       Call DrawLargePoint(y(I1, 1), y(I1, 2), &H808080)
       
   Next I1
    'Call ReadFile("PCA_Ti_n.TXT")
End Sub



Private Sub Cndstart_Click()     '"�������"��ť---ͳһ��������ѧ���ʽΪ�˴�
   Dim I1 As Integer
   Dim d1 As Double
   Dim N2 As Integer, M2 As Integer
       Dim Ir As Integer, Jr As Integer, Nr As Integer
    Dim Sar() As String
    
    '----------------------------------------------------
    Open TDataFileName For Input As #1          '����ȷ���ļ���,���ļ�
    Ir = 0                                 '�ļ���������ֵ=0
    Do Until EOF(1)
       Ir = Ir + 1
       ReDim Preserve Sar(1 To Ir) As String '���¶����ַ������������±�
       Line Input #1, Sar(Ir)      '��һ�С�>������
    Loop
    Close #1                               '�ر��ļ�
    '
    '���¶���x()��y()����,�����Ǹ�ֵ
    ReDim X2(1 To Ir - 1) As Double                 '���¶�������±�
    ReDim Y2(1 To Ir - 1) As Double
    For Jr = 1 To Ir - 1
        Sar(Jr) = LTrim$(RTrim$(Sar(Jr)))
        X2(Jr) = Mid$(Sar(Jr), 14, 13)  'x (��һ����)
        Y2(Jr) = Mid$(Sar(Jr), 28, 13)  'y (�ڶ�����)
    Next Jr
    
    '----------------------------------------------------
    '
    M2 = 20
    N2 = UBound(Y2) / M2 - 1
       ReDim RX(0 To N2)
    ReDim RY(0 To N2)

    Call DrawZoBiao    '������
    
    For I1 = 0 To N2
        d1 = I1 * M2 + 1
        RX(I1) = (X2(d1) + X2(d1 + 1) + X2(d1 + 2) + X2(d1 + 3)) / 4
        RY(I1) = (Y2(d1) + Y2(d1 + 1) + Y2(d1 + 2) + Y2(d1 + 3)) / 4
        'Call DrawZPoint(RX(I1), RY(I1), vbRed)
    Next I1
    
    For I1 = 1 To UBound(X2)   '�����ݵ�ѭ��
       Call DrawLargePoint(X2(I1), Y2(I1), &H808080)
    Next I1
   
   Call DrawZoBiao    '������
  
   '����������
   For d1 = Tfx(1) To Tfx(m) Step 0.001
   'For d1 = 0.001 To 0.999 Step 0.002
       Call hx(d1)
       Call DrawXYPiont2(cc(1), cc(2), vbBlue)
   Next d1
   
            For I1 = 0 To N2
             Call DrawZPoint(RX(I1), RY(I1), vbRed)
            'Call DrawZPoint(RX(I1), RY(I1), vbRed)
         Next I1
   '========������ع�=====����R
End Sub


Private Sub hx(xb)     '"����-���߻ػ�"
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
       MsgBox ("δѡ�������ļ�")
       Exit Sub
    End If
    Open FileName For Input As #1          '����ȷ���ļ���,���ļ�
    Ir = 0                                 '�ļ���������ֵ=0
    Do Until EOF(1)
       Ir = Ir + 1
       ReDim Preserve Sar(1 To Ir) As String 
       Line Input #1, Sar(Ir)      '��һ�С�>������
    Loop
    Close #1                               '�ر��ļ�
    '
    '���¶���x()��y()����,�����Ǹ�ֵ
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



Private Sub CmdDistance_Click()   '���ܾ��� deta f
    Dim i As Integer, I1 As Integer, I2 As Integer
    Dim j As Integer, J1 As Integer, j2 As Integer
    Dim t1 As Double
    Dim d1 As Double, d2 As Double, Dz As Double
    Dim t10 As Double, t11 As Double
    '��PCA-BP���ߵ�->CurcvsPoint
   For J1 = 0 To 9999
       t1 = J1 / 10000
       Call hx(t1)
       CurcvsPoint(J1).x = cc(1): CurcvsPoint(J1).y = cc(2)
   Next J1
   '��Dz
   Dz = 0#
   For I1 = 1 To m    '�����ݵ�ѭ��
       d1 = (Fxy(I1, 1) - CurcvsPoint(0).x) ^ 2 + (Fxy(I1, 2) - CurcvsPoint(0).y) ^ 2
       For j2 = 1 To 9999
          d2 = (Fxy(I1, 1) - CurcvsPoint(j2).x) ^ 2 + (Fxy(I1, 2) - CurcvsPoint(j2).y) ^ 2
          If d2 <= d1 Then d1 = d2
       Next j2
       Dz = Dz + d1
   Next I1
   '��ʾDz
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
  
    '(1)���ļ��ж����ַ�������TextLine��
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
    'ƽ��ʹE?=0,����sqrxy(i)
    
    'Sumx = Sumx / TLine: Sumy = Sumy / TLine
     For i = 1 To TLine
         x(i) = x(i) - Sumx1
         y(i) = y(i) - Sumy1
         'sqrxy(i) = Sqr(x(i) * x(i) + y(i) * (y(i)))
     Next i
    For i = 1 To TLine: x(i) = x(i) / xymax1: y(i) = y(i) / xymax1: Next i
    '
    
    '������0-1֮��
    For i = 1 To TLine
      x(i) = (x(i) + 1) / 2: y(i) = (y(i) + 1) / 2
      CurcvsPoint(i).x = x(i): CurcvsPoint(i).y = y(i)
    Next i
   '��Dz
   Dz = 0#
   For I1 = 1 To m    '�����ݵ�ѭ��
       d1 = (Fxy(I1, 1) - CurcvsPoint(0).x) ^ 2 + (Fxy(I1, 2) - CurcvsPoint(0).y) ^ 2
       For j2 = 1 To TLine
          d2 = (Fxy(I1, 1) - CurcvsPoint(j2).x) ^ 2 + (Fxy(I1, 2) - CurcvsPoint(j2).y) ^ 2
          If d2 <= d1 Then d1 = d2
       Next j2
       Dz = Dz + d1
   Next I1
   '��ʾDz
   TxtHS.Text = Dz: TxtHS.Refresh
   GoTo Sub_EXIT:
WUHS:
   TxtHS.Text = ""
Sub_EXIT:
End Sub


