Public Tfx() As Double    '投影指标
Public Fxy() As Double    '数据点矩阵
                          
Public xymax1 As Double, Sumx1 As Double, Sumy1 As Double
Public Learn_Pause As String


'数据点类型定义
Public Type xy
     x As Double
     y As Double
End Type
Public CurcvsPoint(10000) As xy

Public TDataFileName As String
Public HSDataFileName As String
Public GenDataFileName As String


