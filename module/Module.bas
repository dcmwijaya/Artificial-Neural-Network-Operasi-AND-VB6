Attribute VB_Name = "Module"
Type Data
'Variabel global X-Y baris 1'
Xa0 As Integer
Xa1 As Integer
Xa2 As Integer
Y0 As Integer

'Variabel global X-Y baris 2'
Xb0 As Integer
Xb1 As Integer
Xb2 As Integer
Y1 As Integer

'Variabel global X-Y baris 3'
Xc0 As Integer
Xc1 As Integer
Xc2 As Integer
Y2 As Integer

'Variabel global X-Y baris 4'
Xd0 As Integer
Xd1 As Integer
Xd2 As Integer
Y3 As Integer

'Variabel global W (Bobot) dan Threshold'
W0 As Double
W1 As Double
W2 As Double
T As Double

'Variabel global Hasil'
H0 As Double
H1 As Double
H2 As Double
H3 As Double

'Variabel global Output'
Op0 As Double
Op1 As Double
Op2 As Double
Op3 As Double

'Variabel global Error'
Err0 As Double
Err1 As Double
Err2 As Double
Err3 As Double

'Variabel global Update Weight(Bias)'
UWa0 As Double
UWa1 As Double
UWa2 As Double
UWb0 As Double
UWb1 As Double
UWb2 As Double
UWc0 As Double
UWc1 As Double
UWc2 As Double
UWd0 As Double
UWd1 As Double
UWd2 As Double

End Type

Global IO As Data

