Public Class Form1

    Function CP(ByVal S As String) As Double
        S = Replace(S, " ", "")
        Dim t As String = ""
        Dim f As Integer
        For Each i In S '
            Select Case i '分割最外層括號 x-(y+z)*2(4-2) → x- (y+z) *2* (4-2)
                Case "("
                    If t <> "" And f = 0 Then
                        Select Case Microsoft.VisualBasic.Right(t, 1)
                            Case 0 To 9, " " : t &= "*" 'x(y) > x*(y)
                        End Select
                        t &= " "
                    End If
                    f += 1
                    t &= i
                Case ")"
                    f -= 1
                    If f = 0 Then t &= ") " Else t &= i
                Case Else
                    t &= i
            End Select
        Next
        Dim M() As String = Split(t, " ")
        For i = 0 To UBound(M) '計算括號內的值，傳回值覆蓋式子
            If InStr(M(i), "(") > 0 Then M(i) = CP(Mid(M(i), 2, Len(M(i)) - 2))
        Next

        Dim A As String = Join(M.ToArray, "") '合併括號計算後的值
        A = Replace(A, "--", "+") '負負得正
        Dim M2() As String = Split(A, "+")
        Dim ans As Double = 0
        For Each i In M2
            ans += Subtract(i) '加法運算(優先度最低)
        Next

        Return ans
    End Function

    Function Subtract(ByVal S As String) As Double
        S = Replace(S, "-", " -")
        S = Replace(S, "* -", "*-")
        S = Replace(S, "/ -", "/-")
        S = Replace(S, "^ -", "^-")
        Dim M() As String = Split(S, " ")
        Dim ans As Double = Multiply(M(0))
        For i = 1 To UBound(M)
            ans += Multiply(M(i)) '減法運算
        Next
        Return ans
    End Function

    Function Multiply(ByVal S As String) As Double
        Dim M() As String = Split(S, "*")
        Dim ans As Double = 1
        For Each i In M
            ans *= Divide(i) '乘法運算
        Next
        Return ans
    End Function

    Function Divide(ByVal S As String) As Double
        Dim M() As String = Split(S, "/")
        Dim ans As Double = Power(M(0))
        For i = 1 To UBound(M)
            ans /= Power(M(i)) '除法運算
        Next
        Return ans
    End Function

    Function Power(ByVal S As String) As Double
        Dim M() As String = Split(S, "^")
        Dim ans As Double = Val(M(UBound(M)))
        For i = UBound(M) - 1 To 0 Step -1
            ans = Val(M(i)) ^ ans '冪運算(優先度最高)
        Next
        Return ans
    End Function

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim st As String = TextBox1.Text
        Label2.Text = CP(st)
    End Sub
End Class
