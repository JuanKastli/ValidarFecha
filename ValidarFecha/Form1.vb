Public Class Form1


    Private Sub TextBox1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox1.KeyPress
        If Not e.KeyChar = "/" And Not Char.IsNumber(e.KeyChar) And Not Char.IsControl(e.KeyChar) Then
            e.Handled = True
            Exit Sub
        End If
        'si la tecla pulsada es / y la posicion del cursor es distinto de 2 y 
        'la posiocion del cursor es distinto de 5,entonces habilitar handled y salir
        Dim pos As Integer = TextBox1.SelectionStart
        If e.KeyChar = "/" And pos <> 2 And pos <> 5 Then
            e.Handled = False
        End If
        'si la tecla pulsada es un numero y (pos es 2 o 5 ) entonces habilitar handled y salir 
        If Char.IsNumber(e.KeyChar) And (pos = 2 Or pos = 5) Then
            e.Handled = True
            Exit Sub
         
        End If
    End Sub

    Private Sub TextBox1_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox1.LostFocus
        Dim dia As Integer = CInt(TextBox1.Text.Substring(0, 1))
        Dim mes As Integer = CInt(TextBox1.Text.Substring(3, 4))
        Dim año As Integer = CInt(TextBox1.Text.Substring(6, 9))
        Select Case mes
            Case 1, 3, 5, 7, 8, 10, 12
                If dia < 1 Or dia > 31 Then
                    MsgBox("Mes no admitido")
                    TextBox1.Focus()
                    Exit Sub
                End If
            Case 4, 6, 9, 11
                If dia < 1 Or dia > 30 Then
                    MsgBox("Dia no admitido")
                    TextBox1.Focus()
                End If
            Case 2
                If año Mod 4 = 0 Then
                    If dia < 1 Or dia > 28 Then
                        MsgBox("mes no admitido")
                        TextBox1.Focus()
                    End If
                Else
                    If dia < 1 Or dia > 29 Then
                        MsgBox("Mes no admitido")
                        TextBox1.Focus()
                    End If
                End If
        End Select
    End Sub

    Private Sub TextBox1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox1.TextChanged
        Dim dia As Integer = CInt(TextBox1.Text.Substring(0, 1))
        Dim mes As Integer = CInt(TextBox1.Text.Substring(3, 4))
        Dim año As Integer = CInt(TextBox1.Text.Substring(6, 9))
        If dia < 1 Or dia > 31 Then
            MsgBox("Dia no valido")
            Exit Sub
        End If
        'si el mes es menor a 01 o el mes es mayor a 12 salir
        If mes < 1 Or mes > 12 Then
            MsgBox("Mes no valido")
            TextBox1.Focus()
            Exit Sub
        End If
        'si el año es menor a 1980 o el año es mayor 2018 salir
        If año < 1980 Or año > 2018 Then
            MsgBox("Año no válido")
            TextBox1.Focus()
            Exit Sub
        End If

    End Sub
End Class
