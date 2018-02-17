Public Class Form1
    Private Structure DatStruct
        Public Placa As String
        Public Marca_carro As String
        Public Cedula As Long
        Public Nombre As String
        Public Apellido As String
        Public Tipo_carro As String
    End Structure
    Dim Datpers() As DatStruct
    Dim Index As Int32
    Sub limpiador()
        TextBox1.Text = ""
        ComboBox1.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        TextBox5.Text = ""
    End Sub
    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Label7.Text = ""
        Label8.Text = ""
    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        Label8.Text = ""
        Timer1.Enabled = False
    End Sub
    Private Sub Label8_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Label8.TextChanged
        Timer1.Enabled = True
    End Sub
    'Aceptar
    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        If TextBox1.Text = "" Or ComboBox1.Text = "" Or TextBox3.Text = "" Or TextBox4.Text = "" Or TextBox5.Text = "" Then
            MessageBox.Show("NO SE PUEDEN DEJAR LOS CAMPOS EN BLANCO")
        Else
            Try
                ReDim Preserve Datpers(UBound(Datpers) + 1)
            Catch ex As Exception
                ReDim Preserve Datpers(1)
            End Try
            Index = UBound(Datpers)
            Label7.Text = UBound(Datpers)
            With Datpers(UBound(Datpers))
                .Placa = TextBox1.Text
                .Marca_carro = ComboBox1.Text
                .Cedula = TextBox3.Text
                .Nombre = TextBox4.Text
                .Apellido = TextBox5.Text
                .Tipo_carro = CheckBox1.Text
                .Tipo_carro = CheckBox2.Text
                Label8.Text = "¡Datos almacenados con éxito!"
                Call limpiador()
                TextBox1.Focus()

            End With

        End If
    End Sub
    'Anterior
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If Index > 1 Then
            Index -= 1
            Label7.Text = Index & " / " & UBound(Datpers)
            With Datpers(Index)
                TextBox1.Text = .Placa
                ComboBox1.Text = .Marca_carro
                TextBox3.Text = .Cedula
                TextBox4.Text = .Nombre
                TextBox5.Text = .Apellido
                .Tipo_carro = CheckBox1.Text
                .Tipo_carro = CheckBox2.Text
            End With
        Else
            Label8.Text = "¡Inicio de Archivo Encontrado!"
            Beep()
        End If
    End Sub
    'Siguiente
    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        If Index < UBound(Datpers) Then
            Index += 1
            Label7.Text = Index & " / " & UBound(Datpers)
            With Datpers(Index)
                TextBox1.Text = .Placa
                ComboBox1.Text = .Marca_carro
                TextBox3.Text = .Cedula
                TextBox4.Text = .Nombre
                TextBox5.Text = .Apellido
                CheckBox1.Text = .Tipo_carro
                CheckBox2.Text = .Tipo_carro
            End With
        Else
            Label8.Text = "¡Final de Archivo Encontrado!"
            Beep()
        End If
    End Sub
    'Salvar
    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        Dim FleNme As String = Mid(Application.ExecutablePath, 1, InStrRev(Application.ExecutablePath, "\")) & "Datpers.txt"
        FileOpen(1, FleNme, OpenMode.Binary)
        FilePut(1, UBound(Datpers))
        FilePut(1, Datpers)
        FileClose(1)
    End Sub
    'Cargar
    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Dim FleNme As String = Mid(Application.ExecutablePath, 1, InStrRev(Application.ExecutablePath, "\")) & "Datpers.Txt"
        FileOpen(1, FleNme, OpenMode.Binary)
        FileGet(1, Index)
        ReDim Datpers(Index)
        FileGet(1, Datpers)
        FileClose(1)
        Label7.Text = UBound(Datpers)
    End Sub
    'Finalizar
    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        End
    End Sub
    'Caja de texto de PLACA
    Private Sub TextBox1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox1.KeyPress
        Select Case Asc(e.KeyChar)
            Case Is = 8
            Case Is = 13
                ComboBox1.Focus()
            Case Is = 27
            Case Else
                If InStr("1234567890ABCDEFGHIJKLMNÑPQRSTVYXZ", UCase(e.KeyChar)) = 0 Then
                    Beep()
                    e.Handled = True
                Else : e.KeyChar = UCase(e.KeyChar)
                End If
        End Select
    End Sub

    Private Sub TextBox2_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)
        Select Case Asc(e.KeyChar)
            Case Is = 8
            Case Is = 13 'ENTER
                TextBox3.Focus()
            Case Is = 27 'ESC
                TextBox1.Focus()
            Case Else
                If InStr("ABCDEFGHIJKLMNÑOPQRSTUVWXYZ '.", UCase(e.KeyChar)) = 0 Then
                    Beep()
                    e.Handled = True
                Else
                    e.KeyChar = UCase(e.KeyChar)
                End If
        End Select
    End Sub

    Private Sub TextBox3_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox3.KeyPress
        Select Case Asc(e.KeyChar)
            Case Is = 8
            Case Is = 13
                TextBox4.Focus()
            Case Is = 27
                ComboBox1.Focus()
            Case Else
                If InStr("1234567890", e.KeyChar) = 0 Then
                    Beep()
                    e.Handled = True
                End If
        End Select
    End Sub

    Private Sub TextBox4_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox4.KeyPress
        Select Case Asc(e.KeyChar)
            Case Is = 8
            Case Is = 13
                TextBox5.Focus()
            Case Is = 27
                TextBox3.Focus()
            Case Else
                If InStr("ABCDEFGHIJKLMNÑOPQRSTUVWXYZ '.", UCase(e.KeyChar)) = 0 Then
                    Beep()
                    e.Handled = True
                Else
                    If TextBox4.SelectionStart = 0 Then
                        e.KeyChar = UCase(e.KeyChar)
                    Else
                        If Mid(TextBox4.Text, TextBox4.SelectionStart, 1) = " " Then
                            e.KeyChar = UCase(e.KeyChar)
                        Else
                            e.KeyChar = LCase(e.KeyChar)
                        End If
                    End If
                End If
        End Select
    End Sub

    Private Sub TextBox5_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox5.KeyPress
        Select Case Asc(e.KeyChar)
            Case Is = 8
            Case Is = 13
                Button6.Focus()
            Case Is = 27
                TextBox3.Focus()
            Case Else
                If InStr("ABCDEFGHIJKLMNÑOPQRSTUVWXYZ '.", UCase(e.KeyChar)) = 0 Then
                    Beep()
                    e.Handled = True
                Else
                    If TextBox5.SelectionStart = 0 Then
                        e.KeyChar = UCase(e.KeyChar)
                    Else
                        If Mid(TextBox5.Text, TextBox5.SelectionStart, 1) = " " Then
                            e.KeyChar = UCase(e.KeyChar)
                        Else
                            e.KeyChar = LCase(e.KeyChar)
                        End If
                    End If
                End If
        End Select
    End Sub


    Private Sub TextBox1_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles TextBox1.MouseDown
        Clipboard.Clear()
        Clipboard.SetText(" ")
    End Sub
    Private Sub TextBox2_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs)
        Clipboard.Clear()
        Clipboard.SetText(" ")
    End Sub
    Private Sub TextBox3_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles TextBox3.MouseDown
        Clipboard.Clear()
        Clipboard.SetText(" ")
    End Sub
    Private Sub TextBox4_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles TextBox4.MouseDown
        Clipboard.Clear()
        Clipboard.SetText(" ")
    End Sub
    Private Sub TextBox5_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles TextBox5.MouseDown
        Clipboard.Clear()
        Clipboard.SetText(" ")
    End Sub

    Private Sub CheckBox1_CheckedChanged_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked = True Then
            CheckBox2.Checked = False
        End If
    End Sub

    Private Sub CheckBox2_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox2.CheckedChanged
        If CheckBox2.Checked = True Then
            CheckBox1.Checked = False
        End If
    End Sub
End Class
