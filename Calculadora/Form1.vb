Public Class Form1

    '---------
    'Variables
    '---------
    Dim primerNumero As Double = 0 'contendrá el primer número de la operación que se realice
    Dim segundoNumero As Double = 0 'contendrá el segundo número de la operación que se realice
    Dim numComma As String 'contendrá un valor temporal cuando el usuario introduzca una coma en el número, para evitar que escriba una segunda
    Dim operaciones(4) As Boolean 'arreglo booleano que contiene qué operación se está realizando. En orden: más, menos, multiplicación y división.
    Dim opIndex As Integer = -1 'opIndex será el valor del índice del arreglo operaciones(4)
    Dim escribiendo As Boolean 'variable booleana para determinar si se está escribiendo en la caja de texto
    Dim escribiendo2 As Boolean 'determina si es cribe en la caja de textu el segundo número de la operación
    Dim bOp As Button 'b1 contendrá el botón no numérico presionado por el usuario
    Dim bNum As Button 'Contiene el botón numérico presionado por el usuario
    Dim opPressed As Boolean 'Devuelve True si el último botón presionado fue el de una operación. False, si fue un número

    '-------
    'Eventos
    '-------
    Private Sub num1_Click(sender As Object, e As EventArgs) Handles num9.Click, num8.Click, num7.Click, num6.Click, num5.Click, num4.Click, num3.Click, num2.Click, num1.Click, num0.Click
        'La siguiente variable b contiene el botón clickeado por el usuario y según su contenido realiza la acción que indica el Select.
        bNum = DirectCast(sender, Button)
        Call escribirNumero(bNum.Text)
        btnIgual.Focus()
    End Sub

    'El siguiente evento lo usamos para verificar que el usuario no introduzca más de una coma
    'en el número, y si lo hace, que no se tenga en cuenta.
    Private Sub btnComa_Click(sender As Object, e As EventArgs) Handles btnComa.Click
        If primerNumero = 0 Or segundoNumero <> 0 Then
            numComma = txtNums.Text
            numComma += ","
            If IsNumeric(numComma) Then
                txtNums.Text += ","
                escribiendo = True
            End If
        ElseIf segundoNumero = 0 Then
            numComma = "0"
            numComma += ","
            If IsNumeric(numComma) Then
                txtNums.Text = "0,"
                segundoNumero = 1
            End If
        End If
        btnIgual.Focus()
    End Sub

    Private Sub btnMas_Click(sender As Object, e As EventArgs) Handles btnMult.Click, btnMenos.Click, btnMas.Click, btnDiv.Click
        bOp = DirectCast(sender, Button)

        If Not opPressed And (operaciones(0) Or operaciones(1) Or operaciones(2) Or operaciones(3)) Then
            Call calcular()
        End If

        Select Case bOp.Text
            Case "+"
                opIndex = 0
            Case "-"
                opIndex = 1
            Case "*"
                opIndex = 2
            Case "/"
                opIndex = 3
        End Select
        Call operacion(opIndex)
        Call asignarNumero()
        btnIgual.Focus()
    End Sub

    Private Sub btnIgual_Click(sender As Object, e As EventArgs) Handles btnIgual.Click
        Call calcular()
        primerNumero = 0
        If segundoNumero = 0 Then
            segundoNumero = 0
        Else
            segundoNumero = CDbl(txtNums.Text)
        End If
        operaciones(0) = False
        operaciones(1) = False
        operaciones(2) = False
        operaciones(3) = False
        escribiendo = False
        opIndex = -1
        btnIgual.Focus()
    End Sub

    Private Sub btnC_Click(sender As Object, e As EventArgs) Handles btnC.Click
        txtNums.Text = 0
        primerNumero = 0
        segundoNumero = 0
        operaciones(0) = False
        operaciones(1) = False
        operaciones(2) = False
        operaciones(3) = False
        escribiendo = False
        escribiendo2 = False
        numComma = "0"
        opIndex = -1
        btnIgual.Focus()
    End Sub

    Private Sub btnBorrar_Click(sender As Object, e As EventArgs) Handles btnBorrar.Click
        If IsNumeric(txtNums.Text) Then 'Revisa que el contenido de la caja de texto sea numérico
            If txtNums.Text <> 0 Then 'Comprueba que sea distinto de 0
                If txtNums.Text.Length > 1 Then 'Revisa la longitud de la cadena de la caja de texto
                    txtNums.Text = txtNums.Text.Remove(txtNums.Text.Length - 1) 'Si hay más de un caracter, borra el último
                Else
                    txtNums.Text = 0 'Si hay un solo caracter, escribe 0 en la caja de texto
                    escribiendo = False
                    escribiendo2 = False
                End If
            Else    'Para los casos en los que se haya escrito "0," en la caja de texto, borra la coma.
                txtNums.Text = 0
                escribiendo = False
                escribiendo2 = False
            End If
        Else    'Si el contenido de la caja de texto no es numérico ("Error - División por 0"), llama al evento del botón C.
            Call btnC_Click(e, e)
        End If
        btnIgual.Focus()
    End Sub

    Private Sub btnNeg_Click(sender As Object, e As EventArgs) Handles btnNeg.Click
        If IsNumeric(txtNums.Text) Then
            txtNums.Text = txtNums.Text * -1
        End If
        btnIgual.Focus()
    End Sub

    '------------------
    'Eventos de teclado
    '------------------
    'Form1_KeyPress: usado para capturar la tecla presionada por el usuario
    Private Sub Form1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles MyBase.KeyPress
        'El siguiente If establece lo que sucede al presionar uno de los 10 dígitos en el teclado númerico.
        'Con Asc(e.KeyChar) obtenemos el código Ascii de la tecla presionada, mientras que con Chr(Asc(e.KeyChar))
        'lo convertimos a formato caracter para que sea escrito en el textbox
        If Asc(e.KeyChar) >= 48 And Asc(e.KeyChar) <= 57 Then
            Call escribirNumero(Chr(Asc(e.KeyChar)))
        End If

        'Si no fue presionado uno de los 10 dígitos, mediante el siguiente If se establece si se presionó la
        'tecla de alguna de las 4 operaciones principales. 
        If Asc(e.KeyChar) = 42 Or Asc(e.KeyChar) = 43 Or Asc(e.KeyChar) = 45 Or Asc(e.KeyChar) = 47 Then
            If operaciones(0) Or operaciones(1) Or operaciones(2) Or operaciones(3) Then
                Call calcular()
            End If

            Select Case Asc(e.KeyChar)
                Case 43 '+
                    opIndex = 0
                Case 45 '-
                    opIndex = 1
                Case 42 '*
                    opIndex = 2
                Case 47 '/
                    opIndex = 3
            End Select
            Call operacion(opIndex)
            Call asignarNumero()
        End If

        'El siguiente If captura si se presiona un punto o una coma para escribir un número decimal
        If Asc(e.KeyChar) = 44 Or Asc(e.KeyChar) = 46 Then
            Call btnComa_Click(e, e)
        End If

        'Por último, se analiza si fue presionada la tecla Backspace (Retroceso) para borrar.
        If Asc(e.KeyChar) = 8 Then
            Call btnBorrar_Click(e, e)
        End If

    End Sub

    'Evento que captura lo que sucede al presionar la tecla Enter
    Private Sub Form1_PreviewKeyDown(sender As Object, e As PreviewKeyDownEventArgs) Handles MyBase.PreviewKeyDown
        If e.KeyCode = Keys.Enter Then
            Call btnIgual_Click(e, e)
        End If
    End Sub


    '-------
    'Métodos
    '-------
    'escribirNumero([num]) escribe los números presionados, sea mediante clic o teclado, en la caja de texto
    Sub escribirNumero([num])
        If Not escribiendo Then
            If primerNumero = 0 Then
                txtNums.Text = num
                primerNumero = CDbl(txtNums.Text)
            Else
                txtNums.Text += num
            End If
        Else
            If Not escribiendo2 Then
                txtNums.Text = num
                segundoNumero = CDbl(txtNums.Text)
                If num <> "0" Then
                    escribiendo2 = True
                End If
            Else
                txtNums.Text += num
                segundoNumero = CDbl(txtNums.Text)
            End If
        End If
        opPressed = False
    End Sub

    'operacion([opIndex]) captura cuál de las 4 operaciones va a realizarse según lo que haya indicado el usuario y
    'vuelve True el valor de esa operacion en el arreglo y False el resto. Luego, establece los valores de
    'primerNumero o segundoNumero según corresponda.
    Sub operacion([opIndex])
        For i = 0 To 3 Step 1
            If i = opIndex Then
                operaciones(i) = True
            Else
                operaciones(i) = False
            End If
        Next
        opPressed = True
        escribiendo = True
        escribiendo2 = False
    End Sub

    'asignarNumero() establece los valores de las variables primerNumero o segundoNumero según corresponda
    Sub asignarNumero()
        If IsNumeric(txtNums.Text) Then
            If escribiendo Then
                primerNumero = CDbl(txtNums.Text)
                segundoNumero = 0
            Else
                If Not opPressed Then
                    segundoNumero = CDbl(txtNums.Text)
                Else
                    segundoNumero = 0
                End If
            End If
        End If
    End Sub

    'calcular() realiza el cálculo correspondiente
    Sub calcular()
        If operaciones(0) Then
            txtNums.Text = primerNumero + segundoNumero
        ElseIf operaciones(1) Then
            txtNums.Text = primerNumero - segundoNumero
        ElseIf operaciones(2) Then
            txtNums.Text = primerNumero * segundoNumero
        ElseIf operaciones(3) Then
            If segundoNumero = 0 Then
                txtNums.Text = "Error - División por 0"
                primerNumero = 0
                segundoNumero = 0
                escribiendo = False
                escribiendo2 = False
            Else
                txtNums.Text = primerNumero / segundoNumero
            End If
        End If
        primerNumero = 0
    End Sub

End Class
