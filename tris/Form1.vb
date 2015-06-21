Public Class Form1
    Const MAXC As Integer = 2
    Const MAXF As Integer = 1
    Const MAXP As Integer = 500
    Dim crea, matrice(MAXC, MAXC), px(MAXF, MAXP), py(MAXF, MAXP), numP(MAXF), scia(MAXF, MAXP), errore(MAXF), mossa As Integer
    Dim premuto As Boolean
    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        reset()
    End Sub

    Private Sub pcb1_MouseDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles pcb1.MouseDown, pcb2.MouseDown, pcb3.MouseDown, pcb4.MouseDown, pcb5.MouseDown, pcb6.MouseDown, pcb7.MouseDown, pcb8.MouseDown, pcb9.MouseDown
        Dim i, j As Integer

        For i = 0 To MAXF Step 1
            errore(i) = 0
            For j = 0 To MAXP Step 1
                scia(i, j) = 0
            Next
        Next
        premuto = True
        calcola(e.X, e.Y)
    End Sub

    Private Sub pcb1_MouseMove(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles pcb1.MouseMove, pcb2.MouseMove, pcb3.MouseMove, pcb4.MouseMove, pcb5.MouseMove, pcb6.MouseMove, pcb7.MouseMove, pcb8.MouseMove, pcb9.MouseMove
        Dim pcbImage As PictureBox
        Dim grafico As Graphics
        pcbImage = CType(sender, PictureBox)
        If premuto = True Then
            calcola(e.X, e.Y)
            grafico = pcbImage.CreateGraphics()
            grafico.DrawEllipse(Pens.Black, e.X, e.Y, 1, 1)
        End If
    End Sub
    Private Sub pcb1_MouseUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles pcb1.MouseUp, pcb2.MouseUp, pcb3.MouseUp, pcb4.MouseUp, pcb5.MouseUp, pcb6.MouseUp, pcb7.MouseUp, pcb8.MouseUp, pcb9.MouseUp
        Dim i, j, rig, col As Integer
        Dim pcbImage As PictureBox
        Dim grafico As Graphics
        pcbImage = CType(sender, PictureBox)
        grafico = pcbImage.CreateGraphics

        premuto = False
        grafico.Clear(Color.White)
        If crea = MAXF + 1 Then
            For i = 0 To MAXF Step 1

                j = 0
                While (j < numP(i) And errore(i) = 0)
                    If scia(i, j) = 0 Then
                        errore(i) = 1
                    Else
                        j += 1
                    End If
                End While

                If errore(i) = 0 Then
                    If i = 0 Then
                        pcbImage.Image = Image.FromFile("X.bmp")
                    Else
                        pcbImage.Image = Image.FromFile("O.bmp")
                    End If
                    rig = System.Math.Truncate(pcbImage.Tag / 3)
                    pcbImage.Enabled = False
                    col = pcbImage.Tag Mod 3
                    matrice(rig, col) = i + 1
                    mossa += 1
                    controlla()
                End If
            Next i
        Else
            If crea = MAXF Then
                clearTab()
            Else
                creaSimboli(1)
            End If
            crea += 1
            End If
    End Sub

    Private Sub calcola(ByVal X As Integer, ByVal Y As Integer)
        Dim i, j, fuori As Integer

        If crea <> MAXF + 1 Then
            If numP(crea) < MAXP + 1 Then
                px(crea, numP(crea)) = X
                py(crea, numP(crea)) = Y
                numP(crea) += 1
            End If
        Else
            For i = 0 To MAXF Step 1
                If errore(i) = 0 Then
                    fuori = False
                    For j = 0 To numP(i) - 1 Step 1
                        If System.Math.Sqrt((px(i, j) - X) ^ 2 + (py(i, j) - Y) ^ 2) < 35 Then
                            scia(i, j) = 1
                        Else
                            fuori += 1
                        End If
                    Next

                    If fuori = numP(i) Then
                        errore(i) = 1
                    End If
                End If
            Next
        End If
    End Sub
    Private Sub creaSimboli(ByVal i)
        If (i = 0) Then
            pcb1.Image = Image.FromFile("X.bmp")
            pcb1.Visible = True
            pcb5.Image = Image.FromFile("sfondo.bmp")
            pcb5.Visible = True
        Else
            pcb1.Visible = False
            pcb3.Image = Image.FromFile("O.bmp")
            pcb3.Visible = True
        End If
    End Sub
    Private Sub clearTab()
        pcb1.Image = Image.FromFile("sfondo.bmp")
        pcb2.Image = Image.FromFile("sfondo.bmp")
        pcb3.Image = Image.FromFile("sfondo.bmp")
        pcb4.Image = Image.FromFile("sfondo.bmp")
        pcb5.Image = Image.FromFile("sfondo.bmp")
        pcb6.Image = Image.FromFile("sfondo.bmp")
        pcb7.Image = Image.FromFile("sfondo.bmp")
        pcb8.Image = Image.FromFile("sfondo.bmp")
        pcb9.Image = Image.FromFile("sfondo.bmp")
        pcb1.Visible = True
        pcb2.Visible = True
        pcb3.Visible = True
        pcb4.Visible = True
        pcb5.Visible = True
        pcb6.Visible = True
        pcb7.Visible = True
        pcb8.Visible = True
        pcb9.Visible = True
        lblTesto1.Visible = False
    End Sub
    Private Sub controlla()
        Dim i, j, g As Integer
        Dim vinto As Boolean

        For g = 1 To 2 Step 1
            vinto = False
            For i = 0 To MAXC Step 1
                j = 0
                While matrice(i, j) = g And j <> MAXC
                    j += 1
                End While
                If matrice(i, j) = g Then
                    vinto = True
                Else
                    j = 0
                    While matrice(j, i) = g And j <> MAXC
                        j += 1
                    End While
                    If (matrice(j, i) = g) Then
                        vinto = True
                    End If
                End If
            Next
            If Not vinto Then
                i = 0
                While matrice(i, i) = g And i <> MAXC
                    i += 1
                End While
                If matrice(i, i) = g Then
                    vinto = True
                Else
                    i = 0
                    While matrice(i, MAXC - i) = g And i <> MAXC
                        i += 1
                    End While
                    If matrice(i, MAXC - i) = g Then
                        vinto = True
                    End If
                End If
            End If
            If vinto Then
                MessageBox.Show("Giocatore " & g.ToString & " ha VINTO !!!", "Vittoria", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Reset()
            ElseIf mossa = 9 Then
                MessageBox.Show("NESSUN giocatore ha VINTO", "Pareggio", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Reset()
            End If
        Next
    End Sub
    Private Sub reset()
        pcb1.Image = Nothing
        pcb2.Image = Nothing
        pcb3.Image = Nothing
        pcb4.Image = Nothing
        pcb5.Image = Nothing
        pcb6.Image = Nothing
        pcb7.Image = Nothing
        pcb8.Image = Nothing
        pcb9.Image = Nothing
        pcb1.Visible = False
        pcb2.Visible = False
        pcb3.Visible = False
        pcb4.Visible = False
        pcb5.Visible = False
        pcb6.Visible = False
        pcb7.Visible = False
        pcb8.Visible = False
        pcb9.Visible = False
        pcb1.Enabled = True
        pcb2.Enabled = True
        pcb3.Enabled = True
        pcb4.Enabled = True
        pcb5.Enabled = True
        pcb6.Enabled = True
        pcb7.Enabled = True
        pcb8.Enabled = True
        pcb9.Enabled = True
        lblTesto1.Visible = True
        Array.Clear(numP, 0, 2)
        Array.Clear(matrice, 0, 9)

        premuto = False
        crea = 0
        creaSimboli(0)
        mossa = 0
    End Sub
End Class

