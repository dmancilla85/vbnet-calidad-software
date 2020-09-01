Public Class frmMain

   Private iFase As Integer

   Private dFuncAdFuncional As Double
   Private dFuncPrecision As Double

   Private dUsabEntendible As Double
   Private dUsabAprendizaje As Double
   Private dUsabCapOper As Double
   Private dUsabAtractivo As Double

   Private dMantAnalizable As Double
   Private dMantModificable As Double

   Private dPortInstalable As Double
   Private dPortCoexistencia As Double
   Private dPortReemplazable As Double

   Private dFiabMadurez As Double
   Private dFiabRecuperable As Double

   Private dEficComportamiento As Double
   Private dEficRecursos As Double

   Private Enum Estado
      Inicio = 0
      Funcionalidad = 1
      Usabilidad = 2
      Mantenibilidad = 3
      Portabilidad = 4
      Fiabilidad = 5
      Eficiencia = 6
      Resultado = 7
   End Enum

   Private Sub btnSalir_Click(sender As System.Object, e As System.EventArgs) Handles btnSalir.Click
      Close()
   End Sub

   Private Sub frmMain_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load

      CenterForm()
      iFase = Estado.Inicio
      determinaPan(iFase)
      btnAnterior.Enabled = False
      btnSiguiente.Text = "Iniciar"
   End Sub

   Private Sub btnSiguiente_Click(sender As System.Object, e As System.EventArgs) Handles btnSiguiente.Click

      If iFase < Estado.Resultado Then
         iFase = iFase + 1
      End If

      If iFase >= Estado.Funcionalidad And iFase < Estado.Resultado Then
         btnSiguiente.Text = "Siguiente"

         If EsValido(iFase) And iFase > Estado.Inicio And
             iFase < Estado.Resultado Then
            determinaPan(iFase)
         Else
            iFase = iFase - 1
         End If
      Else
         determinaPan(iFase)

      End If

      btnAnterior.Enabled = iFase > Estado.Inicio
      btnSiguiente.Enabled = iFase < Estado.Resultado

   End Sub

   Private Sub determinaPan(iNum As Integer)

      Select Case iNum
         Case Estado.Inicio
            Limpiar()
            panResultado.Visible = False
            panForm.Visible = False
            panMain.Visible = True
            CenterFrame(panMain)

         Case Estado.Resultado
            panMain.Visible = False
            panForm.Visible = False
            panResultado.Visible = True
            CenterFrame(panResultado)
            CalcularResultado()

         Case Estado.Funcionalidad
            panResultado.Visible = False
            panMain.Visible = False
            panForm.Visible = True
            CenterFrame(panForm)
            gbFunc.Enabled = True
            gbUsab.Enabled = False
            gbMant.Enabled = False
            gbPorta.Enabled = False
            gbFiab.Enabled = False
            gbEfic.Enabled = False
            mskFunCor.Focus()
         Case Estado.Usabilidad
            gbFunc.Enabled = True
            gbUsab.Enabled = True
            gbMant.Enabled = False
            gbPorta.Enabled = False
            gbFiab.Enabled = False
            gbEfic.Enabled = False
            mskUsa_a.Focus()

         Case Estado.Mantenibilidad
            gbFunc.Enabled = True
            gbUsab.Enabled = True
            gbMant.Enabled = True
            gbPorta.Enabled = False
            gbFiab.Enabled = False
            gbEfic.Enabled = False
            mskMant_a.Focus()

         Case Estado.Portabilidad
            gbFunc.Enabled = True
            gbUsab.Enabled = True
            gbMant.Enabled = True
            gbPorta.Enabled = True
            gbFiab.Enabled = False
            gbEfic.Enabled = False
            mskPorta_a.Focus()

         Case Estado.Fiabilidad
            gbFunc.Enabled = True
            gbUsab.Enabled = True
            gbMant.Enabled = True
            gbPorta.Enabled = True
            gbFiab.Enabled = True
            gbEfic.Enabled = False
            mskFiab_a.Focus()

         Case Estado.Eficiencia
            gbFunc.Enabled = True
            gbUsab.Enabled = True
            gbMant.Enabled = True
            gbPorta.Enabled = True
            gbFiab.Enabled = True
            gbEfic.Enabled = True
            mskEfic_a.Focus()

      End Select
   End Sub

   Public Sub CenterForm()
      Top = (Screen.PrimaryScreen.Bounds.Height - Height) \ 2
      Left = (Screen.PrimaryScreen.Bounds.Width - Width) \ 2
   End Sub

   Public Sub CenterFrame(ByRef F As Panel)
      ' Center the specified form within the screen
      On Error Resume Next
      F.Top = (Height - F.Height) \ 2 - 30
      F.Left = (Width - F.Width) \ 2
      On Error GoTo 0
   End Sub

   Private Sub btnAnterior_Click(sender As System.Object, e As System.EventArgs) Handles btnAnterior.Click

      If iFase > Estado.Inicio And iFase < Estado.Resultado Then
         iFase = iFase - 1
      Else
         iFase = Estado.Funcionalidad
      End If

      If iFase = Estado.Inicio Then
         panMain.Visible = True
         CenterFrame(panMain)
         panForm.Visible = False
      End If

      If iFase >= Estado.Funcionalidad And iFase < Estado.Resultado Then
         panMain.Visible = False
         CenterFrame(panForm)
         determinaPan(iFase)
         panForm.Visible = True
      End If

      btnAnterior.Enabled = iFase > Estado.Inicio
      btnSiguiente.Enabled = iFase < Estado.Resultado
   End Sub

   Private Function EsValido(iFase As Integer) As Boolean

      Dim bVale As Boolean
      bVale = True

      If iFase = Estado.Inicio Or iFase > Estado.Resultado Then
         Return True
      End If

      Try
         Select Case iFase
            Case Estado.Funcionalidad + 1
               If mskFuncInc.Text = "" Or _
                   mskFunCor.Text = "" Or _
                   mskPrecOk.Text = "" Or
                   mskPrecTotal.Text = "" Then
                  MsgBox("Falta ingresar valores", MsgBoxStyle.Exclamation)
                  bVale = False
               End If

               If Integer.Parse(mskPrecOk.Text) > Integer.Parse(mskPrecTotal.Text) Then
                  MsgBox("Casos totales debe ser mayor o igual que casos exitosos.", _
                         MsgBoxStyle.Exclamation)
                  bVale = False
               End If

            Case Estado.Usabilidad + 1
               If mskUsa_a.Text = "" Or
                   mskUsa_b.Text = "" Or
                   mskUsa_c.Text = "" Or
                   mskUsa_d.Text = "" Or
                   mskUsa_e.Text = "" Or
                   mskUsa_f.Text = "" Then
                  MsgBox("Falta ingresar valores", MsgBoxStyle.Exclamation)
                  bVale = False
               End If

               If Integer.Parse(mskUsa_c.Text) > Integer.Parse(mskUsa_d.Text) Then
                  MsgBox("Total de tareas debe ser mayor o igual" _
                         + " que tareas con ayuda.", MsgBoxStyle.Exclamation)
                  bVale = False
               End If

            Case Estado.Mantenibilidad + 1
               If mskMant_a.Text = "" Or
                   mskMant_b.Text = "" Or
                   mskMant_c.Text = "" Or
                   mskMant_d.Text = "" Or
                   mskMant_e.Text = "" Or
                   mskMant_f.Text = "" Then
                  MsgBox("Falta ingresar valores", MsgBoxStyle.Exclamation)
                  bVale = False
               End If

               If Integer.Parse(mskMant_a.Text) > Integer.Parse(mskMant_b.Text) Then
                  MsgBox("Total de fallas debe ser mayor o igual" _
                         + " que fallas informadas.", MsgBoxStyle.Exclamation)
                  bVale = False
               End If

               If Integer.Parse(mskMant_c.Text) > Integer.Parse(mskMant_d.Text) Then
                  MsgBox("Total de fallas debe ser mayor o igual" _
                         + " que fallas muy localizables.", MsgBoxStyle.Exclamation)
                  bVale = False
               End If

            Case Estado.Portabilidad + 1
               If mskPorta_a.Text = "" Or
                   mskPorta_b.Text = "" Or
                   mskPorta_c.Text = "" Or
                   mskPorta_d.Text = "" Or
                   mskPorta_e.Text = "" Or
                   mskPorta_f.Text = "" Then
                  MsgBox("Falta ingresar valores", MsgBoxStyle.Exclamation)
                  bVale = False
               End If

               If Integer.Parse(mskPorta_a.Text) > Integer.Parse(mskPorta_b.Text) Then
                  MsgBox("Total de instalaciones realizadas debe ser mayor o igual" _
                         + " que instalaciones exitosas.", MsgBoxStyle.Exclamation)
                  bVale = False
               End If

               If Integer.Parse(mskPorta_e.Text) > Integer.Parse(mskPorta_f.Text) Then
                  MsgBox("Datos totales debe ser mayor o igual" _
                         + " que datos reemplazados.", MsgBoxStyle.Exclamation)
                  bVale = False
               End If

            Case Estado.Fiabilidad + 1
               If mskFiab_a.Text = "" Or
                   mskFiab_b.Text = "" Or
                   mskFiab_c.Text = "" Or
                   mskFiab_d.Text = "" Then
                  MsgBox("Falta ingresar valores", MsgBoxStyle.Exclamation)
                  bVale = False
               End If

               If Integer.Parse(mskFiab_a.Text) > Integer.Parse(mskFiab_b.Text) Then
                  MsgBox("Total de fallas debe ser mayor o igual" _
                         + " que cantidad de pruebas realizadas.", MsgBoxStyle.Exclamation)
                  bVale = False
               End If

               If Integer.Parse(mskFiab_c.Text) > Integer.Parse(mskFiab_d.Text) Then
                  MsgBox("Restauraciones totales debe ser mayor o igual" _
                         + " que restauraciones exitosas.", MsgBoxStyle.Exclamation)
                  bVale = False
               End If

            Case Estado.Eficiencia + 1
               If mskEfic_a.Text = "" Or _
                   mskEfic_b.Text = "" Or _
                   mskEfic_c.Text = "" Or
                   mskEfic_d.Text = "" Then
                  MsgBox("Falta ingresar valores", MsgBoxStyle.Exclamation)
                  bVale = False
               End If
         End Select

      Catch ex As System.FormatException
         bVale = False
      End Try


      'If bVale Then
      '   btnSiguiente.Focus()
      'End If

      Return bVale

   End Function

   Private Sub CalcularResultado()

      Dim iRes As Double
      Dim funca As Double
      Dim usab As Double
      Dim mant As Double
      Dim porta As Double
      Dim fiab As Double
      Dim efic As Double

      iRes = 0

      Try
         ' Funcionalidad
         lblFunc_a.Text = Format((((Double.Parse(mskFunCor.Text) / (Double.Parse(mskFuncInc.Text) + Double.Parse(mskFunCor.Text)))) * 100), "#.##")
         lblFunc_b.Text = Format(((Double.Parse(mskPrecOk.Text) / Double.Parse(mskPrecTotal.Text)) * 100), "#.##")
         'Usabilidad
         lblUsab_a.Text = Format(((Double.Parse(mskUsa_a.Text) / (Double.Parse(mskUsa_a.Text) + Double.Parse(mskUsa_b.Text))) * 100), "#.##")
         lblUsab_b.Text = Format(((Double.Parse(mskUsa_c.Text) / Double.Parse(mskUsa_d.Text)) * 100), "#.##")
         lblUsab_c.Text = Format(((Double.Parse(mskUsa_e.Text) / (Double.Parse(mskUsa_e.Text) + Double.Parse(mskUsa_f.Text))) * 100), "#.##")
         'Mantenibilidad
         lblMant_a.Text = Format(((Double.Parse(mskMant_a.Text) / Double.Parse(mskMant_b.Text)) * 100), "#.##")
         lblMant_b.Text = Format(((Double.Parse(mskMant_c.Text) / Double.Parse(mskMant_d.Text)) * 100), "#.##")
         lblMant_c.Text = Format(((Double.Parse(mskMant_e.Text) / (Double.Parse(mskMant_e.Text) + Double.Parse(mskMant_f.Text))) * 100), "#.##")

         'portabilidad
         lblPorta_a.Text = Format(((Double.Parse(mskPorta_a.Text) / Double.Parse(mskPorta_b.Text)) * 100), "#.##")
         lblPorta_b.Text = Format(((Double.Parse(mskPorta_c.Text) / Double.Parse(mskPorta_d.Text)) * 100), "#.##")
         lblPorta_c.Text = Format(((1 - (Double.Parse(mskPorta_e.Text) / Double.Parse(mskPorta_f.Text))) * 100), "#.##")

         'fiabilidad
         lblFiab_a.Text = Format(((Double.Parse(mskFiab_a.Text) / Double.Parse(mskFiab_b.Text)) * 100), "#.##")
         lblFiab_b.Text = Format(((Double.Parse(mskFiab_c.Text) / Double.Parse(mskFiab_d.Text)) * 100), "#.##")

         'eficiencia
         lblEfic_a.Text = Format(((Double.Parse(mskEfic_a.Text) / Double.Parse(mskEfic_b.Text)) * 100), "#.##")
         lblEfic_b.Text = Format(((Double.Parse(mskEfic_c.Text) / Double.Parse(mskEfic_d.Text)) * 100), "#.##")

         funca = Double.Parse(lblFunc_a.Text) + Double.Parse(lblFunc_b.Text)
         funca = funca / 2

         usab = Double.Parse(lblUsab_a.Text) + Double.Parse(lblUsab_c.Text) + Double.Parse(lblUsab_c.Text)
         usab = usab / 3

         fiab = Double.Parse(lblFiab_a.Text) + Double.Parse(lblFiab_b.Text)
         fiab = fiab / 2

         porta = Double.Parse(lblPorta_a.Text) + Double.Parse(lblPorta_b.Text) + Double.Parse(lblPorta_c.Text)
         porta = porta / 3

         mant = Double.Parse(lblMant_a.Text) + Double.Parse(lblMant_b.Text) + Double.Parse(lblMant_c.Text)
         mant = mant / 3

         efic = Double.Parse(lblEfic_a.Text) + Double.Parse(lblEfic_b.Text)
         efic = efic / 2

         iRes = (funca + usab + fiab + porta + mant + efic) / 6
         iRes = FormatNumber(iRes, 2)
      Catch ex As Exception
         MsgBox("Ha ocurrido un error.", MsgBoxStyle.Critical)
      End Try
      

      lblPuntoFinal.Text = iRes.ToString

      If iRes >= 0 And iRes <= 39 Then
         lblResultado.Text = "NO SATISFACTORIO"
         lblResultado.ForeColor = Color.Red
      Else
         If iRes > 39 And iRes <= 69 Then
            lblResultado.Text = "SATISFACTORIO"
            lblResultado.ForeColor = Color.GreenYellow
         Else
            If iRes <= 100 Then
               lblResultado.Text = "EXCELENTE"
               lblResultado.ForeColor = Color.Green
            Else
               lblResultado.Text = "ANDAS VOLANDO!!!"
               lblResultado.ForeColor = Color.SeaGreen
            End If

         End If
      End If

   End Sub

   Private Sub Limpiar()
      mskEfic_a.Clear()
      mskEfic_b.Clear()
      mskEfic_c.Clear()
      mskEfic_d.Clear()
      mskFiab_a.Clear()
      mskFiab_b.Clear()
      mskFiab_c.Clear()
      mskFiab_d.Clear()
      mskFuncInc.Clear()
      mskFunCor.Clear()
      mskMant_a.Clear()
      mskMant_b.Clear()
      mskMant_c.Clear()
      mskMant_d.Clear()
      mskMant_e.Clear()
      mskMant_f.Clear()
      mskPorta_a.Clear()
      mskPorta_b.Clear()
      mskPorta_c.Clear()
      mskPorta_d.Clear()
      mskPorta_e.Clear()
      mskPorta_f.Clear()
      mskPrecOk.Clear()
      mskPrecTotal.Clear()
      mskUsa_a.Clear()
      mskUsa_b.Clear()
      mskUsa_c.Clear()
      mskUsa_d.Clear()
      mskUsa_e.Clear()
      mskUsa_f.Clear()
      lblEfic_a.Text = ""
      lblEfic_b.Text = ""
      lblFiab_b.Text = ""
      lblFiab_a.Text = ""
      lblFunc_a.Text = ""
      lblFunc_b.Text = ""
      lblMant_a.Text = ""
      lblMant_b.Text = ""
      lblMant_c.Text = ""
      lblPorta_a.Text = ""
      lblPorta_b.Text = ""
      lblPorta_c.Text = ""
      lblUsab_a.Text = ""
      lblUsab_b.Text = ""
      lblUsab_c.Text = ""
      lblResultado.Text = ""
      lblResultado.ForeColor = Color.Black
      lblPuntoFinal.Text = ""
   End Sub

   Private Sub mskPrecTotal_LostFocus(sender As Object, e As System.EventArgs) Handles mskPrecTotal.LostFocus
      btnSiguiente.Focus()
   End Sub

   Private Sub mskUsa_f_LostFocus(sender As Object, e As System.EventArgs) Handles mskUsa_f.LostFocus
      btnSiguiente.Focus()
   End Sub

   Private Sub mskPorta_f_LostFocus(sender As Object, e As System.EventArgs) Handles mskPorta_f.LostFocus
      btnSiguiente.Focus()
   End Sub

   Private Sub mskMant_f_LostFocus(sender As Object, e As System.EventArgs) Handles mskMant_f.LostFocus
      btnSiguiente.Focus()
   End Sub

   Private Sub mskFiab_d_LostFocus(sender As Object, e As System.EventArgs) Handles mskFiab_d.LostFocus
      btnSiguiente.Focus()
   End Sub

   Private Sub mskefic_d_LostFocus(sender As Object, e As System.EventArgs) Handles mskEfic_d.LostFocus
      btnSiguiente.Focus()
   End Sub

   Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
      AboutBox1.Show(Me)
   End Sub
End Class
