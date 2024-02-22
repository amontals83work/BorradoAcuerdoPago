Imports System.Data
Imports System.Threading
Imports Microsoft.Office.Interop
Imports System.Text.RegularExpressions

Public Class FormBorradoAcuerdosPagos
#Region "BarraTareasMC"
    Implements TaskbarWindow
    Private mTaskbarButton As ToolStripItem
    Public Property TaskbarButton() As System.Windows.Forms.ToolStripItem Implements TaskbarWindow.TaskbarButton
        Get
            Return mTaskbarButton
        End Get
        Set(ByVal value As System.Windows.Forms.ToolStripItem)
            mTaskbarButton = value
        End Set
    End Property
#End Region

    Private mc As New MCCommand
    Private mcAux As New MCCommand
    Private openFileDialog As New OpenFileDialog

    Private Sub FormBorradoAcuerdosPagos_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Width = 395
    End Sub

    Private Sub btnFichero_Click(sender As Object, e As EventArgs) Handles btnFichero.Click
        If OpenFileDialog.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
            txtFichero.Text = OpenFileDialog.FileName
        Else
            MessageBox.Show("Debe seleccionar un fichero para cargar los expedientes")
            Exit Sub
        End If
    End Sub

    Private Sub btnBorrar_Click(sender As Object, e As EventArgs) Handles btnBorrar.Click
        BorrarAP()
    End Sub

    Private Sub BorrarAP()

        Dim MyConnection As System.Data.OleDb.OleDbConnection = New System.Data.OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + txtFichero.Text.Trim + ";Extended Properties=""Excel 12.0 Xml;HDR=YES;IMEX=1"";")
        Dim hoja As String = nhoja()
        Dim MyCommand As System.Data.OleDb.OleDbDataAdapter = New System.Data.OleDb.OleDbDataAdapter("select * from [" + hoja + "$]", MyConnection)

        Dim ds As System.Data.DataSet = New System.Data.DataSet()
        MyCommand.Fill(ds)
        MyConnection.Close()

        Dim dt As DataTable = ds.Tables(0)

        If Check_Excel(dt) Then
            Dim expediente As String = "", observacion As String = "", formatObser As String = "", fechaIPP As String = "", fechaPP As String = "", fechaChecked As String = "", count As Integer = 0, numExp As Integer = 0, numPP As Integer = 0, idExpediente As Object = DBNull.Value
            Dim aExpRep As New List(Of String)()
            Dim aExpNoPP As New List(Of String)()

            Dim messageThread As New Thread(AddressOf MuestraMensaje)
            messageThread.Start()

            For Each fila As DataRow In dt.Rows
                expediente = fila(0).ToString
                fechaIPP = fila(1).ToString
                observacion = ""
                fechaPP = ""

                'COMPROBAMOS SI HAY COINCIDENCIA ENTRE Expediente Y RefCliente
                If Not String.IsNullOrEmpty(expediente) Then
                    If Not Integer.TryParse(expediente, Nothing) Then
                        mc.CommandText = "SELECT COUNT(*) FROM Expedientes WHERE cast(Expediente as varchar)='" & expediente.ToString() & "' OR RefCliente='" & expediente.ToString() & "'"
                    Else
                        mc.CommandText = "SELECT COUNT(*) FROM Expedientes WHERE Expediente=" & expediente.ToString() & " OR RefCliente='" & expediente.ToString() & "'"
                    End If
                    numExp = mc.ExecuteScalar()
                    If numExp <> 1 Then
                        idExpediente = DBNull.Value
                        aExpRep.Add(expediente.ToString())
                    Else
                        'SI HAY SOLO UNA COINCIDENCIA BUSCAMOS SU idExpediente
                        If Not Integer.TryParse(expediente, Nothing) Then
                            mc.CommandText = "SELECT idExpediente FROM Expedientes WHERE cast(Expediente as varchar)='" & expediente.ToString() & "' OR RefCliente='" & expediente.ToString() & "'"
                        Else
                            mc.CommandText = "SELECT idExpediente FROM Expedientes WHERE Expediente='" & expediente.ToString() & "' OR RefCliente='" & expediente.ToString() & "'"
                        End If
                        idExpediente = mc.ExecuteScalar()

                        Dim result As Integer = If(idExpediente IsNot DBNull.Value, Convert.ToInt32(idExpediente), 0)
                        'SI HAY idExpediente COMPROBAMOS SI TIENE UNA ACUERDO DE PAGO
                        If result <> 0 Then
                            mc.CommandText = "SELECT COUNT(*) FROM Acciones WHERE idExpediente=" & idExpediente.ToString() & " AND idTipoNota=1"
                            numPP = mc.ExecuteScalar()

                            If numPP = 0 Then
                                aExpNoPP.Add(expediente.ToString())
                            Else
                                mc.CommandText = "SELECT Observaciones FROM Acciones WHERE idExpediente=" & idExpediente.ToString() & " AND idTipoNota=1"
                                fechaChecked = Check_Fecha(fechaIPP, expediente)
                                If Not String.IsNullOrEmpty(fechaChecked) Then
                                    observacion = "INCUMPLIMIENTO: " & fechaChecked.ToString() & " // " & mc.ExecuteScalar().Replace("'", "")
                                Else
                                    observacion = "INCUMPLIMIENTO // " & mc.ExecuteScalar().Replace("'", "")
                                End If
                                mc.CommandText = "SELECT Fecha FROM Acciones WHERE idExpediente=" & idExpediente.ToString() & " AND idTipoNota=1"
                                fechaPP = mc.ExecuteScalar()

                                mc.CommandText = "INSERT INTO Acciones (idExpediente, idTipoNota, Fecha, Observaciones, Usuario, Hermes, Borrado, Activa) VALUES ('" & idExpediente.ToString() & "',749,'" & fechaPP.ToString() & "','" & observacion.ToString() & "','Automarcador',0,0,1)"
                                mc.ExecuteNonQuery()

                                mc.CommandText = "UPDATE Acciones SET borrado=1 WHERE idExpediente=" & idExpediente.ToString() & " AND idTipoNota=1"
                                mc.ExecuteNonQuery()

                                mc.CommandText = "UPDATE Expedientes SET AcuerdoPago=0 WHERE idExpediente=" & idExpediente.ToString()
                                mc.ExecuteNonQuery()
                                count += 1
                            End If
                        End If
                    End If
                End If
            Next

            MessageBox.Show("Se han borrado " + count.ToString + " Acuerdos de Pago.")

            If aExpRep.Count > 0 Then
                Dim encabezadoRep As String = "Expedientes y RefClientes que coinciden: " & Environment.NewLine
                For Each repe As String In aExpRep
                    encabezadoRep &= repe & Environment.NewLine
                Next
                MessageBox.Show(encabezadoRep)
            End If
            If aExpNoPP.Count > 0 Then
                Dim encabezadoNoPP As String = "Expedientes sin Acuerdos de Pagos: " & Environment.NewLine
                For Each repe As String In aExpNoPP
                    encabezadoNoPP &= repe & Environment.NewLine
                Next
                MessageBox.Show(encabezadoNoPP)
            End If
        Else
            MessageBox.Show("Después de la modificación vuelva a cargar el archivo con otro nombre")
        End If
    End Sub

    Private Function Check_Fecha(fecha As String, expediente As String) As String
        If Not String.IsNullOrEmpty(fecha) Then
            Dim formatFecha As String = Convert.ToDateTime(fecha).ToString("dd/MM/yyyy")
            If formatFecha.Length = 10 Then
                Dim sRegex As String = "^(0[1-9]|[12][0-9]|3[01])/(0[1-9]|1[0-2])/[0-9]{4}$"
                Dim result As Match = Regex.Match(formatFecha, sRegex)
                If result.Success Then
                    Return formatFecha
                Else
                    MessageBox.Show("Comprobar el formato de celda 'FechaPP' es DD/MM/YYYY")
                    Return Nothing
                End If
            Else
                MessageBox.Show("Comprobar el formato de celda 'FechaPP' es DD/MM/YYYY")
                Return Nothing
            End If
        Else
            MessageBox.Show("No existe fecha en el EXPEDIENTE: " & expediente)
        End If
    End Function

    Private Function Check_Excel(dt As DataTable) As Boolean
        Dim expediente As String = If(dt.Rows(0).Table.Columns.Contains("Expediente"), dt.Rows(0)("Expediente")?.ToString(), Nothing)
        Dim fecha As String = If(dt.Rows(0).Table.Columns.Contains("FechaPP"), dt.Rows(0)("FechaPP")?.ToString().ToLower(), Nothing)

        If Not String.IsNullOrEmpty(expediente) Then
            If Not String.IsNullOrEmpty(fecha) Then
                Check_Fecha(fecha, expediente)
                Return True
            Else
                MessageBox.Show("Comprobar si la segunda columna esta nombrada como 'FechaPP'")
                Return False
            End If
        Else
            MessageBox.Show("Comprobar si la primera columna esta nombrada como 'Expediente'")
            Return False
        End If
    End Function

    Private Shared Sub MuestraMensaje()
        MessageBox.Show("Borrando los Acuerdos de Pago." + vbCrLf + "Este proceso puede tardar varios minutos y ralentizar el programa DespachoMC." + vbCrLf + "Al finalizar aparecerá un aviso.")
    End Sub

    Private Function nhoja() As String
        Dim hoja As String
        Dim xlsApp = New Excel.Application()
        xlsApp.Workbooks.Open(txtFichero.Text.Trim)

        hoja = xlsApp.Sheets(1).Name

        xlsApp.DisplayAlerts = False
        xlsApp.Workbooks.Close()
        xlsApp.DisplayAlerts = True

        xlsApp.Quit()
        xlsApp = Nothing

        Return hoja
    End Function

    Private Sub btnInstrucciones_Click(sender As Object, e As EventArgs) Handles btnInstrucciones.Click
        If Me.Width = 395 Then
            Me.Width = 600
        Else
            Me.Width = 395
        End If
    End Sub

End Class