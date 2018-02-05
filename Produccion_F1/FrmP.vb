Imports System.Windows.Forms
Imports System.IO
Imports System.Data.SqlClient
Imports System.Data
Imports System.Drawing.Text
Imports CrystalDecisions.Shared
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Windows.Forms
Imports CrystalDecisions.ReportSource
Imports System.Web.UI.WebControls
Public Class FrmP
#Region "Variables"
    Public con As New Conexion
    Dim objectCode As String
    Dim itemcode As String
    Dim oCompany As SAPbobsCOM.Company
    Dim connectionString As String = Conexion.ObtenerConexion.ConnectionString
    Public Shared PO As SAPbobsCOM.Documents
    Public Shared GoodsReceiptPO As SAPbobsCOM.Documents
    Public Shared SQL_Conexion As SqlConnection = New SqlConnection()

#Region "Listas"
    Dim batch As New List(Of String)
    Dim descripcion As New List(Of String)
    Dim anchotira As New List(Of Double)
    Dim pesoreal As New List(Of Double)
    Dim bobina As New List(Of String)
    Dim heat As New List(Of String)
    Dim coil As New List(Of String)
    Dim ordencorte As New List(Of String)
#End Region
#Region "Fuentes"
    Private _Font As Font
    Private PATH_FONTS As String = Application.StartupPath + "\Fonts"
#End Region
    Private Const CP_NOCLOSE_BUTTON As Integer = &H200
#End Region
    Protected Overloads Overrides ReadOnly Property CreateParams() As CreateParams
        Get
            Dim myCp As CreateParams = MyBase.CreateParams
            myCp.ClassStyle = myCp.ClassStyle Or CP_NOCLOSE_BUTTON
            Return myCp
        End Get
    End Property
    Public Sub New(ByVal user As String)
        MyBase.New()
        InitializeComponent()
        '  Note which form has called this one
        ToolStripStatusLabel1.Text = user
    End Sub
    Private Function FormatBarCode(code As String)
        Dim barcode As String = String.Empty
        barcode = String.Format("*{0}*", code)
        Return barcode
    End Function
    Private Sub FrmFase1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        DateTimePicker1.Format = DateTimePickerFormat.Custom
        DateTimePicker1.CustomFormat = "yyyy/MM/dd"
        DGV2.Visible = False
    End Sub
    Public Function cargaORDER()
        Dim SQL_da As SqlDataAdapter = New SqlDataAdapter("SELECT T0.DocNum FROM OPOR T0 WHERE T0.DocType ='I' and  T0.CANCELED = 'N' and  T0.DocStatus ='O'", con.ObtenerConexion())
        Dim DT_dat As System.Data.DataTable = New System.Data.DataTable()
        SQL_da.Fill(DT_dat)
        DGV.DataSource = DT_dat
        con.ObtenerConexion.Close()
    End Function


    Private Sub imprime(barcode As String, desc As String, anch As String, pes As String, bob As String, het As String, coi As String)
        Dim Report1 As New CrystalDecisions.CrystalReports.Engine.ReportDocument()
        Report1.PrintOptions.PaperOrientation = PaperOrientation.Portrait
        Report1.Load(Application.StartupPath + "\Report\Informe.rpt", CrystalDecisions.Shared.OpenReportMethod.OpenReportByDefault.OpenReportByDefault)
        Report1.SetParameterValue("CodBatch", barcode)
        'Report1.SetParameterValue("CodBatch", txtBarcode.Text)
        Report1.SetParameterValue("descripcion", desc)
        Report1.SetParameterValue("anchotira", anch)
        Report1.SetParameterValue("pesoreal", pes)
        Report1.SetParameterValue("bobina", bob)
        Report1.SetParameterValue("heat", het)
        Report1.SetParameterValue("coil", coi)
        Report1.SetParameterValue("ordencorte", objectCode)
        Report1.SetParameterValue("fechacorte", Now.ToShortDateString)
        'CrystalReportViewer1.ReportSource = Report1
        Report1.PrintToPrinter(1, False, 0, 0)
    End Sub
    Private Sub generaEntrada()
        Dim iResult As Integer = -1
        Dim iResult2 As Integer = -1
        Dim sResult As String = String.Empty
        Dim sOutput As String = String.Empty
        Dim sql As String
        Dim oRecordSet As SAPbobsCOM.Recordset
        Dim docentry As String

        Try
            Dim result As Integer = MessageBox.Show("Desea Imprimir los Lotes?", "Atencion", MessageBoxButtons.YesNoCancel)
            If result = DialogResult.Cancel Then
                MessageBox.Show("Cancelado")
            ElseIf result = DialogResult.No Then
                MessageBox.Show("No se realizara la orden")
            ElseIf result = DialogResult.Yes Then
                For Each row As DataGridViewRow In DGV2.Rows
                    Dim chk As DataGridViewCheckBoxCell = row.Cells("CHK")
                    If chk.Value IsNot Nothing AndAlso chk.Value = True Then
                        'barcode , desc , anch , pes , bob , het, coi
                        imprime(DGV2.Rows(chk.RowIndex).Cells.Item(4).Value.ToString, DGV2.Rows(chk.RowIndex).Cells.Item(3).Value.ToString, DGV2.Rows(chk.RowIndex).Cells.Item(7).Value.ToString, DGV2.Rows(chk.RowIndex).Cells.Item(8).Value.ToString, DGV2.Rows(chk.RowIndex).Cells.Item(1).Value.ToString, DGV2.Rows(chk.RowIndex).Cells.Item(5).Value.ToString, DGV2.Rows(chk.RowIndex).Cells.Item(6).Value.ToString)
                    End If
                Next

            End If

            '-----------------------------------------------------------------------------------
            MessageBox.Show("Operacion Realizada Exitosamente!")
            DGV2.Visible = False
        Catch ex As Exception
            MsgBox("Error: " + ex.Message.ToString)
        End Try
    End Sub
    Private Sub GR_from_PO()
        Try
            generaEntrada()
        Catch ex As Exception
            MsgBox("Error: " + ex.Message.ToString)
        End Try
    End Sub

    Private Sub DGV_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DGV.CellContentClick
        DGV2.Visible = Enabled
        objectCode = DGV(0, DGV.CurrentCell.RowIndex).Value.ToString()

        Dim SQL_da2 As SqlDataAdapter = New SqlDataAdapter("
        select T1.CardCode,T1.CardName,t1.NumAtCard,T1.DocDate
        from OPDN T1", con.ObtenerConexion())
        Dim DT_dat2 As System.Data.DataTable = New System.Data.DataTable()
        SQL_da2.Fill(DT_dat2)
        Label3.Text = DT_dat2.Rows(0).Item("CardCode").ToString
        Label5.Text = DT_dat2.Rows(0).Item("CardName").ToString
        Label7.Text = DT_dat2.Rows(0).Item("NumAtCard").ToString
        Label9.Text = DT_dat2.Rows(0).Item("DocDate").ToString
        con.ObtenerConexion.Close()

        Panel1.Visible = True

        Dim SQL_da As SqlDataAdapter = New SqlDataAdapter("SELECT t4.ItemCode,t4.ItemName,SUM(CASE T4.Direction when 0 then 1 else -1 end * T4.Quantity) as Quantitys, T4.BatchNum, T3.U_Heat,t3.U_Coi,t3.U_Ancho,t3.U_Correlativo
FROM OITL T0
INNER JOIN OPDN T2 on t2.DocEntry = t0.DocEntry		
INNER JOIN ITL1 T1 ON T0.LogEntry = T1.LogEntry
INNER JOIN OBTN T3 ON T1.MdAbsEntry = T3.AbsEntry
inner join IBT1 T4 on T4.BatchNum = T3.DistNumber
WHERE T0.DocEntry =  '" + DGV(0, DGV.CurrentCell.RowIndex).Value.ToString() + "' AND T0.DocNum =  '" + DGV(0, DGV.CurrentCell.RowIndex).Value.ToString() + "' and T0.BaseEntry = 0
GROUP BY t4.ItemCode,t4.itemname,T4.BatchNum , T3.U_Heat,t3.U_Coi,t3.U_Ancho,t3.U_Correlativo", con.ObtenerConexion())
        Dim DT_dat As System.Data.DataTable = New System.Data.DataTable()
        SQL_da.Fill(DT_dat)
        DGV2.DataSource = DT_dat
        For Each row As DataGridViewRow In DGV2.Rows
            row.Cells("CHK").Value = True
        Next
        con.ObtenerConexion.Close()
    End Sub


    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        GR_from_PO()
        DGV.DataSource = Nothing
        DGV2.DataSource = Nothing
    End Sub

    Private Sub btnFinalizar_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim result As Integer = MessageBox.Show("Desea limpiar el objeto?", "Atencion", MessageBoxButtons.YesNoCancel)
        If result = DialogResult.Cancel Then
            MessageBox.Show("Cancelado")
        ElseIf result = DialogResult.No Then
            MessageBox.Show("Puede continuar!")
        ElseIf result = DialogResult.Yes Then
            PO = Nothing
            GoodsReceiptPO = Nothing
            DGV.DataSource = Nothing
            DGV2.DataSource = Nothing
            MessageBox.Show("Inicie un objeto nuevo")
        End If
        DGV2.Visible = False
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim result As Integer = MessageBox.Show("Desea salir del modulo?", "Atencion", MessageBoxButtons.YesNo)
        If result = DialogResult.No Then
            MessageBox.Show("Puede continuar")
        ElseIf result = DialogResult.Yes Then
            MessageBox.Show("Finalizando modulo")
            Try
                con.oCompany.Disconnect()
            Catch
            End Try
            Application.Exit()
            Me.Close()
        End If
    End Sub

    'Private Sub txtOrder_TextChanged(sender As Object, e As EventArgs) Handles txtOrder.TextChanged
    '    Dim i As DataGridViewCheckBoxColumn = New DataGridViewCheckBoxColumn()
    '    Dim existe As Boolean = DGV2.Columns.Cast(Of DataGridViewColumn).Any(Function(x) x.Name = "CHK")
    '    If existe = False Then
    '        DGV2.Columns.Add(i)
    '        i.HeaderText = "CHK"
    '        i.Name = "CHK"
    '        i.Width = 32
    '        i.DisplayIndex = 0
    '    End If

    '    Dim SQL_da As SqlDataAdapter = New SqlDataAdapter("SELECT T0.[ItemCode], T0.[Quantity], isnull(T0.LineNum,0) as 'No.Linea' FROM POR1 T0 WHERE T0.[LineStatus] = 'O' and T0.[DocEntry] like '" + txtOrder.Text + "%'", con.ObtenerConexion())
    '    Dim DT_dat As System.Data.DataTable = New System.Data.DataTable()
    '    SQL_da.Fill(DT_dat)
    '    DGV2.DataSource = DT_dat
    '    con.ObtenerConexion.Close()
    'End Sub

    Private Sub DateTimePicker1_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker1.ValueChanged
        Dim SQL_da As SqlDataAdapter = New SqlDataAdapter("select DocNum from opdn where CANCELED = 'N' and docdate ='" + DateTimePicker1.Value.ToString("yyyy/MM/dd") + "' ORDER BY DocNum", con.ObtenerConexion())
        Dim DT_dat As System.Data.DataTable = New System.Data.DataTable()
        SQL_da.Fill(DT_dat)
        DGV.DataSource = DT_dat
        txtOrder.Text = DT_dat.Rows(0).Item("DocNum").ToString
        con.ObtenerConexion.Close()
    End Sub
End Class
