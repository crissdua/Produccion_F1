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
    Public Sub New()
        MyBase.New()
        InitializeComponent()
        '  Note which form has called this one
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


    Private Sub imprime(barcode As String, desc As String, anch As String, pes As String, itmcod As String, het As String, coi As String, whs As String)
        Dim Report1 As New CrystalDecisions.CrystalReports.Engine.ReportDocument()
        Report1.PrintOptions.PaperOrientation = PaperOrientation.Portrait
        Report1.Load(Application.StartupPath + "\Report\InformeF1.rpt", CrystalDecisions.Shared.OpenReportMethod.OpenReportByDefault.OpenReportByDefault)
        ''-----------------------------------------ENCABEZADO NO CAMBIA POR IMPRESION------------------------------------------
        Report1.SetParameterValue("cardcode", Label3.Text)
        Report1.SetParameterValue("cardname", Label5.Text)
        Report1.SetParameterValue("docnum", txtOrder.Text)
        Report1.SetParameterValue("docdate", Label9.Text)
        Report1.SetParameterValue("numatcard", Label7.Text)
        Report1.SetParameterValue("ingresodate", Label11.Text)
        ''------------------------------------------DETALLE TRAE DATOS POR PARAMETROS------------------------------------------
        Report1.SetParameterValue("CodBatch", barcode) 'col4
        Report1.SetParameterValue("descripcion", desc) 'col2
        Report1.SetParameterValue("pesoreal", pes) 'col5
        Report1.SetParameterValue("anchotira", anch) 'col6
        Report1.SetParameterValue("bobina", itmcod) 'col1
        Report1.SetParameterValue("heat", het)
        Report1.SetParameterValue("coil", coi)
        Report1.SetParameterValue("almacen", whs)
        'Report1.SetParameterValue("ordencorte", objectCode)
        'Report1.SetParameterValue("fechacorte", Now.ToShortDateString)
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
            Dim result As Integer = MessageBox.Show("Desea Imprimir Ingreso MP?", "Atencion", MessageBoxButtons.YesNoCancel)
            If result = DialogResult.Cancel Then
                MessageBox.Show("Cancelado")
            ElseIf result = DialogResult.No Then
                MessageBox.Show("No se realizara la orden")
            ElseIf result = DialogResult.Yes Then
                For Each row As DataGridViewRow In DGV2.Rows
                    Dim chk As DataGridViewCheckBoxCell = row.Cells("CHK")
                    If chk.Value IsNot Nothing AndAlso chk.Value = True Then
                        'barcode 4 , desc 2 , anch 6, pes 5, itmcod 1, het 7, coi 8,whs 9
                        imprime(DGV2.Rows(chk.RowIndex).Cells.Item(4).Value.ToString, DGV2.Rows(chk.RowIndex).Cells.Item(2).Value.ToString, DGV2.Rows(chk.RowIndex).Cells.Item(6).Value.ToString, DGV2.Rows(chk.RowIndex).Cells.Item(3).Value.ToString, DGV2.Rows(chk.RowIndex).Cells.Item(1).Value.ToString, DGV2.Rows(chk.RowIndex).Cells.Item(7).Value.ToString, DGV2.Rows(chk.RowIndex).Cells.Item(8).Value.ToString, DGV2.Rows(chk.RowIndex).Cells.Item(9).Value.ToString)
                    End If
                Next

            End If

            '-----------------------------------------------------------------------------------
            MessageBox.Show("Operacion Realizada Exitosamente!")
            DGV2.Visible = False
            Panel1.Visible = False
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
        select T1.CardCode,T1.CardName,T1.DocNum,t1.NumAtCard,convert(varchar(10),convert(date,T1.DocDate,106),103) as DocDate
        from OPDN T1 where T1.DocEntry =  '" + DGV(0, DGV.CurrentCell.RowIndex).Value.ToString() + "'", con.ObtenerConexion())
        Dim DT_dat2 As System.Data.DataTable = New System.Data.DataTable()
        SQL_da2.Fill(DT_dat2)
        txtOrder.Text = DGV(0, DGV.CurrentCell.RowIndex).Value.ToString()
        Label3.Text = DT_dat2.Rows(0).Item("CardCode").ToString
        Label5.Text = DT_dat2.Rows(0).Item("CardName").ToString
        Label7.Text = DT_dat2.Rows(0).Item("NumAtCard").ToString
        Label9.Text = DT_dat2.Rows(0).Item("DocDate").ToString
        Label11.Text = DateTime.Now.ToShortDateString
        con.ObtenerConexion.Close()

        Panel1.Visible = True

        Dim SQL_da As SqlDataAdapter = New SqlDataAdapter("SELECT t4.ItemCode as 'Codigo de Articulo',t4.ItemName as 'Descripcion del Articulo',SUM(CASE T4.Direction when 0 then 1 else -1 end * T4.Quantity) as TM, T4.BatchNum as Lote,'peso' as Peso,CONVERT(int,t3.U_Ancho) as Ancho,T3.U_Heat as Heat,t3.U_Coi as Coil,t4.whscode as Almacen
FROM OITL T0
INNER JOIN OPDN T2 on t2.DocEntry = t0.DocEntry		
INNER JOIN ITL1 T1 ON T0.LogEntry = T1.LogEntry
INNER JOIN OBTN T3 ON T1.MdAbsEntry = T3.AbsEntry
inner join IBT1 T4 on T4.BatchNum = T3.DistNumber
WHERE T0.DocEntry =  '" + DGV(0, DGV.CurrentCell.RowIndex).Value.ToString() + "' AND T0.DocNum =  '" + DGV(0, DGV.CurrentCell.RowIndex).Value.ToString() + "' and T0.BaseEntry = 0 and t4.WhsCode = 'BMP2'
GROUP BY t4.ItemCode,t4.itemname,T4.BatchNum , T3.U_Heat,t3.U_Coi,t3.U_Ancho,t3.U_Correlativo,T4.whscode", con.ObtenerConexion())
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
        End If
        DGV2.Visible = False
        Panel1.Visible = False
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim result As Integer = MessageBox.Show("Desea salir del modulo?", "Atencion", MessageBoxButtons.YesNo)
        If result = DialogResult.No Then
            MessageBox.Show("Puede continuar")
        ElseIf result = DialogResult.Yes Then
            Try
                con.oCompany.Disconnect()
            Catch
            End Try
            Me.Hide()
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

    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Dim SQL_da As SqlDataAdapter = New SqlDataAdapter("select isnull(DocNum,'0') as DocNum from opdn where CANCELED = 'N' and docdate ='" + DateTimePicker1.Value.ToString("yyyy/MM/dd") + "' ORDER BY DocNum", con.ObtenerConexion())
        Dim DT_dat As System.Data.DataTable = New System.Data.DataTable()
        SQL_da.Fill(DT_dat)
        DGV.DataSource = DT_dat
        If DT_dat.Rows.Count = 0 Then
            MessageBox.Show("No Hay Documentos en esta Fecha")
        Else
            txtOrder.Text = DT_dat.Rows(0).Item("DocNum").ToString
        End If

        con.ObtenerConexion.Close()
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click

        Try
            DGV2.Visible = Enabled

            Dim SQL_da2 As SqlDataAdapter = New SqlDataAdapter("
            select T1.CardCode,T1.CardName,T1.DocNum,t1.NumAtCard,convert(varchar(10),convert(date,T1.DocDate,106),103) as DocDate
            from OPDN T1 where T1.DocEntry =  '" + TextBox1.Text + "'", con.ObtenerConexion())
            Dim DT_dat2 As System.Data.DataTable = New System.Data.DataTable()
            SQL_da2.Fill(DT_dat2)
            txtOrder.Text = TextBox1.Text
            Label3.Text = DT_dat2.Rows(0).Item("CardCode").ToString
            Label5.Text = DT_dat2.Rows(0).Item("CardName").ToString
            Label7.Text = DT_dat2.Rows(0).Item("NumAtCard").ToString
            Label9.Text = DT_dat2.Rows(0).Item("DocDate").ToString
            Label11.Text = DateTime.Now.ToShortDateString
            con.ObtenerConexion.Close()

            Panel1.Visible = True

            Dim SQL_da As SqlDataAdapter = New SqlDataAdapter("SELECT t4.ItemCode as 'Codigo de Articulo',t4.ItemName as 'Descripcion del Articulo',SUM(CASE T4.Direction when 0 then 1 else -1 end * T4.Quantity) as TM, T4.BatchNum as Lote,'peso' as Peso,CONVERT(int,t3.U_Ancho) as Ancho,T3.U_Heat as Heat,t3.U_Coi as Coil,t4.whscode as Almacen
    FROM OITL T0
    INNER JOIN OPDN T2 on t2.DocEntry = t0.DocEntry		
    INNER JOIN ITL1 T1 ON T0.LogEntry = T1.LogEntry
    INNER JOIN OBTN T3 ON T1.MdAbsEntry = T3.AbsEntry
    inner join IBT1 T4 on T4.BatchNum = T3.DistNumber
    WHERE T0.DocEntry =  '" + TextBox1.Text + "' AND T0.DocNum =  '" + TextBox1.Text + "' and T0.BaseEntry = 0 and t4.WhsCode = 'BMP2'
    GROUP BY t4.ItemCode,t4.itemname,T4.BatchNum , T3.U_Heat,t3.U_Coi,t3.U_Ancho,t3.U_Correlativo,T4.whscode", con.ObtenerConexion())
            Dim DT_dat As System.Data.DataTable = New System.Data.DataTable()
            SQL_da.Fill(DT_dat)
            DGV2.DataSource = DT_dat
            For Each row As DataGridViewRow In DGV2.Rows
                row.Cells("CHK").Value = True
            Next
            con.ObtenerConexion.Close()
        Catch ex As Exception
            MessageBox.Show("La Entrada no Existe")
        End Try

    End Sub
End Class
