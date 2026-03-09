using ClosedXML.Excel;
using ClosedXML.Excel.Drawings;
using ExcelDataReader;
using MySql.Data.MySqlClient;
using Mysqlx;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Forms.DataVisualization.Charting;

namespace Administracion_OMAJA
{
    public partial class ADMINOMAJA : Form
    {
        private DatabaseManager dbManager = new DatabaseManager();
        private Timer timerUltimaCarga;
        private bool filtrosInicializados;
        private const string ColumnaSeleccionExportar = "SeleccionExportar";

        private readonly Dictionary<string, DataRow> seleccionPersistente = new Dictionary<string, DataRow>(StringComparer.OrdinalIgnoreCase);
        // Colores por EstatusGuia 

        private readonly Dictionary<string, Color> coloresEstatus = new Dictionary<string, Color>(StringComparer.OrdinalIgnoreCase)
        {
            { "DOCUMENTADO", Color.FromArgb(219, 234, 254) },   // azul pastel
            { "PENDIENTE", Color.FromArgb(255, 247, 214) },     // amarillo pastel
            { "EN RUTA", Color.FromArgb(213, 245, 227) },       // verde menta pastel
            { "ULTIMA MILLA", Color.FromArgb(237, 233, 254) },  // lila pastel (se mantiene)
            { "ENTREGADO", Color.FromArgb(221, 236, 255) },     // azul muy suave para entregado
            { "COMPLETADO", Color.FromArgb(189, 216, 255) }     // azul pastel para completado
        };

        // Colores por TipoCobro (prioridad 1)
        private readonly Dictionary<string, Color> coloresTipoCobro = new Dictionary<string, Color>(StringComparer.OrdinalIgnoreCase)
        {
            { "POR COBRAR", Color.FromArgb(255, 236, 179) },    // ámbar pastel
            { "CRÉDITO",   Color.FromArgb(198, 246, 213) },     // verde menta clara
            { "CREDITO",   Color.FromArgb(198, 246, 213) },
            { "PAGADO",    Color.FromArgb(209, 247, 196) },     // verde claro pastel
            { "CANCELADO", Color.FromArgb(255, 204, 204) }      // coral/rojo pastel
        };

        private void ColorearCeldaTipoCobro(DataGridView dgv)
        {
            if (dgv?.Rows == null || !dgv.Columns.Contains("TipoCobro"))
            {
                return;
            }

            foreach (DataGridViewRow row in dgv.Rows)
            {
                if (row.IsNewRow) continue;

                var cell = row.Cells["TipoCobro"];
                string valor = Convert.ToString(cell?.Value ?? string.Empty).Trim();

                if (string.IsNullOrWhiteSpace(valor))
                {
                    continue;
                }

                if (coloresTipoCobro.TryGetValue(valor.ToUpperInvariant(), out Color color))
                {
                    cell.Style.BackColor = color;
                    cell.Style.SelectionBackColor = ControlPaint.Dark(color);
                    cell.Style.SelectionForeColor = Color.Black;
                }
            }
        }

        private readonly Color colorFacturaConValor = Color.FromArgb(255, 255, 102); // Fosforito amarillo para facturas con valor

        private void ColorearFacturaConValor(DataGridView dgv)
        {
            if (dgv?.Rows == null || !dgv.Columns.Contains("Factura"))
            {
                return;
            }

            // Detectar la columna de folio (puede ser FolioGuia o folio_guia según la fuente de datos)
            string colFolio = dgv.Columns.Contains("FolioGuia")
                ? "FolioGuia"
                : dgv.Columns.Contains("folio_guia")
                    ? "folio_guia"
                    : null;

            foreach (DataGridViewRow row in dgv.Rows)
            {
                if (row.IsNewRow) continue;

                var facturaCell = row.Cells["Factura"];
                string valor = Convert.ToString(facturaCell?.Value ?? string.Empty).Trim();
                if (!string.IsNullOrEmpty(valor))
                {
                    // Pintar celda Factura
                    facturaCell.Style.BackColor = colorFacturaConValor;
                    facturaCell.Style.SelectionBackColor = ControlPaint.Dark(colorFacturaConValor);
                    facturaCell.Style.SelectionForeColor = Color.Black;

                    // Pintar también la celda de Folio si existe
                    if (colFolio != null)
                    {
                        var folioCell = row.Cells[colFolio];
                        if (folioCell != null)
                        {
                            folioCell.Style.BackColor = colorFacturaConValor;
                            folioCell.Style.SelectionBackColor = ControlPaint.Dark(colorFacturaConValor);
                            folioCell.Style.SelectionForeColor = Color.Black;
                        }
                    }
                }
            }
        }

        private Color? ObtenerColorPrioritario(string tipoCobro, string estatusGuia)
        {
            // Prioridad: TipoCobro > EstatusGuia (solo para Última Milla / Entregado / Completado y otros estatus conocidos)
            if (!string.IsNullOrWhiteSpace(tipoCobro) &&
                coloresTipoCobro.TryGetValue(tipoCobro.Trim().ToUpperInvariant(), out var colorCobro))
            {
                return colorCobro;
            }

            if (!string.IsNullOrWhiteSpace(estatusGuia) &&
                coloresEstatus.TryGetValue(estatusGuia.Trim().ToUpperInvariant(), out var colorEstatus))
            {
                return colorEstatus;
            }

            return null;
        }

        public ADMINOMAJA()
        {
            InitializeComponent();
            toolStripTextBoxBcliente.TextChanged += toolStripTextBoxBcliente_TextChanged;
            toolStripTextBoxBcliente.KeyDown += toolStripTextBoxBcliente_KeyDown;
            toolStripTextBoxBuscarFactura.KeyDown += toolStripTextBoxBuscarFactura_KeyDown;
            toolStripButtonBuscarFactura.Click += toolStripButtonBuscarFactura_Click;
            comboBoxSucursales.SelectedIndexChanged += comboBoxSucursales_SelectedIndexChanged;
            comboBoxSucursalDestino.SelectedIndexChanged += comboBoxSucursalDestino_SelectedIndexChanged;
            radioButtonSinfiltro.CheckedChanged += radioButtonSinfiltro_CheckedChanged;
            radioButtonPagadasEntregadas.CheckedChanged += radioButtonPagadasEntregadas_CheckedChanged;
            textBox1clientes.CharacterCasing = CharacterCasing.Upper;
            toolStriptxtBusqueda.KeyDown += toolStriptxtBusqueda_KeyDown;

            dataGridViewPrincipal.RowPostPaint += dataGridViewPrincipal_RowPostPaint;
            dataGridViewPrincipal.RowHeadersWidth = 56;
            dataGridViewPrincipal.RowHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dateTimePickerInicioCanceladas.ValueChanged += dateTimePickerInicioCanceladas_ValueChanged;
            dateTimePickerFinCanceladas.ValueChanged += dateTimePickerFinCanceladas_ValueChanged;
            comboBoxSucursalCancelaciones.SelectedIndexChanged += comboBoxSucursalCancelaciones_SelectedIndexChanged;

            dataGridViewPrincipal.DataBindingComplete += dataGridViewPrincipal_DataBindingComplete;
            dataGridViewPrincipal.CellEndEdit += dataGridViewPrincipal_CellEndEdit;
            dataGridViewPrincipal.KeyDown += dataGridViewPrincipal_KeyDown;

            checkBoxEditable.CheckedChanged += checkBoxEditable_CheckedChanged;
            checkBoxExportarTodo.CheckedChanged += checkBoxExportarTodo_CheckedChanged;
            checkBoxExportarSeleccion.CheckedChanged += checkBoxExportarSeleccion_CheckedChanged;
            buttonbuscarestatus.Click += buttonbuscarestatus_Click;
            buttonExportarStatus.Click += buttonExportarStatus_Click;

            dataGridViewPrincipal.CellValueChanged += dataGridViewPrincipal_CellValueChanged;
            dataGridViewPrincipal.CurrentCellDirtyStateChanged += dataGridViewPrincipal_CurrentCellDirtyStateChanged;

            this.Shown += ADMINOMAJA_Shown_InicializarSeguimiento; 
            dataGridViewSeguimientos.CellClick += dataGridViewSeguimientos_CellClick;

            dataGridViewPrincipal.DataError += dataGridViewPrincipal_DataError;
        }

        private void ADMINOMAJA_Shown_InicializarSeguimiento(object sender, EventArgs e)
        {
            ConfigurarComboOrigenesSeguimiento();
            ConfigurarComboDestinosSeguimiento(); // NUEVO
            CargarResumenSeguimientos();
        }

        private void dataGridViewPrincipal_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {
            var grid = sender as DataGridView;
            var colName = grid?.Columns[e.ColumnIndex]?.Name ?? "(col)";
            var val = grid?.Rows[e.RowIndex]?.Cells[e.ColumnIndex]?.Value;
            MessageBox.Show($"DataError en columna '{colName}'. Valor: '{val ?? "null"}'. Error: {e.Exception?.Message}",
                "DataError", MessageBoxButtons.OK, MessageBoxIcon.Error);
            e.ThrowException = true; // mantiene la excepción visible tras el mensaje
        }

        private void dataGridViewPrincipal_RowPostPaint(object sender, DataGridViewRowPostPaintEventArgs e)
        {
            var grid = sender as DataGridView;
            if (grid == null)
            {
                return;
            }

            string numero = (e.RowIndex + 1).ToString(CultureInfo.InvariantCulture);
            Rectangle bounds = new Rectangle(e.RowBounds.Left, e.RowBounds.Top, grid.RowHeadersWidth, e.RowBounds.Height);

            TextRenderer.DrawText(
                e.Graphics,
                numero,
                grid.RowHeadersDefaultCellStyle.Font ?? grid.Font,
                bounds,
                grid.RowHeadersDefaultCellStyle.ForeColor,
                TextFormatFlags.VerticalCenter | TextFormatFlags.Right);
        }

        private void ActualizarDataGridView()
        {
            DataTable datos = dbManager.ObtenerTodasGuias();
            dataGridViewPrincipal.DataSource = datos;
            FormatearEncabezadosDataGridView(dataGridViewPrincipal);
            AplicarEstiloAzulClaroDataGridView(dataGridViewPrincipal);
        }

        private void ADMINOMAJA_Load(object sender, EventArgs e)
        {
            CargarEstadoCargaExcel();
            ActualizarEstadoCargaExcel();
            ActualizarDataGridView();
            IniciarTimerUltimaCarga();

            if (comboBoxSucursales.Items.Count > 0)
            {
                comboBoxSucursales.SelectedIndex = 0;
            }

            if (comboBoxSucursalDestino.Items.Count > 0)
            {
                comboBoxSucursalDestino.SelectedIndex = 0;
            }

            comboBoxSucursales.Enabled = true;
            comboBoxSucursalDestino.Enabled = true;
            toolStripProgressBar1.Visible = false;

            radioButtonSinfiltro.Checked = true;
            OcultarIndicadoresResumen();

            CargarGraficoVentas();

            filtrosInicializados = true;

            if (comboBoxSucursalCancelaciones.Items.Count > 0)
            {
                int indexTodas = comboBoxSucursalCancelaciones.FindStringExact("TODAS");
                comboBoxSucursalCancelaciones.SelectedIndex = indexTodas >= 0 ? indexTodas : 0;
            }
        }


        private class EstadoCargaExcel
        {
            public List<RegistroCarga> Historial { get; set; } = new List<RegistroCarga>();
        }

        private string EstadoCargaPath => Path.Combine(Application.StartupPath, "estadocargaexcel.json");

        private void GuardarEstadoCargaExcel(int documentosCargados, int nuevos, int actualizados)
        {
            EstadoCargaExcel estado;
            if (File.Exists(EstadoCargaPath))
            {
                estado = JsonConvert.DeserializeObject<EstadoCargaExcel>(File.ReadAllText(EstadoCargaPath));
            }
            else
            {
                estado = new EstadoCargaExcel();
            }

            estado.Historial.Add(new RegistroCarga
            {
                FechaHora = DateTime.Now,
                DocumentosCargados = documentosCargados,
                Nuevos = nuevos,
                Actualizados = actualizados
            });

            File.WriteAllText(EstadoCargaPath, JsonConvert.SerializeObject(estado));
        }

        private void CargarEstadoCargaExcel()
        {
            if (File.Exists(EstadoCargaPath))
            {
                var estado = JsonConvert.DeserializeObject<EstadoCargaExcel>(File.ReadAllText(EstadoCargaPath));
                dbManager.CargarHistorialDesdeEstado(estado.Historial);
            }
        }




        // === BOTÓN 5: ELIMINAR ===
        private void btnEliminar_Click(object sender, EventArgs e)
        {
            if (dataGridViewPrincipal.CurrentRow != null)
            {
                string folioGuia = dataGridViewPrincipal.CurrentRow.Cells["folio_guia"].Value?.ToString();

                if (!string.IsNullOrEmpty(folioGuia))
                {
                    DialogResult result = MessageBox.Show(
                        $"¿Está seguro de eliminar la guía {folioGuia}?",
                        "Confirmar eliminación",
                        MessageBoxButtons.YesNo,
                        MessageBoxIcon.Warning);

                    if (result == DialogResult.Yes)
                    {
                        bool eliminado = dbManager.EliminarGuia(folioGuia);
                        if (eliminado)
                        {
                            MessageBox.Show("Guía eliminada exitosamente", "Éxito",
                                MessageBoxButtons.OK, MessageBoxIcon.Information);
                            buttonBuscar_Click(sender, e);
                        }
                    }
                }
            }
        }



        // ==================== BOTÓN 7: LIMPIAR ====================
        private void btnLimpiar_Click(object sender, EventArgs e)
        {
            dataGridViewPrincipal.DataSource = null;
            toolStriptxtBusqueda.Clear();
        }


        private void toolStripDropDownButton1_Click(object sender, EventArgs e)
        {

        }

        private void toolStripDropDownButton1_Click_1(object sender, EventArgs e)
        {
            toolStripDropDownButton1.ShowDropDown();
        }

        private void labelGuiasOrigen_Click(object sender, EventArgs e)
        {
        }

        // ==================== MÉTODO IMPORTAR EXCEL ====================




        private void importarExcelToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Archivos Excel|*.xlsx;*.xls";

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                toolStripProgressBar1.Value = 0;
                toolStripProgressBar1.Visible = true;
                toolStripProgressBar1.Minimum = 0;
                toolStripProgressBar1.Maximum = 100;

                Task.Run(() =>
                {
                    ImportarDesdeExcelConProgreso(openFileDialog.FileName, dataGridViewPrincipal, (actual, total) =>
                    {
                        int percent = (int)((actual * 100.0) / total);
                        toolStripProgressBar1.GetCurrentParent().Invoke((Action)(() =>
                        {
                            toolStripProgressBar1.Value = percent;
                        }));
                    });

                    this.Invoke((Action)(() =>
                    {
                        toolStripProgressBar1.Visible = false;
                        MessageBox.Show("Importación de Excel completada.", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        FormatearEncabezadosDataGridView(dataGridViewPrincipal);
                        AplicarEstiloAzulClaroDataGridView(dataGridViewPrincipal);
                        ActualizarEstadoCargaExcel();
                        CargarGraficoVentas();
                    }));
                });
            }
        }


        private void ImportarDesdeExcelConProgreso(string filePath, DataGridView dataGridView, Action<int, int> reportProgress)
        {
            dbManager.ImportarDesdeExcel(filePath, dataGridView, reportProgress);

            int nuevos = 0;
            int actualizados = 0;

            FormatearEncabezadosDataGridView(dataGridView);
            AplicarEstiloAzulClaroDataGridView(dataGridViewPrincipal);

            int documentosCargados = dataGridView.Rows.Cast<DataGridViewRow>().Count(r => !r.IsNewRow);
            GuardarEstadoCargaExcel(documentosCargados, nuevos, actualizados);

            if (InvokeRequired)
            {
                this.Invoke((Action)(() =>
                {
                    ActualizarEstadoCargaExcel();
                    CargarGraficoVentas();
                }));
            }
            else
            {
                ActualizarEstadoCargaExcel();
                CargarGraficoVentas();
            }
        }

        private IEnumerable<(DateTime Inicio, DateTime Fin, string Nombre)> ObtenerRangosMensuales(DateTime inicio, DateTime fin)
        {
            if (fin < inicio)
            {
                var tmp = inicio;
                inicio = fin;
                fin = tmp;
            }

            DateTime cursor = new DateTime(inicio.Year, inicio.Month, 1);
            while (cursor <= fin)
            {
                var monthStart = cursor;
                var monthEnd = monthStart.AddMonths(1).AddDays(-1);

                var rangoInicio = monthStart < inicio ? inicio : monthStart;
                var rangoFin = monthEnd > fin ? fin : monthEnd;

                yield return (rangoInicio, rangoFin, $"{monthStart:yyyy-MM}");
                cursor = monthStart.AddMonths(1);
            }
        }

        private (int GuiasTotales, int GuiasPorCobrar, decimal MontoPorCobrar, int GuiasCanceladas, decimal MontoCanceladas) CalcularEstadisticasMes(DataTable dt)
        {
            if (dt == null || dt.Rows.Count == 0)
            {
                return (0, 0, 0m, 0, 0m);
            }

            var filasNoCanceladas = dt.AsEnumerable().Where(r => !EsGuiaCancelada(r)).ToList();
            int guiasTotales = filasNoCanceladas.Count;

            var porCobrar = filasNoCanceladas
                .Where(r => string.Equals(Convert.ToString(r["TipoCobro"]), "POR COBRAR", StringComparison.OrdinalIgnoreCase))
                .ToList();

            int guiasPorCobrar = porCobrar.Count;
            decimal montoPorCobrar = porCobrar.Sum(r => ConvertToDecimal(r["Total"]));

            var canceladas = dt.AsEnumerable()
                .Where(r => string.Equals(Convert.ToString(r["EstatusGuia"]), "CANCELADO", StringComparison.OrdinalIgnoreCase))
                .ToList();
            int guiasCanceladas = canceladas.Count;
            decimal montoCanceladas = canceladas.Sum(r => ConvertToDecimal(r["Total"]));

            return (guiasTotales, guiasPorCobrar, montoPorCobrar, guiasCanceladas, montoCanceladas);
        }

        private void exportarExcelToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            // Exporta la vista actual (filtros/orden/columnas visibles) con encabezado, resumen, gráfica y colores.
            ExportarVistaActualConFormato();
        }

        private bool ExportarReporteExcel(string filePath)
        {
            DateTime fechaInicio = dtpFechaInicio.Value.Date;
            DateTime fechaFin = dtpFechaFin.Value.Date;
            if (fechaFin < fechaInicio) { var t = fechaInicio; fechaInicio = fechaFin; fechaFin = t; }

            DataTable datos = ObtenerDatosFiltradosParaExport(fechaInicio, fechaFin);
            if (datos == null || datos.Rows.Count == 0)
            {
                MessageBox.Show("No hay datos para exportar con los filtros actuales.",
                    "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }

            string sucursal = comboBoxSucursales.SelectedItem?.ToString() ?? "TODAS";
            string destino = comboBoxSucursalDestino.SelectedItem?.ToString() ?? "TODAS";

            if (chartVentasdiarias.Series.Count == 0)
            {
                CargarGraficoVentas();
            }

            using (var workbook = new XLWorkbook())
            {
                var resumenSheet = workbook.Worksheets.Add("Resumen");
                string mesTitulo = $"Reporte de {fechaInicio:MMMM yyyy} OMAJA";
                InsertarEncabezadoMes(resumenSheet, mesTitulo, sucursal, destino);

                var indicadoresCalc = ConstruirIndicadoresResumen(fechaInicio, fechaFin, sucursal, destino, datos);
                int nextRow = EscribirIndicadoresTabla(resumenSheet, 5, indicadoresCalc);

                var indicadoresGraf = FiltrarIndicadoresParaGrafico(indicadoresCalc);
                int chartRow = nextRow + 1;
                InsertarGraficoIndicadores(resumenSheet, chartRow, 1, indicadoresGraf);

                using (var chartStream = new MemoryStream())
                {
                    chartVentasdiarias.SaveImage(chartStream, ChartImageFormat.Png);
                    chartStream.Position = 0;
                    var picture = resumenSheet.AddPicture(chartStream, XLPictureFormat.Png, "Ventas");
                    picture.MoveTo(resumenSheet.Cell(chartRow + 22, 1));
                    picture.WithSize(800, 400);
                }

                var datosSheet = workbook.Worksheets.Add("Datos");
                var tablaDatos = datosSheet.Cell(1, 1).InsertTable(datos, "Guias");
                tablaDatos.Theme = XLTableTheme.None;
                tablaDatos.ShowRowStripes = false;
                tablaDatos.ShowColumnStripes = false;
                AplicarColoresFilasExcel(datosSheet, datos, startRow: 2, startCol: 1);
                datosSheet.Columns().AdjustToContents();

                var totSheet = workbook.Worksheets.Add("Totales");
                totSheet.Cell("A1").Value = "Indicador";
                totSheet.Cell("B1").Value = "Cantidad";
                totSheet.Cell("C1").Value = "Monto";
                totSheet.Range("A1:C1").Style.Font.Bold = true;

                int rowTot = 2;
                foreach (var ind in indicadoresCalc)
                {
                    totSheet.Cell(rowTot, 1).Value = ind.Titulo;
                    totSheet.Cell(rowTot, 2).Value = ind.Cantidad;
                    totSheet.Cell(rowTot, 3).Value = ind.Monto;
                    rowTot++;
                }

                totSheet.Column(3).Style.NumberFormat.Format = "$#,##0.00";
                totSheet.Columns().AdjustToContents();

                workbook.SaveAs(filePath);
            }

            return true;
        }

        private string ObtenerNombreColumna(DataTable dt, params string[] nombresPosibles)
        {
            if (dt == null || dt.Columns.Count == 0 || nombresPosibles == null || nombresPosibles.Length == 0)
            {
                return null;
            }

            foreach (string nombre in nombresPosibles)
            {
                var columna = dt.Columns.Cast<DataColumn>()
                    .FirstOrDefault(c => c.ColumnName.Equals(nombre, StringComparison.OrdinalIgnoreCase));

                if (columna != null)
                {
                    return columna.ColumnName;
                }
            }

            return null;
        }

        private void MostrarMensajeExportacionCompletada(string filePath)
        {
            var resultado = MessageBox.Show(
                "Exportación completada correctamente.\n¿Deseas abrir el archivo?",
                "Exportación completada",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Information);

            if (resultado == DialogResult.Yes)
            {
                System.Diagnostics.Process.Start(filePath);
            }
        }

        private DataTable ObtenerDatosFiltradosParaExport(DateTime fechaInicio, DateTime fechaFin)
        {
            if (fechaFin < fechaInicio)
            {
                var temp = fechaInicio;
                fechaInicio = fechaFin;
                fechaFin = temp;
            }

            string sucursal = comboBoxSucursales.SelectedItem?.ToString() ?? "TODAS";
            string destino = comboBoxSucursalDestino.SelectedItem?.ToString() ?? "TODAS";

            if (radioButtonGuiasporcobrar.Checked)
            {
                return dbManager.ObtenerGuiasPorCobrar(fechaInicio, fechaFin, sucursal, destino);
            }

            if (radioButtonGuiasconcredito.Checked)
            {
                return dbManager.ObtenerGuiasConCredito(fechaInicio, fechaFin, sucursal, destino);
            }

            if (radioButtonGuiasPagadas.Checked)
            {
                return dbManager.ObtenerGuiasPagadas(fechaInicio, fechaFin, sucursal, destino);
            }

            if (radioButtonPagadasEntregadas.Checked)
            {
                DataTable origen = dbManager.ObtenerTodasGuiasFiltrado(fechaInicio, fechaFin, sucursal, destino);
                return FiltrarGuiasPorEstatusPagado(origen, "ULTIMA MILLA", "ENTREGADO", "COMPLETADO");
            }

            if (radioButtonGuiasCanceladas.Checked)
            {
                DataTable todas = dbManager.ObtenerTodasGuiasFiltrado(fechaInicio, fechaFin, sucursal, destino);
                if (todas == null)
                {
                    return null;
                }

                DataTable canceladas = todas.Clone();
                foreach (DataRow row in todas.Rows)
                {
                    string estatus = (row["EstatusGuia"]?.ToString() ?? string.Empty).Trim();
                    if (estatus.Equals("Cancelado", StringComparison.OrdinalIgnoreCase))
                    {
                        canceladas.ImportRow(row);
                    }
                }

                return canceladas;
            }

            return dbManager.ObtenerTodasGuiasFiltrado(fechaInicio, fechaFin, sucursal, destino);
        }
        private void toolStrip1_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {

        }

        private void toolStripBusqueda_Click(object sender, EventArgs e)
        {
            string folio = toolStriptxtBusqueda.Text.Trim();

            if (!string.IsNullOrEmpty(folio))
            {
                string error;
                DataTable resultados = dbManager.BuscarGuias("FolioGuia", folio, out error);

                if (!string.IsNullOrEmpty(error))
                {
                    MessageBox.Show(error, "Error de búsqueda", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                if (resultados != null && resultados.Rows.Count > 0)
                {
                    DataTable dtOrdenada = new DataTable();

                    string colFolio = "FolioGuia";
                    dtOrdenada.Columns.Add(colFolio, resultados.Columns[colFolio].DataType);

                    string colFecha = resultados.Columns.Contains("FechaElaboracion") ? "FechaElaboracion" : null;
                    if (colFecha != null && colFecha != colFolio)
                        dtOrdenada.Columns.Add(colFecha, resultados.Columns[colFecha].DataType);

                    foreach (DataColumn col in resultados.Columns)
                    {
                        if (col.ColumnName != colFolio && col.ColumnName != colFecha)
                            dtOrdenada.Columns.Add(col.ColumnName, col.DataType);
                    }

                    foreach (DataRow row in resultados.Rows)
                    {
                        var newRow = dtOrdenada.NewRow();
                        foreach (DataColumn col in dtOrdenada.Columns)
                        {
                            newRow[col.ColumnName] = row[col.ColumnName];
                        }
                        dtOrdenada.Rows.Add(newRow);
                    }

                    CargarGridConSeleccionPersistente(dtOrdenada);
                }
                else
                {
                    dataGridViewPrincipal.DataSource = null;
                    MessageBox.Show("No se encontró el folio especificado.", "Sin resultados", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            else
            {
                MessageBox.Show("Por favor, ingresa un folio para buscar.", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void toolStripLabel3_Click(object sender, EventArgs e)
        {

        }

        private void ActualizarEstadoCargaExcel()
        {
            if (dbManager.SeCargoHoy)
            {
                toolStripLabel3.Text = $"Hoy se han cargado {dbManager.DocumentosCargadosHoy} registros desde Excel.";
            }
            else
            {
                toolStripLabel3.Text = "No se ha cargado ningún documento de Excel el día de hoy.";
            }
        }

        private void toolStripDropDownButton2_Click(object sender, EventArgs e)
        {

        }

        private void importarBasesqlToolStripMenuItem_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Filter = "Archivo SQL|*.sql",
                Title = "Importar base de datos MySQL"
            };

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                string mysqlPath = @"C:\Program Files\MySQL\MySQL Server 8.0\bin\mysql.exe";
                string filePath = openFileDialog.FileName;
                string user = "root";
                string password = "omaja123";
                string database = "adminomaja";
                string arguments = $"-u{user} -p{password} {database}";

                try
                {
                    var process = new System.Diagnostics.Process
                    {
                        StartInfo = new System.Diagnostics.ProcessStartInfo
                        {
                            FileName = mysqlPath,
                            Arguments = arguments,
                            RedirectStandardInput = true,
                            UseShellExecute = false,
                            CreateNoWindow = true
                        }
                    };

                    process.Start();

                    string sql = System.IO.File.ReadAllText(filePath, Encoding.UTF8);
                    process.StandardInput.WriteLine(sql);
                    process.StandardInput.Close();

                    process.WaitForExit();

                    MessageBox.Show("Importación completada exitosamente.", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Error al importar la base de datos: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void exportarBasesqlToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog
            {
                Filter = "Archivo SQL|*.sql",
                Title = "Exportar base de datos MySQL"
            };

            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                string mysqldumpPath = @"C:\Program Files\MySQL\MySQL Server 8.0\bin\mysqldump.exe";
                string filePath = saveFileDialog.FileName;
                string user = "root";
                string password = "omaja123";
                string database = "adminomaja";
                string arguments = $"--routines --triggers --events --single-transaction --databases {database} -u{user} -p{password}";

                try
                {
                    var process = new System.Diagnostics.Process
                    {
                        StartInfo = new System.Diagnostics.ProcessStartInfo
                        {
                            FileName = mysqldumpPath,
                            Arguments = arguments,
                            RedirectStandardOutput = true,
                            UseShellExecute = false,
                            CreateNoWindow = true
                        }
                    };

                    process.Start();

                    using (var fileStream = new System.IO.StreamWriter(filePath, false, Encoding.UTF8))
                    {
                        fileStream.Write(process.StandardOutput.ReadToEnd());
                    }

                    process.WaitForExit();

                    MessageBox.Show("Exportación completada exitosamente.", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Error al exportar la base de datos: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void buttonBuscar_Click(object sender, EventArgs e)
        {
            DateTime fechaInicio = dtpFechaInicio.Value.Date;
            DateTime fechaFin = dtpFechaFin.Value.Date;
            string sucursal = comboBoxSucursales.SelectedItem?.ToString() ?? "TODAS";
            string destino = comboBoxSucursalDestino.SelectedItem?.ToString() ?? "TODAS";
            bool requiereSucursalEspecifica = !sucursal.Equals("TODAS", StringComparison.OrdinalIgnoreCase);
            bool requiereDestinoEspecifico = !destino.Equals("TODAS", StringComparison.OrdinalIgnoreCase);

            if (fechaFin < fechaInicio)
            {
                MessageBox.Show("La fecha final no puede ser menor que la fecha inicial.", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            bool esCanceladas = radioButtonGuiasCanceladas.Checked;
            bool esTodosLosFiltros = radioButtonTodoslosFiltros.Checked;
            bool esPagadasEntregadas = radioButtonPagadasEntregadas.Checked;
            DataTable dt = null;
            int totalCanceladas = 0;
            decimal sumaTotalCanceladas = 0m;

            OcultarIndicadoresResumen();

            if (esCanceladas)
            {
                DataTable todas = dbManager.ObtenerTodasGuiasFiltrado(fechaInicio, fechaFin, sucursal, destino);
                DataTable dtCanceladas = new DataTable();

                if (todas != null && todas.Rows.Count > 0)
                {
                    dtCanceladas = todas.Clone();

                    foreach (DataRow row in todas.Rows)
                    {
                        string estatus = (row["EstatusGuia"]?.ToString() ?? string.Empty).Trim();
                        if (estatus.Equals("Cancelado", StringComparison.OrdinalIgnoreCase))
                        {
                            dtCanceladas.ImportRow(row);
                            totalCanceladas++;
                            if (todas.Columns.Contains("Total") &&
                                decimal.TryParse(row["Total"]?.ToString(), out decimal valor))
                            {
                                sumaTotalCanceladas += valor;
                            }
                        }
                    }
                }

                dt = dtCanceladas;
                ActualizarLabelResumen(labelGuiasCanceladas, totalCanceladas, sumaTotalCanceladas);
            }
            else if (radioButtonGuiasporcobrar.Checked)
            {
                dt = dbManager.ObtenerGuiasPorCobrar(fechaInicio, fechaFin, sucursal, destino);
                var filasValidas = FilasNoCanceladas(dt).ToList();

                labelnumerodeguiasporcobrar.Text = filasValidas.Count.ToString();
                labelnumerodeguiasporcobrar.Font = new Font(labelnumerodeguiasporcobrar.Font, FontStyle.Bold);
                labelnumerodeguiasporcobrar.ForeColor = Color.Black;
                labelnumerodeguiasporcobrar.Visible = true;

                decimal sumaTotal = 0m;
                if (dt.Columns.Contains("Total"))
                {
                    foreach (DataRow row in filasValidas)
                    {
                        if (decimal.TryParse(row["Total"]?.ToString(), out decimal valor))
                        {
                            sumaTotal += valor;
                        }
                    }
                }

                labelmontototalporcobrar.Text = $"{filasValidas.Count}\n{sumaTotal:C}";
                labelmontototalporcobrar.Font = new Font(labelmontototalporcobrar.Font, FontStyle.Bold);
                labelmontototalporcobrar.ForeColor = sumaTotal < 50000m ? Color.Green : Color.Red;
                labelmontototalporcobrar.Visible = true;
            }
            else if (radioButtonGuiasconcredito.Checked)
            {
                dt = dbManager.ObtenerGuiasConCredito(fechaInicio, fechaFin, sucursal, destino);
            }
            else if (radioButtonGuiasPagadas.Checked)
            {
                DataTable dtPagadas = dbManager.ObtenerGuiasPagadas(fechaInicio, fechaFin, sucursal, destino);
                dt = dtPagadas;
                ActualizarResumenPagosPorUbicacion(fechaInicio, fechaFin, sucursal, destino);
            }
            else if (esPagadasEntregadas)
            {
                DataTable filtradas = dbManager.ObtenerTodasGuiasFiltrado(fechaInicio, fechaFin, sucursal, destino);
                dt = FiltrarGuiasPorEstatusPagado(filtradas, "ULTIMA MILLA", "ENTREGADO", "COMPLETADO");
            }
            else if (esTodosLosFiltros)
            {
                dt = dbManager.ObtenerTodasGuiasFiltrado(fechaInicio, fechaFin, "TODAS", "TODAS");

                if (requiereSucursalEspecifica || requiereDestinoEspecifico)
                {
                    dt = FiltrarGuiasPorUbicacion(dt, sucursal, destino);
                }

                ActualizarIndicadoresTodosLosFiltros(dt);
                ActualizarResumenPagosPorUbicacion(fechaInicio, fechaFin, sucursal, destino);
            }
            else
            {
                dt = dbManager.ObtenerTodasGuiasFiltrado(fechaInicio, fechaFin, sucursal, destino);
            }

            if (dt == null || dt.Rows.Count == 0)
            {
                dataGridViewPrincipal.DataSource = null;
                MessageBox.Show("No se encontraron resultados para los criterios seleccionados.", "Sin resultados", MessageBoxButtons.OK, MessageBoxIcon.Information);
                labelGuiastotales.Text = "Total de guías: 0";
                return;
            }

            CargarGridConSeleccionPersistente(dt);

            if (radioButtonGuiasporcobrar.Checked)
            {
                var filasValidas = FilasNoCanceladas(dt).ToList();

                labelnumerodeguiasporcobrar.Text = filasValidas.Count.ToString();
                labelnumerodeguiasporcobrar.Font = new Font(labelnumerodeguiasporcobrar.Font, FontStyle.Bold);
                labelnumerodeguiasporcobrar.ForeColor = Color.Black;
                labelnumerodeguiasporcobrar.Visible = true;

                decimal sumaTotal = 0m;
                if (dt.Columns.Contains("Total"))
                {
                    foreach (DataRow row in filasValidas)
                    {
                        if (decimal.TryParse(row["Total"]?.ToString(), out decimal valor))
                        {
                            sumaTotal += valor;
                        }
                    }
                }

                labelmontototalporcobrar.Text = $"{filasValidas.Count}\n{sumaTotal:C}";
                labelmontototalporcobrar.Font = new Font(labelmontototalporcobrar.Font, FontStyle.Bold);
                labelmontototalporcobrar.ForeColor = sumaTotal < 50000m ? Color.Green : Color.Red;
                labelmontototalporcobrar.Visible = true;
            }

            if (esPagadasEntregadas || esTodosLosFiltros)
            {
                ActualizarIndicadoresEstadosEntrega(dt);
            }
            else
            {
                OcultarIndicadoresEntrega();
            }

            int totalGuias = esCanceladas ? totalCanceladas : ContarGuiasNoCanceladas(dt);
            labelGuiastotales.Text = totalGuias.ToString();
            labelGuiastotales.Font = new Font(labelGuiastotales.Font, FontStyle.Bold);
            labelGuiastotales.ForeColor = Color.Navy;
        }


        private static DataTable FiltrarGuiasPorEstatusPagado(DataTable origen, params string[] estatusPermitidos)
        {
            if (origen == null || origen.Rows.Count == 0 || estatusPermitidos == null || estatusPermitidos.Length == 0)
            {
                return origen;
            }

            var estatus = new HashSet<string>(estatusPermitidos, StringComparer.OrdinalIgnoreCase);
            var filas = origen.AsEnumerable()
                .Where(row =>
                {
                    string valorEstatus = (row["EstatusGuia"]?.ToString() ?? string.Empty).Trim();
                    string valorCobro = row["TipoCobro"]?.ToString() ?? string.Empty;
                    return estatus.Contains(valorEstatus) && ContienePagado(valorCobro);
                })
                .ToList();

            return filas.Any() ? filas.CopyToDataTable() : origen.Clone();
        }

        private static bool ContienePagado(string valorTipoCobro)
        {
            return !string.IsNullOrWhiteSpace(valorTipoCobro) &&
                   valorTipoCobro.IndexOf("PAGADO", StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private static DataTable FiltrarGuiasPorTipoCobro(
    DataTable origen,
    string[] tiposCobro = null,
    string[] estatusGuia = null,
    bool requiereTipoCobroPagado = false)
        {
            if (origen == null || origen.Rows.Count == 0)
            {
                return origen;
            }

            HashSet<string> tiposPermitidos = tiposCobro != null && tiposCobro.Length > 0
                ? new HashSet<string>(tiposCobro, StringComparer.OrdinalIgnoreCase)
                : null;

            HashSet<string> estatusPermitidos = estatusGuia != null && estatusGuia.Length > 0
                ? new HashSet<string>(estatusGuia, StringComparer.OrdinalIgnoreCase)
                : null;

            var filas = origen.AsEnumerable()
                .Where(row =>
                {
                    string valorTipoCobro = (row["TipoCobro"]?.ToString() ?? string.Empty).Trim();
                    string valorEstatus = (row["EstatusGuia"]?.ToString() ?? string.Empty).Trim();

                    bool coincideTipo = tiposPermitidos == null || tiposPermitidos.Contains(valorTipoCobro);
                    bool coincideEstatus = estatusPermitidos == null || estatusPermitidos.Contains(valorEstatus);
                    bool coincidePagado = !requiereTipoCobroPagado || ContienePagado(valorTipoCobro);

                    return coincideTipo && coincideEstatus && coincidePagado;
                })
                .ToList();

            return filas.Any() ? filas.CopyToDataTable() : origen.Clone();
        }

        private void ActualizarIndicadoresEstadosEntrega(DataTable dt)
        {
            if (dt == null || dt.Rows.Count == 0)
            {
                OcultarIndicadoresEntrega();
                return;
            }

            ActualizarIndicadorEstado(labelUltimaMilla, dt, "ULTIMA MILLA", Color.SteelBlue);
            ActualizarIndicadorEstado(labelEntregado, dt, "ENTREGADA", Color.MediumPurple);
            ActualizarIndicadorEstado(labelCompletado, dt, "COMPLETADO", Color.ForestGreen);
        }

        private void ActualizarIndicadorEstado(Label label, DataTable dt, string estatusObjetivo, Color color)
        {
            if (label == null)
            {
                return;
            }

            var filas = dt.AsEnumerable()
                .Where(row =>
                    string.Equals((row["EstatusGuia"]?.ToString() ?? string.Empty).Trim(), estatusObjetivo, StringComparison.OrdinalIgnoreCase) &&
                    ContienePagado(row["TipoCobro"]?.ToString() ?? string.Empty))
                .ToList();

            decimal monto = filas.Sum(row => ConvertToDecimal(row["Total"]));
            label.Text = $"{filas.Count}{Environment.NewLine}{monto:C}";
            label.Font = new Font(label.Font, FontStyle.Bold);
            label.ForeColor = color;
            label.Visible = true;
        }

        private void OcultarIndicadoresEntrega()
        {
            OcultarLabel(labelUltimaMilla);
            OcultarLabel(labelEntregado);
            OcultarLabel(labelCompletado);
        }

        private DataTable FiltrarGuiasPorUbicacion(DataTable origen, string sucursalSeleccionada, string destinoSeleccionado)
        {
            if (origen == null || origen.Rows.Count == 0)
            {
                return origen;
            }

            IEnumerable<DataRow> filas = origen.AsEnumerable();

            bool filtraSucursal = !string.IsNullOrWhiteSpace(sucursalSeleccionada) &&
                                  !sucursalSeleccionada.Equals("TODAS", StringComparison.OrdinalIgnoreCase);
            if (filtraSucursal)
            {
                string sucursalNormalizada = NormalizarSucursalNombre(sucursalSeleccionada);

                var columnas = new List<string>();
                if (origen.Columns.Contains("Sucursal"))
                {
                    columnas.Add("Sucursal");
                }

                if (origen.Columns.Contains("Origen"))
                {
                    columnas.Add("Destino");
                }

                if (origen.Columns.Contains("Destino"))
                {
                    columnas.Add("Origen");
                }

                if (columnas.Count > 0)
                {
                    filas = filas.Where(row => columnas.Any(columna =>
                    {
                        string valor = row[columna]?.ToString();
                        return !string.IsNullOrWhiteSpace(valor) &&
                               NormalizarSucursalNombre(valor) == sucursalNormalizada;
                    }));
                }
            }

            bool filtraDestino = !string.IsNullOrWhiteSpace(destinoSeleccionado) &&
                                 !destinoSeleccionado.Equals("TODAS", StringComparison.OrdinalIgnoreCase) &&
                                 origen.Columns.Contains("Destino");
            if (filtraDestino)
            {
                string destinoNormalizado = NormalizarSucursalNombre(destinoSeleccionado);
                filas = filas.Where(row => NormalizarSucursalNombre(row["Destino"]?.ToString()) == destinoNormalizado);
            }

            return filas.Any() ? filas.CopyToDataTable() : origen.Clone();
        }


        private static string NormalizarSucursalNombre(string valor)
        {
            if (string.IsNullOrWhiteSpace(valor))
            {
                return string.Empty;
            }

            string resultado = valor.Trim().ToUpperInvariant();
            resultado = resultado.Normalize(NormalizationForm.FormD);

            var sb = new StringBuilder(resultado.Length);
            foreach (char c in resultado)
            {
                var categoria = CharUnicodeInfo.GetUnicodeCategory(c);
                if (categoria == UnicodeCategory.NonSpacingMark)
                {
                    continue;
                }

                if (char.IsLetterOrDigit(c) || c == '/' || c == ' ')
                {
                    sb.Append(c);
                }
            }

            resultado = sb.ToString()
                         .Replace("(", string.Empty)
                         .Replace(")", string.Empty)
                         .Replace("-", string.Empty)
                         .Replace("  ", " ")
                         .Trim();

            return resultado;
        }

        private void ActualizarLabelResumen(Label label, int cantidad, decimal monto)
        {
            label.Text = $"{cantidad}\n{monto:C}";
            label.Font = new Font(label.Font, FontStyle.Bold);
            label.ForeColor = monto < 50000 ? Color.Green : Color.Red;
            label.Visible = true;
        }

        private void comboBoxSucursales_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (!filtrosInicializados)
            {
                return;
            }

            bool requiereActualizarResumen = radioButtonGuiasPagadas.Checked || radioButtonTodoslosFiltros.Checked;

            if (requiereActualizarResumen)
            {
                LimpiarIndicadoresGuiasPagadas();
            }

            if (radioButtonGuiasPagadas.Checked || radioButtonGuiasporcobrar.Checked ||
                radioButtonGuiasconcredito.Checked || radioButtonGuiasCanceladas.Checked ||
                radioButtonTodoslosFiltros.Checked)
            {
                dataGridViewPrincipal.DataSource = null;
                labelGuiastotales.Text = "Total de guías: 0";
                buttonBuscar.PerformClick();
            }
        }

        private void radioButtonGuiasPagadas_CheckedChanged(object sender, EventArgs e)
        {
            HabilitarComboBoxSucursales();

            if (radioButtonGuiasPagadas.Checked)
            {
                LimpiarIndicadoresGuiasPagadas();
                dataGridViewPrincipal.DataSource = null;
                labelGuiastotales.Text = "Total de guías: 0";
            }
            else
            {
                LimpiarIndicadoresGuiasPagadas();
            }
        }

        private void OcultarLabel(Label label)
        {
            label.Text = string.Empty;
            label.Visible = false;
        }

        private void OcultarIndicadoresResumen()
        {
            OcultarLabel(labelGuiasOrigen);
            OcultarLabel(labelGuiasCobradasdestino);
            OcultarLabel(labelGuiasCanceladas);
            OcultarLabel(labelnumerodeguiasporcobrar);
            OcultarLabel(labelmontototalporcobrar);
            OcultarIndicadoresEntrega();
        }

        private void LimpiarIndicadoresGuiasPagadas()
        {
            OcultarLabel(labelGuiasOrigen);
            OcultarLabel(labelGuiasCobradasdestino);
        }

        private void radioButtonSinfiltro_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButtonSinfiltro.Checked)
            {
                OcultarIndicadoresResumen();
            }
        }

        private void importarcsvClientesToolStripMenuItem_Click(object sender, EventArgs e)
        {
            using (var dialog = new OpenFileDialog
            {
                Filter = "Archivo CSV|*.csv",
                Title = "Importar clientes desde CSV"
            })
            {
                if (dialog.ShowDialog() != DialogResult.OK)
                {
                    return;
                }

                toolStripProgressBar1.Minimum = 0;
                toolStripProgressBar1.Maximum = 100;
                toolStripProgressBar1.Value = 0;
                toolStripProgressBar1.Visible = true;

                Task.Run(() =>
                {
                    try
                    {
                        dbManager.ImportarClientesDesdeCsv(dialog.FileName, (actual, total) =>
                        {
                            int porcentaje = total == 0 ? 0 : (int)((actual * 100.0) / total);
                            toolStripProgressBar1.GetCurrentParent()?.BeginInvoke((Action)(() =>
                            {
                                toolStripProgressBar1.Value = Math.Max(toolStripProgressBar1.Minimum,
                                    Math.Min(toolStripProgressBar1.Maximum, porcentaje));
                            }));
                        });

                        BeginInvoke((Action)(() =>
                        {
                            MostrarClientesEnDataGridView();
                            MessageBox.Show("Clientes importados correctamente.", "Éxito",
                                MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }));
                    }
                    catch (Exception ex)
                    {
                        BeginInvoke((Action)(() =>
                            MessageBox.Show($"Error al importar Clientes: {ex.Message}", "Error",
                                MessageBoxButtons.OK, MessageBoxIcon.Error)));
                    }
                    finally
                    {
                        BeginInvoke((Action)(() => toolStripProgressBar1.Visible = false));
                    }
                });
            }
        }

        private void exportarcsvClientesToolStripMenuItem_Click(object sender, EventArgs e)
        {
            using (var dialog = new SaveFileDialog
            {
                Filter = "Archivo CSV|*.csv",
                Title = "Exportar clientes a CSV",
                FileName = $"Clientes_{DateTime.Now:yyyyMMdd_HHmm}.csv"
            })
            {
                if (dialog.ShowDialog() != DialogResult.OK)
                {
                    return;
                }

                try
                {
                    dbManager.ExportarClientesACsv(dialog.FileName);
                    MessageBox.Show("Clientes exportados correctamente.", "Éxito",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Error al exportar clientes: {ex.Message}", "Error",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void toolStripButtoiniciarbusqueda_Click(object sender, EventArgs e)
        {
            string criterio = toolStripTextBoxBcliente.Text.Trim();

            if (string.IsNullOrWhiteSpace(criterio))
            {
                MostrarClientesEnDataGridView();
                return;
            }

            DataTable resultado = dbManager.BuscarClientesPorTexto(criterio);

            if (resultado == null || resultado.Rows.Count == 0)
            {
                MessageBox.Show("No se encontraron clientes que coincidan con el criterio indicado.", "Sin resultados",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            MostrarClientesEnDataGridView(resultado);
        }

        private void toolStripProgressBar1_Click(object sender, EventArgs e)
        {
        }

        private void dtpFechaInicio_ValueChanged(object sender, EventArgs e)
        {
            if (dtpFechaFin.Value < dtpFechaInicio.Value)
            {
                dtpFechaFin.Value = dtpFechaInicio.Value;
            }

            CargarGraficoVentas();
        }

        private void dtpFechaFin_ValueChanged(object sender, EventArgs e)
        {
            if (dtpFechaFin.Value < dtpFechaInicio.Value)
            {
                dtpFechaFin.Value = dtpFechaInicio.Value;
            }

            CargarGraficoVentas();
        }

        private void label3_Click(object sender, EventArgs e)
        {
        }

        private void labelmontototalporcobrar_Click(object sender, EventArgs e)
        {
        }

        private void labelnumerodeguiasporcobrar_Click(object sender, EventArgs e)
        {
        }

        private void radioButtonGuiasporcobrar_CheckedChanged_1(object sender, EventArgs e)
        {
            HabilitarComboBoxSucursales();

            if (radioButtonGuiasporcobrar.Checked)
            {
                OcultarIndicadoresResumen();
            }
        }

        private void radioButtonGuiasconCredito_CheckedChanged(object sender, EventArgs e)
        {
            HabilitarComboBoxSucursales();

            if (radioButtonGuiasconcredito.Checked)
            {
                OcultarIndicadoresResumen();
            }
        }

        private void toolStripTextBoxBcliente_TextChanged(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(toolStripTextBoxBcliente.Text))
            {
                MostrarClientesEnDataGridView();
            }
        }

        private void toolStripTextBoxBcliente_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                toolStripButtoiniciarbusqueda_Click(sender, EventArgs.Empty);
                e.SuppressKeyPress = true;
            }
        }



        private void MostrarClientesEnDataGridView()
        {
            MostrarClientesEnDataGridView(dbManager.ObtenerTodosClientes());
        }

        private void MostrarClientesEnDataGridView(DataTable datos)
        {
            dataGridViewPrincipal.DataSource = datos;

            if (dataGridViewPrincipal.Columns.Contains("Activo") &&
                !(dataGridViewPrincipal.Columns["Activo"] is DataGridViewCheckBoxColumn))
            {
                int index = dataGridViewPrincipal.Columns["Activo"].Index;
                dataGridViewPrincipal.Columns.Remove("Activo");
                var checkColumn = new DataGridViewCheckBoxColumn
                {
                    Name = "Activo",
                    HeaderText = "Activo",
                    DataPropertyName = "Activo",
                    TrueValue = true,
                    FalseValue = false
                };
                dataGridViewPrincipal.Columns.Insert(index, checkColumn);
            }

            FormatearEncabezadosDataGridView(dataGridViewPrincipal);
            AplicarEstiloAzulClaroDataGridView(dataGridViewPrincipal);
        }

        private void MostrarUsuariosEnDataGridView()
        {
            MostrarUsuariosEnDataGridView(dbManager.ObtenerTodosUsuarios());
        }

        private void MostrarUsuariosEnDataGridView(DataTable datos)
        {
            dataGridViewPrincipal.DataSource = datos;

            if (dataGridViewPrincipal.Columns.Contains("Activo") &&
                !(dataGridViewPrincipal.Columns["Activo"] is DataGridViewCheckBoxColumn))
            {
                int index = dataGridViewPrincipal.Columns["Activo"].Index;
                dataGridViewPrincipal.Columns.Remove("Activo");
                var checkColumn = new DataGridViewCheckBoxColumn
                {
                    Name = "Activo",
                    HeaderText = "Activo",
                    DataPropertyName = "Activo",
                    TrueValue = true,
                    FalseValue = false
                };
                dataGridViewPrincipal.Columns.Insert(index, checkColumn);
            }

            FormatearEncabezadosDataGridView(dataGridViewPrincipal);
            AplicarEstiloAzulClaroDataGridView(dataGridViewPrincipal);
        }

        private void HabilitarComboBoxSucursales()
        {
            bool puedeHabilitar = radioButtonGuiasPagadas.Checked || radioButtonGuiasporcobrar.Checked ||
                                  radioButtonGuiasconcredito.Checked || radioButtonGuiasCanceladas.Checked ||
                                  radioButtonTodoslosFiltros.Checked || radioButtonPagadasEntregadas.Checked;

            comboBoxSucursales.Enabled = puedeHabilitar;
        }

        private void radioButtonSinfiltro_CheckedChanged_1(object sender, EventArgs e)
        {

        }

        private void radioButtonGuiasCanceladas_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void IniciarTimerUltimaCarga()
        {
            if (timerUltimaCarga == null)
            {
                timerUltimaCarga = new Timer { Interval = 1000 };
                timerUltimaCarga.Tick += TimerUltimaCarga_Tick;
            }

            timerUltimaCarga.Start();
        }

        private void TimerUltimaCarga_Tick(object sender, EventArgs e)
        {
            var estado = LeerEstadoCarga();
            var hoy = DateTime.Today;
            var historialHoy = estado.Historial
                .Where(r => r.FechaHora.Date == hoy)
                .OrderByDescending(r => r.FechaHora)
                .ToList();

            if (!historialHoy.Any())
            {
                toolStripLabel3.Text = "No se ha cargado ningún documento de Excel el día de hoy.";
                return;
            }

            int totalHoy = historialHoy.Sum(r => r.DocumentosCargados);
            DateTime ultima = historialHoy.First().FechaHora;
            TimeSpan transcurrido = DateTime.Now - ultima;

            toolStripLabel3.Text = $"Hoy: {totalHoy} guías. Última carga hace {transcurrido.Hours}h {transcurrido.Minutes}m {transcurrido.Seconds}s";
        }

        private EstadoCargaExcel LeerEstadoCarga()
        {
            if (!File.Exists(EstadoCargaPath))
            {
                return new EstadoCargaExcel();
            }

            try
            {
                return JsonConvert.DeserializeObject<EstadoCargaExcel>(File.ReadAllText(EstadoCargaPath)) ?? new EstadoCargaExcel();
            }
            catch
            {
                return new EstadoCargaExcel();
            }
        }

        private void CargarGraficoVentas()
        {
            if (chartVentasdiarias.ChartAreas.Count == 0)
            {
                return;
            }

            const string currencyFormat = "$#,##0.00";
            DateTime fechaInicio = dtpFechaInicio.Value.Date;
            DateTime fechaFin = dtpFechaFin.Value.Date;

            if (fechaFin < fechaInicio)
            {
                var temp = fechaInicio;
                fechaInicio = fechaFin;
                fechaFin = temp;
            }

            DataTable datos = dbManager.ObtenerSumaTotalPorSucursalConCanceladas(fechaInicio, fechaFin);

            chartVentasdiarias.Series.Clear();
            chartVentasdiarias.Legends.Clear();
            chartVentasdiarias.Titles.Clear();
            chartVentasdiarias.Titles.Add(new Title($"Ventas {fechaInicio:dd/MM/yyyy} - {fechaFin:dd/MM/yyyy}")
            {
                Font = new Font("Segoe UI", 14f, FontStyle.Bold),
                ForeColor = Color.FromArgb(40, 40, 40),
                Alignment = ContentAlignment.MiddleCenter
            });

            var chartArea = chartVentasdiarias.ChartAreas[0];
            chartArea.AxisX.Interval = 1;
            chartArea.AxisX.LabelStyle.Angle = 0;
            chartArea.AxisX.LabelStyle.IsEndLabelVisible = true;
            chartArea.AxisX.LabelStyle.Font = new Font("Segoe UI", 8f, FontStyle.Regular);
            chartArea.AxisY.LabelStyle.Format = currencyFormat;
            chartArea.AxisY.Title = "Monto (MXN)";
            chartArea.AxisY.MajorGrid.LineDashStyle = ChartDashStyle.Dot;

            var seriesValidas = new Series("Ingresos válidos")
            {
                ChartType = SeriesChartType.Column,
                Color = Color.FromArgb(0, 122, 204),
                IsValueShownAsLabel = true,
                LabelFormat = currencyFormat
            };

            var seriesCanceladas = new Series("Canceladas")
            {
                ChartType = SeriesChartType.Column,
                Color = Color.FromArgb(220, 53, 69),
                IsValueShownAsLabel = true,
                LabelFormat = currencyFormat,
                LabelForeColor = Color.FromArgb(220, 53, 69)
            };

            foreach (DataRow row in datos.Rows)
            {
                string sucursal = row["Sucursal"]?.ToString() ?? "Sin dato";
                seriesValidas.Points.AddXY(sucursal, ConvertToDecimal(row["TotalValidas"]));
                seriesCanceladas.Points.AddXY(sucursal, ConvertToDecimal(row["TotalCanceladas"]));
            }

            chartVentasdiarias.Series.Add(seriesValidas);
            chartVentasdiarias.Series.Add(seriesCanceladas);
            chartVentasdiarias.Legends.Add(new Legend());
        }

        private void FormatearEncabezadosDataGridView(DataGridView dgv)
        {
            if (dgv == null || dgv.Columns.Count == 0)
            {
                return;
            }

            var encabezados = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                { "FolioGuia", "Folio Guía" },
                { "FechaElaboracion", "Fecha Elaboración" },
                { "HoraElaboracion", "Hora Elaboración" },
                { "EstatusGuia", "Estatus Guía" },
                { "UbicacionActual", "Ubicación Actual" },
                { "TipoCobro", "Tipo Cobro" },
                { "ZonaOperativaEntrega", "Zona Operativa Entrega" },
                { "TipoEntrega", "Tipo de Entrega" },
                { "FechaEntrega", "Fecha Entrega" },
                { "HoraEntrega", "Hora Entrega" },
                { "FolioInforme", "Folio Informe" },
                { "FolioEmbarque", "Folio Embarque" },
                { "UsuarioDocumento", "Usuario Documento" },
                { "FechaCancelacion", "Fecha Cancelación" },
                { "UsuarioCancelacion", "Usuario Cancelación" },
                { "ValorDeclarado", "Valor Declarado" },
                { "TipoCobroInicial", "Tipo Cobro Inicial" },
                { "FechaUltimaMilla", "Fecha Última Milla" },
                { "MotivoCancelacion", "Motivo Cancelación" },
            };

            foreach (DataGridViewColumn col in dgv.Columns)
            {
                if (encabezados.TryGetValue(col.Name, out string header))
                {
                    col.HeaderText = header;
                    continue;
                }

                string texto = (col.HeaderText ?? col.Name)
                    .Replace("_", " ")
                    .Replace("-", " ");

                col.HeaderText = CultureInfo.CurrentCulture.TextInfo.ToTitleCase(texto.ToLowerInvariant()).Trim();
            }

            if (dgv.Columns.Contains("Subtotal"))
            {
                dgv.Columns["Subtotal"].DefaultCellStyle.Format = "C";
                dgv.Columns["Subtotal"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            }

            if (dgv.Columns.Contains("Total"))
            {
                dgv.Columns["Total"].DefaultCellStyle.Format = "C";
                dgv.Columns["Total"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            }

            if (dgv.Columns.Contains("ValorDeclarado"))
            {
                dgv.Columns["ValorDeclarado"].DefaultCellStyle.Format = "C";
                dgv.Columns["ValorDeclarado"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            }
        }

        private void AplicarEstiloAzulClaroDataGridView(DataGridView dgv)
        {
            if (dgv == null)
            {
                return;
            }

            dgv.EnableHeadersVisualStyles = false;
            dgv.ColumnHeadersDefaultCellStyle.BackColor = Color.LightSteelBlue;
            dgv.ColumnHeadersDefaultCellStyle.ForeColor = Color.Black;
            dgv.ColumnHeadersDefaultCellStyle.Font = new Font(dgv.Font, FontStyle.Bold);

            dgv.DefaultCellStyle.BackColor = Color.AliceBlue;
            dgv.DefaultCellStyle.ForeColor = Color.Black;
            dgv.DefaultCellStyle.SelectionBackColor = Color.LightSalmon;
            dgv.DefaultCellStyle.SelectionForeColor = Color.Black;
            dgv.AlternatingRowsDefaultCellStyle.BackColor = Color.White;
        }







        private void ActualizarResumenPagosPorUbicacion(DateTime fechaInicio, DateTime fechaFin, string sucursal, string destino)
        {
            var resumenOrigen = dbManager.ObtenerResumenPagosOrigen(fechaInicio, fechaFin, sucursal, destino);
            labelGuiasOrigen.Text = $"{resumenOrigen.Cantidad}\n{resumenOrigen.Monto:C}";
            labelGuiasOrigen.Font = new Font(labelGuiasOrigen.Font, FontStyle.Bold);
            labelGuiasOrigen.ForeColor = Color.MediumSeaGreen;
            labelGuiasOrigen.Visible = true;

            var resumenDestino = dbManager.ObtenerResumenPagosDestino(fechaInicio, fechaFin, sucursal, destino);
            labelGuiasCobradasdestino.Text = $"{resumenDestino.Cantidad}\n{resumenDestino.Monto:C}";
            labelGuiasCobradasdestino.Font = new Font(labelGuiasCobradasdestino.Font, FontStyle.Bold);
            labelGuiasCobradasdestino.ForeColor = Color.DarkOrange;
            labelGuiasCobradasdestino.Visible = true;
        }


        private static decimal ConvertToDecimal(object value)
        {
            if (value == null)
            {
                return 0m;
            }

            return decimal.TryParse(Convert.ToString(value), NumberStyles.Any, CultureInfo.InvariantCulture, out decimal result)
                ? result
                : 0m;
        }

        private static bool EsGuiaCancelada(DataRow row)
        {
            if (row == null || row.Table == null || !row.Table.Columns.Contains("EstatusGuia"))
            {
                return false;
            }

            string estatus = Convert.ToString(row["EstatusGuia"])?.Trim();
            return string.Equals(estatus, "CANCELADO", StringComparison.OrdinalIgnoreCase);
        }

        private static IEnumerable<DataRow> FilasNoCanceladas(DataTable dt)
        {
            return dt?.AsEnumerable().Where(row => !EsGuiaCancelada(row)) ?? Enumerable.Empty<DataRow>();
        }

        private static int ContarGuiasNoCanceladas(DataTable dt)
        {
            return FilasNoCanceladas(dt).Count();
        }

        private void truncarGuiasToolStripMenuItem_Click(object sender, EventArgs e)
        {
            const string mensaje = "Esta acción eliminará de forma permanente todos los registros de la tabla guias. ¿Deseas continuar?";
            var confirmacion = MessageBox.Show(mensaje, "Confirmar eliminación masiva", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2);

            if (confirmacion != DialogResult.Yes)
            {
                return;
            }

            System.Windows.Forms.Cursor previousCursor = System.Windows.Forms.Cursor.Current;
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor;

            try
            {
                if (dbManager.TruncarGuias())
                {
                    dataGridViewPrincipal.DataSource = null;
                    MessageBox.Show("Todos los registros de guías fueron eliminados.", "Operación completada", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    ActualizarEstadoCargaExcel();
                    CargarGraficoVentas();
                }
            }
            finally
            {
                System.Windows.Forms.Cursor.Current = previousCursor;
            }
        }

        private void radioButtonTodoslosFiltros_CheckedChanged(object sender, EventArgs e)
        {
            HabilitarComboBoxSucursales();

            if (!radioButtonTodoslosFiltros.Checked)
            {
                return;
            }

            OcultarIndicadoresResumen();
            dataGridViewPrincipal.DataSource = null;
            labelGuiastotales.Text = "0";

            if (filtrosInicializados)
            {
                buttonBuscar.PerformClick();
            }
        }



        private static void MostrarResumenEnLabel(Label label, (int Cantidad, decimal Monto) datos, Color color)
        {
            label.Text = $"{datos.Cantidad}  Monto: {datos.Monto:C}";
            label.Font = new Font(label.Font, FontStyle.Bold);
            label.ForeColor = color;
            label.Visible = true;
        }

        private void ActualizarIndicadoresTodosLosFiltros(DataTable dt)
        {
            if (dt == null || dt.Rows.Count == 0)
            {
                labelnumerodeguiasporcobrar.Text = "0";
                labelnumerodeguiasporcobrar.Font = new Font(labelnumerodeguiasporcobrar.Font, FontStyle.Bold);
                labelnumerodeguiasporcobrar.ForeColor = Color.Black;
                labelnumerodeguiasporcobrar.Visible = true;

                labelmontototalporcobrar.Text = $"0\n{0m:C}";
                labelmontototalporcobrar.Font = new Font(labelmontototalporcobrar.Font, FontStyle.Bold);
                labelmontototalporcobrar.ForeColor = Color.Green;
                labelmontototalporcobrar.Visible = true;

                labelGuiastotales.Text = "0";
                labelGuiastotales.Font = new Font(labelGuiastotales.Font, FontStyle.Bold);
                labelGuiastotales.ForeColor = Color.Navy;

                ActualizarLabelResumen(labelGuiasCanceladas, 0, 0m);
                return;
            }

            var rows = dt.AsEnumerable();
            var filasNoCanceladas = rows.Where(r => !EsGuiaCancelada(r)).ToList();

            var porCobrar = filasNoCanceladas
                .Where(r => string.Equals(Convert.ToString(r["TipoCobro"]), "POR COBRAR", StringComparison.OrdinalIgnoreCase))
                .ToList();

            labelnumerodeguiasporcobrar.Text = porCobrar.Count.ToString();
            labelnumerodeguiasporcobrar.Font = new Font(labelnumerodeguiasporcobrar.Font, FontStyle.Bold);
            labelnumerodeguiasporcobrar.ForeColor = Color.Black;
            labelnumerodeguiasporcobrar.Visible = true;

            decimal montoPorCobrar = porCobrar.Sum(r => ConvertToDecimal(r["Total"]));
            labelmontototalporcobrar.Text = $"{porCobrar.Count}\n{montoPorCobrar:C}";
            labelmontototalporcobrar.Font = new Font(labelmontototalporcobrar.Font, FontStyle.Bold);
            labelmontototalporcobrar.ForeColor = montoPorCobrar < 50000m ? Color.Green : Color.Red;
            labelmontototalporcobrar.Visible = true;

            labelGuiastotales.Text = filasNoCanceladas.Count.ToString();
            labelGuiastotales.Font = new Font(labelGuiastotales.Font, FontStyle.Bold);
            labelGuiastotales.ForeColor = Color.Navy;

            var canceladas = rows
                .Where(r => string.Equals(Convert.ToString(r["EstatusGuia"]), "CANCELADO", StringComparison.OrdinalIgnoreCase))
                .ToList();

            decimal montoCanceladas = canceladas.Sum(r => ConvertToDecimal(r["Total"]));
            ActualizarLabelResumen(labelGuiasCanceladas, canceladas.Count, montoCanceladas);
        }



        private void textBox1clientes_TextChanged(object sender, EventArgs e)
        {

        }

        private void dateTimePicker1clientes_ValueChanged(object sender, EventArgs e)
        {

        }

        private void dateTimePicker2clientes_ValueChanged(object sender, EventArgs e)
        {

        }


        private void button1Buscarstatsclientes_Click(object sender, EventArgs e)
        {
            string cliente = textBox1clientes.Text?.Trim();
            if (string.IsNullOrWhiteSpace(cliente))
            {
                MessageBox.Show("Ingresa el nombre del cliente en mayúsculas.", "Dato requerido", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                textBox1clientes.Focus();
                return;
            }

            DateTime fechaInicio = dateTimePickerInicioclientes.Value.Date;
            DateTime fechaFin = dateTimePickerFinclientes.Value.Date;
            if (fechaFin < fechaInicio)
            {
                MessageBox.Show("La fecha final no puede ser menor que la inicial.", "Rango inválido", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            System.Windows.Forms.Cursor cursorAnterior = System.Windows.Forms.Cursor.Current;
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor;
            try
            {
                var datos = dbManager.ObtenerGuiasPorCliente(cliente, fechaInicio, fechaFin);
                if (datos == null || datos.Rows.Count == 0)
                {
                    dataGridViewPrincipal.DataSource = null;
                    LimpiarIndicadoresClientes();
                    MessageBox.Show("No se encontraron guías para el cliente y rango indicados.", "Sin resultados", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                dataGridViewPrincipal.DataSource = datos;
                FormatearEncabezadosDataGridView(dataGridViewPrincipal);
                AplicarEstiloAzulClaroDataGridView(dataGridViewPrincipal);

                ClienteEstadisticas estadisticas = dbManager.ObtenerEstadisticasCliente(cliente, fechaInicio, fechaFin);
                int totalNoCanceladas = ContarGuiasNoCanceladas(datos);
                ActualizarIndicadoresClientes(estadisticas, totalNoCanceladas);
            }
            catch (Exception ex)
            {
                LimpiarIndicadoresClientes();
                MessageBox.Show($"Error al consultar las guías del cliente: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                System.Windows.Forms.Cursor.Current = cursorAnterior;
            }
        }

        private void labelGuiasxcobrarclientes_Click(object sender, EventArgs e)
        {

        }

        private void labeltotalxcobrarclientes_Click(object sender, EventArgs e)
        {

        }

        private void labelguiastotalesclientes_Click(object sender, EventArgs e)
        {

        }

        private void labelguiaspagadasorigenclientes_Click(object sender, EventArgs e)
        {

        }

        private void labelguiaspagadasdestinoclientes_Click(object sender, EventArgs e)
        {

        }

        private void labelguiascanceladascliente_Click(object sender, EventArgs e)
        {

        }

        private void label20_Click(object sender, EventArgs e)
        {

        }

        private void label25_Click(object sender, EventArgs e)
        {

        }

        private void comboBoxSucursalDestino_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (!filtrosInicializados)
            {
                return;
            }

            bool requiereActualizarResumen = radioButtonGuiasPagadas.Checked || radioButtonTodoslosFiltros.Checked;

            if (requiereActualizarResumen)
            {
                LimpiarIndicadoresGuiasPagadas();
            }

            if (radioButtonGuiasPagadas.Checked || radioButtonGuiasporcobrar.Checked ||
                radioButtonGuiasconcredito.Checked || radioButtonGuiasCanceladas.Checked ||
                radioButtonTodoslosFiltros.Checked)
            {
                dataGridViewPrincipal.DataSource = null;
                labelGuiastotales.Text = "Total de guías: 0";
                buttonBuscar.PerformClick();
            }

        }

        private void checkBoxSucursalDestino_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void LimpiarIndicadoresClientes()
        {
            OcultarLabel(labelGuiasxcobrarclientes);
            OcultarLabel(labeltotalxcobrarclientes);
            OcultarLabel(labelguiaspagadasorigenclientes);
            OcultarLabel(labelguiaspagadasdestinoclientes);
            OcultarLabel(labelguiascanceladascliente);
            OcultarLabel(labelguiastotalesclientes);
            OcultarLabel(labelPaquetesEnviados);
            labelSucursales.Text = string.Empty;
            labelSucursales.Visible = false;
        }

        private void ActualizarIndicadoresClientes(ClienteEstadisticas estadisticas, int totalGuias)
        {
            labelGuiasxcobrarclientes.Text = $": {estadisticas.GuiasPorCobrar}";
            labelGuiasxcobrarclientes.Font = new Font(labelGuiasxcobrarclientes.Font, FontStyle.Bold);
            labelGuiasxcobrarclientes.ForeColor = Color.Black;
            labelGuiasxcobrarclientes.Visible = true;

            labeltotalxcobrarclientes.Text = $": {estadisticas.MontoPorCobrar:C}";
            labeltotalxcobrarclientes.Font = new Font(labeltotalxcobrarclientes.Font, FontStyle.Bold);
            labeltotalxcobrarclientes.ForeColor = estadisticas.MontoPorCobrar < 50000m ? Color.Green : Color.Red;
            labeltotalxcobrarclientes.Visible = true;

            labelguiaspagadasorigenclientes.Text = $": {estadisticas.GuiasPagadasOrigen}   Monto: {estadisticas.MontoPagadasOrigen:C}";
            labelguiaspagadasorigenclientes.Font = new Font(labelguiaspagadasorigenclientes.Font, FontStyle.Bold);
            labelguiaspagadasorigenclientes.ForeColor = Color.MediumSeaGreen;
            labelguiaspagadasorigenclientes.Visible = true;

            labelguiaspagadasdestinoclientes.Text = $": {estadisticas.GuiasPagadasDestino}   Monto: {estadisticas.MontoPagadasDestino:C}";
            labelguiaspagadasdestinoclientes.Font = new Font(labelguiaspagadasdestinoclientes.Font, FontStyle.Bold);
            labelguiaspagadasdestinoclientes.ForeColor = Color.DarkOrange;
            labelguiaspagadasdestinoclientes.Visible = true;

            labelguiascanceladascliente.Text = $": {estadisticas.GuiasCanceladas}   Monto: {estadisticas.MontoCanceladas:C}";
            labelguiascanceladascliente.Font = new Font(labelguiascanceladascliente.Font, FontStyle.Bold);
            labelguiascanceladascliente.ForeColor = Color.Firebrick;
            labelguiascanceladascliente.Visible = true;

            labelguiastotalesclientes.Text = $": {totalGuias}";
            labelguiastotalesclientes.Font = new Font(labelguiastotalesclientes.Font, FontStyle.Bold);
            labelguiastotalesclientes.ForeColor = Color.Navy;
            labelguiastotalesclientes.Visible = true;

            labelPaquetesEnviados.Text = $"Paquetes enviados: {estadisticas.PaquetesEnviados:N0}";
            labelPaquetesEnviados.Font = new Font(labelPaquetesEnviados.Font, FontStyle.Bold);
            labelPaquetesEnviados.ForeColor = Color.SteelBlue;
            labelPaquetesEnviados.Visible = true;

            labelSucursales.Text = estadisticas.Destinos.Count == 0
                 ? "Sin destinos registrados."
                : string.Join(Environment.NewLine, estadisticas.Destinos.Select(d => $"{d.Destino}: {d.TotalCajas:N0} pz"));
            labelSucursales.Font = new Font(labelSucursales.Font, FontStyle.Bold);
            labelSucursales.ForeColor = Color.DimGray;
            labelSucursales.Visible = true;
        }

        private void groupBox1_Enter(object sender, EventArgs e)
        {

        }

        private void radioButtonPagadasEntregadas_CheckedChanged(object sender, EventArgs e)
        {
            HabilitarComboBoxSucursales();

            if (!filtrosInicializados)
            {
                return;
            }

            if (radioButtonPagadasEntregadas.Checked)
            {
                dataGridViewPrincipal.DataSource = null;
                labelGuiastotales.Text = "Total de guías: 0";
                buttonBuscar.PerformClick();
            }
            else
            {
                OcultarIndicadoresEntrega();
            }
        }

        private void importarUsuariosToolStripMenuItem_Click(object sender, EventArgs e)
        {
            using (var dialog = new OpenFileDialog
            {
                Filter = "Archivos Excel|*.xlsx;*.xls",
                Title = "Importar usuarios desde Excel"
            })
            {
                if (dialog.ShowDialog() != DialogResult.OK)
                {
                    return;
                }

                toolStripProgressBar1.Minimum = 0;
                toolStripProgressBar1.Maximum = 100;
                toolStripProgressBar1.Value = 0;
                toolStripProgressBar1.Visible = true;

                Task.Run(() =>
                {
                    try
                    {
                        dbManager.ImportarUsuariosDesdeExcel(dialog.FileName, (actual, total) =>
                        {
                            int porcentaje = total == 0 ? 0 : (int)((actual * 100.0) / total);
                            toolStripProgressBar1.GetCurrentParent()?.BeginInvoke((Action)(() =>
                            {
                                toolStripProgressBar1.Value = Math.Max(toolStripProgressBar1.Minimum,
                                    Math.Min(toolStripProgressBar1.Maximum, porcentaje));
                            }));
                        });

                        BeginInvoke((Action)(() =>
                        {
                            MostrarUsuariosEnDataGridView();
                            MessageBox.Show("Usuarios importados correctamente.", "Éxito",
                                MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }));
                    }
                    catch (Exception ex)
                    {
                        BeginInvoke((Action)(() =>
                            MessageBox.Show($"Error al importar usuarios: {ex.Message}", "Error",
                                MessageBoxButtons.OK, MessageBoxIcon.Error)));
                    }
                    finally
                    {
                        BeginInvoke((Action)(() => toolStripProgressBar1.Visible = false));
                    }
                });
            }
        }

        private void comboBoxSucursalCancelaciones_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (!filtrosInicializados)
            {
                return;
            }

            ActualizarCancelacionesPorSucursal();
        }

        private void ActualizarCancelacionesPorSucursal()
        {
            DateTime fechaInicio = dateTimePickerInicioCanceladas.Value.Date;
            DateTime fechaFin = dateTimePickerFinCanceladas.Value.Date;
            if (fechaFin < fechaInicio)
            {
                var temp = fechaInicio;
                fechaInicio = fechaFin;
                fechaFin = temp;
            }

            string sucursal = comboBoxSucursalCancelaciones.SelectedItem?.ToString() ?? "TODAS";
            sucursal = sucursal.Trim();

            DataTable canceladas = dbManager.ObtenerGuiasCanceladasPorSucursal(fechaInicio, fechaFin, sucursal);
            CargarGridConSeleccionPersistente(canceladas);

            DataTable resumen = dbManager.ObtenerResumenCancelacionesPorUsuario(fechaInicio, fechaFin, sucursal);
            int totalGuias = dbManager.ContarGuiasTotales(fechaInicio, fechaFin, sucursal);

            if (resumen == null || resumen.Rows.Count == 0 || totalGuias == 0)
            {
                labelUsuarioCancelaccion.Text = "Sin datos.";
                labelIndicadorCancelaciones.Text = "Sin datos.";
                return;
            }

            var rtf = new StringBuilder();
            rtf.Append(@"{\rtf1\ansi{\colortbl ;\red0\green0\blue0;\red0\green128\blue0;\red200\green0\blue0;}");
            foreach (DataRow r in resumen.Rows)
            {
                string nombre = ObtenerPrimerNombre(r["UsuarioDocumento"]?.ToString());
                int gt = Convert.ToInt32(r["TotalGuias"]);
                int gc = Convert.ToInt32(r["Canceladas"]);

                rtf.Append(@"\cf1 ").Append(nombre)
                   .Append(@": \cf2 GT: ").Append(gt)
                   .Append(@"  \cf3 GC: ").Append(gc)
                   .Append(@"\line");
            }
            rtf.Append("}");
            richTextBoxUsuarioCancelaccion.Rtf = rtf.ToString();

            labelIndicadorCancelaciones.ForeColor = Color.Black;
            labelIndicadorCancelaciones.Text = string.Join(Environment.NewLine,
                resumen.AsEnumerable().Select(r =>
                {
                    string nombre = ObtenerPrimerNombre(r["UsuarioDocumento"]?.ToString());
                    int gc = Convert.ToInt32(r["Canceladas"]);
                    double porcentaje = totalGuias == 0 ? 0 : (gc * 100.0 / totalGuias);
                    return $"{nombre}: {porcentaje:0.#}%";
                }));

            var rtfPct = new StringBuilder();
            rtfPct.Append(@"{\rtf1\ansi{\colortbl ;\red0\green0\blue0;\red0\green0\blue200;/* Azul */}");
            foreach (DataRow r in resumen.Rows)
            {
                string nombre = ObtenerPrimerNombre(r["UsuarioDocumento"]?.ToString());
                int gc = Convert.ToInt32(r["Canceladas"]);
                double porcentaje = totalGuias == 0 ? 0 : (gc * 100.0 / totalGuias);

                rtfPct.Append(@"\cf1 ").Append(nombre)
                      .Append(@": \cf2 ").AppendFormat("{0:0.#}%", porcentaje)
                      .Append(@"\line");
            }
            rtfPct.Append("}");
            richTextBoxIndicadorCancelaciones.Rtf = rtfPct.ToString();
        }

        private static string ObtenerPrimerNombre(string nombreCompleto)
        {
            if (string.IsNullOrWhiteSpace(nombreCompleto))
            {
                return "SIN";
            }

            var partes = nombreCompleto.Trim().Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
            return partes.Length > 0 ? partes[0] : nombreCompleto.Trim();
        }

        private void dateTimePickerInicioCanceladas_ValueChanged(object sender, EventArgs e)
        {
            if (!filtrosInicializados)
            {
                return;
            }

            ActualizarCancelacionesPorSucursal();
        }

        private void dateTimePickerFinCanceladas_ValueChanged(object sender, EventArgs e)
        {
            if (!filtrosInicializados)
            {
                return;
            }

            ActualizarCancelacionesPorSucursal();
        }

        private void buttonCanceladoPequeño_Click(object sender, EventArgs e)
        {
            if (!filtrosInicializados)
            {
                return;
            }
            ActualizarCancelacionesPorSucursal();
        }

        private void buttonbuscarestatus_Click(object sender, EventArgs e)
        {
            DateTime inicio = dateTimePickerInicioSeguimiento.Value.Date;
            DateTime fin = dateTimePickerFinSeguimiento.Value.Date;

            if (fin < inicio)
            {
                MessageBox.Show("La fecha final no puede ser menor que la inicial.", "Rango inválido", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            var estatus = ObtenerEstatusSeleccionados();
            if (estatus.Count == 0 || checkBoxTodoslosFiltros.Checked)
            {
                estatus.Clear(); // sin filtro de estatus
            }

            var dt = dbManager.ObtenerGuiasPorEstatus(inicio, fin, estatus);

            // Filtro por Origen
            string origenFiltro = ValorComboOrigenesSeguimiento();
            dt = FiltrarPorOrigenSeguimiento(dt, origenFiltro);

            // NUEVO: filtro por Destino
            string destinoFiltro = ValorComboDestinoSeguimiento();
            dt = FiltrarPorDestinoSeguimiento(dt, destinoFiltro);

            CargarGridConSeleccionPersistente(dt);
        }


        private List<string> ObtenerEstatusSeleccionados()
        {
            var list = new List<string>();
            if (checkBoxDocumentado.Checked) list.Add("DOCUMENTADO");
            if (checkBoxPendiente.Checked) list.Add("PENDIENTE");
            if (checkBoxEnRuta.Checked) list.Add("EN RUTA");
            if (checkBoxUltimaMilla.Checked) list.Add("ULTIMA MILLA");
            if (checkBoxCompletado.Checked) list.Add("COMPLETADO");
            return list;
        }

        private string ObtenerEstatusFila(DataGridViewRow row)
        {
            if (row == null) return string.Empty;
            var cols = row.DataGridView?.Columns;
            if (cols == null) return string.Empty;

            string valor(string colName) =>
                cols.Contains(colName) ? Convert.ToString(row.Cells[colName]?.Value ?? string.Empty).Trim() : null;

            return valor("EstatusGuia")
                   ?? valor("Estatus Guía")
                   ?? string.Empty;
        }

        private void ColorearFilasEstatus(DataGridView dgv)
        {
            if (dgv?.Rows == null) return;

            foreach (DataGridViewRow row in dgv.Rows)
            {
                if (row.IsNewRow) continue;
                string estatus = ObtenerEstatusFila(row);
                if (coloresEstatus.TryGetValue(estatus.ToUpperInvariant(), out Color color))
                {
                    row.DefaultCellStyle.BackColor = color;
                    row.DefaultCellStyle.SelectionBackColor = ControlPaint.Dark(color);
                    row.DefaultCellStyle.SelectionForeColor = Color.Black;
                }
            }
        }

        private void dataGridViewPrincipal_DataBindingComplete(object sender, DataGridViewBindingCompleteEventArgs e)
        {
            SincronizarModoEditable();
            ColorearFilasEstatus(dataGridViewPrincipal);
            ColorearFacturaConValor(dataGridViewPrincipal);
            ColorearCeldaTipoCobro(dataGridViewPrincipal);
            ReAplicarSeleccionPersistenteEnGrid();
        }



        private void checkBoxEditable_CheckedChanged(object sender, EventArgs e)
        {
            SincronizarModoEditable();
            if (!checkBoxEditable.Checked)
            {
                LimpiarSeleccionPersistente();
            }
        }

        private void SincronizarModoEditable()
        {
            bool editable = checkBoxEditable.Checked;

            if (editable)
            {
                if (!dataGridViewPrincipal.Columns.Contains(ColumnaSeleccionExportar))
                {
                    var col = new DataGridViewCheckBoxColumn
                    {
                        Name = ColumnaSeleccionExportar,
                        HeaderText = "Exportar",
                        Width = 60,
                        TrueValue = true,
                        FalseValue = false
                    };
                    dataGridViewPrincipal.Columns.Insert(0, col);
                }
            }
            else
            {
                if (dataGridViewPrincipal.Columns.Contains(ColumnaSeleccionExportar))
                {
                    dataGridViewPrincipal.Columns.Remove(ColumnaSeleccionExportar);
                }
            }

            if (dataGridViewPrincipal.Columns.Contains("Observaciones"))
            {
                dataGridViewPrincipal.Columns["Observaciones"].ReadOnly = !editable;
            }
        }

        private void dataGridViewPrincipal_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter && checkBoxEditable.Checked)
            {
                if (dataGridViewPrincipal.CurrentCell != null &&
                    dataGridViewPrincipal.Columns[dataGridViewPrincipal.CurrentCell.ColumnIndex].Name.Equals("Observaciones", StringComparison.OrdinalIgnoreCase))
                {
                    e.Handled = true;
                    e.SuppressKeyPress = true;
                    dataGridViewPrincipal.EndEdit();
                    GuardarObservacionesFila(dataGridViewPrincipal.CurrentRow);
                }
            }
        }

        private void dataGridViewPrincipal_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            if (!checkBoxEditable.Checked)
            {
                return;
            }

            if (e.RowIndex >= 0 &&
                dataGridViewPrincipal.Columns[e.ColumnIndex].Name.Equals("Observaciones", StringComparison.OrdinalIgnoreCase))
            {
                GuardarObservacionesFila(dataGridViewPrincipal.Rows[e.RowIndex]);
            }
        }

        private void GuardarObservacionesFila(DataGridViewRow row)
        {
            if (row == null || row.IsNewRow) return;

            string folio = ObtenerValorCelda(row, "FolioGuia") ?? ObtenerValorCelda(row, "folio_guia");
            string obs = Convert.ToString(row.Cells["Observaciones"]?.Value ?? string.Empty);

            if (string.IsNullOrWhiteSpace(folio))
            {
                MessageBox.Show("No se encontró el FolioGuia para guardar observaciones.", "Dato faltante", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            bool ok = dbManager.ActualizarObservaciones(folio, obs);
            if (!ok)
            {
                MessageBox.Show("No se pudo guardar Observaciones.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private static string ObtenerValorCelda(DataGridViewRow row, string columnName)
        {
            return row.DataGridView.Columns.Contains(columnName)
                ? Convert.ToString(row.Cells[columnName]?.Value)
                : null;
        }

        private void checkBoxExportarTodo_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBoxExportarTodo.Checked && checkBoxExportarSeleccion.Checked)
            {
                MessageBox.Show("Solo puede estar activo uno: Exportar Todo o Exportar Selección.", "Selección inválida", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                checkBoxExportarSeleccion.Checked = false;
            }
        }

        private void checkBoxExportarSeleccion_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBoxExportarSeleccion.Checked && checkBoxExportarTodo.Checked)
            {
                MessageBox.Show("Solo puede estar activo uno: Exportar Todo o Exportar Selección.", "Selección inválida", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                checkBoxExportarTodo.Checked = false;
            }
        }

        private void buttonExportarStatus_Click(object sender, EventArgs e)
        {
            if (dataGridViewPrincipal.DataSource == null || dataGridViewPrincipal.Rows.Count == 0)
            {
                MessageBox.Show("No hay datos para exportar.", "Sin datos", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            if (checkBoxExportarTodo.Checked && checkBoxExportarSeleccion.Checked)
            {
                MessageBox.Show("Seleccione solo uno: Exportar Todo o Exportar Selección.", "Selección inválida", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            var filas = new List<DataGridViewRow>();
            foreach (DataGridViewRow r in dataGridViewPrincipal.Rows)
            {
                if (r.IsNewRow) continue;

                if (checkBoxExportarSeleccion.Checked)
                {
                    if (EstaSeleccionada(r))
                    {
                        filas.Add(r);
                    }
                }
                else
                {
                    filas.Add(r); // exportar todo
                }
            }

            if (checkBoxExportarSeleccion.Checked && filas.Count == 0)
            {
                MessageBox.Show("No hay filas seleccionadas para exportar.", "Sin selección", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            using (var dialog = new SaveFileDialog
            {
                Filter = "Archivo Excel|*.xlsx",
                Title = "Exportar estatus de guías",
                FileName = $"Guias_Estatus_{DateTime.Now:yyyyMMdd_HHmm}.xlsx"
            })
            {
                if (dialog.ShowDialog() != DialogResult.OK)
                {
                    return;
                }

                DataTable dtExport = ConstruirDataTableDesdeDgv(dataGridViewPrincipal, filas);
                ExportarDataTableAExcel(dtExport, dialog.FileName);
                if (MessageBox.Show("Exportación completada. ¿Deseas abrir el archivo?", "Éxito", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                {
                    System.Diagnostics.Process.Start(dialog.FileName);
                }
            }
        }


        private DataTable ConstruirDataTableDesdeDgv(DataGridView dgv, IEnumerable<DataGridViewRow> filas)
        {
            var dt = new DataTable();

            // Columnas visibles en orden de DisplayIndex
            var columnas = dgv.Columns.Cast<DataGridViewColumn>()
                                      .Where(c => c.Visible && c.Name != ColumnaSeleccionExportar)
                                      .OrderBy(c => c.DisplayIndex)
                                      .ToList();

            foreach (var col in columnas)
            {
                dt.Columns.Add(col.Name, typeof(string));
            }

            foreach (var row in filas)
            {
                var newRow = dt.NewRow();
                foreach (var col in columnas)
                {
                    newRow[col.Name] = Convert.ToString(row.Cells[col.Name]?.Value);
                }
                dt.Rows.Add(newRow);
            }

            return dt;
        }

        private void ExportarDataTableAExcel(DataTable dt, string filePath)
        {
            using (var wb = new XLWorkbook())
            {
                var wsDatos = wb.Worksheets.Add("Vista");
                var table = wsDatos.Cell(1, 1).InsertTable(dt, "VistaActual");
                table.Theme = XLTableTheme.None;
                table.ShowRowStripes = false;
                table.ShowColumnStripes = false;

                string colTipoCobro = dt.Columns.Cast<DataColumn>()
                    .Select(c => c.ColumnName)
                    .FirstOrDefault(c => c.Equals("TipoCobro", StringComparison.OrdinalIgnoreCase) || c.Equals("Tipo Cobro", StringComparison.OrdinalIgnoreCase));

                string colEstatus = dt.Columns.Cast<DataColumn>()
                    .Select(c => c.ColumnName)
                    .FirstOrDefault(c => c.Equals("EstatusGuia", StringComparison.OrdinalIgnoreCase) ||
                                         c.Equals("Estatus Guía", StringComparison.OrdinalIgnoreCase) ||
                                         c.Equals("Estatus Guia", StringComparison.OrdinalIgnoreCase));

                string[] columnasMoneda = { "Subtotal", "Total", "Valor", "TotalGuias", "MontoPorCobrar", "MontoPagadasOrigen", "MontoPagadasDestino", "MontoCanceladas" };
                foreach (var col in columnasMoneda)
                {
                    if (dt.Columns.Contains(col))
                    {
                        int colNumber = dt.Columns[col].Ordinal + 1;
                        wsDatos.Column(colNumber).Style.NumberFormat.Format = "$#,##0.00";
                        wsDatos.Column(colNumber).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                        table.Column(colNumber).Cells().Style.NumberFormat.Format = "$#,##0.00";
                        table.Column(colNumber).Cells().Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                    }
                }

                AplicarColoresFilasExcel(wsDatos, dt, startRow: 2, startCol: 1);

                wsDatos.Columns().AdjustToContents();
                wb.SaveAs(filePath);
            }
        }

        private void ExportarVistaActualConFormato()
        {
            DateTime fechaInicio = dtpFechaInicio.Value.Date;
            DateTime fechaFin = dtpFechaFin.Value.Date;
            if (fechaFin < fechaInicio) { var tmp = fechaInicio; fechaInicio = fechaFin; fechaFin = tmp; }

            string sucursal = comboBoxSucursales.SelectedItem?.ToString() ?? "TODAS";
            string destino = comboBoxSucursalDestino.SelectedItem?.ToString() ?? "TODAS";

            var rangos = ObtenerRangosMensuales(fechaInicio, fechaFin).ToList();
            if (rangos.Count == 0)
            {
                MessageBox.Show("Rango de fechas vacío.", "Sin datos", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            using (var dialog = new SaveFileDialog
            {
                Filter = "Archivo Excel|*.xlsx",
                Title = "Exportar vista actual a Excel",
                FileName = $"Reportes_OMAJA_{DateTime.Now:yyyyMMdd_HHmm}.xlsx"
            })
            {
                if (dialog.ShowDialog() != DialogResult.OK)
                {
                    return;
                }

                if (rangos.Count == 1)
                {
                    if (ExportarReporteExcel(dialog.FileName))
                    {
                        MostrarMensajeExportacionCompletada(dialog.FileName);
                    }

                    return;
                }

                var resumenMeses = new List<(string Mes, int GuiasTotales, int GuiasPorCobrar, decimal MontoPorCobrar, int GuiasCanceladas, decimal MontoCanceladas)>();
                var acumuladoIndicadores = new Dictionary<string, (int Cant, decimal Monto)>(StringComparer.OrdinalIgnoreCase);

                using (var wb = new XLWorkbook())
                {
                    foreach (var rango in rangos)
                    {
                        var datosMes = ObtenerDatosFiltradosParaExport(rango.Inicio, rango.Fin);
                        if (datosMes == null) continue;

                        var ws = wb.Worksheets.Add(rango.Nombre);
                        string tituloMes = $"Reporte Mensual OMAJA - {rango.Nombre}";
                        InsertarEncabezadoMes(ws, tituloMes, sucursal, destino);

                        var inds = ConstruirIndicadoresResumen(rango.Inicio, rango.Fin, sucursal, destino, datosMes);
                        int nextRow = EscribirIndicadoresTabla(ws, 5, inds);

                        var indsGraf = FiltrarIndicadoresParaGrafico(inds);
                        int chartRow = nextRow + 1;
                        InsertarGraficoIndicadores(ws, chartRow, 1, indsGraf);

                        int dataStartRow = chartRow + 22;
                        var tablaMes = ws.Cell(dataStartRow, 1).InsertTable(datosMes, "Guias");
                        tablaMes.Theme = XLTableTheme.None;
                        tablaMes.ShowRowStripes = false;
                        tablaMes.ShowColumnStripes = false;
                        AplicarColoresFilasExcel(ws, datosMes, startRow: dataStartRow + 1, startCol: 1);
                        ws.Columns().AdjustToContents();

                        var est = CalcularEstadisticasMes(datosMes);
                        resumenMeses.Add((rango.Nombre, est.GuiasTotales, est.GuiasPorCobrar, est.MontoPorCobrar, est.GuiasCanceladas, est.MontoCanceladas));

                        foreach (var ind in inds)
                        {
                            if (acumuladoIndicadores.TryGetValue(ind.Titulo, out var acc))
                            {
                                acumuladoIndicadores[ind.Titulo] = (acc.Cant + ind.Cantidad, acc.Monto + ind.Monto);
                            }
                            else
                            {
                                acumuladoIndicadores[ind.Titulo] = (ind.Cantidad, ind.Monto);
                            }
                        }
                    }

                    var wsTot = wb.Worksheets.Add("Totales");
                    wsTot.Cell("A1").Value = "Mes";
                    wsTot.Cell("B1").Value = "Guías Totales";
                    wsTot.Cell("C1").Value = "Guías por Cobrar";
                    wsTot.Cell("D1").Value = "Monto por Cobrar";
                    wsTot.Cell("E1").Value = "Guías Canceladas";
                    wsTot.Cell("F1").Value = "Monto Canceladas";
                    wsTot.Range("A1:F1").Style.Font.Bold = true;

                    int row = 2;
                    foreach (var r in resumenMeses)
                    {
                        wsTot.Cell(row, 1).Value = r.Mes;
                        wsTot.Cell(row, 2).Value = r.GuiasTotales;
                        wsTot.Cell(row, 3).Value = r.GuiasPorCobrar;
                        wsTot.Cell(row, 4).Value = r.MontoPorCobrar;
                        wsTot.Cell(row, 5).Value = r.GuiasCanceladas;
                        wsTot.Cell(row, 6).Value = r.MontoCanceladas;
                        row++;
                    }

                    wsTot.Cell(row, 1).Value = "TOTAL";
                    wsTot.Cell(row, 1).Style.Font.Bold = true;
                    wsTot.Cell(row, 2).FormulaA1 = $"SUM(B2:B{row - 1})";
                    wsTot.Cell(row, 3).FormulaA1 = $"SUM(C2:C{row - 1})";
                    wsTot.Cell(row, 4).FormulaA1 = $"SUM(D2:D{row - 1})";
                    wsTot.Cell(row, 5).FormulaA1 = $"SUM(E2:E{row - 1})";
                    wsTot.Cell(row, 6).FormulaA1 = $"SUM(F2:F{row - 1})";
                    wsTot.Range(row, 1, row, 6).Style.Font.Bold = true;
                    wsTot.Columns(4, 6).Style.NumberFormat.Format = "$#,##0.00";

                    int startInd = row + 2;
                    wsTot.Cell(startInd, 1).Value = "Indicador";
                    wsTot.Cell(startInd, 2).Value = "Cantidad";
                    wsTot.Cell(startInd, 3).Value = "Monto";
                    wsTot.Range(startInd, 1, startInd, 3).Style.Font.Bold = true;

                    int ri = startInd + 1;
                    foreach (var kv in acumuladoIndicadores)
                    {
                        wsTot.Cell(ri, 1).Value = kv.Key;
                        wsTot.Cell(ri, 2).Value = kv.Value.Cant;
                        wsTot.Cell(ri, 3).Value = kv.Value.Monto;
                        ri++;
                    }

                    wsTot.Column(3).Style.NumberFormat.Format = "$#,##0.00";
                    wsTot.Columns().AdjustToContents();

                    wb.SaveAs(dialog.FileName);
                }

                MostrarMensajeExportacionCompletada(dialog.FileName);
            }
        }


        private string ObtenerFolio(DataRow row)
        {
            if (row == null || row.Table == null) return null;
            var cols = row.Table.Columns;
            if (cols.Contains("FolioGuia")) return Convert.ToString(row["FolioGuia"]);
            if (cols.Contains("folio_guia")) return Convert.ToString(row["folio_guia"]);
            return null;
        }

        private string ObtenerFolio(DataGridViewRow row)
        {
            if (row == null || row.DataGridView == null) return null;
            string valor(string colName) =>
                row.DataGridView.Columns.Contains(colName) ? Convert.ToString(row.Cells[colName]?.Value) : null;
            return valor("FolioGuia") ?? valor("folio_guia");
        }


        private DataRow ClonarDataRow(DataRow origen)
        {
            if (origen == null) return null;
            var tabla = origen.Table?.Clone() ?? new DataTable();
            var copia = tabla.NewRow();
            copia.ItemArray = (object[])origen.ItemArray.Clone();
            tabla.Rows.Add(copia);
            return copia;
        }

        private DataTable UnirConSeleccionPersistente(DataTable dt)
        {
            var baseTable = dt != null ? dt.Copy() : null;
            if (baseTable == null)
            {
                var fila = seleccionPersistente.Values.FirstOrDefault();
                baseTable = fila?.Table.Clone() ?? new DataTable();
            }

            var existentes = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (DataRow r in baseTable.Rows)
            {
                var k = ObtenerFolio(r);
                if (!string.IsNullOrWhiteSpace(k))
                    existentes.Add(k);
            }

            foreach (var par in seleccionPersistente)
            {
                if (existentes.Contains(par.Key)) continue;
                var nuevo = baseTable.NewRow();
                foreach (DataColumn c in baseTable.Columns)
                {
                    if (par.Value.Table.Columns.Contains(c.ColumnName))
                    {
                        nuevo[c.ColumnName] = par.Value[c.ColumnName];
                    }
                }
                baseTable.Rows.Add(nuevo);
            }

            return baseTable;
        }

        private void ReAplicarSeleccionPersistenteEnGrid()
        {
            if (!dataGridViewPrincipal.Columns.Contains(ColumnaSeleccionExportar))
            {
                return;
            }

            foreach (DataGridViewRow row in dataGridViewPrincipal.Rows)
            {
                if (row.IsNewRow) continue;
                string folio = ObtenerFolio(row);
                if (!string.IsNullOrWhiteSpace(folio) && seleccionPersistente.ContainsKey(folio))
                {
                    row.Cells[ColumnaSeleccionExportar].Value = true;
                }
            }
        }

        private void ActualizarSeleccionPersistente(DataGridViewRow row)
        {
            string folio = ObtenerFolio(row);
            if (string.IsNullOrWhiteSpace(folio)) return;

            bool seleccionado = false;
            if (row.DataGridView.Columns.Contains(ColumnaSeleccionExportar))
            {
                bool.TryParse(Convert.ToString(row.Cells[ColumnaSeleccionExportar].Value), out seleccionado);
            }

            if (seleccionado)
            {
                var drv = row.DataBoundItem as DataRowView;
                var dataRow = drv?.Row;
                var copia = ClonarDataRow(dataRow);
                if (copia != null)
                {
                    seleccionPersistente[folio] = copia;
                }
            }
            else
            {
                seleccionPersistente.Remove(folio);
            }
        }

        private void LimpiarSeleccionPersistente()
        {
            seleccionPersistente.Clear();
        }

        private bool EstaSeleccionada(DataGridViewRow row)
        {
            if (row == null || row.IsNewRow) return false;
            bool val = false;
            if (row.DataGridView.Columns.Contains(ColumnaSeleccionExportar))
            {
                bool.TryParse(Convert.ToString(row.Cells[ColumnaSeleccionExportar].Value), out val);
            }

            if (val) return true;

            string folio = ObtenerFolio(row);
            return !string.IsNullOrWhiteSpace(folio) && seleccionPersistente.ContainsKey(folio);
        }

        private void CargarGridConSeleccionPersistente(DataTable dt, bool aplicarColoresEstatus = true)
        {
            var data = UnirConSeleccionPersistente(dt);
            dataGridViewPrincipal.DataSource = data;
            FormatearEncabezadosDataGridView(dataGridViewPrincipal);
            AplicarEstiloAzulClaroDataGridView(dataGridViewPrincipal);
            SincronizarModoEditable();
            ReAplicarSeleccionPersistenteEnGrid();
            if (aplicarColoresEstatus)
            {
                ColorearFilasEstatus(dataGridViewPrincipal);
            }
        }

        private void dataGridViewPrincipal_CurrentCellDirtyStateChanged(object sender, EventArgs e)
        {
            if (dataGridViewPrincipal.IsCurrentCellDirty)
            {
                dataGridViewPrincipal.CommitEdit(DataGridViewDataErrorContexts.Commit);
            }
        }

        private void dataGridViewPrincipal_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0) return;
            if (!dataGridViewPrincipal.Columns.Contains(ColumnaSeleccionExportar)) return;
            if (dataGridViewPrincipal.Columns[e.ColumnIndex].Name != ColumnaSeleccionExportar) return;
            ActualizarSeleccionPersistente(dataGridViewPrincipal.Rows[e.RowIndex]);
        }


        private void buttonExportarStatus_Click_1(object sender, EventArgs e)
        {

        }

        private void toolStripButtonBuscarFactura_Click(object sender, EventArgs e)
        {
            string factura = toolStripTextBoxBuscarFactura.Text.Trim();

            if (string.IsNullOrWhiteSpace(factura))
            {
                MessageBox.Show("Ingresa un número o clave de factura para buscar.", "Dato requerido",
                MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            string error;
            DataTable resultados = dbManager.BuscarGuiasPorFactura(factura, out error);

            if (!string.IsNullOrEmpty(error))
            {
                MessageBox.Show(error, "Error de búsqueda", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (resultados != null && resultados.Rows.Count > 0)
            {
                dataGridViewPrincipal.DataSource = resultados;
                AplicarEstiloAzulClaroDataGridView(dataGridViewPrincipal);
                FormatearEncabezadosDataGridView(dataGridViewPrincipal);
            }
            else
            {
                dataGridViewPrincipal.DataSource = null;
                MessageBox.Show("No se encontraron guías con la factura indicada.", "Sin resultados",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void toolStripTextBoxBuscarFactura_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                toolStripButtonBuscarFactura_Click(sender, EventArgs.Empty);
                e.SuppressKeyPress = true;
            }
        }

        private void toolStriptxtBusqueda_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                toolStripBusqueda_Click(sender, EventArgs.Empty);
                e.SuppressKeyPress = true;
            }
        }

        private void toolStripTextBoxBuscarFactura_Click(object sender, EventArgs e)
        {

        }

        private void comboBoxOrigenesSeguimiento_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void comboBoxDestinoSeguimient_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        // NUEVO: alias para comboBoxOrigenesSeguimiento
        private class ComboOpcionSeguimiento
        {
            public string Texto { get; set; }
            public string Valor { get; set; }
            public override string ToString() => Texto;
        }

        private readonly List<ComboOpcionSeguimiento> sucursalesAliasSeguimiento =
            new List<ComboOpcionSeguimiento>
            {
        new ComboOpcionSeguimiento { Texto = "TODAS",    Valor = "TODAS" },
        new ComboOpcionSeguimiento { Texto = "AGS",      Valor = "AGUASCALIENTES, AGUASCALIENTES" },
        new ComboOpcionSeguimiento { Texto = "CHICO",    Valor = "CHICONCUAC, ESTADO DE MEXICO" },
        new ComboOpcionSeguimiento { Texto = "FER",      Valor = "CIUDAD DE MEXICO FER" },
        new ComboOpcionSeguimiento { Texto = "GDL",      Valor = "GUADALAJARA, JALISCO" },
        new ComboOpcionSeguimiento { Texto = "LEON",     Valor = "LEON, GUANAJUATO" },
        new ComboOpcionSeguimiento { Texto = "MIX",      Valor = "CIUDAD DE MEXICO MIXC" },
        new ComboOpcionSeguimiento { Texto = "SFCO",     Valor = "SAN FRANCISCO DEL RINCON, GUANAJUATO" },
        new ComboOpcionSeguimiento { Texto = "TABLA H",  Valor = "TABLA HONDA, ESTADO DE MEXICO" },
        new ComboOpcionSeguimiento { Texto = "TEXTI",    Valor = "TEXTICUITZEO" },
        new ComboOpcionSeguimiento { Texto = "URIANGATO",Valor = "URIANGATO, GUANAJUATO" },
        new ComboOpcionSeguimiento { Texto = "VILLA",    Valor = "VILLA HIDALGO, JALISCO" },
        new ComboOpcionSeguimiento { Texto = "ZAPO",     Valor = "ZAPOTLANEJO, JALISCO"},
            };


        // NUEVO: inicialización y helper del combo de seguimiento
        private void ConfigurarComboOrigenesSeguimiento()
        {
            comboBoxOrigenesSeguimiento.DisplayMember = "Texto";
            comboBoxOrigenesSeguimiento.ValueMember = "Valor";
            comboBoxOrigenesSeguimiento.DataSource = sucursalesAliasSeguimiento.ToList();
        }

        private string ValorComboOrigenesSeguimiento()
        {
            var val = comboBoxOrigenesSeguimiento.SelectedValue as string
                      ?? comboBoxOrigenesSeguimiento.SelectedItem?.ToString();
            return string.IsNullOrWhiteSpace(val) ? "TODAS" : val;
        }

        // NUEVO: filtrar DataTable por Origen usando el valor real
        private DataTable FiltrarPorOrigenSeguimiento(DataTable origen, string origenFiltro)
        {
            if (origen == null || origen.Rows.Count == 0 || string.IsNullOrWhiteSpace(origenFiltro))
                return origen;

            if (origenFiltro.Equals("TODAS", StringComparison.OrdinalIgnoreCase))
                return origen;

            if (!origen.Columns.Contains("Origen"))
                return origen;

            string filtroNorm = NormalizarSucursalNombre(origenFiltro);

            var filas = origen.AsEnumerable()
                .Where(r => NormalizarSucursalNombre(Convert.ToString(r["Origen"])) == filtroNorm)
                .ToList();

            return filas.Any() ? filas.CopyToDataTable() : origen.Clone();
        }

        // NUEVO: inicialización y helpers del combo de destinos
        private void ConfigurarComboDestinosSeguimiento()
        {
            comboBoxDestinoSeguimient.DisplayMember = "Texto";
            comboBoxDestinoSeguimient.ValueMember = "Valor";
            comboBoxDestinoSeguimient.DataSource = sucursalesAliasSeguimiento.ToList();
        }

        private string ValorComboDestinoSeguimiento()
        {
            var val = comboBoxDestinoSeguimient.SelectedValue as string
                      ?? comboBoxDestinoSeguimient.SelectedItem?.ToString();
            return string.IsNullOrWhiteSpace(val) ? "TODAS" : val;
        }

        // NUEVO: filtrar DataTable por Destino usando el valor real normalizado
        private DataTable FiltrarPorDestinoSeguimiento(DataTable origen, string destinoFiltro)
        {
            if (origen == null || origen.Rows.Count == 0 || string.IsNullOrWhiteSpace(destinoFiltro))
                return origen;

            if (destinoFiltro.Equals("TODAS", StringComparison.OrdinalIgnoreCase))
                return origen;

            if (!origen.Columns.Contains("Destino"))
                return origen;

            string filtroNorm = NormalizarSucursalNombre(destinoFiltro);

            var filas = origen.AsEnumerable()
                .Where(r => NormalizarSucursalNombre(Convert.ToString(r["Destino"])) == filtroNorm)
                .ToList();

            return filas.Any() ? filas.CopyToDataTable() : origen.Clone();
        }

        private void dataGridViewSeguimientos_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        // Añade esta clase y lista (cerca de las definiciones de alias)
        private class SucursalSeguimientoInfo
        {
            public string Nombre { get; set; }
            public string PrefijoFolio { get; set; }
            public string Alias { get; set; } // para referencia, no se usa en filtro
        }

        private readonly List<SucursalSeguimientoInfo> sucursalesSeguimientoResumen = new List<SucursalSeguimientoInfo>
        {
            new SucursalSeguimientoInfo { Nombre = "AGUASCALIENTES",            PrefijoFolio = "FG-AGS",   Alias = "AGS" },
            new SucursalSeguimientoInfo { Nombre = "CHICONCUAC",                PrefijoFolio = "FG-CHICO", Alias = "CHICO" },
            new SucursalSeguimientoInfo { Nombre = "FERRETERIA",                PrefijoFolio = "FG-FER",   Alias = "FER" },
            new SucursalSeguimientoInfo { Nombre = "GUADALAJARA",               PrefijoFolio = "FG-GDL",   Alias = "GDL" },
            new SucursalSeguimientoInfo { Nombre = "LEON",                      PrefijoFolio = "FG-LEON",  Alias = "LEON" },
            new SucursalSeguimientoInfo { Nombre = "MIXCALCO",                  PrefijoFolio = "FG.MIX",   Alias = "MIX" },
            new SucursalSeguimientoInfo { Nombre = "SAN FRANCISCO DEL RINCON",  PrefijoFolio = "FG-SFR",   Alias = "SFCO" },
            new SucursalSeguimientoInfo { Nombre = "TABLA HONDA",               PrefijoFolio = "FG-TH",    Alias = "TABLA H" },
            new SucursalSeguimientoInfo { Nombre = "TEXTICUITZEO",              PrefijoFolio = "FG-TEXTI", Alias = "TEXTI" },
            new SucursalSeguimientoInfo { Nombre = "URIANGATO",                 PrefijoFolio = "TG-CEDIS", Alias = "URIANGATO" },
            new SucursalSeguimientoInfo { Nombre = "VILLA HIDALGO",             PrefijoFolio = "FG-VILLA", Alias = "VILLA" },
            new SucursalSeguimientoInfo { Nombre = "ZAPOTLANEJO",               PrefijoFolio = "FG-ZAPO",  Alias = "ZAPO" },
        };

        // Añade estos métodos en Form1
        private DateTime FechaCorteSeguimiento() => DateTime.Today.AddDays(-3);

        // Reemplaza el método de resumen por este (incluye UM/ER y nuevo orden)
        private void CargarResumenSeguimientos()
        {
            var dt = new DataTable();
            dt.Columns.Add("Sucursal", typeof(string));
            dt.Columns.Add("UltimaMilla", typeof(int));
            dt.Columns.Add("EnRuta", typeof(int));
            dt.Columns.Add("Documentado", typeof(int));
            dt.Columns.Add("Pendiente", typeof(int));

            DateTime corte = FechaCorteSeguimiento(); // hoy - 3 días

            foreach (var s in sucursalesSeguimientoResumen)
            {
                var c = dbManager.ObtenerConteoSeguimientoPorPrefijoAvanzado(corte, s.PrefijoFolio);
                var row = dt.NewRow();
                row["Sucursal"] = s.Nombre;
                row["UltimaMilla"] = c.UltimaMilla;
                row["EnRuta"] = c.EnRuta;
                row["Documentado"] = c.Documentado;
                row["Pendiente"] = c.Pendiente;
                dt.Rows.Add(row);
            }

            dataGridViewSeguimientos.DataSource = dt;
            dataGridViewSeguimientos.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            dataGridViewSeguimientos.ReadOnly = true;
            dataGridViewSeguimientos.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            dataGridViewSeguimientos.MultiSelect = false;
        }

        // Reemplaza el handler de clic por este (filtra UM/ER/DOC/PEN con atraso >=3 días)
        private void dataGridViewSeguimientos_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0) return;

            var sucursalNombre = Convert.ToString(dataGridViewSeguimientos.Rows[e.RowIndex].Cells["Sucursal"]?.Value)?.Trim().ToUpperInvariant();
            var info = sucursalesSeguimientoResumen.FirstOrDefault(x => x.Nombre.Equals(sucursalNombre, StringComparison.OrdinalIgnoreCase));
            if (info == null) return;

            // Estatus fijos del resumen
            var estatus = new List<string> { "ULTIMA MILLA", "EN RUTA", "DOCUMENTADO", "PENDIENTE" };

            DateTime corte = FechaCorteSeguimiento(); // hoy - 3 días
            var detalle = dbManager.ObtenerGuiasSeguimientoPorPrefijo(corte, info.PrefijoFolio, estatus);

            // Filtros de combos (si no están en TODAS)
            string origenFiltro = ValorComboOrigenesSeguimiento();
            detalle = FiltrarPorOrigenSeguimiento(detalle, origenFiltro);

            string destinoFiltro = ValorComboDestinoSeguimiento();
            detalle = FiltrarPorDestinoSeguimiento(detalle, destinoFiltro);

            CargarGridConSeleccionPersistente(detalle);
        }

        private List<(string Titulo, int Cantidad, decimal Monto)> ConstruirIndicadoresResumen(DateTime fechaInicio, DateTime fechaFin, string sucursal, string destino, DataTable dt)
        {
            var indicadores = new List<(string Titulo, int Cantidad, decimal Monto)>();
            if (dt == null) return indicadores;

            var filas = dt.AsEnumerable();
            var noCanceladas = FilasNoCanceladas(dt).ToList();
            var porCobrar = noCanceladas.Where(r => string.Equals(Convert.ToString(r["TipoCobro"]), "POR COBRAR", StringComparison.OrdinalIgnoreCase)).ToList();
            var canceladas = filas.Where(r => string.Equals(Convert.ToString(r["EstatusGuia"]), "CANCELADO", StringComparison.OrdinalIgnoreCase)).ToList();

            var um = filas.Where(r => string.Equals(Convert.ToString(r["EstatusGuia"]), "ULTIMA MILLA", StringComparison.OrdinalIgnoreCase) &&
                                      ContienePagado(r["TipoCobro"]?.ToString() ?? string.Empty)).ToList();
            var entregado = filas.Where(r => string.Equals(Convert.ToString(r["EstatusGuia"]), "ENTREGADA", StringComparison.OrdinalIgnoreCase) &&
                                             ContienePagado(r["TipoCobro"]?.ToString() ?? string.Empty)).ToList();
            var completado = filas.Where(r => string.Equals(Convert.ToString(r["EstatusGuia"]), "COMPLETADO", StringComparison.OrdinalIgnoreCase) &&
                                              ContienePagado(r["TipoCobro"]?.ToString() ?? string.Empty)).ToList();

            var resumenOrigen = dbManager.ObtenerResumenPagosOrigen(fechaInicio, fechaFin, sucursal, destino);
            var resumenDestino = dbManager.ObtenerResumenPagosDestino(fechaInicio, fechaFin, sucursal, destino);

            decimal Suma(IEnumerable<DataRow> rows, string col) => rows.Sum(r => ConvertToDecimal(r[col]));

            indicadores.Add(("Guías por cobrar", porCobrar.Count, Suma(porCobrar, "Total")));
            indicadores.Add(("Monto total por cobrar", porCobrar.Count, Suma(porCobrar, "Total"))); // mismo monto, etiqueta distinta
            indicadores.Add(("Total de guías", noCanceladas.Count, Suma(noCanceladas, "Total")));
            indicadores.Add(("Guías canceladas", canceladas.Count, Suma(canceladas, "Total")));
            indicadores.Add(("Última milla", um.Count, Suma(um, "Total")));
            indicadores.Add(("Entregado", entregado.Count, Suma(entregado, "Total")));
            indicadores.Add(("Completado", completado.Count, Suma(completado, "Total")));
            indicadores.Add(("Guías origen", resumenOrigen.Cantidad, resumenOrigen.Monto));
            indicadores.Add(("Guías destino", resumenDestino.Cantidad, resumenDestino.Monto));

            return indicadores;
        }

        // NUEVO helper: encabezado
        private void InsertarEncabezadoMes(IXLWorksheet ws, string titulo, string sucursal, string destino)
        {
            var header = ws.Range("A1:F2");
            header.Merge();
            header.Value = titulo;
            header.Style.Font.Bold = true;
            header.Style.Font.FontColor = XLColor.White;
            header.Style.Fill.BackgroundColor = XLColor.Black;
            header.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            header.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;

            ws.Row(1).Height = 18;
            ws.Row(2).Height = 18;

            ws.Cell("A3").Value = "Sucursal origen:";
            ws.Cell("A3").Style.Font.Bold = true;
            ws.Cell("B3").Value = sucursal;

            ws.Cell("C3").Value = "Sucursal destino:";
            ws.Cell("C3").Style.Font.Bold = true;
            ws.Cell("D3").Value = destino;
        }

        // NUEVO helper: tabla de indicadores
        private int EscribirIndicadoresTabla(IXLWorksheet ws, int startRow, List<(string Titulo, int Cantidad, decimal Monto)> indicadores)
        {
            ws.Cell(startRow, 1).Value = "Indicador";
            ws.Cell(startRow, 2).Value = "Cantidad";
            ws.Cell(startRow, 3).Value = "Monto";
            ws.Range(startRow, 1, startRow, 3).Style.Font.Bold = true;

            int row = startRow + 1;
            foreach (var ind in indicadores)
            {
                ws.Cell(row, 1).Value = ind.Titulo;
                ws.Cell(row, 2).Value = ind.Cantidad;
                ws.Cell(row, 3).Value = ind.Monto;
                row++;
            }

            ws.Column(3).Style.NumberFormat.Format = "$#,##0.00";
            ws.Columns(1, 3).AdjustToContents();
            return row;
        }

        // NUEVO helper: filtra indicadores para el gráfico
        private List<(string Titulo, int Cantidad, decimal Monto)> FiltrarIndicadoresParaGrafico(List<(string Titulo, int Cantidad, decimal Monto)> indicadores)
        {
            var nombres = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
            {
                "Guías por cobrar", "Guías canceladas", "Última milla", "Entregado", "Completado", "Guías origen", "Guías destino"
            };
            return indicadores.Where(i => nombres.Contains(i.Titulo)).ToList();
        }

        // NUEVO helper: gráfico de indicadores (cantidad como etiqueta, monto como valor)
        private void InsertarGraficoIndicadores(IXLWorksheet ws, int topRow, int leftCol, List<(string Titulo, int Cantidad, decimal Monto)> indicadores)
        {
            if (indicadores == null || indicadores.Count == 0) return;

            using (var chart = new System.Windows.Forms.DataVisualization.Charting.Chart())
            {
                chart.Width = 800;
                chart.Height = 400;
                var area = new ChartArea("ca");
                area.AxisX.Interval = 1;
                area.AxisX.LabelStyle.Angle = -15;
                area.AxisY.LabelStyle.Format = "$#,##0.00";
                area.AxisY.Title = "Monto";
                chart.ChartAreas.Add(area);

                var series = new Series("Indicadores")
                {
                    ChartType = SeriesChartType.Column,
                    IsValueShownAsLabel = true,
                    LabelFormat = "$#,##0.00"
                };

                foreach (var ind in indicadores)
                {
                    int idx = series.Points.AddXY(ind.Titulo, ind.Monto);
                    var dp = series.Points[idx];
                    dp.LabelForeColor = Color.Black;
                    dp.ToolTip = $"Cantidad: {ind.Cantidad}\nMonto: {ind.Monto:C}";
                }

                chart.Series.Add(series);

                using (var ms = new MemoryStream())
                {
                    chart.SaveImage(ms, ChartImageFormat.Png);
                    ms.Position = 0;
                    var pic = ws.AddPicture(ms, XLPictureFormat.Png, "GraficoIndicadores");
                    pic.MoveTo(ws.Cell(topRow, leftCol));
                    pic.WithSize(800, 400);
                }
            }
        }



        // NUEVO helper: aplica colores por TipoCobro/EstatusGuia a las celdas (ClosedXML)
        private void AplicarColoresFilasExcel(IXLWorksheet ws, DataTable dt, int startRow = 2, int startCol = 1)
        {
            if (ws == null || dt == null || dt.Rows.Count == 0)
            {
                return;
            }

            if (dt.Columns.Count > 0 && startRow > 1)
            {
                var encabezado = ws.Range(startRow - 1, startCol, startRow - 1, startCol + dt.Columns.Count - 1);
                encabezado.Style.Fill.BackgroundColor = XLColor.LightSteelBlue;
                encabezado.Style.Font.Bold = true;
                encabezado.Style.Font.FontColor = XLColor.Black;
            }

            string colEstatus = ObtenerNombreColumna(dt, "EstatusGuia", "Estatus Guía", "Estatus Guia");
            string colTipoCobro = ObtenerNombreColumna(dt, "TipoCobro", "Tipo Cobro");
            string colFactura = ObtenerNombreColumna(dt, "Factura");
            string colFolio = ObtenerNombreColumna(dt, "FolioGuia", "folio_guia", "Folio Guía", "Folio Guia");

            int idxEstatus = colEstatus != null ? dt.Columns[colEstatus].Ordinal : -1;
            int idxTipoCobro = colTipoCobro != null ? dt.Columns[colTipoCobro].Ordinal : -1;
            int idxFactura = colFactura != null ? dt.Columns[colFactura].Ordinal : -1;
            int idxFolio = colFolio != null ? dt.Columns[colFolio].Ordinal : -1;

            for (int i = 0; i < dt.Rows.Count; i++)
            {
                var row = dt.Rows[i];
                int excelRow = startRow + i;

                var colorBaseFila = (i % 2 == 0) ? Color.AliceBlue : Color.White;
                var rangoFila = ws.Range(excelRow, startCol, excelRow, startCol + dt.Columns.Count - 1);
                rangoFila.Style.Fill.BackgroundColor = XLColor.FromColor(colorBaseFila);
                rangoFila.Style.Font.FontColor = XLColor.Black;

                if (idxEstatus >= 0)
                {
                    string estatus = Convert.ToString(row[colEstatus] ?? string.Empty).Trim();
                    Color colorEstatus;
                    if (!string.IsNullOrWhiteSpace(estatus) &&
                        coloresEstatus.TryGetValue(estatus.ToUpperInvariant(), out colorEstatus))
                    {
                        rangoFila.Style.Fill.BackgroundColor = XLColor.FromColor(colorEstatus);
                        rangoFila.Style.Font.FontColor = XLColor.Black;
                    }
                }

                if (idxTipoCobro >= 0)
                {
                    string tipoCobro = Convert.ToString(row[colTipoCobro] ?? string.Empty).Trim();
                    Color colorTipoCobro;
                    if (!string.IsNullOrWhiteSpace(tipoCobro) &&
                        coloresTipoCobro.TryGetValue(tipoCobro.ToUpperInvariant(), out colorTipoCobro))
                    {
                        var celdaTipoCobro = ws.Cell(excelRow, startCol + idxTipoCobro);
                        celdaTipoCobro.Style.Fill.BackgroundColor = XLColor.FromColor(colorTipoCobro);
                        celdaTipoCobro.Style.Font.FontColor = XLColor.Black;
                    }
                }

                if (idxFactura >= 0)
                {
                    string factura = Convert.ToString(row[colFactura] ?? string.Empty).Trim();
                    if (!string.IsNullOrWhiteSpace(factura))
                    {
                        var celdaFactura = ws.Cell(excelRow, startCol + idxFactura);
                        celdaFactura.Style.Fill.BackgroundColor = XLColor.FromColor(colorFacturaConValor);
                        celdaFactura.Style.Font.FontColor = XLColor.Black;

                        if (idxFolio >= 0)
                        {
                            var celdaFolio = ws.Cell(excelRow, startCol + idxFolio);
                            celdaFolio.Style.Fill.BackgroundColor = XLColor.FromColor(colorFacturaConValor);
                            celdaFolio.Style.Font.FontColor = XLColor.Black;
                        }
                    }
                }
            }
        }


    }
}

