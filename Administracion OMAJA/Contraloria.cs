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
using Microsoft.Web.WebView2.Core;

namespace Administracion_OMAJA
{
    public partial class Contraloria : Form
    {
        private readonly DatabaseManager dbManager = new DatabaseManager();
        private readonly ToolStripProgressBar toolStripProgressBarImportacion = new ToolStripProgressBar();

        private readonly Dictionary<string, Color> coloresEstatus = new Dictionary<string, Color>(StringComparer.OrdinalIgnoreCase)
        {
            { "DOCUMENTADO", Color.FromArgb(219, 234, 254) },
            { "PENDIENTE", Color.FromArgb(255, 247, 214) },
            { "EN RUTA", Color.FromArgb(213, 245, 227) },
            { "ULTIMA MILLA", Color.FromArgb(237, 233, 254) },
            { "ENTREGADO", Color.FromArgb(221, 236, 255) },
            { "COMPLETADO", Color.FromArgb(189, 216, 255) }
        };

        private readonly Dictionary<string, Color> coloresTipoCobro = new Dictionary<string, Color>(StringComparer.OrdinalIgnoreCase)
        {
            { "POR COBRAR", Color.FromArgb(255, 236, 179) },
            { "CRÉDITO", Color.FromArgb(198, 246, 213) },
            { "CREDITO", Color.FromArgb(198, 246, 213) },
            { "PAGADO", Color.FromArgb(209, 247, 196) },
            { "CANCELADO", Color.FromArgb(255, 204, 204) }
        };

        private readonly Color colorFacturaConValor = Color.FromArgb(255, 255, 102);
        private readonly Color colorDescuentoContraloria = Color.FromArgb(153, 255, 102);

        private bool filtrosInicializados;

        private readonly Dictionary<int, string> overridesEstatusGuiasEnCortesContraloria = new Dictionary<int, string>();

        private static readonly string[] opcionesManualEstatusGuiasEnCortesContraloria =
        {
            "Faltante.",
            "Pagada en Destino.",
            "Matriz.",
            "Pagada en Origen.",
            "Pagada en Origen/No Entregado.",
            "Pendiente de pago."
        };

        private bool configurandoComboEstatusGuiasEnCortesContraloria;

        public Contraloria()
        {
            InitializeComponent();
            InicializarEventos();
            InicializarBarraProgresoImportacion();
            InicializarFiltros();
            CargarFacturacionDesdeBaseContraloria();
        }

        private void InicializarEventos()
        {
            toolStripBusqueda.Click += toolStripBusqueda_Click;
            toolStripButtoiniciarbusqueda.Click += toolStripButtoiniciarbusqueda_Click;
            toolStripButtonBuscarFactura.Click += toolStripButtonBuscarFactura_Click;
            buttonFiltrarTipodeCobro.Click += buttonFiltrarTipodeCobroContraloria_Click;
            buttonFiltrarFacturado.Click += buttonFiltrarFacturadoContraloria_Click;

            toolStriptxtBusqueda.KeyDown += toolStriptxtBusqueda_KeyDown;
            toolStripTextBoxBcliente.KeyDown += toolStripTextBoxBcliente_KeyDown;
            toolStripTextBoxBuscarFactura.KeyDown += toolStripTextBoxBuscarFactura_KeyDown;
            toolStripTextBoxBcliente.TextChanged += toolStripTextBoxBcliente_TextChanged;

            dataGridViewPrincipal.DataBindingComplete += dataGridViewPrincipal_DataBindingComplete;
            dataGridViewPrincipal.DataError += dataGridViewPrincipal_DataError;
            dataGridViewPrincipal.RowPostPaint += dataGridViewPrincipal_RowPostPaint;

            dtpFechaInicio.ValueChanged += dtpFechaInicio_ValueChanged;
            dtpFechaFin.ValueChanged += dtpFechaFin_ValueChanged;
            comboBoxSucursales.SelectedIndexChanged += comboBoxSucursales_SelectedIndexChanged;
            comboBoxSucursalDestino.SelectedIndexChanged += comboBoxSucursalDestino_SelectedIndexChanged;

            dataGridViewPrincipal.RowHeadersWidth = 56;
            dataGridViewPrincipal.RowHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;

            dataGridViewContraloria.DataBindingComplete += dataGridViewContraloria_DataBindingCompleteContraloria;
            dataGridViewContraloria.DataError += dataGridViewContraloria_DataErrorContraloria;
            dataGridViewContraloria.RowPostPaint += dataGridViewContraloria_RowPostPaintContraloria;
            dataGridViewContraloria.CellEndEdit += dataGridViewContraloria_CellEndEditGuardarObservacionesContraloria;
            dataGridViewContraloria.CellDoubleClick += dataGridViewContraloria_CellDoubleClickEditarObservacionesContraloria;
            dataGridViewContraloria.CellClick += dataGridViewContraloria_CellClickEstatusGuiasEnCortesContraloria;
            dataGridViewContraloria.CurrentCellDirtyStateChanged += dataGridViewContraloria_CurrentCellDirtyStateChangedEstatusGuiasEnCortesContraloria;
            dataGridViewContraloria.CellValueChanged += dataGridViewContraloria_CellValueChangedEstatusGuiasEnCortesContraloria;
            dataGridViewContraloria.KeyDown += dataGridViewContraloria_KeyDownCopiarContraloria;
            dataGridViewContraloria.RowHeadersWidth = 56;
            dataGridViewContraloria.RowHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
        }

        private void InicializarBarraProgresoImportacion()
        {
            toolStripProgressBarImportacion.Name = "toolStripProgressBarImportacion";
            toolStripProgressBarImportacion.Alignment = ToolStripItemAlignment.Right;
            toolStripProgressBarImportacion.AutoSize = false;
            toolStripProgressBarImportacion.Size = new Size(160, 16);
            toolStripProgressBarImportacion.Minimum = 0;
            toolStripProgressBarImportacion.Maximum = 100;
            toolStripProgressBarImportacion.Value = 0;
            toolStripProgressBarImportacion.Visible = false;

            toolStrip1.Items.Add(toolStripProgressBarImportacion);
        }

        private void PrepararBarraProgresoImportacion()
        {
            toolStripProgressBarImportacion.Minimum = 0;
            toolStripProgressBarImportacion.Maximum = 100;
            toolStripProgressBarImportacion.Value = 0;
            toolStripProgressBarImportacion.Visible = true;
        }

        private void OcultarBarraProgresoImportacion()
        {
            toolStripProgressBarImportacion.Value = 0;
            toolStripProgressBarImportacion.Visible = false;
        }

        private void importarContraloriaToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            using (var openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "Archivos Excel|*.xlsx;*.xls";
                openFileDialog.Title = "Importar Excel de guías";

                if (openFileDialog.ShowDialog() != DialogResult.OK)
                {
                    return;
                }

                PrepararBarraProgresoImportacion();

                Task.Run(() =>
                {
                    int ultimoPorcentaje = -1;

                    try
                    {
                        var resultado = dbManager.ImportarDesdeExcel(openFileDialog.FileName, dataGridViewPrincipal, (actual, total) =>
                        {
                            int porcentaje = total <= 0 ? 0 : (int)((actual * 100.0) / total);
                            if (porcentaje == ultimoPorcentaje)
                            {
                                return;
                            }

                            ultimoPorcentaje = porcentaje;

                            if (toolStrip1.IsHandleCreated)
                            {
                                toolStrip1.BeginInvoke((Action)(() =>
                                {
                                    toolStripProgressBarImportacion.Value = Math.Max(
                                        toolStripProgressBarImportacion.Minimum,
                                        Math.Min(toolStripProgressBarImportacion.Maximum, porcentaje));
                                }));
                            }
                        });

                        if (IsHandleCreated)
                        {
                            BeginInvoke((Action)(() =>
                            {
                                toolStripProgressBarImportacion.Value = 100;
                                AplicarFormatoGeneralGrid(dataGridViewPrincipal);
                                OcultarBarraProgresoImportacion();

                                MessageBox.Show(
                                    string.Format(
                                        CultureInfo.CurrentCulture,
                                        "Importación de Excel completada.\nNuevos: {0}\nActualizados: {1}",
                                        resultado.nuevos,
                                        resultado.actualizados),
                                    "Éxito",
                                    MessageBoxButtons.OK,
                                    MessageBoxIcon.Information);
                            }));
                        }
                    }
                    catch (Exception ex)
                    {
                        if (IsHandleCreated)
                        {
                            BeginInvoke((Action)(() =>
                            {
                                OcultarBarraProgresoImportacion();
                                MessageBox.Show(
                                    "Error al importar el archivo de Excel.\n" + ex.Message,
                                    "Error",
                                    MessageBoxButtons.OK,
                                    MessageBoxIcon.Error);
                            }));
                        }
                    }
                });
            }
        }

        private void toolStripBusqueda_Click(object sender, EventArgs e)
        {
            string folio = toolStriptxtBusqueda.Text.Trim();

            if (string.IsNullOrWhiteSpace(folio))
            {
                MessageBox.Show("Por favor, ingresa un folio para buscar.", "Advertencia",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            string error;
            DataTable resultados = dbManager.BuscarGuias("FolioGuia", folio, out error);

            if (!string.IsNullOrEmpty(error))
            {
                MessageBox.Show(error, "Error de búsqueda", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (resultados != null && resultados.Rows.Count > 0)
            {
                CargarGuiasEnGrid(OrdenarResultadoGuias(resultados));
            }
            else
            {
                dataGridViewPrincipal.DataSource = null;
                MessageBox.Show("No se encontró el folio especificado.", "Sin resultados",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
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
                CargarGuiasEnGrid(resultados);
            }
            else
            {
                dataGridViewPrincipal.DataSource = null;
                MessageBox.Show("No se encontraron guías con la factura indicada.", "Sin resultados",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
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

            AplicarFormatoGeneralGrid(dataGridViewPrincipal);
        }

        private void CargarGuiasEnGrid(DataTable datos)
        {
            dataGridViewPrincipal.DataSource = datos;
            AplicarFormatoGeneralGrid(dataGridViewPrincipal);
        }

        private static DataTable OrdenarResultadoGuias(DataTable resultados)
        {
            if (resultados == null || resultados.Rows.Count == 0 || !resultados.Columns.Contains("FolioGuia"))
            {
                return resultados;
            }

            DataTable dtOrdenada = new DataTable();

            string colFolio = "FolioGuia";
            dtOrdenada.Columns.Add(colFolio, resultados.Columns[colFolio].DataType);

            string colFecha = resultados.Columns.Contains("FechaElaboracion") ? "FechaElaboracion" : null;
            if (colFecha != null && colFecha != colFolio)
            {
                dtOrdenada.Columns.Add(colFecha, resultados.Columns[colFecha].DataType);
            }

            foreach (DataColumn col in resultados.Columns)
            {
                if (col.ColumnName != colFolio && col.ColumnName != colFecha)
                {
                    dtOrdenada.Columns.Add(col.ColumnName, col.DataType);
                }
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

            return dtOrdenada;
        }

        private void AplicarFormatoGeneralGrid(DataGridView dgv)
        {
            FormatearEncabezadosDataGridView(dgv);
            AplicarEstiloAzulClaroDataGridView(dgv);
            ColorearFilasEstatus(dgv);
            ColorearFacturaConValor(dgv);
            ColorearCeldaTipoCobro(dgv);
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
                { "MotivoCancelacion", "Motivo Cancelación" }
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

            dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCells;
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

        private void ColorearCeldaTipoCobro(DataGridView dgv)
        {
            if (dgv?.Rows == null || !dgv.Columns.Contains("TipoCobro"))
            {
                return;
            }

            foreach (DataGridViewRow row in dgv.Rows)
            {
                if (row.IsNewRow)
                {
                    continue;
                }

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

        private void ColorearFacturaConValor(DataGridView dgv)
        {
            if (dgv?.Rows == null || !dgv.Columns.Contains("Factura"))
            {
                return;
            }

            string colFolio = dgv.Columns.Contains("FolioGuia")
                ? "FolioGuia"
                : dgv.Columns.Contains("folio_guia")
                    ? "folio_guia"
                    : null;

            foreach (DataGridViewRow row in dgv.Rows)
            {
                if (row.IsNewRow)
                {
                    continue;
                }

                var facturaCell = row.Cells["Factura"];
                string valor = Convert.ToString(facturaCell?.Value ?? string.Empty).Trim();

                if (!string.IsNullOrEmpty(valor))
                {
                    facturaCell.Style.BackColor = colorFacturaConValor;
                    facturaCell.Style.SelectionBackColor = ControlPaint.Dark(colorFacturaConValor);
                    facturaCell.Style.SelectionForeColor = Color.Black;

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

        private string ObtenerEstatusFila(DataGridViewRow row)
        {
            if (row == null)
            {
                return string.Empty;
            }

            var cols = row.DataGridView?.Columns;
            if (cols == null)
            {
                return string.Empty;
            }

            string Valor(string colName)
            {
                return cols.Contains(colName)
                    ? Convert.ToString(row.Cells[colName]?.Value ?? string.Empty).Trim()
                    : null;
            }

            return Valor("EstatusGuia")
                ?? Valor("Estatus Guía")
                ?? Valor("Estatus guias en cortes")
                ?? string.Empty;
        }

        private void ColorearFilasEstatus(DataGridView dgv)
        {
            if (dgv?.Rows == null)
            {
                return;
            }

            foreach (DataGridViewRow row in dgv.Rows)
            {
                if (row.IsNewRow)
                {
                    continue;
                }

                row.DefaultCellStyle.BackColor = dgv.DefaultCellStyle.BackColor;
                row.DefaultCellStyle.SelectionBackColor = dgv.DefaultCellStyle.SelectionBackColor;
                row.DefaultCellStyle.SelectionForeColor = dgv.DefaultCellStyle.SelectionForeColor;

                string estatus = ObtenerEstatusFila(row);
                if (string.IsNullOrWhiteSpace(estatus))
                {
                    continue;
                }

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
            AplicarFormatoGeneralGrid(dataGridViewPrincipal);
        }

        private void dataGridViewPrincipal_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {
            var grid = sender as DataGridView;
            var colName = grid?.Columns[e.ColumnIndex]?.Name ?? "(col)";
            var val = grid?.Rows[e.RowIndex]?.Cells[e.ColumnIndex]?.Value;

            MessageBox.Show(
                string.Format(
                    CultureInfo.CurrentCulture,
                    "DataError en columna '{0}'. Valor: '{1}'. Error: {2}",
                    colName,
                    val ?? "null",
                    e.Exception?.Message),
                "DataError",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);

            e.ThrowException = true;
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

        private void exportarContraloriaToolStripMenuItem1_Click(object sender, EventArgs e)
        {
        }

        private void importarCorteToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ImportarCorteContraloria();
        }

        private void ImportarCorteContraloria()
        {
            using (var openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "Archivos Excel|*.xlsx;*.xls";
                openFileDialog.Title = "Importar Excel de facturación";

                if (openFileDialog.ShowDialog() != DialogResult.OK)
                {
                    return;
                }

                PrepararBarraProgresoImportacion();

                Task.Run(() =>
                {
                    int ultimoPorcentaje = -1;

                    try
                    {
                        var resultado = dbManager.ImportarFacturacionDesdeExcelContraloria(openFileDialog.FileName, (actual, total) =>
                        {
                            int porcentaje = total <= 0 ? 0 : (int)((actual * 100.0) / total);
                            if (porcentaje == ultimoPorcentaje)
                            {
                                return;
                            }

                            ultimoPorcentaje = porcentaje;

                            if (toolStrip1.IsHandleCreated)
                            {
                                toolStrip1.BeginInvoke((Action)(() =>
                                {
                                    toolStripProgressBarImportacion.Value = Math.Max(
                                        toolStripProgressBarImportacion.Minimum,
                                        Math.Min(toolStripProgressBarImportacion.Maximum, porcentaje));
                                }));
                            }
                        });

                        if (IsHandleCreated)
                        {
                            BeginInvoke((Action)(() =>
                            {
                                CargarFacturacionDesdeBaseContraloria();

                                toolStripProgressBarImportacion.Value = 100;
                                OcultarBarraProgresoImportacion();

                                MessageBox.Show(
                                    string.Format(
                                        CultureInfo.CurrentCulture,
                                        "Importación de facturación completada.\nNuevos: {0}\nActualizados: {1}",
                                        resultado.nuevos,
                                        resultado.actualizados),
                                    "Éxito",
                                    MessageBoxButtons.OK,
                                    MessageBoxIcon.Information);
                            }));
                        }
                    }
                    catch (Exception ex)
                    {
                        if (IsHandleCreated)
                        {
                            BeginInvoke((Action)(() =>
                            {
                                OcultarBarraProgresoImportacion();
                                MessageBox.Show(
                                    "Error al importar el archivo de facturación.\n" + ex.Message,
                                    "Error",
                                    MessageBoxButtons.OK,
                                    MessageBoxIcon.Error);
                            }));
                        }
                    }
                });
            }
        }


        private void dataGridViewContraloria_DataBindingCompleteContraloria(object sender, DataGridViewBindingCompleteEventArgs e)
        {
            AplicarFormatoDataGridViewFacturacionContraloria();
        }

        private void dataGridViewContraloria_DataErrorContraloria(object sender, DataGridViewDataErrorEventArgs e)
        {
            var grid = sender as DataGridView;

            if (grid != null &&
                e.ColumnIndex >= 0 &&
                e.ColumnIndex < grid.Columns.Count &&
                EsColumnaEstatusGuiasEnCortesContraloria(grid.Columns[e.ColumnIndex].Name))
            {
                e.ThrowException = false;
                return;
            }

            var colName = grid?.Columns[e.ColumnIndex]?.Name ?? "(col)";
            var val = grid?.Rows[e.RowIndex]?.Cells[e.ColumnIndex]?.Value;

            MessageBox.Show(
                string.Format(
                    CultureInfo.CurrentCulture,
                    "DataError en columna '{0}'. Valor: '{1}'. Error: {2}",
                    colName,
                    val ?? "null",
                    e.Exception?.Message),
                "DataError",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);

            e.ThrowException = true;
        }

        private void dataGridViewContraloria_RowPostPaintContraloria(object sender, DataGridViewRowPostPaintEventArgs e)
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

        private void AplicarFormatoDataGridViewFacturacionContraloria()
        {
            ConfigurarComboBoxEstatusGuiasEnCortesContraloria();
            FormatearEncabezadosDataGridViewFacturacionContraloria(dataGridViewContraloria);
            AplicarEstiloAzulClaroDataGridViewFacturacionContraloria(dataGridViewContraloria);
            ColorearFilasEstatus(dataGridViewContraloria);
            ColorearCeldasDescuentoContraloria(dataGridViewContraloria);
            ConfigurarEdicionDataGridViewFacturacionContraloria();
            ConfigurarCopiadoDataGridViewFacturacionContraloria();
            AsegurarColumnaObservacionesAuditoriaVisibleContraloria();
        }

        private void ConfigurarEdicionDataGridViewFacturacionContraloria()
        {
            if (dataGridViewContraloria == null || dataGridViewContraloria.Columns.Count == 0)
            {
                return;
            }

            dataGridViewContraloria.ReadOnly = false;
            dataGridViewContraloria.EditMode = DataGridViewEditMode.EditProgrammatically;
            dataGridViewContraloria.SelectionMode = DataGridViewSelectionMode.CellSelect;
            dataGridViewContraloria.MultiSelect = true;
            dataGridViewContraloria.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None;

            foreach (DataGridViewColumn col in dataGridViewContraloria.Columns)
            {
                bool esEditable =
                    EsColumnaObservacionesAuditoriaContraloria(col.Name) ||
                    EsColumnaEstatusGuiasEnCortesContraloria(col.Name);

                col.ReadOnly = !esEditable;
            }

            if (dataGridViewContraloria.Columns.Contains("Observaciones de auditoria"))
            {
                var columnaObservaciones = dataGridViewContraloria.Columns["Observaciones de auditoria"];
                columnaObservaciones.Visible = true;
                columnaObservaciones.ReadOnly = false;
                columnaObservaciones.DefaultCellStyle.WrapMode = DataGridViewTriState.True;
                columnaObservaciones.SortMode = DataGridViewColumnSortMode.NotSortable;
                columnaObservaciones.AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells;
            }

            if (dataGridViewContraloria.Columns.Contains("Estatus guias en cortes"))
            {
                var columnaEstatus = dataGridViewContraloria.Columns["Estatus guias en cortes"];
                columnaEstatus.Visible = true;
                columnaEstatus.ReadOnly = false;
            }
        }


        private void AsegurarColumnaObservacionesAuditoriaVisibleContraloria()
        {
            if (dataGridViewContraloria == null || dataGridViewContraloria.Columns.Count == 0)
            {
                return;
            }

            if (!dataGridViewContraloria.Columns.Contains("Observaciones de auditoria"))
            {
                return;
            }

            var columna = dataGridViewContraloria.Columns["Observaciones de auditoria"];
            columna.Visible = true;
            columna.ReadOnly = false;
            columna.DefaultCellStyle.WrapMode = DataGridViewTriState.True;
            columna.SortMode = DataGridViewColumnSortMode.NotSortable;

            if (columna.AutoSizeMode != DataGridViewAutoSizeColumnMode.DisplayedCells)
            {
                columna.AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells;
            }
        }

        private void ConfigurarCopiadoDataGridViewFacturacionContraloria()
        {
            if (dataGridViewContraloria == null)
            {
                return;
            }

            dataGridViewContraloria.SelectionMode = DataGridViewSelectionMode.CellSelect;
            dataGridViewContraloria.MultiSelect = true;
            dataGridViewContraloria.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableWithoutHeaderText;
        }

        private void dataGridViewContraloria_KeyDownCopiarContraloria(object sender, KeyEventArgs e)
        {
            if (!e.Control || e.KeyCode != Keys.C)
            {
                return;
            }

            var grid = sender as DataGridView;
            if (grid == null)
            {
                return;
            }

            try
            {
                if (grid.GetCellCount(DataGridViewElementStates.Selected) > 0)
                {
                    var contenido = grid.GetClipboardContent();
                    if (contenido != null)
                    {
                        Clipboard.SetDataObject(contenido);
                        e.Handled = true;
                        e.SuppressKeyPress = true;
                        return;
                    }
                }

                if (grid.CurrentCell != null && grid.CurrentCell.Value != null)
                {
                    Clipboard.SetText(Convert.ToString(grid.CurrentCell.Value, CultureInfo.CurrentCulture));
                    e.Handled = true;
                    e.SuppressKeyPress = true;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    "No se pudo copiar el contenido seleccionado.\n" + ex.Message,
                    "Copiar",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
            }
        }

        private void dataGridViewContraloria_CellDoubleClickEditarObservacionesContraloria(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0 || e.ColumnIndex < 0)
            {
                return;
            }

            var grid = sender as DataGridView;
            if (grid == null)
            {
                return;
            }

            var columna = grid.Columns[e.ColumnIndex];
            if (columna == null || !columna.Name.Equals("Observaciones de auditoria", StringComparison.OrdinalIgnoreCase))
            {
                return;
            }

            grid.CurrentCell = grid.Rows[e.RowIndex].Cells[e.ColumnIndex];
            grid.BeginEdit(true);
        }

        private void dataGridViewContraloria_CellEndEditGuardarObservacionesContraloria(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0 || e.ColumnIndex < 0)
            {
                return;
            }

            var grid = sender as DataGridView;
            if (grid == null)
            {
                return;
            }

            var columna = grid.Columns[e.ColumnIndex];
            if (columna == null || !columna.Name.Equals("Observaciones de auditoria", StringComparison.OrdinalIgnoreCase))
            {
                return;
            }

            GuardarObservacionesAuditoriaDataGridViewContraloria(grid.Rows[e.RowIndex]);
        }

        private void GuardarObservacionesAuditoriaDataGridViewContraloria(DataGridViewRow row)
        {
            if (row == null || row.IsNewRow || row.DataGridView == null)
            {
                return;
            }

            if (!row.DataGridView.Columns.Contains("id"))
            {
                MessageBox.Show(
                    "No se encontró la columna Id para guardar observaciones de auditoría.",
                    "Dato faltante",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
                return;
            }

            int id;
            if (!int.TryParse(Convert.ToString(row.Cells["id"]?.Value ?? string.Empty, CultureInfo.CurrentCulture), out id))
            {
                MessageBox.Show(
                    "No se pudo identificar el Id del registro para guardar observaciones de auditoría.",
                    "Dato inválido",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
                return;
            }

            string observaciones = row.DataGridView.Columns.Contains("Observaciones de auditoria")
                ? Convert.ToString(row.Cells["Observaciones de auditoria"]?.Value ?? string.Empty)
                : string.Empty;

            bool ok = dbManager.ActualizarObservacionesAuditoriaContraloria(id, observaciones);
            if (!ok)
            {
                MessageBox.Show(
                    "No se pudieron guardar las observaciones de auditoría.",
                    "Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        private void FormatearEncabezadosDataGridViewFacturacionContraloria(DataGridView dgv)
        {
            if (dgv == null || dgv.Columns.Count == 0)
            {
                return;
            }

            var encabezados = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
     {
         { "id", "Id" },
         { "sucursal", "Sucursal" },
         { "fecha", "Fecha" },
         { "numero", "Número" },
         { "cliente", "Cliente" },
         { "documento", "Documento" },
         { "nota_de_debito", "Nota de Débito" },
         { "uuid", "UUID" },
         { "descuento", "Descuento" },
         { "sub_total", "Sub Total" },
         { "iva", "IVA" },
         { "retencion", "Retención" },
         { "total", "Total" },
         { "moneda", "Moneda" },
         { "estatus", "Estatus" },
         { "folio_fiscal_uuid", "Folio Fiscal UUID" },
         { "destino", "Destino" },
         { "origen", "Origen" },
         { "no_viaje", "No. Viaje" },
         { "Estatus guias en cortes", "Estatus guías en cortes" },
         { "Busqueda en cortes", "Búsqueda en cortes" },
         { "Observaciones de auditoria", "Observaciones de auditoría" }
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

            string[] columnasMonedaContraloria =
            {
        "nota_de_debito",
        "descuento",
        "sub_total",
        "iva",
        "retencion",
        "total"
    };

            foreach (string nombreColumna in columnasMonedaContraloria)
            {
                if (dgv.Columns.Contains(nombreColumna))
                {
                    dgv.Columns[nombreColumna].DefaultCellStyle.Format = "C";
                    dgv.Columns[nombreColumna].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                }
            }
        }

        private void AplicarEstiloAzulClaroDataGridViewFacturacionContraloria(DataGridView dgv)
        {
            if (dgv == null)
            {
                return;
            }

            dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCells;
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

        private void exportarCorteToolStripMenuItem_Click(object sender, EventArgs e)
        {
        }

        private void InicializarFiltros()
        {
            filtrosInicializados = false;

            if (comboBoxSucursales.Items.Count > 0)
            {
                int indexTodas = comboBoxSucursales.FindStringExact("TODAS");
                comboBoxSucursales.SelectedIndex = indexTodas >= 0 ? indexTodas : 0;
            }

            if (comboBoxSucursalDestino.Items.Count > 0)
            {
                int indexTodas = comboBoxSucursalDestino.FindStringExact("TODAS");
                comboBoxSucursalDestino.SelectedIndex = indexTodas >= 0 ? indexTodas : 0;
            }

            DateTime hoy = DateTime.Today;
            dtpFechaFin.Value = hoy;
            dtpFechaInicio.Value = new DateTime(hoy.Year, hoy.Month, 1);

            filtrosInicializados = true;
            CargarGuiasPorFiltros();
        }

        private void CargarGuiasPorFiltros(bool mostrarMensajeSinResultados = false)
        {
            AplicarFiltrosContraloria(mostrarMensajeSinResultados);
        }

        private void dtpFechaInicio_ValueChanged(object sender, EventArgs e)
        {
            if (dtpFechaFin.Value < dtpFechaInicio.Value)
            {
                dtpFechaFin.Value = dtpFechaInicio.Value;
            }

            if (!filtrosInicializados)
            {
                return;
            }

            CargarGuiasPorFiltros();
        }

        private void dtpFechaFin_ValueChanged(object sender, EventArgs e)
        {
            if (dtpFechaFin.Value < dtpFechaInicio.Value)
            {
                dtpFechaFin.Value = dtpFechaInicio.Value;
            }

            if (!filtrosInicializados)
            {
                return;
            }

            CargarGuiasPorFiltros();
        }

        private void comboBoxSucursales_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (!filtrosInicializados)
            {
                return;
            }

            CargarGuiasPorFiltros();
        }

        private void comboBoxSucursalDestino_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (!filtrosInicializados)
            {
                return;
            }

            CargarGuiasPorFiltros();
        }


        private void buttonFiltrarTipodeCobroContraloria_Click(object sender, EventArgs e)
        {
            AplicarFiltrosContraloria(true);
        }



        private void TruncarCortesContraloria()
        {
            const string mensaje = "Esta acción eliminará de forma permanente todos los registros de la tabla Cortes. ¿Deseas continuar?";
            var confirmacion = MessageBox.Show(
                mensaje,
                "Confirmar eliminación masiva",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Warning,
                MessageBoxDefaultButton.Button2);

            if (confirmacion != DialogResult.Yes)
            {
                return;
            }

            System.Windows.Forms.Cursor cursorAnterior = System.Windows.Forms.Cursor.Current;
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor;

            try
            {
                if (dbManager.TruncarCortesContraloria())
                {
                    dataGridViewContraloria.DataSource = null;
                    MessageBox.Show(
                        "Todos los registros de facturación fueron eliminados.",
                        "Operación completada",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information);
                }
            }
            finally
            {
                System.Windows.Forms.Cursor.Current = cursorAnterior;
            }
        }

        private void truncarCortesToolStripMenuItem_Click_1(object sender, EventArgs e)
        {
            TruncarCortesContraloria();
        }

        private void CargarFacturacionDesdeBaseContraloria()
        {
            DataTable datos = dbManager.ObtenerFacturacionContraloria();
            AplicarOverridesEstatusGuiasEnCortesContraloria(datos);
            dataGridViewContraloria.DataSource = datos;
            AplicarFormatoDataGridViewFacturacionContraloria();
        }

        private void buttonFiltrarFacturadoContraloria_Click(object sender, EventArgs e)
        {
            AplicarFiltrosContraloria(true);
        }

        private void AplicarFiltrosContraloria(bool mostrarMensajeSinResultados)
        {
            DateTime fechaInicio = dtpFechaInicio.Value.Date;
            DateTime fechaFin = dtpFechaFin.Value.Date;

            if (fechaFin < fechaInicio)
            {
                var temp = fechaInicio;
                fechaInicio = fechaFin;
                fechaFin = temp;
            }

            string sucursal = comboBoxSucursales.SelectedItem?.ToString() ?? "TODAS";
            string destino = comboBoxSucursalDestino.SelectedItem?.ToString() ?? "TODAS";

            bool incluirPagado = checkBox1.Checked;
            bool incluirPorCobrar = checkBox2.Checked;
            bool incluirCancelado = checkBox3.Checked;
            bool incluirCompletado = checkBox4.Checked;
            bool incluirEntregada = checkBox5.Checked;
            bool incluirFacturada = checkBox6.Checked;
            bool incluirNoFacturada = checkBox7.Checked;

            DataTable datos = dbManager.ObtenerGuiasPorFiltroTipoCobroContraloria(
                fechaInicio,
                fechaFin,
                sucursal,
                destino,
                incluirPagado,
                incluirPorCobrar,
                incluirCancelado,
                incluirCompletado,
                incluirEntregada,
                incluirFacturada,
                incluirNoFacturada);

            if (datos == null || datos.Rows.Count == 0)
            {
                dataGridViewPrincipal.DataSource = null;

                if (mostrarMensajeSinResultados)
                {
                    MessageBox.Show(
                        ConstruirMensajeSinResultadosContraloria(sucursal, destino),
                        "Sin resultados",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information);
                }

                return;
            }

            CargarGuiasEnGrid(datos);
        }

        private string ConstruirDescripcionFiltrosContraloria()
        {
            var filtros = new List<string>();

            if (checkBox1.Checked)
            {
                filtros.Add("Pagado");
            }

            if (checkBox2.Checked)
            {
                filtros.Add("Por Cobrar");
            }

            if (checkBox3.Checked)
            {
                filtros.Add("Cancelado");
            }

            if (checkBox4.Checked)
            {
                filtros.Add("Completado");
            }

            if (checkBox5.Checked)
            {
                filtros.Add("Entregada");
            }

            if (checkBox6.Checked)
            {
                filtros.Add("Facturada");
            }

            if (checkBox7.Checked)
            {
                filtros.Add("No facturada");
            }

            return filtros.Count == 0
                ? "Sin filtros de checkboxes"
                : string.Join(", ", filtros);
        }

        private string ConstruirMensajeSinResultadosContraloria(string sucursal, string destino)
        {
            var sb = new StringBuilder();

            sb.AppendLine("No se encontraron resultados para la combinación seleccionada.");
            sb.AppendLine();
            sb.AppendLine("Filtros aplicados:");
            sb.AppendLine(" - CheckBoxes: " + ConstruirDescripcionFiltrosContraloria());
            sb.AppendLine(" - Sucursal: " + sucursal);
            sb.AppendLine(" - Destino: " + destino);
            sb.AppendLine(" - Fecha inicio: " + dtpFechaInicio.Value.Date.ToString("dd/MM/yyyy", CultureInfo.CurrentCulture));
            sb.AppendLine(" - Fecha fin: " + dtpFechaFin.Value.Date.ToString("dd/MM/yyyy", CultureInfo.CurrentCulture));

            if (checkBox1.Checked && checkBox2.Checked)
            {
                sb.AppendLine();
                sb.AppendLine("Nota: 'Pagado' y 'Por Cobrar' al mismo tiempo normalmente no tendrán coincidencias.");
            }

            if (checkBox6.Checked && checkBox7.Checked)
            {
                sb.AppendLine();
                sb.AppendLine("Nota: 'Facturada' y 'No facturada' al mismo tiempo normalmente no tendrán coincidencias.");
            }

            return sb.ToString();
        }

        private static bool EsColumnaEstatusGuiasEnCortesContraloria(string nombreColumna)
        {
            return !string.IsNullOrWhiteSpace(nombreColumna) &&
                   nombreColumna.Equals("Estatus guias en cortes", StringComparison.OrdinalIgnoreCase);
        }

        private static bool EsColumnaObservacionesAuditoriaContraloria(string nombreColumna)
        {
            return !string.IsNullOrWhiteSpace(nombreColumna) &&
                   nombreColumna.Equals("Observaciones de auditoria", StringComparison.OrdinalIgnoreCase);
        }

        private void AplicarOverridesEstatusGuiasEnCortesContraloria(DataTable datos)
        {
            if (datos == null ||
                !datos.Columns.Contains("id") ||
                !datos.Columns.Contains("Estatus guias en cortes"))
            {
                return;
            }

            foreach (DataRow row in datos.Rows)
            {
                int id;
                if (!int.TryParse(Convert.ToString(row["id"] ?? string.Empty, CultureInfo.CurrentCulture), out id))
                {
                    continue;
                }

                string valorOverride;
                if (overridesEstatusGuiasEnCortesContraloria.TryGetValue(id, out valorOverride))
                {
                    row["Estatus guias en cortes"] = valorOverride;
                }
            }
        }

        private void ConfigurarComboBoxEstatusGuiasEnCortesContraloria()
        {
            if (configurandoComboEstatusGuiasEnCortesContraloria)
            {
                return;
            }

            if (dataGridViewContraloria == null ||
                dataGridViewContraloria.Columns.Count == 0 ||
                !dataGridViewContraloria.Columns.Contains("Estatus guias en cortes"))
            {
                return;
            }

            configurandoComboEstatusGuiasEnCortesContraloria = true;

            try
            {
                var columnaActual = dataGridViewContraloria.Columns["Estatus guias en cortes"];
                int index = columnaActual.Index;
                string dataPropertyName = string.IsNullOrWhiteSpace(columnaActual.DataPropertyName)
                    ? columnaActual.Name
                    : columnaActual.DataPropertyName;
                string headerText = string.IsNullOrWhiteSpace(columnaActual.HeaderText)
                    ? columnaActual.Name
                    : columnaActual.HeaderText;
                bool visible = columnaActual.Visible;

                if (!(columnaActual is DataGridViewComboBoxColumn))
                {
                    dataGridViewContraloria.Columns.Remove(columnaActual);

                    var combo = new DataGridViewComboBoxColumn
                    {
                        Name = "Estatus guias en cortes",
                        DataPropertyName = dataPropertyName,
                        HeaderText = headerText,
                        DisplayStyle = DataGridViewComboBoxDisplayStyle.DropDownButton,
                        DisplayStyleForCurrentCellOnly = true,
                        FlatStyle = FlatStyle.Flat,
                        SortMode = DataGridViewColumnSortMode.NotSortable
                    };

                    dataGridViewContraloria.Columns.Insert(index, combo);
                }

                var columnaCombo = dataGridViewContraloria.Columns["Estatus guias en cortes"] as DataGridViewComboBoxColumn;
                if (columnaCombo == null)
                {
                    return;
                }

                columnaCombo.Visible = visible;
                columnaCombo.DisplayStyle = DataGridViewComboBoxDisplayStyle.DropDownButton;
                columnaCombo.DisplayStyleForCurrentCellOnly = true;
                columnaCombo.FlatStyle = FlatStyle.Flat;
                columnaCombo.MaxDropDownItems = 8;

                var opciones = ObtenerOpcionesDisponiblesEstatusGuiasEnCortesContraloria();

                columnaCombo.Items.Clear();
                foreach (string opcion in opciones)
                {
                    columnaCombo.Items.Add(opcion);
                }

                foreach (DataGridViewRow row in dataGridViewContraloria.Rows)
                {
                    if (!row.IsNewRow)
                    {
                        ConfigurarOpcionesCeldaEstatusGuiasEnCortesContraloria(row);
                    }
                }
            }
            finally
            {
                configurandoComboEstatusGuiasEnCortesContraloria = false;
            }
        }

        private void ConfigurarOpcionesCeldaEstatusGuiasEnCortesContraloria(DataGridViewRow row)
        {
            if (row == null ||
                row.IsNewRow ||
                row.DataGridView == null ||
                !row.DataGridView.Columns.Contains("Estatus guias en cortes"))
            {
                return;
            }

            var cell = row.Cells["Estatus guias en cortes"] as DataGridViewComboBoxCell;
            if (cell == null)
            {
                return;
            }

            cell.DisplayStyle = DataGridViewComboBoxDisplayStyle.DropDownButton;
            cell.FlatStyle = FlatStyle.Flat;
        }


        private void dataGridViewContraloria_CellClickEstatusGuiasEnCortesContraloria(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0 || e.ColumnIndex < 0)
            {
                return;
            }

            var grid = sender as DataGridView;
            if (grid == null)
            {
                return;
            }

            var columna = grid.Columns[e.ColumnIndex];
            if (columna == null || !EsColumnaEstatusGuiasEnCortesContraloria(columna.Name))
            {
                return;
            }

            grid.CurrentCell = grid.Rows[e.RowIndex].Cells[e.ColumnIndex];
            grid.BeginEdit(true);

            BeginInvoke((Action)(() =>
            {
                var combo = grid.EditingControl as ComboBox;
                if (combo != null)
                {
                    combo.DroppedDown = true;
                }
            }));
        }

        private void dataGridViewContraloria_CurrentCellDirtyStateChangedEstatusGuiasEnCortesContraloria(object sender, EventArgs e)
        {
            var grid = sender as DataGridView;
            if (grid == null || !grid.IsCurrentCellDirty || grid.CurrentCell == null)
            {
                return;
            }

            if (!EsColumnaEstatusGuiasEnCortesContraloria(grid.CurrentCell.OwningColumn?.Name))
            {
                return;
            }

            grid.CommitEdit(DataGridViewDataErrorContexts.Commit);
        }

        private void dataGridViewContraloria_CellValueChangedEstatusGuiasEnCortesContraloria(object sender, DataGridViewCellEventArgs e)
        {
            if (configurandoComboEstatusGuiasEnCortesContraloria || e.RowIndex < 0 || e.ColumnIndex < 0)
            {
                return;
            }

            var grid = sender as DataGridView;
            if (grid == null)
            {
                return;
            }

            var columna = grid.Columns[e.ColumnIndex];
            if (columna == null || !EsColumnaEstatusGuiasEnCortesContraloria(columna.Name))
            {
                return;
            }

            var row = grid.Rows[e.RowIndex];

            if (!grid.Columns.Contains("id"))
            {
                return;
            }

            int id;
            if (!int.TryParse(Convert.ToString(row.Cells["id"]?.Value ?? string.Empty, CultureInfo.CurrentCulture), out id))
            {
                return;
            }

            string valor = Convert.ToString(row.Cells[columna.Name]?.Value ?? string.Empty, CultureInfo.CurrentCulture).Trim();

            if (string.IsNullOrWhiteSpace(valor))
            {
                overridesEstatusGuiasEnCortesContraloria.Remove(id);
            }
            else
            {
                overridesEstatusGuiasEnCortesContraloria[id] = valor;
            }
        }

        private List<string> ObtenerOpcionesDisponiblesEstatusGuiasEnCortesContraloria()
        {
            var opciones = new List<string>();
            var existentes = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            foreach (string opcion in opcionesManualEstatusGuiasEnCortesContraloria)
            {
                if (!string.IsNullOrWhiteSpace(opcion) && existentes.Add(opcion))
                {
                    opciones.Add(opcion);
                }
            }

            if (dataGridViewContraloria != null &&
                dataGridViewContraloria.Columns.Contains("Estatus guias en cortes"))
            {
                foreach (DataGridViewRow row in dataGridViewContraloria.Rows)
                {
                    if (row == null || row.IsNewRow)
                    {
                        continue;
                    }

                    string valorActual = Convert.ToString(
                        row.Cells["Estatus guias en cortes"]?.Value ?? string.Empty,
                        CultureInfo.CurrentCulture).Trim();

                    if (!string.IsNullOrWhiteSpace(valorActual) && existentes.Add(valorActual))
                    {
                        opciones.Add(valorActual);
                    }
                }
            }

            return opciones;
        }
        private void ColorearCeldasDescuentoContraloria(DataGridView dgv)
        {
            if (dgv?.Rows == null)
            {
                return;
            }

            string nombreColumnaDescuento = dgv.Columns.Contains("descuento")
                ? "descuento"
                : dgv.Columns.Contains("Descuento")
                    ? "Descuento"
                    : null;

            if (nombreColumnaDescuento == null)
            {
                return;
            }

            foreach (DataGridViewRow row in dgv.Rows)
            {
                if (row.IsNewRow)
                {
                    continue;
                }

                var cell = row.Cells[nombreColumnaDescuento];
                if (cell == null)
                {
                    continue;
                }

                decimal descuento;
                bool esNumero =
                    decimal.TryParse(Convert.ToString(cell.Value ?? string.Empty, CultureInfo.CurrentCulture), NumberStyles.Any, CultureInfo.CurrentCulture, out descuento) ||
                    decimal.TryParse(Convert.ToString(cell.Value ?? string.Empty, CultureInfo.InvariantCulture), NumberStyles.Any, CultureInfo.InvariantCulture, out descuento);

                if (!esNumero || descuento <= 0m)
                {
                    cell.Style.BackColor = dgv.DefaultCellStyle.BackColor;
                    cell.Style.SelectionBackColor = dgv.DefaultCellStyle.SelectionBackColor;
                    cell.Style.SelectionForeColor = dgv.DefaultCellStyle.SelectionForeColor;
                    continue;
                }

                cell.Style.BackColor = colorDescuentoContraloria;
                cell.Style.SelectionBackColor = ControlPaint.Dark(colorDescuentoContraloria);
                cell.Style.SelectionForeColor = Color.Black;
            }
        }
    }
}
