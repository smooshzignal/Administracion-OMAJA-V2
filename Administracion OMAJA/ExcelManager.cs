using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.Windows.Forms;
using ExcelDataReader;

namespace Administracion_Omaja
{
    public class ExcelManager
    {
        private static readonly string[] FechaHoraFormats = { "dd-MM-yyyy HH:mm", "dd/MM/yyyy HH:mm", "yyyy-MM-dd HH:mm" };
        private static readonly string[] FechaFormats = { "dd-MM-yyyy", "dd/MM/yyyy", "yyyy-MM-dd" };

        // ==================== MÉTODO: LEER EXCEL Y PROCESAR ====================
        public List<Dictionary<string, object>> LeerExcel(string filePath)
        {
            var registros = new List<Dictionary<string, object>>();

            try
            {
                System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

                using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
                {
                    using (var reader = ExcelReaderFactory.CreateReader(stream))
                    {
                        var result = reader.AsDataSet(new ExcelDataSetConfiguration()
                        {
                            ConfigureDataTable = (_) => new ExcelDataTableConfiguration()
                            {
                                UseHeaderRow = true
                            }
                        });

                        if (result.Tables.Count > 0)
                        {
                            DataTable dt = result.Tables[0];

                            foreach (DataRow row in dt.Rows)
                            {
                                var registro = ProcesarFila(row);
                                registros.Add(registro);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error al leer Excel: {ex.Message}", "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            return registros;
        }

        // ==================== MÉTODO: PROCESAR UNA FILA DEL EXCEL ====================
        private Dictionary<string, object> ProcesarFila(DataRow row)
        {
            var datos = new Dictionary<string, object>();

            object Cell(params string[] nombresColumnas) => GetCellValue(row, nombresColumnas);

            try
            {
                var fechaElaboracion = ParseDateTime(
                    Cell("Fecha/Hora Elaboración", "Fecha Hora Elaboracion", "FechaElaboracion", "Fecha Elaboracion"),
                    FechaHoraFormats);
                var fechaEntregaCompleta = ParseDateTime(
                    Cell("Fecha y Hora de Entrega", "Fecha y Hora Entrega", "FechaHoraEntrega", "Fecha Entrega", "FechaEntrega"),
                    FechaHoraFormats,
                    "NO ENTREGADO", "NO ENTREGADA");
                var fechaCancelacion = ParseDateTime(
                    Cell("Fecha de Cancelación", "Fecha Cancelacion", "FechaCancelacion"),
                    FechaFormats);
                var fechaUltimaMilla = ParseDateTime(
                    Cell("Fecha última milla", "Fecha Ultima Milla", "FechaUltimaMilla"),
                    FechaFormats,
                    "NO ENVIADA");

                datos["fechaElab"] = fechaElaboracion?.Date;
                datos["horaElab"] = fechaElaboracion?.TimeOfDay;
                datos["fechaEntrega"] = fechaEntregaCompleta?.Date;
                datos["horaEntrega"] = fechaEntregaCompleta?.TimeOfDay;
                datos["fechaCancel"] = fechaCancelacion;
                datos["fechaUltimaMilla"] = fechaUltimaMilla;

                datos["folio"] = ParseString(Cell("Folio Guía", "Folio Guia", "FolioGuia", "Folio"));
                datos["estatus"] = ParseString(Cell("Estatus Guía", "Estatus Guia", "EstatusGuia", "Estatus"));
                datos["cliente"] = ParseString(Cell("Cliente"));
                datos["ubicacion"] = ParseString(Cell("Ubicación Actual", "Ubicacion Actual", "UbicacionActual"));
                datos["origen"] = ParseString(Cell("Origen"));
                datos["destino"] = ParseString(Cell("Destino"));
                datos["tipoCobro"] = ParseString(Cell("Tipo cobro", "Tipo Cobro", "TipoCobro"));
                datos["zona"] = ParseString(Cell("Zona Operativa Entrega", "ZonaOperativaEntrega"));
                datos["tipoEntrega"] = ParseString(Cell("Tipo de entrega", "Tipo de Entrega", "TipoEntrega"));
                datos["tracking"] = ParseString(Cell("Tracking"));
                datos["referencia"] = ParseString(Cell("Referencia"));
                datos["subtotal"] = ParseDecimal(Cell("Subtotal"));
                datos["total"] = ParseDecimal(Cell("Total", "Total (MXN)", "Importe Total"));
                datos["sucursal"] = ParseString(Cell("Sucursal"));
                datos["folioInforme"] = ParseString(Cell("Folio Informe", "FolioInforme"));
                datos["folioEmbarque"] = ParseString(Cell("Folio Embarque", "FolioEmbarque"));
                datos["usuarioDoc"] = ParseString(Cell("Usuario Documento", "UsuarioDocumento"));
                datos["usuarioCancel"] = ParseString(Cell("Usuario de Cancelación", "Usuario Cancelacion", "UsuarioCancelacion"));
                datos["remitente"] = ParseString(Cell("Remitente"));
                datos["destinatario"] = ParseString(Cell("Destinatario"));
                datos["cajas"] = ParseInt(Cell("Cajas"));
                datos["valorDeclarado"] = ParseDecimal(Cell("Valor declarado", "Valor Declarado", "ValorDeclarado"));
                datos["observaciones"] = ParseString(Cell("Observaciones"));
                datos["factura"] = ParseString(Cell("Factura"));
                datos["timbradoSat"] = ParseString(Cell("Timbrado SAT", "TimbradoSat"));
                datos["folioErp"] = ParseString(Cell("Folio ERP", "FolioErp"));
                datos["tipoCobroInicial"] = ParseString(Cell("Tipo de cobro inicial", "Tipo de Cobro Inicial", "TipoCobroInicial"));
                datos["motivoCancel"] = ParseString(Cell("Motivo cancelación", "Motivo Cancelacion", "MotivoCancelacion"));
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error procesando fila: {ex.Message}");
            }

            return datos;
        }

        private static object GetCellValue(DataRow row, params string[] columnNames)
        {
            if (row?.Table == null || columnNames == null || columnNames.Length == 0)
            {
                return null;
            }

            foreach (var columnName in columnNames)
            {
                if (!string.IsNullOrWhiteSpace(columnName) && row.Table.Columns.Contains(columnName))
                {
                    return row[columnName];
                }
            }

            return null;
        }

        private static string ParseString(object value)
        {
            return value?.ToString().Trim() ?? string.Empty;
        }

        private static DateTime? ParseDateTime(object value, string[] formats, params string[] invalidTokens)
        {
            if (value == null || value == DBNull.Value)
                return null;

            if (value is double dbl)
                return DateTime.FromOADate(dbl);

            if (value is DateTime fecha)
                return fecha;

            string texto = value.ToString().Trim();
            if (string.IsNullOrEmpty(texto))
                return null;

            if (invalidTokens != null)
            {
                foreach (var token in invalidTokens)
                {
                    if (texto.Equals(token, StringComparison.OrdinalIgnoreCase))
                        return null;
                }
            }

            if (DateTime.TryParseExact(texto, formats, CultureInfo.InvariantCulture, DateTimeStyles.None, out var exacta))
                return exacta;

            if (DateTime.TryParse(texto, CultureInfo.InvariantCulture, DateTimeStyles.None, out var generica) ||
                DateTime.TryParse(texto, CultureInfo.CurrentCulture, DateTimeStyles.None, out generica))
            {
                return generica;
            }

            return null;
        }

        private static decimal ParseDecimal(object value)
        {
            if (value == null || value == DBNull.Value)
                return 0m;

            if (value is decimal dec)
                return dec;

            if (value is double dbl)
                return Convert.ToDecimal(dbl);

            decimal convertido;
            return decimal.TryParse(value.ToString(), NumberStyles.Any, CultureInfo.InvariantCulture, out convertido) ||
                   decimal.TryParse(value.ToString(), NumberStyles.Any, CultureInfo.CurrentCulture, out convertido)
                ? convertido
                : 0m;
        }

        private static int ParseInt(object value)
        {
            if (value == null || value == DBNull.Value)
                return 0;

            if (value is int entero)
                return entero;

            if (value is long largo)
                return (int)largo;

            if (value is double dbl)
                return (int)Math.Round(dbl);

            int convertido;
            return int.TryParse(value.ToString(), NumberStyles.Any, CultureInfo.InvariantCulture, out convertido) ||
                   int.TryParse(value.ToString(), NumberStyles.Any, CultureInfo.CurrentCulture, out convertido)
                ? convertido
                : 0;
        }

        // ==================== MÉTODO: MOSTRAR DATOS EN DATAGRIDVIEW ====================
        public void MostrarEnDataGridView(List<Dictionary<string, object>> registros, DataGridView dgv)
        {
            if (registros.Count == 0) return;

            // Crear DataTable
            DataTable dt = new DataTable();

            // Agregar columnas
            foreach (var key in registros[0].Keys)
            {
                dt.Columns.Add(key);
            }

            // Agregar filas
            foreach (var registro in registros)
            {
                DataRow row = dt.NewRow();
                foreach (var item in registro)
                {
                    row[item.Key] = item.Value;
                }
                dt.Rows.Add(row);
            }

            // Asignar al DataGridView
            dgv.DataSource = dt;
        }
    }
}