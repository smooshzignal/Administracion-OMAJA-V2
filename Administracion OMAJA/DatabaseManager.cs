using Administracion_Omaja;
using DocumentFormat.OpenXml.Office.Word;
using ExcelDataReader;
using MySql.Data.MySqlClient;
using Mysqlx;
using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Administracion_OMAJA
{
    internal class DatabaseManager
    {
        private readonly string _connectionString = "Server=localhost;Database=adminomaja;Uid=root;Pwd=omaja123;Port=3306;";

        private static bool DebeFiltrarSucursal(string sucursal)
        {
            return !string.IsNullOrWhiteSpace(sucursal) && !sucursal.Equals("TODAS", StringComparison.OrdinalIgnoreCase);
        }

        private static bool DebeFiltrarDestino(string destino)
        {
            return !string.IsNullOrWhiteSpace(destino) && !destino.Equals("TODAS", StringComparison.OrdinalIgnoreCase);
        }

        private static bool DebeFiltrarOrigen(string origen)
        {
            return !string.IsNullOrWhiteSpace(origen) && !origen.Equals("TODAS", StringComparison.OrdinalIgnoreCase);
        }

        private DataTable EjecutarConsultaGuias(DateTime fechaInicio, DateTime fechaFin, string sucursal, string destino, string condicionExtra = null, Action<MySqlCommand> configureCondicion = null)
        {
            var dt = new DataTable();
            using (var conn = new MySqlConnection(_connectionString))
            {
                conn.Open();
                var query = new StringBuilder("SELECT * FROM guias WHERE FechaElaboracion BETWEEN @inicio AND @fin");

                if (!string.IsNullOrWhiteSpace(condicionExtra))
                {
                    query.Append(" AND ").Append(condicionExtra);
                }

                bool filtraSucursal = DebeFiltrarSucursal(sucursal);
                if (filtraSucursal)
                {
                    query.Append(" AND Sucursal = @sucursal");
                }

                bool filtraDestino = DebeFiltrarDestino(destino);
                if (filtraDestino)
                {
                    query.Append(" AND UPPER(TRIM(Destino)) = @destino");
                }

                query.Append(" ORDER BY FechaElaboracion DESC");

                using (var cmd = new MySqlCommand(query.ToString(), conn))
                {
                    cmd.Parameters.AddWithValue("@inicio", fechaInicio);
                    cmd.Parameters.AddWithValue("@fin", fechaFin);

                    if (filtraSucursal)
                    {
                        cmd.Parameters.AddWithValue("@sucursal", sucursal);
                    }

                    if (filtraDestino)
                    {
                        cmd.Parameters.AddWithValue("@destino", destino.Trim().ToUpperInvariant());
                    }

                    configureCondicion?.Invoke(cmd);

                    using (var adapter = new MySqlDataAdapter(cmd))
                    {
                        adapter.Fill(dt);
                    }
                }
            }

            return dt;
        }

        private (string EstatusGuia, string TipoCobro, string FolioInforme, string Factura, DateTime? FechaUltimaMilla) ObtenerCamposActualizables(string folioGuia, MySqlConnection conn)
        {
            string query = @"SELECT EstatusGuia, TipoCobro, FolioInforme, Factura, FechaUltimaMilla FROM guias WHERE FolioGuia = @folio";
            using (var cmd = new MySqlCommand(query, conn))
            {
                cmd.Parameters.AddWithValue("@folio", folioGuia);
                using (var reader = cmd.ExecuteReader())
                {
                    if (reader.Read())
                    {
                        return (
                            reader["EstatusGuia"]?.ToString(),
                            reader["TipoCobro"]?.ToString(),
                            reader["FolioInforme"]?.ToString(),
                            reader["Factura"]?.ToString(),
                            reader["FechaUltimaMilla"] == DBNull.Value ? (DateTime?)null : Convert.ToDateTime(reader["FechaUltimaMilla"])
                        );
                    }
                }
            }
            return (null, null, null, null, null);
        }

        public (int nuevos, int actualizados) InsertarDesdeRegistros(List<Dictionary<string, object>> registros)
        {
            int nuevos = 0;
            int actualizados = 0;

            using (var conn = new MySqlConnection(_connectionString))
            {
                conn.Open();

                foreach (var datos in registros)
                {
                    try
                    {
                        string folioGuia = datos.ContainsKey("folio") ? datos["folio"]?.ToString() : null;
                        if (string.IsNullOrEmpty(folioGuia))
                        {
                            continue;
                        }

                        string existeQuery = "SELECT COUNT(*) FROM guias WHERE FolioGuia = @folio";
                        using (var existeCmd = new MySqlCommand(existeQuery, conn))
                        {
                            existeCmd.Parameters.AddWithValue("@folio", folioGuia);
                            int existe = Convert.ToInt32(existeCmd.ExecuteScalar());

                            if (existe > 0)
                            {
                                var actuales = ObtenerCamposActualizables(folioGuia, conn);

                                string nuevoEstatus = datos.ContainsKey("estatus") ? datos["estatus"]?.ToString() : null;
                                string nuevoTipoCobro = datos.ContainsKey("tipoCobro") ? datos["tipoCobro"]?.ToString() : null;
                                string nuevoFolioInforme = datos.ContainsKey("folioInforme") ? datos["folioInforme"]?.ToString() : null;
                                string nuevaFactura = datos.ContainsKey("factura") ? datos["factura"]?.ToString() : null;
                                DateTime? nuevaFechaUltimaMilla = ObtenerFecha(datos, "fechaUltimaMilla");

                                if (actuales.EstatusGuia != nuevoEstatus ||
                                    actuales.TipoCobro != nuevoTipoCobro ||
                                    actuales.FolioInforme != nuevoFolioInforme ||
                                    actuales.Factura != nuevaFactura ||
                                    actuales.FechaUltimaMilla != nuevaFechaUltimaMilla)
                                {
                                    const string updateQuery = @"UPDATE guias SET 
                                            EstatusGuia = @EstatusGuia,
                                            TipoCobro = @TipoCobro,
                                            FolioInforme = @FolioInforme,
                                            Factura = @Factura,
                                            FechaUltimaMilla = @FechaUltimaMilla
                                            WHERE FolioGuia = @FolioGuia";

                                    using (var updateCmd = new MySqlCommand(updateQuery, conn))
                                    {
                                        updateCmd.Parameters.AddWithValue("@EstatusGuia", (object)nuevoEstatus ?? DBNull.Value);
                                        updateCmd.Parameters.AddWithValue("@TipoCobro", (object)nuevoTipoCobro ?? DBNull.Value);
                                        updateCmd.Parameters.AddWithValue("@FolioInforme", (object)nuevoFolioInforme ?? DBNull.Value);
                                        updateCmd.Parameters.AddWithValue("@Factura", (object)nuevaFactura ?? DBNull.Value);
                                        updateCmd.Parameters.AddWithValue("@FechaUltimaMilla", (object)nuevaFechaUltimaMilla ?? DBNull.Value);
                                        updateCmd.Parameters.AddWithValue("@FolioGuia", folioGuia);
                                        updateCmd.ExecuteNonQuery();
                                        actualizados++;
                                    }
                                }
                            }
                            else
                            {
                                const string insertQuery = @"INSERT INTO guias (
                                        FechaElaboracion, HoraElaboracion, FolioGuia, EstatusGuia, Cliente, UbicacionActual, Origen, 
                                        Destino, TipoCobro, ZonaOperativaEntrega, TipoEntrega, FechaEntrega, HoraEntrega, Tracking, 
                                        Referencia, Subtotal, Total, Sucursal, FolioInforme, FolioEmbarque, UsuarioDocumento, 
                                        FechaCancelacion, UsuarioCancelacion, Remitente, Destinatario, Cajas, ValorDeclarado, 
                                        Observaciones, Factura, TimbradoSAT, FolioERP, TipoCobroInicial, FechaUltimaMilla, MotivoCancelacion
                                    ) VALUES (
                                        @FechaElaboracion, @HoraElaboracion, @FolioGuia, @EstatusGuia, @Cliente, @UbicacionActual, @Origen, 
                                        @Destino, @TipoCobro, @ZonaOperativaEntrega, @TipoEntrega, @FechaEntrega, @HoraEntrega, @Tracking, 
                                        @Referencia, @Subtotal, @Total, @Sucursal, @FolioInforme, @FolioEmbarque, @UsuarioDocumento, 
                                        @FechaCancelacion, @UsuarioCancelacion, @Remitente, @Destinatario, @Cajas, @ValorDeclarado, 
                                        @Observaciones, @Factura, @TimbradoSAT, @FolioERP, @TipoCobroInicial, @FechaUltimaMilla, @MotivoCancelacion
                                    )";

                                using (var cmd = new MySqlCommand(insertQuery, conn))
                                {
                                    DateTime? fechaElab = ObtenerFecha(datos, "fechaElab");
                                    TimeSpan? horaElab = ObtenerTiempo(datos, "horaElab");
                                    DateTime? fechaEntrega = ObtenerFecha(datos, "fechaEntrega");
                                    TimeSpan? horaEntrega = ObtenerTiempo(datos, "horaEntrega");
                                    DateTime? fechaCancel = ObtenerFecha(datos, "fechaCancel");
                                    DateTime? fechaUltimaMilla = ObtenerFecha(datos, "fechaUltimaMilla");

                                    cmd.Parameters.AddWithValue("@FechaElaboracion", (object)fechaElab ?? DBNull.Value);
                                    cmd.Parameters.AddWithValue("@HoraElaboracion", (object)horaElab ?? DBNull.Value);
                                    cmd.Parameters.AddWithValue("@FolioGuia", (object)datoValor(datos, "folio") ?? string.Empty);
                                    cmd.Parameters.AddWithValue("@EstatusGuia", (object)datoValor(datos, "estatus") ?? string.Empty);
                                    cmd.Parameters.AddWithValue("@Cliente", (object)datoValor(datos, "cliente") ?? string.Empty);
                                    cmd.Parameters.AddWithValue("@UbicacionActual", (object)datoValor(datos, "ubicacion") ?? string.Empty);
                                    cmd.Parameters.AddWithValue("@Origen", (object)datoValor(datos, "origen") ?? string.Empty);
                                    cmd.Parameters.AddWithValue("@Destino", (object)datoValor(datos, "destino") ?? string.Empty);
                                    cmd.Parameters.AddWithValue("@TipoCobro", (object)datoValor(datos, "tipoCobro") ?? string.Empty);
                                    cmd.Parameters.AddWithValue("@ZonaOperativaEntrega", (object)datoValor(datos, "zona") ?? string.Empty);
                                    cmd.Parameters.AddWithValue("@TipoEntrega", (object)datoValor(datos, "tipoEntrega") ?? string.Empty);
                                    cmd.Parameters.AddWithValue("@FechaEntrega", (object)fechaEntrega ?? DBNull.Value);
                                    cmd.Parameters.AddWithValue("@HoraEntrega", (object)horaEntrega ?? DBNull.Value);
                                    cmd.Parameters.AddWithValue("@Tracking", (object)datoValor(datos, "tracking") ?? string.Empty);
                                    cmd.Parameters.AddWithValue("@Referencia", (object)datoValor(datos, "referencia") ?? string.Empty);
                                    cmd.Parameters.AddWithValue("@Subtotal", ObtenerDecimal(datos, "subtotal"));
                                    cmd.Parameters.AddWithValue("@Total", ObtenerDecimal(datos, "total"));
                                    cmd.Parameters.AddWithValue("@Sucursal", (object)datoValor(datos, "sucursal") ?? string.Empty);
                                    cmd.Parameters.AddWithValue("@FolioInforme", (object)datoValor(datos, "folioInforme") ?? string.Empty);
                                    cmd.Parameters.AddWithValue("@FolioEmbarque", (object)datoValor(datos, "folioEmbarque") ?? string.Empty);
                                    cmd.Parameters.AddWithValue("@UsuarioDocumento", (object)datoValor(datos, "usuarioDoc") ?? string.Empty);
                                    cmd.Parameters.AddWithValue("@FechaCancelacion", (object)fechaCancel ?? DBNull.Value);
                                    cmd.Parameters.AddWithValue("@UsuarioCancelacion", (object)datoValor(datos, "usuarioCancel") ?? string.Empty);
                                    cmd.Parameters.AddWithValue("@Remitente", (object)datoValor(datos, "remitente") ?? string.Empty);
                                    cmd.Parameters.AddWithValue("@Destinatario", (object)datoValor(datos, "destinatario") ?? string.Empty);
                                    cmd.Parameters.AddWithValue("@Cajas", ObtenerEntero(datos, "cajas"));
                                    cmd.Parameters.AddWithValue("@ValorDeclarado", ObtenerDecimal(datos, "valorDeclarado"));
                                    cmd.Parameters.AddWithValue("@Observaciones", (object)datoValor(datos, "observaciones") ?? string.Empty);
                                    cmd.Parameters.AddWithValue("@Factura", (object)datoValor(datos, "factura") ?? string.Empty);
                                    cmd.Parameters.AddWithValue("@TimbradoSAT", (object)datoValor(datos, "timbradoSat") ?? string.Empty);
                                    cmd.Parameters.AddWithValue("@FolioERP", (object)datoValor(datos, "folioErp") ?? string.Empty);
                                    cmd.Parameters.AddWithValue("@TipoCobroInicial", (object)datoValor(datos, "tipoCobroInicial") ?? string.Empty);
                                    cmd.Parameters.AddWithValue("@FechaUltimaMilla", (object)fechaUltimaMilla ?? DBNull.Value);
                                    cmd.Parameters.AddWithValue("@MotivoCancelacion", (object)datoValor(datos, "motivoCancel") ?? string.Empty);
                                    cmd.ExecuteNonQuery();
                                    nuevos++;
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error al insertar/actualizar fila: {ex.Message}");
                    }
                }
            }

            return (nuevos, actualizados);
        }

        public DataTable ObtenerTodasGuias()
        {
            var dt = new DataTable();

            try
            {
                using (var conn = new MySqlConnection(_connectionString))
                {
                    conn.Open();
                    const string query = "SELECT * FROM Guias ORDER BY FechaCreacion DESC";
                    using (var cmd = new MySqlCommand(query, conn))
                    using (var adapter = new MySqlDataAdapter(cmd))
                    {
                        adapter.Fill(dt);
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error al obtener guías: {ex.Message}");
            }

            return dt;
        }

        public DataTable ObtenerTodasGuiasFiltrado(DateTime fechaInicio, DateTime fechaFin, string sucursal, string destino = "TODAS")
        {
            return EjecutarConsultaGuias(fechaInicio, fechaFin, sucursal, destino);
        }

        public DataTable ObtenerGuiasPorCobrar(DateTime fechaInicio, DateTime fechaFin, string sucursal, string destino = "TODAS")
        {
            return EjecutarConsultaGuias(
                fechaInicio,
                fechaFin,
                sucursal,
                destino,
                "UPPER(TipoCobro) = @tipoCobro",
                cmd => cmd.Parameters.AddWithValue("@tipoCobro", "POR COBRAR"));
        }

        public DataTable ObtenerGuiasConCredito(DateTime fechaInicio, DateTime fechaFin, string sucursal, string destino = "TODAS")
        {
            return EjecutarConsultaGuias(
                fechaInicio,
                fechaFin,
                sucursal,
                destino,
                "UPPER(TipoCobro) = @tipoCobroCredito",
                cmd => cmd.Parameters.AddWithValue("@tipoCobroCredito", "CRÉDITO"));
        }

        public DataTable ObtenerGuiasPagadas(DateTime fechaInicio, DateTime fechaFin, string sucursal, string destino = "TODAS")
        {
            const string condicion = "(UPPER(TipoCobro) = @tipoCobroPagado OR UPPER(TipoCobroInicial) = @tipoCobroPagado)";
            return EjecutarConsultaGuias(
                fechaInicio,
                fechaFin,
                sucursal,
                destino,
                condicion,
                cmd => cmd.Parameters.AddWithValue("@tipoCobroPagado", "PAGADO"));
        }

        private (int Cantidad, decimal Monto) ObtenerResumenPagos(
            DateTime fechaInicio,
            DateTime fechaFin,
            string sucursal,
            string destino,
            string columnaUbicacion,
            string columnaTipoCobro,
            string valorTipoCobro)
        {
            using (var conn = new MySqlConnection(_connectionString))
            {
                conn.Open();

                var query = new StringBuilder($@"SELECT COUNT(*) AS Cantidad, IFNULL(SUM(Total), 0) AS Monto
                                 FROM guias
                                 WHERE FechaElaboracion BETWEEN @inicio AND @fin
                                   AND UPPER(COALESCE({columnaTipoCobro}, '')) = @valorTipoCobro
                                   AND COALESCE(UPPER(EstatusGuia), '') <> 'CANCELADO'");

                bool filtraSucursal = DebeFiltrarSucursal(sucursal);
                if (filtraSucursal)
                {
                    query.Append($" AND UPPER(COALESCE({columnaUbicacion}, '')) = @sucursal");
                }

                bool filtraDestino = DebeFiltrarDestino(destino);
                if (filtraDestino)
                {
                    query.Append(" AND UPPER(COALESCE(Destino, '')) = @destino");
                }

                using (var cmd = new MySqlCommand(query.ToString(), conn))
                {
                    cmd.Parameters.AddWithValue("@inicio", fechaInicio);
                    cmd.Parameters.AddWithValue("@fin", fechaFin);
                    cmd.Parameters.AddWithValue("@valorTipoCobro", valorTipoCobro.Trim().ToUpperInvariant());

                    if (filtraSucursal)
                    {
                        cmd.Parameters.AddWithValue("@sucursal", sucursal.Trim().ToUpperInvariant());
                    }

                    if (filtraDestino)
                    {
                        cmd.Parameters.AddWithValue("@destino", destino.Trim().ToUpperInvariant());
                    }

                    using (var reader = cmd.ExecuteReader())
                    {
                        if (reader.Read())
                        {
                            int cantidad = reader["Cantidad"] == DBNull.Value ? 0 : Convert.ToInt32(reader["Cantidad"]);
                            decimal monto = reader["Monto"] == DBNull.Value ? 0m : Convert.ToDecimal(reader["Monto"]);
                            return (cantidad, monto);
                        }
                    }
                }
            }

            return (0, 0m);
        }

        public (int Cantidad, decimal Monto) ObtenerResumenPagosOrigen(DateTime fechaInicio, DateTime fechaFin, string sucursal, string destino = "TODAS")
        {
            return ObtenerResumenPagos(fechaInicio, fechaFin, sucursal, destino, "Sucursal", "TipoCobroInicial", "PAGADO");
        }

        public (int Cantidad, decimal Monto) ObtenerResumenPagosDestino(DateTime fechaInicio, DateTime fechaFin, string sucursal, string destino = "TODAS")
        {
            return ObtenerResumenPagos(fechaInicio, fechaFin, sucursal, destino, "Sucursal", "TipoCobro", "POR COBRAR");
        }

        public DataTable ObtenerSumaTotalPorSucursalConCanceladas(DateTime fechaInicio, DateTime fechaFin)
        {
            var dt = new DataTable();

            using (var conn = new MySqlConnection(_connectionString))
            {
                conn.Open();
                const string query = @"
                    SELECT 
                        Sucursal,
                        SUM(CASE WHEN UPPER(EstatusGuia) = 'CANCELADO' THEN Total ELSE 0 END) AS TotalCanceladas,
                        SUM(CASE WHEN UPPER(EstatusGuia) <> 'CANCELADO' THEN Total ELSE 0 END) AS TotalValidas
                    FROM guias
                    WHERE FechaElaboracion BETWEEN @inicio AND @fin
                    GROUP BY Sucursal
                    ORDER BY Sucursal";

                using (var cmd = new MySqlCommand(query, conn))
                {
                    cmd.Parameters.AddWithValue("@inicio", fechaInicio);
                    cmd.Parameters.AddWithValue("@fin", fechaFin);

                    using (var adapter = new MySqlDataAdapter(cmd))
                    {
                        adapter.Fill(dt);
                    }
                }
            }

            return dt;
        }

        public void InsertarDesdeDataTable(DataTable dt)
        {
            using (var conn = new MySqlConnection(_connectionString))
            {
                conn.Open();

                foreach (DataRow row in dt.Rows)
                {
                    try
                    {
                        string fechaHoraOriginal = row.Table.Columns.Contains("Fecha/Hora Elaboración") ? row["Fecha/Hora Elaboración"]?.ToString() ?? string.Empty : string.Empty;
                        DateTime? fechaElaboracion = null;
                        TimeSpan? horaElaboracion = null;

                        if (!string.IsNullOrEmpty(fechaHoraOriginal) && fechaHoraOriginal.Contains(" "))
                        {
                            var partes = fechaHoraOriginal.Split(' ');
                            if (DateTime.TryParse(partes[0], out DateTime fecha))
                            {
                                fechaElaboracion = fecha;
                            }

                            if (TimeSpan.TryParse(partes[1], out TimeSpan hora))
                            {
                                horaElaboracion = hora;
                            }
                        }

                        string fechaHoraEntrega = row.Table.Columns.Contains("Fecha y Hora de Entrega") ? row["Fecha y Hora de Entrega"]?.ToString() ?? string.Empty : string.Empty;
                        DateTime? fechaEntrega = null;
                        TimeSpan? horaEntrega = null;

                        if (!string.IsNullOrEmpty(fechaHoraEntrega) && fechaHoraEntrega.Contains(" "))
                        {
                            var partes = fechaHoraEntrega.Split(' ');
                            if (DateTime.TryParse(partes[0], out DateTime fecha))
                            {
                                fechaEntrega = fecha;
                            }

                            if (TimeSpan.TryParse(partes[1], out TimeSpan hora))
                            {
                                horaEntrega = hora;
                            }
                        }

                        const string insertQuery = @"INSERT INTO guias (
                            FechaElaboracion, HoraElaboracion, FolioGuia, EstatusGuia, Cliente, UbicacionActual, Origen, 
                            Destino, TipoCobro, ZonaOperativaEntrega, TipoEntrega, FechaEntrega, HoraEntrega, Tracking, 
                            Referencia, Subtotal, Total, Sucursal, FolioInforme, FolioEmbarque, UsuarioDocumento, 
                            FechaCancelacion, UsuarioCancelacion, Remitente, Destinatario, Cajas, ValorDeclarado, 
                            Observaciones, Factura, TimbradoSAT, FolioERP, TipoCobroInicial, FechaUltimaMilla, MotivoCancelacion
                        ) VALUES (
                            @FechaElaboracion, @HoraElaboracion, @FolioGuia, @EstatusGuia, @Cliente, @UbicacionActual, @Origen, 
                            @Destino, @TipoCobro, @ZonaOperativaEntrega, @TipoEntrega, @FechaEntrega, @HoraEntrega, @Tracking, 
                            @Referencia, @Subtotal, @Total, @Sucursal, @FolioInforme, @FolioEmbarque, @UsuarioDocumento, 
                            @FechaCancelacion, @UsuarioCancelacion, @Remitente, @Destinatario, @Cajas, @ValorDeclarado, 
                            @Observaciones, @Factura, @TimbradoSAT, @FolioERP, @TipoCobroInicial, @FechaUltimaMilla, @MotivoCancelacion
                        )";

                        using (var cmd = new MySqlCommand(insertQuery, conn))
                        {
                            cmd.Parameters.AddWithValue("@FechaElaboracion", fechaElaboracion ?? (object)DBNull.Value);
                            cmd.Parameters.AddWithValue("@HoraElaboracion", horaElaboracion ?? (object)DBNull.Value);
                            cmd.Parameters.AddWithValue("@FolioGuia", row.Table.Columns.Contains("Folio Guía") ? row["Folio Guía"]?.ToString() ?? string.Empty : string.Empty);
                            cmd.Parameters.AddWithValue("@EstatusGuia", row.Table.Columns.Contains("Estatus Guía") ? row["Estatus Guía"]?.ToString() ?? string.Empty : string.Empty);
                            cmd.Parameters.AddWithValue("@Cliente", row.Table.Columns.Contains("Cliente") ? row["Cliente"]?.ToString() ?? string.Empty : string.Empty);
                            cmd.Parameters.AddWithValue("@UbicacionActual", row.Table.Columns.Contains("Ubicación Actual") ? row["Ubicación Actual"]?.ToString() ?? string.Empty : string.Empty);
                            cmd.Parameters.AddWithValue("@Origen", row.Table.Columns.Contains("Origen") ? row["Origen"]?.ToString() ?? string.Empty : string.Empty);
                            cmd.Parameters.AddWithValue("@Destino", row.Table.Columns.Contains("Destino") ? row["Destino"]?.ToString() ?? string.Empty : string.Empty);
                            cmd.Parameters.AddWithValue("@TipoCobro", row.Table.Columns.Contains("Tipo cobro") ? row["Tipo cobro"]?.ToString() ?? string.Empty : string.Empty);
                            cmd.Parameters.AddWithValue("@ZonaOperativaEntrega", row.Table.Columns.Contains("Zona Operativa Entrega") ? row["Zona Operativa Entrega"]?.ToString() ?? string.Empty : string.Empty);
                            cmd.Parameters.AddWithValue("@TipoEntrega", row.Table.Columns.Contains("Tipo de entrega") ? row["Tipo de entrega"]?.ToString() ?? string.Empty : string.Empty);
                            cmd.Parameters.AddWithValue("@FechaEntrega", fechaEntrega ?? (object)DBNull.Value);
                            cmd.Parameters.AddWithValue("@HoraEntrega", horaEntrega ?? (object)DBNull.Value);
                            cmd.Parameters.AddWithValue("@Tracking", row.Table.Columns.Contains("Tracking") ? row["Tracking"]?.ToString() ?? string.Empty : string.Empty);
                            cmd.Parameters.AddWithValue("@Referencia", row.Table.Columns.Contains("Referencia") ? row["Referencia"]?.ToString() ?? string.Empty : string.Empty);
                            cmd.Parameters.AddWithValue("@Subtotal", decimal.TryParse(row.Table.Columns.Contains("Subtotal") ? row["Subtotal"]?.ToString() : "0", out decimal subtotal) ? subtotal : 0);
                            cmd.Parameters.AddWithValue("@Total", decimal.TryParse(row.Table.Columns.Contains("Total") ? row["Total"]?.ToString() : "0", out decimal total) ? total : 0);
                            cmd.Parameters.AddWithValue("@Sucursal", row.Table.Columns.Contains("Sucursal") ? row["Sucursal"]?.ToString() ?? string.Empty : string.Empty);
                            cmd.Parameters.AddWithValue("@FolioInforme", row.Table.Columns.Contains("Folio Informe") ? row["Folio Informe"]?.ToString() ?? string.Empty : string.Empty);
                            cmd.Parameters.AddWithValue("@FolioEmbarque", row.Table.Columns.Contains("Folio Embarque") ? row["Folio Embarque"]?.ToString() ?? string.Empty : string.Empty);
                            cmd.Parameters.AddWithValue("@UsuarioDocumento", row.Table.Columns.Contains("Usuario Documento") ? row["Usuario Documento"]?.ToString() ?? string.Empty : string.Empty);
                            cmd.Parameters.AddWithValue("@FechaCancelacion", DateTime.TryParse(row.Table.Columns.Contains("Fecha de Cancelación") ? row["Fecha de Cancelación"]?.ToString() : string.Empty, out DateTime fechaCancelacion) ? (object)fechaCancelacion : DBNull.Value);
                            cmd.Parameters.AddWithValue("@UsuarioCancelacion", row.Table.Columns.Contains("Usuario de Cancelación") ? row["Usuario de Cancelación"]?.ToString() ?? string.Empty : string.Empty);
                            cmd.Parameters.AddWithValue("@Remitente", row.Table.Columns.Contains("Remitente") ? row["Remitente"]?.ToString() ?? string.Empty : string.Empty);
                            cmd.Parameters.AddWithValue("@Destinatario", row.Table.Columns.Contains("Destinatario") ? row["Destinatario"]?.ToString() ?? string.Empty : string.Empty);
                            cmd.Parameters.AddWithValue("@Cajas", int.TryParse(row.Table.Columns.Contains("Cajas") ? row["Cajas"]?.ToString() : "0", out int cajas) ? cajas : 0);
                            cmd.Parameters.AddWithValue("@ValorDeclarado", decimal.TryParse(row.Table.Columns.Contains("Valor declarado") ? row["Valor declarado"]?.ToString() : "0", out decimal valorDeclarado) ? valorDeclarado : 0);
                            cmd.Parameters.AddWithValue("@Observaciones", row.Table.Columns.Contains("Observaciones") ? row["Observaciones"]?.ToString() ?? string.Empty : string.Empty);
                            cmd.Parameters.AddWithValue("@Factura", row.Table.Columns.Contains("Factura") ? row["Factura"]?.ToString() ?? string.Empty : string.Empty);
                            cmd.Parameters.AddWithValue("@TimbradoSAT", row.Table.Columns.Contains("Timbrado SAT") ? row["Timbrado SAT"]?.ToString() ?? string.Empty : string.Empty);
                            cmd.Parameters.AddWithValue("@FolioERP", row.Table.Columns.Contains("Folio ERP") ? row["Folio ERP"]?.ToString() ?? string.Empty : string.Empty);
                            cmd.Parameters.AddWithValue("@TipoCobroInicial", row.Table.Columns.Contains("Tipo de cobro inicial") ? row["Tipo de cobro inicial"]?.ToString() ?? string.Empty : string.Empty);
                            cmd.Parameters.AddWithValue("@FechaUltimaMilla", DateTime.TryParse(row.Table.Columns.Contains("Fecha última milla") ? row["Fecha última milla"]?.ToString() : string.Empty, out DateTime fechaUltimaMilla) ? (object)fechaUltimaMilla : DBNull.Value);
                            cmd.Parameters.AddWithValue("@MotivoCancelacion", row.Table.Columns.Contains("Motivo cancelación") ? row["Motivo cancelación"]?.ToString() ?? string.Empty : string.Empty);

                            cmd.ExecuteNonQuery();
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error al insertar fila: {ex.Message}");
                    }
                }
            }
        }

        public DataTable BuscarGuias(string criterio, string valor, out string error)
        {
            var dt = new DataTable();
            error = null;

            try
            {
                using (var conn = new MySqlConnection(_connectionString))
                {
                    conn.Open();

                    string query;
                    MySqlCommand cmd;

                    if (criterio == "FolioGuia" && valor.All(char.IsDigit))
                    {
                        query = $"SELECT * FROM Guias WHERE {criterio} LIKE @valor ORDER BY FechaElaboracion DESC";
                        cmd = new MySqlCommand(query, conn);
                        cmd.Parameters.AddWithValue("@valor", "%" + valor);
                    }
                    else
                    {
                        query = $"SELECT * FROM Guias WHERE {criterio} = @valor ORDER BY FechaElaboracion DESC";
                        cmd = new MySqlCommand(query, conn);
                        cmd.Parameters.AddWithValue("@valor", valor);
                    }

                    using (var adapter = new MySqlDataAdapter(cmd))
                    {
                        adapter.Fill(dt);
                    }
                }
            }
            catch (Exception ex)
            {
                error = $"Error al buscar: {ex.Message}\nConsulta: SELECT * FROM Guias WHERE {criterio} = '{valor}'";
                MessageBox.Show(error, "Error SQL/Conexión", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            if (dt.Rows.Count == 0 && error == null)
            {
                error = $"No se encontraron resultados para FolioGuia = '{valor}'.";
            }

            return dt;
        }


        public DataTable BuscarGuiasPorFactura(string factura, out string error)
        {
            var dt = new DataTable();
            error = null;
    
            try
            {
                using (var conn = new MySqlConnection(_connectionString))
                {       
                    conn.Open();
                    string query = "SELECT * FROM Guias WHERE Factura LIKE @factura ORDER BY FechaElaboracion DESC";

                    using (var cmd = new MySqlCommand(query, conn))
                    {
                        cmd.Parameters.AddWithValue("@factura", "%" + factura + "%");
                        using (var adapter = new MySqlDataAdapter(cmd))
                        {
                            adapter.Fill(dt);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                error = $"Error al buscar factura: {ex.Message}";
            }

            return dt;
        }


        // ==================== BOTÓN 4: ACTUALIZAR GUIAS ====================
        public bool ActualizarGuia(string folioGuia, Dictionary<string, object> datos)
        {
            try
            {
                using (MySqlConnection conn = new MySqlConnection(_connectionString))
                {
                    conn.Open();

                    string setClause = "";
                    foreach (var dato in datos)
                    {
                        if (setClause != "") setClause += ", ";
                        setClause += $"{dato.Key} = @{dato.Key}";
                    }

                    string query = $"UPDATE Guias SET {setClause} WHERE folio_guia = @folio";

                    using (MySqlCommand cmd = new MySqlCommand(query, conn))
                    {
                        cmd.Parameters.AddWithValue("@folio", folioGuia);

                        foreach (var dato in datos)
                        {
                            cmd.Parameters.AddWithValue($"@{dato.Key}", dato.Value ?? DBNull.Value);
                        }

                        int rows = cmd.ExecuteNonQuery();
                        return rows > 0;
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error al actualizar: {ex.Message}");
                return false;
            }
        }

        public DataTable GenerarReporte(string fechaInicio, string fechaFin)
        {
            DataTable dt = new DataTable();

            try
            {
                using (MySqlConnection conn = new MySqlConnection(_connectionString))
                {
                    conn.Open();

                    string query = @"SELECT 
                        fecha_elaboracion,
                        COUNT(*) as total_guias,
                        SUM(total) as monto_total,
                        estatus_guia,
                        destino
                        FROM Guias 
                        WHERE fecha_elaboracion BETWEEN @fechaInicio AND @fechaFin
                        GROUP BY fecha_elaboracion, estatus_guia, destino
                        ORDER BY fecha_elaboracion DESC";

                    using (MySqlCommand cmd = new MySqlCommand(query, conn))
                    {
                        cmd.Parameters.AddWithValue("@fechaInicio", fechaInicio);
                        cmd.Parameters.AddWithValue("@fechaFin", fechaFin);

                        using (MySqlDataAdapter adapter = new MySqlDataAdapter(cmd))
                        {
                            adapter.Fill(dt);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error al generar reporte: {ex.Message}");
            }

            return dt;
        }

        public bool EliminarGuia(string folioGuia)
        {
            try
            {
                using (MySqlConnection conn = new MySqlConnection(_connectionString))
                {
                    conn.Open();

                    string query = "DELETE FROM Guias WHERE folio_guia = @folio";

                    using (MySqlCommand cmd = new MySqlCommand(query, conn))
                    {
                        cmd.Parameters.AddWithValue("@folio", folioGuia);

                        int rowsAffected = cmd.ExecuteNonQuery();
                        return rowsAffected > 0;
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error al eliminar: {ex.Message}");
                return false;
            }
        }

        public Dictionary<string, int> ObtenerEstadisticas()
        {
            var estadisticas = new Dictionary<string, int>();

            try
            {
                using (MySqlConnection conn = new MySqlConnection(_connectionString))
                {
                    conn.Open();

                    // Total de guías
                    string query1 = "SELECT COUNT(*) FROM Guias";
                    using (MySqlCommand cmd1 = new MySqlCommand(query1, conn))
                    {
                        estadisticas["Total Guías"] = Convert.ToInt32(cmd1.ExecuteScalar());
                    }

                    // Guías por cobrar
                    string query2 = "SELECT COUNT(*) FROM Guias WHERE tipo_cobro = 'Por Cobrar'";
                    using (MySqlCommand cmd2 = new MySqlCommand(query2, conn))
                    {
                        estadisticas["Por Cobrar"] = Convert.ToInt32(cmd2.ExecuteScalar());
                    }

                    // Guías en ruta
                    string query3 = "SELECT COUNT(*) FROM Guias WHERE estatus_guia = 'En Ruta'";

                    using (MySqlCommand cmd3 = new MySqlCommand(query3, conn))
                    {
                        estadisticas["En Ruta"] = Convert.ToInt32(cmd3.ExecuteScalar());
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error al obtener estadísticas: {ex.Message}");
            }

            return estadisticas;
        }

        public void ActualizarEstadoCarga(DateTime fecha, int cantidad)
        {
            FechaUltimaCarga = fecha;
            DocumentosCargadosHoy = cantidad;
        }

        public bool SeCargoHoy
        {
            get { return FechaUltimaCarga.Date == DateTime.Today && DocumentosCargadosHoy > 0; }
        }
        public DateTime FechaUltimaCarga { get; set; }
        public int DocumentosCargadosHoy { get; set; }

       public void CargarHistorialDesdeEstado(List<RegistroCarga> historial)
        {
            if (historial == null || historial.Count == 0)
            {
                FechaUltimaCarga = DateTime.MinValue;
                DocumentosCargadosHoy = 0;
                return;
            }

            var hoy = DateTime.Today;
            int totalHoy = 0;
            DateTime? ultimaCarga = null;

            foreach (var registro in historial)
            {
                DateTime fecha = registro.FechaHora;
                int cantidad = registro.DocumentosCargados;

                if (fecha.Date == hoy)
                {
                    totalHoy += cantidad;
                    if (!ultimaCarga.HasValue || fecha > ultimaCarga.Value)
                        ultimaCarga = fecha;
                }
            }

            FechaUltimaCarga = ultimaCarga ?? DateTime.MinValue;
            DocumentosCargadosHoy = totalHoy;
        }

        public void ExportarClientesACsv(string filePath)
        {
            using (var conn = new MySql.Data.MySqlClient.MySqlConnection(_connectionString))
            {
                conn.Open();
                string query = "SELECT IdCliente, NroCliente, TipoCliente, RFC, Nombre, NombreCorto, Sucursal, Activo, FechaCreacion, FechaActualizacion FROM clientes";
                using (var cmd = new MySql.Data.MySqlClient.MySqlCommand(query, conn))
                using (var reader = cmd.ExecuteReader())
                using (var writer = new StreamWriter(filePath, false, Encoding.UTF8))
                {
                    // Escribir encabezados
                    writer.WriteLine("IdCliente,NroCliente,TipoCliente,RFC,Nombre,NombreCorto,Sucursal,Activo,FechaCreacion,FechaActualizacion");
                    while (reader.Read())
                    {
                        var values = new List<string>();
                        for (int i = 0; i < reader.FieldCount; i++)
                        {
                            var val = reader[i]?.ToString().Replace("\"", "\"\"");
                            if (val != null && (val.Contains(",") || val.Contains("\"")))
                                val = $"\"{val}\"";
                            values.Add(val);
                        }
                        writer.WriteLine(string.Join(",", values));
                    }
                }
            }
        }

        public void ImportarClientesDesdeCsv(string filePath, Action<int, int> reportProgress = null)
        {
            using (var conn = new MySql.Data.MySqlClient.MySqlConnection(_connectionString))
            {
                conn.Open();
                using (var reader = new StreamReader(filePath, Encoding.UTF8))
                {
                    string headerLine = reader.ReadLine(); // Saltar encabezados
                    var allLines = new List<string>();
                    while (!reader.EndOfStream)
                    {
                        string line = reader.ReadLine();
                        if (!string.IsNullOrWhiteSpace(line))
                            allLines.Add(line);
                    }

                    int total = allLines.Count;
                    for (int i = 0; i < total; i++)
                    {
                        string line = allLines[i];
                        var fields = ParseCsvLine(line);

                        int idCliente = int.Parse(fields[0]);
                        string nroCliente = fields[1];
                        string tipoCliente = fields[2];
                        string rfc = fields[3];
                        string nombre = fields[4];
                        string nombreCorto = fields[5];
                        string sucursal = fields[6];
                        string activoStr = fields[7].Trim().ToUpper();
                        string activo = (activoStr == "SI") ? "SI" : "NO";

                        string existeQuery = "SELECT COUNT(*) FROM clientes WHERE IdCliente = @IdCliente";
                        using (var existeCmd = new MySql.Data.MySqlClient.MySqlCommand(existeQuery, conn))
                        {
                            existeCmd.Parameters.AddWithValue("@IdCliente", idCliente);
                            int existe = Convert.ToInt32(existeCmd.ExecuteScalar());

                            if (existe > 0)
                            {
                                string update = @"UPDATE clientes SET 
                                    NroCliente=@NroCliente, TipoCliente=@TipoCliente, RFC=@RFC, Nombre=@Nombre, NombreCorto=@NombreCorto, 
                                    Sucursal=@Sucursal, Activo=@Activo, FechaActualizacion=NOW()
                                    WHERE IdCliente=@IdCliente";
                                using (var cmd = new MySql.Data.MySqlClient.MySqlCommand(update, conn))
                                {
                                    cmd.Parameters.AddWithValue("@NroCliente", nroCliente);
                                    cmd.Parameters.AddWithValue("@TipoCliente", (object)tipoCliente ?? DBNull.Value);
                                    cmd.Parameters.AddWithValue("@RFC", (object)rfc ?? DBNull.Value);
                                    cmd.Parameters.AddWithValue("@Nombre", nombre);
                                    cmd.Parameters.AddWithValue("@NombreCorto", (object)nombreCorto ?? DBNull.Value);
                                    cmd.Parameters.AddWithValue("@Sucursal", (object)sucursal ?? DBNull.Value);
                                    cmd.Parameters.AddWithValue("@Activo", activo);
                                    cmd.Parameters.AddWithValue("@IdCliente", idCliente);
                                    cmd.ExecuteNonQuery();
                                }
                            }
                            else
                            {
                                string insert = @"INSERT INTO clientes 
                                    (IdCliente, NroCliente, TipoCliente, RFC, Nombre, NombreCorto, Sucursal, Activo) 
                                    VALUES (@IdCliente, @NroCliente, @TipoCliente, @RFC, @Nombre, @NombreCorto, @Sucursal, @Activo)";
                                using (var cmd = new MySql.Data.MySqlClient.MySqlCommand(insert, conn))
                                {
                                    cmd.Parameters.AddWithValue("@IdCliente", idCliente);
                                    cmd.Parameters.AddWithValue("@NroCliente", nroCliente);
                                    cmd.Parameters.AddWithValue("@TipoCliente", (object)tipoCliente ?? DBNull.Value);
                                    cmd.Parameters.AddWithValue("@RFC", (object)rfc ?? DBNull.Value);
                                    cmd.Parameters.AddWithValue("@Nombre", nombre);
                                    cmd.Parameters.AddWithValue("@NombreCorto", (object)nombreCorto ?? DBNull.Value);
                                    cmd.Parameters.AddWithValue("@Sucursal", (object)sucursal ?? DBNull.Value);
                                    cmd.Parameters.AddWithValue("@Activo", activo);
                                    cmd.ExecuteNonQuery();
                                }
                            }
                        }
                        reportProgress?.Invoke(i + 1, total);
                    }
                }
            }
        }

        // Utilidad para parsear líneas CSV (soporta comillas y comas)
        private static List<string> ParseCsvLine(string line)
        {
            var result = new List<string>();
            bool inQuotes = false;
            StringBuilder value = new StringBuilder();
            for (int i = 0; i < line.Length; i++)
            {
                char c = line[i];
                if (c == '\"')
                {
                    if (inQuotes && i + 1 < line.Length && line[i + 1] == '\"')
                    {
                        value.Append('\"');
                        i++;
                    }
                    else
                    {
                        inQuotes = !inQuotes;
                    }
                }
                else if (c == ',' && !inQuotes)
                {
                    result.Add(value.ToString());
                    value.Clear();
                }
                else
                {
                    value.Append(c);
                }
            }
            result.Add(value.ToString());
            return result;
        }

        private static void NormalizarColumnaActivo(DataTable dt)
        {
            if (dt == null || !dt.Columns.Contains("Activo") || dt.Columns["Activo"].DataType == typeof(bool))
            {
                return;
            }

            var colBool = new DataColumn("ActivoBool", typeof(bool));
            dt.Columns.Add(colBool);

            foreach (DataRow row in dt.Rows)
            {
                string valor = row["Activo"]?.ToString().Trim().ToUpperInvariant();
                row["ActivoBool"] = valor == "SI" || valor == "1" || valor == "TRUE";
            }

            dt.Columns.Remove("Activo");
            colBool.ColumnName = "Activo";
        }

        public DataTable ObtenerTodosClientes()
        {
            DataTable dt = new DataTable();
            using (var conn = new MySqlConnection(_connectionString))
            {
                conn.Open();
                const string query = "SELECT IdCliente, NroCliente, TipoCliente, RFC, Nombre, NombreCorto, Sucursal, Activo, FechaCreacion, FechaActualizacion FROM clientes";
                using (var cmd = new MySqlCommand(query, conn))
                using (var adapter = new MySqlDataAdapter(cmd))
                {
                    adapter.Fill(dt);
                }
            }

            NormalizarColumnaActivo(dt);
            return dt;
        }

        public DataTable BuscarClientesPorTexto(string texto)
        {
            if (string.IsNullOrWhiteSpace(texto))
            {
                return ObtenerTodosClientes();
            }

            var dt = new DataTable();
            using (var conn = new MySqlConnection(_connectionString))
            {
                conn.Open();
                const string query = @"
                    SELECT IdCliente, NroCliente, TipoCliente, RFC, Nombre, NombreCorto, Sucursal, Activo, FechaCreacion, FechaActualizacion
                    FROM clientes
                    WHERE 
                        LOWER(NroCliente) LIKE @pattern OR
                        LOWER(Nombre) LIKE @pattern OR
                        LOWER(NombreCorto) LIKE @pattern OR
                        LOWER(RFC) LIKE @pattern OR
                        LOWER(Sucursal) LIKE @pattern
                    ORDER BY Nombre";
                using (var cmd = new MySqlCommand(query, conn))
                {
                    string pattern = $"%{texto.Trim().ToLowerInvariant()}%";
                    cmd.Parameters.AddWithValue("@pattern", pattern);

                    using (var adapter = new MySqlDataAdapter(cmd))
                    {
                        adapter.Fill(dt);
                    }
                }
            }

            NormalizarColumnaActivo(dt);
            return dt;
        }

        private static DateTime? ObtenerFecha(Dictionary<string, object> origen, string clave)
        {
            object valor;
            if (!origen.TryGetValue(clave, out valor) || valor == null)
                return null;

            if (valor is DateTime fecha)
                return fecha;

            DateTime fechaConvertida;
            return DateTime.TryParse(valor.ToString(), out fechaConvertida) ? fechaConvertida : (DateTime?)null;
        }

        private static TimeSpan? ObtenerTiempo(Dictionary<string, object> origen, string clave)
        {
            object valor;
            if (!origen.TryGetValue(clave, out valor) || valor == null)
                return null;

            if (valor is TimeSpan tiempo)
                return tiempo;

            if (valor is DateTime fecha)
                return fecha.TimeOfDay;

            TimeSpan tiempoConvertido;
            if (TimeSpan.TryParse(valor.ToString(), out tiempoConvertido))
                return tiempoConvertido;

            DateTime fechaConvertida;
            return DateTime.TryParse(valor.ToString(), out fechaConvertida) ? fechaConvertida.TimeOfDay : (TimeSpan?)null;
        }

        private static decimal ObtenerDecimal(Dictionary<string, object> origen, string clave)
        {
            object valor;
            if (!origen.TryGetValue(clave, out valor) || valor == null)
                return 0m;

            if (valor is decimal dec)
                return dec;

            if (valor is double dbl)
                return Convert.ToDecimal(dbl);

            decimal decimalConvertido;
            return decimal.TryParse(valor.ToString(), NumberStyles.Any, CultureInfo.InvariantCulture, out decimalConvertido) ||
                   decimal.TryParse(valor.ToString(), NumberStyles.Any, CultureInfo.CurrentCulture, out decimalConvertido)
                ? decimalConvertido
                : 0m;
        }

        private static int ObtenerEntero(Dictionary<string, object> origen, string clave)
        {
            object valor;
            if (!origen.TryGetValue(clave, out valor) || valor == null)
                return 0;

            if (valor is int entero)
                return entero;

            if (valor is long largo)
                return (int)largo;

            if (valor is double dbl)
                return (int)Math.Round(dbl);

            int enteroConvertido;
            return int.TryParse(valor.ToString(), NumberStyles.Any, CultureInfo.InvariantCulture, out enteroConvertido) ||
                   int.TryParse(valor.ToString(), NumberStyles.Any, CultureInfo.CurrentCulture, out enteroConvertido)
                ? enteroConvertido
                : 0;
        }

        private static object datoValor(Dictionary<string, object> origen, string clave)
        {
            object valor;
            return origen.TryGetValue(clave, out valor) ? valor : null;
        }

        public void ImportarDesdeExcel(string filePath, DataGridView dataGridView, Action<int, int> reportProgress)
        {
            try
            {
                System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

                DataTable table;
                using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    var conf = new ExcelDataSetConfiguration
                    {
                        ConfigureDataTable = _ => new ExcelDataTableConfiguration
                        {
                            UseHeaderRow = true
                        }
                    };

                    var result = reader.AsDataSet(conf);
                    if (result.Tables.Count == 0)
                    {
                        MessageBox.Show("El archivo no contiene hojas con datos.", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }

                    table = result.Tables[0];
                }

                if (table.Rows.Count == 0)
                {
                    MessageBox.Show("El archivo no contiene registros para importar.", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                for (int i = 0; i < table.Rows.Count; i++)
                {
                    reportProgress?.Invoke(i + 1, table.Rows.Count);
                }

                dataGridView.Invoke((Action)(() =>
                {
                    dataGridView.DataSource = table;
                }));

                InsertarDesdeDataTable(table);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error al importar: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public bool TruncarGuias()
        {
            try
            {
                using (var conn = new MySqlConnection(_connectionString))
                {
                    conn.Open();
                    using (var cmd = new MySqlCommand("TRUNCATE TABLE guias", conn))
                    {
                        cmd.ExecuteNonQuery();
                    }
                }

                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"No se pudieron eliminar las guías: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
        }

        private static string ConstruirPatronCliente(string cliente)
        {
            return $"%{cliente.Trim().ToUpperInvariant()}%";
        }

        public DataTable ObtenerGuiasPorCliente(string cliente, DateTime fechaInicio, DateTime fechaFin)
        {
            if (string.IsNullOrWhiteSpace(cliente))
            {
                throw new ArgumentException("El cliente es obligatorio.", nameof(cliente));
            }

            var dt = new DataTable();
            using (var conn = new MySqlConnection(_connectionString))
            {
                conn.Open();
                const string query = @"SELECT * FROM guias
                                       WHERE FechaElaboracion BETWEEN @inicio AND @fin
                                         AND UPPER(COALESCE(Cliente, '')) LIKE @cliente
                                       ORDER BY FechaElaboracion DESC";
                using (var cmd = new MySqlCommand(query, conn))
                {
                    cmd.Parameters.AddWithValue("@inicio", fechaInicio);
                    cmd.Parameters.AddWithValue("@fin", fechaFin);
                    cmd.Parameters.AddWithValue("@cliente", ConstruirPatronCliente(cliente));

                    using (var adapter = new MySqlDataAdapter(cmd))
                    {
                        adapter.Fill(dt);
                    }
                }
            }

            return dt;
        }

        public ClienteEstadisticas ObtenerEstadisticasCliente(string cliente, DateTime fechaInicio, DateTime fechaFin)
        {
            if (string.IsNullOrWhiteSpace(cliente))
            {
                throw new ArgumentException("El cliente es obligatorio.", nameof(cliente));
            }

            var resultado = new ClienteEstadisticas();
            using (var conn = new MySqlConnection(_connectionString))
            {
                conn.Open();

                const string resumenQuery = @"
                    SELECT
                        SUM(CASE WHEN UPPER(TipoCobro) = 'POR COBRAR'
                                 AND COALESCE(UPPER(EstatusGuia), '') <> 'CANCELADO' THEN 1 ELSE 0 END) AS GuiasPorCobrar,
                        IFNULL(SUM(CASE WHEN UPPER(TipoCobro) = 'POR COBRAR'
                                        AND COALESCE(UPPER(EstatusGuia), '') <> 'CANCELADO' THEN Total ELSE 0 END), 0) AS MontoPorCobrar,
                        SUM(CASE WHEN UPPER(TipoCobroInicial) = 'PAGADO'
                                 AND COALESCE(UPPER(EstatusGuia), '') <> 'CANCELADO' THEN 1 ELSE 0 END) AS GuiasPagadasOrigen,
                        IFNULL(SUM(CASE WHEN UPPER(TipoCobroInicial) = 'PAGADO'
                                        AND COALESCE(UPPER(EstatusGuia), '') <> 'CANCELADO' THEN Total ELSE 0 END), 0) AS MontoPagadasOrigen,
                        SUM(CASE WHEN UPPER(TipoCobro) = 'PAGADO'
                                 AND COALESCE(UPPER(EstatusGuia), '') <> 'CANCELADO' THEN 1 ELSE 0 END) AS GuiasPagadasDestino,
                        IFNULL(SUM(CASE WHEN UPPER(TipoCobro) = 'PAGADO'
                                        AND COALESCE(UPPER(EstatusGuia), '') <> 'CANCELADO' THEN Total ELSE 0 END), 0) AS MontoPagadasDestino,
                        SUM(CASE WHEN UPPER(EstatusGuia) = 'CANCELADO' THEN 1 ELSE 0 END) AS GuiasCanceladas,
                        IFNULL(SUM(CASE WHEN UPPER(EstatusGuia) = 'CANCELADO' THEN Total ELSE 0 END), 0) AS MontoCanceladas,
                        IFNULL(SUM(CASE WHEN COALESCE(UPPER(EstatusGuia), '') <> 'CANCELADO' THEN IFNULL(Cajas, 0) ELSE 0 END), 0) AS PaquetesEnviados
                    FROM guias
                    WHERE FechaElaboracion BETWEEN @inicio AND @fin
                      AND UPPER(COALESCE(Cliente, '')) LIKE @cliente;";

                using (var cmd = new MySqlCommand(resumenQuery, conn))
                {
                    cmd.Parameters.AddWithValue("@inicio", fechaInicio);
                    cmd.Parameters.AddWithValue("@fin", fechaFin);
                    cmd.Parameters.AddWithValue("@cliente", ConstruirPatronCliente(cliente));

                    using (var reader = cmd.ExecuteReader())
                    {
                        if (reader.Read())
                        {
                            resultado.GuiasPorCobrar = reader["GuiasPorCobrar"] == DBNull.Value ? 0 : Convert.ToInt32(reader["GuiasPorCobrar"]);
                            resultado.MontoPorCobrar = reader["MontoPorCobrar"] == DBNull.Value ? 0m : Convert.ToDecimal(reader["MontoPorCobrar"]);
                            resultado.GuiasPagadasOrigen = reader["GuiasPagadasOrigen"] == DBNull.Value ? 0 : Convert.ToInt32(reader["GuiasPagadasOrigen"]);
                            resultado.MontoPagadasOrigen = reader["MontoPagadasOrigen"] == DBNull.Value ? 0m : Convert.ToDecimal(reader["MontoPagadasOrigen"]);
                            resultado.GuiasPagadasDestino = reader["GuiasPagadasDestino"] == DBNull.Value ? 0 : Convert.ToInt32(reader["GuiasPagadasDestino"]);
                            resultado.MontoPagadasDestino = reader["MontoPagadasDestino"] == DBNull.Value ? 0m : Convert.ToDecimal(reader["MontoPagadasDestino"]);
                            resultado.GuiasCanceladas = reader["GuiasCanceladas"] == DBNull.Value ? 0 : Convert.ToInt32(reader["GuiasCanceladas"]);
                            resultado.MontoCanceladas = reader["MontoCanceladas"] == DBNull.Value ? 0m : Convert.ToDecimal(reader["MontoCanceladas"]);
                            resultado.PaquetesEnviados = reader["PaquetesEnviados"] == DBNull.Value ? 0 : Convert.ToInt32(reader["PaquetesEnviados"]);
                        }
                    }
                }

                const string destinosQuery = @"
                    SELECT 
                        CASE WHEN TRIM(IFNULL(Destino, '')) = '' THEN 'SIN DESTINO' ELSE TRIM(Destino) END AS Destino,
                        IFNULL(SUM(IFNULL(Cajas, 0)), 0) AS TotalCajas
                    FROM guias
                    WHERE FechaElaboracion BETWEEN @inicio AND @fin
                      AND UPPER(COALESCE(Cliente, '')) LIKE @cliente
                      AND COALESCE(UPPER(EstatusGuia), '') <> 'CANCELADO'
                    GROUP BY Destino
                    ORDER BY TotalCajas DESC, Destino;"; 

                using (var cmd = new MySqlCommand(destinosQuery, conn))
                {
                    cmd.Parameters.AddWithValue("@inicio", fechaInicio);
                    cmd.Parameters.AddWithValue("@fin", fechaFin);
                    cmd.Parameters.AddWithValue("@cliente", ConstruirPatronCliente(cliente));

                    using (var reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            string destino = reader["Destino"]?.ToString() ?? "SIN DESTINO";
                            int totalCajas = reader["TotalCajas"] == DBNull.Value ? 0 : Convert.ToInt32(reader["TotalCajas"]);
                            resultado.Destinos.Add((destino, totalCajas));
                        }
                    }
                }
            }

            return resultado;
        }

        public void ImportarUsuariosDesdeExcel(string filePath, Action<int, int> reportProgress = null)
        {
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

            DataTable table;
            using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
            using (var reader = ExcelReaderFactory.CreateReader(stream))
            {
                var conf = new ExcelDataSetConfiguration
                {
                    ConfigureDataTable = _ => new ExcelDataTableConfiguration
                    {
                        UseHeaderRow = true
                    }
                };

                var result = reader.AsDataSet(conf);
                if (result.Tables.Count == 0)
                {
                    return;
                }

                table = result.Tables[0];
            }

            if (table.Rows.Count == 0)
            {
                return;
            }

            using (var conn = new MySqlConnection(_connectionString))
            {
                conn.Open();

                int total = table.Rows.Count;
                for (int i = 0; i < total; i++)
                {
                    var row = table.Rows[i];

                    object GetCell(params string[] columnNames)
                    {
                        foreach (var name in columnNames)
                        {
                            if (!string.IsNullOrWhiteSpace(name) && table.Columns.Contains(name))
                            {
                                return row[name];
                            }
                        }
                        return null;
                    }

                    string usuario = Convert.ToString(GetCell("Usuario", "usuario"))?.Trim();
                    if (string.IsNullOrWhiteSpace(usuario))
                    {
                        reportProgress?.Invoke(i + 1, total);
                        continue;
                    }

                    string nombre = Convert.ToString(GetCell("Nombre", "nombre"))?.Trim();
                    string tipoUsuario = Convert.ToString(GetCell("Tipo Usuario", "TipoUsuario", "tipo_usuario"))?.Trim().ToUpperInvariant();
                    string sucursal = Convert.ToString(GetCell("Sucursal", "sucursal"))?.Trim();

                    object activoRaw = GetCell("Activo", "activo");
                    bool activo = false;
                    if (activoRaw is bool boolValue)
                    {
                        activo = boolValue;
                    }
                    else if (activoRaw is double doubleValue)
                    {
                        activo = Math.Abs(doubleValue) > 0.0001;
                    }
                    else if (activoRaw != null)
                    {
                        var texto = activoRaw.ToString().Trim();
                        activo = texto.Equals("SI", StringComparison.OrdinalIgnoreCase) ||
                                 texto.Equals("TRUE", StringComparison.OrdinalIgnoreCase) ||
                                 texto == "1";
                    }

                    object fechaRaw = GetCell("Fecha último inicio de sesión", "Fecha ultimo inicio de sesion",
                        "Fecha último inicio de sesion", "Fecha ultimo inicio de sesión");
                    DateTime? fechaUltimoInicio = null;
                    if (fechaRaw is DateTime fecha)
                    {
                        fechaUltimoInicio = fecha.Date;
                    }
                    else if (fechaRaw is double fechaOa)
                    {
                        fechaUltimoInicio = DateTime.FromOADate(fechaOa).Date;
                    }
                    else if (fechaRaw != null)
                    {
                        if (DateTime.TryParse(fechaRaw.ToString(), CultureInfo.InvariantCulture, DateTimeStyles.None, out var fechaInv) ||
                            DateTime.TryParse(fechaRaw.ToString(), CultureInfo.CurrentCulture, DateTimeStyles.None, out fechaInv))
                        {
                            fechaUltimoInicio = fechaInv.Date;
                        }
                    }

                    if (string.IsNullOrWhiteSpace(tipoUsuario) || (tipoUsuario != "ADMINISTRADOR" && tipoUsuario != "USUARIO"))
                    {
                        tipoUsuario = "USUARIO";
                    }

                    const string existeQuery = "SELECT COUNT(*) FROM usuarios WHERE usuario = @usuario";
                    using (var existeCmd = new MySqlCommand(existeQuery, conn))
                    {
                        existeCmd.Parameters.AddWithValue("@usuario", usuario);
                        int existe = Convert.ToInt32(existeCmd.ExecuteScalar());

                        if (existe > 0)
                        {
                            const string update = @"
                                UPDATE usuarios SET
                                    nombre = @nombre,
                                    tipo_usuario = @tipo_usuario,
                                    activo = @activo,
                                    sucursal = @sucursal,
                                    fecha_ultimo_inicio_sesion = @fecha_ultimo_inicio_sesion,
                                    updated_at = NOW()
                                WHERE usuario = @usuario";
                            using (var cmd = new MySqlCommand(update, conn))
                            {
                                cmd.Parameters.AddWithValue("@usuario", usuario);
                                cmd.Parameters.AddWithValue("@nombre", (object)nombre ?? DBNull.Value);
                                cmd.Parameters.AddWithValue("@tipo_usuario", tipoUsuario);
                                cmd.Parameters.AddWithValue("@activo", activo);
                                cmd.Parameters.AddWithValue("@sucursal", (object)sucursal ?? DBNull.Value);
                                cmd.Parameters.AddWithValue("@fecha_ultimo_inicio_sesion", (object)fechaUltimoInicio ?? DBNull.Value);
                                cmd.ExecuteNonQuery();
                            }
                        }
                        else
                        {
                            const string insert = @"
                                INSERT INTO usuarios
                                    (usuario, nombre, tipo_usuario, activo, sucursal, fecha_ultimo_inicio_sesion)
                                VALUES
                                    (@usuario, @nombre, @tipo_usuario, @activo, @sucursal, @fecha_ultimo_inicio_sesion)";
                            using (var cmd = new MySqlCommand(insert, conn))
                            {
                                cmd.Parameters.AddWithValue("@usuario", usuario);
                                cmd.Parameters.AddWithValue("@nombre", (object)nombre ?? DBNull.Value);
                                cmd.Parameters.AddWithValue("@tipo_usuario", tipoUsuario);
                                cmd.Parameters.AddWithValue("@activo", activo);
                                cmd.Parameters.AddWithValue("@sucursal", (object)sucursal ?? DBNull.Value);
                                cmd.Parameters.AddWithValue("@fecha_ultimo_inicio_sesion", (object)fechaUltimoInicio ?? DBNull.Value);
                                cmd.ExecuteNonQuery();
                            }
                        }
                    }

                    reportProgress?.Invoke(i + 1, total);
                }
            }
        }

        public DataTable ObtenerTodosUsuarios()
        {
            DataTable dt = new DataTable();
            using (var conn = new MySqlConnection(_connectionString))
            {
                conn.Open();
                const string query = @"
                    SELECT 
                        usuario AS Usuario,
                        nombre AS Nombre,
                        tipo_usuario AS `Tipo Usuario`,
                        activo AS Activo,
                        sucursal AS Sucursal,
                        fecha_ultimo_inicio_sesion AS `Fecha ultimo inicio de sesion`,
                        created_at AS Creado,
                        updated_at AS Actualizado
                    FROM usuarios
                    ORDER BY usuario";
                using (var cmd = new MySqlCommand(query, conn))
                using (var adapter = new MySqlDataAdapter(cmd))
                {
                    adapter.Fill(dt);
                }
            }

            NormalizarColumnaActivo(dt);
            return dt;
        }

        public DataTable ObtenerGuiasCanceladasPorSucursal(DateTime fechaInicio, DateTime fechaFin, string origen = "TODAS")
        {
            var dt = new DataTable();
            using (var conn = new MySqlConnection(_connectionString))
            {
                conn.Open();
                string token = string.IsNullOrWhiteSpace(origen) || origen.Equals("TODAS", StringComparison.OrdinalIgnoreCase)
                    ? null
                    : origen.Split(',')[0].Trim().ToUpperInvariant();
                string patron = token == null ? null : $"%{token}%";

                var query = new StringBuilder(@"
            SELECT *
            FROM guias
            WHERE FechaElaboracion BETWEEN @inicio AND @fin
              AND UPPER(COALESCE(EstatusGuia,'')) = 'CANCELADO'");

                if (patron != null)
                {
                    query.Append(" AND UPPER(COALESCE(Origen,'')) LIKE @origen");
                }

                query.Append(" ORDER BY FechaElaboracion DESC");

                using (var cmd = new MySqlCommand(query.ToString(), conn))
                {
                    cmd.Parameters.AddWithValue("@inicio", fechaInicio);
                    cmd.Parameters.AddWithValue("@fin", fechaFin);
                    if (patron != null)
                    {
                        cmd.Parameters.AddWithValue("@origen", patron);
                    }

                    using (var adapter = new MySqlDataAdapter(cmd))
                    {
                        adapter.Fill(dt);
                    }
                }
            }
            return dt;
        }

        public DataTable ObtenerResumenCancelacionesPorUsuario(DateTime fechaInicio, DateTime fechaFin, string origen = "TODAS")
        {
            var dt = new DataTable();
            using (var conn = new MySqlConnection(_connectionString))
            {
                conn.Open();
                string token = string.IsNullOrWhiteSpace(origen) || origen.Equals("TODAS", StringComparison.OrdinalIgnoreCase)
                    ? null
                    : origen.Split(',')[0].Trim().ToUpperInvariant();
                string patron = token == null ? null : $"%{token}%";

                var sb = new StringBuilder(@"
            SELECT COALESCE(UsuarioDocumento, 'SIN') AS UsuarioDocumento,
                   COUNT(*) AS TotalGuias,
                   SUM(CASE WHEN UPPER(COALESCE(EstatusGuia,'')) = 'CANCELADO' THEN 1 ELSE 0 END) AS Canceladas
            FROM guias
            WHERE FechaElaboracion BETWEEN @inicio AND @fin");

                if (patron != null)
                {
                    sb.Append(" AND UPPER(COALESCE(Origen,'')) LIKE @origen");
                }

                sb.Append(" GROUP BY COALESCE(UsuarioDocumento, 'SIN')")
                  .Append(" HAVING Canceladas > 0")
                  .Append(" ORDER BY Canceladas DESC");

                using (var cmd = new MySqlCommand(sb.ToString(), conn))
                {
                    cmd.Parameters.AddWithValue("@inicio", fechaInicio);
                    cmd.Parameters.AddWithValue("@fin", fechaFin);
                    if (patron != null)
                    {
                        cmd.Parameters.AddWithValue("@origen", patron);
                    }

                    using (var adapter = new MySqlDataAdapter(cmd))
                    {
                        adapter.Fill(dt);
                    }
                }
            }
            return dt;
        }

        public int ContarGuiasTotales(DateTime fechaInicio, DateTime fechaFin, string origen = "TODAS")
        {
            using (var conn = new MySqlConnection(_connectionString))
            {
                conn.Open();
                string token = string.IsNullOrWhiteSpace(origen) || origen.Equals("TODAS", StringComparison.OrdinalIgnoreCase)
                    ? null
                    : origen.Split(',')[0].Trim().ToUpperInvariant();
                string patron = token == null ? null : $"%{token}%";

                var sb = new StringBuilder(@"
            SELECT COUNT(*) 
            FROM guias
            WHERE FechaElaboracion BETWEEN @inicio AND @fin");

                if (patron != null)
                {
                    sb.Append(" AND UPPER(COALESCE(Origen,'')) LIKE @origen");
                }

                using (var cmd = new MySqlCommand(sb.ToString(), conn))
                {
                    cmd.Parameters.AddWithValue("@inicio", fechaInicio);
                    cmd.Parameters.AddWithValue("@fin", fechaFin);
                    if (patron != null)
                    {
                        cmd.Parameters.AddWithValue("@origen", patron);
                    }

                    object result = cmd.ExecuteScalar();
                    return (result == null || result == DBNull.Value) ? 0 : Convert.ToInt32(result);
                }
            }
        }

        public DataTable ObtenerGuiasPorEstatus(DateTime fechaInicio, DateTime fechaFin, IList<string> estatus)
        {
            var dt = new DataTable();
            using (var conn = new MySqlConnection(_connectionString))
            {
                conn.Open();

                var sql = new StringBuilder(@"
                    SELECT *
                    FROM guias
                    WHERE FechaElaboracion BETWEEN @inicio AND @fin");

                if (estatus != null && estatus.Count > 0)
                {
                    var placeholders = estatus.Select((s, i) => "@e" + i).ToArray();
                    sql.Append(" AND UPPER(COALESCE(EstatusGuia,'')) IN (");
                    sql.Append(string.Join(",", placeholders));
                    sql.Append(")");
                }

                sql.Append(" ORDER BY FechaElaboracion DESC");

                using (var cmd = new MySqlCommand(sql.ToString(), conn))
                {
                    cmd.Parameters.AddWithValue("@inicio", fechaInicio);
                    cmd.Parameters.AddWithValue("@fin", fechaFin);

                    if (estatus != null)
                    {
                        for (int i = 0; i < estatus.Count; i++)
                        {
                            cmd.Parameters.AddWithValue("@e" + i, estatus[i].ToUpperInvariant());
                        }
                    }

                    using (var adapter = new MySqlDataAdapter(cmd))
                    {
                        adapter.Fill(dt);
                    }
                }
            }
            return dt;
        }

        public bool ActualizarObservaciones(string folioGuia, string observaciones)
        {
            if (string.IsNullOrWhiteSpace(folioGuia))
            {
                return false;
            }

            using (var conn = new MySqlConnection(_connectionString))
            {
                conn.Open();
                const string sql = @"UPDATE guias 
                                     SET Observaciones = @obs, FechaActualizacion = NOW() 
                                     WHERE FolioGuia = @folio";
                using (var cmd = new MySqlCommand(sql, conn))
                {
                    cmd.Parameters.AddWithValue("@obs", (object)observaciones ?? DBNull.Value);
                    cmd.Parameters.AddWithValue("@folio", folioGuia);
                    int rows = cmd.ExecuteNonQuery();
                    return rows > 0;
                }
            }
        }

        // Agrega debajo de ObtenerGuiasPorEstatus
        public (int Documentado, int Pendiente) ObtenerConteoSeguimientoPorPrefijo(DateTime fechaCorte, string prefijoFolio)
        {
            try
            {
                using (var conn = new MySqlConnection(_connectionString))
                {
                    conn.Open();
                    const string sql = @"
                        SELECT 
                            SUM(CASE WHEN UPPER(COALESCE(EstatusGuia,'')) = 'DOCUMENTADO' THEN 1 ELSE 0 END) AS Doc,
                            SUM(CASE WHEN UPPER(COALESCE(EstatusGuia,'')) = 'PENDIENTE' THEN 1 ELSE 0 END) AS Pen
                        FROM guias
                        WHERE FolioGuia LIKE @prefijo
                          AND FechaElaboracion IS NOT NULL
                          AND FechaElaboracion <= @corte";
                    using (var cmd = new MySqlCommand(sql, conn))
                    {
                        cmd.Parameters.AddWithValue("@prefijo", prefijoFolio + "%");
                        cmd.Parameters.AddWithValue("@corte", fechaCorte);
                        using (var rd = cmd.ExecuteReader())
                        {
                            if (rd.Read())
                            {
                                int doc = rd["Doc"] == DBNull.Value ? 0 : Convert.ToInt32(rd["Doc"]);
                                int pen = rd["Pen"] == DBNull.Value ? 0 : Convert.ToInt32(rd["Pen"]);
                                return (doc, pen);
                            }
                        }
                    }
                }
            }
            catch
            {
                // silenciar y devolver ceros
            }
            return (0, 0);
        }

        // NUEVO: consulta avanzada para seguimiento por prefijo (4 estatus)
        public (int UltimaMilla, int EnRuta, int Documentado, int Pendiente) ObtenerConteoSeguimientoPorPrefijoAvanzado(DateTime fechaCorte, string prefijoFolio)
        {
            try
            {
                using (var conn = new MySqlConnection(_connectionString))
                {
                    conn.Open();
                    const string sql = @"
                        SELECT 
                            SUM(CASE WHEN UPPER(COALESCE(EstatusGuia,'')) = 'ULTIMA MILLA' THEN 1 ELSE 0 END) AS UM,
                            SUM(CASE WHEN UPPER(COALESCE(EstatusGuia,'')) = 'EN RUTA' THEN 1 ELSE 0 END) AS ER,
                            SUM(CASE WHEN UPPER(COALESCE(EstatusGuia,'')) = 'DOCUMENTADO' THEN 1 ELSE 0 END) AS DOC,
                            SUM(CASE WHEN UPPER(COALESCE(EstatusGuia,'')) = 'PENDIENTE' THEN 1 ELSE 0 END) AS PEN
                        FROM guias
                        WHERE FolioGuia LIKE @prefijo
                          AND FechaElaboracion IS NOT NULL
                          AND FechaElaboracion <= @corte";
                    using (var cmd = new MySqlCommand(sql, conn))
                    {
                        cmd.Parameters.AddWithValue("@prefijo", prefijoFolio + "%");
                        cmd.Parameters.AddWithValue("@corte", fechaCorte);
                        using (var rd = cmd.ExecuteReader())
                        {
                            if (rd.Read())
                            {
                                return (
                                    rd["UM"] == DBNull.Value ? 0 : Convert.ToInt32(rd["UM"]),
                                    rd["ER"] == DBNull.Value ? 0 : Convert.ToInt32(rd["ER"]),
                                    rd["DOC"] == DBNull.Value ? 0 : Convert.ToInt32(rd["DOC"]),
                                    rd["PEN"] == DBNull.Value ? 0 : Convert.ToInt32(rd["PEN"])
                                );
                            }
                        }
                    }
                }
            }
            catch
            {
                // silencioso, devuelve ceros
            }
            return (0, 0, 0, 0);
        }

        // NUEVO: detalle de guías por prefijo y corte (>= 3 días de atraso)
        public DataTable ObtenerGuiasSeguimientoPorPrefijo(DateTime fechaCorte, string prefijoFolio, IList<string> estatus = null)
        {
            var dt = new DataTable();
            using (var conn = new MySqlConnection(_connectionString))
            {
                conn.Open();

                var sql = new StringBuilder(@"
                    SELECT *
                    FROM guias
                    WHERE FolioGuia LIKE @prefijo
                      AND FechaElaboracion IS NOT NULL
                      AND FechaElaboracion <= @corte");

                if (estatus != null && estatus.Count > 0)
                {
                    var placeholders = estatus.Select((s, i) => "@e" + i).ToArray();
                    sql.Append(" AND UPPER(COALESCE(EstatusGuia,'')) IN (");
                    sql.Append(string.Join(",", placeholders));
                    sql.Append(")");
                }

                sql.Append(" ORDER BY FechaElaboracion DESC");

                using (var cmd = new MySqlCommand(sql.ToString(), conn))
                {
                    cmd.Parameters.AddWithValue("@prefijo", prefijoFolio + "%");
                    cmd.Parameters.AddWithValue("@corte", fechaCorte);

                    if (estatus != null)
                    {
                        for (int i = 0; i < estatus.Count; i++)
                        {
                            cmd.Parameters.AddWithValue("@e" + i, estatus[i].ToUpperInvariant());
                        }
                    }

                    using (var adapter = new MySqlDataAdapter(cmd))
                    {
                        adapter.Fill(dt);
                    }
                }
            }
            return dt;
        }
    }

    internal sealed class ClienteEstadisticas
    {
        public int GuiasPorCobrar { get; set; }
        public decimal MontoPorCobrar { get; set; }
        public int GuiasPagadasOrigen { get; set; }
        public decimal MontoPagadasOrigen { get; set; }
        public int GuiasPagadasDestino { get; set; }
        public decimal MontoPagadasDestino { get; set; }
        public int GuiasCanceladas { get; set; }
        public decimal MontoCanceladas { get; set; }
        public int PaquetesEnviados { get; set; }
        public List<(string Destino, int TotalCajas)> Destinos { get; } = new List<(string Destino, int TotalCajas)>();
    }
}
