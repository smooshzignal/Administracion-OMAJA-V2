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


        public (int nuevos, int actualizados) InsertarDesdeDataTable(DataTable dt, Action<int, int> reportProgress = null)
        {
            if (dt == null || dt.Rows.Count == 0)
            {
                return (0, 0);
            }

            string ObtenerTexto(DataRow row, string columnName)
            {
                return row.Table.Columns.Contains(columnName)
                    ? row[columnName]?.ToString() ?? string.Empty
                    : string.Empty;
            }

            using (var conn = new MySqlConnection(_connectionString))
            {
                conn.Open();

                using (var transaction = conn.BeginTransaction())
                using (var existeCmd = new MySqlCommand("SELECT COUNT(*) FROM guias WHERE FolioGuia = @folio", conn, transaction))
                using (var insertCmd = new MySqlCommand(
                    @"INSERT INTO guias (
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
            )", conn, transaction))
                using (var updateCmd = new MySqlCommand(
                    @"UPDATE guias SET
                FechaElaboracion = @FechaElaboracion,
                HoraElaboracion = @HoraElaboracion,
                EstatusGuia = @EstatusGuia,
                Cliente = @Cliente,
                UbicacionActual = @UbicacionActual,
                Origen = @Origen,
                Destino = @Destino,
                TipoCobro = @TipoCobro,
                ZonaOperativaEntrega = @ZonaOperativaEntrega,
                TipoEntrega = @TipoEntrega,
                FechaEntrega = @FechaEntrega,
                HoraEntrega = @HoraEntrega,
                Tracking = @Tracking,
                Referencia = @Referencia,
                Subtotal = @Subtotal,
                Total = @Total,
                Sucursal = @Sucursal,
                FolioInforme = @FolioInforme,
                FolioEmbarque = @FolioEmbarque,
                UsuarioDocumento = @UsuarioDocumento,
                FechaCancelacion = @FechaCancelacion,
                UsuarioCancelacion = @UsuarioCancelacion,
                Remitente = @Remitente,
                Destinatario = @Destinatario,
                Cajas = @Cajas,
                ValorDeclarado = @ValorDeclarado,
                Observaciones = @Observaciones,
                Factura = @Factura,
                TimbradoSAT = @TimbradoSAT,
                FolioERP = @FolioERP,
                TipoCobroInicial = @TipoCobroInicial,
                FechaUltimaMilla = @FechaUltimaMilla,
                MotivoCancelacion = @MotivoCancelacion
            WHERE FolioGuia = @FolioGuiaWhere", conn, transaction))
                {
                    existeCmd.Parameters.Add("@folio", MySqlDbType.VarChar);

                    insertCmd.Parameters.Add("@FechaElaboracion", MySqlDbType.DateTime);
                    insertCmd.Parameters.Add("@HoraElaboracion", MySqlDbType.Time);
                    insertCmd.Parameters.Add("@FolioGuia", MySqlDbType.VarChar);
                    insertCmd.Parameters.Add("@EstatusGuia", MySqlDbType.VarChar);
                    insertCmd.Parameters.Add("@Cliente", MySqlDbType.VarChar);
                    insertCmd.Parameters.Add("@UbicacionActual", MySqlDbType.VarChar);
                    insertCmd.Parameters.Add("@Origen", MySqlDbType.VarChar);
                    insertCmd.Parameters.Add("@Destino", MySqlDbType.VarChar);
                    insertCmd.Parameters.Add("@TipoCobro", MySqlDbType.VarChar);
                    insertCmd.Parameters.Add("@ZonaOperativaEntrega", MySqlDbType.VarChar);
                    insertCmd.Parameters.Add("@TipoEntrega", MySqlDbType.VarChar);
                    insertCmd.Parameters.Add("@FechaEntrega", MySqlDbType.DateTime);
                    insertCmd.Parameters.Add("@HoraEntrega", MySqlDbType.Time);
                    insertCmd.Parameters.Add("@Tracking", MySqlDbType.VarChar);
                    insertCmd.Parameters.Add("@Referencia", MySqlDbType.VarChar);
                    insertCmd.Parameters.Add("@Subtotal", MySqlDbType.Decimal);
                    insertCmd.Parameters.Add("@Total", MySqlDbType.Decimal);
                    insertCmd.Parameters.Add("@Sucursal", MySqlDbType.VarChar);
                    insertCmd.Parameters.Add("@FolioInforme", MySqlDbType.VarChar);
                    insertCmd.Parameters.Add("@FolioEmbarque", MySqlDbType.VarChar);
                    insertCmd.Parameters.Add("@UsuarioDocumento", MySqlDbType.VarChar);
                    insertCmd.Parameters.Add("@FechaCancelacion", MySqlDbType.DateTime);
                    insertCmd.Parameters.Add("@UsuarioCancelacion", MySqlDbType.VarChar);
                    insertCmd.Parameters.Add("@Remitente", MySqlDbType.VarChar);
                    insertCmd.Parameters.Add("@Destinatario", MySqlDbType.VarChar);
                    insertCmd.Parameters.Add("@Cajas", MySqlDbType.Int32);
                    insertCmd.Parameters.Add("@ValorDeclarado", MySqlDbType.Decimal);
                    insertCmd.Parameters.Add("@Observaciones", MySqlDbType.VarChar);
                    insertCmd.Parameters.Add("@Factura", MySqlDbType.VarChar);
                    insertCmd.Parameters.Add("@TimbradoSAT", MySqlDbType.VarChar);
                    insertCmd.Parameters.Add("@FolioERP", MySqlDbType.VarChar);
                    insertCmd.Parameters.Add("@TipoCobroInicial", MySqlDbType.VarChar);
                    insertCmd.Parameters.Add("@FechaUltimaMilla", MySqlDbType.DateTime);
                    insertCmd.Parameters.Add("@MotivoCancelacion", MySqlDbType.VarChar);

                    updateCmd.Parameters.Add("@FechaElaboracion", MySqlDbType.DateTime);
                    updateCmd.Parameters.Add("@HoraElaboracion", MySqlDbType.Time);
                    updateCmd.Parameters.Add("@EstatusGuia", MySqlDbType.VarChar);
                    updateCmd.Parameters.Add("@Cliente", MySqlDbType.VarChar);
                    updateCmd.Parameters.Add("@UbicacionActual", MySqlDbType.VarChar);
                    updateCmd.Parameters.Add("@Origen", MySqlDbType.VarChar);
                    updateCmd.Parameters.Add("@Destino", MySqlDbType.VarChar);
                    updateCmd.Parameters.Add("@TipoCobro", MySqlDbType.VarChar);
                    updateCmd.Parameters.Add("@ZonaOperativaEntrega", MySqlDbType.VarChar);
                    updateCmd.Parameters.Add("@TipoEntrega", MySqlDbType.VarChar);
                    updateCmd.Parameters.Add("@FechaEntrega", MySqlDbType.DateTime);
                    updateCmd.Parameters.Add("@HoraEntrega", MySqlDbType.Time);
                    updateCmd.Parameters.Add("@Tracking", MySqlDbType.VarChar);
                    updateCmd.Parameters.Add("@Referencia", MySqlDbType.VarChar);
                    updateCmd.Parameters.Add("@Subtotal", MySqlDbType.Decimal);
                    updateCmd.Parameters.Add("@Total", MySqlDbType.Decimal);
                    updateCmd.Parameters.Add("@Sucursal", MySqlDbType.VarChar);
                    updateCmd.Parameters.Add("@FolioInforme", MySqlDbType.VarChar);
                    updateCmd.Parameters.Add("@FolioEmbarque", MySqlDbType.VarChar);
                    updateCmd.Parameters.Add("@UsuarioDocumento", MySqlDbType.VarChar);
                    updateCmd.Parameters.Add("@FechaCancelacion", MySqlDbType.DateTime);
                    updateCmd.Parameters.Add("@UsuarioCancelacion", MySqlDbType.VarChar);
                    updateCmd.Parameters.Add("@Remitente", MySqlDbType.VarChar);
                    updateCmd.Parameters.Add("@Destinatario", MySqlDbType.VarChar);
                    updateCmd.Parameters.Add("@Cajas", MySqlDbType.Int32);
                    updateCmd.Parameters.Add("@ValorDeclarado", MySqlDbType.Decimal);
                    updateCmd.Parameters.Add("@Observaciones", MySqlDbType.VarChar);
                    updateCmd.Parameters.Add("@Factura", MySqlDbType.VarChar);
                    updateCmd.Parameters.Add("@TimbradoSAT", MySqlDbType.VarChar);
                    updateCmd.Parameters.Add("@FolioERP", MySqlDbType.VarChar);
                    updateCmd.Parameters.Add("@TipoCobroInicial", MySqlDbType.VarChar);
                    updateCmd.Parameters.Add("@FechaUltimaMilla", MySqlDbType.DateTime);
                    updateCmd.Parameters.Add("@MotivoCancelacion", MySqlDbType.VarChar);
                    updateCmd.Parameters.Add("@FolioGuiaWhere", MySqlDbType.VarChar);

                    int nuevos = 0;
                    int actualizados = 0;
                    int totalFilas = dt.Rows.Count;

                    for (int i = 0; i < totalFilas; i++)
                    {
                        var row = dt.Rows[i];

                        try
                        {
                            string folioGuia = ObtenerTexto(row, "Folio Guía").Trim();
                            if (string.IsNullOrWhiteSpace(folioGuia))
                            {
                                if ((i + 1) % 100 == 0 || i + 1 == totalFilas)
                                {
                                    reportProgress?.Invoke(i + 1, totalFilas);
                                }

                                continue;
                            }

                            string fechaHoraOriginal = ObtenerTexto(row, "Fecha/Hora Elaboración");
                            DateTime? fechaElaboracion = null;
                            TimeSpan? horaElaboracion = null;

                            if (!string.IsNullOrEmpty(fechaHoraOriginal) && fechaHoraOriginal.Contains(" "))
                            {
                                var partes = fechaHoraOriginal.Split(' ');
                                DateTime fecha;
                                TimeSpan hora;

                                if (partes.Length > 0 && DateTime.TryParse(partes[0], out fecha))
                                {
                                    fechaElaboracion = fecha;
                                }

                                if (partes.Length > 1 && TimeSpan.TryParse(partes[1], out hora))
                                {
                                    horaElaboracion = hora;
                                }
                            }

                            string fechaHoraEntrega = ObtenerTexto(row, "Fecha y Hora de Entrega");
                            DateTime? fechaEntrega = null;
                            TimeSpan? horaEntrega = null;

                            if (!string.IsNullOrEmpty(fechaHoraEntrega) && fechaHoraEntrega.Contains(" "))
                            {
                                var partes = fechaHoraEntrega.Split(' ');
                                DateTime fecha;
                                TimeSpan hora;

                                if (partes.Length > 0 && DateTime.TryParse(partes[0], out fecha))
                                {
                                    fechaEntrega = fecha;
                                }

                                if (partes.Length > 1 && TimeSpan.TryParse(partes[1], out hora))
                                {
                                    horaEntrega = hora;
                                }
                            }

                            DateTime fechaCancelacion;
                            DateTime fechaUltimaMilla;
                            decimal subtotal;
                            decimal total;
                            int cajas;
                            decimal valorDeclarado;

                            Action<MySqlCommand> asignarValores = cmd =>
                            {
                                cmd.Parameters["@FechaElaboracion"].Value = (object)fechaElaboracion ?? DBNull.Value;
                                cmd.Parameters["@HoraElaboracion"].Value = (object)horaElaboracion ?? DBNull.Value;
                                cmd.Parameters["@EstatusGuia"].Value = ObtenerTexto(row, "Estatus Guía");
                                cmd.Parameters["@Cliente"].Value = ObtenerTexto(row, "Cliente");
                                cmd.Parameters["@UbicacionActual"].Value = ObtenerTexto(row, "Ubicación Actual");
                                cmd.Parameters["@Origen"].Value = ObtenerTexto(row, "Origen");
                                cmd.Parameters["@Destino"].Value = ObtenerTexto(row, "Destino");
                                cmd.Parameters["@TipoCobro"].Value = ObtenerTexto(row, "Tipo cobro");
                                cmd.Parameters["@ZonaOperativaEntrega"].Value = ObtenerTexto(row, "Zona Operativa Entrega");
                                cmd.Parameters["@TipoEntrega"].Value = ObtenerTexto(row, "Tipo de entrega");
                                cmd.Parameters["@FechaEntrega"].Value = (object)fechaEntrega ?? DBNull.Value;
                                cmd.Parameters["@HoraEntrega"].Value = (object)horaEntrega ?? DBNull.Value;
                                cmd.Parameters["@Tracking"].Value = ObtenerTexto(row, "Tracking");
                                cmd.Parameters["@Referencia"].Value = ObtenerTexto(row, "Referencia");
                                cmd.Parameters["@Subtotal"].Value = decimal.TryParse(ObtenerTexto(row, "Subtotal"), out subtotal) ? subtotal : 0m;
                                cmd.Parameters["@Total"].Value = decimal.TryParse(ObtenerTexto(row, "Total"), out total) ? total : 0m;
                                cmd.Parameters["@Sucursal"].Value = ObtenerTexto(row, "Sucursal");
                                cmd.Parameters["@FolioInforme"].Value = ObtenerTexto(row, "Folio Informe");
                                cmd.Parameters["@FolioEmbarque"].Value = ObtenerTexto(row, "Folio Embarque");
                                cmd.Parameters["@UsuarioDocumento"].Value = ObtenerTexto(row, "Usuario Documento");
                                cmd.Parameters["@FechaCancelacion"].Value = DateTime.TryParse(ObtenerTexto(row, "Fecha de Cancelación"), out fechaCancelacion) ? (object)fechaCancelacion : DBNull.Value;
                                cmd.Parameters["@UsuarioCancelacion"].Value = ObtenerTexto(row, "Usuario de Cancelación");
                                cmd.Parameters["@Remitente"].Value = ObtenerTexto(row, "Remitente");
                                cmd.Parameters["@Destinatario"].Value = ObtenerTexto(row, "Destinatario");
                                cmd.Parameters["@Cajas"].Value = int.TryParse(ObtenerTexto(row, "Cajas"), out cajas) ? cajas : 0;
                                cmd.Parameters["@ValorDeclarado"].Value = decimal.TryParse(ObtenerTexto(row, "Valor declarado"), out valorDeclarado) ? valorDeclarado : 0m;
                                cmd.Parameters["@Observaciones"].Value = ObtenerTexto(row, "Observaciones");
                                cmd.Parameters["@Factura"].Value = ObtenerTexto(row, "Factura");
                                cmd.Parameters["@TimbradoSAT"].Value = ObtenerTexto(row, "Timbrado SAT");
                                cmd.Parameters["@FolioERP"].Value = ObtenerTexto(row, "Folio ERP");
                                cmd.Parameters["@TipoCobroInicial"].Value = ObtenerTexto(row, "Tipo de cobro inicial");
                                cmd.Parameters["@FechaUltimaMilla"].Value = DateTime.TryParse(ObtenerTexto(row, "Fecha última milla"), out fechaUltimaMilla) ? (object)fechaUltimaMilla : DBNull.Value;
                                cmd.Parameters["@MotivoCancelacion"].Value = ObtenerTexto(row, "Motivo cancelación");
                            };

                            existeCmd.Parameters["@folio"].Value = folioGuia;
                            int existe = Convert.ToInt32(existeCmd.ExecuteScalar());

                            if (existe > 0)
                            {
                                asignarValores(updateCmd);
                                updateCmd.Parameters["@FolioGuiaWhere"].Value = folioGuia;
                                updateCmd.ExecuteNonQuery();
                                actualizados++;
                            }
                            else
                            {
                                asignarValores(insertCmd);
                                insertCmd.Parameters["@FolioGuia"].Value = folioGuia;
                                insertCmd.ExecuteNonQuery();
                                nuevos++;
                            }
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"Error al insertar/actualizar fila {i + 1}: {ex.Message}");
                        }

                        if ((i + 1) % 100 == 0 || i + 1 == totalFilas)
                        {
                            reportProgress?.Invoke(i + 1, totalFilas);
                        }
                    }

                    transaction.Commit();
                    return (nuevos, actualizados);
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

        public (int nuevos, int actualizados) ImportarDesdeExcel(string filePath, DataGridView dataGridView, Action<int, int> reportProgress)
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
                        return (0, 0);
                    }

                    table = result.Tables[0];
                }

                if (table.Rows.Count == 0)
                {
                    MessageBox.Show("El archivo no contiene registros para importar.", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return (0, 0);
                }

                reportProgress?.Invoke(0, table.Rows.Count);

                dataGridView.Invoke((Action)(() =>
                {
                    dataGridView.DataSource = table;
                }));

                return InsertarDesdeDataTable(table, reportProgress);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error al importar: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return (0, 0);
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

        private static DataColumn ObtenerColumnaPorAliasFacturacionContraloria(DataTable tabla, params string[] nombres)
        {
            if (tabla == null || nombres == null)
            {
                return null;
            }

            foreach (string nombre in nombres)
            {
                var columna = tabla.Columns.Cast<DataColumn>()
                    .FirstOrDefault(c => c.ColumnName.Equals(nombre, StringComparison.OrdinalIgnoreCase));

                if (columna != null)
                {
                    return columna;
                }
            }

            return null;
        }

        private static object ObtenerValorColumnaFacturacionContraloria(DataRow row, params string[] nombres)
        {
            if (row == null || row.Table == null)
            {
                return null;
            }

            var columna = ObtenerColumnaPorAliasFacturacionContraloria(row.Table, nombres);
            return columna == null ? null : row[columna];
        }

        private static string ObtenerTextoColumnaFacturacionContraloria(DataRow row, params string[] nombres)
        {
            object valor = ObtenerValorColumnaFacturacionContraloria(row, nombres);
            return valor == null || valor == DBNull.Value
                ? null
                : Convert.ToString(valor).Trim();
        }

        private static int? ObtenerEnteroColumnaFacturacionContraloria(DataRow row, params string[] nombres)
        {
            object valor = ObtenerValorColumnaFacturacionContraloria(row, nombres);
            if (valor == null || valor == DBNull.Value)
            {
                return null;
            }

            if (valor is int entero)
            {
                return entero;
            }

            if (valor is long largo)
            {
                return (int)largo;
            }

            if (valor is double doble)
            {
                return (int)Math.Round(doble);
            }

            int resultado;
            return int.TryParse(Convert.ToString(valor), NumberStyles.Any, CultureInfo.InvariantCulture, out resultado) ||
                   int.TryParse(Convert.ToString(valor), NumberStyles.Any, CultureInfo.CurrentCulture, out resultado)
                ? resultado
                : (int?)null;
        }

        private static DateTime? ObtenerFechaColumnaFacturacionContraloria(DataRow row, params string[] nombres)
        {
            object valor = ObtenerValorColumnaFacturacionContraloria(row, nombres);
            if (valor == null || valor == DBNull.Value)
            {
                return null;
            }

            if (valor is DateTime fecha)
            {
                return fecha.Date;
            }

            if (valor is double oaDate)
            {
                return DateTime.FromOADate(oaDate).Date;
            }

            DateTime resultado;
            return DateTime.TryParse(Convert.ToString(valor), CultureInfo.InvariantCulture, DateTimeStyles.None, out resultado) ||
                   DateTime.TryParse(Convert.ToString(valor), CultureInfo.CurrentCulture, DateTimeStyles.None, out resultado)
                ? resultado.Date
                : (DateTime?)null;
        }

        private static decimal? ObtenerDecimalColumnaFacturacionContraloria(DataRow row, params string[] nombres)
        {
            object valor = ObtenerValorColumnaFacturacionContraloria(row, nombres);
            if (valor == null || valor == DBNull.Value)
            {
                return null;
            }

            if (valor is decimal decimalValor)
            {
                return decimalValor;
            }

            if (valor is double doble)
            {
                return Convert.ToDecimal(doble);
            }

            decimal resultado;
            return decimal.TryParse(Convert.ToString(valor), NumberStyles.Any, CultureInfo.InvariantCulture, out resultado) ||
                   decimal.TryParse(Convert.ToString(valor), NumberStyles.Any, CultureInfo.CurrentCulture, out resultado)
                ? resultado
                : (decimal?)null;
        }

        public (DataTable datos, int nuevos, int actualizados) ImportarFacturacionDesdeExcelContraloria(
    string filePath,
    Action<int, int> reportProgress = null)
        {
            try
            {
                Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

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
                        return (new DataTable(), 0, 0);
                    }

                    table = result.Tables[0];
                }

                if (table.Rows.Count == 0)
                {
                    MessageBox.Show("El archivo no contiene registros para importar.", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return (table, 0, 0);
                }

                int nuevos = 0;
                int actualizados = 0;
                var erroresContraloria = new List<string>();

                reportProgress?.Invoke(0, table.Rows.Count);

                using (var conn = new MySqlConnection(_connectionString))
                {
                    conn.Open();

                    using (var transaction = conn.BeginTransaction())
                    using (var existeCmd = new MySqlCommand("SELECT COUNT(*) FROM facturacion WHERE id = @id", conn, transaction))
                    using (var insertCmd = new MySqlCommand(@"
                INSERT INTO facturacion
                (
                    id, sucursal, fecha, numero, cliente, documento, nota_de_debito, uuid,
                    descuento, sub_total, iva, retencion, total, moneda, estatus,
                    folio_fiscal_uuid, destino, origen, no_viaje
                )
                VALUES
                (
                    @id, @sucursal, @fecha, @numero, @cliente, @documento, @nota_de_debito, @uuid,
                    @descuento, @sub_total, @iva, @retencion, @total, @moneda, @estatus,
                    @folio_fiscal_uuid, @destino, @origen, @no_viaje
                )", conn, transaction))
                    using (var updateCmd = new MySqlCommand(@"
                UPDATE facturacion SET
                    sucursal = @sucursal,
                    fecha = @fecha,
                    numero = @numero,
                    cliente = @cliente,
                    documento = @documento,
                    nota_de_debito = @nota_de_debito,
                    uuid = @uuid,
                    descuento = @descuento,
                    sub_total = @sub_total,
                    iva = @iva,
                    retencion = @retencion,
                    total = @total,
                    moneda = @moneda,
                    estatus = @estatus,
                    folio_fiscal_uuid = @folio_fiscal_uuid,
                    destino = @destino,
                    origen = @origen,
                    no_viaje = @no_viaje
                WHERE id = @id", conn, transaction))
                    {
                        existeCmd.Parameters.Add("@id", MySqlDbType.Int32);

                        int siguienteIdFacturacionContraloria = ObtenerSiguienteIdFacturacionContraloria(conn, transaction);

                        for (int i = 0; i < table.Rows.Count; i++)
                        {
                            DataRow row = table.Rows[i];

                            try
                            {
                                int? idExcel = ObtenerEnteroColumnaFacturacionContraloria(row, "id", "ID", "Id");
                                int idOperacion;
                                bool existe;

                                if (idExcel.HasValue)
                                {
                                    idOperacion = idExcel.Value;
                                    existeCmd.Parameters["@id"].Value = idOperacion;
                                    existe = Convert.ToInt32(existeCmd.ExecuteScalar()) > 0;

                                    if (idOperacion >= siguienteIdFacturacionContraloria)
                                    {
                                        siguienteIdFacturacionContraloria = idOperacion + 1;
                                    }
                                }
                                else
                                {
                                    idOperacion = siguienteIdFacturacionContraloria;
                                    siguienteIdFacturacionContraloria++;
                                    existe = false;
                                }

                                string sucursal = ObtenerTextoColumnaFacturacionContraloria(row, "sucursal", "Sucursal");
                                DateTime? fecha = ObtenerFechaColumnaFacturacionContraloria(row, "fecha", "Fecha");
                                string numero = ObtenerTextoColumnaFacturacionContraloria(row, "numero", "Número", "Numero");
                                string cliente = ObtenerTextoColumnaFacturacionContraloria(row, "cliente", "Cliente");
                                string documento = ObtenerTextoColumnaFacturacionContraloria(row, "documento", "Documento");
                                decimal? notaDebito = ObtenerDecimalColumnaFacturacionContraloria(row, "nota_de_debito", "Nota de Débito", "Nota de Debito", "Nota Debito");
                                string uuid = ObtenerTextoColumnaFacturacionContraloria(row, "uuid", "UUID");
                                decimal? descuento = ObtenerDecimalColumnaFacturacionContraloria(row, "descuento", "Descuento");
                                decimal? subTotal = ObtenerDecimalColumnaFacturacionContraloria(row, "sub_total", "Sub Total", "Subtotal");
                                decimal? iva = ObtenerDecimalColumnaFacturacionContraloria(row, "iva", "IVA");
                                decimal? retencion = ObtenerDecimalColumnaFacturacionContraloria(row, "retencion", "Retención", "Retencion");
                                decimal? total = ObtenerDecimalColumnaFacturacionContraloria(row, "total", "Total");
                                string moneda = ObtenerTextoColumnaFacturacionContraloria(row, "moneda", "Moneda");
                                string estatus = ObtenerTextoColumnaFacturacionContraloria(row, "estatus", "Estatus");
                                string folioFiscalUuid = ObtenerTextoColumnaFacturacionContraloria(row, "folio_fiscal_uuid", "Folio Fiscal UUID", "Folio Fiscal Uuid");
                                string destino = ObtenerTextoColumnaFacturacionContraloria(row, "destino", "Destino");
                                string origen = ObtenerTextoColumnaFacturacionContraloria(row, "origen", "Origen");
                                string noViaje = ObtenerTextoColumnaFacturacionContraloria(row, "no_viaje", "No Viaje", "No. Viaje", "NoViaje");

                                MySqlCommand cmd = existe ? updateCmd : insertCmd;
                                cmd.Parameters.Clear();

                                cmd.Parameters.AddWithValue("@id", idOperacion);
                                cmd.Parameters.AddWithValue("@sucursal", (object)sucursal ?? DBNull.Value);
                                cmd.Parameters.AddWithValue("@fecha", (object)fecha ?? DBNull.Value);
                                cmd.Parameters.AddWithValue("@numero", (object)numero ?? DBNull.Value);
                                cmd.Parameters.AddWithValue("@cliente", (object)cliente ?? DBNull.Value);
                                cmd.Parameters.AddWithValue("@documento", (object)documento ?? DBNull.Value);
                                cmd.Parameters.AddWithValue("@nota_de_debito", (object)notaDebito ?? DBNull.Value);
                                cmd.Parameters.AddWithValue("@uuid", (object)uuid ?? DBNull.Value);
                                cmd.Parameters.AddWithValue("@descuento", descuento ?? 0m);
                                cmd.Parameters.AddWithValue("@sub_total", subTotal ?? 0m);
                                cmd.Parameters.AddWithValue("@iva", iva ?? 0m);
                                cmd.Parameters.AddWithValue("@retencion", retencion ?? 0m);
                                cmd.Parameters.AddWithValue("@total", total ?? 0m);
                                cmd.Parameters.AddWithValue("@moneda", string.IsNullOrWhiteSpace(moneda) ? "PESOS" : moneda);
                                cmd.Parameters.AddWithValue("@estatus", (object)estatus ?? DBNull.Value);
                                cmd.Parameters.AddWithValue("@folio_fiscal_uuid", (object)folioFiscalUuid ?? DBNull.Value);
                                cmd.Parameters.AddWithValue("@destino", (object)destino ?? DBNull.Value);
                                cmd.Parameters.AddWithValue("@origen", (object)origen ?? DBNull.Value);
                                cmd.Parameters.AddWithValue("@no_viaje", (object)noViaje ?? DBNull.Value);

                                cmd.ExecuteNonQuery();

                                if (existe)
                                {
                                    actualizados++;
                                }
                                else
                                {
                                    nuevos++;
                                }
                            }
                            catch (Exception ex)
                            {
                                erroresContraloria.Add("Fila " + (i + 1).ToString(CultureInfo.InvariantCulture) + ": " + ex.Message);
                            }

                            reportProgress?.Invoke(i + 1, table.Rows.Count);
                        }

                        transaction.Commit();
                    }
                }

                if (erroresContraloria.Count > 0)
                {
                    string detalle = string.Join(Environment.NewLine, erroresContraloria.Take(10));
                    MessageBox.Show(
                        "La importación terminó con errores parciales.\n" +
                        "Nuevos: " + nuevos.ToString(CultureInfo.InvariantCulture) +
                        "\nActualizados: " + actualizados.ToString(CultureInfo.InvariantCulture) +
                        "\nErrores: " + erroresContraloria.Count.ToString(CultureInfo.InvariantCulture) +
                        "\n\nPrimeros errores:\n" + detalle,
                        "Importación parcial",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);
                }

                return (table, nuevos, actualizados);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error al importar facturación Contraloria: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return (new DataTable(), 0, 0);
            }
        }

        public DataTable ObtenerGuiasPorFiltroTipoCobroContraloria(
    DateTime fechaInicio,
    DateTime fechaFin,
    string sucursal,
    string destino,
    bool incluirPagado,
    bool incluirPorCobrar,
    bool incluirCancelado)
{
    var condiciones = new List<string>();

    if (incluirPagado)
    {
        condiciones.Add("(UPPER(COALESCE(TipoCobro, '')) = 'PAGADO' OR UPPER(COALESCE(TipoCobroInicial, '')) = 'PAGADO')");
    }

    if (incluirPorCobrar)
    {
        condiciones.Add("UPPER(COALESCE(TipoCobro, '')) = 'POR COBRAR'");
    }

    if (incluirCancelado)
    {
        condiciones.Add("UPPER(COALESCE(EstatusGuia, '')) = 'CANCELADO'");
    }

    if (condiciones.Count == 0)
    {
        return ObtenerTodasGuiasFiltrado(fechaInicio, fechaFin, sucursal, destino);
    }

    string condicionExtra = "(" + string.Join(" OR ", condiciones) + ")";
    return EjecutarConsultaGuias(fechaInicio, fechaFin, sucursal, destino, condicionExtra);
}

private static int ObtenerSiguienteIdFacturacionContraloria(MySqlConnection conn, MySqlTransaction transaction)
{
    using (var cmd = new MySqlCommand("SELECT IFNULL(MAX(id), 0) + 1 FROM facturacion", conn, transaction))
    {
        object result = cmd.ExecuteScalar();
        return result == null || result == DBNull.Value ? 1 : Convert.ToInt32(result);
    }
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
