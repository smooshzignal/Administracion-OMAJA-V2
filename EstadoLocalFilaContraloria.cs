private void AsegurarColumnasLocalesContraloria(DataTable datos)
{
    if (datos == null)
    {
        return;
    }

    if (!datos.Columns.Contains("Estatus guias en cortes"))
    {
        datos.Columns.Add("Estatus guias en cortes", typeof(string));
    }

    if (!datos.Columns.Contains("Busqueda en cortes"))
    {
        datos.Columns.Add("Busqueda en cortes", typeof(string));
    }

    if (!datos.Columns.Contains("Observaciones de auditoria"))
    {
        datos.Columns.Add("Observaciones de auditoria", typeof(string));
    }
}

private void CargarEstadoLocalDesdeJsonContraloria()
{
    overridesEstatusGuiasEnCortesContraloria.Clear();
    overridesBusquedaEnCortesContraloria.Clear();
    overridesObservacionesAuditoriaContraloria.Clear();

    try
    {
        if (!File.Exists(rutaEstadoLocalContraloria))
        {
            return;
        }

        string json = File.ReadAllText(rutaEstadoLocalContraloria, Encoding.UTF8);
        var items = JsonConvert.DeserializeObject<List<EstadoLocalFilaContraloria>>(json);

        if (items == null)
        {
            return;
        }

        foreach (var item in items)
        {
            if (item == null || item.Id <= 0)
            {
                continue;
            }

            if (!string.IsNullOrWhiteSpace(item.EstatusGuiasEnCortes))
            {
                overridesEstatusGuiasEnCortesContraloria[item.Id] = item.EstatusGuiasEnCortes;
            }

            if (!string.IsNullOrWhiteSpace(item.BusquedaEnCortes))
            {
                overridesBusquedaEnCortesContraloria[item.Id] = item.BusquedaEnCortes;
            }

            if (!string.IsNullOrWhiteSpace(item.ObservacionesAuditoria))
            {
                overridesObservacionesAuditoriaContraloria[item.Id] = item.ObservacionesAuditoria;
            }
        }
    }
    catch (Exception ex)
    {
        MessageBox.Show(
            "No se pudo cargar el estado local de Contraloria.\n" + ex.Message,
            "Persistencia local",
            MessageBoxButtons.OK,
            MessageBoxIcon.Warning);
    }
}

private void AplicarEstadoLocalColumnasContraloria(DataTable datos)
{
    if (datos == null || !datos.Columns.Contains("id"))
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

        string valorEstatus;
        if (overridesEstatusGuiasEnCortesContraloria.TryGetValue(id, out valorEstatus))
        {
            row["Estatus guias en cortes"] = valorEstatus;
        }

        string valorBusqueda;
        if (overridesBusquedaEnCortesContraloria.TryGetValue(id, out valorBusqueda))
        {
            row["Busqueda en cortes"] = valorBusqueda;
        }

        string valorObservaciones;
        if (overridesObservacionesAuditoriaContraloria.TryGetValue(id, out valorObservaciones))
        {
            row["Observaciones de auditoria"] = valorObservaciones;
        }
    }
}

private void GuardarEstadoLocalEnJsonContraloria()
{
    try
    {
        string directorio = Path.GetDirectoryName(rutaEstadoLocalContraloria);
        if (!string.IsNullOrWhiteSpace(directorio))
        {
            Directory.CreateDirectory(directorio);
        }

        var ids = new HashSet<int>();

        foreach (var id in overridesEstatusGuiasEnCortesContraloria.Keys)
        {
            ids.Add(id);
        }

        foreach (var id in overridesBusquedaEnCortesContraloria.Keys)
        {
            ids.Add(id);
        }

        foreach (var id in overridesObservacionesAuditoriaContraloria.Keys)
        {
            ids.Add(id);
        }

        var items = new List<EstadoLocalFilaContraloria>();

        foreach (int id in ids.OrderBy(x => x))
        {
            string estatus;
            string busqueda;
            string observaciones;

            overridesEstatusGuiasEnCortesContraloria.TryGetValue(id, out estatus);
            overridesBusquedaEnCortesContraloria.TryGetValue(id, out busqueda);
            overridesObservacionesAuditoriaContraloria.TryGetValue(id, out observaciones);

            items.Add(new EstadoLocalFilaContraloria
            {
                Id = id,
                EstatusGuiasEnCortes = estatus,
                BusquedaEnCortes = busqueda,
                ObservacionesAuditoria = observaciones
            });
        }

        string json = JsonConvert.SerializeObject(items, Formatting.Indented);
        File.WriteAllText(rutaEstadoLocalContraloria, json, Encoding.UTF8);
    }
    catch (Exception ex)
    {
        MessageBox.Show(
            "No se pudo guardar el estado local de Contraloria.\n" + ex.Message,
            "Persistencia local",
            MessageBoxButtons.OK,
            MessageBoxIcon.Warning);
    }
}

private bool ObtenerIdFilaContraloria(DataGridViewRow row, out int id)
{
    id = 0;

    if (row == null || row.DataGridView == null || !row.DataGridView.Columns.Contains("id"))
    {
        return false;
    }

    return int.TryParse(
        Convert.ToString(row.Cells["id"]?.Value ?? string.Empty, CultureInfo.CurrentCulture),
        out id);
}

private void CapturarEstadoActualGridContraloria()
{
    if (dataGridViewContraloria == null || dataGridViewContraloria.Rows == null)
    {
        return;
    }

    foreach (DataGridViewRow row in dataGridViewContraloria.Rows)
    {
        if (row == null || row.IsNewRow)
        {
            continue;
        }

        int id;
        if (!ObtenerIdFilaContraloria(row, out id))
        {
            continue;
        }

        string estatus = dataGridViewContraloria.Columns.Contains("Estatus guias en cortes")
            ? Convert.ToString(row.Cells["Estatus guias en cortes"]?.Value ?? string.Empty).Trim()
            : string.Empty;

        string busqueda = dataGridViewContraloria.Columns.Contains("Busqueda en cortes")
            ? Convert.ToString(row.Cells["Busqueda en cortes"]?.Value ?? string.Empty).Trim()
            : string.Empty;

        string observaciones = dataGridViewContraloria.Columns.Contains("Observaciones de auditoria")
            ? Convert.ToString(row.Cells["Observaciones de auditoria"]?.Value ?? string.Empty)
            : string.Empty;

        if (string.IsNullOrWhiteSpace(estatus))
        {
            overridesEstatusGuiasEnCortesContraloria.Remove(id);
        }
        else
        {
            overridesEstatusGuiasEnCortesContraloria[id] = estatus;
        }

        if (string.IsNullOrWhiteSpace(busqueda))
        {
            overridesBusquedaEnCortesContraloria.Remove(id);
        }
        else
        {
            overridesBusquedaEnCortesContraloria[id] = busqueda;
        }

        if (string.IsNullOrWhiteSpace(observaciones))
        {
            overridesObservacionesAuditoriaContraloria.Remove(id);
        }
        else
        {
            overridesObservacionesAuditoriaContraloria[id] = observaciones;
        }
    }
}

private void Contraloria_FormClosingPersistenciaLocalContraloria(object sender, FormClosingEventArgs e)
{
    if (dataGridViewContraloria != null && dataGridViewContraloria.IsCurrentCellInEditMode)
    {
        dataGridViewContraloria.EndEdit();
    }

    CapturarEstadoActualGridContraloria();
    GuardarEstadoLocalEnJsonContraloria();
}

private sealed class EstadoLocalFilaContraloria
{
    public int Id { get; set; }
    public string EstatusGuiasEnCortes { get; set; }
    public string BusquedaEnCortes { get; set; }
    public string ObservacionesAuditoria { get; set; }
}