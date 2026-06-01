[index.html](https://github.com/user-attachments/files/28478688/index.html)
<!DOCTYPE html>
<html lang="es" style="height: 100%;">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0, user-scalable=yes, viewport-fit=cover">
    <title>FARMACIA ESE HOSPITAL LOCAL SAN ZENON | DASHBOARD INTEGRAL</title>
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0-beta3/css/all.min.css">
    <script src="https://cdn.sheetjs.com/xlsx-0.20.2/package/dist/xlsx.full.min.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js"></script>
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        html, body { height: 100%; margin: 0; padding: 0; overflow: hidden; }
        body { font-family: 'Segoe UI', Roboto, 'Helvetica Neue', sans-serif; background: linear-gradient(135deg, #0f172a 0%, #1e293b 100%); display: flex; align-items: center; justify-content: center; padding: 20px; }
        .login-container { width: 100%; max-width: 520px; background: white; border-radius: 40px; box-shadow: 0 25px 50px -12px rgba(0,0,0,0.25); overflow: hidden; }
        .login-header { background: #1e3c72; padding: 32px 28px; text-align: center; color: white; }
        .login-header h2 { font-size: 2rem; margin-bottom: 8px; }
        .login-body { padding: 32px 28px; }
        .input-field { margin-bottom: 24px; }
        .input-field label { display: block; margin-bottom: 8px; font-weight: 600; color: #1e293b; }
        .input-field input { width: 100%; padding: 14px 18px; border: 1px solid #cbd5e1; border-radius: 32px; font-size: 1rem; }
        .btn-primary { background: #1e3c72; color: white; border: none; width: 100%; padding: 14px; border-radius: 40px; font-weight: bold; font-size: 1.1rem; cursor: pointer; transition: 0.2s; }
        .btn-primary:hover { background: #0f2a4a; }
        .app-container { width: 100%; height: 100%; background: white; border-radius: 32px; display: flex; flex-direction: column; overflow: hidden; }
        .app-header { background: #1e3c72; color: white; padding: 12px 24px; display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap; gap: 12px; flex-shrink: 0; }
        .user-info { display: flex; gap: 20px; align-items: center; flex-wrap: wrap; }
        .employee-name, .user-role { background: rgba(255,255,255,0.2); padding: 6px 14px; border-radius: 40px; font-size: 0.85rem; }
        .notification-bell { position: relative; cursor: pointer; font-size: 1.4rem; }
        .badge { position: absolute; top: -8px; right: -12px; background-color: #dc2626; color: white; border-radius: 50%; padding: 2px 6px; font-size: 0.65rem; }
        .tabs { display: flex; gap: 12px; background: #f1f5f9; padding: 10px 20px 0 20px; flex-wrap: wrap; flex-shrink: 0; }
        .tab-btn { padding: 10px 24px; background: #e2e8f0; border: none; border-radius: 40px; font-weight: 600; cursor: pointer; font-size: 0.95rem; transition: all 0.2s; margin-bottom: 8px; }
        .tab-btn.active { background: #1e3c72; color: white; box-shadow: 0 4px 8px rgba(0,0,0,0.1); }
        .tab-content { display: none; flex: 1; overflow-y: auto; padding: 20px; }
        .tab-content.active { display: flex; flex-direction: column; }
        .card { background: #f8fafc; border-radius: 24px; padding: 20px; margin-bottom: 24px; box-shadow: 0 1px 3px rgba(0,0,0,0.05); }
        .form-grid { display: grid; grid-template-columns: repeat(2, 1fr); gap: 20px; align-items: end; }
        .form-field { display: flex; flex-direction: column; gap: 6px; }
        .form-field label { font-weight: 600; color: #1e293b; font-size: 0.85rem; }
        .form-field input, .form-field select { padding: 10px 12px; border: 1px solid #cbd5e1; border-radius: 16px; font-size: 0.9rem; width: 100%; }
        .btn-secondary { background: #3b82f6; color: white; border: none; padding: 10px 16px; border-radius: 28px; cursor: pointer; font-weight: 500; transition: 0.2s; }
        .btn-secondary:hover { background: #2563eb; }
        .btn-outline { background: transparent; border: 1px solid #64748b; padding: 10px 16px; border-radius: 28px; cursor: pointer; font-weight: 500; }
        .btn-danger { background: #dc2626; color: white; border: none; padding: 10px 16px; border-radius: 28px; cursor: pointer; }
        .btn-warning { background: #f59e0b; color: white; border: none; padding: 10px 16px; border-radius: 28px; cursor: pointer; }
        .btn-success { background: #10b981; color: white; border: none; padding: 10px 16px; border-radius: 28px; cursor: pointer; }
        .table-wrapper { flex: 1; overflow: auto; border-radius: 20px; max-height: 55vh; border: 1px solid #e2e8f0; }
        table { width: 100%; border-collapse: collapse; background: white; min-width: 1000px; }
        th, td { padding: 12px 16px; text-align: left; border-bottom: 1px solid #e2e8f0; }
        th { background-color: #f1f5f9; position: sticky; top: 0; }
        .expired-row { background-color: #fee2e2 !important; }
        .stock-highlight { background: #dbeafe; border-radius: 20px; padding: 2px 10px; display: inline-block; }
        .cart-item { background: white; border-radius: 20px; padding: 12px 16px; margin-bottom: 10px; display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap; gap: 12px; border: 1px solid #e2e8f0; }
        .cart-item-info { display: flex; flex-wrap: wrap; gap: 16px; align-items: center; flex: 2; }
        .cart-item-actions { display: flex; gap: 8px; }
        .cart-item-actions .btn-icon { background: transparent; border: none; font-size: 1.3rem; cursor: pointer; padding: 4px 8px; }
        .cart-item-actions .btn-icon-edit { color: #3b82f6; }
        .cart-item-actions .btn-icon-delete { color: #dc2626; }
        .modal { position: fixed; top: 0; left: 0; width: 100%; height: 100%; background: rgba(0,0,0,0.6); display: flex; align-items: center; justify-content: center; z-index: 2000; padding: 20px; }
        .modal-content { background: white; border-radius: 32px; width: 100%; max-width: 800px; max-height: 90vh; overflow-y: auto; padding: 28px; }
        .flex-between { display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap; gap: 16px; margin-bottom: 20px; }
        .btn-icon { padding: 6px 10px; margin: 0 2px; background: transparent; border: none; cursor: pointer; font-size: 1.2rem; }
        .btn-icon-edit { color: #3b82f6; }
        .btn-icon-delete { color: #dc2626; }
        .alert-tooltip { position: fixed; background: white; border-radius: 20px; padding: 16px; box-shadow: 0 10px 25px rgba(0,0,0,0.2); z-index: 1000; right: 30px; top: 90px; width: 320px; display: none; max-height: 80vh; overflow-y: auto; }
        .umbral-bar { background: #eef2ff; padding: 12px 20px; border-radius: 50px; margin-bottom: 20px; display: flex; flex-wrap: wrap; gap: 16px; align-items: center; }
        .filter-bar { display: flex; flex-wrap: wrap; gap: 12px; margin-bottom: 20px; align-items: flex-end; background: #f1f5f9; padding: 15px; border-radius: 60px; }
        .filter-bar select, .filter-bar input { border-radius: 40px !important; padding: 10px 18px !important; }
        .config-list { margin-bottom: 24px; border-bottom: 1px solid #e2e8f0; padding-bottom: 16px; }
        .config-item { display: flex; justify-content: space-between; align-items: center; background: #f1f5f9; padding: 8px 12px; border-radius: 20px; margin-bottom: 8px; }
        .add-item { display: flex; gap: 12px; margin-top: 12px; align-items: center; }
        .add-item input { flex: 1; padding: 12px 16px; border: 1px solid #cbd5e1; border-radius: 40px !important; font-size: 0.9rem; }
        .add-item button { border-radius: 40px !important; padding: 10px 20px; background: #10b981; color: white; border: none; cursor: pointer; }
        .user-enabled-checkbox { width: 20px; height: 20px; cursor: pointer; }
        .disabled-user-row { opacity: 0.6; background-color: #fef2f2; }
        .license-overlay { position: fixed; top:0; left:0; width:100%; height:100%; background: rgba(0,0,0,0.85); z-index: 10000; display: flex; align-items: center; justify-content: center; }
        .license-modal { background: white; max-width: 400px; border-radius: 32px; padding: 32px; text-align: center; }
        .dashboard-kpi { display: flex; flex-wrap: wrap; gap: 20px; margin-bottom: 30px; }
        .kpi-card { background: white; border-radius: 28px; padding: 20px; flex: 1 1 200px; box-shadow: 0 4px 12px rgba(0,0,0,0.05); text-align: center; border-left: 6px solid #1e3c72; }
        .kpi-number { font-size: 2.2rem; font-weight: 800; color: #1e3c72; }
        .kpi-label { font-size: 0.85rem; color: #475569; margin-top: 6px; }
        .dashboard-charts { display: flex; flex-wrap: wrap; gap: 24px; margin-bottom: 30px; }
        .chart-box { background: white; border-radius: 28px; padding: 16px; flex: 1 1 300px; box-shadow: 0 2px 8px rgba(0,0,0,0.05); }
        .dashboard-tables { margin-top: 16px; }
        @media (max-width: 768px) { .form-grid { grid-template-columns: 1fr; } .tabs { gap: 8px; } .tab-btn { padding: 6px 16px; font-size: 0.8rem; } .dashboard-charts { flex-direction: column; } }
        .forgot-link { text-align: center; margin-top: 15px; font-size: 0.9rem; cursor: pointer; color: #1e3c72; text-decoration: underline; }
        .radio-group { display: flex; gap: 20px; margin: 15px 0; }
        .radio-group label { font-weight: normal; margin-left: 5px; }
        /* Botón re-dispensar siempre visible */
        .btn-redispensar { background-color: #0ea5e9; color: white; border: none; padding: 6px 12px; border-radius: 20px; cursor: pointer; font-size: 0.85rem; display: inline-flex; align-items: center; gap: 6px; }
        .btn-redispensar:hover { background-color: #0284c7; }
        #historialEntregasTable td:last-child button { background-color: #0ea5e9; border: none; color: white; padding: 6px 12px; border-radius: 20px; cursor: pointer; white-space: nowrap; }
        #historialEntregasTable td:last-child button:hover { background-color: #0284c7; }
    </style>
</head>
<body>
<div id="loginView" style="display: flex; justify-content: center; align-items: center; width: 100%;">
    <div class="login-container"><div class="login-header"><h2><i class="fas fa-pills"></i> FARMACIA ESE HOSPITAL LOCAL SAN ZENON</h2><p>Iniciar sesión</p></div>
    <div class="login-body"><div class="input-field"><label>Usuario</label><input type="text" id="loginUser" placeholder="Ingrese su usuario"></div>
    <div class="input-field"><label>Contraseña</label><input type="password" id="loginPass" placeholder="••••••"></div>
    <button class="btn-primary" id="loginBtn"><i class="fas fa-sign-in-alt"></i> Acceder</button>
    <div class="forgot-link" id="forgotPasswordLink">¿Olvidó su usuario o contraseña? Restablezca aquí</div>
    <p style="margin-top: 20px; text-align: center; font-size: 0.8rem;">@Copyright - CoorporacionDacs</p></div></div>
</div>
<div id="appView" class="app-container" style="display: none;">
    <div class="app-header"><h2><i class="fas fa-chart-line"></i> FARMACIA ESE HOSPITAL LOCAL SAN ZENON</h2>
        <div class="user-info"><div class="notification-bell" id="bellIcon"><i class="fas fa-bell"></i><span id="alertBadge" class="badge" style="display: none;">0</span></div>
        <button id="configBtn" class="btn-outline" style="background: rgba(255,255,255,0.2); color:white; border: none;"><i class="fas fa-cog"></i> Configurar</button>
        <span class="employee-name" id="currentUserLabel"></span><span class="user-role" id="userRoleLabel"></span><button class="btn-outline" id="logoutBtn">Cerrar sesión</button></div>
    </div>
    <div id="alertPanel" class="alert-tooltip"></div>
    <div class="tabs">
        <button class="tab-btn active" data-tab="dashboard">📊 Dashboard</button>
        <button class="tab-btn" data-tab="inventario">📦 Inventario</button>
        <button class="tab-btn" data-tab="movimientos">🔄 Movimientos</button>
        <button class="tab-btn" data-tab="entregas">💊 Entrega por Área</button>
        <button class="tab-btn" data-tab="historialEntregas">📜 Historial Entregas</button>
        <button class="tab-btn" data-tab="reportes">📊 Reportes</button>
        <button class="tab-btn" data-tab="usuarios">👥 Usuarios</button>
        <button class="tab-btn" data-tab="licencia" id="licenciaTabBtn" style="display: none;">🔐 Licencia</button>
    </div>

    <!-- Dashboard -->
    <div id="tabDashboard" class="tab-content active">
        <div class="card">
            <h3><i class="fas fa-chart-simple"></i> Panel de Control - Indicadores Clave</h3>
            <div class="dashboard-kpi" id="kpiContainer"></div>
            <div class="dashboard-charts">
                <div class="chart-box"><canvas id="tipoChart" width="400" height="250"></canvas><p class="kpi-label" style="text-align:center;">Distribución por tipo</p></div>
                <div class="chart-box"><canvas id="topEntregasChart" width="400" height="250"></canvas><p class="kpi-label" style="text-align:center;">Top 5 productos más entregados</p></div>
                <div class="chart-box"><canvas id="topDestinatariosChart" width="400" height="250"></canvas><p class="kpi-label" style="text-align:center;">Top 5 destinatarios que más reciben</p></div>
            </div>
            <div class="dashboard-kpi" style="justify-content: space-between;">
                <div class="kpi-card"><div class="kpi-number" id="masEntregadoProd">-</div><div class="kpi-label">Producto más entregado</div></div>
                <div class="kpi-card"><div class="kpi-number" id="masVencidoProd">-</div><div class="kpi-label">Producto más vencido (existencia vencida)</div></div>
                <div class="kpi-card"><div class="kpi-number" id="stockCriticoProd">-</div><div class="kpi-label">Producto con stock más crítico</div></div>
            </div>
            <div class="dashboard-tables"><h4>📋 Últimas entregas realizadas</h4><div class="table-wrapper" style="max-height: 300px;"><table id="ultimasEntregasTable"><thead>运转<th>Fecha</th><th>Área</th><th>Destinatario</th><th>Productos</th><th>Cantidad total</th></tr></thead><tbody></tbody></table></div></div>
        </div>
    </div>

    <!-- Inventario -->
    <div id="tabInventario" class="tab-content">
        <div class="card">
            <div class="umbral-bar"><label>⚠️ Alertar vencimiento (días):</label><input type="number" id="expiryDaysInput" min="1" value="30" style="width:90px;"><button class="btn-secondary" id="setExpiryDaysBtn">Aplicar</button><span>🔴 Vencidos | 🟡 Stock bajo / Próximo venc.</span></div>
            <div class="filter-bar">
                <select id="invSearchType"><option value="codigo">Código</option><option value="descripcion">Descripción</option><option value="tipo">Tipo</option><option value="estante">Estante</option></select>
                <input type="text" id="invSearchValue" placeholder="Buscar...">
                <button class="btn-secondary" id="btnBuscarInv">Buscar</button>
                <button class="btn-outline" id="btnLimpiarInv">Limpiar</button>
                <button class="btn-warning" id="btnProximosVencer">📅 Próx. vencer</button>
                <button class="btn-danger" id="btnBajoStockInv">⚠️ Bajo stock</button>
                <button class="btn-secondary" id="exportInventarioBtn">Exportar Excel</button>
                <button class="btn-success" id="importInventarioBtn">Importar Excel</button>
            </div>
            <div class="table-wrapper"><table id="inventoryTable"><thead>运转<th>Código</th><th>Descripción</th><th>Tipo</th><th>Existencia</th><th>Stock Mínimo</th><th>Vencimiento</th><th>Estante</th><th>Lote</th><th>Registro INVIMA</th><th>Acciones</th></tr></thead><tbody></tbody></table></div>
        </div>
        <div class="card">
            <h3>Agregar / Editar Producto</h3>
            <div class="form-grid">
                <div class="form-field"><label>Código *</label><input type="text" id="prodCodigo"></div>
                <div class="form-field"><label>Descripción *</label><input type="text" id="prodDescripcion"></div>
                <div class="form-field"><label>Tipo</label><select id="prodTipo"><option>Medicamento</option><option>Insumo</option></select></div>
                <div class="form-field"><label>Existencia</label><input type="number" id="prodExistencia" value="0"></div>
                <div class="form-field"><label>Stock Mínimo</label><input type="number" id="prodStockMinimo" value="0"></div>
                <div class="form-field"><label>Vencimiento</label><input type="date" id="prodVencimiento"></div>
                <div class="form-field"><label>Estante</label><select id="prodEstante"></select></div>
                <div class="form-field"><label>Lote</label><input type="text" id="prodLote" placeholder="Número de lote"></div>
                <div class="form-field"><label>Registro INVIMA</label><input type="text" id="prodRegInvima" placeholder="Registro sanitario INVIMA"></div>
            </div>
            <div style="display: flex; gap: 12px; margin-top: 20px;">
                <button class="btn-primary" id="addProductBtn">Agregar Producto</button>
                <button class="btn-outline" id="cancelEditBtn" style="display: none;">Cancelar Edición</button>
            </div>
        </div>
    </div>

    <!-- Movimientos (con destinatario y área) -->
    <div id="tabMovimientos" class="tab-content"><div class="card"><h3>Registrar Movimiento</h3><div class="form-grid"><div class="form-field"><label>Código *</label><input type="text" id="movCodigo"></div><div class="form-field"><label>Descripción</label><input type="text" id="movDescripcion"></div><div class="form-field"><label>Existencia actual</label><input type="text" id="movExistencia" readonly></div><div class="form-field"><label>Stock Mínimo</label><input type="number" id="movStockMinimo"></div><div class="form-field"><label>Tipo Movimiento</label><select id="movTipo"></select></div><div class="form-field"><label>Cantidad *</label><input type="number" id="movCantidad" min="1"></div><div class="form-field"><label>Tipo Producto</label><select id="movTipoProducto"></select></div><div class="form-field"><label>Lote</label><input id="movLote" placeholder="Lote"></div><div class="form-field"><label>Registro INVIMA</label><input id="movRegistroInvima" placeholder="Registro INVIMA"></div><div class="form-field"><label>Vencimiento</label><input type="date" id="movVencimiento"></div><div class="form-field"><label>Estante</label><select id="movEstante"></select></div></div><div style="display:flex; gap:12px; margin-top:20px;"><button class="btn-primary" id="registrarMovBtn">Registrar</button><button class="btn-outline" id="limpiarFormBtn">Limpiar</button><button class="btn-secondary" id="buscarProductoBtn">🔍 Buscar producto (flotante)</button></div></div>
        <div class="card"><h3>Historial de movimientos</h3><div style="display: flex; gap: 12px; margin-bottom: 16px;"><button class="btn-secondary" id="exportMovimientosBtn">Exportar Excel</button><button class="btn-success" id="importMovimientosBtn">Importar Excel</button></div><div class="table-wrapper"><table id="movimientosTable"><thead>运转<th>Fecha</th><th>Usuario</th><th>Tipo</th><th>Código</th><th>Descripción</th><th>Cantidad</th><th>Lote</th><th>Reg. Invima</th><th>Destinatario</th><th>Área</th><th>Acciones</th></tr></thead><tbody></tbody></table></div></div></div>

    <!-- Entregas (carrito con iconos editar/eliminar) -->
    <div id="tabEntregas" class="tab-content"><div class="card"><h3><i class="fas fa-truck-medical"></i> Dispensación a Áreas / Pacientes</h3><p>Seleccione productos, ajuste cantidades, complete datos del receptor y realice la entrega. Se imprimirá un recibo automáticamente.</p></div>
    <div class="card"><h4>1. Agregar productos al carrito</h4><div class="form-grid"><div class="form-field"><label>Producto</label><select id="selectProductoEntrega"><option value="">-- Seleccionar --</option></select></div><div class="form-field"><label>Cantidad</label><input type="number" id="cantidadEntrega" min="1" value="1"></div><div class="form-field"><label>Stock disponible</label><input type="text" id="stockDisponibleEntrega" readonly></div><div class="form-field" style="justify-content: flex-end; flex-direction: row; gap: 8px;"><button class="btn-primary" id="agregarAlCarritoBtn"><i class="fas fa-cart-plus"></i> Agregar</button><button class="btn-outline" id="limpiarFormularioEntregaBtn">Limpiar</button></div></div><div><h4>Carrito de entrega</h4><div id="cartContainer" style="max-height: 280px; overflow-y: auto; border:1px solid #e2e8f0; border-radius:20px; padding:12px; background: #f9fafb;"></div></div></div>
    <div class="card"><h4>2. Datos del receptor</h4><div class="form-grid"><div class="form-field"><label>Área / Servicio *</label><select id="areaDestino"></select><input type="text" id="otroArea" placeholder="Especificar otra área" style="margin-top:8px; display:none;"></div><div class="form-field"><label>Nombre del paciente / responsable *</label><input type="text" id="nombreDestinatario" placeholder="Ej: María López"></div></div><div style="display:flex; gap:12px; margin-top:20px;"><button class="btn-success" id="confirmarEntregaBtn"><i class="fas fa-check-circle"></i> Realizar entrega e imprimir recibo</button><button class="btn-outline" id="limpiarCarritoBtn">Vaciar carrito</button></div></div></div>

    <!-- Historial Entregas (botón Re-dispensar siempre visible) -->
    <div id="tabHistorialEntregas" class="tab-content"><div class="card"><h3>Historial de Entregas por Área</h3><div class="form-grid" style="margin-bottom:20px;"><div class="form-field"><label>Área</label><select id="filtroAreaHistorial"><option value="">Todas</option></select></div><div class="form-field"><label>Destinatario</label><input type="text" id="filtroDestinatarioHistorial"></div><div class="form-field"><label>Fecha desde</label><input type="date" id="filtroFechaDesdeHistorial"></div><div class="form-field"><label>Fecha hasta</label><input type="date" id="filtroFechaHastaHistorial"></div><div class="form-field"><button class="btn-secondary" id="aplicarFiltrosHistorial">Filtrar</button><button class="btn-outline" id="limpiarFiltrosHistorial">Limpiar</button></div></div><div class="table-wrapper"><table id="historialEntregasTable"><thead>运转<th>Fecha/Hora</th><th>Usuario</th><th>Área</th><th>Destinatario</th><th>Productos (Código - Cantidad)</th><th>Re-dispensar</th></tr></thead><tbody></tbody></table></div></div></div>

    <!-- Reportes -->
    <div id="tabReportes" class="tab-content"><div class="card"><h3>Reportes</h3><div style="display:flex; gap:12px; flex-wrap:wrap;"><button class="btn-secondary" id="reporteInventarioBtn">📄 Inventario</button><button class="btn-secondary" id="reporteMovimientosBtn">📄 Movimientos</button><button class="btn-warning" id="reporteVencimientoBtn">⚠️ Próx. vencer</button><button class="btn-danger" id="reporteBajoStockBtn">⚠️ Bajo stock</button><button class="btn-success" id="reporteEntregasBtn">📄 Entregas</button></div></div></div>

    <!-- Usuarios -->
    <div id="tabUsuarios" class="tab-content"><div class="card"><h3>Cambiar contraseña</h3><div class="form-grid"><div class="form-field"><label>Usuario</label><input id="userActual" readonly></div><div class="form-field"><label>Contraseña actual</label><input type="password" id="pwdActual"></div><div class="form-field"><label>Nueva contraseña</label><input type="password" id="pwdNueva"></div><div class="form-field"><label>Confirmar</label><input type="password" id="pwdConfirm"></div></div><button class="btn-primary" id="cambiarPwdBtn">Cambiar contraseña</button></div><div class="card"><h3>Gestión de usuarios</h3><button class="btn-secondary" id="toggleUsersBtn">Ver / ocultar usuarios</button><div id="usersList" style="display:none; margin-top:16px;"><div class="table-wrapper"><table id="usersTable"><thead>运转<th>Usuario</th><th>Nombre</th><th>Teléfono</th><th>Email</th><th>Rol</th><th>Estado</th><th>Acciones</th><tr></thead><tbody id="usersTableBody"></tbody></table></div></div><div id="adminAddUser" style="display:none; margin-top:20px;"><h4>➕ Agregar usuario</h4><div class="form-grid"><div class="form-field"><label>Usuario</label><input id="newUser"></div><div class="form-field"><label>Contraseña</label><input type="password" id="newPass"></div><div class="form-field"><label>Nombre completo</label><input id="newFullName"></div><div class="form-field"><label>Teléfono</label><input id="newPhone"></div><div class="form-field"><label>Email</label><input id="newEmail" type="email"></div><div class="form-field"><label>Rol</label><select id="newRole"><option>user</option><option>admin</option></select></div></div><button class="btn-secondary" id="addUserBtn">Crear usuario</button></div></div></div>

    <!-- Licencia -->
    <div id="tabLicencia" class="tab-content"><div class="card"><h3>Licencia y respaldo</h3><div id="licenseInfo"></div><button class="btn-primary" id="renewLicenseBtn">Renovar licencia</button><hr><button class="btn-secondary" id="exportAllDataBtn">Exportar datos (JSON)</button><button class="btn-warning" id="resetDataBtn">Restablecer datos</button></div></div>
</div>

<script>
    // ==================== CONFIGURACIÓN INICIAL ====================
    const STORAGE_CONFIG_AREAS = "app_config_areas";
    const STORAGE_CONFIG_TIPOS_PRODUCTO = "app_config_tipos_producto";
    const STORAGE_CONFIG_TIPOS_MOV = "app_config_tipos_movimiento";
    const STORAGE_CONFIG_ESTANTES = "app_config_estantes";
    const STORAGE_USERS = "farmapp_users";
    const STORAGE_PRODUCTOS = "farmapp_productos";
    const STORAGE_INSUMOS = "farmapp_insumos";
    const STORAGE_MOVEMENTS = "farmapp_movements";
    const STORAGE_EXPIRY_DAYS = "farmapp_expiry_days";
    const STORAGE_LICENSE_CHECK = "farmapp_license_last_check";
    const HABILITATION_PASSWORD = "Dacs1989***";

    let currentUser = null, currentUserFullName = "", currentUserRole = "";
    let tipoChartInstance = null, topEntregasChartInstance = null, topDestinatariosChartInstance = null;
    let cart = [];
    let editingProductCode = null;

    const defaultAreas = ["Urgencia", "Odontología", "Promotoras", "Otro"];
    const defaultTiposProducto = ["Medicamento", "Insumo"];
    const defaultTiposMovimiento = ["Entrada", "Salida"];
    const defaultEstantes = ["A1", "A2", "B1", "B2", "C1", "C2"];

    function loadConfig(key, defaultValue) { const stored = localStorage.getItem(key); return stored ? JSON.parse(stored) : defaultValue; }
    function saveConfig(key, value) { localStorage.setItem(key, JSON.stringify(value)); }
    function getAreas() { return loadConfig(STORAGE_CONFIG_AREAS, defaultAreas); }
    function getTiposProducto() { return loadConfig(STORAGE_CONFIG_TIPOS_PRODUCTO, defaultTiposProducto); }
    function getTiposMovimiento() { return loadConfig(STORAGE_CONFIG_TIPOS_MOV, defaultTiposMovimiento); }
    function getEstantes() { return loadConfig(STORAGE_CONFIG_ESTANTES, defaultEstantes); }
    function getExpiryDays() { return parseInt(localStorage.getItem(STORAGE_EXPIRY_DAYS) || "30"); }
    function setExpiryDays(days) { localStorage.setItem(STORAGE_EXPIRY_DAYS, days); renderInventory(); updateNotificationBell(); }

    // ==================== CRIPTO ====================
    async function hashPassword(pw) { const encoder = new TextEncoder(); const data = encoder.encode(pw); const hash = await crypto.subtle.digest('SHA-256', data); return Array.from(new Uint8Array(hash)).map(b=>b.toString(16).padStart(2,'0')).join(''); }
    async function verifyPassword(plain, hash) { return (await hashPassword(plain)) === hash; }

    // ==================== DATOS ====================
    function getInventory() { return [...(JSON.parse(localStorage.getItem(STORAGE_PRODUCTOS)||"[]")), ...(JSON.parse(localStorage.getItem(STORAGE_INSUMOS)||"[]"))]; }
    function saveProduct(p) { const key = p.tipoProducto === "Medicamento" ? STORAGE_PRODUCTOS : STORAGE_INSUMOS; const list = JSON.parse(localStorage.getItem(key)||"[]"); const idx = list.findIndex(i=>i.codigo===p.codigo); if(idx!==-1) list[idx]=p; else list.push(p); localStorage.setItem(key,JSON.stringify(list)); }
    function getUsers() { return JSON.parse(localStorage.getItem(STORAGE_USERS)||"{}"); }
    function saveUsers(u) { localStorage.setItem(STORAGE_USERS,JSON.stringify(u)); }
    function getMovements() { return JSON.parse(localStorage.getItem(STORAGE_MOVEMENTS)||"[]"); }
    function saveMovements(m) { localStorage.setItem(STORAGE_MOVEMENTS,JSON.stringify(m)); }

    async function registrarMovimiento(data, extra={}) {
        let inv = getInventory(); let idx = inv.findIndex(p=>p.codigo===data.codigo);
        if(idx===-1 && data.tipo==="Salida") throw new Error(`Producto ${data.codigo} no existe`);
        if(idx===-1 && data.tipo==="Entrada") {
            const newP={ codigo:data.codigo, descripcion:data.descripcion||"Nuevo", tipoProducto:data.tipoProducto||"Medicamento", stockMinimo:parseInt(data.stockMinimo)||0, existencia:data.cantidad, lote:data.lote||"", registroInvima:data.registroInvima||"", vencimiento:data.vencimiento||"", estante:data.estante||"" };
            saveProduct(newP);
        } else {
            let prod = inv[idx];
            if(data.tipo==="Entrada") prod.existencia += data.cantidad;
            else { if(prod.existencia<data.cantidad) throw new Error(`Stock insuficiente`); prod.existencia -= data.cantidad; }
            if(data.descripcion) prod.descripcion=data.descripcion;
            if(data.lote) prod.lote=data.lote;
            if(data.registroInvima) prod.registroInvima=data.registroInvima;
            if(data.vencimiento) prod.vencimiento=data.vencimiento;
            if(data.estante !== undefined) prod.estante=data.estante;
            if(data.stockMinimo !== undefined) prod.stockMinimo = parseInt(data.stockMinimo);
            saveProduct(prod);
        }
        const movs = getMovements();
        const newMov = { id:Date.now(), fecha:new Date().toISOString(), usuario:currentUser, tipo:data.tipo, codigo:data.codigo, descripcion:data.descripcion||"", cantidad:data.cantidad, lote:data.lote||"", registroInvima:data.registroInvima||"", estante:data.estante||"", ...extra };
        movs.push(newMov); saveMovements(movs);
        renderDashboard(); renderInventory(); renderMovementsTable(); updateNotificationBell();
        return newMov;
    }

    // ==================== DASHBOARD ====================
    function renderDashboard() {
        const inventory = getInventory();
        const medicamentos = inventory.filter(p => p.tipoProducto === "Medicamento").length;
        const insumos = inventory.filter(p => p.tipoProducto === "Insumo").length;
        const movements = getMovements();
        const entregas = movements.filter(m => m.tipo === "Salida" && m.destinatario);
        document.getElementById("kpiContainer").innerHTML = `
            <div class="kpi-card"><div class="kpi-number">${inventory.length}</div><div class="kpi-label">Total Productos</div></div>
            <div class="kpi-card"><div class="kpi-number">${medicamentos}</div><div class="kpi-label">Medicamentos</div></div>
            <div class="kpi-card"><div class="kpi-number">${insumos}</div><div class="kpi-label">Insumos</div></div>
            <div class="kpi-card"><div class="kpi-number">${movements.length}</div><div class="kpi-label">Movimientos</div></div>
            <div class="kpi-card"><div class="kpi-number">${entregas.length}</div><div class="kpi-label">Entregas</div></div>
        `;
        const consumo = {};
        entregas.forEach(m => { consumo[m.codigo] = (consumo[m.codigo] || 0) + m.cantidad; });
        let maxConsumo = 0, maxProd = null;
        for (let [cod, cant] of Object.entries(consumo)) if (cant > maxConsumo) { maxConsumo = cant; maxProd = cod; }
        const prodMasEntregado = maxProd ? (inventory.find(p=>p.codigo===maxProd)?.descripcion || maxProd) : "Ninguno";
        document.getElementById("masEntregadoProd").innerText = prodMasEntregado;
        const hoy = new Date(); hoy.setHours(0,0,0,0);
        const vencidos = inventory.filter(p => p.vencimiento && new Date(p.vencimiento) < hoy);
        let maxVencido = null, maxCantVencida = 0;
        vencidos.forEach(p => { if(p.existencia > maxCantVencida) { maxCantVencida = p.existencia; maxVencido = p.descripcion; } });
        document.getElementById("masVencidoProd").innerText = maxVencido ? `${maxVencido} (${maxCantVencida} und)` : "Sin vencidos";
        let critico = null, minRatio = Infinity;
        inventory.forEach(p => { if(p.stockMinimo > 0) { let ratio = p.existencia / p.stockMinimo; if(ratio < minRatio) { minRatio = ratio; critico = p.descripcion; } } });
        document.getElementById("stockCriticoProd").innerText = critico ? `${critico} (ratio ${minRatio.toFixed(2)})` : "N/A";
        const ctxTipo = document.getElementById('tipoChart').getContext('2d');
        if(tipoChartInstance) tipoChartInstance.destroy();
        tipoChartInstance = new Chart(ctxTipo, { type: 'pie', data: { labels: ['Medicamentos', 'Insumos'], datasets: [{ data: [medicamentos, insumos], backgroundColor: ['#3b82f6', '#10b981'] }] }, options: { responsive: true } });
        const top5Prod = Object.entries(consumo).sort((a,b)=>b[1]-a[1]).slice(0,5);
        const labelsProd = top5Prod.map(([cod])=> { let prod = inventory.find(p=>p.codigo===cod); return prod ? prod.descripcion.substring(0,20) : cod; });
        const dataProd = top5Prod.map(([,cant])=>cant);
        const ctxTopProd = document.getElementById('topEntregasChart').getContext('2d');
        if(topEntregasChartInstance) topEntregasChartInstance.destroy();
        topEntregasChartInstance = new Chart(ctxTopProd, { type: 'bar', data: { labels: labelsProd, datasets: [{ label: 'Cantidad entregada', data: dataProd, backgroundColor: '#f59e0b' }] }, options: { responsive: true, scales: { y: { beginAtZero: true } } } });
        const consumoDestinatario = {};
        entregas.forEach(m => { const dest = m.destinatario; if(dest) consumoDestinatario[dest] = (consumoDestinatario[dest] || 0) + m.cantidad; });
        const top5Dest = Object.entries(consumoDestinatario).sort((a,b)=>b[1]-a[1]).slice(0,5);
        const labelsDest = top5Dest.map(([dest])=>dest.length>20?dest.substring(0,17)+'...':dest);
        const dataDest = top5Dest.map(([,cant])=>cant);
        const ctxDest = document.getElementById('topDestinatariosChart').getContext('2d');
        if(topDestinatariosChartInstance) topDestinatariosChartInstance.destroy();
        topDestinatariosChartInstance = new Chart(ctxDest, { type: 'bar', data: { labels: labelsDest, datasets: [{ label: 'Cantidad total recibida', data: dataDest, backgroundColor: '#10b981' }] }, options: { responsive: true, scales: { y: { beginAtZero: true } } } });
        const ultimas = entregas.sort((a,b)=>new Date(b.fecha)-new Date(a.fecha)).slice(0,10);
        const tbodyUlt = document.querySelector("#ultimasEntregasTable tbody");
        tbodyUlt.innerHTML = "";
        ultimas.forEach(m => { let row = tbodyUlt.insertRow(); row.insertCell(0).innerText = new Date(m.fecha).toLocaleString(); row.insertCell(1).innerText = m.area || "—"; row.insertCell(2).innerText = m.destinatario || "—"; row.insertCell(3).innerText = `${m.codigo} - ${m.descripcion}`; row.insertCell(4).innerText = m.cantidad; });
    }

    // ==================== INVENTARIO ====================
    function renderInventory(filteredList = null) {
        let inv = filteredList || getInventory();
        const searchType = document.getElementById("invSearchType").value;
        const searchVal = document.getElementById("invSearchValue").value.trim().toLowerCase();
        if(searchVal && !filteredList) {
            inv = inv.filter(p => {
                if(searchType === "codigo") return p.codigo.toLowerCase().includes(searchVal);
                if(searchType === "descripcion") return p.descripcion.toLowerCase().includes(searchVal);
                if(searchType === "tipo") return p.tipoProducto.toLowerCase().includes(searchVal);
                if(searchType === "estante") return (p.estante||"").toLowerCase().includes(searchVal);
                return true;
            });
        }
        const tbody = document.querySelector("#inventoryTable tbody");
        tbody.innerHTML = "";
        const hoy = new Date(); hoy.setHours(0,0,0,0);
        const expiryLimit = new Date(); expiryLimit.setDate(hoy.getDate() + getExpiryDays());
        inv.forEach(p => {
            const row = tbody.insertRow();
            const isExpired = p.vencimiento && new Date(p.vencimiento) < hoy;
            const isLowStock = p.existencia <= p.stockMinimo;
            const isNearExpiry = p.vencimiento && new Date(p.vencimiento) >= hoy && new Date(p.vencimiento) <= expiryLimit;
            if(isExpired) row.classList.add("expired-row");
            else if(isLowStock || isNearExpiry) row.style.backgroundColor = "#fff3e0";
            row.insertCell(0).innerText = p.codigo;
            row.insertCell(1).innerText = p.descripcion;
            row.insertCell(2).innerText = p.tipoProducto;
            row.insertCell(3).innerHTML = `<span class="stock-highlight">${p.existencia}</span>`;
            row.insertCell(4).innerText = p.stockMinimo;
            let vencHtml = p.vencimiento || "N/A";
            if(isExpired) vencHtml = `<span style="color:#dc2626;">${p.vencimiento} (VENCIDO)</span>`;
            else if(isNearExpiry) vencHtml = `<span style="color:#f59e0b;">${p.vencimiento} (Próximo)</span>`;
            row.insertCell(5).innerHTML = vencHtml;
            row.insertCell(6).innerText = p.estante || "-";
            row.insertCell(7).innerText = p.lote || "-";
            row.insertCell(8).innerText = p.registroInvima || "-";
            const acc = row.insertCell(9);
            if(currentUserRole==="admin"||currentUserRole==="root") {
                const editBtn = document.createElement("button"); editBtn.innerHTML = '<i class="fas fa-edit"></i>'; editBtn.className = "btn-icon btn-icon-edit"; editBtn.title = "Editar";
                editBtn.onclick = () => mostrarModalEdicionProducto(p);
                const delBtn = document.createElement("button"); delBtn.innerHTML = '<i class="fas fa-trash-alt"></i>'; delBtn.className = "btn-icon btn-icon-delete"; delBtn.title = "Eliminar";
                delBtn.onclick = () => { if(confirm("¿Eliminar?")) { const key = p.tipoProducto==="Medicamento"?STORAGE_PRODUCTOS:STORAGE_INSUMOS; const list=JSON.parse(localStorage.getItem(key)); localStorage.setItem(key,JSON.stringify(list.filter(i=>i.codigo!==p.codigo))); renderInventory(); renderMovementsTable(); renderDashboard(); } };
                acc.append(editBtn, delBtn);
            } else acc.innerText = "---";
        });
        updateNotificationBell();
        renderDashboard();
    }

    function mostrarModalEdicionProducto(prod) {
        if(currentUserRole!=="admin" && currentUserRole!=="root") return;
        const modal = document.createElement("div"); modal.className = "modal";
        modal.innerHTML = `<div class="modal-content"><h3>Editar producto: ${prod.codigo}</h3><div class="form-grid"><div class="form-field"><label>Descripción</label><input id="editDesc" value="${prod.descripcion.replace(/"/g, '&quot;')}"></div><div class="form-field"><label>Existencia</label><input type="number" id="editExist" value="${prod.existencia}"></div><div class="form-field"><label>Stock Mínimo</label><input type="number" id="editStockMin" value="${prod.stockMinimo}"></div><div class="form-field"><label>Lote</label><input id="editLote" value="${(prod.lote||'').replace(/"/g, '&quot;')}"></div><div class="form-field"><label>Registro Invima</label><input id="editRegInv" value="${(prod.registroInvima||'').replace(/"/g, '&quot;')}"></div><div class="form-field"><label>Vencimiento</label><input type="date" id="editVen" value="${prod.vencimiento||''}"></div><div class="form-field"><label>Estante</label><select id="editEstante">${getEstantes().map(e=>`<option value="${e}" ${prod.estante===e?'selected':''}>${e}</option>`).join('')}<option value="">-- Sin estante --</option></select></div></div><div style="margin-top:20px;"><button class="btn-primary" id="saveEditProd">Guardar</button><button class="btn-outline" id="closeModalEdit">Cancelar</button></div></div>`;
        document.body.appendChild(modal);
        document.getElementById("saveEditProd").onclick = () => {
            prod.descripcion = document.getElementById("editDesc").value;
            prod.existencia = parseInt(document.getElementById("editExist").value);
            prod.stockMinimo = parseInt(document.getElementById("editStockMin").value);
            prod.lote = document.getElementById("editLote").value;
            prod.registroInvima = document.getElementById("editRegInv").value;
            prod.vencimiento = document.getElementById("editVen").value;
            prod.estante = document.getElementById("editEstante").value;
            saveProduct(prod);
            renderInventory(); renderMovementsTable(); renderDashboard();
            modal.remove();
        };
        document.getElementById("closeModalEdit").onclick = () => modal.remove();
    }

    function cancelEditProduct() {
        editingProductCode = null;
        document.getElementById("prodCodigo").value = "";
        document.getElementById("prodDescripcion").value = "";
        document.getElementById("prodTipo").value = "Medicamento";
        document.getElementById("prodExistencia").value = "0";
        document.getElementById("prodStockMinimo").value = "0";
        document.getElementById("prodVencimiento").value = "";
        document.getElementById("prodEstante").value = "";
        document.getElementById("prodLote").value = "";
        document.getElementById("prodRegInvima").value = "";
        document.getElementById("addProductBtn").innerText = "Agregar Producto";
        document.getElementById("cancelEditBtn").style.display = "none";
    }

    function addOrUpdateProduct() {
        const codigo = document.getElementById("prodCodigo").value.trim();
        if(!codigo) return alert("El código es obligatorio");
        const descripcion = document.getElementById("prodDescripcion").value.trim();
        if(!descripcion) return alert("La descripción es obligatoria");
        const tipoProducto = document.getElementById("prodTipo").value;
        let existencia = parseInt(document.getElementById("prodExistencia").value);
        if(isNaN(existencia)) existencia = 0;
        const stockMinimo = parseInt(document.getElementById("prodStockMinimo").value) || 0;
        const vencimiento = document.getElementById("prodVencimiento").value;
        const estante = document.getElementById("prodEstante").value;
        const lote = document.getElementById("prodLote").value;
        const registroInvima = document.getElementById("prodRegInvima").value;
        const product = { codigo, descripcion, tipoProducto, existencia, stockMinimo, vencimiento, estante, lote, registroInvima };
        const inventory = getInventory();
        if(editingProductCode) {
            const idx = inventory.findIndex(p => p.codigo === editingProductCode);
            if(idx !== -1) { const updatedProduct = { ...inventory[idx], ...product }; saveProduct(updatedProduct); }
            else saveProduct(product);
        } else {
            if(inventory.find(p => p.codigo === codigo)) return alert("Ya existe un producto con ese código");
            saveProduct(product);
        }
        renderInventory();
        cancelEditProduct();
    }

    // ==================== MOVIMIENTOS ====================
    function renderMovementsTable() {
        const movs = getMovements();
        const tbody = document.querySelector("#movimientosTable tbody");
        tbody.innerHTML = "";
        movs.slice().reverse().forEach(m => {
            const row = tbody.insertRow();
            row.insertCell(0).innerText = new Date(m.fecha).toLocaleString();
            row.insertCell(1).innerText = m.usuario;
            row.insertCell(2).innerText = m.tipo;
            row.insertCell(3).innerText = m.codigo;
            row.insertCell(4).innerText = m.descripcion;
            row.insertCell(5).innerText = m.cantidad;
            row.insertCell(6).innerText = m.lote || "";
            row.insertCell(7).innerText = m.registroInvima || "";
            row.insertCell(8).innerText = m.destinatario || "—";
            row.insertCell(9).innerText = m.area || "—";
            if(currentUserRole==="admin"||currentUserRole==="root") {
                const cellActions = row.insertCell(10);
                const editBtn = document.createElement("button"); editBtn.innerHTML = '<i class="fas fa-edit"></i>'; editBtn.className = "btn-icon btn-icon-edit"; editBtn.title = "Editar";
                editBtn.onclick = () => mostrarModalEditarMovimiento(m);
                const delBtn = document.createElement("button"); delBtn.innerHTML = '<i class="fas fa-trash-alt"></i>'; delBtn.className = "btn-icon btn-icon-delete"; delBtn.title = "Eliminar";
                delBtn.onclick = () => { if(confirm("¿Eliminar movimiento?")){ let mvs=getMovements(); mvs=mvs.filter(mm=>mm.id!==m.id); saveMovements(mvs); renderMovementsTable(); renderInventory(); renderDashboard(); } };
                cellActions.append(editBtn, delBtn);
            } else row.insertCell(10).innerText = "---";
        });
        renderDashboard();
    }

    function mostrarModalEditarMovimiento(mov) {
        if(currentUserRole!=="admin" && currentUserRole!=="root") return;
        const modal = document.createElement("div"); modal.className = "modal";
        modal.innerHTML = `<div class="modal-content"><h3>Editar Movimiento ID: ${mov.id}</h3><div class="form-grid"><div class="form-field"><label>Descripción</label><input id="editMovDesc" value="${mov.descripcion.replace(/"/g, '&quot;')}"></div><div class="form-field"><label>Lote</label><input id="editMovLote" value="${(mov.lote||'').replace(/"/g, '&quot;')}"></div><div class="form-field"><label>Registro Invima</label><input id="editMovRegInv" value="${(mov.registroInvima||'').replace(/"/g, '&quot;')}"></div><div class="form-field"><label>Estante</label><select id="editMovEstante">${getEstantes().map(e=>`<option value="${e}" ${mov.estante===e?'selected':''}>${e}</option>`).join('')}<option value="">-- Sin estante --</option></select></div></div><div style="margin-top:20px;"><button class="btn-primary" id="saveEditMov">Guardar</button><button class="btn-outline" id="closeEditMov">Cancelar</button></div></div>`;
        document.body.appendChild(modal);
        document.getElementById("saveEditMov").onclick = () => {
            mov.descripcion = document.getElementById("editMovDesc").value;
            mov.lote = document.getElementById("editMovLote").value;
            mov.registroInvima = document.getElementById("editMovRegInv").value;
            mov.estante = document.getElementById("editMovEstante").value;
            const movements = getMovements();
            const idx = movements.findIndex(m => m.id === mov.id);
            if(idx !== -1) movements[idx] = mov;
            saveMovements(movements);
            renderMovementsTable();
            modal.remove();
            const prod = getInventory().find(p => p.codigo === mov.codigo);
            if(prod) { if(mov.lote) prod.lote = mov.lote; if(mov.registroInvima) prod.registroInvima = mov.registroInvima; if(mov.estante) prod.estante = mov.estante; if(mov.descripcion) prod.descripcion = mov.descripcion; saveProduct(prod); renderInventory(); }
            renderDashboard();
        };
        document.getElementById("closeEditMov").onclick = () => modal.remove();
    }

    function limpiarFormularioMovimiento() {
        document.getElementById("movCodigo").value = "";
        document.getElementById("movDescripcion").value = "";
        document.getElementById("movExistencia").value = "";
        document.getElementById("movStockMinimo").value = "";
        document.getElementById("movCantidad").value = "";
        document.getElementById("movLote").value = "";
        document.getElementById("movRegistroInvima").value = "";
        document.getElementById("movVencimiento").value = "";
        document.getElementById("movEstante").value = "";
    }

    function openProductSearchModal() {
        const modal = document.createElement("div"); modal.className = "modal";
        modal.innerHTML = `<div class="modal-content"><h3><i class="fas fa-search"></i> Buscar producto</h3><input type="text" id="searchProductInput" placeholder="🔎 Código o descripción..." style="width:100%; margin-bottom:16px; padding:12px; border-radius:28px; border:1px solid #cbd5e1;"><div class="table-wrapper"><table id="searchProductTable" style="min-width:500px;"><thead>运转<th>Código</th><th>Descripción</th><th>Tipo</th><th>Stock</th><th>Vencimiento</th><th>Estante</th><th></th></tr></thead><tbody></tbody></table></div><div style="margin-top:16px; text-align:right;"><button class="btn-outline" id="closeSearchModal">Cerrar</button></div></div>`;
        document.body.appendChild(modal);
        const tbody = modal.querySelector("#searchProductTable tbody");
        function renderSearch(term) { const inv = getInventory(); const filtered = inv.filter(p=>p.codigo.toLowerCase().includes(term) || p.descripcion.toLowerCase().includes(term)); tbody.innerHTML = filtered.map(p=>`<tr><td>${p.codigo}</td><td>${p.descripcion}</td><td>${p.tipoProducto}</td><td>${p.existencia}</td><td>${p.vencimiento||"N/A"}</td><td>${p.estante||"-"}</td><td><button class="btn-secondary btn-icon" data-cod="${p.codigo}">Seleccionar</button></td></tr>`).join(''); modal.querySelectorAll("[data-cod]").forEach(btn=>{ btn.onclick = () => { const cod = btn.getAttribute("data-cod"); const prod = getInventory().find(p=>p.codigo===cod); if(prod){ document.getElementById("movCodigo").value = prod.codigo; document.getElementById("movDescripcion").value = prod.descripcion; document.getElementById("movExistencia").value = prod.existencia; document.getElementById("movStockMinimo").value = prod.stockMinimo; document.getElementById("movTipoProducto").value = prod.tipoProducto; document.getElementById("movLote").value = prod.lote||""; document.getElementById("movRegistroInvima").value = prod.registroInvima||""; document.getElementById("movVencimiento").value = prod.vencimiento||""; document.getElementById("movEstante").value = prod.estante||""; modal.remove(); } }; }); }
        const input = modal.querySelector("#searchProductInput"); input.addEventListener("input", e=> renderSearch(e.target.value.toLowerCase())); renderSearch("");
        modal.querySelector("#closeSearchModal").onclick = () => modal.remove();
    }

    // ==================== ENTREGAS (CARRITO CON ICONOS) ====================
    function actualizarStockItem(item) { const prod = getInventory().find(p=>p.codigo===item.codigo); if(prod) item.stockActual = prod.existencia; else item.stockActual = 0; return item; }
    function renderCart() {
        const container = document.getElementById("cartContainer");
        if(!container) return;
        if(cart.length===0){ container.innerHTML="<p style='text-align:center; color:gray;'>🛒 Carrito vacío</p>"; return; }
        cart = cart.map(item => actualizarStockItem(item));
        container.innerHTML = cart.map((item,idx)=>`
            <div class="cart-item">
                <div class="cart-item-info">
                    <strong>${item.codigo}</strong> - ${item.descripcion}
                    <span>Cantidad: ${item.cantidad}</span>
                    <span>📦 Stock disp: ${item.stockActual}</span>
                </div>
                <div class="cart-item-actions">
                    <button class="btn-icon btn-icon-edit" data-edit="${idx}" title="Editar cantidad"><i class="fas fa-edit"></i></button>
                    <button class="btn-icon btn-icon-delete" data-del="${idx}" title="Eliminar producto"><i class="fas fa-trash-alt"></i></button>
                </div>
            </div>
        `).join('');
        document.querySelectorAll("[data-edit]").forEach(btn=>{ btn.onclick = () => { const i = parseInt(btn.getAttribute("data-edit")); const item = cart[i]; const prodActual = getInventory().find(p=>p.codigo===item.codigo); const stockActual = prodActual ? prodActual.existencia : 0; const nuevaCant = prompt(`Editar cantidad para ${item.codigo}\nStock disponible AHORA: ${stockActual}\nCantidad actual: ${item.cantidad}`, item.cantidad); if(nuevaCant && !isNaN(parseInt(nuevaCant))) { const newCant = parseInt(nuevaCant); if(newCant <= 0) return alert("Cantidad inválida"); if(newCant > stockActual) return alert(`No puede entregar más de ${stockActual} (stock actual)`); item.cantidad = newCant; item.stockActual = stockActual; renderCart(); } }; });
        document.querySelectorAll("[data-del]").forEach(btn=>{ btn.onclick = () => { const i = parseInt(btn.getAttribute("data-del")); cart.splice(i,1); renderCart(); }; });
    }

    async function realizarEntrega(items, destinatario, area) {
        if(!items.length) throw new Error("Carrito vacío");
        if(!destinatario.trim()) throw new Error("Nombre requerido");
        const deliveryId = Date.now();
        for(let item of items) {
            await registrarMovimiento({ codigo:item.codigo, cantidad:item.cantidad, tipo:"Salida", descripcion:item.descripcion, tipoProducto:item.tipoProducto, lote:item.lote, registroInvima:item.registroInvima, vencimiento:item.vencimiento, estante:item.estante, stockMinimo:0 }, { destinatario, area, deliveryId });
        }
    }

    function imprimirRecibo(destinatario, area, items) {
        const win = window.open();
        win.document.write(`<html><head><title>Recibo Entrega</title><style>body{font-family:monospace;padding:20px;} .recibo{border:1px solid #ccc;padding:20px;max-width:400px;margin:auto;}</style></head><body><div class="recibo"><h3>FARMACIA ESE HOSPITAL LOCAL SAN ZENON</h3><p><strong>Fecha:</strong> ${new Date().toLocaleString()}</p><p><strong>Entregó:</strong> ${currentUserFullName}</p><p><strong>Área:</strong> ${area}</p><p><strong>Destinatario:</strong> ${destinatario}</p><hr><table style="width:100%"><tr><th>Producto</th><th>Cant</th></tr>${items.map(i=>`<tr><td>${i.codigo} - ${i.descripcion}</td><td style="text-align:center">${i.cantidad}</td></tr>`).join('')}</table><hr><p>Total ítems: ${items.reduce((s,i)=>s+i.cantidad,0)}</p><p style="text-align:center">--- Válido como comprobante ---</p></div><script>window.print();setTimeout(()=>window.close(),500);<\/script></body></html>`);
        win.document.close();
    }

    function cargarSelectProductosEntrega() {
        const inv = getInventory();
        const select = document.getElementById("selectProductoEntrega");
        select.innerHTML = '<option value="">-- Seleccionar --</option>';
        inv.forEach(p=>{ const opt = document.createElement("option"); opt.value = p.codigo; opt.textContent = `${p.codigo} - ${p.descripcion} (Stock:${p.existencia})`; select.appendChild(opt); });
        select.onchange = () => { const prod = inv.find(p=>p.codigo===select.value); document.getElementById("stockDisponibleEntrega").value = prod ? prod.existencia : ""; };
    }

    function setupEventosEntrega() {
        document.getElementById("agregarAlCarritoBtn").onclick = () => {
            const cod = document.getElementById("selectProductoEntrega").value;
            if(!cod) return alert("Seleccione producto");
            const cant = parseInt(document.getElementById("cantidadEntrega").value);
            if(isNaN(cant)||cant<=0) return alert("Cantidad válida");
            const prod = getInventory().find(p=>p.codigo===cod);
            if(!prod) return alert("Producto no existe");
            if(prod.existencia < cant) return alert(`Stock insuficiente: solo ${prod.existencia}`);
            const exist = cart.find(i=>i.codigo===cod);
            if(exist) {
                if(exist.cantidad + cant > prod.existencia) return alert(`No puede exceder el stock actual (${prod.existencia})`);
                exist.cantidad += cant;
                exist.stockActual = prod.existencia;
            } else {
                cart.push({ codigo:prod.codigo, descripcion:prod.descripcion, cantidad:cant, tipoProducto:prod.tipoProducto, lote:prod.lote, registroInvima:prod.registroInvima, vencimiento:prod.vencimiento, estante:prod.estante, stockActual:prod.existencia });
            }
            renderCart();
            document.getElementById("cantidadEntrega").value = 1;
            document.getElementById("stockDisponibleEntrega").value = "";
        };
        document.getElementById("limpiarFormularioEntregaBtn").onclick = () => { document.getElementById("selectProductoEntrega").value = ""; document.getElementById("cantidadEntrega").value = 1; document.getElementById("stockDisponibleEntrega").value = ""; };
        document.getElementById("confirmarEntregaBtn").onclick = async () => {
            let area = document.getElementById("areaDestino").value;
            if(area === "Otro") area = document.getElementById("otroArea").value.trim() || "Otro";
            const dest = document.getElementById("nombreDestinatario").value.trim();
            if(cart.length===0) return alert("Carrito vacío");
            if(!dest) return alert("Ingrese nombre del destinatario");
            try {
                await realizarEntrega(cart, dest, area);
                imprimirRecibo(dest, area, cart);
                alert("✅ Entrega realizada. Stock actualizado.");
                cart = [];
                renderCart();
                document.getElementById("nombreDestinatario").value = "";
                document.getElementById("otroArea").value = "";
                renderInventory();
                renderMovementsTable();
                renderHistorialEntregas();
                updateNotificationBell();
                cargarSelectProductosEntrega();
                renderDashboard();
            } catch(e){ alert("Error: "+e.message); }
        };
        document.getElementById("limpiarCarritoBtn").onclick = () => { cart = []; renderCart(); };
        document.getElementById("areaDestino").addEventListener("change", function(){ document.getElementById("otroArea").style.display = this.value === "Otro" ? "block" : "none"; });
        cargarSelectProductosEntrega();
    }

    // ==================== HISTORIAL ENTREGAS (RE-DISPENSAR SIEMPRE VISIBLE) ====================
    function getEntregas() {
        const movs = getMovements().filter(m => m.tipo === "Salida" && m.destinatario && m.area);
        const grupos = new Map();
        movs.forEach(m => { const key = m.deliveryId ? `delivery_${m.deliveryId}` : `single_${m.id}`; if(!grupos.has(key)) grupos.set(key, []); grupos.get(key).push(m); });
        const pedidos = [];
        for(let [key, items] of grupos.entries()) { pedidos.push({ fecha: items[0].fecha, usuario: items[0].usuario, area: items[0].area, destinatario: items[0].destinatario, productos: items.map(i => ({ codigo: i.codigo, cantidad: i.cantidad, descripcion: i.descripcion, tipoProducto: i.tipoProducto, lote: i.lote, registroInvima: i.registroInvima, vencimiento: i.vencimiento, estante: i.estante })), deliveryId: key }); }
        pedidos.sort((a,b) => new Date(b.fecha) - new Date(a.fecha));
        return pedidos;
    }

    function renderHistorialEntregas() {
        let pedidos = getEntregas();
        const areaFiltro = document.getElementById("filtroAreaHistorial").value;
        const destFiltro = document.getElementById("filtroDestinatarioHistorial").value.toLowerCase();
        const desde = document.getElementById("filtroFechaDesdeHistorial").value;
        const hasta = document.getElementById("filtroFechaHastaHistorial").value;
        if(areaFiltro) pedidos = pedidos.filter(p => p.area === areaFiltro);
        if(destFiltro) pedidos = pedidos.filter(p => p.destinatario.toLowerCase().includes(destFiltro));
        if(desde) pedidos = pedidos.filter(p => new Date(p.fecha) >= new Date(desde));
        if(hasta) pedidos = pedidos.filter(p => new Date(p.fecha) <= new Date(hasta));
        const tbody = document.querySelector("#historialEntregasTable tbody");
        tbody.innerHTML = "";
        pedidos.forEach(pedido => {
            const row = tbody.insertRow();
            row.insertCell(0).innerText = new Date(pedido.fecha).toLocaleString();
            row.insertCell(1).innerText = pedido.usuario;
            row.insertCell(2).innerText = pedido.area;
            row.insertCell(3).innerText = pedido.destinatario;
            row.insertCell(4).innerText = pedido.productos.map(prod => `${prod.codigo} x${prod.cantidad}`).join(', ');
            const btn = document.createElement("button"); 
            btn.innerHTML = "<i class='fas fa-redo-alt'></i> Re-dispensar";
            btn.className = "btn-redispensar";
            btn.style.cursor = "pointer";
            btn.onclick = () => {
                cart = pedido.productos.map(prod => {
                    const prodActual = getInventory().find(p => p.codigo === prod.codigo);
                    return { codigo: prod.codigo, descripcion: prod.descripcion, cantidad: prod.cantidad, tipoProducto: prod.tipoProducto, lote: prod.lote, registroInvima: prod.registroInvima, vencimiento: prod.vencimiento, estante: prod.estante, stockActual: prodActual ? prodActual.existencia : 0 };
                });
                renderCart();
                document.getElementById("nombreDestinatario").value = pedido.destinatario;
                const areaSel = document.getElementById("areaDestino");
                if(areaSel.querySelector(`option[value="${pedido.area}"]`)) areaSel.value = pedido.area;
                else areaSel.value = "Otro";
                document.getElementById("otroArea").value = pedido.area === "Otro" ? pedido.area : "";
                document.getElementById("otroArea").style.display = areaSel.value === "Otro" ? "block" : "none";
                alert("Productos cargados en el carrito. Puede modificar cantidades.");
                document.querySelector(".tab-btn[data-tab='entregas']").click();
            };
            row.insertCell(5).appendChild(btn);
        });
        renderDashboard();
    }

    // ==================== USUARIOS ====================
    function renderUsersTable() {
        const users = getUsers();
        const tbody = document.getElementById("usersTableBody");
        tbody.innerHTML = "";
        let usuariosAMostrar = users;
        if(currentUserRole === "user") { usuariosAMostrar = { [currentUser]: users[currentUser] }; }
        for(let [user, data] of Object.entries(usuariosAMostrar)) {
            const row = tbody.insertRow();
            if(data.enabled === false) row.classList.add("disabled-user-row");
            row.insertCell(0).innerText = user;
            row.insertCell(1).innerText = data.fullName || user;
            row.insertCell(2).innerText = data.phone || "";
            row.insertCell(3).innerText = data.email || "";
            row.insertCell(4).innerText = data.role;
            const cellEstado = row.insertCell(5);
            if(currentUserRole === "admin" || currentUserRole === "root") {
                const chk = document.createElement("input"); chk.type = "checkbox"; chk.checked = data.enabled !== false; chk.className = "user-enabled-checkbox";
                chk.onchange = () => { const newUsers = getUsers(); newUsers[user].enabled = chk.checked; saveUsers(newUsers); renderUsersTable(); if(user === currentUser && !chk.checked) { alert("Usuario deshabilitado, será desconectado."); sessionStorage.clear(); location.reload(); } };
                cellEstado.appendChild(chk);
            } else cellEstado.innerText = data.enabled === false ? "Inactivo" : "Activo";
            const cellAcc = row.insertCell(6);
            if(currentUserRole === "admin" || currentUserRole === "root") {
                const editBtn = document.createElement("button"); editBtn.innerHTML = '<i class="fas fa-edit"></i>'; editBtn.className = "btn-icon btn-icon-edit";
                editBtn.onclick = () => mostrarModalEditarUsuario(user, data);
                const delBtn = document.createElement("button"); delBtn.innerHTML = '<i class="fas fa-trash-alt"></i>'; delBtn.className = "btn-icon btn-icon-delete";
                if(user !== currentUser && !(currentUserRole === "root" && user === "root")) delBtn.onclick = () => { if(confirm(`¿Eliminar usuario ${user}?`)) { const newUsers = getUsers(); delete newUsers[user]; saveUsers(newUsers); renderUsersTable(); } };
                else delBtn.disabled = true;
                cellAcc.append(editBtn, delBtn);
            } else cellAcc.innerText = "---";
        }
    }

    function mostrarModalEditarUsuario(username, userData) {
        const modal = document.createElement("div"); modal.className = "modal";
        modal.innerHTML = `<div class="modal-content"><h3>Editar usuario: ${username}</h3><div class="form-grid"><div class="form-field"><label>Nombre completo</label><input id="editFullName" value="${(userData.fullName||'').replace(/"/g, '&quot;')}"></div><div class="form-field"><label>Teléfono</label><input id="editPhone" value="${(userData.phone||'').replace(/"/g, '&quot;')}"></div><div class="form-field"><label>Email</label><input id="editEmail" value="${(userData.email||'').replace(/"/g, '&quot;')}"></div><div class="form-field"><label>Rol</label><select id="editRole"><option value="user">user</option><option value="admin">admin</option></select></div></div><div style="margin-top:20px;"><button class="btn-outline" id="cancelEditUser">Cancelar</button><button class="btn-primary" id="saveEditUser">Guardar</button></div></div>`;
        document.body.appendChild(modal);
        document.getElementById("editRole").value = userData.role;
        document.getElementById("cancelEditUser").onclick = () => modal.remove();
        document.getElementById("saveEditUser").onclick = () => {
            const users = getUsers();
            if(users[username]) {
                users[username].fullName = document.getElementById("editFullName").value.trim();
                users[username].phone = document.getElementById("editPhone").value.trim();
                users[username].email = document.getElementById("editEmail").value.trim();
                users[username].role = document.getElementById("editRole").value;
                saveUsers(users);
                if(username === currentUser) currentUserFullName = users[username].fullName;
                document.getElementById("currentUserLabel").innerText = currentUserFullName;
                alert("Usuario actualizado");
                modal.remove();
                renderUsersTable();
            } else modal.remove();
        };
    }

    async function addUser() {
        if(currentUserRole !== "admin" && currentUserRole !== "root") return alert("No tiene permisos");
        const username = document.getElementById("newUser").value.trim();
        const pass = document.getElementById("newPass").value;
        const full = document.getElementById("newFullName").value.trim();
        const phone = document.getElementById("newPhone").value.trim();
        const email = document.getElementById("newEmail").value.trim();
        if(!username || !pass || !full) return alert("Complete usuario, contraseña y nombre");
        const users = getUsers();
        if(users[username]) return alert("Usuario ya existe");
        users[username] = { usuario: username, passwordHash: await hashPassword(pass), role: document.getElementById("newRole").value, createdAt: new Date().toISOString(), fullName: full, phone: phone, email: email, enabled: true };
        saveUsers(users);
        alert("Usuario creado");
        document.getElementById("newUser").value = ""; document.getElementById("newPass").value = ""; document.getElementById("newFullName").value = ""; document.getElementById("newPhone").value = ""; document.getElementById("newEmail").value = "";
        if(document.getElementById("usersList").style.display === "block") renderUsersTable();
    }

    // ==================== CONFIGURACIÓN ====================
    function openConfigModal() {
        // AHORA CUALQUIER USUARIO PUEDE ACCEDER A LA CONFIGURACIÓN (YA NO SE RESTRINGE SOLO A ADMIN/ROOT)
        const modal = document.createElement("div"); modal.className = "modal";
        modal.innerHTML = `<div class="modal-content"><h3>Configuración del Sistema</h3>
            <div class="config-list"><h4>📌 Áreas / Servicios</h4><div id="areasList"></div><div class="add-item"><input type="text" id="newArea" placeholder="Nueva área"><button class="btn-secondary" id="addAreaBtn">+ Agregar</button></div></div>
            <div class="config-list"><h4>📦 Tipo Producto</h4><div id="tiposProductoList"></div><div class="add-item"><input type="text" id="newTipoProd" placeholder="Nuevo tipo"><button class="btn-secondary" id="addTipoProdBtn">+ Agregar</button></div></div>
            <div class="config-list"><h4>🔄 Tipo Movimiento</h4><div id="tiposMovList"></div><div class="add-item"><input type="text" id="newTipoMov" placeholder="Nuevo tipo"><button class="btn-secondary" id="addTipoMovBtn">+ Agregar</button></div></div>
            <div class="config-list"><h4>🏷️ Estantes (Ubicaciones)</h4><div id="estantesList"></div><div class="add-item"><input type="text" id="newEstante" placeholder="Nuevo estante"><button class="btn-secondary" id="addEstanteBtn">+ Agregar</button></div></div>
            <button class="btn-outline" id="closeConfigModal">Cerrar</button></div>`;
        document.body.appendChild(modal);
        function renderList(containerId, items, storageKey) {
            const container = document.getElementById(containerId);
            container.innerHTML = items.map((item,idx)=>`<div class="config-item"><span>${item}</span><button class="btn-danger btn-icon" data-idx="${idx}">🗑️</button></div>`).join('');
            container.querySelectorAll("[data-idx]").forEach(btn=>{ btn.onclick = () => { const idx = parseInt(btn.dataset.idx); items.splice(idx,1); saveConfig(storageKey, items); renderList(containerId, items, storageKey); updateAllDynamicSelects(); }; });
        }
        function refreshAll() {
            renderList("areasList", getAreas(), STORAGE_CONFIG_AREAS);
            renderList("tiposProductoList", getTiposProducto(), STORAGE_CONFIG_TIPOS_PRODUCTO);
            renderList("tiposMovList", getTiposMovimiento(), STORAGE_CONFIG_TIPOS_MOV);
            renderList("estantesList", getEstantes(), STORAGE_CONFIG_ESTANTES);
        }
        document.getElementById("addAreaBtn").onclick = () => { const val = document.getElementById("newArea").value.trim(); if(val && !getAreas().includes(val)) { const areas = getAreas(); areas.push(val); saveConfig(STORAGE_CONFIG_AREAS, areas); refreshAll(); updateAllDynamicSelects(); document.getElementById("newArea").value=""; } };
        document.getElementById("addTipoProdBtn").onclick = () => { const val = document.getElementById("newTipoProd").value.trim(); if(val && !getTiposProducto().includes(val)) { const tipos = getTiposProducto(); tipos.push(val); saveConfig(STORAGE_CONFIG_TIPOS_PRODUCTO, tipos); refreshAll(); updateAllDynamicSelects(); document.getElementById("newTipoProd").value=""; } };
        document.getElementById("addTipoMovBtn").onclick = () => { const val = document.getElementById("newTipoMov").value.trim(); if(val && !getTiposMovimiento().includes(val)) { const tipos = getTiposMovimiento(); tipos.push(val); saveConfig(STORAGE_CONFIG_TIPOS_MOV, tipos); refreshAll(); updateAllDynamicSelects(); document.getElementById("newTipoMov").value=""; } };
        document.getElementById("addEstanteBtn").onclick = () => { const val = document.getElementById("newEstante").value.trim(); if(val && !getEstantes().includes(val)) { const estantes = getEstantes(); estantes.push(val); saveConfig(STORAGE_CONFIG_ESTANTES, estantes); refreshAll(); updateAllDynamicSelects(); document.getElementById("newEstante").value=""; } };
        document.getElementById("closeConfigModal").onclick = () => modal.remove();
        refreshAll();
    }

    function updateAllDynamicSelects() {
        const movTipo = document.getElementById("movTipo"); if(movTipo) movTipo.innerHTML = getTiposMovimiento().map(t=>`<option value="${t}">${t}</option>`).join('');
        const tipoProd = document.getElementById("movTipoProducto"); if(tipoProd) tipoProd.innerHTML = getTiposProducto().map(t=>`<option value="${t}">${t}</option>`).join('');
        const areaSelect = document.getElementById("areaDestino"); if(areaSelect) { areaSelect.innerHTML = getAreas().map(a=>`<option value="${a}">${a}</option>`).join(''); areaSelect.value = getAreas()[0]; document.getElementById("otroArea").style.display = areaSelect.value === "Otro" ? "block" : "none"; }
        const filtroArea = document.getElementById("filtroAreaHistorial"); if(filtroArea) filtroArea.innerHTML = '<option value="">Todas</option>' + getAreas().map(a=>`<option value="${a}">${a}</option>`).join('');
        const prodEstante = document.getElementById("prodEstante"); if(prodEstante) prodEstante.innerHTML = '<option value="">-- Sin estante --</option>' + getEstantes().map(e=>`<option value="${e}">${e}</option>`).join('');
        const movEstante = document.getElementById("movEstante"); if(movEstante) movEstante.innerHTML = '<option value="">-- Sin estante --</option>' + getEstantes().map(e=>`<option value="${e}">${e}</option>`).join('');
    }

    // ==================== NOTIFICACIONES ====================
    function updateNotificationBell() {
        const inv = getInventory();
        const hoy = new Date(); hoy.setHours(0,0,0,0);
        const expiryLimit = new Date(); expiryLimit.setDate(hoy.getDate() + getExpiryDays());
        const vencidos = inv.filter(p => p.vencimiento && new Date(p.vencimiento) < hoy);
        const proximos = inv.filter(p => p.vencimiento && new Date(p.vencimiento) >= hoy && new Date(p.vencimiento) <= expiryLimit);
        const bajoStock = inv.filter(p => p.existencia <= p.stockMinimo);
        const totalAlertas = vencidos.length + proximos.length + bajoStock.length;
        const badge = document.getElementById("alertBadge");
        if(totalAlertas > 0) { badge.style.display = "inline-block"; badge.innerText = totalAlertas; } else { badge.style.display = "none"; }
        const alertPanel = document.getElementById("alertPanel");
        let html = "";
        if(vencidos.length) html += `<strong style="color:#b91c1c;">⛔ VENCIDOS (${vencidos.length}):</strong><ul>${vencidos.map(p=>`<li>${p.codigo} - ${p.descripcion} (vence: ${p.vencimiento})</li>`).join('')}</ul>`;
        if(proximos.length) html += `<strong style="color:#f59e0b;">📅 PRÓXIMOS A VENCER (${getExpiryDays()} días):</strong><ul>${proximos.map(p=>`<li>${p.codigo} - ${p.descripcion} (vence: ${p.vencimiento})</li>`).join('')}</ul>`;
        if(bajoStock.length) html += `<strong style="color:#dc2626;">⚠️ STOCK BAJO (${bajoStock.length}):</strong><ul>${bajoStock.map(p=>`<li>${p.codigo} - ${p.descripcion} (stock: ${p.existencia} / min: ${p.stockMinimo})</li>`).join('')}</ul>`;
        if(!totalAlertas) html = "<p style='text-align:center;'>✅ Sin alertas activas</p>";
        alertPanel.innerHTML = html;
    }

    // ==================== RECUPERACIÓN ====================
    function mostrarRecuperacion() {
        const modal = document.createElement("div"); modal.className = "modal";
        modal.innerHTML = `<div class="modal-content"><h3>Recuperar acceso</h3>
            <p>¿Qué desea restablecer?</p>
            <div class="radio-group">
                <label><input type="radio" name="recoveryType" value="password" checked> Contraseña</label>
                <label><input type="radio" name="recoveryType" value="user"> Usuario</label>
            </div>
            <div id="recoveryStepEmail">
                <p>Ingrese su correo electrónico registrado:</p>
                <input type="email" id="recoveryEmail" placeholder="ejemplo@hospital.com" style="width:100%; padding:12px; border-radius:28px; margin-bottom:16px;">
                <button id="buscarEmailBtn" class="btn-primary">Buscar cuenta</button>
            </div>
            <div id="resultadoRecuperacion" style="margin-top:16px;"></div>
            <button class="btn-outline" id="closeRecoveryModal" style="margin-top:16px;">Cerrar</button>
        </div>`;
        document.body.appendChild(modal);
        const buscarBtn = modal.querySelector('#buscarEmailBtn');
        const emailInput = modal.querySelector('#recoveryEmail');
        const resultDiv = modal.querySelector('#resultadoRecuperacion');

        async function procesarRecuperacion() {
            const type = modal.querySelector('input[name="recoveryType"]:checked').value;
            const email = emailInput.value.trim().toLowerCase();
            if(!email) { resultDiv.innerHTML = '<p style="color:red;">❌ Ingrese un correo electrónico.</p>'; return; }
            const users = getUsers();
            let encontrado = null;
            for (let [user, data] of Object.entries(users)) {
                if (data.email && data.email.toLowerCase() === email) { encontrado = { user, data }; break; }
            }
            if (!encontrado) {
                resultDiv.innerHTML = '<p style="color:red;">❌ No existe un usuario con ese correo electrónico.</p>';
                return;
            }
            if (type === 'user') {
                resultDiv.innerHTML = `<p style="color:green;">✅ Su usuario es: <strong>${encontrado.user}</strong></p><p>Se ha enviado un enlace de recuperación a su correo (simulado).</p>`;
                alert(`Simulación: Se ha enviado un correo a ${email} con su usuario: ${encontrado.user}`);
            } else {
                resultDiv.innerHTML = `<p style="color:green;">✅ Cuenta encontrada: Usuario <strong>${encontrado.user}</strong></p><p>Ingrese nueva contraseña:</p><input type="password" id="newPwdRec" placeholder="Nueva contraseña" style="width:100%; padding:8px; border-radius:20px;"><button id="resetPwdConfirmBtn" class="btn-secondary" style="margin-top:8px;">Restablecer contraseña</button>`;
                const resetBtn = resultDiv.querySelector('#resetPwdConfirmBtn');
                resetBtn.onclick = async () => {
                    const newPass = resultDiv.querySelector('#newPwdRec').value;
                    if (!newPass) return alert("Ingrese nueva contraseña");
                    users[encontrado.user].passwordHash = await hashPassword(newPass);
                    saveUsers(users);
                    alert("Contraseña restablecida exitosamente. Ahora puede iniciar sesión.");
                    modal.remove();
                };
            }
        }
        buscarBtn.onclick = procesarRecuperacion;
        modal.querySelector('#closeRecoveryModal').onclick = () => modal.remove();
    }

    // ==================== EXPORTAR / IMPORTAR ====================
    function exportToExcel(data, filename) {
        const ws = XLSX.utils.json_to_sheet(data);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, "Datos");
        XLSX.writeFile(wb, filename + ".xlsx");
    }
    async function importInventoryFromExcel(file) { 
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.onload = e => {
                const wb = XLSX.read(e.target.result, { type: 'array' });
                const sheet = wb.Sheets["Inventario"];
                if(sheet) {
                    const data = XLSX.utils.sheet_to_json(sheet);
                    const medicamentos = data.filter(p => p.tipoProducto === "Medicamento");
                    const insumos = data.filter(p => p.tipoProducto === "Insumo");
                    localStorage.setItem(STORAGE_PRODUCTOS, JSON.stringify(medicamentos));
                    localStorage.setItem(STORAGE_INSUMOS, JSON.stringify(insumos));
                    renderInventory(); renderMovementsTable(); renderDashboard();
                    resolve();
                } else reject("Hoja 'Inventario' no encontrada");
            };
            reader.onerror = reject;
            reader.readAsArrayBuffer(file);
        });
    }
    async function importMovementsFromExcel(file) {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.onload = e => {
                const wb = XLSX.read(e.target.result, { type: 'array' });
                const sheet = wb.Sheets["Movimientos"];
                if(sheet) {
                    const data = XLSX.utils.sheet_to_json(sheet);
                    localStorage.setItem(STORAGE_MOVEMENTS, JSON.stringify(data));
                    renderMovementsTable(); renderDashboard();
                    resolve();
                } else reject("Hoja 'Movimientos' no encontrada");
            };
            reader.onerror = reject;
            reader.readAsArrayBuffer(file);
        });
    }

    // ==================== INICIALIZACIÓN Y LOGIN ====================
    async function initData() {
        if(!localStorage.getItem(STORAGE_USERS)) {
            const rh=await hashPassword("admin123"); const ah=await hashPassword("admin"); const uh=await hashPassword("user");
            saveUsers({ "root":{passwordHash:rh,role:"root",fullName:"Dueño del Sistema",phone:"",email:"root@hospital.local",enabled:true,createdAt:new Date().toISOString()}, "admin":{passwordHash:ah,role:"admin",fullName:"Administrador",phone:"",email:"admin@hospital.local",enabled:true,createdAt:new Date().toISOString()}, "user":{passwordHash:uh,role:"user",fullName:"Usuario Normal",phone:"",email:"user@hospital.local",enabled:true,createdAt:new Date().toISOString()} });
        }
        if(!localStorage.getItem(STORAGE_PRODUCTOS) && !localStorage.getItem(STORAGE_INSUMOS)) {
            const sample=[{codigo:"123456",descripcion:"Solución Salina 0.9%",tipoProducto:"Medicamento",stockMinimo:10,existencia:45,lote:"L01",registroInvima:"INV001",vencimiento:"2026-04-24",estante:"A1"},{codigo:"PARA001",descripcion:"Paracetamol 500mg",tipoProducto:"Medicamento",stockMinimo:20,existencia:15,lote:"L02",registroInvima:"INV002",vencimiento:"2025-12-01",estante:"B3"},{codigo:"VENCIDO01",descripcion:"Ibuprofeno 400mg",tipoProducto:"Medicamento",stockMinimo:5,existencia:6,lote:"L03",registroInvima:"INV003",vencimiento:"2023-01-15",estante:"D4"},{codigo:"JERINGA10",descripcion:"Jeringa 10ml",tipoProducto:"Insumo",stockMinimo:50,existencia:120,lote:"LJ01",registroInvima:"INV004",vencimiento:"2027-01-01",estante:"C2"}];
            localStorage.setItem(STORAGE_PRODUCTOS,JSON.stringify(sample.filter(p=>p.tipoProducto==="Medicamento")));
            localStorage.setItem(STORAGE_INSUMOS,JSON.stringify(sample.filter(p=>p.tipoProducto==="Insumo")));
        }
        if(!localStorage.getItem(STORAGE_MOVEMENTS)) localStorage.setItem(STORAGE_MOVEMENTS,"[]");
        if(!localStorage.getItem(STORAGE_EXPIRY_DAYS)) localStorage.setItem(STORAGE_EXPIRY_DAYS,"30");
        if(!localStorage.getItem(STORAGE_CONFIG_AREAS)) saveConfig(STORAGE_CONFIG_AREAS, defaultAreas);
        if(!localStorage.getItem(STORAGE_CONFIG_TIPOS_PRODUCTO)) saveConfig(STORAGE_CONFIG_TIPOS_PRODUCTO, defaultTiposProducto);
        if(!localStorage.getItem(STORAGE_CONFIG_TIPOS_MOV)) saveConfig(STORAGE_CONFIG_TIPOS_MOV, defaultTiposMovimiento);
        if(!localStorage.getItem(STORAGE_CONFIG_ESTANTES)) saveConfig(STORAGE_CONFIG_ESTANTES, defaultEstantes);
    }

    function setupApp() {
        renderInventory(); renderMovementsTable(); setupEventosEntrega(); renderHistorialEntregas(); updateAllDynamicSelects(); renderDashboard();
        const bell = document.getElementById("bellIcon");
        const panel = document.getElementById("alertPanel");
        bell.addEventListener("click", (e) => { e.stopPropagation(); panel.style.display = panel.style.display === "block" ? "none" : "block"; });
        document.addEventListener("click", (e) => { if (!bell.contains(e.target)) panel.style.display = "none"; });
        document.getElementById("btnBuscarInv").onclick = () => renderInventory();
        document.getElementById("btnLimpiarInv").onclick = () => { document.getElementById("invSearchValue").value = ""; renderInventory(); };
        document.getElementById("btnProximosVencer").onclick = () => { const inv = getInventory(); const hoy = new Date(); hoy.setHours(0,0,0,0); const limit = new Date(); limit.setDate(hoy.getDate() + getExpiryDays()); const prox = inv.filter(p => p.vencimiento && new Date(p.vencimiento) >= hoy && new Date(p.vencimiento) <= limit); if(prox.length) renderInventory(prox); else alert("No hay productos próximos a vencer"); };
        document.getElementById("btnBajoStockInv").onclick = () => { const inv = getInventory(); const bajo = inv.filter(p => p.existencia <= p.stockMinimo); if(bajo.length) renderInventory(bajo); else alert("No hay productos con bajo stock"); };
        document.getElementById("setExpiryDaysBtn").onclick = () => setExpiryDays(parseInt(document.getElementById("expiryDaysInput").value));
        document.getElementById("exportInventarioBtn").onclick = () => exportToExcel(getInventory(), "inventario_completo");
        document.getElementById("importInventarioBtn").onclick = () => { const input = document.createElement("input"); input.type="file"; input.accept=".xlsx"; input.onchange = async (e) => { if(e.target.files.length) { await importInventoryFromExcel(e.target.files[0]); alert("Inventario importado"); } }; input.click(); };
        document.getElementById("addProductBtn").onclick = addOrUpdateProduct;
        document.getElementById("cancelEditBtn").onclick = cancelEditProduct;
        document.getElementById("registrarMovBtn").onclick = async () => { try { const cod = document.getElementById("movCodigo").value.trim(); if(!cod) throw new Error("Código obligatorio"); const cantidad = parseInt(document.getElementById("movCantidad").value); if(isNaN(cantidad) || cantidad <= 0) throw new Error("Cantidad válida > 0"); await registrarMovimiento({ codigo:cod, cantidad:cantidad, tipo:document.getElementById("movTipo").value, descripcion:document.getElementById("movDescripcion").value, tipoProducto:document.getElementById("movTipoProducto").value, lote:document.getElementById("movLote").value, registroInvima:document.getElementById("movRegistroInvima").value, vencimiento:document.getElementById("movVencimiento").value, estante:document.getElementById("movEstante").value, stockMinimo:parseInt(document.getElementById("movStockMinimo").value)||0 }); alert("Movimiento registrado exitosamente"); limpiarFormularioMovimiento(); renderInventory(); renderMovementsTable(); updateNotificationBell(); renderDashboard(); } catch(e){ alert("Error: "+e.message); } };
        document.getElementById("limpiarFormBtn").onclick = limpiarFormularioMovimiento;
        document.getElementById("buscarProductoBtn").onclick = openProductSearchModal;
        document.getElementById("exportMovimientosBtn").onclick = () => exportToExcel(getMovements(), "movimientos_completos");
        document.getElementById("importMovimientosBtn").onclick = () => { const input = document.createElement("input"); input.type="file"; input.accept=".xlsx"; input.onchange = async (e) => { if(e.target.files.length) { await importMovementsFromExcel(e.target.files[0]); alert("Movimientos importados"); } }; input.click(); };
        document.getElementById("reporteInventarioBtn").onclick = () => exportToExcel(getInventory(), "inventario_reporte");
        document.getElementById("reporteMovimientosBtn").onclick = () => exportToExcel(getMovements(), "movimientos_reporte");
        document.getElementById("reporteVencimientoBtn").onclick = () => { const inv = getInventory(); const hoy = new Date(); const limit = new Date(); limit.setDate(hoy.getDate() + getExpiryDays()); const prox = inv.filter(p => p.vencimiento && new Date(p.vencimiento) >= hoy && new Date(p.vencimiento) <= limit); exportToExcel(prox, "proximos_a_vencer"); };
        document.getElementById("reporteBajoStockBtn").onclick = () => { const inv = getInventory(); const bajo = inv.filter(p => p.existencia <= p.stockMinimo); exportToExcel(bajo, "bajo_stock"); };
        document.getElementById("reporteEntregasBtn").onclick = () => { const entregas = getEntregas(); const flat = entregas.flatMap(p => p.productos.map(prod => ({ fecha: p.fecha, usuario: p.usuario, area: p.area, destinatario: p.destinatario, codigo: prod.codigo, cantidad: prod.cantidad, descripcion: prod.descripcion }))); exportToExcel(flat, "entregas_por_area"); };
        document.getElementById("cambiarPwdBtn").onclick = async () => { const old = document.getElementById("pwdActual").value; const newp = document.getElementById("pwdNueva").value; const conf = document.getElementById("pwdConfirm").value; if(newp !== conf) return alert("Las nuevas contraseñas no coinciden"); const users = getUsers(); if(await verifyPassword(old, users[currentUser].passwordHash)) { users[currentUser].passwordHash = await hashPassword(newp); saveUsers(users); alert("Contraseña cambiada"); document.getElementById("pwdActual").value = ""; document.getElementById("pwdNueva").value = ""; document.getElementById("pwdConfirm").value = ""; } else alert("Contraseña actual incorrecta"); };
        document.getElementById("logoutBtn").onclick = () => { sessionStorage.clear(); location.reload(); };
        document.getElementById("configBtn").onclick = openConfigModal;
        document.getElementById("toggleUsersBtn").onclick = () => { const div = document.getElementById("usersList"); if(div.style.display === "none") { renderUsersTable(); div.style.display = "block"; if(currentUserRole==="admin"||currentUserRole==="root") document.getElementById("adminAddUser").style.display = "block"; } else div.style.display = "none"; };
        document.getElementById("addUserBtn").onclick = addUser;
        document.getElementById("aplicarFiltrosHistorial").onclick = renderHistorialEntregas;
        document.getElementById("limpiarFiltrosHistorial").onclick = () => { document.getElementById("filtroAreaHistorial").value = ""; document.getElementById("filtroDestinatarioHistorial").value = ""; document.getElementById("filtroFechaDesdeHistorial").value = ""; document.getElementById("filtroFechaHastaHistorial").value = ""; renderHistorialEntregas(); };
        document.getElementById("exportAllDataBtn").onclick = () => { const allData = { productos: getInventory(), movimientos: getMovements(), usuarios: getUsers(), configuraciones: { areas: getAreas(), tiposProducto: getTiposProducto(), tiposMovimiento: getTiposMovimiento(), estantes: getEstantes(), expiryDays: getExpiryDays() } }; exportToExcel([allData], "backup_completo"); };
        document.getElementById("resetDataBtn").onclick = () => { if(confirm("Restablecer datos de ejemplo?")) { localStorage.clear(); location.reload(); } };
        document.getElementById("renewLicenseBtn").onclick = () => { if(confirm("¿Renovar licencia por 3 meses? (requiere contraseña)")) { const pwd = prompt("Contraseña de habilitación:"); if(pwd === HABILITATION_PASSWORD) { localStorage.setItem(STORAGE_LICENSE_CHECK, new Date().toISOString()); alert("Licencia renovada por 3 meses."); renderDashboard(); } else alert("Contraseña incorrecta."); } };
        if(currentUserRole === "root") document.getElementById("licenciaTabBtn").style.display = "block";
        document.querySelectorAll(".tab-btn").forEach(btn => {
            btn.onclick = () => {
                document.querySelectorAll(".tab-btn").forEach(b=>b.classList.remove("active"));
                btn.classList.add("active");
                document.querySelectorAll(".tab-content").forEach(t=>t.classList.remove("active"));
                document.getElementById(`tab${btn.dataset.tab.charAt(0).toUpperCase()+btn.dataset.tab.slice(1)}`).classList.add("active");
                if(btn.dataset.tab === "entregas") cargarSelectProductosEntrega();
                if(btn.dataset.tab === "historialEntregas") renderHistorialEntregas();
                if(btn.dataset.tab === "usuarios" && document.getElementById("usersList").style.display === "block") renderUsersTable();
                if(btn.dataset.tab === "dashboard") renderDashboard();
            };
        });
    }

    window.onload = async () => {
        await initData();
        const loginView = document.getElementById("loginView"), appView = document.getElementById("appView");
        document.getElementById("forgotPasswordLink").onclick = () => mostrarRecuperacion();
        document.getElementById("loginBtn").onclick = async () => {
            const user = document.getElementById("loginUser").value, pass = document.getElementById("loginPass").value;
            const users = getUsers();
            if(users[user] && await verifyPassword(pass, users[user].passwordHash)) {
                if(users[user].enabled === false) { alert("Usuario deshabilitado. Contacte al administrador."); return; }
                currentUser = user; currentUserRole = users[user].role; currentUserFullName = users[user].fullName;
                sessionStorage.setItem("currentUser",user); sessionStorage.setItem("currentUserRole",currentUserRole);
                document.getElementById("currentUserLabel").innerText = currentUserFullName;
                document.getElementById("userRoleLabel").innerText = "Rol: "+currentUserRole;
                document.getElementById("userActual").value = user;
                loginView.style.display = "none"; appView.style.display = "flex";
                setupApp();
            } else alert("Credenciales incorrectas");
        };
        if(sessionStorage.getItem("currentUser")) {
            currentUser = sessionStorage.getItem("currentUser");
            currentUserRole = sessionStorage.getItem("currentUserRole");
            const users = getUsers();
            if(users[currentUser]) {
                currentUserFullName = users[currentUser].fullName;
                document.getElementById("currentUserLabel").innerText = currentUserFullName;
            } else document.getElementById("currentUserLabel").innerText = currentUser;
            document.getElementById("userRoleLabel").innerText = "Rol: "+currentUserRole;
            document.getElementById("userActual").value = currentUser;
            loginView.style.display = "none"; appView.style.display = "flex";
            setupApp();
        }
    };
</script>
</body>
</html>
