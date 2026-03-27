// Creado por Vibras Positivas para Distrileco - 2026
document.getElementById('btnProcesar').addEventListener('click', async () => {
    const files = {
        ventas: document.getElementById('inputVentas').files[0],
        inventario: document.getElementById('inputInventario').files[0],
        cxp: document.getElementById('inputCXP').files[0],
        compras: document.getElementById('inputCompras').files[0]
    };

    if (!files.ventas || !files.inventario || !files.cxp || !files.compras) {
        alert("Socio, por favor carga los 4 archivos para poder cruzar la información.");
        return;
    }

    const data = {};
    
    try {
        for (const key in files) {
            data[key] = await readExcel(files[key]);
        }
        procesarDashboard(data);
    } catch (error) {
        console.error("Error procesando archivos:", error);
        alert("Hubo un error al leer los archivos. Revisa que sean Excel válidos.");
    }
});

function readExcel(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = (e) => {
            const workbook = XLSX.read(e.target.result, { type: 'binary' });
            const sheetName = workbook.SheetNames[0];
            const json = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName]);
            resolve(json);
        };
        reader.onerror = reject;
        reader.readAsBinaryString(file);
    });
}

function procesarDashboard(data) {
    const select = document.getElementById('selectProveedor');
    const proveedores = [...new Set(data.compras.map(item => item.Proveedor || item.PROVEEDOR || item.nombre))].sort();

    select.innerHTML = '<option value="">-- Elige un proveedor para analizar --</option>';
    proveedores.forEach(prov => {
        const opt = document.createElement('option');
        opt.value = prov;
        opt.textContent = prov;
        select.appendChild(opt);
    });

    document.getElementById('welcomeMessage').style.display = 'none';
    document.getElementById('dashboardResult').style.display = 'block';

    select.onchange = () => mostrarKpis(select.value, data);
}

function mostrarKpis(proveedor, data) {
    if (!proveedor) return;

    // Filtrado Inteligente
    const comprasProv = data.compras.filter(i => (i.Proveedor || i.PROVEEDOR) === proveedor);
    const productosDeEsteProv = [...new Set(comprasProv.map(i => i.Producto || i.ARTICULO || i.COD_REF))];
    
    const ventasProv = data.ventas.filter(i => productosDeEsteProv.includes(i.Producto || i.ARTICULO));
    const invProv = data.inventario.filter(i => productosDeEsteProv.includes(i.Producto || i.ARTICULO));
    const cxpProv = data.cxp.filter(i => (i.Proveedor || i.PROVEEDOR) === proveedor);

    // Cálculos
    const totalVentas = ventasProv.reduce((acc, curr) => acc + (Number(curr.Total || curr.VENTA) || 0), 0);
    const totalInv = invProv.reduce((acc, curr) => acc + (Number(curr.CostoTotal || curr.VALOR) || 0), 0);
    const totalCxp = cxpProv.reduce((acc, curr) => acc + (Number(curr.Saldo || curr.VALOR_DEUDA) || 0), 0);

    const kpiContainer = document.getElementById('kpiContainer');
    kpiContainer.innerHTML = `
        <div class="row text-center">
            <div class="col-md-4">
                <div class="p-3 border rounded bg-white">
                    <h6>Ventas Mes</h6>
                    <h4 class="text-success">$${totalVentas.toLocaleString()}</h4>
                </div>
            </div>
            <div class="col-md-4">
                <div class="p-3 border rounded bg-white">
                    <h6>Inventario Actual</h6>
                    <h4 class="text-primary">$${totalInv.toLocaleString()}</h4>
                </div>
            </div>
            <div class="col-md-4">
                <div class="p-3 border rounded bg-white">
                    <h6>Saldo Pendiente (CXP)</h6>
                    <h4 class="text-danger">$${totalCxp.toLocaleString()}</h4>
                </div>
            </div>
        </div>
    `;

    // Recomendación Estratégica
    const recBox = document.getElementById('recomendacionesBox');
    let recomendacion = "✅ Relación estable.";
    if (totalCxp > totalVentas * 1.5) {
        recomendacion = "⚠️ **Alerta:** La deuda supera con creces la venta mensual. Sugerencia: Negociar devolución de excedentes o ampliar plazo.";
    } else if (totalInv < (totalVentas / 4)) {
        recomendacion = "🚀 **Oportunidad:** Inventario bajo para el ritmo de venta. ¡Pide más producto!";
    }

    recBox.innerHTML = `
        <div class="alert alert-info">
            <strong>Sugerencia de Negociación:</strong><br>${recomendacion}
        </div>
    `;
}
