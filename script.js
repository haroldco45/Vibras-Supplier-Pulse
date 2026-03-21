let baseDatos = { prov: {}, fac: {}, vend: {}, cli: {} };

document.getElementById('btnProcesar').addEventListener('click', async function() {
    const files = {
        ventas: document.getElementById('inputVentas').files[0],
        inv: document.getElementById('inputInventario').files[0],
        cxp: document.getElementById('inputCXP').files[0],
        compras: document.getElementById('inputCompras').files[0]
    };

    if (!files.ventas || !files.inv || !files.cxp || !files.compras) {
        alert("Socio, ¡falta artillería! Carga los 4 archivos."); return;
    }

    const [v, inv, cxp, comp] = await Promise.all([
        leerExcel(files.ventas, "A"), leerExcel(files.inv, "A"),
        leerExcel(files.cxp, "A"), leerExcel(files.compras, "A")
    ]);

    procesarTodo(v, inv, cxp, comp);
});

function procesarTodo(v, inv, cxp, comp) {
    // REINICIAR BASE DE DATOS
    baseDatos = { prov: {}, fac: {}, vend: {}, cli: {} };

    // 1. MAPEO DE COMPRAS (LLAVE PARA TODO)
    comp.forEach(f => {
        if(f.B) baseDatos.fac[f.B.toString().trim()] = { prov: f.E, total: f.L, saldo: 0 };
    });

    // 2. PROCESAR VENTAS (PROV, VEND, CLI)
    v.forEach(f => {
        let p = f.H || "SIN PROVEEDOR";
        let vend = f.L || "SIN VENDEDOR";
        let cli = f.AJ || "CLIENTE VARIOS";
        let total = (parseFloat(f.E) || 0) * (parseFloat(f.F) || 0);
        if(f.A == 2) total = -total;

        // Agrupar Proveedores
        if(!baseDatos.prov[p]) baseDatos.prov[p] = { vta: 0, deuda: 0 };
        baseDatos.prov[p].vta += total;

        // Agrupar Vendedores
        if(!baseDatos.vend[vend]) baseDatos.vend[vend] = { vta: 0 };
        baseDatos.vend[vend].vta += total;

        // Agrupar Clientes
        if(!baseDatos.cli[cli]) baseDatos.cli[cli] = { vta: 0 };
        baseDatos.cli[cli].vta += total;
    });

    // 3. PROCESAR CXP
    cxp.forEach(f => {
        let fac = f.B ? f.B.toString().trim() : "";
        let saldo = parseFloat(f.E) || 0;
        if(baseDatos.fac[fac]) {
            baseDatos.fac[fac].saldo = saldo;
            let p = baseDatos.fac[fac].prov;
            if(baseDatos.prov[p]) baseDatos.prov[p].deuda += saldo;
        }
    });

    document.getElementById('welcomeMessage').style.display = 'none';
    document.getElementById('dashboardContent').style.display = 'block';
    cambiarTab('proveedores'); // Tab por defecto
}

function cambiarTab(tipo) {
    const areaSel = document.getElementById('areaSelector');
    const areaRes = document.getElementById('areaResultados');
    areaRes.innerHTML = "";
    
    let options = `<option value="">-- Selecciona un ${tipo} --</option>`;
    let dataObj = baseDatos[tipo === 'proveedores' ? 'prov' : tipo === 'vendedores' ? 'vend' : tipo === 'clientes' ? 'cli' : 'fac'];

    Object.keys(dataObj).sort().forEach(k => {
        options += `<option value="${k}">${k}</option>`;
    });

    areaSel.innerHTML = `<select class="form-select shadow-sm" onchange="verDetalle('${tipo}', this.value)">${options}</select>`;
}

function verDetalle(tipo, llave) {
    if(!llave) return;
    let data = baseDatos[tipo === 'proveedores' ? 'prov' : tipo === 'vendedores' ? 'vend' : tipo === 'clientes' ? 'cli' : 'fac'][llave];
    let html = `<div class="mt-4 p-4 border rounded bg-white shadow-sm border-primary">
                    <h4>Análisis de ${llave}</h4><hr>`;
    
    if(tipo === 'proveedores') {
        let semaforo = data.vta > data.deuda ? '🟢 SALUDABLE' : '🔴 CRÍTICO';
        html += `<p>Ventas del Mes: <strong>$${data.vta.toLocaleString()}</strong></p>
                 <p>Deuda Total: <strong>$${data.deuda.toLocaleString()}</strong></p>
                 <div class="alert ${data.vta > data.deuda ? 'alert-success' : 'alert-danger'}">${semaforo}</div>`;
    } else {
        html += `<p>Ventas Totales: <strong>$${data.vta ? data.vta.toLocaleString() : 'N/A'}</strong></p>`;
        if(tipo === 'facturas') html += `<p>Saldo Pendiente: <strong>$${data.saldo.toLocaleString()}</strong></p>`;
    }
    
    html += `</div>`;
    document.getElementById('areaResultados').innerHTML = html;
}

async function leerExcel(file, header) {
    return new Promise((resolve) => {
        const reader = new FileReader();
        reader.onload = (e) => {
            const wb = XLSX.read(new Uint8Array(e.target.result), {type: 'array'});
            resolve(XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], {header: header}));
        };
        reader.readAsArrayBuffer(file);
    });
}
