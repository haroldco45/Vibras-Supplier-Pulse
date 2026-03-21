let baseDatos = { prov: {}, fac: {}, vend: {}, cli: {}, productos: {} };

document.getElementById('btnProcesar').addEventListener('click', async function() {
    const ids = ['inputVentas', 'inputInventario', 'inputCXP', 'inputCompras'];
    const files = ids.map(id => document.getElementById(id).files[0]);

    if (files.some(f => !f)) {
        alert("Socio, te falta cargar algún archivo. ¡Deben ser los 4!");
        return;
    }

    document.getElementById('welcomeMessage').innerHTML = '<h3 class="text-primary">Abriendo bodega de datos... ⏳</h3>';

    try {
        // Leemos los 4 archivos de forma ultra-compatible
        const [v, inv, cxp, comp] = await Promise.all(files.map(f => leerPro(f)));
        procesarTodo(v, inv, cxp, comp);
    } catch (error) {
        console.error(error);
        alert("Error: El navegador no pudo abrir los archivos. Intenta cerrando el Excel en tu PC antes de subirlos.");
    }
});

function leerPro(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = (e) => {
            try {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, {type: 'array'});
                const sheet = workbook.Sheets[workbook.SheetNames[0]];
                resolve(XLSX.utils.sheet_to_json(sheet, {defval: ""}));
            } catch (err) { reject(err); }
        };
        reader.readAsArrayBuffer(file);
    });
}

function procesarTodo(v, inv, cxp, comp) {
    baseDatos = { prov: {}, fac: {}, vend: {}, cli: {}, productos: {} };

    // 1. COMPRAS
    comp.forEach(f => {
        let n = f.factura || f.documento || f.numero;
        if(n) baseDatos.fac[n.toString().trim()] = f.razon_social || f.proveedor;
    });

    // 2. INVENTARIO
    inv.forEach(f => {
        let r = f.referencia || f.codigo;
        if(r) baseDatos.productos[r.toString().trim()] = { 
            s: parseFloat(f.totalinventario) || 0, 
            c: parseFloat(f.costo_ponderado_final) || 0 
        };
    });

    // 3. VENTAS
    v.forEach(f => {
        let p = f.referencia_proveedor || "SIN PROVEEDOR";
        let r = f.referencia ? f.referencia.toString().trim() : "";
        let d = (f.nombre_producto || "").toUpperCase();
        
        if (p === "SIN PROVEEDOR" || p === "") {
            if (d.includes("COLOMBINA")) p = "COLOMBINA";
            else if (d.includes("FAMILIA") || d.includes("SANCELA")) p = "FAMILIA SANCELA";
        }

        let total = (parseFloat(f.cantidad) || 0) * (parseFloat(f.precio_base_venta) || 0);
        if(f.tipo == "2") total = -total;

        if(!baseDatos.prov[p]) baseDatos.prov[p] = { vta: 0, deu: 0, invV: 0 };
        baseDatos.prov[p].vta += total;

        if(baseDatos.productos[r]) {
            baseDatos.prov[p].invV += (baseDatos.productos[r].s * baseDatos.productos[ref].c);
        }
    });

    // 4. CXP
    cxp.forEach(f => {
        let n = f.documento || f.numero;
        let p = baseDatos.fac[n] || f.razon_social || "OTROS";
        let s = parseFloat(f.totalvalor) || parseFloat(f.saldo) || 0;
        if(!baseDatos.prov[p]) baseDatos.prov[p] = { vta: 0, deu: 0, invV: 0 };
        baseDatos.prov[p].deu += s;
    });

    document.getElementById('welcomeMessage').style.display = 'none';
    document.getElementById('dashboardContent').style.display = 'block';
    cambiarTab('proveedores');
}

function cambiarTab(t) {
    const area = document.getElementById('areaSelector');
    let obj = baseDatos[t === 'proveedores' ? 'prov' : t === 'vendedores' ? 'vend' : 'cli'];
    let opts = `<option value="">-- SELECCIONA ${t.toUpperCase()} --</option>`;
    Object.keys(obj).sort().forEach(k => opts += `<option value="${k}">${k}</option>`);
    area.innerHTML = `<select class="form-select p-3 border-primary" onchange="verDetalle('${t}', this.value)">${opts}</select>`;
}

function verDetalle(t, ll) {
    let d = baseDatos[t === 'proveedores' ? 'prov' : t === 'vendedores' ? 'vend' : 'cli'][ll];
    let rentable = d.vta > d.deu;
    document.getElementById('areaResultados').innerHTML = `
        <div class="mt-4 p-4 border rounded bg-white shadow text-center">
            <h3 class="text-primary">${ll}</h3><hr>
            <div class="row">
                <div class="col-4">Ventas<br><h5 class="text-success">$${Math.round(d.vta).toLocaleString()}</h5></div>
                <div class="col-4">Deuda<br><h5 class="text-danger">$${Math.round(d.deu).toLocaleString()}</h5></div>
                <div class="col-4">Inv.<br><h5 class="text-warning">$${Math.round(d.invV).toLocaleString()}</h5></div>
            </div>
            <div class="alert ${rentable ? 'alert-success' : 'alert-danger'} mt-3">
                <h3>${rentable ? '🟢 RENTABLE' : '🔴 NO RENTABLE'}</h3>
            </div>
        </div>`;
}
