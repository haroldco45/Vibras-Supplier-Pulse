let baseDatos = { prov: {}, fac: {}, vend: {}, cli: {}, productos: {} };

document.getElementById('btnProcesar').addEventListener('click', async function() {
    const ids = ['inputVentas', 'inputInventario', 'inputCXP', 'inputCompras'];
    const files = ids.map(id => document.getElementById(id).files[0]);

    if (files.some(f => !f)) {
        alert("Harold, como CEO debes asegurar que los 4 reportes estén cargados.");
        return;
    }

    document.getElementById('welcomeMessage').innerHTML = `
        <div class="text-center py-5">
            <div class="spinner-border text-primary mb-3"></div>
            <h4>Vibras Positivas está cruzando millones de datos...</h4>
        </div>`;

    try {
        const [v, inv, cxp, comp] = await Promise.all(files.map(f => leerArchivoExcel(f)));
        procesarInteligencia(v, inv, cxp, comp);
    } catch (err) {
        console.error(err);
        alert("Error de lectura. Asegúrate de que los archivos sean Excel (.xlsx) válidos.");
    }
});

function leerArchivoExcel(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = (e) => {
            try {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type: 'array' });
                const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
                resolve(XLSX.utils.sheet_to_json(firstSheet));
            } catch (err) { reject(err); }
        };
        reader.readAsArrayBuffer(file);
    });
}

function procesarInteligencia(v, inv, cxp, comp) {
    baseDatos = { prov: {}, fac: {}, vend: {}, cli: {}, productos: {} };

    // 1. Mapeo de Compras (Factura -> Proveedor)
    comp.forEach(f => {
        let fac = f.factura || f.documento || f.numero;
        if(fac) baseDatos.fac[fac.toString().trim()] = f.razon_social || f.proveedor;
    });

    // 2. Mapeo de Inventario (Ref -> Stock/Costo)
    inv.forEach(f => {
        let ref = f.referencia || f.codigo;
        if(ref) baseDatos.productos[ref.toString().trim()] = { 
            s: parseFloat(f.totalinventario) || 0, 
            c: parseFloat(f.costo_ponderado_final) || 0 
        };
    });

    // 3. Procesar Ventas (El corazón del ingreso)
    v.forEach(f => {
        let p = f.referencia_proveedor || "SIN PROVEEDOR";
        let ref = f.referencia ? f.referencia.toString().trim() : "";
        let desc = (f.nombre_producto || "").toUpperCase();
        
        // Genialidad de Combos en Blanco
        if (p === "SIN PROVEEDOR" || p === "") {
            if (desc.includes("COLOMBINA")) p = "COLOMBINA";
            else if (desc.includes("FAMILIA") || desc.includes("SANCELA")) p = "FAMILIA SANCELA";
        }

        let total = (parseFloat(f.cantidad) || 0) * (parseFloat(f.precio_base_venta) || 0);
        if(f.tipo == "2") total = -total; // Devolución

        if(!baseDatos.prov[p]) baseDatos.prov[p] = { vta: 0, deu: 0, invVal: 0 };
        baseDatos.prov[p].vta += total;

        if(baseDatos.productos[ref]) {
            baseDatos.prov[p].invVal += (baseDatos.productos[ref].s * baseDatos.productos[ref].c);
        }

        let ven = f.vendedor || "OFICINA";
        if(!baseDatos.vend[ven]) baseDatos.vend[ven] = { vta: 0 };
        baseDatos.vend[ven].vta += total;
    });

    // 4. Mapear CXP (La cruda realidad)
    cxp.forEach(f => {
        let fac = f.documento || f.numero;
        let p = baseDatos.fac[fac] || f.razon_social || "OTROS";
        let saldo = parseFloat(f.totalvalor) || parseFloat(f.saldo) || 0;
        if(!baseDatos.prov[p]) baseDatos.prov[p] = { vta: 0, deu: 0, invVal: 0 };
        baseDatos.prov[p].deu += saldo;
    });

    document.getElementById('welcomeMessage').style.display = 'none';
    document.getElementById('dashboardContent').style.display = 'block';
    cambiarTab('proveedores');
}

function cambiarTab(t) {
    const selArea = document.getElementById('areaSelector');
    let obj = baseDatos[t === 'proveedores' ? 'prov' : t === 'vendedores' ? 'vend' : 'cli'];
    let opts = `<option value="">-- SELECCIONA ${t.toUpperCase()} --</option>`;
    Object.keys(obj).sort().forEach(k => { if(k !== "undefined") opts += `<option value="${k}">${k}</option>`; });
    selArea.innerHTML = `<select class="form-select form-select-lg border-primary shadow-sm" onchange="verDetalle('${t}', this.value)">${opts}</select>`;
}

function verDetalle(t, ll) {
    let d = baseDatos[t === 'proveedores' ? 'prov' : t === 'vendedores' ? 'vend' : 'cli'][ll];
    let esBueno = d.vta > d.deu;
    let html = `<div class="mt-4 p-4 border rounded bg-white shadow-sm border-primary">
                <div class="d-flex justify-content-between align-items-center">
                    <h3 class="text-primary mb-0">${ll}</h3>
                    <span class="badge ${esBueno ? 'bg-success' : 'bg-danger'} p-2 fs-6">${esBueno ? 'RENTABLE' : 'ALERTA'}</span>
                </div><hr>`;

    if(t === 'proveedores') {
        let vtaDia = d.vta / 30;
        let diasInv = d.vta > 0 ? Math.round(d.invVal / (vtaDia || 1)) : "∞";
        html += `
            <div class="row text-center mb-4">
                <div class="col-4"><h6>Ventas</h6><h4 class="text-success">$${Math.round(d.vta).toLocaleString()}</h4></div>
                <div class="col-4"><h6>Deuda</h6><h4 class="text-danger">$${Math.round(d.deu).toLocaleString()}</h4></div>
                <div class="col-4"><h6>Días Stock</h6><h4 class="text-warning">${diasInv}</h4></div>
            </div>
            <div class="p-3 rounded bg-light border-start border-4 ${esBueno ? 'border-success' : 'border-danger'}">
                <strong>💡 Estrategia CEO:</strong> ${d.deu > d.vta ? 'Socio, frena compras inmediatamente. Negocia devolución de stock muerto o exige bonificación adicional.' : 'Proveedor estrella. Pide descuento del 3-5% por pronto pago.'}
            </div>`;
    } else {
        html += `<h2 class="text-center text-primary py-3">$${Math.round(d.vta).toLocaleString()}</h2><p class="text-center text-muted">Total facturado en el periodo</p>`;
    }
    document.getElementById('areaResultados').innerHTML = html;
}
