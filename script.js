let baseDatos = { prov: {}, fac: {}, vend: {}, cli: {}, productos: {} };

document.getElementById('btnProcesar').addEventListener('click', async function() {
    // Referencias a los inputs
    const fVentas = document.getElementById('inputVentas').files[0];
    const fInv = document.getElementById('inputInventario').files[0];
    const fCXP = document.getElementById('inputCXP').files[0];
    const fCompras = document.getElementById('inputCompras').files[0];

    if (!fVentas || !fInv || !fCXP || !fCompras) {
        alert("Socio, te falta cargar algún archivo. ¡Revisa los 4!");
        return;
    }

    document.getElementById('welcomeMessage').innerHTML = '<h3 class="text-primary">Moviendo datos en Vibras Positivas... ⏳</h3>';

    try {
        const [v, inv, cxp, comp] = await Promise.all([
            leerMaster(fVentas),
            leerMaster(fInv),
            leerMaster(fCXP),
            leerMaster(fCompras)
        ]);

        procesarTodo(v, inv, cxp, comp);
    } catch (error) {
        console.error("Error crítico:", error);
        alert("Error al leer: Asegúrate de que los archivos NO estén abiertos en Excel.");
    }
});

// LA FUNCIÓN QUE NO FALLA
function leerMaster(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = (e) => {
            try {
                const bstr = e.target.result;
                const wb = XLSX.read(bstr, { type: 'binary', cellDates: true });
                const wsname = wb.SheetNames[0];
                const data = XLSX.utils.sheet_to_json(wb.Sheets[wsname]);
                resolve(data);
            } catch (err) { reject(err); }
        };
        reader.readAsBinaryString(file);
    });
}

function procesarTodo(v, inv, cxp, comp) {
    baseDatos = { prov: {}, fac: {}, vend: {}, cli: {}, productos: {} };

    // 1. MAPEAR COMPRAS
    comp.forEach(f => {
        let n = f.factura || f.documento || f.numero;
        if(n) baseDatos.fac[n.toString().trim()] = f.razon_social || f.proveedor;
    });

    // 2. MAPEAR INVENTARIO
    inv.forEach(f => {
        let r = f.referencia || f.codigo;
        if(r) baseDatos.productos[r.toString().trim()] = { 
            s: parseFloat(f.totalinventario) || 0, 
            c: parseFloat(f.costo_ponderado_final) || 0 
        };
    });

    // 3. MAPEAR VENTAS
    v.forEach(f => {
        let p = f.referencia_proveedor || "SIN PROVEEDOR";
        let r = f.referencia ? f.referencia.toString().trim() : "";
        let d = (f.nombre_producto || "").toUpperCase();
        
        if (p === "SIN PROVEEDOR" || p === "") {
            if (d.includes("COLOMBINA")) p = "COLOMBINA";
            else if (d.includes("FAMILIA") || d.includes("SANCELA")) p = "FAMILIA SANCELA";
        }

        let cant = parseFloat(f.cantidad) || 0;
        let pre = parseFloat(f.precio_base_venta) || 0;
        let sub = cant * pre;
        if(f.tipo == "2") sub = -sub;

        if(!baseDatos.prov[p]) baseDatos.prov[p] = { vta: 0, deu: 0, invV: 0 };
        baseDatos.prov[p].vta += sub;

        if(baseDatos.productos[r]) {
            baseDatos.prov[p].invV += (baseDatos.productos[r].s * baseDatos.productos[r].c);
        }
    });

    // 4. MAPEAR CXP
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
    let esBueno = d.vta > d.deu;
    document.getElementById('areaResultados').innerHTML = `
        <div class="mt-4 p-4 border rounded bg-white shadow-lg border-primary text-center">
            <h3>${ll}</h3><hr>
            <div class="row">
                <div class="col-4">Ventas: <br><strong class="text-success">$${d.vta.toLocaleString()}</strong></div>
                <div class="col-4">Deuda: <br><strong class="text-danger">$${d.deu.toLocaleString()}</strong></div>
                <div class="col-4">Inv: <br><strong class="text-warning">$${d.invV.toLocaleString()}</strong></div>
            </div>
            <div class="alert ${esBueno ? 'alert-success' : 'alert-danger'} mt-3">
                <h2>${esBueno ? '🟢 RENTABLE' : '🔴 RIESGO'}</h2>
            </div>
        </div>`;
}
