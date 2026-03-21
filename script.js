let baseDatos = { prov: {}, fac: {}, vend: {}, cli: {}, productos: {} };

document.addEventListener('click', async function(e) {
    if (e.target && e.target.id === 'btnProcesar') {
        const fVentas = document.getElementById('inputVentas').files[0];
        const fInv = document.getElementById('inputInventario').files[0];
        const fCXP = document.getElementById('inputCXP').files[0];
        const fCompras = document.getElementById('inputCompras').files[0];

        if (!fVentas || !fInv || !fCXP || !fCompras) {
            alert("Socio, ¡faltan archivos! Carga los 4 para el análisis maestro.");
            return;
        }

        document.getElementById('welcomeMessage').innerHTML = '<div class="text-center py-5"><div class="spinner-border text-primary"></div><h4 class="mt-3">Vibras Positivas analizando rentabilidad...</h4></div>';

        try {
            const [v, inv, cxp, comp] = await Promise.all([
                leerArchivo(fVentas), leerArchivo(fInv),
                leerArchivo(fCXP), leerArchivo(fCompras)
            ]);

            procesarTodo(v, inv, cxp, comp);
        } catch (err) {
            console.error(err);
            alert("Error al leer Excel. Asegúrate que no tengan contraseña.");
        }
    }
});

function leerArchivo(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = (e) => {
            try {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type: 'array' });
                const json = XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]]);
                resolve(json);
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

        let cant = parseFloat(f.cantidad) || 0;
        let pre = parseFloat(f.precio_base_venta) || 0;
        let sub = cant * pre;
        if(f.tipo == "2") sub = -sub;

        if(!baseDatos.prov[p]) baseDatos.prov[p] = { vta: 0, deu: 0, invV: 0 };
        baseDatos.prov[p].vta += sub;

        if(baseDatos.productos[r]) {
            baseDatos.prov[p].invV += (baseDatos.productos[r].s * baseDatos.productos[r].c);
        }

        let ven = f.vendedor || "OFICINA";
        if(!baseDatos.vend[ven]) baseDatos.vend[ven] = { vta: 0 };
        baseDatos.vend[ven].vta += sub;
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
    area.innerHTML = `<select class="form-select p-3 border-primary shadow-sm" onchange="verDetalle('${t}', this.value)">${opts}</select>`;
}

function verDetalle(t, ll) {
    let d = baseDatos[t === 'proveedores' ? 'prov' : t === 'vendedores' ? 'vend' : 'cli'][ll];
    if(t === 'proveedores') {
        let rentable = d.vta > d.deu;
        let vtaDia = d.vta / 30;
        let diasInv = d.vta > 0 ? Math.round(d.invV / (vtaDia || 1)) : 0;

        document.getElementById('areaResultados').innerHTML = `
            <div class="mt-4 p-4 border rounded bg-white shadow-sm border-primary">
                <h4 class="text-primary">${ll}</h4><hr>
                <div class="row text-center">
                    <div class="col-4">Ventas<br><h5 class="text-success">$${Math.round(d.vta).toLocaleString()}</h5></div>
                    <div class="col-4">Deuda<br><h5 class="text-danger">$${Math.round(d.deu).toLocaleString()}</h5></div>
                    <div class="col-4">Días Stock<br><h5 class="text-warning">${diasInv}</h5></div>
                </div>
                <div class="alert ${rentable ? 'alert-success' : 'alert-danger'} mt-3 text-center">
                    <h3>${rentable ? '🟢 RENTABLE' : '🔴 RIESGO'}</h3>
                </div>
                <div class="bg-light p-3 rounded mt-2">
                    <strong>Estrategia CEO:</strong> ${d.deu > d.vta ? 'Frenar compras. Negociar devolución de stock muerto.' : 'Pide descuento por volumen.'}
                </div>
            </div>`;
    } else {
        document.getElementById('areaResultados').innerHTML = `<div class="mt-4 p-4 border rounded bg-white text-center shadow-sm">
            <h4>${ll}</h4><h2 class="text-primary">$${Math.round(d.vta).toLocaleString()}</h2><p>Total Movimiento</p></div>`;
    }
}
