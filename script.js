let baseDatos = { prov: {}, fac: {}, vend: {}, cli: {}, productos: {} };

document.getElementById('btnProcesar').onclick = async function() {
    const fVentas = document.getElementById('inputVentas').files[0];
    const fInv = document.getElementById('inputInventario').files[0];
    const fCXP = document.getElementById('inputCXP').files[0];
    const fCompras = document.getElementById('inputCompras').files[0];

    if (!fVentas || !fInv || !fCXP || !fCompras) {
        alert("Socio, cargue los 4 archivos para activar la inteligencia.");
        return;
    }

    document.getElementById('welcomeMessage').innerHTML = '<div class="text-center py-5"><div class="spinner-border text-primary mb-3"></div><h4>Analizando el ADN de Distrileco...</h4></div>';

    try {
        const [v, inv, cxp, comp] = await Promise.all([
            leerArchivo(fVentas), leerArchivo(fInv),
            leerArchivo(fCXP), leerArchivo(fCompras)
        ]);
        procesarTodo(v, inv, cxp, comp);
    } catch (error) {
        console.error(error);
        alert("Error: Asegúrese de que los archivos sean Excel (.xlsx) y no estén abiertos.");
    }
};

function leerArchivo(file) {
    return new Promise((resolve) => {
        const reader = new FileReader();
        reader.onload = (e) => {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, {type: 'array'});
            resolve(XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]], {defval: ""}));
        };
        reader.readAsArrayBuffer(file);
    });
}

function procesarTodo(v, inv, cxp, comp) {
    baseDatos = { prov: {}, fac: {}, vend: {}, cli: {}, productos: {} };

    // 1. MAPEAR COMPRAS (Diccionario: factura -> razon_social)
    comp.forEach(f => {
        let n = f.factura || f.documento || f.numero;
        if(n) baseDatos.fac[n.toString().trim()] = f.razon_social || f.proveedor;
    });

    // 2. INVENTARIO (referencia -> stock / costo)
    inv.forEach(f => {
        let r = f.referencia || f.codigo;
        if(r) {
            baseDatos.productos[r.toString().trim()] = { 
                s: parseFloat(f.totalinventario) || 0, 
                c: parseFloat(f.costo_ponderado_final) || 0 
            };
        }
    });

    // 3. VENTAS (El motor de ingresos)
    v.forEach(f => {
        let p = f.referencia_proveedor || "OTROS";
        let r = f.referencia ? f.referencia.toString().trim() : "";
        let d = (f.nombre_producto || "").toUpperCase();
        
        // Lógica de Combos Vibras Positivas
        if (p === "OTROS" || p === "" || p === "0") {
            if (d.includes("COLOMBINA")) p = "COLOMBINA";
            else if (d.includes("FAMILIA") || d.includes("SANCELA")) p = "FAMILIA SANCELA";
        }

        let subtotal = (parseFloat(f.valor) || 0) * (parseFloat(f.cantidad) || 0);
        if(f.tipo == "2" || f.tipo == 2) subtotal = -subtotal;

        if(!baseDatos.prov[p]) baseDatos.prov[p] = { vta: 0, deu: 0, invV: 0 };
        baseDatos.prov[p].vta += subtotal;

        if(baseDatos.productos[r]) {
            baseDatos.prov[p].invV += (baseDatos.productos[r].s * baseDatos.productos[r].c);
        }
    });

    // 4. CXP (La deuda real)
    cxp.forEach(f => {
        let n = f.documento || f.numero;
        let p = baseDatos.fac[n.toString().trim()] || f.razon_social || "OTROS";
        let s = parseFloat(f.totalvalor) || parseFloat(f.saldo) || 0;
        
        if(!baseDatos.prov[p]) baseDatos.prov[p] = { vta: 0, deu: 0, invV: 0 };
        baseDatos.prov[p].deu += s;
    });

    // Mostrar el Dashboard
    const sel = document.getElementById('selectProveedor');
    sel.innerHTML = '<option value="">-- ELIGE UN PROVEEDOR --</option>';
    Object.keys(baseDatos.prov).sort().forEach(p => {
        let opt = document.createElement('option');
        opt.value = p; opt.textContent = p; sel.appendChild(opt);
    });

    document.getElementById('welcomeMessage').style.display = 'none';
    document.getElementById('dashboardResult').style.display = 'block';
}

// ESTA FUNCIÓN ES LA QUE PINTA EL ANÁLISIS
document.getElementById('selectProveedor').onchange = function() {
    let p = this.value;
    if(!p) return;
    let d = baseDatos.prov[p];
    let rentable = d.vta > d.deu;
    let vtaDia = d.vta / 30;
    let diasInv = d.vta > 0 ? Math.round(d.invV / (vtaDia || 1)) : "Sin Ventas";

    document.getElementById('kpiContainer').innerHTML = `
        <div class="row text-center mb-4">
            <div class="col-4"><div class="p-3 border rounded">Ventas<br><h4 class="text-success">$${Math.round(d.vta).toLocaleString()}</h4></div></div>
            <div class="col-4"><div class="p-3 border rounded">Deuda CXP<br><h4 class="text-danger">$${Math.round(d.deu).toLocaleString()}</h4></div></div>
            <div class="col-4"><div class="p-3 border rounded">Stock (Días)<br><h4 class="text-warning">${diasInv}</h4></div></div>
        </div>
        <div class="alert ${rentable ? 'alert-success' : 'alert-danger'} text-center shadow">
            <h3>${rentable ? '🟢 PROVEEDOR RENTABLE' : '🔴 ALERTA: RIESGO FINANCIERO'}</h3>
        </div>
    `;

    document.getElementById('recomendacionesBox').innerHTML = `
        <div class="card border-primary">
            <div class="card-body">
                <h5 class="text-primary">📝 Estrategia CEO Distrileco:</h5>
                <p>${d.deu > d.vta ? 'Socio, frena compras. Este proveedor está consumiendo tu capital. Negocia devolución de stock muerto.' : 'Proveedor estrella. Tienes flujo positivo, pide bonificación por volumen.'}</p>
            </div>
        </div>`;
};
