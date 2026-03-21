document.getElementById('btnProcesar').addEventListener('click', async function() {
    const files = {
        ventas: document.getElementById('inputVentas').files[0],
        inv: document.getElementById('inputInventario').files[0],
        cxp: document.getElementById('inputCXP').files[0],
        compras: document.getElementById('inputCompras').files[0]
    };

    if (!files.ventas || !files.inv || !files.cxp || !files.compras) {
        alert("Socio, carga los 4 archivos para que la inteligencia de Vibras funcione.");
        return;
    }

    const dVentas = await leerExcel(files.ventas, "A");
    const dInv = await leerExcel(files.inv, "A");
    const dCXP = await leerExcel(files.cxp, "A");
    const dCompras = await leerExcel(files.compras, "A");

    procesarInforme(dVentas, dInv, dCXP, dCompras);
});

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

function procesarInforme(v, inv, cxp, comp) {
    let proveedores = {};
    let mapaFacturas = {}; // Para saber de quién es cada factura

    // 1. LIMPIEZA DE COMPRAS (Columna E es Proveedor, B es Factura)
    comp.forEach(f => {
        if(f.B && f.E) mapaFacturas[f.B.toString().trim()] = f.E.toString().trim();
    });

    // 2. PROCESAR VENTAS (H: Prov, D: Desc, E*F: Total, A: Tipo)
    v.forEach(f => {
        if (!f.D) return;
        let p = f.H ? f.H.toString().trim() : "SIN_PROVEEDOR";
        let desc = f.D.toUpperCase();

        // Lógica de Combos en blanco (Tu genialidad)
        if (p === "SIN_PROVEEDOR" || p === "") {
            if (desc.includes("COLOMBINA")) p = "COLOMBINA";
            else if (desc.includes("FAMILIA") || desc.includes("SANCELA")) p = "FAMILIA SANCELA";
            else return; // Si no es ninguno, lo ignoramos para no ensuciar
        }

        if (!proveedores[p]) proveedores[p] = { ventas: 0, deuda: 0, stock: 0 };
        let subtotal = (parseFloat(f.E) || 0) * (parseFloat(f.F) || 0);
        proveedores[p].ventas += (f.A == 2 ? -subtotal : subtotal);
    });

    // 3. PROCESAR CXP (B: Factura, E: Saldo)
    cxp.forEach(f => {
        if(!f.B) return;
        let fac = f.B.toString().trim();
        let provReal = mapaFacturas[fac] || "OTROS_PROVEEDORES";
        
        if (!proveedores[provReal]) proveedores[provReal] = { ventas: 0, deuda: 0, stock: 0 };
        proveedores[provReal].deuda += (parseFloat(f.E) || 0);
    });

    // 4. LLENAR EL SELECTOR (SIN NÚMEROS DE FACTURA)
    const selector = document.getElementById('selectProveedor');
    selector.innerHTML = '<option value="">-- Selecciona un Proveedor Real --</option>';
    
    Object.keys(proveedores).sort().forEach(p => {
        if(p !== "OTROS_PROVEEDORES") {
            let opt = document.createElement('option');
            opt.value = p;
            opt.textContent = p;
            selector.appendChild(opt);
        }
    });

    // Mostrar Dashboard
    document.getElementById('welcomeMessage').style.display = 'none';
    document.getElementById('dashboardResult').style.display = 'block';

    // Evento al cambiar de proveedor
    selector.onchange = function() {
        let elegido = this.value;
        if(!elegido) return;
        
        let datos = proveedores[elegido];
        let semaforo = datos.ventas > datos.deuda ? '🟢 SALUDABLE' : '🔴 RIESGO FINANCIERO';
        
        document.getElementById('kpiContainer').innerHTML = `
            <div class="row">
                <div class="col-6 mb-3"><div class="kpi-box"><h5>Ventas</h5><h3>$${datos.ventas.toLocaleString()}</h3></div></div>
                <div class="col-6 mb-3"><div class="kpi-box"><h5>Deuda (CXP)</h5><h3>$${datos.deuda.toLocaleString()}</h3></div></div>
            </div>
            <div class="alert ${datos.ventas > datos.deuda ? 'alert-success' : 'alert-danger'}">
                <strong>Estado: ${semaforo}</strong><br>
                ${datos.ventas < datos.deuda ? 'Socio, estás debiendo más de lo que vendes. ¡Toca apretar al proveedor!' : 'Vas bien, el flujo de caja soporta la deuda.'}
            </div>
        `;
    };
}
