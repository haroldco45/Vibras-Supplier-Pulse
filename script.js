document.getElementById('btnProcesar').addEventListener('click', async function() {
    const files = {
        ventas: document.getElementById('inputVentas').files[0],
        inv: document.getElementById('inputInventario').files[0],
        cxp: document.getElementById('inputCXP').files[0],
        compras: document.getElementById('inputCompras').files[0]
    };

    if (!files.ventas || !files.inv || !files.cxp || !files.compras) {
        alert("Socio, carga los 4 archivos para que Vibras Positivas haga la magia.");
        return;
    }

    const btn = this;
    btn.innerText = "Procesando Datos...";
    btn.disabled = true;

    try {
        const dataVentas = await leerExcel(files.ventas, "A");
        const dataInv = await leerExcel(files.inv, "A");
        const dataCXP = await leerExcel(files.cxp, "A");
        const dataCompras = await leerExcel(files.compras, "A");

        generarDashboard(dataVentas, dataInv, dataCXP, dataCompras);
    } catch (error) {
        console.error(error);
        alert("Hubo un error leyendo los archivos. Revisa que sean Excel válidos.");
    } finally {
        btn.innerText = "GENERAR INFORME MAESTRO";
        btn.disabled = false;
    }
});

async function leerExcel(file, header) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = (e) => {
            try {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, {type: 'array'});
                const sheet = workbook.Sheets[workbook.SheetNames[0]];
                resolve(XLSX.utils.sheet_to_json(sheet, {header: header}));
            } catch (err) { reject(err); }
        };
        reader.readAsArrayBuffer(file);
    });
}

function generarDashboard(v, inv, cxp, comp) {
    let proveedores = {};

    // 1. Mapear Compras para identificar facturas (Llave: Columna B)
    let facturasAProveedor = {};
    comp.forEach(f => { if(f.B) facturasAProveedor[f.B.toString().trim()] = f.E; });

    // 2. Procesar Ventas y Combos (H prov, D desc, E cant, F precio)
    v.forEach((f, idx) => {
        if(idx === 0) return;
        let prov = f.H ? f.H.toString().trim().toUpperCase() : "SIN_PROVEEDOR";
        let desc = f.D ? f.D.toString().toUpperCase() : "";

        // Lógica de Combos Vibras Positivas
        if (prov === "SIN_PROVEEDOR" || prov === "") {
            if (desc.includes("COLOMBINA")) prov = "COLOMBINA";
            else if (desc.includes("FAMILIA") || desc.includes("SANCELA")) prov = "FAMILIA SANCELA";
        }

        if (!proveedores[prov]) proveedores[prov] = { ventas: 0, deuda: 0, items: 0 };
        let total = (parseFloat(f.E) || 0) * (parseFloat(f.F) || 0);
        proveedores[prov].ventas += (f.A == 2 ? -total : total);
        proveedores[prov].items++;
    });

    // 3. Procesar CXP (B factura, E valor/saldo)
    cxp.forEach((f, idx) => {
        if(idx === 0) return;
        let idFactura = f.B ? f.B.toString().trim() : "";
        let provReal = facturasAProveedor[idFactura] || f.A || "OTROS";
        provReal = provReal.toString().toUpperCase();

        if (!proveedores[provReal]) proveedores[provReal] = { ventas: 0, deuda: 0, items: 0 };
        proveedores[provReal].deuda += (parseFloat(f.E) || 0);
    });

    // 4. Llenar el Selector de Proveedores
    const select = document.getElementById('selectProveedor');
    select.innerHTML = '<option value="">-- Elige un proveedor --</option>';
    Object.keys(proveedores).sort().forEach(p => {
        let opt = document.createElement('option');
        opt.value = p;
        opt.innerText = p;
        select.appendChild(opt);
    });

    document.getElementById('welcomeMessage').style.display = 'none';
    document.getElementById('dashboardResult').style.display = 'block';

    // Función para actualizar los KPIs al cambiar el selector
    select.onchange = function() {
        const p = this.value;
        if(!p) return;
        
        const data = proveedores[p];
        const esSano = data.ventas > data.deuda;
        const color = esSano ? 'success' : 'danger';

        document.getElementById('kpiContainer').innerHTML = `
            <div class="row text-center">
                <div class="col-6 mb-3">
                    <div class="p-3 bg-white border rounded shadow-sm">
                        <h6>Ventas del Mes</h6>
                        <h4 class="text-primary">$${data.ventas.toLocaleString()}</h4>
                    </div>
                </div>
                <div class="col-6 mb-3">
                    <div class="p-3 bg-white border rounded shadow-sm">
                        <h6>Deuda Actual (CXP)</h6>
                        <h4 class="text-danger">$${data.deuda.toLocaleString()}</h4>
                    </div>
                </div>
            </div>
            <div class="alert alert-${color} mt-3">
                <h5 class="alert-heading">Semáforo: ${esSano ? 'PROVEEDOR RENTABLE' : 'ALERTA DE LIQUIDEZ'}</h5>
                <p>Análisis: ${esSano ? 'Este proveedor tiene una rotación sana.' : 'La deuda acumulada supera las ventas del mes. ¡Atención!'}</p>
            </div>
        `;

        document.getElementById('recomendacionesBox').innerHTML = `
            <div class="card bg-light border-0 p-3">
                <h6>🚀 Recomendación de Vibras Positivas:</h6>
                <p class="small italic">${esSano ? 'Pide bonificaciones por volumen y mantén el stock.' : 'Socio, negocia 15 días extra de plazo o haz un combo con Arroz Superarroz.'}</p>
            </div>
        `;
    };
}

