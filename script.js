let baseDatos = { prov: {}, fac: {}, vend: {}, cli: {}, productos: {} };

document.getElementById('btnProcesar').addEventListener('click', async function() {
    const files = {
        ventas: document.getElementById('inputVentas').files[0],
        inv: document.getElementById('inputInventario').files[0],
        cxp: document.getElementById('inputCXP').files[0],
        compras: document.getElementById('inputCompras').files[0]
    };

    if (!Object.values(files).every(f => f)) {
        alert("Socio, carga los 4 archivos para poder analizar."); return;
    }

    document.getElementById('welcomeMessage').innerHTML = '<h3 class="text-primary">Mapeando base de datos de Distrileco... ⏳</h3>';

    try {
        // Leemos con la primera fila como nombres de columnas
        const [v, inv, cxp, comp] = await Promise.all([
            leerExcel(files.ventas), leerExcel(files.inv),
            leerExcel(files.cxp), leerExcel(files.compras)
        ]);

        procesarTodo(v, inv, cxp, comp);
    } catch (error) {
        console.error(error);
        alert("Error al leer los archivos. Asegúrate que sean Excel o CSV.");
    }
});

function procesarTodo(v, inv, cxp, comp) {
    baseDatos = { prov: {}, fac: {}, vend: {}, cli: {}, productos: {} };

    // 1. MAPEAR COMPRAS (Columna: factura, razon_social)
    let mapaFacturas = {};
    comp.forEach(f => {
        let nroFac = f.factura || f.documento || f.numero;
        if(nroFac) mapaFacturas[nroFac.toString().trim()] = f.razon_social || f.proveedor;
    });

    // 2. INVENTARIO (Columna: referencia, totalinventario, costo_ponderado_final)
    inv.forEach(f => {
        let ref = f.referencia || f.codigo;
        if(ref) baseDatos.productos[ref.toString().trim()] = { 
            stock: parseFloat(f.totalinventario) || 0, 
            costo: parseFloat(f.costo_ponderado_final) || 0 
        };
    });

    // 3. VENTAS (Columna: referencia_proveedor, referencia, nombre_producto, cantidad, precio_base_venta, vendedor, razon_social, tipo)
    v.forEach(f => {
        let p = f.referencia_proveedor || "SIN PROVEEDOR";
        let desc = (f.nombre_producto || "").toUpperCase();
        let ref = f.referencia ? f.referencia.toString().trim() : "";
        
        // Lógica de Combos
        if (p === "SIN PROVEEDOR" || p === "") {
            if (desc.includes("COLOMBINA")) p = "COLOMBINA";
            else if (desc.includes("FAMILIA") || desc.includes("SANCELA")) p = "FAMILIA SANCELA";
        }

        let cant = parseFloat(f.cantidad) || 0;
        let precio = parseFloat(f.precio_base_venta) || 0;
        let subtotal = cant * precio;
        if(f.tipo == "2") subtotal = -subtotal; // Devoluciones

        if(!baseDatos.prov[p]) baseDatos.prov[p] = { vta: 0, deuda: 0, valorInv: 0 };
        baseDatos.prov[p].vta += subtotal;

        // Sumar valor de inventario al proveedor
        if(baseDatos.productos[ref]) {
            baseDatos.prov[p].valorInv += (baseDatos.productos[ref].stock * baseDatos.productos[ref].costo);
        }

        // Vendedores y Clientes
        let vend = f.vendedor || "OFICINA";
        if(!baseDatos.vend[vend]) baseDatos.vend[vend] = { vta: 0 };
        baseDatos.vend[vend].vta += subtotal;

        let cli = f.razon_social || "VARIO";
        if(!baseDatos.cli[cli]) baseDatos.cli[cli] = { vta: 0 };
        baseDatos.cli[cli].vta += subtotal;
    });

    // 4. CXP (Columna: documento, totalvalor o saldo)
    cxp.forEach(f => {
        let nroFac = f.documento || f.numero;
        let p = mapaFacturas[nroFac] || f.razon_social || "OTROS";
        let deuda = parseFloat(f.totalvalor) || parseFloat(f.saldo) || 0;
        if(!baseDatos.prov[p]) baseDatos.prov[p] = { vta: 0, deuda: 0, valorInv: 0 };
        baseDatos.prov[p].deuda += deuda;
    });

    document.getElementById('welcomeMessage').style.display = 'none';
    document.getElementById('dashboardContent').style.display = 'block';
    cambiarTab('proveedores');
}

function cambiarTab(tipo) {
    const areaSel = document.getElementById('areaSelector');
    let dataObj = baseDatos[tipo === 'proveedores' ? 'prov' : tipo === 'vendedores' ? 'vend' : tipo === 'clientes' ? 'cli' : 'fac'];
    let keys = Object.keys(dataObj).filter(k => k !== "undefined").sort();
    let options = `<option value="">-- SELECCIONA ${tipo.toUpperCase()} --</option>`;
    keys.forEach(k => options += `<option value="${k}">${k}</option>`);
    areaSel.innerHTML = `<select class="form-select p-3 shadow border-primary" onchange="verDetalle('${tipo}', this.value)">${options}</select>`;
}

function verDetalle(tipo, llave) {
    let data = baseDatos[tipo === 'proveedores' ? 'prov' : tipo === 'vendedores' ? 'vend' : tipo === 'clientes' ? 'cli' : 'fac'][llave];
    let html = `<div class="mt-4 p-4 border rounded bg-white shadow-lg border-primary">
                <h4 class="text-primary">ANÁLISIS: ${llave}</h4><hr>`;

    if(tipo === 'proveedores') {
        let vtaDia = data.vta / 30;
        let diasInv = data.vta > 0 ? Math.round(data.valorInv / (vtaDia || 1)) : "SIN VENTAS";
        let colorSemaforo = data.vta > data.deuda ? 'alert-success' : 'alert-danger';

        html += `
            <div class="row text-center mb-3">
                <div class="col-md-4"><h6>Ventas Mes</h6><h3 class="text-success">$${data.vta.toLocaleString()}</h3></div>
                <div class="col-md-4"><h6>Deuda</h6><h3 class="text-danger">$${data.deuda.toLocaleString()}</h3></div>
                <div class="col-md-4"><h6>Días Bodega</h6><h3 class="text-warning">${diasInv}</h3></div>
            </div>
            <div class="alert ${colorSemaforo} text-center">
                <h2>${data.vta > data.deuda ? '🟢 RENTABLE' : '🔴 NO RENTABLE'}</h2>
            </div>
            <div class="row">
                <div class="col-md-6"><div class="card p-2 border-primary"><h6>Acción Proveedor</h6><p>${data.deuda > data.vta ? 'Frenar compras ya.' : 'Pedir descuento pronto pago.'}</p></div></div>
                <div class="col-md-6"><div class="card p-2 border-success"><h6>Fuerza Ventas</h6><p>${diasInv > 30 ? 'Hacer combo con Arroz.' : 'Prioridad en ruta.'}</p></div></div>
            </div>`;
    } else {
        html += `<h2 class="text-center">$${data.vta.toLocaleString()}</h2><p class="text-center">Total Movimiento</p>`;
    }
    document.getElementById('areaResultados').innerHTML = html;
}

async function leerExcel(file) {
    return new Promise(res => {
        const reader = new FileReader();
        reader.onload = e => {
            const wb = XLSX.read(new Uint8Array(e.target.result), {type:'array'});
            res(XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]]));
        };
        reader.readAsArrayBuffer(file);
    });
}
