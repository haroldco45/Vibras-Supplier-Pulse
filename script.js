// Variable global para guardar los datos procesados
let proveedoresData = {};

document.getElementById('btnProcesar').onclick = async function() {
    const vFile = document.getElementById('inputVentas').files[0];
    const iFile = document.getElementById('inputInventario').files[0];
    const cFile = document.getElementById('inputCXP').files[0];
    const compFile = document.getElementById('inputCompras').files[0];

    if (!vFile || !iFile || !cFile || !compFile) {
        alert("❌ Socio, carga los 4 archivos primero.");
        return;
    }

    try {
        const v = await leerExcel(vFile);
        const inv = await leerExcel(iFile);
        const cxp = await leerExcel(cFile);
        const comp = await leerExcel(compFile);

        procesarLogica(v, inv, cxp, comp);
        alert("✅ Datos procesados. Selecciona un proveedor arriba.");
    } catch (e) {
        alert("🚨 Error: " + e.message);
    }
};

async function leerExcel(file) {
    return new Promise((resolve) => {
        const reader = new FileReader();
        reader.onload = (e) => {
            const data = new Uint8Array(e.target.result);
            const wb = XLSX.read(data, {type: 'array'});
            resolve(XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], {header: "A"}));
        };
        reader.readAsArrayBuffer(file);
    });
}

function procesarLogica(v, inv, cxp, comp) {
    proveedoresData = {};
    let facturasAProv = {};
    
    // 1. Mapear Compras (B factura, E proveedor)
    comp.forEach(f => { if(f.B) facturasAProv[f.B.toString().trim()] = f.E; });

    // 2. Procesar Ventas y Combos
    v.forEach((f, idx) => {
        if(idx === 0) return;
        let p = f.H ? f.H.toString().trim().toUpperCase() : "";
        let d = f.D ? f.D.toString().toUpperCase() : "";

        if (!p || p === "") {
            if (d.includes("COLOMBINA")) p = "COLOMBINA";
            else if (d.includes("FAMILIA") || d.includes("SANCELA")) p = "FAMILIA SANCELA";
            else p = "OTROS / COMBOS";
        }

        if (!proveedoresData[p]) proveedoresData[p] = { vta: 0, deuda: 0, desc: p };
        let total = (parseFloat(f.E) || 0) * (parseFloat(f.F) || 0);
        proveedoresData[p].vta += (f.A == 2 ? -total : total);
    });

    // 3. Procesar CXP (B factura, E saldo)
    cxp.forEach((f, idx) => {
        if(idx === 0) return;
        let idFact = f.B ? f.B.toString().trim() : "";
        let pReal = facturasAProv[idFact] || f.A || "OTROS";
        pReal = pReal.toString().toUpperCase();
        
        if (!proveedoresData[pReal]) proveedoresData[pReal] = { vta: 0, deuda: 0, desc: pReal };
        proveedoresData[pReal].deuda += (parseFloat(f.E) || 0);
    });

    // 4. Llenar Selector
    const select = document.getElementById('selectProveedor');
    select.innerHTML = '<option value="">-- Selecciona --</option>';
    Object.keys(proveedoresData).sort().forEach(p => {
        select.innerHTML += `<option value="${p}">${p}</option>`;
    });

    document.getElementById('welcomeMessage').style.display = 'none';
    document.getElementById('dashboardResult').style.display = 'block';
}

// Lógica al cambiar el proveedor
document.getElementById('selectProveedor').onchange = function() {
    const p = this.value;
    if(!p) return;
    const d = proveedoresData[p];
    const esSano = d.vta > d.deuda;

    document.getElementById('kpiContainer').innerHTML = `
        <div class="row text-center mt-3">
            <div class="col-6"><div class="p-3 border rounded bg-white"><h6>Ventas</h6><h4 class="text-primary">$${d.vta.toLocaleString()}</h4></div></div>
            <div class="col-6"><div class="p-3 border rounded bg-white"><h6>Deuda</h6><h4 class="text-danger">$${d.deuda.toLocaleString()}</h4></div></div>
        </div>
        <div class="alert alert-${esSano?'success':'danger'} mt-3">
            <h5>${esSano?'Socio Estratégico':'⚠️ Alerta de Liquidez'}</h5>
            <p>${esSano?'La rotación es mayor a la deuda.':'La deuda supera las ventas. Estrategia: Pedir más plazo.'}</p>
        </div>
    `;
};

// BOTÓN PDF
document.getElementById('btnPDF').onclick = function() {
    const p = document.getElementById('selectProveedor').value;
    if(!p) { alert("Socio, elige un proveedor primero."); return; }
    
    const { jsPDF } = window.jspdf;
    const doc = new jsPDF();
    const d = proveedoresData[p];

    doc.setFontSize(18);
    doc.text("Informe de Negociación - Vibras Positivas", 10, 20);
    doc.setFontSize(12);
    doc.text(`Proveedor: ${p}`, 10, 35);
    doc.text(`Ventas Mes: $${d.vta.toLocaleString()}`, 10, 45);
    doc.text(`Deuda Actual: $${d.deuda.toLocaleString()}`, 10, 55);
    doc.text("--------------------------------------------------", 10, 65);
    doc.text("App desarrollada por Vibras Positivas (3117700431)", 10, 80);
    
    doc.save(`Negociacion_${p}.pdf`);
};
