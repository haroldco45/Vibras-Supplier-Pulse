// 1. Verificación inicial de conexión
console.log("Vibras Positivas: Script cargado con éxito");

document.getElementById('btnProcesar').onclick = async function() {
    alert("🚀 Iniciando lectura de archivos... ¡Deme un segundo, socio!");

    const vFile = document.getElementById('inputVentas').files[0];
    const iFile = document.getElementById('inputInventario').files[0];
    const cFile = document.getElementById('inputCXP').files[0];
    const compFile = document.getElementById('inputCompras').files[0];

    if (!vFile || !iFile || !cFile || !compFile) {
        alert("❌ ¡Error! Te faltan archivos por seleccionar. Revisa los 4 campos.");
        return;
    }

    try {
        alert("📂 Procesando Ventas e Inventario...");
        const vData = await leerExcelSimple(vFile);
        const iData = await leerExcelSimple(iFile);
        
        alert("📂 Procesando Compras y CXP...");
        const cData = await leerExcelSimple(cFile);
        const compData = await leerExcelSimple(compFile);

        alert("✅ Archivos leídos. ¡Generando el informe de Vibras Positivas!");
        
        // Aquí llamamos a la lógica de visualización
        renderizarInforme(vData, iData, cData, compData);

    } catch (e) {
        alert("🚨 Ocurrió un error técnico: " + e.message);
        console.error(e);
    }
};

async function leerExcelSimple(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = (e) => {
            try {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, {type: 'array'});
                const sheet = workbook.Sheets[workbook.SheetNames[0]];
                resolve(XLSX.utils.sheet_to_json(sheet, {header: "A"}));
            } catch (err) { reject(err); }
        };
        reader.onerror = () => reject(new Error("No se pudo leer el archivo"));
        reader.readAsArrayBuffer(file);
    });
}

function renderizarInforme(v, inv, cxp, comp) {
    // Esta función limpia el dashboard y muestra los datos
    document.getElementById('welcomeMessage').style.display = 'none';
    const dash = document.getElementById('dashboardResult');
    dash.style.display = 'block';
    
    // Mostramos un resumen rápido para probar que hay datos
    document.getElementById('kpiContainer').innerHTML = `
        <div class="alert alert-success">
            <h5>¡Éxito Socio!</h5>
            <p>Se procesaron ${v.length} registros de ventas y ${cxp.length} de cartera.</p>
            <p>Use el selector de arriba para filtrar (en desarrollo).</p>
        </div>
    `;
    
    alert("🔥 ¡Informe generado! Mira la pantalla a la derecha.");
}
