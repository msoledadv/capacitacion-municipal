let datosAgrupados = [];
let nombreSecretariaDetectada = "";

// Parámetros de la URL para ver qué CSV cargar (Ej: ?csv=secretaria2)
const urlParams = new URLSearchParams(window.location.search);
const archivoCSV = urlParams.get('csv') ? `${urlParams.get('csv')}.csv` : 'secretaria1.csv';
console.log("Archivo CSV solicitado para carga:", archivoCSV);

// Carga principal con PapaParse
Papa.parse(archivoCSV, {
    download: true,
    header: true,
    dynamicTyping: true,
    skipEmptyLines: true,
    complete: function(results) {
        console.log("CSV Cargado exitosamente. Filas detectadas:", results.data.length);
        procesarDatos(results.data);
    },
    error: function(err) {
        console.error(`Error al cargar el archivo ${archivoCSV}:`, err);
        // Auxiliar: si falla el dinámico, carga por defecto el 1
        Papa.parse("secretaria1.csv", {
            download: true,
            header: true,
            dynamicTyping: true,
            skipEmptyLines: true,
            complete: function(res) { procesarDatos(res.data); }
        });
    }
});

// Convierte fechas raras de Excel o ISO a formato DD/MM/AAAA
function formatearFecha(val) {
    if (val === undefined || val === null || String(val).trim() === "" || String(val).toLowerCase() === "s/d") {
        return 's/d';
    }
    let num = Number(val);
    if (!isNaN(num) && num > 0) {
        if (num > 30000 && num < 60000) {
            try {
                const utc_days  = Math.floor(num - 25569);
                const utc_value = utc_days * 86400;
                const date_info = new Date(utc_value * 1000);
                const dia = String(date_info.getUTCDate()).padStart(2, '0');
                const mes = String(date_info.getUTCMonth() + 1).padStart(2, '0');
                const anio = date_info.getUTCFullYear();
                return `${dia}/${mes}/${anio}`;
            } catch(e) { console.error(e); }
        }
    }
    const stringFecha = String(val).trim();
    if (stringFecha.includes('T')) {
        const partes = stringFecha.split('T')[0].split('-');
        if(partes.length === 3) return `${partes[2]}/${partes[1]}/${partes[0]}`;
    }
    return stringFecha || 's/d';
}

// Agrupa las filas del CSV por legajo y calcula los totales
function procesarDatos(filas) {
    const mapa = new Map();
    let uSec = "", uOfi = "SIN OFICINA", uLeg = "", uNom = "", uCar = "";

    // Parsea números y limpia las comas de Excel
    const num = (v) => {
        if (v === undefined || v === null || v === "") return 0;
        let n = v.toString().replace(',', '.');
        return parseFloat(n) || 0;
    };

    filas.forEach(fila => {
        // Lógica Fill Down para celdas combinadas del reporte
        if (fila["SECRETARIA"] && fila["SECRETARIA"].toString().trim() !== "") uSec = fila["SECRETARIA"].toString().trim();
        if (fila["OFICINA"] && fila["OFICINA"].toString().trim() !== "") uOfi = fila["OFICINA"].toString().trim();
        if (fila["LEGAJO"]) uLeg = fila["LEGAJO"].toString().trim();
        if (fila["NOMBRE COMPLETO"] && fila["NOMBRE COMPLETO"].toString().trim() !== "") uNom = fila["NOMBRE COMPLETO"].toString().trim();
        if (fila["CARGO ESCALAFON"] && fila["CARGO ESCALAFON"].toString().trim() !== "") uCar = fila["CARGO ESCALAFON"].toString().trim();

        if (!uLeg || uLeg === "0") return; 

        if (!mapa.has(uLeg)) {
            mapa.set(uLeg, {
                LEGAJO: uLeg,
                NOMBRE: uNom || 'Sin Nombre',
                SECRETARIA: uSec || 'General',
                OFICINA: uOfi,
                CARGO: uCar || 's/d',
                CURSOS: [], 
                CREDITOS: 0,
                OBJETIVO: 0,
                SALDO_RESTANTE: 0
            });
        }

        const p = mapa.get(uLeg);
        
        const cursoVal = fila["CAPACITACION"];
        if (cursoVal && cursoVal.toString().trim() !== "0" && cursoVal.toString().trim() !== "" && cursoVal.toString().toLowerCase() !== "s/d") {
            const fechaVal = formatearFecha(fila["Fecha Aprobación"]);
            
            const yaExiste = p.CURSOS.some(c => c.nombre === cursoVal.toString().trim() && c.fecha === fechaVal);
            if (!yaExiste) {
                p.CURSOS.push({
                    nombre: cursoVal.toString().trim(),
                    fecha: fechaVal
                });
            }
        }
        
        let creditosFila = num(fila["Suma de CREDITOS"]);
        let objetivoFila = num(fila["Suma de OBJETIVO"]);
        let saldoFila = num(fila["Suma de SALDO RESTANTE"]);

        p.CREDITOS = creditosFila;
        if (objetivoFila > 0) p.OBJETIVO = objetivoFila;
        p.SALDO_RESTANTE = saldoFila; 
    });

    datosAgrupados = Array.from(mapa.values());
    
    if (datosAgrupados.length > 0) {
        nombreSecretariaDetectada = datosAgrupados[0].SECRETARIA;
        actualizarInterfazTitulo();
    }
    
    poblarCargos();
    poblarOficinas(); 
    renderTable(datosAgrupados);
    inicializarEventos();
}

function actualizarInterfazTitulo() {
    const txtSec = document.getElementById('nombreSecretariaHeader');
    if (txtSec) {
        txtSec.innerText = nombreSecretariaDetectada;
    }
}

// Listeners de los filtros del sidebar
function inicializarEventos() {
    const ids = ['selectOficina', 'selectCargo', 'inputNombre', 'inputLegajo', 'selectEstado'];
    
    ids.forEach(id => {
        const el = document.getElementById(id);
        if (el) {
            el.addEventListener(id.includes('input') ? 'input' : 'change', () => {
                filtrar();
            });
        }
    });

    document.getElementById('btnLimpiar').addEventListener('click', () => {
        ids.forEach(id => {
            const el = document.getElementById(id);
            if (el) el.value = "";
        });
        renderTable(datosAgrupados);
    });
}

// Lógica de filtrado en tiempo real
function filtrar() {
    const ofi = document.getElementById('selectOficina').value;
    const car = document.getElementById('selectCargo').value;
    const est = document.getElementById('selectEstado').value; 
    const nom = document.getElementById('inputNombre').value.toLowerCase().trim();
    const leg = document.getElementById('inputLegajo').value.toLowerCase().trim();

    const filtrados = datosAgrupados.filter(p => {
        
        let estadoReal = "SIN INICIAR";
        if (p.SALDO_RESTANTE <= 0) {
            estadoReal = "COMPLETO";
        } else if (p.CREDITOS > 0 && p.SALDO_RESTANTE > 0) {
            estadoReal = "EN PROCESO";
        }

        const matchOfi = (ofi === "" || p.OFICINA === ofi);
        const matchCar = (car === "" || p.CARGO === car);
        const matchEst = (est === "" || estadoReal === est);
        const matchNom = (nom === "" || p.NOMBRE.toLowerCase().includes(nom));
        const matchLeg = (leg === "" || p.LEGAJO.toString().toLowerCase().includes(leg));

        return matchOfi && matchCar && matchEst && matchNom && matchLeg;
    });

    renderTable(filtrados);
}

// Dibuja las filas de los agentes y procesa las capacitaciones vigentes e históricas
function renderTable(data) {
    const tbody = document.getElementById('tbody');
    tbody.innerHTML = '';
    
    if (data.length === 0) {
        tbody.innerHTML = `<tr><td colspan="8" style="text-align:center; color:#777; padding: 25px;">No se encontró personal con los filtros seleccionados.</td></tr>`;
        document.getElementById('contador').innerText = `Personal encontrado: 0`;
        return;
    }

    data.forEach(p => {
        const faltanVisual = p.SALDO_RESTANTE < 0 ? 0 : p.SALDO_RESTANTE;
        
        let clase = "pendiente", texto = "🚨 SIN INICIAR";
        if (p.SALDO_RESTANTE <= 0) {
            clase = "cumplido"; texto = "✅ COMPLETO";
        } else if (p.CREDITOS > 0 && p.SALDO_RESTANTE > 0) {
            clase = "proceso"; texto = "⏳ EN PROCESO";
        }

        let listaCursosVisual = "";
        if (p.CURSOS && p.CURSOS.length > 0) {
            listaCursosVisual = `<ul style="margin:0; padding-left:12px; list-style-type:disc;">`;
            
            p.CURSOS.forEach(c => {
                // Filtro para separar lo anterior al 01/04/2024 (2022 y 2023 pasan a gris)
                const esHistorico = c.fecha.includes('/2022') || c.fecha.includes('/2023');
                
                const colorTexto = esHistorico ? '#9ca3af' : '#1f2937'; 
                const colorFecha = esHistorico ? '#cbd5e1' : '#666';
                const etiquetaHistorico = esHistorico ? ' <span style="font-size:10px; font-style:italic; font-weight:normal; color:#cbd5e1;">(Histórico)</span>' : '';

                listaCursosVisual += `
                    <li style="color: ${colorTexto};">
                        <strong style="font-weight: ${esHistorico ? 'normal' : 'bold'};">${c.nombre}</strong> 
                        <span style="color:${colorFecha}; font-size:11px; margin-left:3px;">(${c.fecha})</span>${etiquetaHistorico}
                    </li>`;
            });
            
            listaCursosVisual += `</ul>`;
        } else {
            listaCursosVisual = '<span style="color:#aaa; font-style:italic;">Sin cursos registrados</span>';
        }

        const tr = document.createElement('tr');
        tr.innerHTML = `
            <td><mark style="background:none; font-weight:bold; color:#0056b3; font-family:monospace;">${p.LEGAJO}</mark></td>
            <td><strong>${p.NOMBRE}</strong><br><small style="color:#555; font-weight:500;">${p.OFICINA}</small></td>
            <td><small>${p.CARGO}</small></td>
            <td class="col-capa">${listaCursosVisual}</td>
            <td style="text-align:center; font-weight:bold; color:#0056b3;">${p.CREDITOS.toFixed(1).replace('.0', '')}</td>
            <td style="text-align:center; color:#444;">${p.OBJETIVO.toFixed(1).replace('.0', '')}</td>
            <td style="text-align:center; font-weight:bold; color:${faltanVisual > 0 ? '#b45309' : '#10b981'}">
                ${faltanVisual.toFixed(1).replace('.0', '')}
            </td>
            <td><span class="badge ${clase}">${texto}</span></td>
        `;
        tbody.appendChild(tr);
    });
    
    document.getElementById('contador').innerText = `Personal total filtrado: ${data.length}`;
}

// Carga las opciones únicas en el select de oficinas
function poblarOficinas() {
    const sOfi = document.getElementById('selectOficina');
    if (!sOfi) return;
    sOfi.innerHTML = '<option value="">Todas las Oficinas</option>';
    
    const oficinas = [...new Set(datosAgrupados.map(p => p.OFICINA))].filter(Boolean).sort();
    oficinas.forEach(o => sOfi.innerHTML += `<option value="${o}">${o}</option>`);
}

// Carga las opciones únicas en el select de cargos
function poblarCargos() {
    const sCar = document.getElementById('selectCargo');
    if (!sCar) return;
    sCar.innerHTML = '<option value="">Todos los Cargos</option>';
    const cargos = [...new Set(datosAgrupados.map(p => p.CARGO))].filter(Boolean).sort();
    cargos.forEach(c => sCar.innerHTML += `<option value="${c}">${c}</option>`);
}
