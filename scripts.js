let oficinas = [];
let personas = [];

/* ===============================
   BOTÓN PROCESAR
=============================== */
document.getElementById("btnProcesar").addEventListener("click", procesar);

/* ===============================
   CARGA XLSX AUTOMÁTICA
   (también soporta apertura por file://)
=============================== */
function cargarXLSXAutomatico() {
    const statusEl = document.getElementById("xlsxStatus");

    fetch("OFICINAS NOMENCLATURAS.xlsx")
        .then(res => {
            if (!res.ok) throw new Error("HTTP " + res.status);
            return res.arrayBuffer();
        })
        .then(data => procesarXLSX(data, statusEl))
        .catch(() => {
            // Protocolo file:// bloquea fetch — pedir al usuario que seleccione el archivo
            statusEl.textContent = "⚠️ No se pudo cargar el catálogo automáticamente. Selecciona el XLSX manualmente.";
            statusEl.className = "status-badge status-warn";

            const container = document.getElementById("xlsxStatusContainer");
            const label = document.createElement("label");
            label.className = "file-btn";
            label.style.marginTop = "8px";
            label.innerHTML = `<span>📁 Seleccionar OFICINAS NOMENCLATURAS.xlsx</span>
                <input type="file" accept=".xlsx" hidden id="xlsxManual">`;
            container.appendChild(label);

            document.addEventListener("change", e => {
                if (e.target.id !== "xlsxManual") return;
                const reader = new FileReader();
                reader.onload = ev => procesarXLSX(ev.target.result, statusEl);
                reader.readAsArrayBuffer(e.target.files[0]);
            });
        });
}

function procesarXLSX(data, statusEl) {
    const wb = XLSX.read(data, { type: "array" });
    const hoja = wb.Sheets[wb.SheetNames[0]];

    oficinas = XLSX.utils.sheet_to_json(hoja, { defval: "" })
        .map(o => {
            let claveOf = String(o.CLAVE).trim();

            // Normalizar claves numéricas a 4 dígitos
            if (/^\d+$/.test(claveOf)) {
                claveOf = claveOf.padStart(4, "0");
            }

            return {
                CLAVE: claveOf,
                NOMBRE: String(o.NOMBRE).trim(),
                NOMENCLATURA: String(o.NOMENCLATURA).trim().toUpperCase(),
                DIRECCION: String(o.DIRECCION).trim()
            };
        })
        .filter(o => o.CLAVE && o.NOMENCLATURA);

    statusEl.textContent = `✅ Catálogo cargado: ${oficinas.length} oficinas`;
    statusEl.className = "status-badge status-ok";
    console.log("📘 Catálogo de oficinas cargado:", oficinas.length);
}

/* ===============================
   CARGA CSV
=============================== */
document.getElementById("csvPersonas").addEventListener("change", e => {
    const archivo = e.target.files[0];
    if (!archivo) return;

    // Mostrar nombre del archivo seleccionado
    document.getElementById("csvNombre").textContent = "📄 " + archivo.name;

    const reader = new FileReader();

    reader.onload = ev => {
        const lines = ev.target.result.replace(/\r/g, "")
            .split("\n")
            .filter(l => l.trim());

        if (lines.length < 2) {
            document.getElementById("csvStatus").textContent = "❌ CSV vacío o sin datos";
            document.getElementById("csvStatus").className = "status-badge status-error";
            return;
        }

        // Detectar separador
        let sep = ",";
        if (lines[0].includes("\t")) sep = "\t";
        else if (lines[0].includes(";")) sep = ";";

        // Función para limpiar cada celda: quitar comillas y espacios
        const limpiar = str => str.trim().replace(/^"|"$/g, "").trim();

        const headers = lines[0].split(sep).map(h => limpiar(h).toLowerCase());

        const idxClave   = headers.findIndex(h => h.includes("clave"));
        const idxActivo  = headers.findIndex(h => h.includes("activo"));
        const idxNombre  = headers.findIndex(h =>
            h.includes("nombre") || h.includes("descripci") // cubre "descripción" e "descripcion"
        );

        if (idxClave === -1) {
            document.getElementById("csvStatus").textContent = "❌ El CSV no tiene columna 'Clave'";
            document.getElementById("csvStatus").className = "status-badge status-error";
            return;
        }

        personas = [];

        for (let i = 1; i < lines.length; i++) {
            const cols = lines[i].split(sep).map(limpiar); // ← quitar comillas en TODOS los valores
            const clave = (cols[idxClave] || "").trim().toUpperCase();
            if (!clave) continue;

            // Extraer nomenclatura: primeros 3 caracteres alfabéticos del inicio
            const matchNom = clave.match(/^([A-Z]{3})/);
            const nomenclatura = matchNom ? matchNom[1] : clave.substring(0, 3);

            personas.push({
                clave,
                claveCompleta: clave,
                nombreReal: idxNombre >= 0
                    ? (cols[idxNombre] || "").trim()
                    : (cols[1] || "").trim(),   // fallback columna 2
                nomenclatura,
                activo: idxActivo >= 0
                    ? (cols[idxActivo] || "").trim().toLowerCase() === "true"
                    : true // si no hay columna activo, asumir activo
            });
        }

        const total = personas.length;
        const activos = personas.filter(p => p.activo).length;
        const statusEl = document.getElementById("csvStatus");
        statusEl.textContent = `✅ ${total} claves cargadas (${activos} activas)`;
        statusEl.className = "status-badge status-ok";
        console.log("📗 CSV cargado:", total, "registros");
    };

    // Intentar UTF-8; si hay caracteres inválidos, releer como Latin-1
    reader.readAsText(archivo, "UTF-8");
    reader.onerror = () => reader.readAsText(archivo, "ISO-8859-1");
});

/* ===============================
   PROCESAR
=============================== */

function procesar() {
    // Validaciones previas
    const empleadoRaw = document.getElementById("empleado").value.trim();
    if (!empleadoRaw) {
        alert("⚠️ Ingresa el número de empleado");
        return;
    }
    if (!personas.length) {
        alert("⚠️ Carga primero el archivo CSV con las claves del usuario");
        return;
    }
    if (!oficinas.length) {
        alert("⚠️ El catálogo de oficinas no está disponible. Espera a que cargue o selecciónalo manualmente.");
        return;
    }

    const empleado = empleadoRaw.padStart(6, "0");

    // Leer oficinas a asignar (normalizar a mayúsculas y recortar)
    const oficinasInput = document.getElementById("oficinasInput").value
        .split("\n")
        .map(o => o.trim().toUpperCase())
        .filter(Boolean);

    const oficinasQuitarRaw = document.getElementById("oficinasQuitarInput").value
        .split("\n")
        .map(o => o.trim().toUpperCase())
        .filter(Boolean);

    // No se necesita detectar modo: cada valor se busca directamente en el catálogo

    // Mapa rápido por CLAVE y por NOMENCLATURA
    const mapOficinas   = {};
    const mapPorNom     = {};
    oficinas.forEach(o => {
        mapOficinas[o.CLAVE] = o;
        mapPorNom[o.NOMENCLATURA] = o;
    });

    // Índices del CSV del usuario
    const nomUsuarioActivas   = new Set();
    const nomUsuarioInactivas = new Set();
    const personasActivasPorNom = {};

    personas.forEach(p => {
        if (p.activo) {
            nomUsuarioActivas.add(p.nomenclatura);
            if (!personasActivasPorNom[p.nomenclatura]) {
                personasActivasPorNom[p.nomenclatura] = [];
            }
            personasActivasPorNom[p.nomenclatura].push(p);
        } else {
            nomUsuarioInactivas.add(p.nomenclatura);
        }
    });

    const tiene = [], faltan = [], inactivas = [], bajas = [], errores = [];
    const debeTener = new Set();
    const oficinasQuitarSet = new Set();

    // Construir set de oficinas a quitar
    // Estrategia: buscar cada valor como CLAVE (con padding si es numérico),
    // luego como NOMENCLATURA directa. Así soporta: "3472", "1C", "V417", "MXS", etc.
    let errorEnQuitar = false;
    oficinasQuitarRaw.forEach(o => {
        if (errorEnQuitar) return;

        // Intentar como CLAVE numérica (padding a 4 dígitos)
        const claveNorm = /^\d+$/.test(o) ? o.padStart(4, "0") : o;

        // Buscar en el catálogo: primero por CLAVE, luego por NOMENCLATURA
        const of = mapOficinas[claveNorm] || mapPorNom[o];

        if (of) {
            oficinasQuitarSet.add(of.NOMENCLATURA);
        } else {
            alert(`⚠️ "${o}" no se encontró como CLAVE ni como NOMENCLATURA en el catálogo XLSX.`);
            errorEnQuitar = true;
        }
    });

    if (errorEnQuitar) return;

    // Clasificar oficinas a asignar
    oficinasInput.forEach(cod => {
        // Normalizar numérica a 4 dígitos
        let codigoNorm = /^\d+$/.test(cod) ? cod.padStart(4, "0") : cod;

        const of = mapOficinas[codigoNorm] || mapPorNom[codigoNorm];
        if (!of) {
            errores.push([cod, "Oficina no existe en catálogo XLSX"]);
            return;
        }

        const claveFinal = of.NOMENCLATURA + "0" + empleado;
        debeTener.add(of.NOMENCLATURA);

        if (nomUsuarioActivas.has(of.NOMENCLATURA)) {
            tiene.push([of.CLAVE, claveFinal, of.NOMBRE, of.NOMENCLATURA, of.DIRECCION, "ACTIVA"]);
        } else if (nomUsuarioInactivas.has(of.NOMENCLATURA)) {
            inactivas.push([of.CLAVE, claveFinal, of.NOMBRE, of.NOMENCLATURA, of.DIRECCION, "REACTIVAR"]);
        } else {
            faltan.push([of.CLAVE, claveFinal, of.NOMBRE, of.NOMENCLATURA, of.DIRECCION, "AGREGAR"]);
        }
    });

    // Calcular bajas
    const modoQuitarEspecifico = oficinasQuitarSet.size > 0;

    nomUsuarioActivas.forEach(nom => {
        // Nunca inhabilitar algo que se va a agregar
        if (debeTener.has(nom)) return;

        // Si hay lista específica a quitar, solo esas
        if (modoQuitarEspecifico && !oficinasQuitarSet.has(nom)) return;

        const personasCsv = personasActivasPorNom[nom] || [];
        personasCsv.forEach(p => {
            bajas.push([
                p.claveCompleta,
                p.nombreReal,
                nom,
                "ACTIVA EN CSV",
                "INHABILITAR"
            ]);
        });
    });

    // Pintar tablas y mostrar conteos
    pintar("tablaTiene",    ["N° OFICINA", "CLAVE", "NOMBRE", "NOM", "DIR", "ACCIÓN"], tiene,    "emptyTiene",    "conteoTiene");
    pintar("tablaInactivas",["N° OFICINA", "CLAVE", "NOMBRE", "NOM", "DIR", "ACCIÓN"], inactivas,"emptyInactivas","conteoInactivas");
    pintar("tablaFaltan",   ["N° OFICINA", "CLAVE", "NOMBRE", "NOM", "DIR", "ACCIÓN"], faltan,   "emptyFaltan",   "conteoFaltan");
    pintar("tablaBajas",    ["CLAVE CSV", "DESCRIPCIÓN CSV", "NOM", "ESTADO CSV", "ACCIÓN"], bajas, "emptyBajas", "conteoBajas");
    pintar("tablaErrores",  ["OFICINA", "ERROR"], errores, "emptyErrores", "conteoErrores");

    // Resumen general
    mostrarResumen(tiene, inactivas, faltan, bajas, errores);

    // Hacer scroll al resultado
    setTimeout(() => {
        document.querySelector(".dashboard").scrollIntoView({ behavior: "smooth" });
    }, 100);
}

/* ===============================
   RESUMEN BANNER
=============================== */
function mostrarResumen(tiene, inactivas, faltan, bajas, errores) {
    const banner = document.getElementById("resumenBanner");
    banner.style.display = "flex";
    banner.innerHTML = `
        <div class="resumen-item"><span class="dot dot-green"></span><strong>${tiene.length}</strong> activas</div>
        <div class="resumen-item"><span class="dot dot-indigo"></span><strong>${inactivas.length}</strong> reactivar</div>
        <div class="resumen-item"><span class="dot dot-yellow"></span><strong>${faltan.length}</strong> agregar</div>
        <div class="resumen-item"><span class="dot dot-red"></span><strong>${bajas.length}</strong> inhabilitar</div>
        <div class="resumen-item"><span class="dot dot-dark"></span><strong>${errores.length}</strong> inválidas</div>
    `;
}

/* ===============================
   PINTAR TABLAS
=============================== */
function pintar(idTabla, headers, data, idEmpty, idConteo) {
    const t = document.getElementById(idTabla);
    const emptyEl = document.getElementById(idEmpty);
    const conteoEl = document.getElementById(idConteo);

    t.innerHTML = "";

    if (data.length === 0) {
        emptyEl.style.display = "block";
        conteoEl.textContent = "";
    } else {
        emptyEl.style.display = "none";
        conteoEl.textContent = `(${data.length})`;
        t.innerHTML = "<tr>" + headers.map(h => `<th>${h}</th>`).join("") + "</tr>";
        data.forEach(r => {
            t.innerHTML += "<tr>" + r.map(c => `<td>${c}</td>`).join("") + "</tr>";
        });
    }
}

/* ===============================
   EXPORTAR EXCEL
=============================== */
document.getElementById("btnExportar").addEventListener("click", () => {
    const wb = XLSX.utils.book_new();

    const tablas = [
        { id: "tablaTiene",    nombre: "Claves_Tiene"     },
        { id: "tablaInactivas",nombre: "Claves_Inactivas" },
        { id: "tablaFaltan",   nombre: "Claves_Faltan"    },
        { id: "tablaErrores",  nombre: "Oficinas_Invalidas"},
        { id: "tablaBajas",    nombre: "Claves_Inhabilitar"}
    ];

    tablas.forEach(tabla => {
        const table = document.getElementById(tabla.id);

        // CORRECCIÓN: <= 1 porque la fila 0 siempre es el encabezado
        if (!table || table.rows.length <= 1) {
            const ws = XLSX.utils.aoa_to_sheet([["No hay datos"]]);
            XLSX.utils.book_append_sheet(wb, ws, tabla.nombre);
        } else {
            const ws = XLSX.utils.table_to_sheet(table);
            XLSX.utils.book_append_sheet(wb, ws, tabla.nombre);
        }
    });

    const empleado = document.getElementById("empleado").value.trim() || "empleado";
    XLSX.writeFile(wb, `Claves_${empleado}.xlsx`);
});

/* ===============================
   INICIALIZACIÓN
=============================== */
document.addEventListener("DOMContentLoaded", () => {
    cargarXLSXAutomatico();
});
