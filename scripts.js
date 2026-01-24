let oficinas = [];
let personas = [];

/* ===============================
   BOTÓN PROCESAR
=============================== */
document.getElementById("btnProcesar").addEventListener("click", procesar);

/* ===============================
   CARGA XLSX AUTOMÁTICA
=============================== */
function cargarXLSXAutomatico() {
    fetch("OFICINAS NOMENCLATURAS.xlsx")
        .then(res => {
            if (!res.ok) throw new Error("No se pudo cargar OFICINAS NOMENCLATURAS.xlsx");
            return res.arrayBuffer();
        })
        .then(data => {
            const wb = XLSX.read(data, { type: "array" });
            const hoja = wb.Sheets[wb.SheetNames[0]];

            oficinas = XLSX.utils.sheet_to_json(hoja, { defval: "" })
                .map(o => {
                    let claveOf = String(o.CLAVE).trim();

                    // Normalizar claves numéricas
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

            console.log("📘 Catálogo de oficinas cargado:", oficinas.length);
        })
        .catch(err => {
            alert("❌ Error cargando el catálogo de oficinas");
            console.error(err);
        });
}


/* ===============================
   CARGA CSV
=============================== */
document.getElementById("csvPersonas").addEventListener("change", e => {
    const reader = new FileReader();

    reader.onload = ev => {
        const lines = ev.target.result.replace(/\r/g, "")
            .split("\n")
            .filter(l => l.trim());

        let sep = ",";
        if (lines[0].includes("\t")) sep = "\t";
        else if (lines[0].includes(";")) sep = ";";
        const headers = lines[0].split(sep).map(h => h.toLowerCase());

        const idxClave = headers.findIndex(h => h.includes("clave"));
        const idxActivo = headers.findIndex(h => h.includes("activo"));
        const idxNombre = headers.findIndex(h =>
        h.includes("nombre") || h.includes("descripcion")
        );

        personas = [];

        for (let i = 1; i < lines.length; i++) {
            const cols = lines[i].split(sep);
            const clave = (cols[idxClave] || "").trim().toUpperCase();
            if (!clave) continue;

            personas.push({
            clave,
            claveCompleta: clave,
            nombreReal: idxNombre >= 0
                ? (cols[idxNombre] || "").trim()
                : (cols[1] || "").trim(), // fallback típico columna 2
            nomenclatura: clave.substring(0, 3),
            activo: (cols[idxActivo] || "").toLowerCase() === "true"
        });
        }
    };

    reader.readAsText(e.target.files[0], "UTF-8");
});

/* ===============================
   PROCESAR
=============================== */

function procesar() {
    if (!personas.length || !oficinas.length) {
        alert("Carga primero el CSV y el XLSX");
        return;
    }

    const empleado = document.getElementById("empleado").value.padStart(6, "0");
    const oficinasInput = document.getElementById("oficinasInput").value
        .split("\n").map(o => o.trim()).filter(Boolean);
        const oficinasQuitarRaw = document.getElementById("oficinasQuitarInput").value
    .split("\n")
    .map(o => o.trim().toUpperCase())
    .filter(Boolean);

let modoQuitar = null; // "NUM" | "NOM"

// Detectar modo por el primer valor
if (oficinasQuitarRaw.length > 0) {
    const primero = oficinasQuitarRaw[0];

    if (/^\d+$/.test(primero)) {
        modoQuitar = "NUM";
    } else if (/^[A-Z]{3}$/.test(primero)) {
        modoQuitar = "NOM";
    } else {
        alert("Formato inválido en oficinas a quitar");
        return;
    }
}

    const mapOficinas = {};
    oficinas.forEach(o => mapOficinas[o.CLAVE] = o);

        const nomUsuarioActivas = new Set();
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

// Normalizar oficinas a quitar
oficinasQuitarRaw.forEach(o => {

    // ===== MODO NUMÉRICO =====
    if (modoQuitar === "NUM") {
        if (!/^\d+$/.test(o)) {
            alert("Todas las oficinas a quitar deben ser números");
            throw new Error("Formato mixto");
        }

        const clave = o.padStart(4, "0");
        const of = mapOficinas[clave];

        if (!of) {
            alert(`La oficina ${clave} no existe en el XLSX`);
            throw new Error("Oficina inválida");
        }

        // Guardamos NOMENCLATURA
        oficinasQuitarSet.add(of.NOMENCLATURA);
    }

    // ===== MODO NOMENCLATURA =====
    if (modoQuitar === "NOM") {
        if (!/^[A-Z]{3}$/.test(o)) {
            alert("Todas las oficinas a quitar deben ser nomenclatura (3 letras)");
            throw new Error("Formato mixto");
        }

        oficinasQuitarSet.add(o);
    }
});

    oficinasInput.forEach(cod => {
        if (/^\d+$/.test(cod)) cod = cod.padStart(4, "0");

        const of = mapOficinas[cod];
        if (!of) {
            errores.push([cod, "Oficina no existe en XLSX"]);
            return;
        }

        const claveFinal = of.NOMENCLATURA + "0" + empleado;
        debeTener.add(of.NOMENCLATURA);

        if (nomUsuarioActivas.has(of.NOMENCLATURA)) {
            tiene.push([claveFinal, of.NOMBRE, of.NOMENCLATURA, of.DIRECCION, "ACTIVA"]);
        } else if (nomUsuarioInactivas.has(of.NOMENCLATURA)) {
            inactivas.push([claveFinal, of.NOMBRE, of.NOMENCLATURA, of.DIRECCION, "REACTIVAR"]);
        } else {
            faltan.push([claveFinal, of.NOMBRE, of.NOMENCLATURA, of.DIRECCION, "AGREGAR"]);
        }
    });

        const modoQuitarEspecifico = oficinasQuitarSet.size > 0;

            nomUsuarioActivas.forEach(nom => {

            // Nunca inhabilitar algo que se pidió agregar
            if (debeTener.has(nom)) return;

            // Si hay lista explícita para quitar
            if (modoQuitarEspecifico && !oficinasQuitarSet.has(nom)) {
                return;
            }

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


    pintar("tablaTiene", ["CLAVE","NOMBRE","NOM","DIR","ACCIÓN"], tiene);
    pintar("tablaInactivas", ["CLAVE","NOMBRE","NOM","DIR","ACCIÓN"], inactivas);
    pintar("tablaFaltan", ["CLAVE","NOMBRE","NOM","DIR","ACCIÓN"], faltan);
    pintar("tablaBajas", ["CLAVE CSV", "DESCRIPCIÓN CSV", "NOM", "ESTADO CSV", "ACCIÓN"],
    bajas
    );
    pintar("tablaErrores", ["OFICINA","ERROR"], errores);
}

/* ===============================
   PINTAR TABLAS
=============================== */
function pintar(id, headers, data) {
    const t = document.getElementById(id);
    t.innerHTML = "<tr>" + headers.map(h => `<th>${h}</th>`).join("") + "</tr>";
    data.forEach(r => {
        t.innerHTML += "<tr>" + r.map(c => `<td>${c}</td>`).join("") + "</tr>";
    });
}
document.addEventListener("DOMContentLoaded", () => {
    cargarXLSXAutomatico();
});

document.getElementById('btnExportar').addEventListener('click', () => {
    // Crear un nuevo libro de Excel
    const wb = XLSX.utils.book_new();

    // Array con id de tablas y nombre de cada sheet
    const tablas = [
        { id: 'tablaTiene', nombre: 'Claves_Tiene' },
        { id: 'tablaInactivas', nombre: 'Claves_Inactivas' },
        { id: 'tablaFaltan', nombre: 'Claves_Faltan' },
        { id: 'tablaErrores', nombre: 'Oficinas_Invalidas' },
        { id: 'tablaBajas', nombre: 'Claves_Inhabilitar' }
    ];

    tablas.forEach(tabla => {
        const table = document.getElementById(tabla.id);

        // Si la tabla está vacía, poner un mensaje
        if (table.rows.length === 0) {
            const ws_data = [["No hay datos"]];
            const ws = XLSX.utils.aoa_to_sheet(ws_data);
            XLSX.utils.book_append_sheet(wb, ws, tabla.nombre);
        } else {
            // Convertir tabla HTML a hoja Excel
            const ws = XLSX.utils.table_to_sheet(table);
            XLSX.utils.book_append_sheet(wb, ws, tabla.nombre);
        }
    });

    // Descargar el archivo
    XLSX.writeFile(wb, 'Claves_Oficinas.xlsx');
});
