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

        const sep = lines[0].includes(";") ? ";" : ",";
        const headers = lines[0].split(sep).map(h => h.toLowerCase());

        const idxClave = headers.findIndex(h => h.includes("clave"));
        const idxActivo = headers.findIndex(h => h.includes("activo"));

        personas = [];

        for (let i = 1; i < lines.length; i++) {
            const cols = lines[i].split(sep);
            const clave = (cols[idxClave] || "").trim().toUpperCase();
            if (!clave) continue;

            personas.push({
                clave,
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

    const mapOficinas = {};
    oficinas.forEach(o => mapOficinas[o.CLAVE] = o);

    const nomUsuarioActivas = new Set();
    const nomUsuarioInactivas = new Set();

    personas.forEach(p => {
        const nom = p.nomenclatura;
        p.activo ? nomUsuarioActivas.add(nom) : nomUsuarioInactivas.add(nom);
    });

    const tiene = [], faltan = [], inactivas = [], bajas = [], errores = [];
    const debeTener = new Set();

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

    nomUsuarioActivas.forEach(nom => {
        if (!debeTener.has(nom)) {
            const of = oficinas.find(o => o.NOMENCLATURA === nom);
            if (of) bajas.push([nom, of.NOMBRE, of.DIRECCION, "INHABILITAR"]);
        }
    });

    pintar("tablaTiene", ["CLAVE","NOMBRE","NOM","DIR","EST"], tiene);
    pintar("tablaInactivas", ["CLAVE","NOMBRE","NOM","DIR","EST"], inactivas);
    pintar("tablaFaltan", ["CLAVE","NOMBRE","NOM","DIR","EST"], faltan);
    pintar("tablaBajas", ["NOM","NOMBRE","DIR","EST"], bajas);
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
