Chart.register(ChartDataLabels);

let db = { motores: [], limites: {}, inspecciones: [], historialExcel: [] };
let chartDesgasteInstance = null;
let chartPredictivoInstance = null;
let miniChartsInstances = {};
let graficosReporteInstances = {};
let inspeccionesMotorActual = [];
let motorIndexSeleccionado = 0;
let listaPronosticosGlobal = [];
let datosReporteParaExportar = [];

const WEBHOOK_OUTLOOK_URL = "https://default6d25241940fe4964bb31e8aaa96d97.66.environment.api.powerplatform.com:443/powerautomate/automations/direct/cu/06/workflows/05dd39f257ea471fbb7b21e0d5eb5c2f/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=F11KYR0Q-qZ3I5CrU9iKeWoTDMgB1ed7_PPoPEzu_FI";
const CORREOS_DESTINO = "hernan.diaz@unacem.pe,braulio.siancas@unacem.pe";
const URL_WEBHOOK_CONSULTA_EXCEL = "https://default6d25241940fe4964bb31e8aaa96d97.66.environment.api.powerplatform.com:443/powerautomate/automations/direct/cu/28/workflows/df4f4ad8529d46e9b4e125d5b888bf86/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=I_0c8OtCiw-zAl-WYH5rR3EuPWYi0jrmqrn1YhvrpNQ";

// NORMALIZACIÓN DE TEXTO
function normalizarTexto(texto) {
    return String(texto || '')
        .trim()
        .toLowerCase()
        .normalize("NFD")
        .replace(/[\u0300-\u036f]/g, "")
        .replace(/[^a-z0-9]/g, "");
}

async function sincronizarBaseDesdePowerAutomate() {
    const btnSync = document.getElementById('btn-sync-cloud');
    const iconoSync = btnSync ? btnSync.querySelector('i') : null;

    if (!URL_WEBHOOK_CONSULTA_EXCEL || URL_WEBHOOK_CONSULTA_EXCEL.includes("PEGA_AQUI")) {
        console.warn("Falta colocar la URL de Power Automate.");
        return;
    }

    try {
        if (iconoSync) iconoSync.classList.add('fa-spin');
        console.log("Conectando con Power Automate...");

        const separador = URL_WEBHOOK_CONSULTA_EXCEL.includes('?') ? '&' : '?';
        const urlPeticion = `${URL_WEBHOOK_CONSULTA_EXCEL}${separador}t=${new Date().getTime()}`;

        const respuesta = await fetch(urlPeticion, { method: 'GET' });

        if (!respuesta.ok) {
            throw new Error(`Error en Power Automate: ${respuesta.status} - ${respuesta.statusText}`);
        }

        const jsonRespuesta = await respuesta.json();

        // 1. Decodificar BASE DE DATOS.xlsm
        if (jsonRespuesta.baseDatos) {
            const binaryBase = atob(jsonRespuesta.baseDatos);
            const bytesBase = new Uint8Array(binaryBase.length);
            for (let i = 0; i < binaryBase.length; i++) bytesBase[i] = binaryBase.charCodeAt(i);
            const wbBase = XLSX.read(bytesBase, { type: 'array' });

            // A) Motores
            const sheetMotoresName = wbBase.SheetNames.find(n => normalizarTexto(n).includes('motor')) || wbBase.SheetNames[0];
            db.motores = extraerMatrizHoja(wbBase.Sheets[sheetMotoresName]);

            // B) Límites
            const sheetLimitesName = wbBase.SheetNames.find(n => {
                const norm = normalizarTexto(n);
                return norm === 'limites' || norm.includes('limite');
            });

            db.limites = {};
            if (sheetLimitesName) {
                const rawLimites = extraerMatrizHoja(wbBase.Sheets[sheetLimitesName]);
                rawLimites.forEach(l => {
                    const cod = l['Ubicacion tecnica'] || l['Ubicación técnica'] || l['UbicacionTecnica'] || l['Codigo de equipo'] || l['Código de equipo'] || l['Column1'];
                    if (cod) {
                        const keyNorm = normalizarTexto(cod);
                        const medNueva = parseFloat(l['Medida Nueva'] || l['MedidaNueva'] || l['ValorNuevo']);
                        const limCambio = parseFloat(l['Limite Cambio'] || l['LimiteCambio'] || l['Limite de cambio'] || l['Limite']);

                        db.limites[keyNorm] = {
                            _keyOriginal: cod,
                            nuevo: !isNaN(medNueva) ? medNueva : 40,
                            limite: !isNaN(limCambio) ? limCambio : (!isNaN(medNueva) ? medNueva / 2 : 20)
                        };
                    }
                });
            }

            // C) Inspecciones
            db.inspecciones = [];
            wbBase.SheetNames.forEach(sName => {
                const norm = normalizarTexto(sName);
                if (norm.includes('power') || norm.includes('bbti') || norm.includes('sisac')) {
                    const filasHoja = extraerMatrizHoja(wbBase.Sheets[sName]);
                    db.inspecciones = db.inspecciones.concat(filasHoja);
                }
            });
            db.inspecciones.forEach(i => i._dateObj = parsearFecha(i['Fecha']));
        }

        // 2. Historial
        if (jsonRespuesta.historial) {
            const binaryHist = atob(jsonRespuesta.historial);
            const bytesHist = new Uint8Array(binaryHist.length);
            for (let i = 0; i < binaryHist.length; i++) bytesHist[i] = binaryHist.charCodeAt(i);
            const wbHist = XLSX.read(bytesHist, { type: 'array' });
            db.historialExcel = extraerMatrizHoja(wbHist.Sheets[wbHist.SheetNames[0]]);
        }

        const hoy = new Date();
        const fechaStr = hoy.toLocaleDateString('es-PE') + ' ' + hoy.toLocaleTimeString('es-PE', { hour: '2-digit', minute: '2-digit' });

        localStorage.removeItem('unacem_db_consolidado');
        localStorage.setItem('unacem_db_consolidado', JSON.stringify(db));
        localStorage.setItem('unacem_db_fecha', fechaStr);

        const lblFecha = document.getElementById('lbl-fecha-actualizacion');
        if (lblFecha) lblFecha.innerText = fechaStr;

        poblarSelectorMotoresSidebar();
        establecerRangoFechasPredeterminado();

        const vistaActiva = document.querySelector('.nav-item.active')?.id.replace('nav-', '') || 'motores';
        cambiarVista(vistaActiva);

    } catch (err) {
        console.warn("Error al sincronizar datos:", err);
    } finally {
        if (iconoSync) iconoSync.classList.remove('fa-spin');
    }
}

document.addEventListener('DOMContentLoaded', function() {
    const logoEl = document.getElementById('img-logo');
    if (logoEl) logoEl.onerror = function() { logoEl.src = '../logo.png'; };

    cargarDatosGuardados();
    cargarConfiguracionEnvio();
    sincronizarBaseDesdePowerAutomate();
    cambiarVista('motores');
    verificarEnvioProgramadoOutlook();
});

function cambiarVista(vista) {
    document.querySelectorAll('.nav-item').forEach(item => item.classList.remove('active'));
    document.getElementById(`nav-${vista}`)?.classList.add('active');

    ['motores', 'inicio', 'predictivo', 'reportes', 'historial', 'configuracion'].forEach(v => {
        const el = document.getElementById(`vista-${v}`);
        if (el) el.style.display = (v === vista) ? 'flex' : 'none';
    });

    if (vista === 'motores') {
        renderizarCatalogoMotores();
    } else if (vista === 'inicio') {
        if (db.motores && db.motores.length > 0) seleccionarMotorPorIndice(motorIndexSeleccionado);
    } else if (vista === 'predictivo') {
        requestAnimationFrame(() => { renderizarModuloPredictivo(); });
    } else if (vista === 'reportes') {
        renderizarReporteConsolidado();
    } else if (vista === 'historial') {
        renderizarHistorialReportesUI();
    }
}

function evaluarTemperaturaIndustrial(tempVal) {
    const t = parseFloat(tempVal);
    if (isNaN(t) || t < 90) return { estado: 'NORMAL', esAlerta: false, recomendacion: 'Operación normal, continuar monitoreo.', color: '#16a34a' };
    if (t > 100) return { estado: 'CRÍTICO', esAlerta: true, recomendacion: 'Temperatura superior a 100°C. Intervención prioritaria.', color: '#dc2626' };
    return { estado: 'ALARMA', esAlerta: true, recomendacion: 'Temperatura entre 90°C y 100°C. Programar inspección.', color: '#ea580c' };
}

function obtenerNombreCarbonConFase(ubTecnica, cKey) {
    const cod = normalizarTexto(ubTecnica);
    const num = parseInt(String(cKey).replace(/\D/g, ''), 10);
    if (isNaN(num)) return cKey;

    if (cod.includes('3375')) {
        if (num >= 1 && num <= 3) return `C${num}-K`;
        if (num >= 4 && num <= 6) return `C${num}-L`;
        if (num >= 7 && num <= 9) return `C${num}-M`;
    }
    if (cod.includes('3335')) {
        if (num >= 1 && num <= 4) return `C${num}-K`;
        if (num >= 5 && num <= 8) return `C${num}-L`;
        if (num >= 9 && num <= 12) return `C${num}-M`;
    }
    return `C${num}`;
}

function parsearFecha(valor) {
    if (!valor) return null;
    if (valor instanceof Date) return isNaN(valor.getTime()) ? null : valor;
    if (typeof valor === 'number') {
        const dateObj = XLSX.SSF.parse_date_code(valor);
        if (dateObj) return new Date(dateObj.y, dateObj.m - 1, dateObj.d);
    }
    if (typeof valor === 'string') {
        const str = valor.trim().split(' ')[0];
        const parts = str.split(/[\/\-]/);
        if (parts.length === 3) {
            const p0 = parseInt(parts[0], 10);
            const p1 = parseInt(parts[1], 10);
            const p2 = parseInt(parts[2], 10);
            if (p0 > 1000) return new Date(p0, p1 - 1, p2);
            if (p2 > 1000) return new Date(p2, p1 - 1, p0);
        }
    }
    const d = new Date(valor);
    return isNaN(d.getTime()) ? null : d;
}

function formatearFechaLegible(dateObj) {
    if (!dateObj || isNaN(dateObj.getTime())) return 'S/F';
    const d = String(dateObj.getDate()).padStart(2, '0');
    const m = String(dateObj.getMonth() + 1).padStart(2, '0');
    return `${d}/${m}/${dateObj.getFullYear()}`;
}

function cargarDatosGuardados() {
    const dbGuardada = localStorage.getItem('unacem_db_consolidado');
    const fechaGuardada = localStorage.getItem('unacem_db_fecha');

    if (dbGuardada) {
        try {
            db = JSON.parse(dbGuardada);
            if (!db.limites) db.limites = {};
            if (db.inspecciones && Array.isArray(db.inspecciones)) {
                db.inspecciones.forEach(i => { i._dateObj = parsearFecha(i['Fecha']); });
            }
            document.getElementById('lbl-fecha-actualizacion').innerText = fechaGuardada || '--/--/----';
            poblarSelectorMotoresSidebar();
            establecerRangoFechasPredeterminado();
        } catch(e) {
            console.error("Error cargando base local:", e);
        }
    }
}

function establecerRangoFechasPredeterminado() {
    const fIniInput = document.getElementById('fecha-inicio');
    const fFinInput = document.getElementById('fecha-fin');

    if (fIniInput && !fIniInput.value) fIniInput.value = "2026-01-01";

    if (fFinInput && !fFinInput.value && db.inspecciones && db.inspecciones.length > 0) {
        const fechasValidas = db.inspecciones.filter(i => i._dateObj).map(i => i._dateObj.getTime());
        if (fechasValidas.length > 0) {
            const maxF = new Date(Math.max(...fechasValidas));
            fFinInput.value = maxF.toISOString().split('T')[0];
        }
    }
}

// POBLAR EL SELECTOR DEL SIDEBAR MOSTRANDO EL CONTRATISTA REAL DE LA HOJA MOTORES
function poblarSelectorMotoresSidebar() {
    const select = document.getElementById('select-motor-filtro');
    const selectContratista = document.getElementById('select-contratista-filtro');
    if (!select) return;
    
    const contratistaSel = selectContratista ? selectContratista.value : "TODOS";
    const valorSeleccionadoPrevio = select.value;

    select.innerHTML = '<option value="TODOS">-- Todos los Motores --</option>';

    if (!db.motores || db.motores.length === 0) return;

    db.motores.forEach((m, idx) => {
        const contratistaMotor = (m['Contratista'] || 'UNACEM').trim().toUpperCase();

        if (contratistaSel === "TODOS" || contratistaMotor === contratistaSel) {
            const ubTecnica = m['Ubicacion tecnica'] || m['Ubicación técnica'] || m['UbicacionTecnica'] || m['Codigo de equipo'] || '';
            const nombreMotor = m['Motor'] || m['NombreMotor'] || ubTecnica;

            const opt = document.createElement('option');
            opt.value = idx;
            opt.innerText = `${ubTecnica ? ubTecnica + ' - ' : ''}${nombreMotor} (${contratistaMotor})`;
            select.appendChild(opt);
        }
    });

    if (valorSeleccionadoPrevio !== "" && select.querySelector(`option[value="${valorSeleccionadoPrevio}"]`)) {
        select.value = valorSeleccionadoPrevio;
    } else {
        select.value = "TODOS";
    }
}

function filtrarPorContratista(contratistaVal) {
    poblarSelectorMotoresSidebar();
    
    const vistaPredictivo = document.getElementById('vista-predictivo');
    if (vistaPredictivo && vistaPredictivo.style.display !== 'none') {
        renderizarModuloPredictivo();
    }
    
    const vistaReportes = document.getElementById('vista-reportes');
    if (vistaReportes && vistaReportes.style.display !== 'none') {
        renderizarReporteConsolidado();
    }

    const vistaMotores = document.getElementById('vista-motores');
    if (vistaMotores && vistaMotores.style.display !== 'none') {
        renderizarCatalogoMotores();
    }
}

function seleccionarMotorDesdeFiltro(indexStr) {
    if (indexStr === "TODOS" || indexStr === "") {
        const vistaPredictivo = document.getElementById('vista-predictivo');
        if (vistaPredictivo && vistaPredictivo.style.display !== 'none') {
            renderizarModuloPredictivo();
        }
        return;
    }
    seleccionarMotorPorIndice(parseInt(indexStr, 10));
}

function extraerMatrizHoja(worksheet) {
    const matrix = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: "" });
    if (!matrix || matrix.length === 0) return [];

    let headerRowIndex = 0;
    for (let i = 0; i < matrix.length; i++) {
        const filaStr = matrix[i].map(c => normalizarTexto(c)).join(' ');
        if (filaStr.includes('ubicacion') || filaStr.includes('motor') || filaStr.includes('fecha') || filaStr.includes('link')) {
            headerRowIndex = i;
            break;
        }
    }

    const headers = matrix[headerRowIndex].map(h => String(h).trim());
    const result = [];

    for (let i = headerRowIndex + 1; i < matrix.length; i++) {
        const row = matrix[i];
        if (!row || row.every(val => val === "" || val === null)) continue;
        const obj = {};
        headers.forEach((h, colIndex) => {
            if (h) obj[h] = row[colIndex] !== undefined ? row[colIndex] : "";
        });
        result.push(obj);
    }
    return result;
}

function cargarImagenConFallbacks(imgEl, ubTecnica, defaultSrc, placeholderEl) {
    if (!ubTecnica) {
        if (imgEl) imgEl.src = defaultSrc;
        return;
    }
    const rutas = [
        `./${ubTecnica}.png`, `./${ubTecnica}.jpg`,
        `../${ubTecnica}.png`, `../${ubTecnica}.jpg`,
        `./motores/${ubTecnica}.png`, `./motores/${ubTecnica}.jpg`
    ];
    let index = 0;
    function intentarSiguiente() {
        if (index < rutas.length) {
            const rutaActual = rutas[index++];
            if (imgEl) {
                imgEl.onerror = intentarSiguiente;
                imgEl.src = rutaActual;
                imgEl.style.display = 'block';
            }
            if (placeholderEl) placeholderEl.style.display = 'none';
        } else {
            if (imgEl) {
                imgEl.onerror = null;
                imgEl.style.display = 'none';
            }
            if (placeholderEl) {
                placeholderEl.style.display = 'block';
                placeholderEl.innerText = `Sin Foto\n(${ubTecnica}.png)`;
            }
        }
    }
    intentarSiguiente();
}

function abrirModalZoomMotor() {
    const imgActual = document.getElementById('img-motor');
    const modal = document.getElementById('modal-zoom-motor');
    const imgModal = document.getElementById('img-modal-motor');
    const titulo = document.getElementById('modal-motor-titulo');

    if (imgActual && imgActual.src && imgActual.style.display !== 'none') {
        imgModal.src = imgActual.src;
        titulo.innerText = document.getElementById('val-motor').innerText + " (" + document.getElementById('val-ubicacion-tecnica').innerText + ")";
        modal.style.display = 'flex';
    }
}

function cerrarModalZoomMotor() {
    const modal = document.getElementById('modal-zoom-motor');
    if (modal) modal.style.display = 'none';
}

function renderizarCatalogoMotores() {
    const grupos = {
        'HORNO 3': document.getElementById('grid-horno-3'),
        'CRUDO 1': document.getElementById('grid-crudo-1'),
        'CRUDO 2': document.getElementById('grid-crudo-2'),
        'HORNO 2': document.getElementById('grid-horno-2')
    };

    const selectContratista = document.getElementById('select-contratista-filtro');
    const contratistaSel = selectContratista ? selectContratista.value : "TODOS";

    for (let key in grupos) { if (grupos[key]) grupos[key].innerHTML = ''; }

    db.motores.forEach((m, idx) => {
        const contratistaMotor = (m['Contratista'] || 'UNACEM').trim().toUpperCase();

        if (contratistaSel === "TODOS" || contratistaMotor === contratistaSel) {
            const ubPlanta = (m['Ubicacion en planta'] || m['Ubicación en planta'] || 'HORNO 3').trim().toUpperCase();
            const ubTecnica = m['Ubicacion tecnica'] || m['Ubicación técnica'] || '';
            const nombreMotor = m['Motor'] || ubTecnica;
            const codEquipo = m['Codigo de equipo'] || m['Código de equipo'] || '';

            const targetGrid = grupos[ubPlanta] || grupos['HORNO 3'];
            const isSelected = (idx === motorIndexSeleccionado) ? 'selected' : '';

            const card = document.createElement('div');
            card.className = `motor-select-card ${isSelected}`;
            card.onclick = function() {
                seleccionarMotorPorIndice(idx);
                cambiarVista('inicio');
            };

            const imgEl = document.createElement('img');
            cargarImagenConFallbacks(imgEl, ubTecnica, 'Motor.png', null);

            const infoDiv = document.createElement('div');
            infoDiv.className = 'motor-select-info';
            infoDiv.innerHTML = `
                <h4>${nombreMotor}</h4>
                <p><strong>Cód:</strong> ${codEquipo} | <strong>Ub:</strong> ${ubTecnica}</p>
                <span class="badge-contratista">${contratistaMotor}</span>
            `;

            card.appendChild(imgEl);
            card.appendChild(infoDiv);
            if (targetGrid) targetGrid.appendChild(card);
        }
    });
}

function seleccionarMotorPorIndice(index) {
    if (index === "" || !db.motores[index]) return;

    motorIndexSeleccionado = index;
    const selectSidebar = document.getElementById('select-motor-filtro');
    if (selectSidebar) selectSidebar.value = index;

    const m = db.motores[index];
    const ubTecnica = m['Ubicacion tecnica'] || m['Ubicación técnica'] || '';
    const nombreMotor = m['Motor'] || ubTecnica;

    document.getElementById('val-motor').innerText = nombreMotor;
    document.getElementById('val-codigo').innerText = m['Codigo de equipo'] || m['Código de equipo'] || '--';
    document.getElementById('val-ubicacion-tecnica').innerText = ubTecnica || '--';
    document.getElementById('val-ubicacion-planta').innerText = m['Ubicacion en planta'] || m['Ubicación en planta'] || '--';
    document.getElementById('val-contratista').innerText = m['Contratista'] || '--';

    document.getElementById('titulo-tabla-carbones').innerText = `${nombreMotor} - Estado Carbones`;
    const tituloParams = document.getElementById('titulo-parametros-relevantes');
    if (tituloParams) tituloParams.innerText = `Parámetros Relevantes - ${nombreMotor} (${ubTecnica})`;

    const imgElement = document.getElementById('img-motor');
    const imgCenter = document.getElementById('img-motor-center');
    const placeholder = document.getElementById('img-placeholder');

    cargarImagenConFallbacks(imgElement, ubTecnica, 'Motor.png', placeholder);
    imgCenter.src = "Motor.png";
    imgCenter.onerror = function() { imgCenter.src = "logo.png"; };

    const ubTecnicaNorm = normalizarTexto(ubTecnica);
    const codEquipoNorm = normalizarTexto(m['Codigo de equipo'] || m['Código de equipo']);

    inspeccionesMotorActual = db.inspecciones.filter(insp => {
        const codInsp = normalizarTexto(insp['Ubicacion tecnica'] || insp['Ubicación técnica'] || insp['UbicacionTecnica'] || insp['Codigo de equipo']);
        return (codInsp === ubTecnicaNorm || codInsp === codEquipoNorm) && codInsp !== '';
    });

    inspeccionesMotorActual.forEach(insp => { if (!insp._dateObj) insp._dateObj = parsearFecha(insp['Fecha']); });
    inspeccionesMotorActual.sort((a, b) => (a._dateObj ? a._dateObj.getTime() : 0) - (b._dateObj ? b._dateObj.getTime() : 0));

    poblarFiltroCarbones();
    aplicarFiltroRangoFechas();
}

function poblarFiltroCarbones() {
    const selectCarbon = document.getElementById('select-carbon-filtro');
    if (!selectCarbon) return;
    selectCarbon.innerHTML = '<option value="TODOS">Todos los carbones</option>';

    if (inspeccionesMotorActual.length === 0) return;

    const m = db.motores[motorIndexSeleccionado];
    const ubTecnica = m ? (m['Ubicacion tecnica'] || m['Ubicación técnica'] || '') : '';

    const carbonesEncontrados = new Set();
    inspeccionesMotorActual.forEach(insp => {
        for (let i = 1; i <= 24; i++) {
            const valRaw = obtenerValorCarbonFlex(insp, i);
            if (valRaw !== null) carbonesEncontrados.add(i);
        }
    });

    carbonesEncontrados.forEach(num => {
        const opt = document.createElement('option');
        const cKeyOriginal = `C${num}`;
        const cKeyNombreFase = obtenerNombreCarbonConFase(ubTecnica, cKeyOriginal);
        opt.value = cKeyOriginal;
        opt.innerText = cKeyNombreFase;
        selectCarbon.appendChild(opt);
    });
}

function aplicarFiltroRangoFechas() {
    const fInicioVal = document.getElementById('fecha-inicio').value;
    const fFinVal = document.getElementById('fecha-fin').value;

    const fInicio = fInicioVal ? new Date(fInicioVal + 'T00:00:00') : null;
    const fFin = fFinVal ? new Date(fFinVal + 'T23:59:59') : null;

    const inspeccionesFiltradas = inspeccionesMotorActual.filter(insp => {
        if (!insp._dateObj) return true;
        if (fInicio && insp._dateObj < fInicio) return false;
        if (fFin && insp._dateObj > fFin) return false;
        return true;
    });

    const m = db.motores[motorIndexSeleccionado];
    const ubTecnica = m ? (m['Ubicacion tecnica'] || m['Ubicación técnica'] || '') : '';

    renderizarVistaInspeccionesRango(inspeccionesFiltradas, ubTecnica);

    const vistaReportes = document.getElementById('vista-reportes');
    if (vistaReportes && vistaReportes.style.display !== 'none') renderizarReporteConsolidado();

    const vistaPredictivo = document.getElementById('vista-predictivo');
    if (vistaPredictivo && vistaPredictivo.style.display !== 'none') renderizarModuloPredictivo();
}

function obtenerValorCampoFlex(obj, posiblesNombres) {
    if (!obj) return '--';
    for (let nombre of posiblesNombres) {
        const nombreNorm = normalizarTexto(nombre);
        for (let key in obj) {
            if (normalizarTexto(key) === nombreNorm) {
                if (obj[key] !== undefined && obj[key] !== "" && obj[key] !== null) return obj[key];
            }
        }
    }
    return '--';
}

function obtenerValorCarbonFlex(insp, indiceNum) {
    const clavesPosibles = [`C${indiceNum}`, `C0${indiceNum}`, `C-${indiceNum}`, `Carbon${indiceNum}`, `Carbon ${indiceNum}`, `${indiceNum}`];
    for (let k of clavesPosibles) {
        if (insp[k] !== undefined && insp[k] !== "" && insp[k] !== null && !isNaN(parseFloat(insp[k]))) {
            return parseFloat(insp[k]);
        }
    }
    return null;
}

function obtenerLimiteEspecificoMotor(ubTecnica, codEquipo) {
    const keyUbNorm = normalizarTexto(ubTecnica);
    const keyCodNorm = normalizarTexto(codEquipo);

    if (db.limites[keyUbNorm]) return db.limites[keyUbNorm];
    if (db.limites[keyCodNorm]) return db.limites[keyCodNorm];

    const limiteEncontrado = Object.values(db.limites).find(l => {
        if (!l || !l._keyOriginal) return false;
        const normOrig = normalizarTexto(l._keyOriginal);
        return normOrig === keyUbNorm || normOrig === keyCodNorm;
    });

    return limiteEncontrado || { nuevo: 40, limite: 20 };
}

function renderizarVistaInspeccionesRango(inspecciones, ubTecnica) {
    const m = db.motores[motorIndexSeleccionado] || {};
    const codEquipo = m['Codigo de equipo'] || m['Código de equipo'] || '';

    const limites = obtenerLimiteEspecificoMotor(ubTecnica, codEquipo);

    if (inspecciones.length === 0) {
        limpiarTablasYGraficos();
        return;
    }

    const ultimaInsp = inspecciones[inspecciones.length - 1];
    const penultimaInsp = inspecciones.length > 1 ? inspecciones[inspecciones.length - 2] : null;

    const corriente = obtenerValorCampoFlex(ultimaInsp, ['Corriente', 'Amp']);
    const tempK = obtenerValorCampoFlex(ultimaInsp, ['K', 'Temp K', 'Temperatura K']);
    const tempM = obtenerValorCampoFlex(ultimaInsp, ['M', 'Temp M', 'Temperatura M']);
    const tempL = obtenerValorCampoFlex(ultimaInsp, ['L', 'Temp L', 'Temperatura L']);
    const tempCarcaza = obtenerValorCampoFlex(ultimaInsp, ['Carcaza', 'Carcasa']);
    const tempLLibre = obtenerValorCampoFlex(ultimaInsp, ['L. Libre', 'L.Libre']);
    const tempLAcople = obtenerValorCampoFlex(ultimaInsp, ['L. Acople', 'L.Acople']);

    const evalK = evaluarTemperaturaIndustrial(tempK);
    const evalM = evaluarTemperaturaIndustrial(tempM);
    const evalL = evaluarTemperaturaIndustrial(tempL);
    const evalLL = evaluarTemperaturaIndustrial(tempLLibre);
    const evalLA = evaluarTemperaturaIndustrial(tempLAcople);

    document.getElementById('val-corriente').innerText = corriente !== '--' ? corriente + ' A' : '-- A';
    document.getElementById('val-temp-k').innerHTML = `${tempK !== '--' ? tempK + ' °C' : '-- °C'} <span class="badge-status" style="background:${evalK.color}; font-size:8px;">${evalK.estado}</span>`;
    document.getElementById('val-temp-m').innerHTML = `${tempM !== '--' ? tempM + ' °C' : '-- °C'} <span class="badge-status" style="background:${evalM.color}; font-size:8px;">${evalM.estado}</span>`;
    document.getElementById('val-temp-l').innerHTML = `${tempL !== '--' ? tempL + ' °C' : '-- °C'} <span class="badge-status" style="background:${evalL.color}; font-size:8px;">${evalL.estado}</span>`;

    document.getElementById('val-temp-carcaza').innerText = tempCarcaza !== '--' ? tempCarcaza + ' °C' : '-- °C';
    document.getElementById('val-temp-llibre').innerHTML = `${tempLLibre !== '--' ? tempLLibre + ' °C' : '-- °C'} <span class="badge-status" style="background:${evalLL.color}; font-size:8px;">${evalLL.estado}</span>`;
    document.getElementById('val-temp-lacople').innerHTML = `${tempLAcople !== '--' ? tempLAcople + ' °C' : '-- °C'} <span class="badge-status" style="background:${evalLA.color}; font-size:8px;">${evalLA.estado}</span>`;

    const carbonesValidos = [];
    for (let i = 1; i <= 24; i++) {
        const valRaw = obtenerValorCarbonFlex(ultimaInsp, i);
        if (valRaw !== null) {
            carbonesValidos.push({
                key: `C${i}`,
                nombreFase: obtenerNombreCarbonConFase(ubTecnica, `C${i}`),
                val: valRaw
            });
        }
    }

    const selectEl = document.getElementById('select-carbon-filtro');
    const carbonFiltroSel = selectEl ? selectEl.value : 'TODOS';

    const container = document.getElementById('contenedor-carbones');
    container.querySelectorAll('.carbon-node').forEach(n => n.remove());

    const totalNodos = carbonesValidos.length;
    const radius = 105;

    const tbody = document.getElementById('tbody-carbones');
    tbody.innerHTML = '';

    let conteoAlerta = 0;
    let conteoBueno = 0;

    carbonesValidos.forEach((item, idx) => {
        let pctVida = 100;
        if (limites.nuevo > 0) pctVida = Math.round((item.val / limites.nuevo) * 100);
        if (pctVida < 0) pctVida = 0;

        let estadoTexto = "CAMBIAR PRONTO";
        let badgeClass = "badge-cambiar";
        let nodeClass = "node-cambiar";

        if (pctVida > 55) {
            estadoTexto = "OPTIMO";
            badgeClass = "badge-optimo";
            nodeClass = "node-optimo";
            conteoBueno++;
        } else if (pctVida >= 52) {
            estadoTexto = "OBSERVACION";
            badgeClass = "badge-observacion";
            nodeClass = "node-observacion";
            conteoAlerta++;
        } else {
            conteoAlerta++;
        }

        const angle = (idx * (360 / totalNodos) - 90) * (Math.PI / 180);
        const x = 150 + radius * Math.cos(angle) - 18;
        const y = 130 + radius * Math.sin(angle) - 18;

        const node = document.createElement('div');
        node.className = `carbon-node ${nodeClass}`;
        node.style.left = `${x}px`;
        node.style.top = `${y}px`;
        node.innerHTML = `<span>${item.nombreFase}</span><span style="font-size:8px;">${pctVida}%</span>`;
        container.appendChild(node);

        const valActual = item.val;
        const valAnterior = penultimaInsp ? obtenerValorCarbonFlex(penultimaInsp, idx + 1) : null;

        let desgasteStr = "--";
        let badgeDesgasteClass = "desgaste-cero";
        let iconDesgaste = "";

        if (valAnterior !== null) {
            const diff = valAnterior - valActual;
            if (diff > 0) {
                desgasteStr = `-${diff.toFixed(2)} mm`;
                badgeDesgasteClass = diff > 1.5 ? "desgaste-alto" : "desgaste-normal";
                iconDesgaste = '<i class="fa-solid fa-arrow-down"></i>';
            } else if (diff === 0) {
                desgasteStr = "0.00 mm";
                badgeDesgasteClass = "desgaste-cero";
                iconDesgaste = '<i class="fa-solid fa-minus"></i>';
            } else {
                desgasteStr = `+${Math.abs(diff).toFixed(2)} mm (Nuevo)`;
                badgeDesgasteClass = "desgaste-normal";
                iconDesgaste = '<i class="fa-solid fa-arrows-rotate"></i>';
            }
        }

        if (carbonFiltroSel === "TODOS" || item.key === carbonFiltroSel) {
            const tr = document.createElement('tr');
            tr.innerHTML = `
                <td>${ubTecnica}</td>
                <td><strong>${item.nombreFase}</strong></td>
                <td>${limites.nuevo.toFixed(2)}</td>
                <td>${valAnterior !== null ? valAnterior.toFixed(2) : '--'}</td>
                <td><strong>${valActual.toFixed(2)}</strong></td>
                <td><strong style="color:#b91c1c;">${limites.limite.toFixed(2)}</strong></td>
                <td><strong>${pctVida}%</strong></td>
                <td><span class="badge-desgaste ${badgeDesgasteClass}">${iconDesgaste} ${desgasteStr}</span></td>
                <td><span class="badge-status ${badgeClass}">${estadoTexto}</span></td>
            `;
            tbody.appendChild(tr);
        }
    });

    const totalCarbonesNum = carbonesValidos.length;
    document.getElementById('kpi-total-carbones').innerText = totalCarbonesNum;
    document.getElementById('kpi-carbones-alerta').innerText = conteoAlerta;
    document.getElementById('kpi-carbones-bueno').innerText = conteoBueno;

    const pctAlerta = totalCarbonesNum > 0 ? ((conteoAlerta / totalCarbonesNum) * 100).toFixed(1) : '0';
    const pctBueno = totalCarbonesNum > 0 ? ((conteoBueno / totalCarbonesNum) * 100).toFixed(1) : '100';

    document.getElementById('kpi-pct-alerta').innerText = `(${pctAlerta}%)`;
    document.getElementById('kpi-pct-bueno').innerText = `(${pctBueno}%)`;

    let itemsParaGrafico = carbonesValidos;
    if (carbonFiltroSel !== "TODOS") {
        itemsParaGrafico = carbonesValidos.filter(c => c.key === carbonFiltroSel);
    }

    renderizarGraficoTendenciaRango(inspecciones, itemsParaGrafico);

    renderizarMinigraficaParametro('chartMiniCorriente', inspecciones, ['Corriente', 'Amp'], '#16a34a');
    renderizarMinigraficaParametro('chartMiniK', inspecciones, ['K', 'Temp K', 'Temperatura K'], '#d90429');
    renderizarMinigraficaParametro('chartMiniM', inspecciones, ['M', 'Temp M', 'Temperatura M'], '#ea580c');
    renderizarMinigraficaParametro('chartMiniL', inspecciones, ['L', 'Temp L', 'Temperatura L'], '#0284c7');
    renderizarMinigraficaParametro('chartMiniCarcaza', inspecciones, ['Carcaza', 'Carcasa'], '#0284c7');
    renderizarMinigraficaParametro('chartMiniLLibre', inspecciones, ['L. Libre', 'L.Libre'], '#2563eb');
    renderizarMinigraficaParametro('chartMiniLAcople', inspecciones, ['L. Acople', 'L.Acople'], '#4f46e5');
}

function configurarOpcionesPanZoom(labelEjeY, dataValores = []) {
    let maxVal = 0;
    if (Array.isArray(dataValores) && dataValores.length > 0) {
        const nums = dataValores.filter(v => typeof v === 'number' && !isNaN(v));
        if (nums.length > 0) maxVal = Math.max(...nums);
    }

    return {
        responsive: true,
        maintainAspectRatio: false,
        interaction: { mode: 'index', intersect: false },
        layout: { padding: { top: 32, bottom: 8, left: 6, right: 6 } },
        plugins: {
            legend: { position: 'top', labels: { boxWidth: 10, font: { size: 10 }, usePointStyle: true } },
            tooltip: {
                enabled: true,
                position: 'nearest',
                backgroundColor: 'rgba(255, 255, 255, 0.96)',
                titleColor: '#0f172a',
                bodyColor: '#1e293b',
                borderColor: '#cbd5e1',
                borderWidth: 1,
                padding: 10,
                callbacks: {
                    label: function(context) {
                        return `${context.dataset.label}: ${context.parsed.y !== null ? context.parsed.y.toFixed(2) : '--'}`;
                    }
                }
            },
            datalabels: {
                display: 'auto',
                clamp: true,
                align: 'top',
                anchor: 'end',
                offset: 8,
                font: { size: 8, weight: 'bold' },
                formatter: function(value) { return value !== null && value !== undefined ? value.toFixed(1) : ''; }
            },
            zoom: {
                pan: { enabled: true, mode: 'x', modifierKey: null, threshold: 0 },
                zoom: { wheel: { enabled: true, speed: 0.1 }, pinch: { enabled: true }, mode: 'x' }
            }
        },
        scales: {
            x: { ticks: { font: { size: 9 }, maxRotation: 45 } },
            y: { 
                title: { display: !!labelEjeY, text: labelEjeY, font: { size: 10, weight: 'bold' } },
                suggestedMax: maxVal > 0 ? maxVal + (maxVal * 0.18) : undefined
            }
        }
    };
}

function desplazarGrafico(canvasId, deltaPixeles) {
    let chartObj = null;
    if (canvasId === 'chartDesgaste') chartObj = chartDesgasteInstance;
    else if (miniChartsInstances[canvasId]) chartObj = miniChartsInstances[canvasId];

    if (chartObj && typeof chartObj.pan === 'function') {
        chartObj.pan({ x: deltaPixeles }, undefined, 'default');
    }
}

function renderizarGraficoTendenciaRango(inspeccionesRango, itemsValidos) {
    const labelsFechas = inspeccionesRango.map(insp => formatearFechaLegible(insp._dateObj));
    const colores = ['#0284c7', '#2563eb', '#7c3aed', '#db2777', '#ea580c', '#d97706', '#65a30d', '#16a34a', '#0d9488', '#0891b2', '#4f46e5', '#c026d3'];

    const datasets = itemsValidos.map((item, idx) => {
        const numIndex = parseInt(item.key.replace('C', ''), 10);
        return {
            label: item.nombreFase,
            data: inspeccionesRango.map(insp => obtenerValorCarbonFlex(insp, numIndex)),
            borderColor: colores[idx % colores.length],
            backgroundColor: colores[idx % colores.length],
            borderWidth: 2,
            pointRadius: 3,
            pointHoverRadius: 6,
            spanGaps: true
        };
    });

    if (chartDesgasteInstance) chartDesgasteInstance.destroy();
    const ctx = document.getElementById('chartDesgaste').getContext('2d');
    chartDesgasteInstance = new Chart(ctx, {
        type: 'line',
        data: { labels: labelsFechas, datasets: datasets },
        options: configurarOpcionesPanZoom('Medida Carbón (mm)')
    });
}

function resetearZoomTendencia() {
    if (chartDesgasteInstance) chartDesgasteInstance.resetZoom();
}

function renderizarMinigraficaParametro(canvasId, inspeccionesRango, nombresCampo, colorLinea) {
    const canvasEl = document.getElementById(canvasId);
    if (!canvasEl) return;

    const labelsFechas = inspeccionesRango.map(insp => formatearFechaLegible(insp._dateObj));
    const dataValores = inspeccionesRango.map(insp => {
        const val = obtenerValorCampoFlex(insp, nombresCampo);
        return (val !== '--' && !isNaN(parseFloat(val))) ? parseFloat(val) : null;
    });

    if (miniChartsInstances[canvasId]) miniChartsInstances[canvasId].destroy();

    const ctx = canvasEl.getContext('2d');
    miniChartsInstances[canvasId] = new Chart(ctx, {
        type: 'line',
        data: {
            labels: labelsFechas,
            datasets: [{
                data: dataValores,
                borderColor: colorLinea,
                backgroundColor: 'transparent',
                borderWidth: 1.5,
                pointRadius: 3,
                spanGaps: true
            }]
        },
        options: configurarOpcionesPanZoom('', dataValores)
    });
}

async function descargarContenedorComoImagen(idContenedor, nombreArchivo) {
    const contenedor = document.getElementById(idContenedor);
    if (!contenedor) return;

    const scrollContainers = contenedor.querySelectorAll('.table-carbones-container');
    const estilosPrevios = [];

    scrollContainers.forEach(sc => {
        estilosPrevios.push({ el: sc, maxHeight: sc.style.maxHeight, overflow: sc.style.overflow });
        sc.style.maxHeight = 'none';
        sc.style.overflow = 'visible';
    });

    try {
        const canvas = await html2canvas(contenedor, { scale: 2, useCORS: true, backgroundColor: '#ffffff' });
        const enlace = document.createElement('a');
        enlace.download = `${nombreArchivo}_${new Date().toISOString().split('T')[0]}.png`;
        enlace.href = canvas.toDataURL('image/png');
        enlace.click();
    } catch (err) {
        console.error("Error al capturar imagen:", err);
        alert("Ocurrió un error al descargar la imagen.");
    } finally {
        estilosPrevios.forEach(item => {
            item.el.style.maxHeight = item.maxHeight;
            item.el.style.overflow = item.overflow;
        });
    }
}

function calcularPrediccionCarbonFlex(inspeccionesMotor, indiceCarbon, mmLimite) {
    let lecturas = [];
    const fInicioVal = document.getElementById('fecha-inicio').value;
    const fFinVal = document.getElementById('fecha-fin').value;
    const fInicio = fInicioVal ? new Date(fInicioVal + 'T00:00:00') : new Date('2026-01-01T00:00:00');
    const fFin = fFinVal ? new Date(fFinVal + 'T23:59:59') : null;

    inspeccionesMotor.forEach(insp => {
        const val = obtenerValorCarbonFlex(insp, indiceCarbon);
        const fechaObj = insp._dateObj || parsearFecha(insp['Fecha']);
        if (val !== null && fechaObj) {
            if (fechaObj.getFullYear() === 2026) {
                if (fechaObj >= fInicio && (!fFin || fechaObj <= fFin)) {
                    lecturas.push({ fecha: fechaObj, mm: val });
                }
            }
        }
    });

    if (lecturas.length < 2) return null;
    lecturas.sort((a, b) => a.fecha.getTime() - b.fecha.getTime());

    let inicioCicloActual = 0;
    for (let i = 1; i < lecturas.length; i++) {
        if (lecturas[i].mm - lecturas[i - 1].mm > 3.0) {
            inicioCicloActual = i;
        }
    }

    const lecturasCicloActual = lecturas.slice(inicioCicloActual);
    if (lecturasCicloActual.length < 2) return null;

    const primeraFecha = lecturasCicloActual[0].fecha;
    const N = lecturasCicloActual.length;

    let sumX = 0, sumY = 0, sumXY = 0, sumX2 = 0;
    lecturasCicloActual.forEach(d => {
        const x = (d.fecha.getTime() - primeraFecha.getTime()) / (1000 * 60 * 60 * 24);
        const y = d.mm;
        sumX += x;
        sumY += y;
        sumXY += x * y;
        sumX2 += x * x;
    });

    const denominador = (N * sumX2 - sumX * sumX);
    if (denominador === 0) return null;

    const m = (N * sumXY - sumX * sumY) / denominador;
    let mmDiaRegresion = Math.abs(m);

    let mmDiaReciente = 0;
    if (lecturasCicloActual.length >= 3) {
        const ult = lecturasCicloActual[lecturasCicloActual.length - 1];
        const penult = lecturasCicloActual[lecturasCicloActual.length - 2];
        const diasDiff = Math.max(1, (ult.fecha.getTime() - penult.fecha.getTime()) / (1000 * 60 * 60 * 24));
        const consumoReciente = penult.mm - ult.mm;
        if (consumoReciente > 0) mmDiaReciente = consumoReciente / diasDiff;
    }

    let mmDiaFinal = mmDiaRegresion;
    if (mmDiaReciente > mmDiaRegresion && mmDiaReciente > 0) {
        mmDiaFinal = (mmDiaReciente * 0.70) + (mmDiaRegresion * 0.30);
    }
    if (mmDiaFinal <= 0.0001) return null;

    const ultimaLectura = lecturasCicloActual[lecturasCicloActual.length - 1];
    const mmRestantes = ultimaLectura.mm - mmLimite;

    const diasTotalProyectados = Math.round(mmRestantes / mmDiaFinal);
    const fechaEstimadaObj = new Date(ultimaLectura.fecha.getTime() + (diasTotalProyectados * 24 * 60 * 60 * 1000));

    const hoy = new Date();
    hoy.setHours(0, 0, 0, 0);
    fechaEstimadaObj.setHours(0, 0, 0, 0);

    const diffTiempo = fechaEstimadaObj - hoy;
    const diasRestantesHoy = Math.ceil(diffTiempo / (1000 * 60 * 60 * 24));

    const proximoViernes = new Date(hoy);
    const diaSemanaHoy = hoy.getDay();
    const diasHastaViernes = (5 - diaSemanaHoy + 7) % 7 || 7;
    proximoViernes.setDate(hoy.getDate() + diasHastaViernes);

    const diasHastaProximoViernes = Math.max(1, Math.round((proximoViernes.getTime() - ultimaLectura.fecha.getTime()) / (1000 * 60 * 60 * 24)));
    const mmEstimadoProximoViernes = Math.max(mmLimite, parseFloat((ultimaLectura.mm - (diasHastaProximoViernes * mmDiaFinal)).toFixed(2)));

    return {
        mmDia: mmDiaFinal,
        diasRestantes: diasRestantesHoy,
        fechaEstimadaStr: formatearFechaLegible(fechaEstimadaObj),
        fechaEstimadaObj: fechaEstimadaObj,
        lecturas: lecturasCicloActual,
        proximoViernesStr: formatearFechaLegible(proximoViernes),
        mmEstimadoProximoViernes: mmEstimadoProximoViernes
    };
}

// RENDERIZAR MÓDULO PREDICTIVO (FILTRADO EXACTO POR CONTRATISTA O MOTOR INDIVIDUAL)
function renderizarModuloPredictivo() {
    const tbody = document.getElementById('tbody-predictivo');
    if (!tbody) return;

    tbody.innerHTML = '';
    if (!db.motores || db.motores.length === 0) return;

    listaPronosticosGlobal = [];
    let contadorCritico = 0, contadorAtencion = 0, contadorOperativo = 0;

    const selectSidebar = document.getElementById('select-motor-filtro');
    const selectContratista = document.getElementById('select-contratista-filtro');

    const motorSeleccionadoVal = selectSidebar ? selectSidebar.value : "TODOS";
    const contratistaSel = selectContratista ? selectContratista.value : "TODOS";

    let motoresAEvaluar = db.motores;

    if (motorSeleccionadoVal !== "TODOS" && motorSeleccionadoVal !== "" && db.motores[motorSeleccionadoVal]) {
        motoresAEvaluar = [db.motores[motorSeleccionadoVal]];
    } else if (contratistaSel !== "TODOS") {
        motoresAEvaluar = db.motores.filter(m => (m['Contratista'] || 'UNACEM').trim().toUpperCase() === contratistaSel);
    }

    motoresAEvaluar.forEach(m => {
        const ubTecnica = m['Ubicacion tecnica'] || m['Ubicación técnica'] || '';
        const codEquipo = m['Codigo de equipo'] || m['Código de equipo'] || '';
        const nombreMotor = m['Motor'] || ubTecnica;
        const contratistaMotor = (m['Contratista'] || 'UNACEM').trim().toUpperCase();

        const limites = obtenerLimiteEspecificoMotor(ubTecnica, codEquipo);

        const keyNorm = normalizarTexto(ubTecnica);
        const codEquipoNorm = normalizarTexto(codEquipo);

        let inspMotor = db.inspecciones.filter(insp => {
            const codInsp = normalizarTexto(insp['Ubicacion tecnica'] || insp['Ubicación técnica'] || insp['UbicacionTecnica'] || insp['Codigo de equipo']);
            return (codInsp === keyNorm || codInsp === codEquipoNorm) && codInsp !== '';
        });

        if (inspMotor.length === 0) return;
        inspMotor.forEach(insp => { if (!insp._dateObj) insp._dateObj = parsearFecha(insp['Fecha']); });
        inspMotor = inspMotor.filter(insp => insp._dateObj !== null);
        inspMotor.sort((a, b) => a._dateObj.getTime() - b._dateObj.getTime());

        if (inspMotor.length === 0) return;
        const ultimaInsp = inspMotor[inspMotor.length - 1];

        for (let i = 1; i <= 24; i++) {
            const valActual = obtenerValorCarbonFlex(ultimaInsp, i);
            if (valActual !== null) {
                const pred = calcularPrediccionCarbonFlex(inspMotor, i, limites.limite);
                if (pred) {
                    let claseRiesgo = "badge-riesgo-normal";
                    let claseDias = "col-dias-normal";
                    let etiquetaRiesgo = "OPERATIVO";

                    if (pred.diasRestantes <= 10) {
                        claseRiesgo = "badge-riesgo-critico";
                        claseDias = "col-dias-critico";
                        etiquetaRiesgo = "CRÍTICO";
                        contadorCritico++;
                    } else if (pred.diasRestantes <= 60) {
                        claseRiesgo = "badge-riesgo-observacion";
                        claseDias = "col-dias-observacion";
                        etiquetaRiesgo = "OBSERVACIÓN";
                        contadorAtencion++;
                    } else {
                        contadorOperativo++;
                    }

                    const nombreConFase = obtenerNombreCarbonConFase(ubTecnica, `C${i}`);
                    listaPronosticosGlobal.push({
                        carbon: nombreConFase,
                        carbonKey: `C${i}`,
                        motor: `${nombreMotor} (${contratistaMotor})`,
                        ubTecnica: ubTecnica,
                        mmUltimo: valActual,
                        limite: limites.limite,
                        desgasteDia: pred.mmDia,
                        diasRestantes: pred.diasRestantes,
                        fechaCambio: pred.fechaEstimadaStr,
                        claseRiesgo: claseRiesgo,
                        claseDias: claseDias,
                        etiquetaRiesgo: etiquetaRiesgo,
                        lecturas: pred.lecturas,
                        proximoViernesStr: pred.proximoViernesStr,
                        mmEstimadoProximoViernes: pred.mmEstimadoProximoViernes
                    });
                }
            }
        }
    });

    document.getElementById('kpi-predictivo-critico').innerText = contadorCritico;
    document.getElementById('kpi-predictivo-atencion').innerText = contadorAtencion;
    document.getElementById('kpi-predictivo-operativo').innerText = contadorOperativo;

    if (listaPronosticosGlobal.length === 0) {
        tbody.innerHTML = '<tr><td colspan="9" style="text-align:center; color:#94a3b8; padding: 20px;">Sin suficientes datos para proyectar desgaste en los filtros seleccionados.</td></tr>';
        if (chartPredictivoInstance) chartPredictivoInstance.destroy();
        return;
    }

    listaPronosticosGlobal.sort((a, b) => a.diasRestantes - b.diasRestantes);

    listaPronosticosGlobal.forEach((item) => {
        const tr = document.createElement('tr');
        tr.style.cursor = 'pointer';
        tr.onclick = function() { renderizarGraficoPredictivoGlobal(item); };
        tr.innerHTML = `
            <td><strong>${item.carbon}</strong></td>
            <td>${item.motor}</td>
            <td>${item.ubTecnica}</td>
            <td>${item.mmUltimo.toFixed(2)}</td>
            <td>${item.limite.toFixed(2)}</td>
            <td>${item.desgasteDia.toFixed(3)}</td>
            <td class="${item.claseDias}">${item.diasRestantes} días</td>
            <td><strong>${item.fechaCambio}</strong></td>
            <td><span class="${item.claseRiesgo}">${item.etiquetaRiesgo}</span></td>
        `;
        tbody.appendChild(tr);
    });

    renderizarGraficoPredictivoGlobal(listaPronosticosGlobal[0]);
}

function renderizarGraficoPredictivoGlobal(pronosticoTop) {
    const canvas = document.getElementById('chartPredictivoGlobal');
    if (!canvas || !pronosticoTop || !pronosticoTop.lecturas) return;

    const lecturas = pronosticoTop.lecturas;
    const labelsFechas = lecturas.map(l => formatearFechaLegible(l.fecha));
    const dataReal = lecturas.map(l => l.mm);
    const ultimaLectura = lecturas[lecturas.length - 1];

    labelsFechas.push(`Próx. Viernes (${pronosticoTop.proximoViernesStr})`);
    labelsFechas.push(`Cambio Estimado (${pronosticoTop.fechaCambio})`);

    const dataProyeccion = new Array(lecturas.length - 1).fill(null);
    dataProyeccion.push(ultimaLectura.mm);
    dataProyeccion.push(pronosticoTop.mmEstimadoProximoViernes);
    dataProyeccion.push(pronosticoTop.limite);

    if (chartPredictivoInstance) chartPredictivoInstance.destroy();

    const ctx = canvas.getContext('2d');
    chartPredictivoInstance = new Chart(ctx, {
        type: 'line',
        data: {
            labels: labelsFechas,
            datasets: [
                {
                    label: `Medición Real - ${pronosticoTop.carbon} (${pronosticoTop.ubTecnica})`,
                    data: dataReal,
                    borderColor: '#0284c7',
                    backgroundColor: '#0284c7',
                    borderWidth: 2,
                    pointRadius: 4
                },
                {
                    label: `Proyección Futura (Próx. Viernes: ${pronosticoTop.mmEstimadoProximoViernes} mm | Límite: ${pronosticoTop.limite.toFixed(2)} mm)`,
                    data: dataProyeccion,
                    borderColor: '#dc2626',
                    borderDash: [5, 5],
                    backgroundColor: 'transparent',
                    borderWidth: 2,
                    pointRadius: 5,
                    pointStyle: 'rectRot'
                }
            ]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            layout: { padding: { top: 32, bottom: 8 } },
            plugins: {
                legend: { position: 'top', labels: { boxWidth: 12, font: { size: 10 } } },
                datalabels: {
                    align: 'top',
                    anchor: 'end',
                    offset: 8,
                    font: { size: 9, weight: 'bold' },
                    formatter: function(value) { return value !== null && value !== undefined ? value.toFixed(2) : ''; }
                },
                zoom: {
                    pan: { enabled: true, mode: 'x', modifierKey: null, threshold: 0 },
                    zoom: { wheel: { enabled: true, speed: 0.1 }, mode: 'x' }
                }
            },
            scales: {
                x: { ticks: { font: { size: 9 } } },
                y: { 
                    title: { display: true, text: 'Longitud del Carbón (mm)' },
                    suggestedMax: Math.max(...dataReal) + 4
                }
            }
        }
    });
}

function renderizarReporteConsolidado() {
    const contenedor = document.getElementById('contenedor-reportes-motores');
    if (!contenedor) return;

    contenedor.innerHTML = '';
    datosReporteParaExportar = [];

    Object.keys(graficosReporteInstances).forEach(k => {
        if (graficosReporteInstances[k]) graficosReporteInstances[k].destroy();
    });
    graficosReporteInstances = {};

    let totalMotores = db.motores.length;
    let observadosOAlerta = 0;

    const fInicioVal = document.getElementById('fecha-inicio').value;
    const fFinVal = document.getElementById('fecha-fin').value;
    const selectContratista = document.getElementById('select-contratista-filtro');
    const contratistaSel = selectContratista ? selectContratista.value : "TODOS";

    const fInicio = fInicioVal ? new Date(fInicioVal + 'T00:00:00') : null;
    const fFin = fFinVal ? new Date(fFinVal + 'T23:59:59') : null;

    db.motores.forEach((m, idxMotor) => {
        const contratistaMotor = (m['Contratista'] || 'UNACEM').trim().toUpperCase();

        if (contratistaSel !== "TODOS" && contratistaMotor !== contratistaSel) return;

        const ubTecnica = m['Ubicacion tecnica'] || m['Ubicación técnica'] || '';
        const codEquipo = m['Codigo de equipo'] || m['Código de equipo'] || '';

        const limites = obtenerLimiteEspecificoMotor(ubTecnica, codEquipo);

        const keyNorm = normalizarTexto(ubTecnica);
        const codEquipoNorm = normalizarTexto(codEquipo);

        let inspMotor = db.inspecciones.filter(insp => {
            const codInsp = normalizarTexto(insp['Ubicacion tecnica'] || insp['Ubicación técnica'] || insp['UbicacionTecnica'] || insp['Codigo de equipo']);
            return (codInsp === keyNorm || codInsp === codEquipoNorm) && codInsp !== '';
        });

        if (inspMotor.length === 0) return;

        inspMotor.forEach(i => i._dateObj = parsearFecha(i['Fecha']));
        inspMotor.sort((a, b) => (a._dateObj ? a._dateObj.getTime() : 0) - (b._dateObj ? b._dateObj.getTime() : 0));

        if (fInicio || fFin) {
            inspMotor = inspMotor.filter(insp => {
                if (!insp._dateObj) return true;
                if (fInicio && insp._dateObj < fInicio) return false;
                if (fFin && insp._dateObj > fFin) return false;
                return true;
            });
        }

        if (inspMotor.length === 0) return;

        const ultimaInsp = inspMotor[inspMotor.length - 1];
        const penultimaInsp = inspMotor.length > 1 ? inspMotor[inspMotor.length - 2] : null;

        let detallesCarbonesAnomalia = [];
        let detallesTemperaturaAnomalia = [];
        let carbonesObservadosItems = [];
        let carbonesNuevosItems = [];

        let tieneObservados = false;
        let tieneCambiados = false;
        let tieneSoloLectura = false;

        for (let k = 1; k <= 24; k++) {
            const valActual = obtenerValorCarbonFlex(ultimaInsp, k);
            const valAnterior = penultimaInsp ? obtenerValorCarbonFlex(penultimaInsp, k) : null;

            if (valActual !== null) {
                let pctVida = limites.nuevo > 0 ? Math.round((valActual / limites.nuevo) * 100) : 100;
                let diffSemanal = valAnterior !== null ? (valAnterior - valActual) : 0;

                const pred = calcularPrediccionCarbonFlex(inspMotor, k, limites.limite);
                const fechaEst = pred ? pred.fechaEstimadaStr : "S/F";
                const diasRest = pred ? pred.diasRestantes : 999;

                let esCriticoPredictivo = (diasRest <= 10);
                let esBajoPct = (pctVida <= 55);
                let esAcelerado = (diffSemanal > 1.5);
                let esCarbonNuevo = (diffSemanal < -3.0);
                let esErrorLectura = (diffSemanal < 0 && diffSemanal >= -3.0);

                if (esCriticoPredictivo || esBajoPct || esAcelerado || esCarbonNuevo || esErrorLectura) {
                    let motivoStr = "Desgaste Acelerado Semanal";
                    let estiloFila = "";
                    const cKeyOriginal = `C${k}`;
                    const nombreConFase = obtenerNombreCarbonConFase(ubTecnica, cKeyOriginal);

                    if (esCarbonNuevo) {
                        motivoStr = "CARBONES CAMBIADOS";
                        estiloFila = 'style="background:#dcfce7; color:#166534;"';
                        carbonesNuevosItems.push({ key: cKeyOriginal, label: nombreConFase });
                        tieneCambiados = true;
                    } else if (esCriticoPredictivo) {
                        motivoStr = `CAMBIO URGENTE: ${diasRest} DÍAS RESTANTES (PREDICTIVO)`;
                        estiloFila = 'style="background:#fee2e2; color:#991b1b; font-weight:800;"';
                        carbonesObservadosItems.push({ key: cKeyOriginal, label: nombreConFase });
                        tieneObservados = true;
                    } else if (esBajoPct) {
                        motivoStr = `Vida Útil Baja (${pctVida}%)`;
                        estiloFila = 'style="background:#fef2f2; color:#b91c1c;"';
                        carbonesObservadosItems.push({ key: cKeyOriginal, label: nombreConFase });
                        tieneObservados = true;
                    } else if (esAcelerado) {
                        motivoStr = "Desgaste Acelerado Semanal";
                        carbonesObservadosItems.push({ key: cKeyOriginal, label: nombreConFase });
                        tieneObservados = true;
                    } else if (esErrorLectura) {
                        motivoStr = "ERROR DE LECTURA";
                        estiloFila = 'style="background:#fef3c7; color:#92400e;"';
                        tieneSoloLectura = true;
                    }

                    detallesCarbonesAnomalia.push({
                        carbon: nombreConFase,
                        carbonKey: cKeyOriginal,
                        medidaAnterior: valAnterior !== null ? `${valAnterior.toFixed(2)} mm` : '--',
                        medidaActual: `${valActual.toFixed(2)} mm`,
                        limiteCambio: `${limites.limite.toFixed(2)} mm`,
                        pct: pctVida,
                        desgasteSemana: diffSemanal > 0 ? `-${diffSemanal.toFixed(2)} mm` : `+${Math.abs(diffSemanal).toFixed(2)} mm`,
                        fechaCambioEstimada: fechaEst,
                        motivo: motivoStr,
                        estiloFila: estiloFila
                    });
                }
            }
        }

        const camposTermicos = [
            { nombre: 'Fase K', val: obtenerValorCampoFlex(ultimaInsp, ['K', 'Temp K']) },
            { nombre: 'Fase M', val: obtenerValorCampoFlex(ultimaInsp, ['M', 'Temp M']) },
            { nombre: 'Fase L', val: obtenerValorCampoFlex(ultimaInsp, ['L', 'Temp L']) },
            { nombre: 'Rodamiento LL', val: obtenerValorCampoFlex(ultimaInsp, ['L. Libre', 'L.Libre']) },
            { nombre: 'Rodamiento LA', val: obtenerValorCampoFlex(ultimaInsp, ['L. Acople', 'L.Acople']) }
        ];

        camposTermicos.forEach(item => {
            const num = parseFloat(item.val);
            if (!isNaN(num)) {
                const evalT = evaluarTemperaturaIndustrial(num);
                if (evalT.esAlerta) {
                    tieneObservados = true;
                    detallesTemperaturaAnomalia.push({
                        componente: item.nombre,
                        temp: `${num.toFixed(1)} °C`,
                        condicion: evalT.estado,
                        recomendacion: evalT.recomendacion,
                        color: evalT.color
                    });
                }
            }
        });

        if (detallesCarbonesAnomalia.length > 0 || detallesTemperaturaAnomalia.length > 0) {
            observadosOAlerta++;

            let badgeTexto = "REQUIERE ATENCIÓN";
            let badgeClase = "badge-cambiar";

            if (tieneObservados || (tieneObservados && tieneCambiados)) {
                badgeTexto = "REQUIERE ATENCIÓN";
                badgeClase = "badge-cambiar";
            } else if (tieneCambiados && !tieneObservados && !tieneSoloLectura) {
                badgeTexto = "CARBONES CAMBIADOS";
                badgeClase = "badge-optimo";
            } else if (tieneSoloLectura && !tieneObservados && !tieneCambiados) {
                badgeTexto = "ERROR DE LECTURA";
                badgeClase = "badge-observacion";
            }

            datosReporteParaExportar.push({
                motor: m['Motor'] || ubTecnica,
                ubTecnica: ubTecnica,
                ubPlanta: m['Ubicacion en planta'] || m['Ubicación en planta'] || '--',
                contratista: contratistaMotor,
                estadoPredominante: badgeTexto,
                carbones: detallesCarbonesAnomalia,
                temperaturas: detallesTemperaturaAnomalia,
                carbonesObservadosItems: carbonesObservadosItems,
                carbonesNuevosItems: carbonesNuevosItems,
                limiteNum: limites.limite,
                inspeccionesMotor: inspMotor
            });

            const canvasIdObs = `chart-obs-motor-${idxMotor}`;
            const canvasIdNuevos = `chart-nuevos-motor-${idxMotor}`;

            const tarjetaDiv = document.createElement('div');
            tarjetaDiv.className = 'tarjeta-motor-reporte';

            let htmlMotor = `
                <div class="cabecera-motor-reporte">
                    <div>
                        <strong style="font-size: 13px; color: var(--primary-blue);">${m['Motor'] || ubTecnica}</strong><br>
                        <small style="color: #64748b;"><strong>Ub. Técnica:</strong> ${ubTecnica} | <strong>Planta:</strong> ${m['Ubicacion en planta'] || '--'} | <strong>Contratista:</strong> ${contratistaMotor}</small>
                    </div>
                    <div>
                        <span class="badge-status ${badgeClase}">${badgeTexto}</span>
                    </div>
                </div>

                <div style="font-size: 11px; font-weight: 700; color: #991b1b; margin-bottom: 6px;">
                    <i class="fa-solid fa-magnifying-glass"></i> Evidencia Técnica de Inspección (${detallesCarbonesAnomalia.length} eventos de carbón | ${detallesTemperaturaAnomalia.length} alertas térmicas):
                </div>
            `;

            if (detallesCarbonesAnomalia.length > 0) {
                htmlMotor += `<table class="table-custom" style="width: 100%; margin-bottom: 12px;">`;
                htmlMotor += `<thead><tr style="background:#fee2e2;"><th>Carbón</th><th>Medida Anterior</th><th>Medida Actual</th><th>Límite Cambio</th><th>% Vida Útil</th><th>Variación Semanal</th><th>Fecha Cambio Estimada</th><th>Diagnóstico / Motivo</th></tr></thead><tbody>`;
                detallesCarbonesAnomalia.forEach(c => {
                    htmlMotor += `<tr ${c.estiloFila}><td><strong>${c.carbon}</strong></td><td>${c.medidaAnterior}</td><td>${c.medidaActual}</td><td><strong>${c.limiteCambio}</strong></td><td><strong>${c.pct}%</strong></td><td>${c.desgasteSemana}</td><td><strong>${c.fechaCambioEstimada}</strong></td><td style="font-weight:bold;">${c.motivo}</td></tr>`;
                });
                htmlMotor += `</tbody></table>`;

                htmlMotor += `<div class="graficos-reporte-columna">`;

                if (carbonesObservadosItems.length > 0) {
                    htmlMotor += `
                        <div class="caja-grafico-reporte">
                            <span style="font-size: 11px; font-weight: bold; color: var(--accent-red);">
                                <i class="fa-solid fa-chart-line"></i> Tendencia de Carbones Observados (${carbonesObservadosItems.map(i=>i.label).join(', ')}) vs Límite (${limites.limite.toFixed(2)} mm)
                            </span>
                            <div style="height: 240px; position: relative; width: 100%; margin-top: 8px;">
                                <canvas id="${canvasIdObs}"></canvas>
                            </div>
                        </div>
                    `;
                }

                if (carbonesNuevosItems.length > 0) {
                    htmlMotor += `
                        <div class="caja-grafico-reporte" style="border-color: #86efac; background: #f0fdf4;">
                            <span style="font-size: 11px; font-weight: bold; color: #166534;">
                                <i class="fa-solid fa-arrows-rotate"></i> Seguimiento Carbones Cambiados (${carbonesNuevosItems.map(i=>i.label).join(', ')})
                            </span>
                            <div style="height: 240px; position: relative; width: 100%; margin-top: 8px;">
                                <canvas id="${canvasIdNuevos}"></canvas>
                            </div>
                        </div>
                    `;
                }

                htmlMotor += `</div>`;
            }

            if (detallesTemperaturaAnomalia.length > 0) {
                htmlMotor += `<table class="table-custom" style="width: 100%; margin-top: 10px;">`;
                htmlMotor += `<thead><tr style="background:#fef3c7;"><th>Componente Térmico</th><th>Temperatura Medida</th><th>Condición</th><th>Recomendación de Campo</th></tr></thead><tbody>`;
                detallesTemperaturaAnomalia.forEach(t => {
                    htmlMotor += `<tr><td><strong>${t.componente}</strong></td><td>${t.temp}</td><td><span class="badge-status" style="background:${t.color};">${t.condicion}</span></td><td>${t.recomendacion}</td></tr>`;
                });
                htmlMotor += `</tbody></table>`;
            }

            tarjetaDiv.innerHTML = htmlMotor;
            contenedor.appendChild(tarjetaDiv);

            setTimeout(() => {
                if (carbonesObservadosItems.length > 0) {
                    renderizarGraficoReporteMotor(canvasIdObs, inspMotor, carbonesObservadosItems, limites.limite, true);
                }
                if (carbonesNuevosItems.length > 0) {
                    renderizarGraficoReporteMotor(canvasIdNuevos, inspMotor, carbonesNuevosItems, limites.limite, false);
                }
            }, 60);
        }
    });

    document.getElementById('reporte-total-motores').innerText = totalMotores;
    document.getElementById('reporte-total-operativos').innerText = totalMotores - observadosOAlerta;
    document.getElementById('reporte-total-criticos').innerText = observadosOAlerta;

    if (observadosOAlerta === 0) {
        contenedor.innerHTML = '<div style="text-align:center; color:#16a34a; font-weight: bold; padding: 30px; background:#f0fdf4; border:1px solid #bbf7d0; border-radius:8px;">Todos los motores operan bajo parámetros normales sin alertas de desgaste o temperatura.</div>';
    }
}

function renderizarGraficoReporteMotor(canvasId, inspeccionesMotor, carbonesItems, mmLimite, incluirLimite) {
    const canvasEl = document.getElementById(canvasId);
    if (!canvasEl) return;

    const labelsFechas = inspeccionesMotor.map(insp => formatearFechaLegible(insp._dateObj));
    const paleta = ['#dc2626', '#ea580c', '#d97706', '#7c3aed', '#0284c7', '#059669', '#2563eb'];

    const datasets = carbonesItems.map((item, idx) => {
        const numIndex = parseInt(item.key.replace('C', ''), 10);
        return {
            label: item.label,
            data: inspeccionesMotor.map(insp => obtenerValorCarbonFlex(insp, numIndex)),
            borderColor: paleta[idx % paleta.length],
            backgroundColor: paleta[idx % paleta.length],
            borderWidth: 1.5,
            pointRadius: 2.5,
            pointHoverRadius: 5,
            spanGaps: true,
            datalabels: {
                display: 'auto',
                align: 'top',
                anchor: 'end',
                offset: 8,
                color: '#0f172a',
                font: { size: 8, weight: 'bold' },
                formatter: function(value) { return value !== null && value !== undefined ? value.toFixed(1) : ''; }
            }
        };
    });

    if (incluirLimite) {
        datasets.push({
            label: `Límite (${mmLimite.toFixed(2)} mm)`,
            data: new Array(labelsFechas.length).fill(mmLimite),
            borderColor: '#991b1b',
            borderDash: [5, 5],
            borderWidth: 1.2,
            pointRadius: 0,
            fill: false,
            datalabels: {
                display: function(context) { return context.dataIndex === context.dataset.data.length - 1; },
                align: 'right',
                anchor: 'end',
                offset: 6,
                color: '#991b1b',
                font: { size: 9, weight: 'bold' },
                formatter: function(value) { return `Límite: ${value.toFixed(1)} mm`; }
            }
        });
    }

    if (graficosReporteInstances[canvasId]) graficosReporteInstances[canvasId].destroy();

    const ctx = canvasEl.getContext('2d');
    graficosReporteInstances[canvasId] = new Chart(ctx, {
        type: 'line',
        data: { labels: labelsFechas, datasets: datasets },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            layout: { padding: { top: 32, right: 40, bottom: 8, left: 8 } },
            plugins: { legend: { position: 'top', labels: { boxWidth: 10, font: { size: 9 }, usePointStyle: true } } },
            scales: {
                x: { ticks: { font: { size: 8 }, maxRotation: 45 } },
                y: { title: { display: true, text: 'Medida (mm)', font: { size: 8, weight: 'bold' } } }
            }
        }
    });
}

function renderizarHistorialReportesUI() {
    const contenedor = document.getElementById('lista-historial-contenedor');
    if (!contenedor) return;

    let listaHistorial = [];
    if (db.historialExcel && db.historialExcel.length > 0) {
        listaHistorial = db.historialExcel.map((row, idx) => {
            const fechaVal = obtenerValorCampoFlex(row, ['FECHA', 'Fecha', 'fecha']);
            const motoresVal = obtenerValorCampoFlex(row, ['TOTAL_MOTORES', 'TotalMotores', 'Total Motores', 'TOTAL MOTORES']);
            let linkVal = obtenerValorCampoFlex(row, ['LINK_PDF', 'LinkPDF', 'Link PDF', 'LINK PDF', 'Link', 'Vínculo al elemento']);

            if (linkVal && linkVal !== '--' && linkVal.includes('sharepoint.com')) {
                let urlLimpia = linkVal.split('?')[0];
                linkVal = `${urlLimpia}?csf=1&web=1`;
            }

            return {
                id: idx,
                fecha: fechaVal !== '--' ? fechaVal : '--/--/----',
                totalObservados: motoresVal !== '--' ? motoresVal : '--',
                linkPdf: linkVal !== '--' ? linkVal : '#'
            };
        }).filter(item => item.linkPdf !== '#' && item.linkPdf !== '' && !item.linkPdf.includes('triggerBody'));
    }

    if (listaHistorial.length === 0) {
        contenedor.innerHTML = '<div style="text-align:center; color:#94a3b8; padding:30px;">No hay reportes archivados actualmente. Haz clic en "Actualizar Base" para sincronizar.</div>';
        return;
    }

    contenedor.innerHTML = '';
    listaHistorial.slice().reverse().forEach((item) => {
        const itemDiv = document.createElement('div');
        itemDiv.className = 'historial-item';
        itemDiv.style.marginBottom = '12px';
        itemDiv.innerHTML = `
            <div class="historial-header" style="display: flex; justify-content: space-between; align-items: center; padding: 14px 18px; background: #ffffff; border: 1px solid #cbd5e1; border-radius: 6px;">
                <div>
                    <strong style="color: var(--primary-blue); font-size: 13px;">
                        <i class="fa-regular fa-file-pdf" style="color: #dc2626; margin-right: 6px; font-size: 16px;"></i> Reporte Emitido: ${item.fecha}
                    </strong>
                    <span style="margin-left: 12px; font-size: 11px; color: #64748b;">(${item.totalObservados} Motores con Eventos)</span>
                </div>
                <div>
                    <a href="${item.linkPdf}" target="_blank" rel="noopener noreferrer" class="btn-action" style="font-size: 11px; background: #0284c7; text-decoration: none; display: inline-flex; align-items: center; gap: 6px; padding: 6px 14px; border-radius: 4px; color: #fff;">
                        <i class="fa-solid fa-arrow-up-right-from-square"></i> Abrir / Descargar PDF
                    </a>
                </div>
            </div>
        `;
        contenedor.appendChild(itemDiv);
    });
}

function guardarConfiguracionEnvio() {
    const dia = document.getElementById('cfg-dia-envio').value;
    const hora = document.getElementById('cfg-hora-envio').value;
    localStorage.setItem('unacem_cfg_envio', JSON.stringify({ dia: dia, hora: hora }));
    const lbl = document.getElementById('lbl-status-envio');
    if (lbl) {
        lbl.innerText = "Configuración guardada con éxito.";
        setTimeout(() => { lbl.innerText = ""; }, 3000);
    }
}

function cargarConfiguracionEnvio() {
    const cfg = JSON.parse(localStorage.getItem('unacem_cfg_envio') || '{"dia":"5","hora":"18:00"}');
    const diaEl = document.getElementById('cfg-dia-envio');
    const horaEl = document.getElementById('cfg-hora-envio');
    if (diaEl) diaEl.value = cfg.dia;
    if (horaEl) horaEl.value = cfg.hora;
}

function verificarEnvioProgramadoOutlook() {
    const cfg = JSON.parse(localStorage.getItem('unacem_cfg_envio') || '{"dia":"5","hora":"18:00"}');
    const hoy = new Date();
    const diaActual = hoy.getDay();
    const horaActualStr = String(hoy.getHours()).padStart(2, '0') + ':' + String(hoy.getMinutes()).padStart(2, '0');

    const diaCoincide = (cfg.dia === 'todos' || parseInt(cfg.dia, 10) === diaActual);
    const horaCoincide = horaActualStr >= cfg.hora;
    const fechaHoyStr = hoy.toISOString().split('T')[0];
    const ultimoEnvio = localStorage.getItem('unacem_ultimo_envio_auto');

    if (diaCoincide && horaCoincide && ultimoEnvio !== fechaHoyStr) {
        if (localStorage.getItem('unacem_db_consolidado')) {
            setTimeout(() => {
                enviarReporteOutlookAutomatico(false);
                localStorage.setItem('unacem_ultimo_envio_auto', fechaHoyStr);
            }, 2500);
        }
    }
}

async function enviarReporteOutlookAutomatico(esManual = false) {
    const vistaReportes = document.getElementById('vista-reportes');
    const elemento = document.getElementById('area-reporte-pdf');
    if (!vistaReportes || !elemento) return;

    if (!db.inspecciones || db.inspecciones.length === 0) {
        const dbGuardada = localStorage.getItem('unacem_db_consolidado');
        if (dbGuardada) db = JSON.parse(dbGuardada);
    }

    if (!db.inspecciones || db.inspecciones.length === 0) {
        if (esManual) alert("No hay datos cargados para enviar.");
        return;
    }

    db.inspecciones.forEach(i => {
        if (!i._dateObj || !(i._dateObj instanceof Date)) i._dateObj = parsearFecha(i['Fecha']);
    });

    establecerRangoFechasPredeterminado();

    const estiloDisplayOriginal = vistaReportes.style.display;
    vistaReportes.style.position = 'fixed';
    vistaReportes.style.left = '-9999px';
    vistaReportes.style.top = '0';
    vistaReportes.style.width = '1080px';
    vistaReportes.style.display = 'block';
    vistaReportes.style.visibility = 'visible';

    renderizarReporteConsolidado();
    await new Promise(resolve => setTimeout(resolve, 1200));

    const hoyStr = new Date().toLocaleDateString('es-PE') + ' ' + new Date().toLocaleTimeString('es-PE', { hour: '2-digit', minute: '2-digit' });
    const lblPdfFecha = document.getElementById('pdf-fecha-generacion');
    if (lblPdfFecha) lblPdfFecha.innerText = hoyStr;

    const opciones = {
        margin: [8, 8, 8, 8],
        filename: `Reporte_Carbones_UNACEM_${new Date().toISOString().split('T')[0]}.pdf`,
        image: { type: 'jpeg', quality: 0.65 },
        html2canvas: { scale: 1.2, useCORS: true, logging: false, removeContainer: true },
        jsPDF: { unit: 'mm', format: 'a4', orientation: 'landscape', compress: true },
        pagebreak: { mode: ['css', 'legacy'] }
    };

    try {
        const botonesAccion = elemento.querySelector('.no-print');
        if (botonesAccion) botonesAccion.style.display = 'none';

        const pdfBase64Full = await html2pdf().set(opciones).from(elemento).outputPdf('datauristring');
        if (botonesAccion) botonesAccion.style.display = 'flex';

        vistaReportes.style.display = estiloDisplayOriginal;
        vistaReportes.style.position = '';
        vistaReportes.style.left = '';
        vistaReportes.style.top = '';
        vistaReportes.style.width = '';
        vistaReportes.style.visibility = '';

        const pdfBase64Limpio = pdfBase64Full.includes(',') ? pdfBase64Full.split(',')[1] : pdfBase64Full;

        const payload = {
            destinatarios: CORREOS_DESTINO,
            asunto: `[UNACEM] Informe de Carbones y Anomalías - ${new Date().toLocaleDateString('es-PE')}`,
            cuerpoMensaje: `
                <div style="font-family: Arial, sans-serif; color: #1e293b;">
                    <h2 style="color: #002244;">UNACEM - Sistema de Inspección y Monitoreo</h2>
                    <p>Estimados,</p>
                    <p>Se adjunta el reporte consolidado correspondiente a la inspección.</p>
                    <p><strong>Fecha de Generación:</strong> ${hoyStr}</p>
                    <hr style="border: 0; border-top: 1px solid #cbd5e1; margin: 15px 0;">
                    <small style="color: #64748b;">Mensaje generado automáticamente por el sistema de mantenimiento predictivo.</small>
                </div>
            `,
            nombreArchivo: `Reporte_Carbones_UNACEM_${new Date().toISOString().split('T')[0]}.pdf`,
            archivoBase64: pdfBase64Limpio,
            fecha: hoyStr,
            totalMotores: datosReporteParaExportar.length
        };

        const respuesta = await fetch(WEBHOOK_OUTLOOK_URL, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(payload)
        });

        if (respuesta.ok) {
            console.log("Reporte enviado con éxito.");
            if (esManual) alert("¡Reporte procesado e informado exitosamente!");
            setTimeout(() => { sincronizarBaseDesdePowerAutomate(); }, 3000);
        } else {
            console.error("Error enviando:", respuesta.status);
            if (esManual) alert("Error al conectar con Power Automate.");
        }

    } catch (err) {
        console.error("Error generación PDF:", err);
        vistaReportes.style.display = estiloDisplayOriginal;
        vistaReportes.style.position = '';
        vistaReportes.style.left = '';
        vistaReportes.style.top = '';
        vistaReportes.style.width = '';
        vistaReportes.style.visibility = '';
    }
}

function descargarReportePDF() {
    const elemento = document.getElementById('area-reporte-pdf');
    if (!elemento) return;

    const hoyStr = new Date().toLocaleDateString('es-PE') + ' ' + new Date().toLocaleTimeString('es-PE', {hour: '2-digit', minute:'2-digit'});
    const lblPdfFecha = document.getElementById('pdf-fecha-generacion');
    if (lblPdfFecha) lblPdfFecha.innerText = hoyStr;

    const opciones = {
        margin: [8, 8, 8, 8],
        filename: `Informe_Inspeccion_Carbones_UNACEM_${new Date().toISOString().split('T')[0]}.pdf`,
        image: { type: 'jpeg', quality: 0.65 },
        html2canvas: { scale: 1.2, useCORS: true, logging: false, removeContainer: true },
        jsPDF: { unit: 'mm', format: 'a4', orientation: 'landscape', compress: true },
        pagebreak: { mode: ['css', 'legacy'] }
    };

    const botonesAccion = elemento.querySelector('.no-print');
    if (botonesAccion) botonesAccion.style.display = 'none';

    html2pdf().set(opciones).from(elemento).save().then(() => {
        if (botonesAccion) botonesAccion.style.display = 'flex';
    }).catch(err => {
        console.error("Error generando PDF:", err);
        if (botonesAccion) botonesAccion.style.display = 'flex';
        alert("Ocurrió un error al generar el PDF.");
    });
}

function descargarReporteExcel() {
    if (!datosReporteParaExportar || datosReporteParaExportar.length === 0) {
        alert("No hay datos de reportes para exportar actualmente.");
        return;
    }

    const filasExportar = [];
    datosReporteParaExportar.forEach(m => {
        m.carbones.forEach(c => {
            filasExportar.push({
                "Motor": m.motor,
                "Ubicación Técnica": m.ubTecnica,
                "Planta": m.ubPlanta,
                "Contratista": m.contratista,
                "Estado Motor": m.estadoPredominante,
                "Elemento / Componente": c.carbon,
                "Medida Anterior": c.medidaAnterior,
                "Medida Actual": c.medidaActual,
                "Límite de Cambio": c.limiteCambio,
                "% Vida Útil": `${c.pct}%`,
                "Variación Semanal": c.desgasteSemana,
                "Fecha Estimada de Cambio": c.fechaCambioEstimada,
                "Condición / Diagnóstico": c.motivo,
                "Recomendación": "Verificar en campo según diagnóstico."
            });
        });

        m.temperaturas.forEach(t => {
            filasExportar.push({
                "Motor": m.motor,
                "Ubicación Técnica": m.ubTecnica,
                "Planta": m.ubPlanta,
                "Contratista": m.contratista,
                "Estado Motor": m.estadoPredominante,
                "Elemento / Componente": t.componente,
                "Medida Anterior": "N/A",
                "Medida Actual": t.temp,
                "Límite de Cambio": "N/A",
                "% Vida Útil": "N/A",
                "Variación Semanal": "N/A",
                "Fecha Estimada de Cambio": "N/A",
                "Condición / Diagnóstico": t.condicion,
                "Recomendación": t.recomendacion
            });
        });
    });

    const worksheet = XLSX.utils.json_to_sheet(filasExportar);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Reporte Eventos UNACEM");
    XLSX.writeFile(workbook, `Reporte_Inspeccion_Carbones_UNACEM_${new Date().toISOString().split('T')[0]}.xlsx`);
}

function limpiarTablasYGraficos() {
    if (chartDesgasteInstance) chartDesgasteInstance.destroy();
    if (chartPredictivoInstance) chartPredictivoInstance.destroy();
    for (let id in miniChartsInstances) { if (miniChartsInstances[id]) miniChartsInstances[id].destroy(); }
    document.getElementById('contenedor-carbones').querySelectorAll('.carbon-node').forEach(n => n.remove());
    document.getElementById('tbody-carbones').innerHTML = '<tr><td colspan="9" style="text-align:center; color:#94a3b8;">Sin inspecciones registradas en este rango</td></tr>';
    document.getElementById('val-corriente').innerText = '-- A';
    document.getElementById('val-temp-k').innerText = '-- °C';
    document.getElementById('val-temp-m').innerText = '-- °C';
    document.getElementById('val-temp-l').innerText = '-- °C';
    document.getElementById('val-temp-carcaza').innerText = '-- °C';
    document.getElementById('val-temp-llibre').innerText = '-- °C';
    document.getElementById('val-temp-lacople').innerText = '-- °C';
}
