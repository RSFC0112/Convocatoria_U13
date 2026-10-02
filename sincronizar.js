const fs = require('fs');
const path = require('path');
const { createClient } = require('@supabase/supabase-js');
const XLSX = require('xlsx');
const axios = require('axios');

const SUPABASE_URL = 'https://fjbbrzhqlvbkliskslqs.supabase.co';
const SUPABASE_KEY = 'sb_publishable_QRpv4Rs_fIJ8kXvTegR25w_t2SRSExp';
const supabase = createClient(SUPABASE_URL, SUPABASE_KEY);
const htmlIndex = fs.readFileSync(path.join(__dirname, '..', 'index.html'), 'utf8');
const normalizarNombre = nombre => String(nombre || '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, ' ')
    .trim();
const jugadoresPorNombre = new Map(
    Array.from(htmlIndex.matchAll(/\{\s*id:\s*(\d+),[^\r\n]*?\bnombre:\s*"([^"]+)",\s*posicion:\s*"([^"]+)"/g),
        ([, id, nombre, posicion]) => [normalizarNombre(nombre), { id: Number(id), posicion }])
);

if (jugadoresPorNombre.size === 0) {
    throw new Error('No se encontraron posiciones de jugadores en index.html.');
}

function parsearFechaExcel(valorExcel) {
    if (valorExcel === undefined || valorExcel === null || valorExcel === '') return null;
    if (valorExcel instanceof Date) return Number.isNaN(valorExcel.getTime()) ? null : valorExcel;

    const num = Number(valorExcel);
    if (Number.isFinite(num) && num > 1000) {
        const utcDays = Math.floor(num - 25569);
        const fechaObj = new Date(utcDays * 86400 * 1000);
        return Number.isNaN(fechaObj.getTime()) ? null : fechaObj;
    }

    const texto = String(valorExcel).trim();
    let partes = texto.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
    if (partes) {
        return new Date(Date.UTC(Number(partes[1]), Number(partes[2]) - 1, Number(partes[3])));
    }

    partes = texto.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
    if (partes) {
        const dia = Number(partes[1]);
        const mes = Number(partes[2]);
        const anio = Number(partes[3]);
        const fecha = new Date(Date.UTC(anio, mes - 1, dia));
        if (fecha.getUTCMonth() !== mes - 1 || fecha.getUTCDate() !== dia) return null;
        return fecha;
    }

    const fechaParsed = new Date(texto);
    return Number.isNaN(fechaParsed.getTime()) ? null : fechaParsed;
}

function formatearFechaDDMMYYYY(valorExcel) {
    if (valorExcel === undefined || valorExcel === null || valorExcel === '') return 'NA';
    const fechaObj = parsearFechaExcel(valorExcel);
    if (!fechaObj) return String(valorExcel).trim();
    const anio = fechaObj.getUTCFullYear();
    const mes = String(fechaObj.getUTCMonth() + 1).padStart(2, '0');
    const dia = String(fechaObj.getUTCDate()).padStart(2, '0');
    return `${dia}/${mes}/${anio}`;
}

function obtenerDiaUTC(fecha) {
    return Date.UTC(fecha.getUTCFullYear(), fecha.getUTCMonth(), fecha.getUTCDate());
}

function esFechaIgualOPosterior(fecha, fechaLimite) {
    return obtenerDiaUTC(fecha) >= obtenerDiaUTC(fechaLimite);
}

async function sincronizarU13() {
    try {
        const urlExcelEnLinea = 'https://docs.google.com/spreadsheets/d/1gpY4TcpxmBebSk9Popp5IK7tYVriTTT-/export?format=xlsx';

        console.log('Descargando datos en tiempo real desde Google Sheets...');
        const respuesta = await axios.get(urlExcelEnLinea, { responseType: 'arraybuffer' });
        const workbook = XLSX.read(respuesta.data, { type: 'buffer' });

        const nombreHoja = 'ACTIVOS 2026';
        if (!workbook.SheetNames.includes(nombreHoja)) {
            console.error(`Error: No se encontró la pestaña "${nombreHoja}".`);
            return;
        }
        console.log(`Leyendo la pestaña en línea: "${nombreHoja}"`);

        const worksheet = workbook.Sheets[nombreHoja];
        const filas = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

        if (filas.length === 0) {
            console.log("La hoja de cálculo está vacía.");
            return;
        }

        const filaFechas = filas[6] || [];
        const normalizarEncabezado = valor => String(valor || '')
            .normalize('NFD')
            .replace(/[\u0300-\u036f]/g, '')
            .toLowerCase()
            .replace(/[^a-z0-9]+/g, ' ')
            .trim();
        let indiceNombre = -1;
        let indiceApellido1 = -1;
        let indiceDorsal = -1;
        let indiceFechaNacimiento = -1;
        let indiceFechaInscripcion = -1;
        let indiceEntrenamientosCumplidos = -1;
        let indiceEntrenamientosOfrecidos = -1;
        let indicePartidosJugados = -1;
        let indicePartidosProgramados = -1;

        for (const fila of filas.slice(0, 10)) {
            (fila || []).forEach((valor, indice) => {
                const encabezado = normalizarEncabezado(valor);
                if (/\bnombres?\b/.test(encabezado) && indiceNombre === -1) {
                    indiceNombre = indice;
                }
                if ((/\bapellido\s*(1|uno)\b/.test(encabezado)
                    || /\bprimer apellido\b/.test(encabezado)
                    || /\bapellido paterno\b/.test(encabezado))
                    && indiceApellido1 === -1) {
                    indiceApellido1 = indice;
                }
                if (encabezado === 'dorsal' && indiceDorsal === -1) {
                    indiceDorsal = indice;
                }
                if (/^fecha nac\b/.test(encabezado) && indiceFechaNacimiento === -1) {
                    indiceFechaNacimiento = indice;
                }
                if ((/\bfecha\b.*\b(inscripcion|ingreso|registro)\b/.test(encabezado)
                    || /\b(inscripcion|ingreso|registro)\b.*\bfecha\b/.test(encabezado))
                    && indiceFechaInscripcion === -1) {
                    indiceFechaInscripcion = indice;
                }
                if (encabezado === 'anual cumplido' && indiceEntrenamientosCumplidos === -1) {
                    indiceEntrenamientosCumplidos = indice;
                }
                if (encabezado === 'anual ofrecido' && indiceEntrenamientosOfrecidos === -1) {
                    indiceEntrenamientosOfrecidos = indice;
                }
                if (encabezado === 'jugados' && indicePartidosJugados === -1) {
                    indicePartidosJugados = indice;
                }
                if (encabezado === 'programados' && indicePartidosProgramados === -1) {
                    indicePartidosProgramados = indice;
                }
            });
            if (indiceNombre !== -1 && indiceApellido1 !== -1 && indiceDorsal !== -1
                && indiceFechaNacimiento !== -1 && indiceFechaInscripcion !== -1
                && indiceEntrenamientosCumplidos !== -1 && indiceEntrenamientosOfrecidos !== -1
                && indicePartidosJugados !== -1 && indicePartidosProgramados !== -1) break;
        }

        if ([indiceNombre, indiceApellido1, indiceDorsal, indiceFechaNacimiento, indiceFechaInscripcion, indiceEntrenamientosCumplidos,
            indiceEntrenamientosOfrecidos, indicePartidosJugados, indicePartidosProgramados].includes(-1)) {
            console.error('Error: Faltan encabezados de identidad, dorsal, fecha de nacimiento, inscripción o estadísticas anuales en las primeras filas.');
            return;
        }
        console.log(`Columnas detectadas: nombre ${indiceNombre + 1}, primer apellido ${indiceApellido1 + 1}, dorsal ${indiceDorsal + 1}, fecha de nacimiento ${indiceFechaNacimiento + 1}, fecha de inscripción ${indiceFechaInscripcion + 1}, entrenamientos ${indiceEntrenamientosCumplidos + 1}/${indiceEntrenamientosOfrecidos + 1}, partidos ${indicePartidosJugados + 1}/${indicePartidosProgramados + 1}`);

        const indicesU13 = [];
        for (let i = 0; i < filas.length; i++) {
            const fila = filas[i];
            if (!fila) continue;
            const nombre = String(fila[indiceNombre] || '').trim();
            const esU13 = fila.some(celda =>
                /\bU\s*-?\s*13\b/i.test(String(celda || '').normalize('NFKC'))
            );
            if (esU13 && nombre !== '') {
                indicesU13.push(i);
            }
        }

        console.log(`Jugadores U13 detectados: ${indicesU13.length}`);

        const jugadoresSinPosicion = indicesU13
            .map(rowIdx => {
                const fila = filas[rowIdx];
                return `${String(fila[indiceNombre]).trim()} ${String(fila[indiceApellido1] || '').trim()}`.trim();
            })
            .filter(nombre => !jugadoresPorNombre.has(normalizarNombre(nombre)));
        if (jugadoresSinPosicion.length > 0) {
            console.error(`Error: No se encontró posición en index.html para: ${jugadoresSinPosicion.join(', ')}.`);
            return;
        }

        if (indicesU13.length === 0) {
            console.error("Error: No se encontraron jugadores U13 en la hoja en línea.");
            return;
        }

        const encabezadosSeccion = filas[3] || [];
        const indiceAsistencia = encabezadosSeccion.findIndex(valor =>
            normalizarEncabezado(valor) === 'asistencia 2026'
        );
        const indiceConvocatorias = encabezadosSeccion.findIndex(valor =>
            normalizarEncabezado(valor) === 'convocatorias 2026'
        );
        if (indiceAsistencia === -1 || indiceConvocatorias <= indiceAsistencia) {
            console.error('Error: No se encontraron los encabezados de asistencia y convocatorias.');
            return;
        }

        const columnasEntrenamientos = [];
        for (let idx = indiceAsistencia + 1; idx < indiceConvocatorias; idx++) {
            const fecha = parsearFechaExcel(filaFechas[idx]);
            if (fecha) columnasEntrenamientos.push({ colIndex: idx, fecha });
        }
        columnasEntrenamientos.sort((a, b) => a.fecha - b.fecha);

        const columnasPartidos = [];
        for (let idx = indiceConvocatorias + 1; idx < filaFechas.length; idx++) {
            const fecha = parsearFechaExcel(filaFechas[idx]);
            if (fecha) columnasPartidos.push({ colIndex: idx, fecha });
        }
        columnasPartidos.sort((a, b) => a.fecha - b.fecha);

        const obtenerMarcaEntrenamiento = valor => {
            if (valor === undefined || valor === null || String(valor).trim() === '') return null;
            const numero = Number(valor);
            return [0, 0.5, 1].includes(numero) ? numero : null;
        };
        const obtenerMarcaPartido = valor => {
            const marca = String(valor ?? '').trim().toUpperCase();
            return ['X', 'P(X)', '1', '1.0'].includes(marca) ? marca : null;
        };

        let contadorU13 = 0;

        for (const rowIdx of indicesU13) {
            const fila = filas[rowIdx];
            const nombre = fila[indiceNombre];
            const apellido1 = fila[indiceApellido1];

            if (!nombre) continue;

            contadorU13++;
            const nombreCompleto = `${String(nombre).trim()} ${String(apellido1 || '').trim()}`;
            const fechaInscripcion = parsearFechaExcel(fila[indiceFechaInscripcion]);
            const limitarPorInscripcion = fechaInscripcion
                && fechaInscripcion.getUTCFullYear() >= 2026;
            const fechaPermitida = fecha => !limitarPorInscripcion || esFechaIgualOPosterior(fecha, fechaInscripcion);
            const esDiaInscripcion = fecha => limitarPorInscripcion
                && obtenerDiaUTC(fecha) === obtenerDiaUTC(fechaInscripcion);
            const obtenerValorAnual = indice => {
                const valor = fila[indice];
                return valor !== undefined && valor !== null && String(valor).trim() !== '' ? valor : 0;
            };
            const entrenamientosDelJugador = columnasEntrenamientos.filter(item =>
                fechaPermitida(item.fecha)
                && (obtenerMarcaEntrenamiento(fila[item.colIndex]) !== null || esDiaInscripcion(item.fecha))
            );
            const entrenamientosAsistidos = obtenerValorAnual(indiceEntrenamientosCumplidos);
            const totalEntrenamientos = obtenerValorAnual(indiceEntrenamientosOfrecidos);

            const partidosDelJugador = columnasPartidos.filter(item =>
                fechaPermitida(item.fecha)
                && obtenerMarcaPartido(fila[item.colIndex]) !== null
            );
            const partidosAsistidos = limitarPorInscripcion
                ? partidosDelJugador.filter(item =>
                    obtenerMarcaPartido(fila[item.colIndex]) === '1'
                    || obtenerMarcaPartido(fila[item.colIndex]) === '1.0'
                ).length
                : obtenerValorAnual(indicePartidosJugados);
            const totalPartidos = limitarPorInscripcion
                ? partidosDelJugador.length
                : obtenerValorAnual(indicePartidosProgramados);

            const textoEntrenamientos = `${entrenamientosAsistidos}/${totalEntrenamientos}`;
            const textoPartidos = `${partidosAsistidos}/${totalPartidos}`;

            const entrenamientosParaUltimasFechas = entrenamientosDelJugador.slice(-5);
            const ultimasFechasData = entrenamientosParaUltimasFechas.map(item => {
                const dia = String(item.fecha.getUTCDate()).padStart(2, '0');
                const mes = ['Ene', 'Feb', 'Mar', 'Abr', 'May', 'Jun', 'Jul', 'Ago', 'Sep', 'Oct', 'Nov', 'Dic'][item.fecha.getUTCMonth()];
                const marca = obtenerMarcaEntrenamiento(fila[item.colIndex]);

                return {
                    fecha: `${dia}-${mes}`,
                    asistio: marca === null && esDiaInscripcion(item.fecha) ? 0 : marca
                };
            });
            while (ultimasFechasData.length < 5) {
                ultimasFechasData.unshift({ fecha: 'NA', asistio: 'NA' });
            }

            const fechaInscripcionLimpia = formatearFechaDDMMYYYY(fila[indiceFechaInscripcion]);
            const jugadorPlantilla = jugadoresPorNombre.get(normalizarNombre(nombreCompleto));
            const valorDorsal = fila[indiceDorsal];
            const dorsal = valorDorsal !== undefined && valorDorsal !== null && String(valorDorsal).trim() !== ''
                && Number.isFinite(Number(valorDorsal)) ? Number(valorDorsal) : null;
            const fechaNacimiento = parsearFechaExcel(fila[indiceFechaNacimiento]);

            const datosJugador = {
                id_jugador: jugadorPlantilla.id,
                nombre: nombreCompleto,
                entrenamientos: textoEntrenamientos,
                partidos: textoPartidos,
                fecha_inscripcion: fechaInscripcionLimpia,
                ultimas_fechas: ultimasFechasData,
                dorsal,
                posicion: jugadorPlantilla.posicion,
                anio_nacimiento: fechaNacimiento ? fechaNacimiento.getUTCFullYear() : null
            };

            const tablaEstadisticas = supabase.from('Estadísticas_U13');
            const { data: registrosExistentes, error: errorConsulta } = await tablaEstadisticas
                .select('id_jugador')
                .eq('nombre', nombreCompleto);

            if (errorConsulta) {
                console.error(`  -> Error al consultar ${datosJugador.nombre}:`, errorConsulta.message);
                continue;
            }
            if (registrosExistentes.length > 1) {
                console.error(`  -> Error: hay varios registros para ${datosJugador.nombre}; no se modificaron.`);
                continue;
            }

            const { error } = registrosExistentes.length === 1
                ? await tablaEstadisticas.update(datosJugador).eq('nombre', nombreCompleto)
                : await tablaEstadisticas.upsert(datosJugador, { onConflict: 'id_jugador' });

            if (error) {
                console.error(`  -> Error al subir a ${datosJugador.nombre}:`, error.message);
            } else {
                console.log(`  -> Sincronizado U13: ${datosJugador.nombre} | Posición: ${datosJugador.posicion} | Entrenamientos: ${textoEntrenamientos} | Partidos: ${textoPartidos}`);
            }
        }

        console.log(`¡Sincronización completa! Se procesaron ${contadorU13} jugadores de la U13 en línea.`);

    } catch (error) {
        console.error('Ocurrió un error general descargando de Google Sheets:', error.message);
    }
}

sincronizarU13();
