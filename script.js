document.addEventListener('DOMContentLoaded', () => {
    const fileInput = document.getElementById('csvFileInput');
    const convertBtn = document.getElementById('convertBtn');
    const previewBtn = document.getElementById('previewBtn');
    const downloadPreviewBtn = document.getElementById('downloadPreviewBtn');
    
    const statusMessage = document.getElementById('statusMessage');
    const previewContainer = document.getElementById('previewContainer');
    const previewTable = document.getElementById('previewTable');

    // Variables globales para mantener el archivo Excel en memoria
    let workbookEnMemoria = null;
    let nombreArchivoEnMemoria = "";

    // Eventos de los botones
    convertBtn.addEventListener('click', () => procesarArchivo('descargar'));
    previewBtn.addEventListener('click', () => procesarArchivo('previa'));
    
    // Botón que aparece debajo de la vista previa
    downloadPreviewBtn.addEventListener('click', () => {
        if (workbookEnMemoria && nombreArchivoEnMemoria) {
            XLSX.writeFile(workbookEnMemoria, nombreArchivoEnMemoria);
            showMessage('¡Archivo Excel descargado con éxito!', 'success');
        }
    });

    function procesarArchivo(accion) {
        const file = fileInput.files[0];

        if (!file) {
            showMessage('Por favor, selecciona un archivo CSV primero.', 'error');
            return;
        }

        if (!file.name.toLowerCase().endsWith('.csv')) {
            showMessage('El archivo debe tener extensión .csv', 'error');
            return;
        }

        showMessage('Procesando archivo...', 'success');
        bloquearBotones(true);
        previewContainer.classList.add('hidden'); // Ocultar tabla si existía de antes

        const reader = new FileReader();
        
        reader.onload = function(e) {
            const csvData = e.target.result;

            Papa.parse(csvData, {
                skipEmptyLines: true,
                complete: function(results) {
                    if (results.errors.length > 0 && results.data.length === 0) {
                        showMessage('Hubo un error al leer el CSV.', 'error');
                        bloquearBotones(false);
                        return;
                    }

                    try {
                        // 1. Crear el objeto de Excel y guardarlo en memoria
                        const worksheet = XLSX.utils.aoa_to_sheet(results.data);
                        workbookEnMemoria = XLSX.utils.book_new();
                        XLSX.utils.book_append_sheet(workbookEnMemoria, worksheet, "Datos");
                        
                        const originalName = file.name.replace(/\.[^/.]+$/, "");
                        nombreArchivoEnMemoria = originalName + "_Limpio.xlsx";

                        // 2. Ejecutar la acción solicitada
                        if (accion === 'descargar') {
                            XLSX.writeFile(workbookEnMemoria, nombreArchivoEnMemoria);
                            showMessage('¡Archivo Excel generado y descargado con éxito!', 'success');
                        } else if (accion === 'previa') {
                            generarTablaHTML(results.data);
                            previewContainer.classList.remove('hidden');
                            showMessage('Vista previa lista. Revisa los datos y haz clic en Descargar.', 'success');
                        }
                    } catch (error) {
                        showMessage('Error al procesar los datos.', 'error');
                        console.error(error);
                    } finally {
                        bloquearBotones(false);
                    }
                }
            });
        };

        reader.onerror = function() {
            showMessage('Error al leer el archivo desde el equipo.', 'error');
            bloquearBotones(false);
        };

        reader.readAsText(file, 'UTF-8'); 
    }

    // Función para construir la tabla en la página web
    function generarTablaHTML(data) {
        previewTable.innerHTML = ''; // Limpiar tabla anterior
        if (data.length === 0) return;

        // Limitar a 100 filas para no colapsar el navegador
        const limiteFilas = Math.min(data.length, 100); 

        // Crear Encabezados (primera fila del CSV)
        const thead = document.createElement('thead');
        const headerRow = document.createElement('tr');
        data[0].forEach(colTexto => {
            const th = document.createElement('th');
            th.textContent = colTexto;
            headerRow.appendChild(th);
        });
        thead.appendChild(headerRow);
        previewTable.appendChild(thead);

        // Crear Cuerpo (resto de los datos)
        const tbody = document.createElement('tbody');
        for (let i = 1; i < limiteFilas; i++) {
            const row = document.createElement('tr');
            data[i].forEach(celdaTexto => {
                const td = document.createElement('td');
                td.textContent = celdaTexto;
                row.appendChild(td);
            });
            tbody.appendChild(row);
        }
        previewTable.appendChild(tbody);

        // Aviso si hay más de 100 filas
        if (data.length > 100) {
            const caption = document.createElement('caption');
            caption.textContent = `Mostrando una vista previa de las primeras 100 filas (de ${data.length} totales). El archivo Excel se descargará completo.`;
            previewTable.appendChild(caption);
        }
    }

    // Funciones auxiliares
    function showMessage(text, type) {
        statusMessage.textContent = text;
        statusMessage.className = `status ${type}`;
    }

    function bloquearBotones(estado) {
        convertBtn.disabled = estado;
        previewBtn.disabled = estado;
    }
});