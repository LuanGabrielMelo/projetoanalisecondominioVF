// === VARIÁVEIS GLOBAIS ===
let originalData = [];
let processedData = [];
let monthlyAnalysis = {};
let waterChart = null;
let energyChart = null;
let comparisonChart = null;
let monthlyTotalChart = null;
let monthlyTrendsChart = null;
let monthlyVariationChart = null;
let isFileValid = false;

// === INICIALIZAÇÃO ===
document.addEventListener("DOMContentLoaded", function () {
    setupEventListeners();
    setDefaultDate();
    showInitialMessage();
});

function setupEventListeners() {
    document
        .getElementById("fileInput")
        .addEventListener("change", handleFileUpload);
    document
        .getElementById("dataForm")
        .addEventListener("submit", handleFormSubmit);
    document
        .getElementById("downloadBtn")
        .addEventListener("click", downloadExcel);
}

function setDefaultDate() {
    const today = new Date().toISOString().split("T")[0];
    document.getElementById("date").value = today;
}

function showInitialMessage() {
    showAlert(
        "Sistema pronto! Você pode carregar um arquivo Excel existente ou inserir dados diretamente.",
        "success"
    );
}

// === FUNÇÕES DE UTILIDADE ===
function showAlert(message, type = "success") {
    const alertDiv = document.createElement("div");
    alertDiv.className = `alert alert-${type}`;
    alertDiv.textContent = message;

    const fileStatus = document.getElementById("fileStatus");
    fileStatus.innerHTML = "";
    fileStatus.appendChild(alertDiv);

    setTimeout(() => {
        if (alertDiv.parentNode) {
            alertDiv.remove();
        }
    }, 5000);
}

function calcularMediaDiaria(valores) {
    if (!valores || valores.length === 0) {
        return {
            media: 0,
            houveNegativo: false,
            valoresUtilizados: 0,
            somaTotal: 0,
        };
    }

    const valoresValidos = valores
        .filter(
            (v) =>
                !isNaN(v) &&
                v !== null &&
                v !== undefined &&
                v !== "" &&
                v !== 0
        )
        .map((v) => parseFloat(v));

    if (valoresValidos.length === 0) {
        return {
            media: 0,
            houveNegativo: false,
            valoresUtilizados: 0,
            somaTotal: 0,
        };
    }

    const houveNegativo = valoresValidos.some((v) => v < 0);
    const soma = valoresValidos.reduce((acc, val) => acc + val, 0);
    const media = soma / valoresValidos.length;

    return {
        media: media,
        houveNegativo: houveNegativo,
        valoresUtilizados: valoresValidos.length,
        somaTotal: soma,
    };
}

// === MANIPULAÇÃO DE ARQUIVOS ===
async function handleFileUpload(event) {
    const file = event.target.files[0];
    if (!file) return;

    try {
        showAlert("Carregando arquivo...", "success");

        const arrayBuffer = await file.arrayBuffer();
        const workbook = XLSX.read(arrayBuffer, {
            cellDates: true,
            cellNF: false,
            cellText: false,
        });

        // Detectar o tipo de arquivo: original ou gerado pelo sistema
        const isGeneratedFile =
            workbook.SheetNames.includes("Dados de Consumo");

        let dadosProcessados = [];

        if (isGeneratedFile) {
            // Arquivo gerado pelo sistema - formato unificado
            dadosProcessados = parseGeneratedFile(workbook);
            showAlert(
                "Arquivo gerado pelo sistema detectado e carregado com sucesso!"
            );
        } else {
            // Arquivo original - formato separado
            dadosProcessados = parseOriginalFile(workbook);
            const dadosCoelba = dadosProcessados.filter(
                (d) => d.source === "coelba"
            );
            const dadosEmbasa = dadosProcessados.filter(
                (d) => d.source === "embasa"
            );
            showAlert(
                `Arquivo original carregado com sucesso! ${dadosCoelba.length} registros Coelba e ${dadosEmbasa.length} registros Embasa encontrados.`
            );
        }

        if (dadosProcessados.length === 0) {
            throw new Error("Nenhum dado válido foi encontrado nas planilhas.");
        }

        originalData = dadosProcessados;
        processedData = [...originalData];
        isFileValid = true;

        monthlyAnalysis = performMonthlyAnalysis(processedData);
        updateDisplay();
    } catch (error) {
        console.error("Erro ao processar arquivo:", error);
        showAlert("Erro ao carregar arquivo: " + error.message, "error");
    }
}

function parseGeneratedFile(workbook) {
    const dadosSheet = workbook.Sheets["Dados de Consumo"];
    const jsonData = XLSX.utils.sheet_to_json(dadosSheet, {
        header: 1,
        defval: "",
        raw: false,
        dateNF: "yyyy-mm-dd",
    });

    const data = [];
    // Pular header (linha 0)
    for (let i = 1; i < jsonData.length; i++) {
        const row = jsonData[i];
        if (
            !row ||
            row.length === 0 ||
            !row.some(
                (cell) => cell !== "" && cell !== null && cell !== undefined
            )
        ) {
            continue;
        }

        const dateStr = row[0]; // Data
        const origem = row[1]; // Origem
        const consumoTotal = parseFloat(row[2]) || 0; // Consumo Total
        const difDia = parseFloat(row[3]) || 0; // Dif_dia

        const parsedDate = parseDate(dateStr);
        if (!parsedDate) continue;

        // Determinar source baseado na origem
        let source;
        if (
            (origem && origem.toLowerCase().includes("embasa")) ||
            (origem && origem.toLowerCase().includes("água"))
        ) {
            source = "embasa";
        } else if (
            (origem && origem.toLowerCase().includes("coelba")) ||
            (origem && origem.toLowerCase().includes("energia"))
        ) {
            source = "coelba";
        } else {
            continue; // Pular se não conseguir identificar
        }

        data.push({
            date: parsedDate,
            source: source,
            consumption: consumoTotal,
            difDia: difDia,
        });
    }

    return data.sort((a, b) => new Date(a.date) - new Date(b.date));
}

function parseOriginalFile(workbook) {
    let coelbaSheet = null;
    let embasaSheet = null;

    workbook.SheetNames.forEach((sheetName) => {
        const lowerName = sheetName.toLowerCase();
        if (
            lowerName.includes("coelba") ||
            lowerName.includes("energia") ||
            lowerName.includes("luz")
        ) {
            coelbaSheet = sheetName;
        } else if (
            lowerName.includes("embasa") ||
            lowerName.includes("agua") ||
            lowerName.includes("water")
        ) {
            embasaSheet = sheetName;
        }
    });

    if (!coelbaSheet && !embasaSheet && workbook.SheetNames.length >= 2) {
        coelbaSheet = workbook.SheetNames[0];
        embasaSheet = workbook.SheetNames[1];
    } else if (!coelbaSheet || !embasaSheet) {
        throw new Error(
            "Não foi possível identificar as planilhas da Coelba e Embasa."
        );
    }

    const coelbaWorksheet = workbook.Sheets[coelbaSheet];
    const coelbaJsonData = XLSX.utils.sheet_to_json(coelbaWorksheet, {
        header: 1,
        defval: "",
        raw: false,
        dateNF: "yyyy-mm-dd",
    });

    const embasaWorksheet = workbook.Sheets[embasaSheet];
    const embasaJsonData = XLSX.utils.sheet_to_json(embasaWorksheet, {
        header: 1,
        defval: "",
        raw: false,
        dateNF: "yyyy-mm-dd",
    });

    const dadosCoelba = parseExcelData(coelbaJsonData, "coelba");
    const dadosEmbasa = parseExcelData(embasaJsonData, "embasa");
    return [...dadosCoelba, ...dadosEmbasa];
}

function parseExcelData(jsonData, forceSource = null) {
    const data = [];
    let startRow = 0;
    let difDiaIndex = -1;
    let dateIndex = 0;

    if (jsonData[0] && Array.isArray(jsonData[0])) {
        jsonData[0].forEach((header, index) => {
            const headerStr = String(header || "")
                .toLowerCase()
                .trim();
            if (headerStr === "dia" || headerStr.includes("data")) {
                dateIndex = index;
                startRow = 1;
            } else if (
                headerStr === "dif_dia" ||
                headerStr === "dif.dia" ||
                headerStr.includes("dif")
            ) {
                difDiaIndex = index;
                startRow = 1;
            }
        });
    }

    if (
        difDiaIndex === -1 &&
        jsonData.length > 1 &&
        jsonData[1] &&
        jsonData[1].length >= 3
    ) {
        difDiaIndex = 2;
    }

    for (let i = startRow; i < jsonData.length; i++) {
        const row = jsonData[i];
        if (
            !row ||
            row.length === 0 ||
            !row.some(
                (cell) => cell !== "" && cell !== null && cell !== undefined
            )
        ) {
            continue;
        }

        const parsedDate = parseDate(row[dateIndex]);
        const consumption = parseFloat(row[1]) || 0;
        const difDia =
            difDiaIndex !== -1 &&
            row[difDiaIndex] !== "" &&
            row[difDiaIndex] !== undefined
                ? parseFloat(row[difDiaIndex]) || 0
                : 0;

        if (!parsedDate || isNaN(consumption)) {
            continue;
        }

        data.push({
            date: parsedDate,
            source: forceSource,
            consumption: consumption,
            difDia: difDia,
        });
    }

    return data.sort((a, b) => new Date(a.date) - new Date(b.date));
}

function parseDate(dateValue) {
    if (!dateValue) return null;
    if (dateValue instanceof Date) return dateValue;

    if (typeof dateValue === "number") {
        if (dateValue > 25569) {
            return new Date((dateValue - 25569) * 86400 * 1000);
        } else {
            const excelEpoch = new Date(1900, 0, 1);
            return new Date(
                excelEpoch.getTime() + (dateValue - 2) * 24 * 60 * 60 * 1000
            );
        }
    }

    if (typeof dateValue === "string") {
        // Tentar primeiro o formato brasileiro dd/mm/yyyy
        const brazilianFormat = dateValue.match(
            /(\d{1,2})\/(\d{1,2})\/(\d{4})/
        );
        if (brazilianFormat) {
            const [, day, month, year] = brazilianFormat;
            const date = new Date(year, month - 1, day);
            if (!isNaN(date.getTime())) {
                return date;
            }
        }

        const formats = [
            dateValue,
            dateValue.replace(/\//g, "-"),
            dateValue.replace(/\./g, "-"),
            dateValue.split("/").reverse().join("-"),
            dateValue.split(".").reverse().join("-"),
        ];

        for (const format of formats) {
            const parsed = new Date(format);
            if (!isNaN(parsed.getTime())) {
                return parsed;
            }
        }
    }

    return null;
}

// === MANIPULAÇÃO DO FORMULÁRIO ===
function handleFormSubmit(event) {
    event.preventDefault();

    const formData = new FormData(event.target);
    const dateValue = formData.get("date");
    const sourceValue = formData.get("source");
    const consumptionValue = formData.get("consumption");

    if (!dateValue || !sourceValue || !consumptionValue) {
        showAlert("Por favor, preencha todos os campos.", "error");
        return;
    }

    const consumption = parseFloat(consumptionValue);
    if (isNaN(consumption)) {
        showAlert(
            "Por favor, insira um valor numérico válido para o consumo.",
            "error"
        );
        return;
    }

    const newRecord = {
        date: new Date(dateValue),
        source: sourceValue,
        consumption: consumption,
        difDia: consumption,
    };

    if (!processedData) {
        processedData = [];
    }

    processedData.push(newRecord);
    processedData.sort((a, b) => new Date(a.date) - new Date(b.date));

    monthlyAnalysis = performMonthlyAnalysis(processedData);

    showAlert("Consumo adicionado com sucesso!");
    updateDisplay();
    event.target.reset();
    setDefaultDate();

    document.getElementById("downloadSection").style.display = "block";
}

// === ANÁLISE MENSAL ===
function performMonthlyAnalysis(data) {
    if (!data || data.length === 0) return {};

    const embasaData = data.filter((r) => r.source === "embasa");
    const coelbaData = data.filter((r) => r.source === "coelba");

    const coelbaMonthly = groupDataByMonthDetailed(coelbaData, "Energia");
    const embasaMonthly = groupDataByMonthDetailed(embasaData, "Água");

    return { energia: coelbaMonthly, agua: embasaMonthly };
}

function groupDataByMonthDetailed(data, tipo) {
    const monthlyGroups = {};

    data.forEach((record) => {
        const monthKey = `${record.date.getFullYear()}-${String(
            record.date.getMonth() + 1
        ).padStart(2, "0")}`;

        if (!monthlyGroups[monthKey]) {
            monthlyGroups[monthKey] = {
                mes: record.date.getMonth() + 1,
                ano: record.date.getFullYear(),
                mesAno: monthKey,
                tipo: tipo,
                registros: [],
                totalConsumo: 0,
                totalDifDia: 0,
                mediaDiariaConsumo: 0,
                mediaDiariaDifDia: 0,
                diasComDados: 0,
                diasComDifDia: 0,
                maiorConsumo: 0,
                menorConsumo: Infinity,
                maiorDifDia: 0,
                menorDifDia: Infinity,
                variacaoPercentual: 0,
            };
        }

        const grupo = monthlyGroups[monthKey];
        grupo.registros.push(record);
        grupo.totalConsumo += record.consumption;
        grupo.diasComDados++;

        if (record.difDia > 0) {
            grupo.totalDifDia += record.difDia;
            grupo.diasComDifDia++;
            grupo.maiorDifDia = Math.max(grupo.maiorDifDia, record.difDia);
            grupo.menorDifDia = Math.min(grupo.menorDifDia, record.difDia);
        }

        grupo.maiorConsumo = Math.max(grupo.maiorConsumo, record.consumption);
        grupo.menorConsumo = Math.min(grupo.menorConsumo, record.consumption);
    });

    const sortedKeys = Object.keys(monthlyGroups).sort();
    sortedKeys.forEach((key, index) => {
        const grupo = monthlyGroups[key];
        grupo.mediaDiariaConsumo = grupo.totalConsumo / grupo.diasComDados;
        grupo.mediaDiariaDifDia =
            grupo.diasComDifDia > 0
                ? grupo.totalDifDia / grupo.diasComDifDia
                : 0;

        if (index > 0) {
            const mesAnterior = monthlyGroups[sortedKeys[index - 1]];
            if (mesAnterior.mediaDiariaDifDia > 0) {
                grupo.variacaoPercentual =
                    ((grupo.mediaDiariaDifDia - mesAnterior.mediaDiariaDifDia) /
                        mesAnterior.mediaDiariaDifDia) *
                    100;
            }
        }

        if (grupo.menorConsumo === Infinity) grupo.menorConsumo = 0;
        if (grupo.menorDifDia === Infinity) grupo.menorDifDia = 0;
    });

    return monthlyGroups;
}

// === ATUALIZAÇÃO DA INTERFACE ===
function updateDisplay() {
    if (!processedData || processedData.length === 0) return;

    updateResultsSection();
    updateDataTable();
    updateMonthlyAnalysisSection();
    updateTrendsCharts();
    updateCharts();

    document.getElementById("resultsSection").style.display = "block";
    document.getElementById("dataTableSection").style.display = "block";
    document.getElementById("monthlyAnalysisSection").style.display = "block";
    document.getElementById("trendsSection").style.display = "block";
    document.getElementById("chartsSection").style.display = "block";
    document.getElementById("downloadSection").style.display = "block";
}

function updateResultsSection() {
    const embasaData = processedData.filter((r) => r.source === "embasa");
    const coelbaData = processedData.filter((r) => r.source === "coelba");

    const embasaDifDiaValues = embasaData.map((r) => r.difDia);
    const coelbaDifDiaValues = coelbaData.map((r) => r.difDia);

    const resultadoEmbasaConsumo = calcularMediaDiaria(embasaDifDiaValues);
    const resultadoCoelbaConsumo = calcularMediaDiaria(coelbaDifDiaValues);

    const embasaTotal = embasaData.reduce(
        (sum, r) => sum + (parseFloat(r.consumption) || 0),
        0
    );
    const coelbaTotal = coelbaData.reduce(
        (sum, r) => sum + (parseFloat(r.consumption) || 0),
        0
    );

    const mediaEmbasaConsumo = resultadoEmbasaConsumo.media || 0;
    const mediaCoelbaConsumo = resultadoCoelbaConsumo.media || 0;

    let monthlyInsights = "";
    if (monthlyAnalysis && (monthlyAnalysis.energia || monthlyAnalysis.agua)) {
        const energiaKeys = monthlyAnalysis.energia
            ? Object.keys(monthlyAnalysis.energia).sort()
            : [];
        const aguaKeys = monthlyAnalysis.agua
            ? Object.keys(monthlyAnalysis.agua).sort()
            : [];

        if (energiaKeys.length > 1) {
            const ultimoMesEnergia =
                monthlyAnalysis.energia[energiaKeys[energiaKeys.length - 1]];
            const variacaoEnergia = ultimoMesEnergia.variacaoPercentual;
            monthlyInsights += `<div class="insight-item">Energia: ${
                variacaoEnergia > 0 ? "+" : ""
            }${variacaoEnergia.toFixed(1)}% vs mês anterior</div>`;
        }

        if (aguaKeys.length > 1) {
            const ultimoMesAgua =
                monthlyAnalysis.agua[aguaKeys[aguaKeys.length - 1]];
            const variacaoAgua = ultimoMesAgua.variacaoPercentual;
            monthlyInsights += `<div class="insight-item">Água: ${
                variacaoAgua > 0 ? "+" : ""
            }${variacaoAgua.toFixed(1)}% vs mês anterior</div>`;
        }
    }

    const resultsHTML = `
                <div class="result-card">
                    <h3>Média Diária - Água</h3>
                    <div class="result-value">${mediaEmbasaConsumo.toFixed(
                        2
                    )}</div>
                    <div class="result-unit">m³/dia</div>
                    <div class="result-details">Soma: ${
                        resultadoEmbasaConsumo.somaTotal?.toFixed(2) || "0"
                    } | Valores: ${
        resultadoEmbasaConsumo.valoresUtilizados || 0
    }</div>
                </div>
                <div class="result-card">
                    <h3>Média Diária - Energia</h3>
                    <div class="result-value">${mediaCoelbaConsumo.toFixed(
                        2
                    )}</div>
                    <div class="result-unit">kWh/dia</div>
                    <div class="result-details">Soma: ${
                        resultadoCoelbaConsumo.somaTotal?.toFixed(2) || "0"
                    } | Valores: ${
        resultadoCoelbaConsumo.valoresUtilizados || 0
    }</div>
                </div>
                <div class="result-card">
                    <h3>Total Água</h3>
                    <div class="result-value">${embasaTotal.toFixed(1)}</div>
                    <div class="result-unit">m³</div>
                </div>
                <div class="result-card">
                    <h3>Total Energia</h3>
                    <div class="result-value">${coelbaTotal.toFixed(1)}</div>
                    <div class="result-unit">kWh</div>
                </div>
                ${
                    monthlyInsights
                        ? `<div class="result-card insights-card">
                    <h3>📈 Tendências Mensais</h3>
                    <div class="insights-content">${monthlyInsights}</div>
                </div>`
                        : ""
                }
            `;

    document.getElementById("resultsGrid").innerHTML = resultsHTML;
}

function updateTrendsCharts() {
    if (!monthlyAnalysis || Object.keys(monthlyAnalysis).length === 0) return;

    // Destruir gráficos existentes
    if (monthlyTrendsChart) {
        monthlyTrendsChart.destroy();
        monthlyTrendsChart = null;
    }
    if (monthlyVariationChart) {
        monthlyVariationChart.destroy();
        monthlyVariationChart = null;
    }

    // Preparar dados para gráficos de tendências
    const energiaData = monthlyAnalysis.energia
        ? Object.values(monthlyAnalysis.energia)
        : [];
    const aguaData = monthlyAnalysis.agua
        ? Object.values(monthlyAnalysis.agua)
        : [];

    // Combinar e ordenar por data
    const allMonthlyData = [];
    energiaData.forEach((mes) => {
        allMonthlyData.push({
            mesAno: mes.mesAno,
            energia: mes.mediaDiariaDifDia,
            agua: 0,
            energiaVariacao: mes.variacaoPercentual,
            aguaVariacao: 0,
        });
    });

    aguaData.forEach((mes) => {
        const existing = allMonthlyData.find(
            (item) => item.mesAno === mes.mesAno
        );
        if (existing) {
            existing.agua = mes.mediaDiariaDifDia;
            existing.aguaVariacao = mes.variacaoPercentual;
        } else {
            allMonthlyData.push({
                mesAno: mes.mesAno,
                energia: 0,
                agua: mes.mediaDiariaDifDia,
                energiaVariacao: 0,
                aguaVariacao: mes.variacaoPercentual,
            });
        }
    });

    allMonthlyData.sort((a, b) => a.mesAno.localeCompare(b.mesAno));

    if (allMonthlyData.length > 0) {
        // Gráfico de Tendências
        const trendsCtx = document
            .getElementById("monthlyTrendsChart")
            .getContext("2d");
        monthlyTrendsChart = new Chart(trendsCtx, {
            type: "line",
            data: {
                labels: allMonthlyData.map((item) => {
                    const [year, month] = item.mesAno.split("-");
                    return `${month}/${year}`;
                }),
                datasets: [
                    {
                        label: "Energia (kWh/dia)",
                        data: allMonthlyData.map((item) => item.energia),
                        borderColor: "#f1c40f",
                        backgroundColor: "rgba(241, 196, 15, 0.1)",
                        fill: false,
                        tension: 0.4,
                        pointRadius: 5,
                        pointHoverRadius: 8,
                        yAxisID: "y",
                    },
                    {
                        label: "Água (m³/dia)",
                        data: allMonthlyData.map((item) => item.agua),
                        borderColor: "#3498db",
                        backgroundColor: "rgba(52, 152, 219, 0.1)",
                        fill: false,
                        tension: 0.4,
                        pointRadius: 5,
                        pointHoverRadius: 8,
                        yAxisID: "y1",
                    },
                ],
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                interaction: {
                    mode: "index",
                    intersect: false,
                },
                plugins: {
                    legend: {
                        display: true,
                        position: "top",
                    },
                },
                scales: {
                    x: {
                        display: true,
                        title: {
                            display: true,
                            text: "Mês/Ano",
                        },
                    },
                    y: {
                        type: "linear",
                        display: true,
                        position: "left",
                        title: {
                            display: true,
                            text: "Energia (kWh/dia)",
                            color: "#f1c40f",
                        },
                        ticks: {
                            color: "#f1c40f",
                        },
                    },
                    y1: {
                        type: "linear",
                        display: true,
                        position: "right",
                        title: {
                            display: true,
                            text: "Água (m³/dia)",
                            color: "#3498db",
                        },
                        ticks: {
                            color: "#3498db",
                        },
                        grid: {
                            drawOnChartArea: false,
                        },
                    },
                },
            },
        });

        // Gráfico de Variação Percentual
        const variationCtx = document
            .getElementById("monthlyVariationChart")
            .getContext("2d");
        monthlyVariationChart = new Chart(variationCtx, {
            type: "bar",
            data: {
                labels: allMonthlyData.slice(1).map((item) => {
                    const [year, month] = item.mesAno.split("-");
                    return `${month}/${year}`;
                }),
                datasets: [
                    {
                        label: "Energia (%)",
                        data: allMonthlyData
                            .slice(1)
                            .map((item) => item.energiaVariacao),
                        backgroundColor: function (context) {
                            const value = context.parsed.y;
                            return value >= 0
                                ? "rgba(76, 175, 80, 0.8)"
                                : "rgba(244, 67, 54, 0.8)";
                        },
                        borderColor: function (context) {
                            const value = context.parsed.y;
                            return value >= 0 ? "#4CAF50" : "#F44336";
                        },
                        borderWidth: 1,
                    },
                    {
                        label: "Água (%)",
                        data: allMonthlyData
                            .slice(1)
                            .map((item) => item.aguaVariacao),
                        backgroundColor: function (context) {
                            const value = context.parsed.y;
                            return value >= 0
                                ? "rgba(33, 150, 243, 0.8)"
                                : "rgba(255, 152, 0, 0.8)";
                        },
                        borderColor: function (context) {
                            const value = context.parsed.y;
                            return value >= 0 ? "#2196F3" : "#FF9800";
                        },
                        borderWidth: 1,
                    },
                ],
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                plugins: {
                    legend: {
                        display: true,
                        position: "top",
                    },
                    tooltip: {
                        callbacks: {
                            label: function (context) {
                                return `${
                                    context.dataset.label
                                }: ${context.parsed.y.toFixed(1)}%`;
                            },
                        },
                    },
                },
                scales: {
                    x: {
                        title: {
                            display: true,
                            text: "Mês/Ano",
                        },
                    },
                    y: {
                        title: {
                            display: true,
                            text: "Variação (%)",
                        },
                        ticks: {
                            callback: function (value) {
                                return value + "%";
                            },
                        },
                    },
                },
            },
        });
    }
}

function updateDataTable() {
    const tbody = document.getElementById("dataTableBody");
    tbody.innerHTML = "";

    processedData.forEach((record) => {
        const row = tbody.insertRow();
        const unit = record.source === "embasa" ? "m³" : "kWh";

        row.innerHTML = `
                    <td>${record.date.toLocaleDateString("pt-BR")}</td>
                    <td>${
                        record.source === "embasa"
                            ? "Embasa (Água)"
                            : "Coelba (Energia)"
                    }</td>
                    <td>${record.consumption.toFixed(2)}</td>
                    <td>${record.difDia.toFixed(2)}</td>
                    <td>${unit}</td>
                `;
    });
}

function updateMonthlyAnalysisSection() {
    if (!monthlyAnalysis || Object.keys(monthlyAnalysis).length === 0) return;

    let monthlyHTML = `
                <div class="monthly-analysis-container">
                    <h2>📊 Análise Mensal Detalhada</h2>
            `;

    if (
        monthlyAnalysis.energia &&
        Object.keys(monthlyAnalysis.energia).length > 0
    ) {
        monthlyHTML += `
                    <div class="monthly-section">
                        <h3>⚡ Energia (Coelba)</h3>
                        <div class="monthly-grid">
                `;

        Object.values(monthlyAnalysis.energia).forEach((mes) => {
            const variacao = mes.variacaoPercentual;
            const variacaoClass =
                variacao > 0
                    ? "increase"
                    : variacao < 0
                    ? "decrease"
                    : "stable";
            const variacaoIcon =
                variacao > 0 ? "📈" : variacao < 0 ? "📉" : "➡️";

            monthlyHTML += `
                        <div class="monthly-card">
                            <div class="monthly-header">
                                <h4>${mes.mes.toString().padStart(2, "0")}/${
                mes.ano
            }</h4>
                                <span class="variation ${variacaoClass}">
                                    ${variacaoIcon} ${Math.abs(
                variacao
            ).toFixed(1)}%
                                </span>
                            </div>
                            <div class="monthly-stats">
                                <div class="stat">
                                    <label>Média Diária (Dif_dia):</label>
                                    <value>${mes.mediaDiariaDifDia.toFixed(
                                        2
                                    )} kWh/dia</value>
                                </div>
                                <div class="stat">
                                    <label>Total Dif_dia:</label>
                                    <value>${mes.totalDifDia.toFixed(
                                        2
                                    )} kWh</value>
                                </div>
                                <div class="stat">
                                    <label>Dias com dados:</label>
                                    <value>${mes.diasComDifDia}/${
                mes.diasComDados
            }</value>
                                </div>
                                <div class="stat">
                                    <label>Maior/Menor Dif_dia:</label>
                                    <value>${mes.maiorDifDia.toFixed(
                                        1
                                    )} / ${mes.menorDifDia.toFixed(
                1
            )} kWh</value>
                                </div>
                            </div>
                        </div>
                    `;
        });

        monthlyHTML += `</div></div>`;
    }

    if (monthlyAnalysis.agua && Object.keys(monthlyAnalysis.agua).length > 0) {
        monthlyHTML += `
                    <div class="monthly-section">
                        <h3>💧 Água (Embasa)</h3>
                        <div class="monthly-grid">
                `;

        Object.values(monthlyAnalysis.agua).forEach((mes) => {
            const variacao = mes.variacaoPercentual;
            const variacaoClass =
                variacao > 0
                    ? "increase"
                    : variacao < 0
                    ? "decrease"
                    : "stable";
            const variacaoIcon =
                variacao > 0 ? "📈" : variacao < 0 ? "📉" : "➡️";

            monthlyHTML += `
                        <div class="monthly-card">
                            <div class="monthly-header">
                                <h4>${mes.mes.toString().padStart(2, "0")}/${
                mes.ano
            }</h4>
                                <span class="variation ${variacaoClass}">
                                    ${variacaoIcon} ${Math.abs(
                variacao
            ).toFixed(1)}%
                                </span>
                            </div>
                            <div class="monthly-stats">
                                <div class="stat">
                                    <label>Média Diária (Dif_dia):</label>
                                    <value>${mes.mediaDiariaDifDia.toFixed(
                                        2
                                    )} m³/dia</value>
                                </div>
                                <div class="stat">
                                    <label>Total Dif_dia:</label>
                                    <value>${mes.totalDifDia.toFixed(
                                        2
                                    )} m³</value>
                                </div>
                                <div class="stat">
                                    <label>Dias com dados:</label>
                                    <value>${mes.diasComDifDia}/${
                mes.diasComDados
            }</value>
                                </div>
                                <div class="stat">
                                    <label>Maior/Menor Dif_dia:</label>
                                    <value>${mes.maiorDifDia.toFixed(
                                        1
                                    )} / ${mes.menorDifDia.toFixed(
                1
            )} m³</value>
                                </div>
                            </div>
                        </div>
                    `;
        });

        monthlyHTML += `</div></div>`;
    }

    monthlyHTML += `</div>`;
    document.getElementById("monthlyAnalysisSection").innerHTML = monthlyHTML;
}

function updateCharts() {
    if (processedData.length === 0) return;

    const embasaData = processedData
        .filter((r) => r.source === "embasa")
        .sort((a, b) => new Date(a.date) - new Date(b.date));
    const coelbaData = processedData
        .filter((r) => r.source === "coelba")
        .sort((a, b) => new Date(a.date) - new Date(b.date));
    const monthlyData = groupDataByMonth();

    // Destruir gráficos existentes
    if (waterChart) {
        waterChart.destroy();
        waterChart = null;
    }
    if (energyChart) {
        energyChart.destroy();
        energyChart = null;
    }
    if (comparisonChart) {
        comparisonChart.destroy();
        comparisonChart = null;
    }
    if (monthlyTotalChart) {
        monthlyTotalChart.destroy();
        monthlyTotalChart = null;
    }

    // Gráfico de Água
    if (embasaData.length > 0) {
        const waterCtx = document.getElementById("waterChart").getContext("2d");
        waterChart = new Chart(waterCtx, {
            type: "line",
            data: {
                labels: embasaData.map((r) =>
                    r.date.toLocaleDateString("pt-BR")
                ),
                datasets: [
                    {
                        label: "Dif_dia Água (m³)",
                        data: embasaData.map((r) => r.difDia),
                        borderColor: "#3498db",
                        backgroundColor: "rgba(52, 152, 219, 0.1)",
                        fill: true,
                        tension: 0.4,
                    },
                ],
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                plugins: {
                    title: {
                        display: true,
                        text: "Consumo Diário de Água (Dif_dia)",
                        font: { size: 16, weight: "bold" },
                    },
                    legend: { display: false },
                },
                scales: {
                    y: {
                        beginAtZero: true,
                        title: { display: true, text: "m³" },
                    },
                    x: { title: { display: true, text: "Data" } },
                },
            },
        });
    }

    // Gráfico de Energia
    if (coelbaData.length > 0) {
        const energyCtx = document
            .getElementById("energyChart")
            .getContext("2d");
        energyChart = new Chart(energyCtx, {
            type: "line",
            data: {
                labels: coelbaData.map((r) =>
                    r.date.toLocaleDateString("pt-BR")
                ),
                datasets: [
                    {
                        label: "Dif_dia Energia (kWh)",
                        data: coelbaData.map((r) => r.difDia),
                        borderColor: "#f1c40f",
                        backgroundColor: "rgba(241, 196, 15, 0.1)",
                        fill: true,
                        tension: 0.4,
                    },
                ],
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                plugins: {
                    title: {
                        display: true,
                        text: "Consumo Diário de Energia (Dif_dia)",
                        font: { size: 16, weight: "bold" },
                    },
                    legend: { display: false },
                },
                scales: {
                    y: {
                        beginAtZero: true,
                        title: { display: true, text: "kWh" },
                    },
                    x: { title: { display: true, text: "Data" } },
                },
            },
        });
    }

    // Gráfico de Comparação Mensal
    if (monthlyData.labels.length > 0) {
        const comparisonCtx = document
            .getElementById("comparisonChart")
            .getContext("2d");
        comparisonChart = new Chart(comparisonCtx, {
            type: "bar",
            data: {
                labels: monthlyData.labels,
                datasets: [
                    {
                        label: "Água (m³)",
                        data: monthlyData.embasaConsumption,
                        backgroundColor: "rgba(52, 152, 219, 0.8)",
                        borderColor: "#3498db",
                        borderWidth: 1,
                    },
                    {
                        label: "Energia (kWh)",
                        data: monthlyData.coelbaConsumption,
                        backgroundColor: "rgba(241, 196, 15, 0.8)",
                        borderColor: "#f1c40f",
                        borderWidth: 1,
                    },
                ],
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                plugins: {
                    title: {
                        display: true,
                        text: "Comparação Mensal de Consumo",
                        font: { size: 16, weight: "bold" },
                    },
                },
                scales: { y: { beginAtZero: true } },
            },
        });
    }

    // Gráfico de Distribuição Total
    const totalEmbasa = processedData
        .filter((r) => r.source === "embasa")
        .reduce((sum, r) => sum + r.consumption, 0);
    const totalCoelba = processedData
        .filter((r) => r.source === "coelba")
        .reduce((sum, r) => sum + r.consumption, 0);

    if (totalEmbasa > 0 || totalCoelba > 0) {
        const monthlyTotalCtx = document
            .getElementById("monthlyTotalChart")
            .getContext("2d");
        monthlyTotalChart = new Chart(monthlyTotalCtx, {
            type: "doughnut",
            data: {
                labels: ["Água", "Energia"],
                datasets: [
                    {
                        data: [totalEmbasa, totalCoelba],
                        backgroundColor: ["#3498db", "#f1c40f"],
                        borderWidth: 3,
                        borderColor: "#fff",
                    },
                ],
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                plugins: {
                    title: {
                        display: true,
                        text: "Distribuição Total do Consumo",
                        font: { size: 16, weight: "bold" },
                    },
                    legend: { position: "bottom" },
                },
            },
        });
    }
}

function groupDataByMonth() {
    const monthlyGroups = {};

    processedData.forEach((record) => {
        const monthKey = `${record.date.getFullYear()}-${String(
            record.date.getMonth() + 1
        ).padStart(2, "0")}`;

        if (!monthlyGroups[monthKey]) {
            monthlyGroups[monthKey] = { embasa: 0, coelba: 0 };
        }

        monthlyGroups[monthKey][record.source] += record.consumption;
    });

    const sortedKeys = Object.keys(monthlyGroups).sort();

    return {
        labels: sortedKeys.map((key) => {
            const [year, month] = key.split("-");
            return `${month}/${year}`;
        }),
        embasaConsumption: sortedKeys.map((key) => monthlyGroups[key].embasa),
        coelbaConsumption: sortedKeys.map((key) => monthlyGroups[key].coelba),
    };
}

// === DOWNLOAD EXCEL ATUALIZADO ===
function downloadExcel() {
    if (processedData.length === 0) {
        showAlert("Nenhum dado disponível para download", "error");
        return;
    }

    const wb = XLSX.utils.book_new();

    // ABA 1: Dados Unificados (formato compatível)
    const unifiedData = processedData.map((record) => ({
        Data: record.date.toLocaleDateString("pt-BR"),
        Origem:
            record.source === "embasa" ? "Embasa (Água)" : "Coelba (Energia)",
        "Consumo Total": record.consumption,
        Dif_dia: record.difDia,
        Unidade: record.source === "embasa" ? "m³" : "kWh",
    }));

    // ABA 2: Formato Original - Coelba
    const coelbaOriginal = processedData
        .filter((r) => r.source === "coelba")
        .map((record) => ({
            Dia: record.date.toLocaleDateString("pt-BR"),
            Medicao: record.consumption,
            "Dif.dia": record.difDia,
        }));

    // ABA 3: Formato Original - Embasa
    const embasaOriginal = processedData
        .filter((r) => r.source === "embasa")
        .map((record) => ({
            Dia: record.date.toLocaleDateString("pt-BR"),
            Consumo: record.consumption,
            Dif_dia: record.difDia,
        }));

    // ABA 4: Estatísticas
    const embasaData = processedData.filter((r) => r.source === "embasa");
    const coelbaData = processedData.filter((r) => r.source === "coelba");
    const embasaDifDiaValues = embasaData.map((r) => r.difDia);
    const coelbaDifDiaValues = coelbaData.map((r) => r.difDia);
    const resultadoEmbasaConsumo = calcularMediaDiaria(embasaDifDiaValues);
    const resultadoCoelbaConsumo = calcularMediaDiaria(coelbaDifDiaValues);

    const statistics = [
        {
            Estatística: "Média Diária - Água",
            Valor: resultadoEmbasaConsumo.media.toFixed(2),
            Unidade: "m³/dia",
            "Soma Total": resultadoEmbasaConsumo.somaTotal?.toFixed(2) || "0",
            "Valores Utilizados": resultadoEmbasaConsumo.valoresUtilizados || 0,
        },
        {
            Estatística: "Média Diária - Energia",
            Valor: resultadoCoelbaConsumo.media.toFixed(2),
            Unidade: "kWh/dia",
            "Soma Total": resultadoCoelbaConsumo.somaTotal?.toFixed(2) || "0",
            "Valores Utilizados": resultadoCoelbaConsumo.valoresUtilizados || 0,
        },
        {
            Estatística: "Total Água",
            Valor: embasaData
                .reduce((sum, r) => sum + r.consumption, 0)
                .toFixed(2),
            Unidade: "m³",
            "Soma Total": "-",
            "Valores Utilizados": embasaData.length,
        },
        {
            Estatística: "Total Energia",
            Valor: coelbaData
                .reduce((sum, r) => sum + r.consumption, 0)
                .toFixed(2),
            Unidade: "kWh",
            "Soma Total": "-",
            "Valores Utilizados": coelbaData.length,
        },
    ];

    // ABA 5: Análise Mensal
    const monthlyDataForExcel = [];
    if (monthlyAnalysis.energia) {
        Object.values(monthlyAnalysis.energia).forEach((mes) => {
            monthlyDataForExcel.push({
                Tipo: "Energia",
                "Mês/Ano": `${mes.mes.toString().padStart(2, "0")}/${mes.ano}`,
                "Média Diária (Dif_dia)": mes.mediaDiariaDifDia.toFixed(2),
                "Total Dif_dia": mes.totalDifDia.toFixed(2),
                "Dias com Dados": `${mes.diasComDifDia}/${mes.diasComDados}`,
                "Variação % vs Mês Anterior": mes.variacaoPercentual.toFixed(1),
                "Maior Dif_dia": mes.maiorDifDia.toFixed(1),
                "Menor Dif_dia": mes.menorDifDia.toFixed(1),
                Unidade: "kWh",
            });
        });
    }

    if (monthlyAnalysis.agua) {
        Object.values(monthlyAnalysis.agua).forEach((mes) => {
            monthlyDataForExcel.push({
                Tipo: "Água",
                "Mês/Ano": `${mes.mes.toString().padStart(2, "0")}/${mes.ano}`,
                "Média Diária (Dif_dia)": mes.mediaDiariaDifDia.toFixed(2),
                "Total Dif_dia": mes.totalDifDia.toFixed(2),
                "Dias com Dados": `${mes.diasComDifDia}/${mes.diasComDados}`,
                "Variação % vs Mês Anterior": mes.variacaoPercentual.toFixed(1),
                "Maior Dif_dia": mes.maiorDifDia.toFixed(1),
                "Menor Dif_dia": mes.menorDifDia.toFixed(1),
                Unidade: "m³",
            });
        });
    }

    // Criar worksheets
    const ws1 = XLSX.utils.json_to_sheet(unifiedData);
    const ws2 = XLSX.utils.json_to_sheet(coelbaOriginal);
    const ws3 = XLSX.utils.json_to_sheet(embasaOriginal);
    const ws4 = XLSX.utils.json_to_sheet(statistics);
    const ws5 = XLSX.utils.json_to_sheet(monthlyDataForExcel);

    // Adicionar worksheets ao workbook
    XLSX.utils.book_append_sheet(wb, ws1, "Dados de Consumo");
    XLSX.utils.book_append_sheet(wb, ws2, "Coelba");
    XLSX.utils.book_append_sheet(wb, ws3, "Embasa");
    XLSX.utils.book_append_sheet(wb, ws4, "Estatísticas");
    XLSX.utils.book_append_sheet(wb, ws5, "Análise Mensal");

    const filename = `relatorio_consumo_${
        new Date().toISOString().split("T")[0]
    }.xlsx`;
    XLSX.writeFile(wb, filename);

    showAlert(
        "Relatório Excel baixado com sucesso! Agora compatível com o sistema - inclui abas no formato original (Coelba/Embasa) e formato unificado."
    );
}
