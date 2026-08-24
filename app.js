/* =========================================================
   SOLAR DGR ANALYTICS
   app.js

   ONLY THESE WORKSHEETS ARE USED:

   1. Dashboard
   2. Annual_KPI
   3. Daily_KPI
   4. PA
   5. Curtailment records

   ALL OTHER WORKSHEETS ARE IGNORED.
========================================================= */


/* =========================================================
   GLOBAL VARIABLES
========================================================= */

let workbook = null;

let charts = {};

let dgrData = {
    dashboard: null,
    annual: null,
    daily: null,
    pa: null,
    curtailment: null
};


/* =========================================================
   CONSTANTS
========================================================= */

const REQUIRED_SHEETS = [
    "Dashboard",
    "Annual_KPI",
    "Daily_KPI",
    "PA",
    "Curtailment records"
];


/* =========================================================
   DOM HELPERS
========================================================= */

function getElement(id) {
    return document.getElementById(id);
}


/* =========================================================
   STATUS MESSAGE
========================================================= */

function setStatus(message) {

    const element = getElement("statusMessage");

    if (element) {
        element.textContent = message;
    }
}


/* =========================================================
   FILE UPLOAD
========================================================= */

const fileInput = getElement("dgrUpload");

const dropZone = getElement("dropZone");

const removeFileButton = getElement("removeFile");


if (fileInput) {

    fileInput.addEventListener(
        "change",
        function(event) {

            const file = event.target.files[0];

            if (file) {
                processDGR(file);
            }

        }
    );

}


/* =========================================================
   DROP ZONE
========================================================= */

if (dropZone) {

    dropZone.addEventListener(
        "click",
        function() {

            if (fileInput) {
                fileInput.click();
            }

        }
    );


    dropZone.addEventListener(
        "dragover",
        function(event) {

            event.preventDefault();

            dropZone.classList.add("dragging");

        }
    );


    dropZone.addEventListener(
        "dragleave",
        function() {

            dropZone.classList.remove("dragging");

        }
    );


    dropZone.addEventListener(
        "drop",
        function(event) {

            event.preventDefault();

            dropZone.classList.remove("dragging");

            const file =
                event.dataTransfer.files[0];

            if (file) {
                processDGR(file);
            }

        }
    );

}


/* =========================================================
   REMOVE FILE
========================================================= */

if (removeFileButton) {

    removeFileButton.addEventListener(
        "click",
        resetApplication
    );

}


/* =========================================================
   PROCESS DGR
========================================================= */

function processDGR(file) {

    if (!file) {
        return;
    }


    const validExtensions = [
        ".xlsx",
        ".xls",
        ".csv"
    ];

    const fileName =
        file.name.toLowerCase();


    const valid =
        validExtensions.some(
            extension =>
                fileName.endsWith(extension)
        );


    if (!valid) {

        alert(
            "Please upload an Excel file (.xlsx / .xls) or CSV file."
        );

        return;
    }


    setStatus(
        "Reading DGR and analysing required worksheets..."
    );


    const reader = new FileReader();


    reader.onload = function(event) {

        try {

            const data =
                new Uint8Array(
                    event.target.result
                );


            workbook =
                XLSX.read(
                    data,
                    {
                        type: "array",
                        cellDates: true
                    }
                );


            readRequiredSheets();

            showFileInformation(file);

            updateSheetStatus();

            createDashboard();

            setStatus(
                "DGR loaded successfully. Dashboard updated."
            );

        }

        catch (error) {

            console.error(error);

            setStatus(
                "Error reading DGR."
            );

            alert(
                "There was an error reading this DGR. Please check the Excel file."
            );

        }

    };


    reader.onerror = function() {

        setStatus(
            "Unable to read the uploaded file."
        );

    };


    reader.readAsArrayBuffer(file);

}


/* =========================================================
   READ REQUIRED WORKSHEETS
========================================================= */

function readRequiredSheets() {

    dgrData.dashboard =
        readSheet(
            "Dashboard"
        );

    dgrData.annual =
        readSheet(
            "Annual_KPI"
        );

    dgrData.daily =
        readSheet(
            "Daily_KPI"
        );

    dgrData.pa =
        readSheet(
            "PA"
        );

    dgrData.curtailment =
        readSheet(
            "Curtailment records"
        );

}


/* =========================================================
   READ SHEET
========================================================= */

function readSheet(sheetName) {

    const actualSheetName =
        findSheetName(sheetName);


    if (!actualSheetName) {

        console.warn(
            "Worksheet not found:",
            sheetName
        );

        return [];

    }


    const worksheet =
        workbook.Sheets[
            actualSheetName
        ];


    return XLSX.utils.sheet_to_json(
        worksheet,
        {
            header: 1,
            defval: null,
            raw: true
        }
    );

}


/* =========================================================
   FIND SHEET
========================================================= */

function findSheetName(requiredName) {

    if (!workbook) {
        return null;
    }


    const exact =
        workbook.SheetNames.find(
            name =>
                name.trim() ===
                requiredName
        );


    if (exact) {
        return exact;
    }


    const insensitive =
        workbook.SheetNames.find(
            name =>
                name.trim().toLowerCase() ===
                requiredName.toLowerCase()
        );


    return insensitive || null;
}


/* =========================================================
   FILE INFORMATION
========================================================= */

function showFileInformation(file) {

    const fileInfo =
        getElement("fileInfo");

    const fileName =
        getElement("fileName");

    const fileDetails =
        getElement("fileDetails");


    if (fileInfo) {
        fileInfo.classList.remove("hidden");
    }


    if (fileName) {
        fileName.textContent =
            file.name;
    }


    if (fileDetails) {

        fileDetails.textContent =
            `${(
                file.size /
                1024 /
                1024
            ).toFixed(2)} MB • ` +
            `${workbook.SheetNames.length} worksheets`;
    }

}


/* =========================================================
   SHEET STATUS
========================================================= */

function updateSheetStatus() {

    updateSingleSheetStatus(
        "dashboardSheetStatus",
        "Dashboard"
    );

    updateSingleSheetStatus(
        "annualSheetStatus",
        "Annual_KPI"
    );

    updateSingleSheetStatus(
        "dailySheetStatus",
        "Daily_KPI"
    );

    updateSingleSheetStatus(
        "paSheetStatus",
        "PA"
    );

    updateSingleSheetStatus(
        "curtailmentSheetStatus",
        "Curtailment records"
    );


    const preview =
        getElement("data-preview");

    if (preview) {
        preview.classList.remove("hidden");
    }

}


function updateSingleSheetStatus(
    elementId,
    sheetName
) {

    const element =
        getElement(elementId);


    if (!element) {
        return;
    }


    const found =
        findSheetName(sheetName);


    if (found) {

        element.textContent =
            "Loaded";

        element.style.color =
            "#27a5ad";

    }

    else {

        element.textContent =
            "Not found";

        element.style.color =
            "#c15b5b";

    }

}


/* =========================================================
   CREATE DASHBOARD
========================================================= */

function createDashboard() {

    updateKPICards();

    createPAChart();

    createCurtailmentChart();

    createPRChart();

    createOperatingHoursChart();

    createSystemLossChart();

    createEnergyChart();

}


/* =========================================================
   KPI CARDS
========================================================= */

function updateKPICards() {

    const daily =
        dgrData.daily;


    /* -----------------------------------------------
       PR
    ------------------------------------------------ */

    const prValues =
        getColumnValues(
            daily,
            21
        );


    const latestPR =
        getLatestNumericValue(
            prValues
        );


    setText(
        "dashboardPR",
        formatPercent(
            latestPR
        )
    );


    /* -----------------------------------------------
       OPERATING HOURS
    ------------------------------------------------ */

    const operatingValues =
        getColumnValues(
            daily,
            8
        );


    const latestOperating =
        getLatestNumericValue(
            operatingValues
        );


    setText(
        "dashboardOperating",
        latestOperating !== null
            ? `${latestOperating.toFixed(2)} h`
            : "—"
    );


    /* -----------------------------------------------
       SYSTEM LOSS
    ------------------------------------------------ */

    const lossValues =
        getColumnValues(
            daily,
            29
        );


    const latestLoss =
        getLatestNumericValue(
            lossValues
        );


    setText(
        "dashboardLoss",
        formatPercent(
            latestLoss
        )
    );


    /* -----------------------------------------------
       PA
    ------------------------------------------------ */

    const paInfo =
        calculatePAStatistics(
            dgrData.pa
        );


    setText(
        "dashboardPA",
        paInfo.availability !== null
            ? `${paInfo.availability.toFixed(2)}%`
            : "—"
    );


    /* -----------------------------------------------
       CURTAILMENT
    ------------------------------------------------ */

    const curtailment =
        extractCurtailmentData(
            dgrData.curtailment
        );


    const totalCurtailment =
        curtailment.values.reduce(
            (sum, value) =>
                sum +
                (
                    Number.isFinite(value)
                        ? value
                        : 0
                ),
            0
        );


    setText(
        "dashboardCurtailment",
        Number.isFinite(
            totalCurtailment
        )
            ? `${totalCurtailment.toFixed(2)} MWh`
            : "—"
    );


    /* -----------------------------------------------
       MEASURED ENERGY
    ------------------------------------------------ */

    const energy =
        extractAnnualEnergy(
            dgrData.annual
        );


    const measuredTotal =
        energy.measured.reduce(
            (sum, value) =>
                sum +
                (
                    Number.isFinite(value)
                        ? value
                        : 0
                ),
            0
        );


    setText(
        "dashboardMeasured",
        Number.isFinite(
            measuredTotal
        )
            ? `${measuredTotal.toFixed(2)}`
            : "—"
    );

}


/* =========================================================
   PLANT AVAILABILITY CHART
========================================================= */

function createPAChart() {

    destroyChart(
        "paChart"
    );


    const pa =
        extractPAData(
            dgrData.pa
        );


    const canvas =
        getElement("paChart");


    if (!canvas) {
        return;
    }


    if (
        pa.labels.length === 0
    ) {

        showChartMessage(
            canvas,
            "No PA breakdown records found."
        );

        return;
    }


    /*
       Chart.js horizontal floating bars.

       Each breakdown becomes:

       start timestamp → end timestamp
    */


    const labels =
        pa.labels;


    const data =
        pa.intervals;


    charts.paChart =
        new Chart(
            canvas,
            {

                type: "bar",

                data: {

                    labels: labels,

                    datasets: [

                        {
                            label:
                                "Breakdown Duration",

                            data: data,

                            backgroundColor:
                                "rgba(39,165,173,0.65)",

                            borderColor:
                                "#27a5ad",

                            borderWidth: 1,

                            borderRadius: 4
                        }

                    ]

                },

                options: {

                    responsive: true,

                    maintainAspectRatio: false,

                    indexAxis: "y",

                    interaction: {

                        mode: "nearest",

                        intersect: false

                    },

                    plugins: {

                        legend: {
                            display: false
                        },

                        tooltip: {

                            callbacks: {

                                label:
                                    function(context) {

                                        const index =
                                            context.dataIndex;

                                        const record =
                                            pa.records[index];

                                        return [
                                            `Start: ${record.startText}`,
                                            `End: ${record.endText}`,
                                            `Duration: ${record.durationText}`
                                        ];

                                    }

                            }

                        },

                        zoom: {

                            pan: {
                                enabled: true,
                                mode: "y"
                            },

                            zoom: {

                                wheel: {
                                    enabled: true
                                },

                                pinch: {
                                    enabled: true
                                },

                                mode: "y"
                            }

                        }

                    },

                    scales: {

                        x: {

                            title: {

                                display: true,

                                text:
                                    "Duration (hours)"

                            },

                            beginAtZero: true

                        },

                        y: {

                            title: {

                                display: true,

                                text:
                                    "Breakdown / Equipment"

                            }

                        }

                    }

                }

            }
        );

}


/* =========================================================
   CURTAILMENT CHART
========================================================= */

function createCurtailmentChart() {

    destroyChart(
        "curtailmentChart"
    );


    const data =
        extractCurtailmentData(
            dgrData.curtailment
        );


    const canvas =
        getElement(
            "curtailmentChart"
        );


    if (!canvas) {
        return;
    }


    if (
        data.labels.length === 0
    ) {

        showChartMessage(
            canvas,
            "No curtailment records found."
        );

        return;
    }


    charts.curtailmentChart =
        new Chart(
            canvas,
            {

                type: "line",

                data: {

                    labels:
                        data.labels,

                    datasets: [

                        {
                            label:
                                "Curtailment Loss",

                            data:
                                data.values,

                            borderColor:
                                "#27a5ad",

                            backgroundColor:
                                "rgba(39,165,173,0.10)",

                            borderWidth: 2,

                            pointRadius: 3,

                            pointHoverRadius: 6,

                            fill: true,

                            tension: 0.25
                        }

                    ]

                },

                options: {

                    responsive: true,

                    maintainAspectRatio: false,

                    interaction: {

                        mode: "index",

                        intersect: false

                    },

                    plugins: {

                        legend: {
                            display: true
                        },

                        tooltip: {

                            callbacks: {

                                label:
                                    function(context) {

                                        return (
                                            "Curtailment Loss: " +
                                            Number(
                                                context.raw
                                            ).toFixed(2) +
                                            " MWh"
                                        );

                                    }

                            }

                        },

                        zoom: {

                            pan: {
                                enabled: true,
                                mode: "x"
                            },

                            zoom: {

                                wheel: {
                                    enabled: true
                                },

                                pinch: {
                                    enabled: true
                                },

                                mode: "x"
                            }

                        }

                    },

                    scales: {

                        x: {

                            title: {

                                display: true,

                                text: "Date"

                            }

                        },

                        y: {

                            beginAtZero: true,

                            title: {

                                display: true,

                                text:
                                    "Curtailment Loss (MWh)"

                            }

                        }

                    }

                }

            }
        );

}


/* =========================================================
   PR SCATTER PLOT
========================================================= */

function createPRChart() {

    destroyChart(
        "prChart"
    );


    const data =
        extractDailyColumn(
            dgrData.daily,
            21
        );


    const canvas =
        getElement(
            "prChart"
        );


    if (!canvas) {
        return;
    }


    const points =
        data.values
            .map(
                function(value, index) {

                    if (
                        !Number.isFinite(
                            value
                        )
                    ) {
                        return null;
                    }


                    return {

                        x:
                            index + 1,

                        y:
                            value

                    };

                }
            )
            .filter(
                point =>
                    point !== null
            );


    if (points.length === 0) {

        showChartMessage(
            canvas,
            "No PR values found in Daily_KPI column V."
        );

        return;
    }


    charts.prChart =
        new Chart(
            canvas,
            {

                type: "scatter",

                data: {

                    datasets: [

                        {
                            label:
                                "Performance Ratio",

                            data:
                                points,

                            backgroundColor:
                                "#27a5ad",

                            borderColor:
                                "#27a5ad",

                            pointRadius: 5,

                            pointHoverRadius: 8
                        }

                    ]

                },

                options: {

                    responsive: true,

                    maintainAspectRatio: false,

                    plugins: {

                        tooltip: {

                            callbacks: {

                                label:
                                    function(context) {

                                        return (
                                            `Day ${context.parsed.x}: ` +
                                            `${context.parsed.y.toFixed(2)}%`
                                        );

                                    }

                            }

                        },

                        zoom: {

                            pan: {
                                enabled: true,
                                mode: "xy"
                            },

                            zoom: {

                                wheel: {
                                    enabled: true
                                },

                                pinch: {
                                    enabled: true
                                },

                                mode: "xy"
                            }

                        }

                    },

                    scales: {

                        x: {

                            type: "linear",

                            title: {

                                display: true,

                                text:
                                    "Day of Month"

                            },

                            ticks: {

                                precision: 0

                            }

                        },

                        y: {

                            title: {

                                display: true,

                                text:
                                    "PR (%)"

                            }

                        }

                    }

                }

            }
        );

}


/* =========================================================
   OPERATING HOURS
========================================================= */

function createOperatingHoursChart() {

    destroyChart(
        "operatingChart"
    );


    const data =
        extractDailyColumn(
            dgrData.daily,
            8
        );


    const canvas =
        getElement(
            "operatingChart"
        );


    if (!canvas) {
        return;
    }


    charts.operatingChart =
        new Chart(
            canvas,
            {

                type: "line",

                data: {

                    labels:
                        data.labels,

                    datasets: [

                        {
                            label:
                                "Operating Hours",

                            data:
                                data.values,

                            borderColor:
                                "#27a5ad",

                            backgroundColor:
                                "rgba(39,165,173,0.10)",

                            borderWidth: 2,

                            pointRadius: 3,

                            pointHoverRadius: 6,

                            fill: true,

                            tension: 0.25
                        }

                    ]

                },

                options: {

                    responsive: true,

                    maintainAspectRatio: false,

                    plugins: {

                        tooltip: {

                            callbacks: {

                                label:
                                    function(context) {

                                        return (
                                            "Operating Hours: " +
                                            Number(
                                                context.raw
                                            ).toFixed(2) +
                                            " h"
                                        );

                                    }

                            }

                        },

                        zoom: {

                            pan: {
                                enabled: true,
                                mode: "x"
                            },

                            zoom: {

                                wheel: {
                                    enabled: true
                                },

                                pinch: {
                                    enabled: true
                                },

                                mode: "x"
                            }

                        }

                    },

                    scales: {

                        x: {

                            title: {

                                display: true,

                                text:
                                    "Day of Month"

                            }

                        },

                        y: {

                            beginAtZero: true,

                            title: {

                                display: true,

                                text:
                                    "Operating Hours"

                            }

                        }

                    }

                }

            }
        );

}


/* =========================================================
   SYSTEM LOSSES
========================================================= */

function createSystemLossChart() {

    destroyChart(
        "systemLossChart"
    );


    const data =
        extractDailyColumn(
            dgrData.daily,
            29
        );


    const canvas =
        getElement(
            "systemLossChart"
        );


    if (!canvas) {
        return;
    }


    charts.systemLossChart =
        new Chart(
            canvas,
            {

                type: "line",

                data: {

                    labels:
                        data.labels,

                    datasets: [

                        {
                            label:
                                "System Losses",

                            data:
                                data.values,

                            borderColor:
                                "#27a5ad",

                            backgroundColor:
                                "rgba(39,165,173,0.12)",

                            borderWidth: 2,

                            pointRadius: 3,

                            pointHoverRadius: 6,

                            fill: true,

                            tension: 0.25
                        }

                    ]

                },

                options: {

                    responsive: true,

                    maintainAspectRatio: false,

                    plugins: {

                        tooltip: {

                            callbacks: {

                                label:
                                    function(context) {

                                        return (
                                            "System Loss: " +
                                            Number(
                                                context.raw
                                            ).toFixed(2) +
                                            "%"
                                        );

                                    }

                            }

                        },

                        zoom: {

                            pan: {
                                enabled: true,
                                mode: "x"
                            },

                            zoom: {

                                wheel: {
                                    enabled: true
                                },

                                pinch: {
                                    enabled: true
                                },

                                mode: "x"
                            }

                        }

                    },

                    scales: {

                        x: {

                            title: {

                                display: true,

                                text:
                                    "Day of Month"

                            }

                        },

                        y: {

                            title: {

                                display: true,

                                text:
                                    "System Losses (%)"

                            }

                        }

                    }

                }

            }
        );

}


/* =========================================================
   ENERGY CHART
========================================================= */

function createEnergyChart() {

    destroyChart(
        "energyChart"
    );


    const energy =
        extractAnnualEnergy(
            dgrData.annual
        );


    const canvas =
        getElement(
            "energyChart"
        );


    if (!canvas) {
        return;
    }


    if (
        energy.labels.length === 0
    ) {

        showChartMessage(
            canvas,
            "No Annual_KPI energy data found."
        );

        return;
    }


    charts.energyChart =
        new Chart(
            canvas,
            {

                type: "bar",

                data: {

                    labels:
                        energy.labels,

                    datasets: [

                        {
                            label:
                                "Budgeted Energy",

                            data:
                                energy.budget,

                            backgroundColor:
                                "rgba(39,165,173,0.70)",

                            borderColor:
                                "#27a5ad",

                            borderWidth: 1,

                            borderRadius: 3
                        },


                        {
                            label:
                                "Measured Energy",

                            data:
                                energy.measured,

                            backgroundColor:
                                "rgba(23,37,42,0.72)",

                            borderColor:
                                "#17252a",

                            borderWidth: 1,

                            borderRadius: 3
                        }

                    ]

                },

                options: {

                    responsive: true,

                    maintainAspectRatio: false,

                    interaction: {

                        mode: "index",

                        intersect: false

                    },

                    plugins: {

                        tooltip: {

                            callbacks: {

                                label:
                                    function(context) {

                                        return (
                                            context.dataset.label +
                                            ": " +
                                            Number(
                                                context.raw
                                            ).toFixed(2)
                                        );

                                    }

                            }

                        },

                        zoom: {

                            pan: {
                                enabled: true,
                                mode: "x"
                            },

                            zoom: {

                                wheel: {
                                    enabled: true
                                },

                                pinch: {
                                    enabled: true
                                },

                                mode: "x"
                            }

                        }

                    },

                    scales: {

                        x: {

                            title: {

                                display: true,

                                text:
                                    "Month"

                            }

                        },

                        y: {

                            beginAtZero: true,

                            title: {

                                display: true,

                                text:
                                    "Energy"

                            }

                        }

                    }

                }

            }
        );


    updateEnergySummary(
        energy
    );

}


/* =========================================================
   EXTRACT DAILY COLUMN
========================================================= */

function extractDailyColumn(
    rows,
    columnIndex
) {

    if (
        !rows ||
        rows.length === 0
    ) {

        return {
            labels: [],
            values: []
        };

    }


    const start =
        findDataStartRow(
            rows
        );


    const labels = [];

    const values = [];


    for (
        let i = start;
        i < rows.length;
        i++
    ) {

        const row =
            rows[i];


        if (
            !row ||
            row.length <= columnIndex
        ) {
            continue;
        }


        const raw =
            row[columnIndex];


        const value =
            parseNumber(
                raw
            );


        if (
            value === null
        ) {
            continue;
        }


        labels.push(
            getDayLabel(
                row,
                i - start + 1
            )
        );


        values.push(
            value
        );

    }


    return {
        labels,
        values
    };

}


/* =========================================================
   GET COLUMN VALUES
========================================================= */

function getColumnValues(
    rows,
    columnIndex
) {

    if (!rows) {
        return [];
    }


    return rows
        .slice(
            findDataStartRow(rows)
        )
        .map(
            row =>
                parseNumber(
                    row?.[columnIndex]
                )
        )
        .filter(
            value =>
                value !== null
        );

}


/* =========================================================
   ANNUAL ENERGY
========================================================= */

function extractAnnualEnergy(
    rows
) {

    const result = {

        labels: [],

        budget: [],

        measured: []

    };


    if (
        !rows ||
        rows.length === 0
    ) {

        return result;
    }


    /*
       E = index 4
       F = index 5

       Rows 10–21 in Excel
       = indexes 9–20
    */


    for (
        let excelRow = 10;
        excelRow <= 21;
        excelRow++
    ) {

        const index =
            excelRow - 1;


        const row =
            rows[index];


        if (!row) {
            continue;
        }


        const budget =
            parseNumber(
                row[4]
            );


        const measured =
            parseNumber(
                row[5]
            );


        result.labels.push(
            getMonthLabel(
                row,
                excelRow
            )
        );


        result.budget.push(
            budget
        );


        result.measured.push(
            measured
        );

    }


    return result;

}


/* =========================================================
   PA DATA
========================================================= */

function extractPAData(
    rows
) {

    const result = {

        labels: [],

        intervals: [],

        records: []

    };


    if (
        !rows ||
        rows.length === 0
    ) {

        return result;
    }


    const header =
        findHeaderRow(
            rows,
            [
                "start",
                "breakdown",
                "from"
            ]
        );


    let startIndex =
        header !== -1
            ? header + 1
            : 1;


    for (
        let i = startIndex;
        i < rows.length;
        i++
    ) {

        const row =
            rows[i];


        if (
            !row ||
            row.length === 0
        ) {
            continue;
        }


        const start =
            findTimeValue(
                row,
                [
                    "start",
                    "breakdown start",
                    "from"
                ]
            );


        const end =
            findTimeValue(
                row,
                [
                    "end",
                    "breakdown end",
                    "to"
                ]
            );


        if (
            start === null ||
            end === null
        ) {
            continue;
        }


        const durationHours =
            calculateDuration(
                start,
                end
            );


        if (
            durationHours === null
        ) {
            continue;
        }


        const equipment =
            findTextValue(
                row,
                [
                    "equipment",
                    "asset",
                    "machine",
                    "name"
                ]
            ) ||
            `Breakdown ${result.records.length + 1}`;


        result.labels.push(
            equipment
        );


        result.intervals.push(
            durationHours
        );


        result.records.push({

            equipment,

            startText:
                formatDateTime(
                    start
                ),

            endText:
                formatDateTime(
                    end
                ),

            durationText:
                formatDuration(
                    durationHours
                )

        });

    }


    return result;

}


/* =========================================================
   PA STATISTICS
========================================================= */

function calculatePAStatistics(
    rows
) {

    const pa =
        extractPAData(
            rows
        );


    if (
        pa.records.length === 0
    ) {

        return {
            availability: null
        };

    }


    const totalDowntime =
        pa.intervals.reduce(
            (sum, value) =>
                sum + value,
            0
        );


    /*
       Availability is calculated from the
       period represented by the PA records.

       This is a fallback calculation.
    */

    const periodHours =
        24 *
        Math.max(
            1,
            pa.records.length
        );


    const availability =
        Math.max(
            0,
            Math.min(
                100,
                (
                    1 -
                    (
                        totalDowntime /
                        periodHours
                    )
                ) *
                100
            )
        );


    return {
        availability
    };

}


/* =========================================================
   CURTAILMENT DATA
========================================================= */

function extractCurtailmentData(
    rows
) {

    const result = {

        labels: [],

        values: []

    };


    if (
        !rows ||
        rows.length === 0
    ) {

        return result;
    }


    const headerRow =
        findHeaderRow(
            rows,
            [
                "curtail",
                "loss",
                "energy"
            ]
        );


    const start =
        headerRow >= 0
            ? headerRow + 1
            : 1;


    for (
        let i = start;
        i < rows.length;
        i++
    ) {

        const row =
            rows[i];


        if (
            !row ||
            row.length === 0
        ) {
            continue;
        }


        const date =
            findDateValue(
                row
            );


        const loss =
            findCurtailmentNumber(
                row
            );


        if (
            loss === null
        ) {
            continue;
        }


        result.labels.push(
            date !== null
                ? formatDate(
                    date
                )
                : `Record ${i - start + 1}`
        );


        result.values.push(
            loss
        );

    }


    return result;

}


/* =========================================================
   FIND CURTAILMENT NUMBER
========================================================= */

function findCurtailmentNumber(
    row
) {

    /*
       Try headers first.
    */

    const possibleHeaders = [
        "loss",
        "curtailment loss",
        "energy loss",
        "loss mwh",
        "curtailment",
        "curtailed energy",
        "mwh"
    ];


    const headerIndex =
        findColumnByHeader(
            dgrData.curtailment,
            possibleHeaders
        );


    if (
        headerIndex !== -1 &&
        row[headerIndex] !== undefined
    ) {

        return parseNumber(
            row[headerIndex]
        );

    }


    /*
       Fallback:
       use the last numeric value in the row.
    */

    for (
        let i = row.length - 1;
        i >= 0;
        i--
    ) {

        const number =
            parseNumber(
                row[i]
            );


        if (
            number !== null
        ) {

            return number;

        }

    }


    return null;

}


/* =========================================================
   FIND HEADER ROW
========================================================= */

function findHeaderRow(
    rows,
    keywords
) {

    const maxRows =
        Math.min(
            rows.length,
            30
        );


    for (
        let i = 0;
        i < maxRows;
        i++
    ) {

        const row =
            rows[i];


        if (!row) {
            continue;
        }


        const text =
            row
                .map(
                    cell =>
                        String(
                            cell ?? ""
                        ).toLowerCase()
                )
                .join(" ");


        const found =
            keywords.some(
                keyword =>
                    text.includes(
                        keyword.toLowerCase()
                    )
            );


        if (found) {
            return i;
        }

    }


    return -1;

}


/* =========================================================
   FIND COLUMN BY HEADER
========================================================= */

function findColumnByHeader(
    rows,
    possibleHeaders
) {

    const headerRow =
        findHeaderRow(
            rows,
            possibleHeaders
        );


    if (
        headerRow === -1
    ) {
        return -1;
    }


    const row =
        rows[headerRow];


    for (
        let i = 0;
        i < row.length;
        i++
    ) {

        const text =
            String(
                row[i] ?? ""
            ).toLowerCase().trim();


        for (
            const header
            of possibleHeaders
        ) {

            if (
                text ===
                header.toLowerCase()
            ) {

                return i;

            }

        }

    }


    return -1;

}


/* =========================================================
   FIND TEXT VALUE
========================================================= */

function findTextValue(
    row,
    keywords
) {

    for (
        const cell
        of row
    ) {

        const text =
            String(
                cell ?? ""
            ).trim();


        if (!text) {
            continue;
        }


        const lower =
            text.toLowerCase();


        const found =
            keywords.some(
                keyword =>
                    lower.includes(
                        keyword.toLowerCase()
                    )
            );


        if (found) {
            return text;
        }

    }


    return null;

}


/* =========================================================
   FIND TIME VALUE
========================================================= */

function findTimeValue(
    row,
    keywords
) {

    for (
        const cell
        of row
    ) {

        if (
            cell instanceof Date
        ) {
            return cell;
        }


        if (
            typeof cell === "number" &&
            cell >= 0 &&
            cell <= 1
        ) {

            const date =
                XLSX.SSF.parse_date_code(
                    cell
                );


            if (date) {

                return new Date(
                    date.y,
                    date.m - 1,
                    date.d,
                    date.H,
                    date.M,
                    date.S
                );

            }

        }


        if (
            typeof cell === "string"
        ) {

            const parsed =
                parseDateString(
                    cell
                );


            if (parsed) {
                return parsed;
            }

        }

    }


    return null;

}


/* =========================================================
   FIND DATE
========================================================= */

function findDateValue(
    row
) {

    for (
        const cell
        of row
    ) {

        if (
            cell instanceof Date
        ) {

            return cell;

        }


        if (
            typeof cell === "number" &&
            cell > 20000
        ) {

            const date =
                XLSX.SSF.parse_date_code(
                    cell
                );


            if (date) {

                return new Date(
                    date.y,
                    date.m - 1,
                    date.d
                );

            }

        }


        if (
            typeof cell === "string"
        ) {

            const parsed =
                parseDateString(
                    cell
                );


            if (parsed) {
                return parsed;
            }

        }

    }


    return null;

}


/* =========================================================
   DATE PARSER
========================================================= */

function parseDateString(
    value
) {

    const text =
        String(
            value
        ).trim();


    if (!text) {
        return null;
    }


    const date =
        new Date(
            text
        );


    if (
        !Number.isNaN(
            date.getTime()
        )
    ) {

        return date;

    }


    return null;

}


/* =========================================================
   CALCULATE DURATION
========================================================= */

function calculateDuration(
    start,
    end
) {

    if (
        !(start instanceof Date) ||
        !(end instanceof Date)
    ) {

        return null;

    }


    let difference =
        end.getTime() -
        start.getTime();


    /*
       If only time is supplied and end
       appears before start, assume it
       crossed midnight.
    */

    if (
        difference < 0
    ) {

        difference +=
            24 *
            60 *
            60 *
            1000;

    }


    return (
        difference /
        (
            1000 *
            60 *
            60
        )
    );

}


/* =========================================================
   FORMAT DURATION
========================================================= */

function formatDuration(
    hours
) {

    const totalMinutes =
        Math.round(
            hours * 60
        );


    const h =
        Math.floor(
            totalMinutes / 60
        );


    const m =
        totalMinutes % 60;


    return `${h}h ${m}m`;

}


/* =========================================================
   FORMAT DATE TIME
========================================================= */

function formatDateTime(
    date
) {

    if (
        !(date instanceof Date)
    ) {

        return "—";

    }


    return date.toLocaleString(
        "en-IN",
        {
            day: "2-digit",
            month: "short",
            year: "numeric",
            hour: "2-digit",
            minute: "2-digit"
        }
    );

}


/* =========================================================
   FORMAT DATE
========================================================= */

function formatDate(
    date
) {

    if (
        !(date instanceof Date)
    ) {

        return String(
            date
        );

    }


    return date.toLocaleDateString(
        "en-IN",
        {
            day: "2-digit",
            month: "short"
        }
    );

}


/* =========================================================
   MONTH LABEL
========================================================= */

function getMonthLabel(
    row,
    excelRow
) {

    /*
       Try to find a month name in the row.
    */

    for (
        const cell
        of row
    ) {

        if (
            typeof cell === "string"
        ) {

            const text =
                cell.trim();


            if (
                text.length >= 3 &&
                text.length <= 15
            ) {

                return text;

            }

        }

    }


    /*
       Fallback.
    */

    const months = [
        "Apr",
        "May",
        "Jun",
        "Jul",
        "Aug",
        "Sep",
        "Oct",
        "Nov",
        "Dec",
        "Jan",
        "Feb",
        "Mar"
    ];


    return (
        months[
            excelRow - 10
        ] ||
        `Month ${excelRow - 9}`
    );

}


/* =========================================================
   DAY LABEL
========================================================= */

function getDayLabel(
    row,
    fallback
) {

    for (
        const cell
        of row
    ) {

        if (
            cell instanceof Date
        ) {

            return cell.getDate();

        }


        if (
            typeof cell === "number" &&
            cell >= 1 &&
            cell <= 31
        ) {

            return cell;

        }


        if (
            typeof cell === "string"
        ) {

            const match =
                cell.match(
                    /\b([1-9]|[12]\d|3[01])\b/
                );


            if (match) {
                return match[1];
            }

        }

    }


    return fallback;

}


/* =========================================================
   FIND DATA START
========================================================= */

function findDataStartRow(
    rows
) {

    if (
        !rows ||
        rows.length === 0
    ) {

        return 0;

    }


    /*
       Find a row that appears to contain
       headers.
    */

    for (
        let i = 0;
        i < Math.min(
            rows.length,
            15
        );
        i++
    ) {

        const row =
            rows[i];


        if (!row) {
            continue;
        }


        const text =
            row
                .map(
                    cell =>
                        String(
                            cell ?? ""
                        ).toLowerCase()
                )
                .join(" ");


        if (
            text.includes("date") ||
            text.includes("day") ||
            text.includes("kpi")
        ) {

            return i + 1;

        }

    }


    /*
       Most DGR Daily_KPI files have
       header rows above the actual data.
       Find the first row with numeric
       values in the target KPI columns.
    */

    for (
        let i = 0;
        i < rows.length;
        i++
    ) {

        const row =
            rows[i];


        if (
            row &&
            (
                parseNumber(
                    row[8]
                ) !== null ||
                parseNumber(
                    row[21]
                ) !== null ||
                parseNumber(
                    row[29]
                ) !== null
            )
        ) {

            return i;

        }

    }


    return 1;

}


/* =========================================================
   PARSE NUMBER
========================================================= */

function parseNumber(
    value
) {

    if (
        value === null ||
        value === undefined ||
        value === ""
    ) {

        return null;

    }


    if (
        typeof value === "number"
    ) {

        return Number.isFinite(
            value
        )
            ? value
            : null;

    }


    let text =
        String(
            value
        ).trim();


    if (!text) {
        return null;
    }


    /*
       Remove commas and units.
    */

    text =
        text.replace(
            /,/g,
            ""
        );


    const match =
        text.match(
            /-?\d+(?:\.\d+)?/
        );


    if (!match) {
        return null;
    }


    const number =
        Number(
            match[0]
        );


    return Number.isFinite(
        number
    )
        ? number
        : null;

}


/* =========================================================
   LATEST NUMERIC VALUE
========================================================= */

function getLatestNumericValue(
    values
) {

    for (
        let i = values.length - 1;
        i >= 0;
        i--
    ) {

        if (
            Number.isFinite(
                values[i]
            )
        ) {

            return values[i];

        }

    }


    return null;

}


/* =========================================================
   FORMAT PERCENT
========================================================= */

function formatPercent(
    value
) {

    if (
        value === null ||
        !Number.isFinite(value)
    ) {

        return "—";

    }


    return `${value.toFixed(2)}%`;

}


/* =========================================================
   UPDATE ENERGY SUMMARY
========================================================= */

function updateEnergySummary(
    energy
) {

    const budget =
        energy.budget.reduce(
            (sum, value) =>
                sum +
                (
                    Number.isFinite(value)
                        ? value
                        : 0
                ),
            0
        );


    const measured =
        energy.measured.reduce(
            (sum, value) =>
                sum +
                (
                    Number.isFinite(value)
                        ? value
                        : 0
                ),
            0
        );


    const variance =
        measured -
        budget;


    setText(
        "totalBudget",
        `${budget.toFixed(2)} GWh`
    );


    setText(
        "totalMeasured",
        `${measured.toFixed(2)}`
    );


    setText(
        "energyVariance",
        `${variance >= 0 ? "+" : ""}${variance.toFixed(2)}`
    );

}


/* =========================================================
   SET TEXT
========================================================= */

function setText(
    id,
    value
) {

    const element =
        getElement(
            id
        );


    if (element) {
        element.textContent =
            value;
    }

}


/* =========================================================
   DESTROY CHART
========================================================= */

function destroyChart(
    id
) {

    if (
        charts[id]
    ) {

        charts[id].destroy();

        delete charts[id];

    }

}


/* =========================================================
   SHOW CHART MESSAGE
========================================================= */

function showChartMessage(
    canvas,
    message
) {

    const parent =
        canvas.parentElement;


    if (!parent) {
        return;
    }


    parent.style.display =
        "flex";

    parent.style.alignItems =
        "center";

    parent.style.justifyContent =
        "center";


    const messageElement =
        document.createElement(
            "div"
        );


    messageElement.textContent =
        message;


    messageElement.style.color =
        "#89979a";

    messageElement.style.fontSize =
        "11px";


    parent.appendChild(
        messageElement
    );

}


/* =========================================================
   NAVIGATION
========================================================= */

document
    .querySelectorAll(
        ".nav-item, .quick-navigation button"
    )
    .forEach(
        button => {

            button.addEventListener(
                "click",
                function() {

                    const targetId =
                        this.dataset.target;


                    const target =
                        getElement(
                            targetId
                        );


                    if (!target) {
                        return;
                    }


                    target.scrollIntoView(
                        {
                            behavior: "smooth",
                            block: "start"
                        }
                    );


                    /*
                       Update sidebar active state.
                    */

                    document
                        .querySelectorAll(
                            ".nav-item"
                        )
                        .forEach(
                            item =>
                                item.classList.remove(
                                    "active"
                                )
                        );


                    const nav =
                        document.querySelector(
                            `.nav-item[data-target="${targetId}"]`
                        );


                    if (nav) {
                        nav.classList.add(
                            "active"
                        );
                    }

                }
            );

        }
    );


/* =========================================================
   RESET ZOOM BUTTONS
========================================================= */

document
    .querySelectorAll(
        ".reset-chart"
    )
    .forEach(
        button => {

            button.addEventListener(
                "click",
                function() {

                    const chartId =
                        this.dataset.chart;


                    const chart =
                        charts[
                            chartId
                        ];


                    if (
                        chart &&
                        chart.resetZoom
                    ) {

                        chart.resetZoom();

                    }

                }
            );

        }
    );


/* =========================================================
   RESET APPLICATION
========================================================= */

function resetApplication() {

    workbook = null;


    dgrData = {

        dashboard: null,

        annual: null,

        daily: null,

        pa: null,

        curtailment: null

    };


    Object.keys(
        charts
    ).forEach(
        id =>
            destroyChart(id)
    );


    if (fileInput) {
        fileInput.value = "";
    }


    const fileInfo =
        getElement(
            "fileInfo"
        );


    if (fileInfo) {
        fileInfo.classList.add(
            "hidden"
        );
    }


    const dataPreview =
        getElement(
            "data-preview"
        );


    if (dataPreview) {
        dataPreview.classList.add(
            "hidden"
        );
    }


    setStatus(
        "Upload your DGR to begin analysis."
    );


    setText(
        "dataStatus",
        "No DGR loaded"
    );


    const dot =
        document.querySelector(
            ".status-dot"
        );


    if (dot) {
        dot.classList.remove(
            "loaded"
        );
    }


    [
        "dashboardPA",
        "dashboardPR",
        "dashboardOperating",
        "dashboardLoss",
        "dashboardCurtailment",
        "dashboardMeasured",
        "totalBudget",
        "totalMeasured",
        "energyVariance"
    ]
    .forEach(
        id =>
            setText(
                id,
                "—"
            )
    );

}


/* =========================================================
   UPDATE DATA STATUS AFTER LOAD
========================================================= */

function updateLoadedStatus() {

    setText(
        "dataStatus",
        "DGR loaded"
    );


    const dot =
        document.querySelector(
            ".status-dot"
        );


    if (dot) {
        dot.classList.add(
            "loaded"
        );
    }

}


/* =========================================================
   OVERRIDE CREATE DASHBOARD
   TO ALSO UPDATE STATUS
========================================================= */

const originalCreateDashboard =
    createDashboard;


/*
   We intentionally call status update
   after the dashboard has been created.
*/

function runDashboard() {

    originalCreateDashboard();

    updateLoadedStatus();

}


/* =========================================================
   PATCH PROCESSING
========================================================= */

const originalProcessDGR =
    processDGR;


/*
   The original processDGR already creates
   the dashboard. This function remains here
   for clarity and future extensions.
*/


/* =========================================================
   INITIAL STATE
========================================================= */

setStatus(
    "Upload your DGR to begin analysis."
);
