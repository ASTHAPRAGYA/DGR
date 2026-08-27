/* =========================================================
   SOLAR DGR ANALYTICS
   app.js
   ========================================================= */

"use strict";

/* =========================================================
   GLOBAL STATE
========================================================= */

let workbook = null;
let uploadedFile = null;

let dashboardData = [];
let dailyKPIData = [];
let paData = [];
let curtailmentData = [];
let annualKPIData = [];

const charts = {};


/* =========================================================
   DOM READY
========================================================= */

document.addEventListener("DOMContentLoaded", () => {

    initialiseNavigation();
    initialiseUpload();
    initialiseRemoveButton();

    hideDashboardUntilUpload();

});


/* =========================================================
   NAVIGATION
========================================================= */

function initialiseNavigation() {

    const navItems = document.querySelectorAll(".nav-item");

    navItems.forEach(item => {

        item.addEventListener("click", () => {

            const targetId = item.dataset.target;
            const target = document.getElementById(targetId);

            if (!target) return;

            navItems.forEach(nav => {
                nav.classList.remove("active");
            });

            item.classList.add("active");

            target.scrollIntoView({
                behavior: "smooth",
                block: "start"
            });

        });

    });

}


/* =========================================================
   UPLOAD
========================================================= */

function initialiseUpload() {

    const fileInput = document.getElementById("dgrFile");
    const dropZone = document.getElementById("dropZone");

    if (!fileInput) return;

    fileInput.addEventListener("change", event => {

        const file = event.target.files?.[0];

        if (file) {
            processDGR(file);
        }

    });


    if (!dropZone) return;


    dropZone.addEventListener("click", () => {
        fileInput.click();
    });


    dropZone.addEventListener("dragover", event => {

        event.preventDefault();

        dropZone.classList.add("dragging");

    });


    dropZone.addEventListener("dragleave", () => {

        dropZone.classList.remove("dragging");

    });


    dropZone.addEventListener("drop", event => {

        event.preventDefault();

        dropZone.classList.remove("dragging");

        const file = event.dataTransfer.files?.[0];

        if (!file) return;

        const valid =
            file.name.toLowerCase().endsWith(".xlsx") ||
            file.name.toLowerCase().endsWith(".xls") ||
            file.name.toLowerCase().endsWith(".csv");

        if (!valid) {

            setStatus(
                "Please upload an Excel (.xlsx/.xls) or CSV file."
            );

            return;
        }

        processDGR(file);

    });

}


/* =========================================================
   REMOVE FILE
========================================================= */

function initialiseRemoveButton() {

    const removeButton = document.getElementById("removeFile");

    if (!removeButton) return;

    removeButton.addEventListener("click", resetApplication);

}


/* =========================================================
   PROCESS DGR
========================================================= */

function processDGR(file) {

    if (typeof XLSX === "undefined") {

        setStatus(
            "SheetJS could not be loaded. Please check the XLSX library."
        );

        return;
    }


    uploadedFile = file;

    setStatus("Reading DGR workbook...");


    const reader = new FileReader();


    reader.onload = event => {

        try {

            const arrayBuffer = event.target.result;

            workbook = XLSX.read(arrayBuffer, {
                type: "array",
                cellDates: true,
                cellNF: true,
                cellText: true
            });


            readWorkbook();

            showUploadedFile();

            showDashboard();

            renderAllCharts();

            setStatus(
                `${file.name} analysed successfully.`
            );

        }

        catch (error) {

            console.error(error);

            setStatus(
                "Unable to read the DGR. Please check the workbook format."
            );

        }

    };


    reader.onerror = () => {

        setStatus(
            "The DGR file could not be read."
        );

    };


    reader.readAsArrayBuffer(file);

}


/* =========================================================
   READ WORKBOOK
========================================================= */

function readWorkbook() {

    dashboardData = getWorksheetData([
        "Dashboard",
        "DASHBOARD"
    ]);

    dailyKPIData = getWorksheetData([
        "Daily_KPI",
        "Daily KPI",
        "DAILY_KPI",
        "DailyKPI"
    ]);

    paData = getWorksheetData([
        "PA",
        "Plant Availability",
        "Plant_Availability"
    ]);

    curtailmentData = getWorksheetData([
        "Curtailment",
        "CURTAILMENT",
        "Grid Curtailment",
        "Grid_Curtailment"
    ]);

    annualKPIData = getWorksheetData([
        "Annual_KPI",
        "Annual KPI",
        "ANNUAL_KPI",
        "AnnualKPI"
    ]);


    updateWorkbookStatus();

}


/* =========================================================
   GET WORKSHEET
========================================================= */

function getWorksheetData(possibleNames) {

    if (!workbook) return [];

    let sheetName = null;


    for (const name of possibleNames) {

        if (
            workbook.SheetNames.some(
                actual =>
                    actual.toLowerCase() === name.toLowerCase()
            )
        ) {

            sheetName = workbook.SheetNames.find(
                actual =>
                    actual.toLowerCase() === name.toLowerCase()
            );

            break;
        }

    }


    if (!sheetName) return [];


    const worksheet = workbook.Sheets[sheetName];

    if (!worksheet) return [];


    return XLSX.utils.sheet_to_json(
        worksheet,
        {
            defval: null,
            raw: true
        }
    );

}


/* =========================================================
   WORKBOOK STATUS
========================================================= */

function updateWorkbookStatus() {

    const statusPanel =
        document.getElementById("workbookStatus");

    const badgeContainer =
        document.getElementById("sheetBadges");

    if (!statusPanel || !badgeContainer) return;


    const requiredSheets = [
        "Dashboard",
        "Annual_KPI",
        "Daily_KPI",
        "PA",
        "Curtailment"
    ];


    badgeContainer.innerHTML = "";


    requiredSheets.forEach(required => {

        const found = workbook.SheetNames.some(
            sheet =>
                sheet.toLowerCase() === required.toLowerCase()
        );


        const badge = document.createElement("span");

        badge.className =
            found
                ? "sheet-badge"
                : "sheet-badge missing";

        badge.textContent =
            found
                ? `${required} ✓`
                : `${required} — Missing`;


        badgeContainer.appendChild(badge);

    });


    statusPanel.classList.remove("hidden");

}


/* =========================================================
   FILE INFORMATION
========================================================= */

function showUploadedFile() {

    const fileInfo =
        document.getElementById("fileInfo");

    const fileName =
        document.getElementById("fileName");

    const fileSheets =
        document.getElementById("fileSheets");

    const sidebarFileName =
        document.getElementById("sidebarFileName");


    if (fileName) {
        fileName.textContent = uploadedFile.name;
    }


    if (fileSheets && workbook) {

        fileSheets.textContent =
            `${workbook.SheetNames.length} worksheets detected`;

    }


    if (sidebarFileName) {
        sidebarFileName.textContent =
            uploadedFile.name;
    }


    if (fileInfo) {
        fileInfo.classList.remove("hidden");
    }

}


/* =========================================================
   SHOW / HIDE DASHBOARD
========================================================= */

function hideDashboardUntilUpload() {

    const sections = document.querySelectorAll(
        ".dashboard-section"
    );

    sections.forEach(section => {
        section.style.display = "none";
    });

}


function showDashboard() {

    const sections = document.querySelectorAll(
        ".dashboard-section"
    );

    sections.forEach(section => {
        section.style.display = "";
    });


    const emptyState =
        document.getElementById("emptyState");

    if (emptyState) {
        emptyState.classList.add("hidden");
    }

}


/* =========================================================
   STATUS
========================================================= */

function setStatus(message) {

    const statusText =
        document.getElementById("statusText");

    if (statusText) {
        statusText.textContent = message;
    }

}


/* =========================================================
   RENDER ALL
========================================================= */

function renderAllCharts() {

    updateKPIs();

    renderPerformanceCharts();

    renderDashboardCharts();

    renderPACChart();

    renderCurtailmentChart();

    renderEnergyChart();

}


/* =========================================================
   KPI CARDS
========================================================= */

function updateKPIs() {

    const pa =
        findLatestValue(
            paData,
            [
                "PA",
                "Plant Availability",
                "Plant Availability (%)",
                "Availability"
            ]
        );


    const pr =
        getColumnValue(
            dailyKPIData,
            21
        );


    const loss =
        getColumnValue(
            dailyKPIData,
            29
        );


    const hours =
        getColumnValue(
            dailyKPIData,
            8
        );


    setText(
        "dashboardPA",
        formatPercent(pa)
    );


    setText(
        "dashboardPR",
        formatPercent(pr)
    );


    setText(
        "dashboardLoss",
        formatPercent(loss)
    );


    setText(
        "dashboardHours",
        formatNumber(hours, 2)
    );

}


/* =========================================================
   PERFORMANCE CHARTS
========================================================= */

function renderPerformanceCharts() {

    const dates =
        dailyKPIData.map(
            row => getDateFromRow(row)
        );


    const prValues =
        dailyKPIData.map(
            row => parseNumber(getColumn(row, 21))
        );


    const hoursValues =
        dailyKPIData.map(
            row => parseNumber(getColumn(row, 8))
        );


    const lossValues =
        dailyKPIData.map(
            row => parseNumber(getColumn(row, 29))
        );


    createDateLineChart(
        "prChart",
        dates,
        prValues,
        "Performance Ratio",
        true
    );


    createDateLineChart(
        "hoursChart",
        dates,
        hoursValues,
        "Operating Hours",
        false
    );


    createDateLineChart(
        "lossChart",
        dates,
        lossValues,
        "System Losses",
        true
    );

}


/* =========================================================
   DASHBOARD CHARTS
========================================================= */

function renderDashboardCharts() {

    const dates =
        dailyKPIData.map(
            row => getDateFromRow(row)
        );


    const prValues =
        dailyKPIData.map(
            row => parseNumber(getColumn(row, 21))
        );


    const lossValues =
        dailyKPIData.map(
            row => parseNumber(getColumn(row, 29))
        );


    createDateLineChart(
        "dashboardPRChart",
        dates,
        prValues,
        "Performance Ratio",
        true
    );


    createDateLineChart(
        "dashboardLossChart",
        dates,
        lossValues,
        "System Losses",
        true
    );

}


/* =========================================================
   DATE CHART
========================================================= */

function createDateLineChart(
    canvasId,
    dates,
    values,
    label,
    percentage
) {

    const canvas =
        document.getElementById(canvasId);

    if (!canvas) return;


    destroyChart(canvasId);


    const validData = [];


    for (let i = 0; i < dates.length; i++) {

        if (
            dates[i] instanceof Date &&
            !isNaN(dates[i].getTime()) &&
            Number.isFinite(values[i])
        ) {

            validData.push({
                x: dates[i],
                y: values[i]
            });

        }

    }


    validData.sort(
        (a, b) => a.x - b.x
    );


    charts[canvasId] =
        new Chart(canvas, {

            type: "line",

            data: {

                datasets: [

                    {
                        label,

                        data: validData,

                        borderWidth: 2,

                        pointRadius: 3,

                        pointHoverRadius: 5,

                        tension: 0.25,

                        fill: false
                    }

                ]

            },


            options: {

                responsive: true,

                maintainAspectRatio: false,

                animation: false,


                interaction: {

                    mode: "nearest",

                    intersect: false

                },


                scales: {

                    x: {

                        type: "linear",

                        title: {
                            display: false
                        },


                        ticks: {

                            autoSkip: false,

                            maxRotation: 0,

                            callback: function(value) {

                                return formatDateTick(
                                    value,
                                    this.chart
                                );

                            }

                        },


                        grid: {
                            display: false
                        }

                    },


                    y: {

                        beginAtZero: false,

                        ticks: {

                            callback: value => {

                                if (percentage) {

                                    return formatPercent(
                                        value
                                    );

                                }

                                return value;

                            }

                        }

                    }

                },


                plugins: {

                    legend: {
                        display: false
                    },


                    tooltip: {

                        callbacks: {

                            title: items => {

                                if (!items.length)
                                    return "";

                                const raw =
                                    items[0].raw;

                                if (
                                    !raw ||
                                    !(raw.x instanceof Date)
                                ) {
                                    return "";
                                }

                                return formatFullDate(
                                    raw.x
                                );

                            },


                            label: context => {

                                const value =
                                    context.parsed.y;

                                return percentage
                                    ? `${label}: ${formatPercent(value)}`
                                    : `${label}: ${value}`;

                            }

                        }

                    }

                }

            }

        });

}


/* =========================================================
   DATE TICK FORMAT
========================================================= */

function formatDateTick(value, chart) {

    const xScale =
        chart.scales.x;

    if (!xScale) return "";


    const min =
        xScale.min;

    const max =
        xScale.max;


    const date =
        new Date(value);


    if (isNaN(date.getTime()))
        return "";


    const day =
        date.getDate();


    const month =
        date.getMonth();


    /*
       IMPORTANT:

       For a normal month:
       1, 3, 5, 7, 9, 11...
       or
       2, 4, 6, 8, 10...

       We deliberately calculate the interval
       instead of allowing Chart.js to generate
       duplicate category labels.
    */


    const totalDays =
        Math.round(
            (max - min) /
            (24 * 60 * 60 * 1000)
        );


    let interval = 2;


    if (totalDays <= 7) {
        interval = 1;
    }
    else if (totalDays <= 15) {
        interval = 2;
    }
    else if (totalDays <= 31) {
        interval = 2;
    }
    else if (totalDays <= 90) {
        interval = 7;
    }
    else {
        interval = 14;
    }


    /*
       Only display ticks at the selected interval.
       This prevents:
       29, 31, 31, 29
       type glitches.
    */

    const firstDay =
        new Date(min).getDate();


    const offset =
        day - firstDay;


    if (
        offset % interval !== 0 &&
        day !== new Date(min).getDate() &&
        day !== new Date(max).getDate()
    ) {

        return "";

    }


    return String(day);

}


/* =========================================================
   PA TIMELINE
========================================================= */

function renderPACChart() {

    const canvas =
        document.getElementById("paChart");

    if (!canvas) return;


    destroyChart("paChart");


    const records =
        extractPARecords();


    if (!records.length) {

        clearCanvasMessage(
            canvas,
            "No PA breakdown records available"
        );

        return;
    }


    /*
       Horizontal floating bars.

       Chart.js does not require a date adapter here.
       Dates are converted into timestamps.
    */


    const labels =
        records.map(
            record =>
                record.equipment || "Equipment"
        );


    const data =
        records.map(record => ({

            x: [
                record.start.getTime(),
                record.end.getTime()
            ],

            y:
                record.equipment ||
                "Equipment"

        }));


    charts.paChart =
        new Chart(canvas, {

            type: "bar",

            data: {

                datasets: [

                    {
                        label: "Breakdown",

                        data,

                        parsing: false,

                        borderWidth: 0,

                        barPercentage: 0.7,

                        categoryPercentage: 0.8
                    }

                ]

            },


            options: {

                indexAxis: "y",

                responsive: true,

                maintainAspectRatio: false,

                animation: false,


                scales: {

                    x: {

                        type: "linear",

                        ticks: {

                            maxRotation: 0,

                            callback: value => {

                                const date =
                                    new Date(value);

                                if (
                                    isNaN(
                                        date.getTime()
                                    )
                                ) {
                                    return "";
                                }

                                return formatShortDateTime(
                                    date
                                );

                            }

                        }

                    },

                    y: {

                        type: "category",

                        labels

                    }

                },


                plugins: {

                    legend: {
                        display: false
                    },


                    tooltip: {

                        callbacks: {

                            title: context => {

                                const index =
                                    context[0].dataIndex;

                                return records[index]
                                    ?.equipment ||
                                    "Breakdown";

                            },


                            label: context => {

                                const record =
                                    records[
                                        context.dataIndex
                                    ];

                                if (!record)
                                    return "";

                                return [
                                    `Start: ${formatFullDateTime(record.start)}`,
                                    `End: ${formatFullDateTime(record.end)}`,
                                    `Duration: ${record.duration}`
                                ];

                            }

                        }

                    }

                }

            }

        });

}


/* =========================================================
   EXTRACT PA
========================================================= */

function extractPARecords() {

    const result = [];


    paData.forEach(row => {

        const start =
            findDateByKeywords(
                row,
                [
                    "start",
                    "from",
                    "breakdown start",
                    "outage start"
                ]
            );


        const end =
            findDateByKeywords(
                row,
                [
                    "end",
                    "to",
                    "breakdown end",
                    "outage end"
                ]
            );


        if (
            start &&
            end &&
            end > start
        ) {

            const equipment =
                findStringByKeywords(
                    row,
                    [
                        "equipment",
                        "asset",
                        "plant",
                        "unit",
                        "description"
                    ]
                );


            result.push({

                start,

                end,

                equipment:
                    equipment ||
                    "PA Breakdown",

                duration:
                    formatDuration(
                        end - start
                    )

            });

        }

    });


    return result;

}


/* =========================================================
   CURTAILMENT
========================================================= */

function renderCurtailmentChart() {

    const canvas =
        document.getElementById(
            "curtailmentChart"
        );

    if (!canvas) return;


    destroyChart(
        "curtailmentChart"
    );


    const points = [];


    curtailmentData.forEach(row => {

        const date =
            getDateFromRow(row);


        const value =
            findNumericByKeywords(
                row,
                [
                    "curtailment",
                    "loss",
                    "mw",
                    "generation loss",
                    "restricted"
                ]
            );


        if (
            date &&
            Number.isFinite(value)
        ) {

            points.push({

                x: date,

                y: value

            });

        }

    });


    points.sort(
        (a, b) => a.x - b.x
    );


    const summary =
        document.getElementById(
            "curtailmentSummary"
        );


    if (!points.length) {

        if (summary) {
            summary.textContent =
                "No curtailment records found";
        }


        clearCanvasMessage(
            canvas,
            "No curtailment data available"
        );

        return;
    }


    if (summary) {

        const total =
            points.reduce(
                (sum, point) =>
                    sum + point.y,
                0
            );

        summary.textContent =
            `${points.length} records · Total ${formatNumber(total, 2)}`;

    }


    charts.curtailmentChart =
        new Chart(canvas, {

            type: "line",

            data: {

                datasets: [

                    {
                        label: "Curtailment Loss",

                        data: points,

                        borderWidth: 2,

                        pointRadius: 3,

                        tension: 0.2,

                        fill: false
                    }

                ]

            },


            options: {

                responsive: true,

                maintainAspectRatio: false,

                animation: false,


                scales: {

                    x: {

                        type: "linear",

                        ticks: {

                            maxRotation: 0,

                            callback: value =>
                                formatDateTick(
                                    value,
                                    charts.curtailmentChart
                                )

                        }

                    },


                    y: {

                        beginAtZero: true

                    }

                },


                plugins: {

                    legend: {
                        display: false
                    },


                    tooltip: {

                        callbacks: {

                            title: items => {

                                const point =
                                    items[0]?.raw;

                                return point
                                    ? formatFullDate(point.x)
                                    : "";

                            }

                        }

                    }

                }

            }

        });

}


/* =========================================================
   ENERGY
========================================================= */

function renderEnergyChart() {

    const canvas =
        document.getElementById(
            "energyChart"
        );

    if (!canvas) return;


    destroyChart("energyChart");


    const records =
        extractEnergyRecords();


    if (!records.length) {

        clearCanvasMessage(
            canvas,
            "No Annual KPI energy data available"
        );

        return;
    }


    const labels =
        records.map(
            record => record.month
        );


    const budget =
        records.map(
            record => record.budget
        );


    const measured =
        records.map(
            record => record.measured
        );


    const totalBudget =
        budget.reduce(
            (sum, value) =>
                sum + value,
            0
        );


    const totalMeasured =
        measured.reduce(
            (sum, value) =>
                sum + value,
            0
        );


    const variance =
        totalMeasured - totalBudget;


    setText(
        "totalBudget",
        formatNumber(totalBudget, 2)
    );


    setText(
        "totalMeasured",
        formatNumber(totalMeasured, 2)
    );


    setText(
        "energyVariance",
        formatNumber(variance, 2)
    );


    charts.energyChart =
        new Chart(canvas, {

            type: "bar",

            data: {

                labels,

                datasets: [

                    {
                        label: "Budgeted Energy",

                        data: budget,

                        borderWidth: 1
                    },

                    {
                        label: "Measured Energy",

                        data: measured,

                        borderWidth: 1
                    }

                ]

            },


            options: {

                responsive: true,

                maintainAspectRatio: false,

                animation: false,


                scales: {

                    y: {

                        beginAtZero: true

                    },

                    x: {

                        ticks: {

                            maxRotation: 0

                        }

                    }

                },


                plugins: {

                    legend: {

                        display: true,

                        position: "top"

                    }

                }

            }

        });

}


/* =========================================================
   EXTRACT ENERGY DATA
========================================================= */

function extractEnergyRecords() {

    const records = [];


    /*
       Annual_KPI:
       E10:E21 = Budgeted Energy
       F10:F21 = Measured Energy

       sheet_to_json() with normal headers means
       the actual property names can vary.

       Therefore we also support positional
       extraction.
    */


    if (!workbook) return records;


    const sheetName =
        workbook.SheetNames.find(
            name =>
                name.toLowerCase() ===
                "annual_kpi"
        );


    if (!sheetName) return records;


    const worksheet =
        workbook.Sheets[sheetName];


    for (let rowNumber = 10; rowNumber <= 21; rowNumber++) {

        const budgetCell =
            worksheet[`E${rowNumber}`];


        const measuredCell =
            worksheet[`F${rowNumber}`];


        const budget =
            parseNumber(
                budgetCell?.v
            );


        const measured =
            parseNumber(
                measuredCell?.v
            );


        if (
            Number.isFinite(budget) ||
            Number.isFinite(measured)
        ) {

            records.push({

                month:
                    getMonthName(
                        rowNumber - 10
                    ),

                budget:
                    Number.isFinite(budget)
                        ? budget
                        : 0,

                measured:
                    Number.isFinite(measured)
                        ? measured
                        : 0

            });

        }

    }


    return records;

}


/* =========================================================
   MONTH NAME
========================================================= */

function getMonthName(index) {

    const months = [
        "Jan",
        "Feb",
        "Mar",
        "Apr",
        "May",
        "Jun",
        "Jul",
        "Aug",
        "Sep",
        "Oct",
        "Nov",
        "Dec"
    ];


    return months[index] || `Month ${index + 1}`;

}


/* =========================================================
   COLUMN HELPERS
========================================================= */

function getColumn(row, zeroBasedIndex) {

    if (!row) return null;


    const keys =
        Object.keys(row);


    return keys[zeroBasedIndex] !== undefined
        ? row[keys[zeroBasedIndex]]
        : null;

}


function getColumnValue(rows, zeroBasedIndex) {

    if (!rows.length) return null;


    const last =
        rows[rows.length - 1];


    return getColumn(
        last,
        zeroBasedIndex
    );

}


/* =========================================================
   DATE EXTRACTION
========================================================= */

function getDateFromRow(row) {

    if (!row) return null;


    /*
       First look for columns whose names
       explicitly indicate date.
    */

    const keys =
        Object.keys(row);


    for (const key of keys) {

        const lower =
            key.toLowerCase();


        if (
            lower.includes("date") ||
            lower === "day" ||
            lower.includes("timestamp")
        ) {

            const date =
                parseDate(row[key]);


            if (date) return date;

        }

    }


    /*
       Otherwise inspect values.
    */

    for (const key of keys) {

        const date =
            parseDate(row[key]);


        if (date) return date;

    }


    return null;

}


/* =========================================================
   DATE PARSER
========================================================= */

function parseDate(value) {

    if (
        value === null ||
        value === undefined ||
        value === ""
    ) {
        return null;
    }


    if (value instanceof Date) {

        if (!isNaN(value.getTime())) {
            return value;
        }

        return null;
    }


    /*
       Excel serial date.
    */

    if (
        typeof value === "number" &&
        value > 20000 &&
        value < 80000
    ) {

        const date =
            XLSX.SSF.parse_date_code(
                value
            );


        if (!date) return null;


        return new Date(
            date.y,
            date.m - 1,
            date.d,
            date.H || 0,
            date.M || 0,
            date.S || 0
        );

    }


    if (typeof value !== "string")
        return null;


    const text =
        value.trim();


    if (!text) return null;


    /*
       DD/MM/YYYY
       DD-MM-YYYY
    */

    const indian =
        text.match(
            /^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})/
        );


    if (indian) {

        const day =
            Number(indian[1]);

        const month =
            Number(indian[2]) - 1;

        const year =
            Number(indian[3]);


        const date =
            new Date(
                year,
                month,
                day
            );


        if (
            date.getFullYear() === year &&
            date.getMonth() === month &&
            date.getDate() === day
        ) {

            return date;

        }

    }


    /*
       ISO / normal JS date.
    */

    const parsed =
        new Date(text);


    if (
        !isNaN(
            parsed.getTime()
        )
    ) {

        return parsed;

    }


    return null;

}


/* =========================================================
   FIND DATE BY COLUMN KEYWORDS
========================================================= */

function findDateByKeywords(
    row,
    keywords
) {

    if (!row) return null;


    const keys =
        Object.keys(row);


    for (const key of keys) {

        const lower =
            key.toLowerCase();


        const match =
            keywords.some(
                keyword =>
                    lower.includes(
                        keyword.toLowerCase()
                    )
            );


        if (match) {

            const date =
                parseDate(
                    row[key]
                );


            if (date) return date;

        }

    }


    return null;

}


/* =========================================================
   FIND STRING
========================================================= */

function findStringByKeywords(
    row,
    keywords
) {

    if (!row) return null;


    const keys =
        Object.keys(row);


    for (const key of keys) {

        const lower =
            key.toLowerCase();


        const match =
            keywords.some(
                keyword =>
                    lower.includes(
                        keyword.toLowerCase()
                    )
            );


        if (
            match &&
            typeof row[key] === "string"
        ) {

            const value =
                row[key].trim();


            if (value)
                return value;

        }

    }


    return null;

}


/* =========================================================
   FIND NUMERIC
========================================================= */

function findNumericByKeywords(
    row,
    keywords
) {

    if (!row) return null;


    const keys =
        Object.keys(row);


    for (const key of keys) {

        const lower =
            key.toLowerCase();


        const match =
            keywords.some(
                keyword =>
                    lower.includes(
                        keyword.toLowerCase()
                    )
            );


        if (match) {

            const value =
                parseNumber(
                    row[key]
                );


            if (
                Number.isFinite(value)
            ) {

                return value;

            }

        }

    }


    return null;

}


/* =========================================================
   FIND LATEST VALUE
========================================================= */

function findLatestValue(
    rows,
    keywords
) {

    if (!rows.length)
        return null;


    for (
        let i = rows.length - 1;
        i >= 0;
        i--
    ) {

        const value =
            findNumericByKeywords(
                rows[i],
                keywords
            );


        if (
            Number.isFinite(value)
        ) {

            return value;

        }

    }


    return null;

}


/* =========================================================
   NUMBER PARSER
========================================================= */

function parseNumber(value) {

    if (
        value === null ||
        value === undefined ||
        value === ""
    ) {
        return NaN;
    }


    if (typeof value === "number")
        return value;


    if (typeof value === "string") {

        let text =
            value
                .replace(/,/g, "")
                .replace(/%/g, "")
                .trim();


        if (!text)
            return NaN;


        const number =
            Number(text);


        return Number.isFinite(number)
            ? number
            : NaN;

    }


    return NaN;

}


/* =========================================================
   PERCENTAGE FORMAT
========================================================= */

function formatPercent(value) {

    const number =
        parseNumber(value);


    if (!Number.isFinite(number))
        return "—";


    /*
       If Excel stores 0.842,
       display 84.2%.

       If Excel stores 84.2,
       display 84.2%.
    */

    const percentage =
        Math.abs(number) <= 1
            ? number * 100
            : number;


    return `${percentage.toFixed(1)}%`;

}


/* =========================================================
   NUMBER FORMAT
========================================================= */

function formatNumber(
    value,
    decimals = 2
) {

    const number =
        parseNumber(value);


    if (!Number.isFinite(number))
        return "—";


    return number.toLocaleString(
        "en-IN",
        {
            minimumFractionDigits:
                decimals,

            maximumFractionDigits:
                decimals
        }
    );

}


/* =========================================================
   DATE FORMATTING
========================================================= */

function formatFullDate(date) {

    if (!(date instanceof Date))
        date = new Date(date);


    if (isNaN(date.getTime()))
        return "";


    return date.toLocaleDateString(
        "en-IN",
        {
            day: "2-digit",
            month: "short",
            year: "numeric"
        }
    );

}


function formatFullDateTime(date) {

    if (!(date instanceof Date))
        date = new Date(date);


    if (isNaN(date.getTime()))
        return "";


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


function formatShortDateTime(date) {

    if (!(date instanceof Date))
        date = new Date(date);


    if (isNaN(date.getTime()))
        return "";


    return date.toLocaleDateString(
        "en-IN",
        {
            day: "2-digit",
            month: "short"
        }
    );

}


/* =========================================================
   DURATION
========================================================= */

function formatDuration(
    milliseconds
) {

    const minutes =
        Math.round(
            milliseconds / 60000
        );


    const hours =
        Math.floor(
            minutes / 60
        );


    const remaining =
        minutes % 60;


    if (hours === 0)
        return `${remaining} min`;


    return `${hours}h ${remaining}m`;

}


/* =========================================================
   SET TEXT
========================================================= */

function setText(
    id,
    value
) {

    const element =
        document.getElementById(id);


    if (element) {
        element.textContent = value;
    }

}


/* =========================================================
   DESTROY CHART
========================================================= */

function destroyChart(id) {

    if (charts[id]) {

        charts[id].destroy();

        delete charts[id];

    }

}


/* =========================================================
   CLEAR CANVAS MESSAGE
========================================================= */

function clearCanvasMessage(
    canvas,
    message
) {

    const context =
        canvas.getContext("2d");


    if (!context) return;


    context.clearRect(
        0,
        0,
        canvas.width,
        canvas.height
    );


    context.save();


    context.textAlign = "center";

    context.textBaseline = "middle";


    context.font =
        "12px Inter, Arial, sans-serif";


    context.fillStyle =
        "#829095";


    context.fillText(
        message,
        canvas.width / 2,
        canvas.height / 2
    );


    context.restore();

}


/* =========================================================
   RESET APPLICATION
========================================================= */

function resetApplication() {

    workbook = null;

    uploadedFile = null;

    dashboardData = [];
    dailyKPIData = [];
    paData = [];
    curtailmentData = [];
    annualKPIData = [];


    Object.keys(charts).forEach(
        id => destroyChart(id)
    );


    const fileInput =
        document.getElementById("dgrFile");


    if (fileInput) {
        fileInput.value = "";
    }


    const fileInfo =
        document.getElementById("fileInfo");


    if (fileInfo) {
        fileInfo.classList.add("hidden");
    }


    const workbookStatus =
        document.getElementById(
            "workbookStatus"
        );


    if (workbookStatus) {
        workbookStatus.classList.add("hidden");
    }


    const emptyState =
        document.getElementById(
            "emptyState"
        );


    if (emptyState) {
        emptyState.classList.remove("hidden");
    }


    const sidebarFileName =
        document.getElementById(
            "sidebarFileName"
        );


    if (sidebarFileName) {

        sidebarFileName.textContent =
            "No DGR uploaded";

    }


    setText(
        "dashboardPA",
        "—"
    );

    setText(
        "dashboardPR",
        "—"
    );

    setText(
        "dashboardLoss",
        "—"
    );

    setText(
        "dashboardHours",
        "—"
    );

    setText(
        "totalBudget",
        "—"
    );

    setText(
        "totalMeasured",
        "—"
    );

    setText(
        "energyVariance",
        "—"
    );


    hideDashboardUntilUpload();


    setStatus(
        "Upload a DGR to generate the analytics."
    );

}


/* =========================================================
   EXTRA SAFETY:
   PREVENT MULTIPLE UPLOAD LABELS FROM BEING GENERATED
========================================================= */

/*
   app.js does NOT create any "Upload DGR" buttons.

   The only upload controls are the two already present
   in HTML:

   1. Header → Upload DGR
   2. Empty state → Upload DGR

   The large drag-and-drop area is clickable but does
   not create another button.
*/


/* =========================================================
   END
========================================================= */
