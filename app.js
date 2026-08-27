/* =========================================================
   SOLAR DGR ANALYTICS
   app.js
========================================================= */

"use strict";

/* =========================================================
   GLOBAL STATE
========================================================= */

let workbook = null;

let charts = {
    dashboardPR: null,
    dashboardLoss: null,
    pa: null,
    pr: null,
    hours: null,
    loss: null,
    curtailment: null,
    energy: null
};

let currentFile = null;


/* =========================================================
   DOM HELPERS
========================================================= */

const $ = (id) => document.getElementById(id);

function setText(id, value) {
    const el = $(id);

    if (el) {
        el.textContent = value;
    }
}

function show(id) {
    const el = $(id);

    if (el) {
        el.classList.remove("hidden");
    }
}

function hide(id) {
    const el = $(id);

    if (el) {
        el.classList.add("hidden");
    }
}


/* =========================================================
   DOM ELEMENTS
========================================================= */

const dgrFile = $("dgrFile");
const dropZone = $("dropZone");
const fileInfo = $("fileInfo");
const workbookStatus = $("workbookStatus");
const sheetBadges = $("sheetBadges");
const removeFile = $("removeFile");
const emptyState = $("emptyState");
const statusText = $("statusText");
const sidebarFileName = $("sidebarFileName");


/* =========================================================
   INITIALISE
========================================================= */

document.addEventListener("DOMContentLoaded", () => {

    initialiseNavigation();
    initialiseUpload();

    /*
       Hide analytics sections until a DGR is uploaded.
       The upload area itself remains visible.
    */

    hideAnalyticsSections();

});


/* =========================================================
   NAVIGATION
========================================================= */

function initialiseNavigation() {

    const navItems = document.querySelectorAll(".nav-item");

    navItems.forEach((button) => {

        button.addEventListener("click", () => {

            const targetId = button.dataset.target;
            const target = $(targetId);

            if (!target) {
                return;
            }

            navItems.forEach((item) => {
                item.classList.remove("active");
            });

            button.classList.add("active");

            target.scrollIntoView({
                behavior: "smooth",
                block: "start"
            });

        });

    });

}


/* =========================================================
   UPLOAD INITIALISATION
========================================================= */

function initialiseUpload() {

    if (!dgrFile) {
        console.error("DGR file input not found.");
        return;
    }

    /*
       Normal file selection
    */

    dgrFile.addEventListener("change", async (event) => {

        const files = event.target.files;

        if (!files || !files.length) {
            return;
        }

        await processDGRFile(files[0]);

    });


    /*
       Drag over
    */

    if (dropZone) {

        dropZone.addEventListener("dragover", (event) => {

            event.preventDefault();

            dropZone.classList.add("dragging");

        });


        /*
           Drag leave
        */

        dropZone.addEventListener("dragleave", () => {

            dropZone.classList.remove("dragging");

        });


        /*
           Drop
        */

        dropZone.addEventListener("drop", async (event) => {

            event.preventDefault();

            dropZone.classList.remove("dragging");

            const files = event.dataTransfer.files;

            if (!files || !files.length) {
                return;
            }

            await processDGRFile(files[0]);

        });


        /*
           Clicking the upload panel opens file picker
        */

        dropZone.addEventListener("click", () => {

            dgrFile.click();

        });

    }


    /*
       Remove uploaded file
    */

    if (removeFile) {

        removeFile.addEventListener("click", () => {

            resetDashboard();

        });

    }

}


/* =========================================================
   PROCESS DGR FILE
========================================================= */

async function processDGRFile(file) {

    if (!file) {
        return;
    }

    const extension = file.name
        .split(".")
        .pop()
        .toLowerCase();

    const allowed = ["xlsx", "xls", "csv"];

    if (!allowed.includes(extension)) {

        alert(
            "Please upload a valid Excel or CSV DGR file.\n\n" +
            "Supported formats: .xlsx, .xls, .csv"
        );

        return;
    }


    try {

        currentFile = file;

        setText(
            "statusText",
            "Reading DGR workbook..."
        );


        /*
           Read file as ArrayBuffer
        */

        const arrayBuffer = await file.arrayBuffer();


        /*
           XLSX is provided by SheetJS in HTML.
        */

        if (typeof XLSX === "undefined") {

            throw new Error(
                "SheetJS library was not loaded. " +
                "Please check the XLSX script in your HTML."
            );

        }


        /*
           Read workbook
        */

        workbook = XLSX.read(arrayBuffer, {
            type: "array",
            cellDates: true,
            cellNF: true,
            cellText: true
        });


        if (
            !workbook ||
            !workbook.SheetNames ||
            workbook.SheetNames.length === 0
        ) {

            throw new Error(
                "No worksheets were found in the uploaded file."
            );

        }


        /*
           Update file information
        */

        setText(
            "fileName",
            file.name
        );

        setText(
            "fileSheets",
            `${workbook.SheetNames.length} worksheet(s) detected`
        );

        setText(
            "sidebarFileName",
            file.name
        );


        /*
           Display workbook information
        */

        renderWorkbookStatus();


        /*
           Hide duplicate upload prompts.
        */

        hide("emptyState");


        /*
           Show uploaded file information.
        */

        show("fileInfo");
        show("workbookStatus");


        /*
           Show analytics
        */

        showAnalyticsSections();


        /*
           Analyse workbook
        */

        analyseWorkbook();


        setText(
            "statusText",
            `DGR loaded successfully · ${file.name}`
        );


    } catch (error) {

        console.error("DGR upload error:", error);

        workbook = null;

        setText(
            "statusText",
            "Unable to read the uploaded DGR."
        );

        alert(
            "The DGR could not be read.\n\n" +
            "Please make sure the file is a valid Excel/CSV file."
        );

    }

}


/* =========================================================
   WORKBOOK STATUS
========================================================= */

function renderWorkbookStatus() {

    if (!sheetBadges || !workbook) {
        return;
    }

    sheetBadges.innerHTML = "";


    const requiredSheets = [
        "Dashboard",
        "Annual_KPI",
        "Daily_KPI",
        "PA",
        "Curtailment"
    ];


    const actualSheets = workbook.SheetNames.map(
        normalizeSheetName
    );


    requiredSheets.forEach((required) => {

        const badge = document.createElement("div");

        badge.className = "sheet-badge";


        const found = actualSheets.some(
            sheet =>
                sheet === normalizeSheetName(required)
        );


        if (found) {

            badge.textContent = `${required} ✓`;

        } else {

            badge.textContent = `${required} — Not found`;

            badge.classList.add("missing");

        }


        sheetBadges.appendChild(badge);

    });

}


/* =========================================================
   SHEET NAME NORMALISATION
========================================================= */

function normalizeSheetName(name) {

    return String(name || "")
        .trim()
        .toLowerCase()
        .replace(/[\s_-]+/g, "");

}


/* =========================================================
   FIND SHEET
========================================================= */

function getSheet(possibleNames) {

    if (!workbook) {
        return null;
    }


    const wanted = possibleNames.map(
        normalizeSheetName
    );


    const actualName = workbook.SheetNames.find(
        name =>
            wanted.includes(
                normalizeSheetName(name)
            )
    );


    if (!actualName) {
        return null;
    }


    return workbook.Sheets[actualName];

}


/* =========================================================
   SHEET TO JSON
========================================================= */

function sheetToRows(sheet) {

    if (!sheet) {
        return [];
    }


    return XLSX.utils.sheet_to_json(
        sheet,
        {
            header: 1,
            defval: null,
            raw: true,
            blankrows: false
        }
    );

}


/* =========================================================
   WORKBOOK ANALYSIS
========================================================= */

function analyseWorkbook() {

    /*
       Destroy old charts first.
    */

    destroyCharts();


    /*
       Dashboard
    */

    analyseDailyKPI();


    /*
       PA
    */

    analysePA();


    /*
       Curtailment
    */

    analyseCurtailment();


    /*
       Annual energy
    */

    analyseAnnualEnergy();

}


/* =========================================================
   DAILY KPI
========================================================= */

function analyseDailyKPI() {

    const sheet = getSheet([
        "Daily_KPI",
        "Daily KPI",
        "DailyKPI"
    ]);


    if (!sheet) {

        setText("dashboardPR", "—");
        setText("dashboardLoss", "—");
        setText("dashboardHours", "—");

        return;

    }


    const rows = sheetToRows(sheet);


    /*
       Column indexes:
       I  = 9  -> index 8
       V  = 22 -> index 21
       AD = 30 -> index 29

       Excel is 1-based.
       JavaScript arrays are 0-based.
    */

    const dateIndex = findDateColumn(rows);

    const hoursIndex = 8;
    const prIndex = 21;
    const lossIndex = 29;


    const records = [];


    for (let i = 0; i < rows.length; i++) {

        const row = rows[i];

        if (!row) {
            continue;
        }


        const date = parseExcelDate(
            dateIndex >= 0 ? row[dateIndex] : null
        );


        const pr = parseNumber(
            row[prIndex]
        );


        const hours = parseNumber(
            row[hoursIndex]
        );


        const loss = parseNumber(
            row[lossIndex]
        );


        /*
           Ignore completely empty records.
        */

        if (
            date === null &&
            pr === null &&
            hours === null &&
            loss === null
        ) {
            continue;
        }


        records.push({
            date,
            pr,
            hours,
            loss
        });

    }


    /*
       Remove probable header rows / invalid dates
       when possible.
    */

    const usableRecords = records.filter(
        record =>
            record.date !== null ||
            record.pr !== null ||
            record.hours !== null ||
            record.loss !== null
    );


    if (!usableRecords.length) {
        return;
    }


    /*
       Sort chronologically.
    */

    usableRecords.sort(
        (a, b) =>
            (a.date || 0) - (b.date || 0)
    );


    /*
       Latest available values
    */

    const latest = [...usableRecords]
        .reverse()
        .find(
            row =>
                row.pr !== null ||
                row.hours !== null ||
                row.loss !== null
        );


    if (latest) {

        setText(
            "dashboardPR",
            formatPercent(latest.pr)
        );

        setText(
            "dashboardHours",
            formatNumber(latest.hours)
        );

        setText(
            "dashboardLoss",
            formatPercent(latest.loss)
        );

    }


    /*
       Create charts
    */

    createPerformanceCharts(
        usableRecords
    );

}


/* =========================================================
   PERFORMANCE CHARTS
========================================================= */

function createPerformanceCharts(records) {

    const labels = records.map(
        record => record.date
    );


    /*
       PR
    */

    const prData = records.map(
        record => record.pr
    );


    createScatterChart(
        "dashboardPRChart",
        "Performance Ratio",
        labels,
        prData,
        true
    );


    createScatterChart(
        "prChart",
        "Performance Ratio",
        labels,
        prData,
        true
    );


    /*
       Hours
    */

    const hoursData = records.map(
        record => record.hours
    );


    createLineChart(
        "hoursChart",
        "Operating Hours",
        labels,
        hoursData,
        false
    );


    /*
       Loss
    */

    const lossData = records.map(
        record => record.loss
    );


    createLineChart(
        "dashboardLossChart",
        "System Losses",
        labels,
        lossData,
        true
    );


    createLineChart(
        "lossChart",
        "System Losses",
        labels,
        lossData,
        true
    );

}


/* =========================================================
   PA ANALYSIS
========================================================= */

function analysePA() {

    const sheet = getSheet([
        "PA",
        "Plant Availability",
        "PA Analysis"
    ]);


    if (!sheet) {
        return;
    }


    const rows = sheetToRows(sheet);


    /*
       Try to identify columns by header names.
    */

    const headerInfo = detectPAColumns(rows);


    const records = [];


    for (
        let i = headerInfo.headerRow + 1;
        i < rows.length;
        i++
    ) {

        const row = rows[i];

        if (!row) {
            continue;
        }


        const start = parseExcelDate(
            row[headerInfo.start]
        );


        const end = parseExcelDate(
            row[headerInfo.end]
        );


        const duration =
            headerInfo.duration >= 0
                ? parseNumber(
                    row[headerInfo.duration]
                )
                : null;


        if (!start && !end) {
            continue;
        }


        records.push({
            start,
            end,
            duration
        });

    }


    if (!records.length) {
        return;
    }


    createPAChart(records);

}


/* =========================================================
   DETECT PA COLUMNS
========================================================= */

function detectPAColumns(rows) {

    const result = {
        headerRow: 0,
        start: 0,
        end: 1,
        duration: 2
    };


    /*
       Search first 15 rows for a useful header.
    */

    for (
        let r = 0;
        r < Math.min(rows.length, 15);
        r++
    ) {

        const row = rows[r] || [];


        for (let c = 0; c < row.length; c++) {

            const value = String(
                row[c] ?? ""
            )
                .trim()
                .toLowerCase();


            if (
                value.includes("start") &&
                (
                    value.includes("time") ||
                    value.includes("date")
                )
            ) {
                result.start = c;
                result.headerRow = r;
            }


            if (
                value.includes("end") &&
                (
                    value.includes("time") ||
                    value.includes("date")
                )
            ) {
                result.end = c;
                result.headerRow = r;
            }


            if (
                value.includes("duration") ||
                value.includes("hours")
            ) {
                result.duration = c;
                result.headerRow = r;
            }

        }

    }


    return result;

}


/* =========================================================
   PA CHART
========================================================= */

function createPAChart(records) {

    const canvas = $("paChart");

    if (!canvas) {
        return;
    }


    /*
       Chart.js does not natively provide a true timeline
       without a time scale plugin.

       We therefore use a horizontal bar representation
       based on duration.
    */

    const labels = records.map(
        (record, index) => {

            if (record.start) {

                return formatDateTime(
                    record.start
                );

            }

            return `Breakdown ${index + 1}`;

        }
    );


    const durations = records.map(
        record => {

            if (record.duration !== null) {
                return record.duration;
            }


            if (
                record.start &&
                record.end
            ) {

                return (
                    record.end -
                    record.start
                ) / 3600000;

            }

            return 0;

        }
    );


    charts.pa = new Chart(
        canvas,
        {
            type: "bar",

            data: {
                labels,

                datasets: [
                    {
                        label: "Duration (hours)",
                        data: durations,

                        borderWidth: 1,

                        borderRadius: 4
                    }
                ]
            },

            options: {

                responsive: true,

                maintainAspectRatio: false,

                animation: false,

                indexAxis: "y",

                plugins: {

                    legend: {
                        display: false
                    },

                    tooltip: {

                        callbacks: {

                            label: function(context) {

                                const value =
                                    context.raw;

                                return `${value.toFixed(2)} hours`;

                            }

                        }

                    }

                },

                scales: {

                    x: {
                        beginAtZero: true,

                        title: {
                            display: true,
                            text: "Duration (hours)"
                        }
                    },

                    y: {

                        ticks: {
                            autoSkip: true
                        }

                    }

                }

            }

        }
    );

}


/* =========================================================
   CURTAILMENT
========================================================= */

function analyseCurtailment() {

    const sheet = getSheet([
        "Curtailment",
        "Curtailment Records",
        "Grid Curtailment"
    ]);


    if (!sheet) {

        setText(
            "curtailmentSummary",
            "Curtailment worksheet not found"
        );

        return;
    }


    const rows = sheetToRows(sheet);


    if (!rows.length) {
        return;
    }


    const dateIndex = findDateColumn(rows);


    /*
       Find a likely MW / loss column.
    */

    let valueIndex = findColumnByKeywords(
        rows,
        [
            "curtailment",
            "mw",
            "loss",
            "generation loss"
        ]
    );


    if (valueIndex < 0) {

        /*
           Fall back to second column.
        */

        valueIndex = 1;

    }


    const records = [];


    for (let i = 0; i < rows.length; i++) {

        const row = rows[i];

        if (!row) {
            continue;
        }


        const date =
            dateIndex >= 0
                ? parseExcelDate(row[dateIndex])
                : null;


        const value =
            parseNumber(row[valueIndex]);


        if (
            date === null ||
            value === null
        ) {
            continue;
        }


        records.push({
            date,
            value
        });

    }


    if (!records.length) {

        setText(
            "curtailmentSummary",
            "No valid curtailment records found"
        );

        return;

    }


    records.sort(
        (a, b) =>
            a.date - b.date
    );


    const total = records.reduce(
        (sum, item) =>
            sum + item.value,
        0
    );


    setText(
        "curtailmentSummary",
        `${records.length} records · Total ${formatNumber(total)}`
    );


    createLineChart(
        "curtailmentChart",
        "Curtailment",
        records.map(
            item => item.date
        ),
        records.map(
            item => item.value
        ),
        false,
        true
    );

}


/* =========================================================
   ANNUAL ENERGY
========================================================= */

function analyseAnnualEnergy() {

    const sheet = getSheet([
        "Annual_KPI",
        "Annual KPI",
        "AnnualKPI"
    ]);


    if (!sheet) {
        return;
    }


    const rows = sheetToRows(sheet);


    /*
       Excel:
       E10:E21 = Budgeted
       F10:F21 = Measured

       JavaScript:
       E = index 4
       F = index 5
    */

    const budgetIndex = 4;
    const measuredIndex = 5;


    const labels = [];
    const budget = [];
    const measured = [];


    /*
       Rows 10-21 are Excel rows.
       JavaScript row indexes are 9-20.
    */

    for (
        let i = 9;
        i <= 20 && i < rows.length;
        i++
    ) {

        const row = rows[i] || [];


        const budgetValue =
            parseNumber(row[budgetIndex]);


        const measuredValue =
            parseNumber(row[measuredIndex]);


        if (
            budgetValue === null &&
            measuredValue === null
        ) {
            continue;
        }


        /*
           Try to find month name from columns
           before E.
        */

        let label = null;


        for (let c = 0; c < 4; c++) {

            const possible =
                row[c];

            if (
                possible !== null &&
                possible !== undefined &&
                String(possible).trim() !== ""
            ) {

                label = String(
                    possible
                );

                break;

            }

        }


        if (!label) {

            label = `Month ${labels.length + 1}`;

        }


        labels.push(label);

        budget.push(
            budgetValue ?? 0
        );

        measured.push(
            measuredValue ?? 0
        );

    }


    if (!labels.length) {
        return;
    }


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
        totalMeasured -
        totalBudget;


    setText(
        "totalBudget",
        formatNumber(totalBudget)
    );


    setText(
        "totalMeasured",
        formatNumber(totalMeasured)
    );


    setText(
        "energyVariance",
        formatNumber(variance)
    );


    createEnergyChart(
        labels,
        budget,
        measured
    );

}


/* =========================================================
   ENERGY CHART
========================================================= */

function createEnergyChart(
    labels,
    budget,
    measured
) {

    const canvas = $("energyChart");

    if (!canvas) {
        return;
    }


    charts.energy = new Chart(
        canvas,
        {
            type: "bar",

            data: {

                labels,

                datasets: [

                    {
                        label: "Budgeted Energy",
                        data: budget,

                        borderWidth: 1,

                        borderRadius: 4
                    },

                    {
                        label: "Measured Energy",
                        data: measured,

                        borderWidth: 1,

                        borderRadius: 4
                    }

                ]

            },

            options: {

                responsive: true,

                maintainAspectRatio: false,

                animation: false,

                interaction: {
                    mode: "index",
                    intersect: false
                },

                plugins: {

                    legend: {
                        display: true
                    }

                },

                scales: {

                    x: {
                        ticks: {
                            autoSkip: false
                        }
                    },

                    y: {

                        beginAtZero: true,

                        title: {
                            display: true,
                            text: "Energy"
                        }

                    }

                }

            }

        }
    );

}


/* =========================================================
   CHART HELPERS
========================================================= */

function createLineChart(
    canvasId,
    label,
    dates,
    values,
    percentage = false,
    showTime = false
) {

    const canvas = $(canvasId);

    if (!canvas) {
        return;
    }


    const chartKey =
        Object.keys(charts).find(
            key =>
                canvasId.toLowerCase()
                    .includes(
                        key.toLowerCase()
                    )
        );


    const labels =
        dates.map(
            date => date
        );


    const config = {

        type: "line",

        data: {

            labels,

            datasets: [

                {
                    label,

                    data: values,

                    tension: 0.25,

                    pointRadius: 3,

                    pointHoverRadius: 5,

                    borderWidth: 2,

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

            plugins: {

                legend: {
                    display: false
                },

                tooltip: {

                    callbacks: {

                        title: function(items) {

                            if (
                                !items ||
                                !items.length
                            ) {
                                return "";
                            }


                            const value =
                                items[0].label;


                            if (
                                value instanceof Date
                            ) {

                                return showTime
                                    ? formatDateTime(value)
                                    : formatDate(value);

                            }


                            return value;

                        },

                        label: function(context) {

                            const value =
                                context.raw;


                            if (
                                percentage &&
                                value !== null &&
                                value !== undefined
                            ) {

                                return `${value.toFixed(2)}%`;

                            }


                            return `${value}`;

                        }

                    }

                }

            },

            scales: {

                x: {

                    type: "category",

                    /*
                       The labels themselves are dates,
                       but category labels prevent Chart.js
                       from doing strange numerical date
                       interpolation.
                    */

                    ticks: {

                        autoSkip: true,

                        maxTicksLimit:
                            calculateMaxTicks(
                                dates.length
                            ),

                        callback: function(
                            value,
                            index
                        ) {

                            const date =
                                dates[index];


                            if (
                                !(date instanceof Date)
                            ) {
                                return date;
                            }


                            if (showTime) {

                                return formatDateTime(
                                    date
                                );

                            }


                            return formatDate(
                                date
                            );

                        }

                    }

                },

                y: {

                    beginAtZero: false,

                    ticks: {

                        callback: function(value) {

                            if (percentage) {
                                return `${value}%`;
                            }

                            return value;

                        }

                    }

                }

            }

        }

    };


    charts[chartKey || canvasId] =
        new Chart(
            canvas,
            config
        );

}


/* =========================================================
   SCATTER CHART
========================================================= */

function createScatterChart(
    canvasId,
    label,
    dates,
    values,
    percentage = false
) {

    const canvas = $(canvasId);

    if (!canvas) {
        return;
    }


    const points = [];


    dates.forEach(
        (date, index) => {

            const value =
                values[index];


            if (
                date instanceof Date &&
                value !== null &&
                value !== undefined &&
                Number.isFinite(value)
            ) {

                points.push({
                    x: date,
                    y: value
                });

            }

        }
    );


    const key =
        Object.keys(charts).find(
            item =>
                canvasId.toLowerCase()
                    .includes(
                        item.toLowerCase()
                    )
        );


    charts[key || canvasId] =
        new Chart(
            canvas,
            {

                type: "scatter",

                data: {

                    datasets: [

                        {
                            label,

                            data: points,

                            showLine: true,

                            tension: 0.25,

                            pointRadius: 3,

                            pointHoverRadius: 5,

                            borderWidth: 2
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

                    plugins: {

                        legend: {
                            display: false
                        },

                        tooltip: {

                            callbacks: {

                                title: function(items) {

                                    if (
                                        !items.length
                                    ) {
                                        return "";
                                    }


                                    const raw =
                                        items[0].raw;


                                    return formatDate(
                                        new Date(raw.x)
                                    );

                                },

                                label: function(context) {

                                    const value =
                                        context.parsed.y;


                                    if (percentage) {

                                        return `${value.toFixed(2)}%`;

                                    }


                                    return value;

                                }

                            }

                        }

                    },

                    scales: {

                        x: {

                            type: "linear",

                            /*
                               IMPORTANT:
                               Dates are stored internally as
                               timestamps. We explicitly control
                               the displayed date intervals.
                            */

                            ticks: {

                                autoSkip: true,

                                maxTicksLimit:
                                    calculateMaxTicks(
                                        dates.length
                                    ),

                                callback: function(value) {

                                    const date =
                                        new Date(value);


                                    return formatDate(
                                        date
                                    );

                                }

                            }

                        },

                        y: {

                            ticks: {

                                callback: function(value) {

                                    if (percentage) {
                                        return `${value}%`;
                                    }

                                    return value;

                                }

                            }

                        }

                    }

                }

            }
        );

}


/* =========================================================
   DATE AXIS TICK CALCULATION
========================================================= */

function calculateMaxTicks(dataLength) {

    /*
       The important part for your July problem.

       We do NOT let Chart.js generate random repeated
       date labels.

       Instead:

       1-7 days    -> every 1 day
       8-15 days   -> every 2 days
       16-31 days  -> every 2-3 days
       32-60 days  -> approximately every 7 days
       etc.

       This prevents:
       29, 31, 31, 29
       type behaviour.
    */

    if (dataLength <= 7) {
        return dataLength;
    }

    if (dataLength <= 15) {
        return 8;
    }

    if (dataLength <= 31) {
        return 10;
    }

    if (dataLength <= 60) {
        return 9;
    }

    if (dataLength <= 120) {
        return 10;
    }

    return 12;

}


/* =========================================================
   FIND DATE COLUMN
========================================================= */

function findDateColumn(rows) {

    if (!rows || !rows.length) {
        return -1;
    }


    /*
       First try headers.
    */

    for (
        let r = 0;
        r < Math.min(rows.length, 10);
        r++
    ) {

        const row = rows[r] || [];


        for (
            let c = 0;
            c < row.length;
            c++
        ) {

            const text =
                String(row[c] ?? "")
                    .trim()
                    .toLowerCase();


            if (
                text === "date" ||
                text.includes("date") ||
                text.includes("day")
            ) {

                return c;

            }

        }

    }


    /*
       If there is no obvious header,
       inspect the first few columns for
       Excel dates.
    */

    const maxColumns =
        Math.min(
            15,
            Math.max(
                ...rows.map(
                    row =>
                        row
                            ? row.length
                            : 0
                )
            )
        );


    for (
        let c = 0;
        c < maxColumns;
        c++
    ) {

        let validDates = 0;


        for (
            let r = 0;
            r < Math.min(rows.length, 20);
            r++
        ) {

            if (
                parseExcelDate(
                    rows[r]?.[c]
                ) !== null
            ) {

                validDates++;

            }

        }


        if (validDates >= 3) {
            return c;
        }

    }


    return -1;

}


/* =========================================================
   FIND COLUMN BY KEYWORDS
========================================================= */

function findColumnByKeywords(
    rows,
    keywords
) {

    if (!rows || !rows.length) {
        return -1;
    }


    for (
        let r = 0;
        r < Math.min(rows.length, 10);
        r++
    ) {

        const row = rows[r] || [];


        for (
            let c = 0;
            c < row.length;
            c++
        ) {

            const text =
                String(row[c] ?? "")
                    .trim()
                    .toLowerCase();


            if (
                keywords.some(
                    keyword =>
                        text.includes(
                            keyword.toLowerCase()
                        )
                )
            ) {

                return c;

            }

        }

    }


    return -1;

}


/* =========================================================
   EXCEL DATE PARSER
========================================================= */

function parseExcelDate(value) {

    if (
        value === null ||
        value === undefined ||
        value === ""
    ) {

        return null;

    }


    /*
       Already a JavaScript Date
    */

    if (
        Object.prototype.toString.call(value) ===
        "[object Date]"
    ) {

        if (
            Number.isNaN(
                value.getTime()
            )
        ) {

            return null;

        }

        return value;

    }


    /*
       Excel serial date
    */

    if (
        typeof value === "number" &&
        Number.isFinite(value)
    ) {

        /*
           Excel dates are generally > 1.
           Ignore implausible values.
        */

        if (
            value > 1 &&
            value < 100000
        ) {

            const parsed =
                XLSX.SSF.parse_date_code(
                    value
                );


            if (parsed) {

                return new Date(
                    parsed.y,
                    parsed.m - 1,
                    parsed.d,
                    parsed.H || 0,
                    parsed.M || 0,
                    parsed.S || 0
                );

            }

        }

        return null;

    }


    /*
       String date
    */

    if (typeof value === "string") {

        const trimmed =
            value.trim();


        if (!trimmed) {
            return null;
        }


        /*
           DD/MM/YYYY
           DD-MM-YYYY
        */

        const match =
            trimmed.match(
                /^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})/
            );


        if (match) {

            const day =
                Number(match[1]);

            const month =
                Number(match[2]) - 1;

            const year =
                Number(match[3]);


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
           Let JavaScript handle ISO-like dates.
        */

        const parsed =
            new Date(trimmed);


        if (
            !Number.isNaN(
                parsed.getTime()
            )
        ) {

            return parsed;

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

        return null;

    }


    if (typeof value === "number") {

        return Number.isFinite(value)
            ? value
            : null;

    }


    let text =
        String(value)
            .trim()
            .replace(/,/g, "")
            .replace(/%/g, "");


    if (!text) {
        return null;
    }


    const number =
        Number(text);


    return Number.isFinite(number)
        ? number
        : null;

}


/* =========================================================
   FORMATTING
========================================================= */

function formatNumber(value) {

    if (
        value === null ||
        value === undefined ||
        !Number.isFinite(value)
    ) {

        return "—";

    }


    return value.toLocaleString(
        "en-IN",
        {
            maximumFractionDigits: 2
        }
    );

}


function formatPercent(value) {

    if (
        value === null ||
        value === undefined ||
        !Number.isFinite(value)
    ) {

        return "—";

    }


    return `${value.toFixed(2)}%`;

}


function formatDate(date) {

    if (!(date instanceof Date)) {
        return String(date ?? "");
    }


    return date.toLocaleDateString(
        "en-GB",
        {
            day: "2-digit",
            month: "short"
        }
    );

}


function formatDateTime(date) {

    if (!(date instanceof Date)) {
        return String(date ?? "");
    }


    return date.toLocaleString(
        "en-GB",
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
   SHOW / HIDE ANALYTICS
========================================================= */

function hideAnalyticsSections() {

    /*
       Do NOT hide the upload panel.

       These are the actual analytics sections.
    */

    [
        "dashboardSection",
        "paSection",
        "performanceSection",
        "curtailmentSection",
        "energySection"
    ].forEach(hide);

}


function showAnalyticsSections() {

    [
        "dashboardSection",
        "paSection",
        "performanceSection",
        "curtailmentSection",
        "energySection"
    ].forEach(show);

}


/* =========================================================
   DESTROY CHARTS
========================================================= */

function destroyCharts() {

    Object.keys(charts).forEach(
        key => {

            if (
                charts[key] &&
                typeof charts[key].destroy ===
                "function"
            ) {

                try {
                    charts[key].destroy();
                } catch (error) {
                    console.warn(
                        `Could not destroy ${key}`,
                        error
                    );
                }

            }


            charts[key] = null;

        }
    );

}


/* =========================================================
   RESET DASHBOARD
========================================================= */

function resetDashboard() {

    destroyCharts();

    workbook = null;
    currentFile = null;


    if (dgrFile) {
        dgrFile.value = "";
    }


    /*
       Reset file UI
    */

    hide("fileInfo");
    hide("workbookStatus");
    show("emptyState");


    /*
       Reset analytics
    */

    hideAnalyticsSections();


    /*
       Reset texts
    */

    setText(
        "fileName",
        "—"
    );

    setText(
        "fileSheets",
        "—"
    );

    setText(
        "sidebarFileName",
        "No DGR uploaded"
    );

    setText(
        "statusText",
        "Upload a DGR to generate the analytics."
    );


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


    setText(
        "curtailmentSummary",
        "Waiting for DGR data"
    );


    if (sheetBadges) {
        sheetBadges.innerHTML = "";
    }

}


/* =========================================================
   GLOBAL ERROR HANDLING
========================================================= */

window.addEventListener(
    "error",
    (event) => {

        console.error(
            "Application error:",
            event.error || event.message
        );

    }
);


window.addEventListener(
    "unhandledrejection",
    (event) => {

        console.error(
            "Unhandled promise rejection:",
            event.reason
        );

    }
);
