/* =========================================================
   SOLAR DGR ANALYTICS
   app.js
   Robust DGR workbook parser + Chart.js dashboard
========================================================= */

"use strict";

/* =========================================================
   GLOBAL STATE
========================================================= */

let workbook = null;

let dailyRows = [];
let paRows = [];
let curtailmentRows = [];
let annualRows = [];

const charts = {};


/* =========================================================
   DOM HELPER
========================================================= */

const $ = id => document.getElementById(id);


/* =========================================================
   FIELD DEFINITIONS
========================================================= */

const DAILY_FIELDS = {
    date: "Date",
    pa: "PA(%)",
    pr: "PR(%)",
    hours: "Operating Hours",
    loss: "System_Losses (%)",
    measured: "Inv_Exp (kWh)",
    netExport: "220kV_Net_Exp (KWh)",
    capacity: "Firm DC Capacity (MWp)"
};

const CURTAILMENT_FIELDS = {
    date: "Date",
    from: "From",
    to: "To",
    duration: "Duration in MIN",
    lossGeneration: "Loss of Generation MWh",
    capacityLoss: "Loss of Capacity in MWH",
    remark: "Remark",
    setPoint: "Set Point (AC MW)"
};

const PA_FIELDS = {
    date: "Date",
    start: "Breakdown Time",
    acknowledgement: "Acknowledgement Time",
    workStart: "Work Start time on Fault",
    completion: "Work Completion time on fault",
    duration: "Breakdown Time",
    equipment: "Affected Equipment",
    category: "Fault Category",
    description: "Breakdown Description",
    energyLoss: "Approximate Energy Loss (KWh)"
};


/* =========================================================
   NUMBER HELPERS
========================================================= */

function cleanNumber(value) {

    if (
        value === null ||
        value === undefined ||
        value === ""
    ) {
        return NaN;
    }

    if (typeof value === "number") {
        return Number.isFinite(value) ? value : NaN;
    }

    let text = String(value)
        .replace(/,/g, "")
        .replace(/%/g, "")
        .trim();

    if (!text) return NaN;

    const n = Number(text);

    return Number.isFinite(n) ? n : NaN;
}


/* =========================================================
   DATE HELPERS
========================================================= */

function toDate(value) {

    if (value instanceof Date && !isNaN(value)) {
        return new Date(
            value.getFullYear(),
            value.getMonth(),
            value.getDate()
        );
    }

    if (typeof value === "number") {

        try {

            const parsed =
                XLSX.SSF.parse_date_code(value);

            if (parsed) {

                return new Date(
                    parsed.y,
                    parsed.m - 1,
                    parsed.d
                );
            }

        } catch (error) {

            console.warn(
                "Excel date conversion failed:",
                value
            );
        }
    }

    if (typeof value === "string") {

        const trimmed = value.trim();

        if (!trimmed) return null;

        const parsed = new Date(trimmed);

        if (!isNaN(parsed)) {

            return new Date(
                parsed.getFullYear(),
                parsed.getMonth(),
                parsed.getDate()
            );
        }
    }

    return null;
}


function isoDate(date) {

    if (!(date instanceof Date) || isNaN(date)) {
        return "";
    }

    const y = date.getFullYear();

    const m = String(
        date.getMonth() + 1
    ).padStart(2, "0");

    const d = String(
        date.getDate()
    ).padStart(2, "0");

    return `${y}-${m}-${d}`;
}


function shortDate(date) {

    if (!(date instanceof Date)) {
        return "";
    }

    return date.toLocaleDateString(
        "en-IN",
        {
            day: "2-digit",
            month: "short"
        }
    );
}


function dayNumber(date) {

    return date instanceof Date
        ? date.getDate()
        : "";
}


function fullDate(date) {

    if (!(date instanceof Date)) {
        return "";
    }

    return date.toLocaleDateString(
        "en-IN",
        {
            day: "2-digit",
            month: "short",
            year: "numeric"
        }
    );
}


/* =========================================================
   TIME HELPERS
========================================================= */

function timeToMinutes(value) {

    if (value === null || value === undefined) {
        return NaN;
    }

    if (value instanceof Date) {

        return (
            value.getHours() * 60 +
            value.getMinutes()
        );
    }

    const text = String(value).trim();

    const match = text.match(
        /^(\d{1,2}):(\d{2})/
    );

    if (!match) return NaN;

    return (
        Number(match[1]) * 60 +
        Number(match[2])
    );
}


function formatTime(value) {

    if (
        value === null ||
        value === undefined ||
        value === ""
    ) {
        return "—";
    }

    if (value instanceof Date) {

        return value.toLocaleTimeString(
            "en-IN",
            {
                hour: "2-digit",
                minute: "2-digit",
                hour12: false
            }
        );
    }

    const text = String(value);

    const match = text.match(
        /^(\d{1,2}):(\d{2})/
    );

    if (!match) return text;

    return `${String(match[1]).padStart(2, "0")}:${match[2]}`;
}


/* =========================================================
   DISPLAY HELPERS
========================================================= */

function formatNumber(value, decimals = 2) {

    if (!Number.isFinite(value)) {
        return "—";
    }

    return value.toLocaleString(
        "en-IN",
        {
            minimumFractionDigits: decimals,
            maximumFractionDigits: decimals
        }
    );
}


function formatPercent(value) {

    if (!Number.isFinite(value)) {
        return "—";
    }

    /*
       DGR stores PA / PR / System Loss as decimals.

       Example:
       0.786237 = 78.62%
       0.019305 = 1.93%
    */

    let percentage = value;

    if (Math.abs(value) <= 1.5) {
        percentage = value * 100;
    }

    return `${percentage.toFixed(2)}%`;
}


/* =========================================================
   ARRAY / SHEET HELPERS
========================================================= */

function sheetToMatrix(sheet) {

    return XLSX.utils.sheet_to_json(
        sheet,
        {
            header: 1,
            raw: true,
            defval: null
        }
    );
}


function findHeaderRow(sheet, requiredHeaders) {

    const matrix = sheetToMatrix(sheet);

    for (
        let i = 0;
        i < Math.min(matrix.length, 30);
        i++
    ) {

        const row = matrix[i] || [];

        const values = row.map(
            value =>
                String(value ?? "")
                    .trim()
                    .toLowerCase()
        );

        const found = requiredHeaders.every(
            required =>
                values.includes(
                    String(required)
                        .trim()
                        .toLowerCase()
                )
        );

        if (found) {
            return i;
        }
    }

    return -1;
}


function sheetToObjects(sheet, requiredHeaders = []) {

    const matrix = sheetToMatrix(sheet);

    if (!matrix.length) {
        return [];
    }

    let headerRow = 0;

    if (requiredHeaders.length) {

        const detected =
            findHeaderRow(
                sheet,
                requiredHeaders
            );

        if (detected >= 0) {
            headerRow = detected;
        }
    }

    const headerValues =
        matrix[headerRow] || [];

    const headers =
        headerValues.map(
            value =>
                String(value ?? "").trim()
        );

    const output = [];

    for (
        let i = headerRow + 1;
        i < matrix.length;
        i++
    ) {

        const row = matrix[i];

        if (!row) continue;

        const object = {};

        headers.forEach(
            (header, index) => {

                if (header) {
                    object[header] =
                        row[index];
                }
            }
        );

        const hasValue =
            Object.values(object).some(
                value =>
                    value !== null &&
                    value !== undefined &&
                    value !== ""
            );

        if (hasValue) {
            output.push(object);
        }
    }

    return output;
}


/* =========================================================
   DAILY KPI PARSER
========================================================= */

function parseDailyKPI(wb) {

    const sheet =
        wb.Sheets["Daily_KPI"];

    if (!sheet) {
        throw new Error(
            "Daily_KPI worksheet was not found."
        );
    }

    const data =
        sheetToObjects(
            sheet,
            [
                "Date",
                "PA(%)",
                "PR(%)"
            ]
        );

    const result = [];

    data.forEach(row => {

        const date =
            toDate(
                row[DAILY_FIELDS.date]
            );

        if (!date) return;

        result.push({

            date,

            pa:
                cleanNumber(
                    row[DAILY_FIELDS.pa]
                ),

            pr:
                cleanNumber(
                    row[DAILY_FIELDS.pr]
                ),

            hours:
                cleanNumber(
                    row[DAILY_FIELDS.hours]
                ),

            loss:
                cleanNumber(
                    row[DAILY_FIELDS.loss]
                ),

            measured:
                cleanNumber(
                    row[DAILY_FIELDS.measured]
                ),

            netExport:
                cleanNumber(
                    row[DAILY_FIELDS.netExport]
                ),

            capacity:
                cleanNumber(
                    row[DAILY_FIELDS.capacity]
                )

        });
    });

    return result
        .filter(row => row.date)
        .sort(
            (a, b) =>
                a.date - b.date
        );
}


/* =========================================================
   PA PARSER
========================================================= */

function parsePA(wb) {

    const sheet =
        wb.Sheets["PA"];

    if (!sheet) {
        console.warn(
            "PA worksheet not found."
        );

        return [];
    }

    const data =
        sheetToObjects(
            sheet,
            [
                "Date",
                "Breakdown Description"
            ]
        );

    return data
        .map(row => {

            const date =
                toDate(
                    row[PA_FIELDS.date]
                );

            if (!date) return null;

            return {

                date,

                start:
                    row[PA_FIELDS.start],

                acknowledgement:
                    row[PA_FIELDS.acknowledgement],

                workStart:
                    row[PA_FIELDS.workStart],

                completion:
                    row[PA_FIELDS.completion],

                duration:
                    cleanNumber(
                        row[PA_FIELDS.duration]
                    ),

                equipment:
                    row[PA_FIELDS.equipment] ||
                    "Equipment",

                category:
                    row[PA_FIELDS.category] ||
                    "Breakdown",

                description:
                    row[PA_FIELDS.description] ||
                    "",

                energyLoss:
                    cleanNumber(
                        row[PA_FIELDS.energyLoss]
                    )

            };

        })
        .filter(Boolean);
}


/* =========================================================
   CURTAILMENT PARSER
========================================================= */

function parseCurtailment(wb) {

    const sheet =
        wb.Sheets["Curtailment records"];

    if (!sheet) {

        console.warn(
            "Curtailment records worksheet not found."
        );

        return [];
    }

    const data =
        sheetToObjects(
            sheet,
            [
                "Date",
                "Loss of Generation MWh"
            ]
        );

    return data
        .map(row => {

            const date =
                toDate(
                    row[CURTAILMENT_FIELDS.date]
                );

            if (!date) return null;

            return {

                date,

                from:
                    row[CURTAILMENT_FIELDS.from],

                to:
                    row[CURTAILMENT_FIELDS.to],

                duration:
                    row[CURTAILMENT_FIELDS.duration],

                lossGeneration:
                    cleanNumber(
                        row[
                            CURTAILMENT_FIELDS
                                .lossGeneration
                        ]
                    ),

                capacityLoss:
                    cleanNumber(
                        row[
                            CURTAILMENT_FIELDS
                                .capacityLoss
                        ]
                    ),

                remark:
                    row[
                        CURTAILMENT_FIELDS.remark
                    ] || "",

                setPoint:
                    cleanNumber(
                        row[
                            CURTAILMENT_FIELDS
                                .setPoint
                        ]
                    )

            };

        })
        .filter(Boolean)
        .sort(
            (a, b) =>
                a.date - b.date
        );
}


/* =========================================================
   ANNUAL KPI PARSER
========================================================= */

function parseAnnualKPI(wb) {

    const sheet =
        wb.Sheets["Annual_KPI"];

    if (!sheet) {

        console.warn(
            "Annual_KPI worksheet not found."
        );

        return [];
    }

    const matrix =
        sheetToMatrix(sheet);

    let headerRow = -1;

    for (
        let i = 0;
        i < matrix.length;
        i++
    ) {

        const values =
            (matrix[i] || [])
                .map(
                    value =>
                        String(
                            value ?? ""
                        )
                            .trim()
                            .toLowerCase()
                );

        if (
            values.includes(
                "budgeted energy (gwh)"
            ) &&
            values.includes(
                "measured energy (mwh) (net exp)"
            )
        ) {

            headerRow = i;
            break;
        }
    }

    if (headerRow < 0) {

        console.warn(
            "Monthwise Annual_KPI table not found."
        );

        return [];
    }

    const headers =
        matrix[headerRow].map(
            value =>
                String(value ?? "").trim()
        );

    const rows = [];

    for (
        let i = headerRow + 1;
        i < matrix.length;
        i++
    ) {

        const row = matrix[i];

        if (!row) continue;

        const obj = {};

        headers.forEach(
            (header, index) => {

                if (header) {
                    obj[header] =
                        row[index];
                }
            }
        );

        const date =
            toDate(
                obj["Year"]
            );

        if (!date) continue;

        const budget =
            cleanNumber(
                obj[
                    "Budgeted Energy (GWh)"
                ]
            );

        const measured =
            cleanNumber(
                obj[
                    "Measured Energy (MWh) (Net Exp)"
                ]
            );

        /*
           Budget is stored in GWh.
           Convert to MWh so both values use
           the same unit.
        */

        rows.push({

            date,

            budgetMWh:
                Number.isFinite(budget)
                    ? budget * 1000
                    : NaN,

            measuredMWh:
                measured

        });
    }

    return rows;
}


/* =========================================================
   CHART HELPERS
========================================================= */

function destroyChart(id) {

    if (charts[id]) {

        charts[id].destroy();

        delete charts[id];
    }
}


function baseOptions() {

    return {

        responsive: true,

        maintainAspectRatio: false,

        animation: false,

        interaction: {
            mode: "index",
            intersect: false
        },

        plugins: {

            legend: {
                display: true,
                position: "top"
            },

            tooltip: {

                callbacks: {

                    title(items) {

                        if (!items.length) {
                            return "";
                        }

                        return items[0].label;
                    }

                }

            }

        },

        scales: {

            x: {

                grid: {
                    display: false
                },

                ticks: {

                    autoSkip: false,

                    maxRotation: 0,

                    minRotation: 0

                }

            },

            y: {

                beginAtZero: false,

                ticks: {
                    maxTicksLimit: 8
                }

            }

        }

    };
}


/* =========================================================
   DAILY DATE AXIS
   IMPORTANT:
   Shows proper date intervals instead of:
   29, 31, 31, 29
========================================================= */

function dateAxisOptions(rows) {

    const dates =
        rows.map(
            row => row.date
        );

    const uniqueMonths =
        [
            ...new Set(
                dates.map(
                    date =>
                        `${date.getFullYear()}-${date.getMonth()}`
                )
            )
        ];

    return {

        grid: {
            display: false
        },

        ticks: {

            autoSkip: false,

            maxRotation: 0,

            minRotation: 0,

            callback(value, index) {

                const date =
                    dates[index];

                if (!date) {
                    return "";
                }

                /*
                   If the selected data spans
                   more than one month, show
                   month + date.

                   If it is one month, show
                   even-numbered dates:
                   2, 4, 6, 8, 10...
                */

                if (uniqueMonths.length > 1) {

                    return `${date.getDate()} ${date.toLocaleDateString(
                        "en-IN",
                        { month: "short" }
                    )}`;
                }

                const day =
                    date.getDate();

                return day % 2 === 0
                    ? day
                    : "";
            }

        }

    };
}


/* =========================================================
   CREATE LINE CHART
========================================================= */

function createLineChart(
    id,
    labels,
    datasets,
    yOptions = {}
) {

    if (!$(id)) return;

    destroyChart(id);

    const options =
        baseOptions();

    options.scales.y =
        {
            ...options.scales.y,
            ...yOptions
        };

    charts[id] =
        new Chart(
            $(id),
            {
                type: "line",

                data: {
                    labels,
                    datasets
                },

                options
            }
        );
}


/* =========================================================
   DAILY CHART DATASET
========================================================= */

function lineDataset(
    label,
    data
) {

    return {

        label,

        data,

        borderWidth: 2,

        pointRadius: 2,

        pointHoverRadius: 4,

        tension: 0.25,

        fill: false

    };
}


/* =========================================================
   FILTERED DAILY DATA
========================================================= */

function getFilteredDailyRows() {

    if (!dailyRows.length) {
        return [];
    }

    const fromValue =
        $("fromDate")?.value;

    const toValue =
        $("toDate")?.value;

    if (fromValue || toValue) {

        const from =
            fromValue
                ? new Date(
                    `${fromValue}T00:00:00`
                )
                : null;

        const to =
            toValue
                ? new Date(
                    `${toValue}T23:59:59`
                )
                : null;

        return dailyRows.filter(
            row =>
                (!from || row.date >= from) &&
                (!to || row.date <= to)
        );
    }

    return dailyRows;
}


/* =========================================================
   DAILY CHARTS
========================================================= */

function renderDailyCharts(rows) {

    if (!rows.length) return;

    const labels =
        rows.map(
            row =>
                shortDate(row.date)
        );

    /* ---------------- PA ---------------- */

    createLineChart(
        "paChart",
        labels,
        [
            lineDataset(
                "PA (%)",
                rows.map(
                    row =>
                        Number.isFinite(row.pa)
                            ? row.pa * 100
                            : NaN
                )
            )
        ],
        {
            min: 0,
            max: 100,
            ticks: {
                callback:
                    value =>
                        `${value}%`
            }
        }
    );


    /* ---------------- PR ---------------- */

    createLineChart(
        "prChart",
        labels,
        [
            lineDataset(
                "PR (%)",
                rows.map(
                    row =>
                        Number.isFinite(row.pr)
                            ? row.pr * 100
                            : NaN
                )
            )
        ],
        {
            min: 0,
            max: 100,
            ticks: {
                callback:
                    value =>
                        `${value}%`
            }
        }
    );


    /* ---------------- HOURS ---------------- */

    createLineChart(
        "hoursChart",
        labels,
        [
            lineDataset(
                "Operating Hours",
                rows.map(
                    row =>
                        row.hours
                )
            )
        ],
        {
            beginAtZero: true,
            ticks: {
                callback:
                    value =>
                        `${value} h`
            }
        }
    );


    /* ---------------- SYSTEM LOSS ---------------- */

    createLineChart(
        "lossChart",
        labels,
        [
            lineDataset(
                "System Loss (%)",
                rows.map(
                    row =>
                        Number.isFinite(row.loss)
                            ? row.loss * 100
                            : NaN
                )
            )
        ],
        {
            beginAtZero: true,
            ticks: {
                callback:
                    value =>
                        `${value}%`
            }
        }
    );


    /* ---------------- DASHBOARD PR ---------------- */

    createLineChart(
        "dashboardPRChart",
        labels,
        [
            lineDataset(
                "PR (%)",
                rows.map(
                    row =>
                        Number.isFinite(row.pr)
                            ? row.pr * 100
                            : NaN
                )
            )
        ],
        {
            min: 0,
            max: 100,
            ticks: {
                callback:
                    value =>
                        `${value}%`
            }
        }
    );


    /* ---------------- DASHBOARD LOSS ---------------- */

    createLineChart(
        "dashboardLossChart",
        labels,
        [
            lineDataset(
                "System Loss (%)",
                rows.map(
                    row =>
                        Number.isFinite(row.loss)
                            ? row.loss * 100
                            : NaN
                )
            )
        ],
        {
            beginAtZero: true,
            ticks: {
                callback:
                    value =>
                        `${value}%`
            }
        }
    );
}


/* =========================================================
   PA BREAKDOWN TIMELINE
========================================================= */

function renderPAChart() {

    if (!paRows.length || !$("paChart")) {
        return;
    }

    /*
       The PA worksheet is event-based.

       We display the breakdown records as a
       duration chart using horizontal bars.
    */

    const labels =
        paRows.map(
            row =>
                `${fullDate(row.date)} — ${row.equipment}`
        );

    const values =
        paRows.map(
            row =>
                Number.isFinite(row.duration)
                    ? row.duration
                    : 0
        );

    destroyChart("paChart");

    charts.paChart =
        new Chart(
            $("paChart"),
            {
                type: "bar",

                data: {

                    labels,

                    datasets: [
                        {
                            label:
                                "Breakdown Duration (hrs)",

                            data: values,

                            borderWidth: 1
                        }
                    ]
                },

                options: {

                    indexAxis: "y",

                    responsive: true,

                    maintainAspectRatio: false,

                    animation: false,

                    plugins: {

                        legend: {
                            display: true
                        },

                        tooltip: {

                            callbacks: {

                                afterLabel(context) {

                                    const row =
                                        paRows[
                                            context.dataIndex
                                        ];

                                    return [
                                        `Start: ${formatTime(row.start)}`,
                                        `Completion: ${formatTime(row.completion)}`,
                                        `Equipment: ${row.equipment}`,
                                        `Category: ${row.category}`
                                    ];
                                }

                            }

                        }

                    },

                    scales: {

                        x: {

                            beginAtZero: true,

                            title: {
                                display: true,
                                text:
                                    "Breakdown Duration (Hours)"
                            }

                        },

                        y: {

                            ticks: {
                                autoSkip: false
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

function renderCurtailmentChart() {

    if (
        !curtailmentRows.length ||
        !$("curtailmentChart")
    ) {
        return;
    }

    const labels =
        curtailmentRows.map(
            row =>
                `${shortDate(row.date)} ${formatTime(row.from)}`
        );

    const values =
        curtailmentRows.map(
            row =>
                Number.isFinite(
                    row.lossGeneration
                )
                    ? row.lossGeneration
                    : 0
        );

    createLineChart(
        "curtailmentChart",
        labels,
        [
            lineDataset(
                "Curtailment Loss (MWh)",
                values
            )
        ],
        {
            beginAtZero: true,
            ticks: {
                callback:
                    value =>
                        `${value} MWh`
            }
        }
    );


    const total =
        values.reduce(
            (sum, value) =>
                sum +
                (
                    Number.isFinite(value)
                        ? value
                        : 0
                ),
            0
        );

    const records =
        curtailmentRows.length;

    if ($("curtailmentSummary")) {

        $("curtailmentSummary").textContent =
            `${records.toLocaleString(
                "en-IN"
            )} curtailment records • ` +
            `${formatNumber(
                total
            )} MWh estimated generation loss`;
    }
}


/* =========================================================
   ENERGY CHART
========================================================= */

function renderEnergyChart() {

    if (
        !annualRows.length ||
        !$("energyChart")
    ) {
        return;
    }

    const labels =
        annualRows.map(
            row =>
                row.date.toLocaleDateString(
                    "en-IN",
                    {
                        month: "short"
                    }
                )
        );

    const budget =
        annualRows.map(
            row =>
                row.budgetMWh
        );

    const measured =
        annualRows.map(
            row =>
                row.measuredMWh
        );

    destroyChart("energyChart");

    charts.energyChart =
        new Chart(
            $("energyChart"),
            {

                type: "bar",

                data: {

                    labels,

                    datasets: [

                        {
                            label:
                                "Budgeted Energy (MWh)",

                            data: budget,

                            borderWidth: 1
                        },

                        {
                            label:
                                "Measured Energy (MWh)",

                            data: measured,

                            borderWidth: 1
                        }

                    ]

                },

                options: {

                    responsive: true,

                    maintainAspectRatio: false,

                    animation: false,

                    plugins: {

                        legend: {
                            display: true,
                            position: "top"
                        }

                    },

                    scales: {

                        x: {
                            grid: {
                                display: false
                            }
                        },

                        y: {

                            beginAtZero: true,

                            ticks: {

                                callback:
                                    value =>
                                        `${Number(
                                            value
                                        ).toLocaleString(
                                            "en-IN"
                                        )} MWh`
                            }

                        }

                    }

                }

            }
        );
}


/* =========================================================
   KPI CARDS
========================================================= */

function renderKPIs(rows) {

    if (!rows.length) return;

    const latest =
        rows[rows.length - 1];


    /* PA */

    if ($("dashboardPA")) {

        $("dashboardPA").textContent =
            formatPercent(
                latest.pa
            );
    }


    /* PR */

    if ($("dashboardPR")) {

        $("dashboardPR").textContent =
            formatPercent(
                latest.pr
            );
    }


    /* SYSTEM LOSS */

    if ($("dashboardLoss")) {

        $("dashboardLoss").textContent =
            formatPercent(
                latest.loss
            );
    }


    /* OPERATING HOURS */

    if ($("dashboardHours")) {

        $("dashboardHours").textContent =
            Number.isFinite(
                latest.hours
            )
                ? `${latest.hours.toFixed(2)} h`
                : "—";
    }
}


/* =========================================================
   ENERGY SUMMARY
========================================================= */

function renderEnergySummary() {

    if (!annualRows.length) return;

    const totalBudget =
        annualRows.reduce(
            (sum, row) =>
                sum +
                (
                    Number.isFinite(
                        row.budgetMWh
                    )
                        ? row.budgetMWh
                        : 0
                ),
            0
        );

    const totalMeasured =
        annualRows.reduce(
            (sum, row) =>
                sum +
                (
                    Number.isFinite(
                        row.measuredMWh
                    )
                        ? row.measuredMWh
                        : 0
                ),
            0
        );

    const variance =
        totalMeasured -
        totalBudget;


    if ($("totalBudget")) {

        $("totalBudget").textContent =
            `${formatNumber(
                totalBudget
            )} MWh`;
    }


    if ($("totalMeasured")) {

        $("totalMeasured").textContent =
            `${formatNumber(
                totalMeasured
            )} MWh`;
    }


    if ($("energyVariance")) {

        $("energyVariance").textContent =
            `${formatNumber(
                variance
            )} MWh`;
    }
}


/* =========================================================
   WORKBOOK STATUS
========================================================= */

function renderWorkbookStatus() {

    if (!$("sheetBadges")) {
        return;
    }

    const requiredSheets = [
        "Dashboard",
        "Annual_KPI",
        "Daily_KPI",
        "PA",
        "Curtailment records"
    ];

    $("sheetBadges").innerHTML = "";

    requiredSheets.forEach(
        sheetName => {

            const badge =
                document.createElement(
                    "span"
                );

            badge.className =
                "sheet-badge";

            if (
                workbook &&
                workbook.SheetNames.includes(
                    sheetName
                )
            ) {

                badge.textContent =
                    `✓ ${sheetName}`;

            } else {

                badge.classList.add(
                    "missing"
                );

                badge.textContent =
                    `✕ ${sheetName}`;
            }

            $("sheetBadges")
                .appendChild(badge);
        }
    );
}


/* =========================================================
   FILE UI
========================================================= */

function showUploadedFile(file) {

    if ($("fileName")) {

        $("fileName").textContent =
            file.name;
    }

    if ($("fileSheets")) {

        $("fileSheets").textContent =
            `${workbook.SheetNames.length} worksheets detected`;
    }

    if ($("sidebarFileName")) {

        $("sidebarFileName").textContent =
            file.name;
    }

    $("fileInfo")?.classList.remove(
        "hidden"
    );

    $("workbookStatus")?.classList.remove(
        "hidden"
    );

    $("dropZone")?.classList.add(
        "hidden"
    );

    $("emptyState")?.classList.add(
        "hidden"
    );
}


/* =========================================================
   DATE FILTER SETUP
========================================================= */

function setupDateFilters() {

    if (!dailyRows.length) return;

    const first =
        dailyRows[0].date;

    const last =
        dailyRows[
            dailyRows.length - 1
        ].date;

    const min =
        isoDate(first);

    const max =
        isoDate(last);


    if ($("fromDate")) {

        $("fromDate").min = min;
        $("fromDate").max = max;
    }

    if ($("toDate")) {

        $("toDate").min = min;
        $("toDate").max = max;
    }
}


/* =========================================================
   FULL RENDER
========================================================= */

function renderAll() {

    if (!dailyRows.length) {
        return;
    }

    const filtered =
        getFilteredDailyRows();

    if (!filtered.length) {
        return;
    }


    renderKPIs(
        filtered
    );


    /*
       PA chart uses PA worksheet events.
       Daily PA trend uses Daily_KPI.
    */

    renderPAChart();


    /*
       The PA worksheet chart is a separate
       event timeline, so create daily charts
       after it.
    */

    renderDailyCharts(
        filtered
    );

    renderCurtailmentChart();

    renderEnergySummary();

    renderEnergyChart();


    if ($("statusText")) {

        const latest =
            filtered[
                filtered.length - 1
            ];

        $("statusText").textContent =
            `${filtered.length} DGR days analysed • ` +
            `Latest data: ${fullDate(
                latest.date
            )}`;
    }
}


/* =========================================================
   FILE PROCESSING
========================================================= */

async function handleFile(file) {

    if (!file) return;

    try {

        if (
            typeof XLSX === "undefined"
        ) {

            throw new Error(
                "SheetJS library is not loaded."
            );
        }


        const buffer =
            await file.arrayBuffer();


        workbook =
            XLSX.read(
                buffer,
                {
                    type: "array",
                    cellDates: true
                }
            );


        /* ---------------- DAILY KPI ---------------- */

        dailyRows =
            parseDailyKPI(
                workbook
            );


        if (!dailyRows.length) {

            throw new Error(
                "No usable daily records were found in the Daily_KPI worksheet."
            );
        }


        /* ---------------- PA ---------------- */

        paRows =
            parsePA(
                workbook
            );


        /* ---------------- CURTAILMENT ---------------- */

        curtailmentRows =
            parseCurtailment(
                workbook
            );


        /* ---------------- ANNUAL KPI ---------------- */

        annualRows =
            parseAnnualKPI(
                workbook
            );


        /* ---------------- UI ---------------- */

        showUploadedFile(
            file
        );

        renderWorkbookStatus();

        setupDateFilters();

        renderAll();


        /*
           Mapping information
        */

        if ($("mappingInfo")) {

            $("mappingInfo").innerHTML =

                `<strong>Daily KPI:</strong> PA, PR, Operating Hours and System Losses<br>
                 <strong>PA:</strong> Breakdown events and durations<br>
                 <strong>Curtailment:</strong> Loss of Generation MWh<br>
                 <strong>Annual KPI:</strong> Monthly Budgeted vs Measured Energy`;
        }


        console.log(
            "DGR successfully loaded",
            {
                dailyRows,
                paRows,
                curtailmentRows,
                annualRows
            }
        );


    } catch (error) {

        console.error(
            "DGR upload error:",
            error
        );

        alert(
            "DGR upload failed:\n\n" +
            error.message
        );
    }
}


/* =========================================================
   RANGE BUTTONS
========================================================= */

function setRange(range) {

    document
        .querySelectorAll(
            ".range-buttons button"
        )
        .forEach(
            button => {

                button.classList.toggle(
                    "selected",
                    button.dataset.range ===
                    String(range)
                );
            }
        );


    if ($("fromDate")) {
        $("fromDate").value = "";
    }

    if ($("toDate")) {
        $("toDate").value = "";
    }


    if (!dailyRows.length) return;


    let filtered;

    if (range === "all") {

        filtered =
            dailyRows;

    } else {

        const count =
            Number(range);

        filtered =
            dailyRows.slice(
                -count
            );
    }


    renderKPIs(
        filtered
    );

    renderDailyCharts(
        filtered
    );
}


/* =========================================================
   DATE FILTER BUTTONS
========================================================= */

function applyDateFilter() {

    document
        .querySelectorAll(
            ".range-buttons button"
        )
        .forEach(
            button =>
                button.classList.remove(
                    "selected"
                )
        );

    renderAll();
}


function resetDateFilter() {

    if ($("fromDate")) {
        $("fromDate").value = "";
    }

    if ($("toDate")) {
        $("toDate").value = "";
    }

    setRange("all");
}


/* =========================================================
   REMOVE FILE
========================================================= */

function removeFile() {

    workbook = null;

    dailyRows = [];

    paRows = [];

    curtailmentRows = [];

    annualRows = [];


    Object.keys(charts)
        .forEach(
            id =>
                destroyChart(id)
        );


    $("fileInfo")?.classList.add(
        "hidden"
    );

    $("workbookStatus")?.classList.add(
        "hidden"
    );

    $("dropZone")?.classList.remove(
        "hidden"
    );

    $("emptyState")?.classList.remove(
        "hidden"
    );


    if ($("sidebarFileName")) {

        $("sidebarFileName").textContent =
            "No DGR uploaded";
    }


    if ($("statusText")) {

        $("statusText").textContent =
            "Upload a DGR to generate the analytics.";
    }


    [
        "dashboardPA",
        "dashboardPR",
        "dashboardLoss",
        "dashboardHours",
        "totalBudget",
        "totalMeasured",
        "energyVariance"
    ]
        .forEach(
            id => {

                if ($(id)) {
                    $(id).textContent =
                        "—";
                }
            }
        );


    const input =
        $("dgrFile");

    if (input) {
        input.value = "";
    }
}


/* =========================================================
   SIDEBAR NAVIGATION
========================================================= */

function setupNavigation() {

    document
        .querySelectorAll(
            ".nav-item"
        )
        .forEach(
            button => {

                button.addEventListener(
                    "click",
                    () => {

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

                        button.classList.add(
                            "active"
                        );


                        const target =
                            $(button.dataset.target);

                        if (target) {

                            target.scrollIntoView(
                                {
                                    behavior: "smooth",
                                    block: "start"
                                }
                            );
                        }

                    }
                );

            }
        );
}


/* =========================================================
   UPLOAD / DRAG & DROP
========================================================= */

function setupUpload() {

    const input =
        $("dgrFile");

    const dropZone =
        $("dropZone");


    if (input) {

        input.addEventListener(
            "change",
            event => {

                const file =
                    event.target.files?.[0];

                handleFile(file);
            }
        );
    }


    if (!dropZone) return;


    dropZone.addEventListener(
        "click",
        () => {

            if (input) {
                input.click();
            }
        }
    );


    [
        "dragenter",
        "dragover"
    ]
        .forEach(
            eventName => {

                dropZone.addEventListener(
                    eventName,
                    event => {

                        event.preventDefault();

                        dropZone.classList.add(
                            "dragging"
                        );
                    }
                );
            }
        );


    [
        "dragleave",
        "drop"
    ]
        .forEach(
            eventName => {

                dropZone.addEventListener(
                    eventName,
                    event => {

                        event.preventDefault();

                        dropZone.classList.remove(
                            "dragging"
                        );
                    }
                );
            }
        );


    dropZone.addEventListener(
        "drop",
        event => {

            const file =
                event.dataTransfer
                    ?.files?.[0];

            handleFile(file);
        }
    );
}


/* =========================================================
   INITIALISE
========================================================= */

document.addEventListener(
    "DOMContentLoaded",
    () => {

        setupNavigation();

        setupUpload();


        /* Remove button */

        $("removeFile")
            ?.addEventListener(
                "click",
                removeFile
            );


        /* Date buttons */

        $("applyDate")
            ?.addEventListener(
                "click",
                applyDateFilter
            );

        $("resetDate")
            ?.addEventListener(
                "click",
                resetDateFilter
            );


        /* Range buttons */

        document
            .querySelectorAll(
                ".range-buttons button"
            )
            .forEach(
                button => {

                    button.addEventListener(
                        "click",
                        () =>
                            setRange(
                                button.dataset.range
                            )
                    );
                }
            );


        console.log(
            "Solar DGR Analytics initialised."
        );
    }
);
