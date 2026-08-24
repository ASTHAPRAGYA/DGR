/* =========================================================
   SRI1PL SOLAR DGR ANALYTICS
   COMPLETE DASHBOARD ENGINE
   ========================================================= */


/* =========================================================
   GLOBAL STATE
   ========================================================= */

let allData = [];
let filteredData = [];

const chartInstances = {};


/* =========================================================
   DOM HELPER
   ========================================================= */

function el(id) {
    return document.getElementById(id);
}


/* =========================================================
   SAFE NUMBER
   ========================================================= */

function num(value) {

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

    let text = String(value)
        .trim()
        .replace(/,/g, "")
        .replace(/%/g, "");

    if (!text) {
        return null;
    }

    const result = parseFloat(text);

    return Number.isFinite(result)
        ? result
        : null;
}


/* =========================================================
   DATE PARSER
   ========================================================= */

function parseDate(value) {

    if (value instanceof Date) {

        return isNaN(value.getTime())
            ? null
            : value;

    }

    if (typeof value === "number") {

        try {

            const d =
                XLSX.SSF.parse_date_code(value);

            if (!d) {
                return null;
            }

            return new Date(
                d.y,
                d.m - 1,
                d.d
            );

        } catch {

            return null;

        }
    }


    if (typeof value === "string") {

        let text =
            value.trim();


        /*
         * DD/MM/YYYY
         */

        const slash =
            text.match(
                /^(\d{1,2})[\/-](\d{1,2})[\/-](\d{4})/
            );


        if (slash) {

            const day =
                parseInt(slash[1]);

            const month =
                parseInt(slash[2]);

            const year =
                parseInt(slash[3]);


            const date =
                new Date(
                    year,
                    month - 1,
                    day
                );


            return isNaN(
                date.getTime()
            )
                ? null
                : date;

        }


        const parsed =
            new Date(text);


        return isNaN(
            parsed.getTime()
        )
            ? null
            : parsed;

    }


    return null;
}


/* =========================================================
   DATE FORMATTING
   ========================================================= */

function isoDate(date) {

    if (!date) {
        return "";
    }

    const year =
        date.getFullYear();

    const month =
        String(
            date.getMonth() + 1
        ).padStart(2, "0");

    const day =
        String(
            date.getDate()
        ).padStart(2, "0");

    return `${year}-${month}-${day}`;
}


function shortDate(date) {

    if (!date) {
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


/* =========================================================
   FIND COLUMN
   ========================================================= */

function findColumn(
    headers,
    candidates
) {

    const cleanHeaders =
        headers.map(
            h =>
                String(h)
                    .trim()
                    .toLowerCase()
        );


    /*
     * First look for exact matches.
     */

    for (
        const candidate
        of candidates
    ) {

        const index =
            cleanHeaders.indexOf(
                candidate.toLowerCase()
            );


        if (index !== -1) {

            return headers[index];

        }

    }


    /*
     * Then look for partial matches.
     */

    for (
        const candidate
        of candidates
    ) {

        const search =
            candidate
                .toLowerCase()
                .replace(/\s/g, "");


        const index =
            cleanHeaders.findIndex(
                header =>
                    header
                        .replace(/\s/g, "")
                        .includes(search)
            );


        if (index !== -1) {

            return headers[index];

        }

    }


    return null;
}


/* =========================================================
   DETECT COLUMNS
   ========================================================= */

function detectColumns(headers) {

    return {

        date:
            findColumn(
                headers,
                [
                    "Date",
                    "date"
                ]
            ),

        pa:
            findColumn(
                headers,
                [
                    "PA(%)",
                    "PA (%)",
                    "Plant Availability (%)",
                    "Plant Availability",
                    "PA"
                ]
            ),

        pr:
            findColumn(
                headers,
                [
                    "PR(%)",
                    "PR (%)",
                    "Performance Ratio (%)",
                    "Performance Ratio",
                    "PR"
                ]
            ),

        ghi:
            findColumn(
                headers,
                [
                    "GHI-UP (KWh/m2)",
                    "GHI-UP(KWh/m2)",
                    "GHI",
                    "GHI-UP"
                ]
            ),

        poaUp:
            findColumn(
                headers,
                [
                    "POA-UP(KWh/m2)",
                    "POA-UP (KWh/m2)",
                    "POA-UP",
                    "POA Up",
                    "GII",
                    "GII-UP"
                ]
            ),

        poaDown:
            findColumn(
                headers,
                [
                    "POA-Down(KWh/m2)",
                    "POA-Down (KWh/m2)",
                    "POA-Down",
                    "POA Down"
                ]
            ),

        measured:
            findColumn(
                headers,
                [
                    "Inv_Exp (kWh)",
                    "Inv_Exp(kWh)",
                    "Inv Exp",
                    "Inverter Export",
                    "Measured Generation",
                    "Measured Generation (kWh)"
                ]
            ),

        actual:
            findColumn(
                headers,
                [
                    "220kV_Net_Exp (KWh)",
                    "220kV_Net_Exp(KWh)",
                    "220kV Net Export",
                    "Actual Generation",
                    "Actual Generation (kWh)",
                    "Net Export"
                ]
            ),

        capacity:
            findColumn(
                headers,
                [
                    "Firm DC Capacity (MWp)",
                    "DC Capacity (MWp)",
                    "Installed DC Capacity",
                    "Capacity (MWp)"
                ]
            )

    };
}


/* =========================================================
   DETECT HEADER ROW
   ========================================================= */

function detectHeaderRow(sheet) {

    const matrix =
        XLSX.utils.sheet_to_json(
            sheet,
            {
                header: 1,
                defval: null,
                raw: true
            }
        );


    /*
     * Search first 30 rows.
     */

    for (
        let rowIndex = 0;
        rowIndex <
        Math.min(
            matrix.length,
            30
        );
        rowIndex++
    ) {

        const row =
            matrix[rowIndex] || [];


        const values =
            row.map(
                value =>
                    String(
                        value ?? ""
                    ).trim()
            );


        const lower =
            values.map(
                value =>
                    value.toLowerCase()
            );


        const hasDate =
            lower.some(
                value =>
                    value === "date"
            );


        const hasPA =
            lower.some(
                value =>
                    value.includes("pa")
            );


        const hasPR =
            lower.some(
                value =>
                    value.includes("pr")
            );


        if (
            hasDate &&
            (hasPA || hasPR)
        ) {

            return rowIndex;

        }

    }


    return -1;
}


/* =========================================================
   READ SHEET
   ========================================================= */

function readSheet(sheet) {

    const headerRow =
        detectHeaderRow(sheet);


    if (headerRow === -1) {

        return null;

    }


    const matrix =
        XLSX.utils.sheet_to_json(
            sheet,
            {
                header: 1,
                defval: null,
                raw: true
            }
        );


    const headers =
        matrix[
            headerRow
        ].map(
            value =>
                String(
                    value ?? ""
                ).trim()
        );


    const objects = [];


    for (
        let i = headerRow + 1;
        i < matrix.length;
        i++
    ) {

        const row =
            matrix[i];


        if (!row) {
            continue;
        }


        const object = {};


        headers.forEach(
            (
                header,
                index
            ) => {

                if (header) {

                    object[header] =
                        row[index];

                }

            }
        );


        objects.push(
            object
        );

    }


    return {

        headers,

        rows: objects,

        headerRow

    };
}


/* =========================================================
   FIND BEST WORKSHEET
   ========================================================= */

function findBestSheet(workbook) {

    let best = null;


    for (
        const sheetName
        of workbook.SheetNames
    ) {

        const sheet =
            workbook.Sheets[
                sheetName
            ];


        const result =
            readSheet(sheet);


        if (!result) {
            continue;
        }


        const columns =
            detectColumns(
                result.headers
            );


        let score = 0;


        if (columns.date) score += 5;
        if (columns.pa) score += 3;
        if (columns.pr) score += 3;
        if (columns.ghi) score += 2;
        if (columns.poaUp) score += 2;
        if (columns.measured) score += 2;
        if (columns.actual) score += 2;


        if (
            !best ||
            score > best.score
        ) {

            best = {

                sheetName,

                result,

                columns,

                score

            };

        }

    }


    return best;
}


/* =========================================================
   NORMALISE DATA
   ========================================================= */

function normalise(
    sheetData
) {

    const {
        rows,
        columns
    } = sheetData;


    const output = [];


    rows.forEach(
        row => {

            const date =
                parseDate(
                    row[
                        columns.date
                    ]
                );


            if (!date) {
                return;
            }


            let pa =
                num(
                    row[
                        columns.pa
                    ]
                );


            let pr =
                num(
                    row[
                        columns.pr
                    ]
                );


            /*
             * Some DGRs store PA/PR as decimals:
             * 0.98 instead of 98.
             *
             * Convert automatically.
             */

            if (
                pa !== null &&
                pa <= 1
            ) {

                pa *= 100;

            }


            if (
                pr !== null &&
                pr <= 1
            ) {

                pr *= 100;

            }


            output.push({

                date,

                pa,

                pr,

                ghi:
                    num(
                        row[
                            columns.ghi
                        ]
                    ),

                poaUp:
                    num(
                        row[
                            columns.poaUp
                        ]
                    ),

                poaDown:
                    num(
                        row[
                            columns.poaDown
                        ]
                    ),

                measured:
                    num(
                        row[
                            columns.measured
                        ]
                    ),

                actual:
                    num(
                        row[
                            columns.actual
                        ]
                    ),

                capacity:
                    num(
                        row[
                            columns.capacity
                        ]
                    )

            });

        }
    );


    output.sort(
        (
            a,
            b
        ) =>
            a.date - b.date
    );


    return output;
}


/* =========================================================
   FORMAT NUMBER
   ========================================================= */

function formatNumber(
    value,
    decimals = 2
) {

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
            minimumFractionDigits:
                decimals,

            maximumFractionDigits:
                decimals
        }
    );
}


/* =========================================================
   CHART COLORS
   ========================================================= */

const chartColors = {

    primary:
        "#27a5ad",

    dark:
        "#17252a",

    secondary:
        "#78c7cc",

    light:
        "#b9e2e4",

    grey:
        "#9aa8ab",

    orange:
        "#e6a23c"

};


/* =========================================================
   DESTROY CHART
   ========================================================= */

function destroyChart(id) {

    if (
        chartInstances[id]
    ) {

        chartInstances[id].destroy();

        delete chartInstances[id];

    }

}


/* =========================================================
   CREATE CHART
   ========================================================= */

function createChart(
    id,
    config
) {

    const canvas =
        el(id);


    if (!canvas) {
        return;
    }


    destroyChart(id);


    chartInstances[id] =
        new Chart(
            canvas,
            config
        );

}


/* =========================================================
   COMMON OPTIONS
   ========================================================= */

function commonOptions(
    yTitle = ""
) {

    return {

        responsive: true,

        maintainAspectRatio: false,

        interaction: {

            mode: "index",

            intersect: false

        },

        plugins: {

            legend: {

                position: "top",

                labels: {

                    usePointStyle: true,

                    padding: 15,

                    font: {

                        family: "Inter",

                        size: 10

                    }

                }

            },

            tooltip: {

                backgroundColor:
                    "#17252a",

                padding: 10,

                cornerRadius: 7

            }

        },

        scales: {

            x: {

                grid: {

                    display: false

                },

                ticks: {

                    maxTicksLimit: 12,

                    font: {

                        size: 9

                    }

                }

            },

            y: {

                grid: {

                    color:
                        "rgba(23,37,42,0.07)"

                },

                title: {

                    display:
                        Boolean(yTitle),

                    text:
                        yTitle,

                    font: {

                        size: 10

                    }

                },

                ticks: {

                    font: {

                        size: 9

                    }

                }

            }

        }

    };
}


/* =========================================================
   LINE CHART
   ========================================================= */

function makeLineChart(
    id,
    labels,
    datasets,
    yTitle
) {

    createChart(
        id,
        {

            type: "line",

            data: {

                labels,

                datasets

            },

            options:
                commonOptions(
                    yTitle
                )

        }
    );

}


/* =========================================================
   BAR CHART
   ========================================================= */

function makeBarChart(
    id,
    labels,
    datasets,
    yTitle
) {

    createChart(
        id,
        {

            type: "bar",

            data: {

                labels,

                datasets

            },

            options:
                commonOptions(
                    yTitle
                )

        }
    );

}


/* =========================================================
   PA CHARTS
   ========================================================= */

function drawPA(
    labels,
    data
) {

    const dataset = {

        label:
            "Plant Availability (%)",

        data,

        borderColor:
            chartColors.primary,

        backgroundColor:
            "rgba(39,165,173,0.10)",

        pointBackgroundColor:
            chartColors.primary,

        pointRadius: 3,

        pointHoverRadius: 6,

        borderWidth: 2,

        tension: 0.3

    };


    makeLineChart(
        "paChart",
        labels,
        [dataset],
        "Availability (%)"
    );


    makeBarChart(

        "paBarChart",

        labels,

        [

            {

                label:
                    "Plant Availability (%)",

                data,

                backgroundColor:
                    "rgba(39,165,173,0.65)",

                borderRadius: 4

            }

        ],

        "Availability (%)"

    );

}


/* =========================================================
   PR CHARTS
   ========================================================= */

function drawPR(
    labels,
    data
) {

    makeLineChart(

        "prChart",

        labels,

        [

            {

                label:
                    "Performance Ratio (%)",

                data,

                borderColor:
                    chartColors.dark,

                backgroundColor:
                    "rgba(23,37,42,0.08)",

                pointBackgroundColor:
                    chartColors.dark,

                pointRadius: 3,

                borderWidth: 2,

                tension: 0.3

            }

        ],

        "PR (%)"

    );


    makeBarChart(

        "prBarChart",

        labels,

        [

            {

                label:
                    "Performance Ratio (%)",

                data,

                backgroundColor:
                    "rgba(23,37,42,0.60)",

                borderRadius: 4

            }

        ],

        "PR (%)"

    );

}


/* =========================================================
   AEY CALCULATION
   ========================================================= */

function calculateAEY(
    data
) {

    let cumulative =
        0;


    let capacity =
        null;


    /*
       Find first valid capacity.
    */

    for (
        const row
        of data
    ) {

        if (
            row.capacity !== null &&
            row.capacity > 0
        ) {

            capacity =
                row.capacity;

            break;

        }

    }


    /*
       If capacity isn't available,
       use a default only for visualization.
    */

    if (
        !capacity ||
        capacity <= 0
    ) {

        return data.map(
            () => null
        );

    }


    const capacityKwp =
        capacity * 1000;


    return data.map(
        row => {

            if (
                row.actual !== null
            ) {

                cumulative +=
                    row.actual;

            }


            return (
                cumulative /
                capacityKwp
            );

        }
    );

}


/* =========================================================
   AEY CHART
   ========================================================= */

function drawAEY(
    labels,
    values
) {

    createChart(

        "aeyChart",

        {

            type: "line",

            data: {

                labels,

                datasets: [

                    {

                        label:
                            "Cumulative AEY (kWh/kWp)",

                        data: values,

                        borderColor:
                            chartColors.primary,

                        backgroundColor:
                            "rgba(39,165,173,0.14)",

                        fill: true,

                        pointRadius: 2,

                        borderWidth: 2,

                        tension: 0.25

                    }

                ]

            },

            options:
                commonOptions(
                    "kWh/kWp"
                )

        }

    );

}


/* =========================================================
   GHI CHARTS
   ========================================================= */

function drawGHI(
    labels,
    values
) {

    makeBarChart(

        "ghiChart",

        labels,

        [

            {

                label:
                    "GHI (kWh/m²)",

                data: values,

                backgroundColor:
                    "rgba(230,162,60,0.70)",

                borderRadius: 4

            }

        ],

        "GHI (kWh/m²)"

    );


    makeLineChart(

        "ghiLineChart",

        labels,

        [

            {

                label:
                    "GHI (kWh/m²)",

                data: values,

                borderColor:
                    chartColors.orange,

                backgroundColor:
                    "rgba(230,162,60,0.12)",

                pointRadius: 3,

                borderWidth: 2,

                tension: 0.3

            }

        ],

        "GHI (kWh/m²)"

    );

}


/* =========================================================
   POA / GII CHART
   ========================================================= */

function drawGII(
    labels,
    poaUp,
    poaDown
) {

    makeLineChart(

        "giiChart",

        labels,

        [

            {

                label:
                    "POA Up (kWh/m²)",

                data:
                    poaUp,

                borderColor:
                    chartColors.primary,

                backgroundColor:
                    "rgba(39,165,173,0.10)",

                pointRadius: 2,

                borderWidth: 2,

                tension: 0.3

            },

            {

                label:
                    "POA Down (kWh/m²)",

                data:
                    poaDown,

                borderColor:
                    chartColors.dark,

                backgroundColor:
                    "rgba(23,37,42,0.08)",

                pointRadius: 2,

                borderWidth: 2,

                tension: 0.3

            }

        ],

        "Irradiance (kWh/m²)"

    );

}


/* =========================================================
   GENERATION CHART
   ========================================================= */

function drawGeneration(
    labels,
    measured,
    actual
) {

    /*
       LINE
    */

    makeLineChart(

        "generationChart",

        labels,

        [

            {

                label:
                    "Measured / Inverter Export",

                data:
                    measured,

                borderColor:
                    chartColors.primary,

                backgroundColor:
                    "rgba(39,165,173,0.08)",

                pointRadius: 2,

                borderWidth: 2,

                tension: 0.25

            },

            {

                label:
                    "Actual / 220kV Net Export",

                data:
                    actual,

                borderColor:
                    chartColors.dark,

                backgroundColor:
                    "rgba(23,37,42,0.08)",

                pointRadius: 2,

                borderWidth: 2,

                tension: 0.25

            }

        ],

        "Generation (kWh)"

    );


    /*
       BAR
    */

    makeBarChart(

        "generationBarChart",

        labels,

        [

            {

                label:
                    "Measured",

                data:
                    measured,

                backgroundColor:
                    "rgba(39,165,173,0.65)",

                borderRadius: 3

            },

            {

                label:
                    "Actual",

                data:
                    actual,

                backgroundColor:
                    "rgba(23,37,42,0.60)",

                borderRadius: 3

            }

        ],

        "Generation (kWh)"

    );


    /*
       PIE / DOUGHNUT
       Based on total measured vs actual.
    */

    const totalMeasured =
        measured
            .filter(
                value =>
                    value !== null
            )
            .reduce(
                (
                    sum,
                    value
                ) =>
                    sum + value,
                0
            );


    const totalActual =
        actual
            .filter(
                value =>
                    value !== null
            )
            .reduce(
                (
                    sum,
                    value
                ) =>
                    sum + value,
                0
            );


    const difference =
        Math.max(
            totalMeasured -
            totalActual,
            0
        );


    createChart(

        "generationPieChart",

        {

            type: "doughnut",

            data: {

                labels: [

                    "Actual Generation",

                    "Generation Difference"

                ],

                datasets: [

                    {

                        data: [

                            totalActual,

                            difference

                        ],

                        backgroundColor: [

                            chartColors.primary,

                            chartColors.light

                        ],

                        borderWidth: 0

                    }

                ]

            },

            options: {

                responsive: true,

                maintainAspectRatio: false,

                cutout: "62%",

                plugins: {

                    legend: {

                        position: "bottom",

                        labels: {

                            usePointStyle: true,

                            padding: 15,

                            font: {

                                size: 10

                            }

                        }

                    },

                    tooltip: {

                        callbacks: {

                            label:
                                function(
                                    context
                                ) {

                                    return (

                                        " " +

                                        context.label +

                                        ": " +

                                        formatNumber(
                                            context.raw / 1000,
                                            2
                                        ) +

                                        " MWh"

                                    );

                                }

                        }

                    }

                }

            }

        }

    );

}


/* =========================================================
   SCATTER PLOT
   ========================================================= */

function drawScatter(
    data
) {

    const points =
        data
            .filter(
                row =>
                    row.ghi !== null &&
                    row.actual !== null
            )
            .map(
                row => ({

                    x: row.ghi,

                    y: row.actual / 1000,

                    date:
                        shortDate(
                            row.date
                        )

                })
            );


    createChart(

        "scatterChart",

        {

            type: "scatter",

            data: {

                datasets: [

                    {

                        label:
                            "GHI vs Actual Generation",

                        data: points,

                        backgroundColor:
                            chartColors.primary,

                        borderColor:
                            chartColors.primary,

                        pointRadius: 5,

                        pointHoverRadius: 7

                    }

                ]

            },

            options: {

                responsive: true,

                maintainAspectRatio: false,

                plugins: {

                    legend: {

                        position: "top",

                        labels: {

                            usePointStyle: true,

                            font: {

                                size: 10

                            }

                        }

                    },

                    tooltip: {

                        callbacks: {

                            label:
                                function(
                                    context
                                ) {

                                    const point =
                                        context.raw;


                                    return (

                                        ` ${point.date} | ` +

                                        `GHI: ${formatNumber(
                                            point.x,
                                            2
                                        )} | ` +

                                        `Generation: ${formatNumber(
                                            point.y,
                                            2
                                        )} MWh`

                                    );

                                }

                        }

                    }

                },

                scales: {

                    x: {

                        title: {

                            display: true,

                            text:
                                "GHI (kWh/m²)"

                        },

                        grid: {

                            color:
                                "rgba(23,37,42,0.07)"

                        }

                    },

                    y: {

                        title: {

                            display: true,

                            text:
                                "Actual Generation (MWh)"

                        },

                        grid: {

                            color:
                                "rgba(23,37,42,0.07)"

                        }

                    }

                }

            }

        }

    );

}


/* =========================================================
   FILTER DATA
   ========================================================= */

function applyFilters() {

    if (!allData.length) {
        return;
    }


    const from =
        el("fromDate").value;

    const to =
        el("toDate").value;


    /*
       Custom date range
    */

    if (from || to) {

        const fromDate =
            from
                ? new Date(
                    from +
                    "T00:00:00"
                )
                : null;


        const toDate =
            to
                ? new Date(
                    to +
                    "T23:59:59"
                )
                : null;


        filteredData =
            allData.filter(
                row => {

                    return (

                        (!fromDate ||
                            row.date >= fromDate)

                        &&

                        (!toDate ||
                            row.date <= toDate)

                    );

                }
            );

    }

    else {

        const selected =
            document.querySelector(
                ".range-button.selected"
            );


        const range =
            selected
                ? selected.dataset.range
                : "30";


        if (range === "all") {

            filteredData =
                [...allData];

        }

        else {

            filteredData =
                allData.slice(
                    -Number(range)
                );

        }

    }


    renderDashboard();

}


/* =========================================================
   RENDER EVERYTHING
   ========================================================= */

function renderDashboard() {

    if (!filteredData.length) {

        return;

    }


    const labels =
        filteredData.map(
            row =>
                shortDate(
                    row.date
                )
        );


    /*
       PA
    */

    drawPA(

        labels,

        filteredData.map(
            row =>
                row.pa
        )

    );


    /*
       PR
    */

    drawPR(

        labels,

        filteredData.map(
            row =>
                row.pr
        )

    );


    /*
       AEY
    */

    const aey =
        calculateAEY(
            filteredData
        );


    drawAEY(
        labels,
        aey
    );


    /*
       GHI
    */

    drawGHI(

        labels,

        filteredData.map(
            row =>
                row.ghi
        )

    );


    /*
       POA / GII
    */

    drawGII(

        labels,

        filteredData.map(
            row =>
                row.poaUp
        ),

        filteredData.map(
            row =>
                row.poaDown
        )

    );


    /*
       Generation
    */

    drawGeneration(

        labels,

        filteredData.map(
            row =>
                row.measured
        ),

        filteredData.map(
            row =>
                row.actual
        )

    );


    /*
       Scatter
    */

    drawScatter(
        filteredData
    );


    /*
       KPIs
    */

    updateKPIs(
        aey
    );


    /*
       Status
    */

    const first =
        filteredData[0];

    const last =
        filteredData[
            filteredData.length - 1
        ];


    el("dataStatus").textContent =

        `${allData.length} DGR records loaded • ` +

        `Showing ${filteredData.length} days • ` +

        `${isoDate(first.date)} → ${isoDate(last.date)}`;

}


/* =========================================================
   KPI UPDATE
   ========================================================= */

function updateKPIs(
    aey
) {

    if (!filteredData.length) {
        return;
    }


    const latest =
        filteredData[
            filteredData.length - 1
        ];


    /*
       PA
    */

    el("kpiPA").textContent =

        latest.pa !== null

            ? formatNumber(
                latest.pa,
                2
            ) + "%"

            : "—";


    /*
       PR
    */

    el("kpiPR").textContent =

        latest.pr !== null

            ? formatNumber(
                latest.pr,
                2
            ) + "%"

            : "—";


    /*
       AEY
    */

    const latestAEY =
        aey[
            aey.length - 1
        ];


    el("kpiAEY").textContent =

        latestAEY !== null &&
        latestAEY !== undefined

            ? formatNumber(
                latestAEY,
                2
            )

            : "—";


    /*
       Generation
    */

    el("kpiGen").textContent =

        latest.actual !== null

            ? formatNumber(
                latest.actual / 1000,
                2
            ) + " MWh"

            : "—";


    /*
       Capacity
    */

    const capacityRow =
        allData.find(
            row =>
                row.capacity !== null
        );


    if (
        capacityRow &&
        capacityRow.capacity
    ) {

        el("plantCapacity").textContent =

            `${formatNumber(
                capacityRow.capacity,
                2
            )} MWp DC`;

    }

}


/* =========================================================
   MAPPING DISPLAY
   ========================================================= */

function displayMapping(
    sheetName,
    columns
) {

    el("mappingInfo").innerHTML = `

        <strong>
            DGR successfully loaded
        </strong>

        <br>

        Sheet:
        ${sheetName}

        <br>

        Date:
        ${columns.date || "Not found"}

        <br>

        PA:
        ${columns.pa || "Not found"}

        ·

        PR:
        ${columns.pr || "Not found"}

        ·

        GHI:
        ${columns.ghi || "Not found"}

        <br>

        POA Up:
        ${columns.poaUp || "Not found"}

        ·

        POA Down:
        ${columns.poaDown || "Not found"}

        <br>

        Measured:
        ${columns.measured || "Not found"}

        ·

        Actual:
        ${columns.actual || "Not found"}

    `;

}


/* =========================================================
   LOAD DGR
   ========================================================= */

async function loadDGR(
    file
) {

    if (!file) {
        return;
    }


    try {

        el("dataStatus").textContent =
            "Reading DGR...";


        /*
         * Read file.
         */

        const buffer =
            await file.arrayBuffer();


        /*
         * Parse workbook.
         */

        const workbook =
            XLSX.read(
                buffer,
                {
                    type: "array",
                    cellDates: true
                }
            );


        /*
         * Find best worksheet.
         */

        const best =
            findBestSheet(
                workbook
            );


        if (!best) {

            throw new Error(

                "No worksheet containing usable DGR data was found."

            );

        }


        /*
         * Check minimum requirement.
         */

        if (!best.columns.date) {

            throw new Error(

                "The DGR does not contain a Date column."

            );

        }


        /*
         * Normalise.
         */

        allData =
            normalise(
                best.result
            );


        if (!allData.length) {

            throw new Error(

                "The worksheet was found, but no valid daily records could be read."

            );

        }


        /*
         * Display detected fields.
         */

        displayMapping(
            best.sheetName,
            best.columns
        );


        /*
         * Date limits.
         */

        const first =
            allData[0].date;

        const last =
            allData[
                allData.length - 1
            ].date;


        el("fromDate").min =
            isoDate(first);

        el("fromDate").max =
            isoDate(last);

        el("toDate").min =
            isoDate(first);

        el("toDate").max =
            isoDate(last);


        /*
         * Apply current filter.
         */

        applyFilters();


        /*
         * Success message.
         */

        el("dataStatus").textContent =

            `DGR loaded successfully • ` +

            `${allData.length} daily records • ` +

            `Sheet: ${best.sheetName}`;

    }

    catch (error) {

        console.error(
            error
        );


        el("dataStatus").textContent =
            "DGR could not be loaded.";


        alert(

            "DGR upload error:\n\n" +
            error.message

        );

    }

}


/* =========================================================
   FILE UPLOAD
   ========================================================= */

el("dgrInput")
    .addEventListener(
        "change",
        event => {

            const file =
                event.target.files[0];


            loadDGR(
                file
            );

        }
    );


/* =========================================================
   DRAG & DROP
   ========================================================= */

const dropzone =
    el("dropzone");


if (dropzone) {

    dropzone.addEventListener(
        "click",
        () => {

            el("dgrInput").click();

        }
    );


    dropzone.addEventListener(
        "dragover",
        event => {

            event.preventDefault();

            dropzone.classList.add(
                "dragging"
            );

        }
    );


    dropzone.addEventListener(
        "dragleave",
        () => {

            dropzone.classList.remove(
                "dragging"
            );

        }
    );


    dropzone.addEventListener(
        "drop",
        event => {

            event.preventDefault();


            dropzone.classList.remove(
                "dragging"
            );


            const file =
                event.dataTransfer.files[0];


            loadDGR(
                file
            );

        }
    );

}


/* =========================================================
   RANGE BUTTONS
   ========================================================= */

document
    .querySelectorAll(
        ".range-button"
    )
    .forEach(
        button => {

            button.addEventListener(
                "click",
                () => {

                    document
                        .querySelectorAll(
                            ".range-button"
                        )
                        .forEach(
                            item =>
                                item.classList.remove(
                                    "selected"
                                )
                        );


                    button.classList.add(
                        "selected"
                    );


                    /*
                     * Clear custom dates.
                     */

                    el("fromDate").value =
                        "";

                    el("toDate").value =
                        "";


                    applyFilters();

                }
            );

        }
    );


/* =========================================================
   CUSTOM DATE
   ========================================================= */

el("applyDate")
    .addEventListener(
        "click",
        () => {

            const from =
                el("fromDate").value;

            const to =
                el("toDate").value;


            if (
                from &&
                to &&
                from > to
            ) {

                alert(
                    "From date cannot be later than To date."
                );

                return;

            }


            document
                .querySelectorAll(
                    ".range-button"
                )
                .forEach(
                    button =>
                        button.classList.remove(
                            "selected"
                        )
                );


            applyFilters();

        }
    );


/* =========================================================
   RESET
   ========================================================= */

el("resetDate")
    .addEventListener(
        "click",
        () => {

            el("fromDate").value =
                "";

            el("toDate").value =
                "";


            document
                .querySelectorAll(
                    ".range-button"
                )
                .forEach(
                    button =>
                        button.classList.remove(
                            "selected"
                        )
                );


            const defaultButton =
                document.querySelector(
                    '.range-button[data-range="30"]'
                );


            if (defaultButton) {

                defaultButton.classList.add(
                    "selected"
                );

            }


            applyFilters();

        }
    );


/* =========================================================
   INITIAL STATE
   ========================================================= */

window.addEventListener(
    "DOMContentLoaded",
    () => {

        el("dataStatus").textContent =

            "No DGR loaded — upload an Excel or CSV Daily Generation Report.";

    }
);
