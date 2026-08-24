/* =========================================================
   SRI1PL SOLAR DGR DASHBOARD
   APPLICATION JAVASCRIPT
   ========================================================= */


/* =========================================================
   GLOBAL VARIABLES
   ========================================================= */

let rows = [];

let filteredRows = [];

const charts = {};


/* =========================================================
   DGR COLUMN MAPPING
   Based on the uploaded SRI1PL DGR
   ========================================================= */

const FIELD = {

    date:
        "Date",

    plantAvailability:
        "PA(%)",

    performanceRatio:
        "PR(%)",

    ghi:
        "GHI-UP (KWh/m2)",

    poaUp:
        "POA-UP(KWh/m2)",

    poaDown:
        "POA-Down(KWh/m2)",

    measuredGeneration:
        "Inv_Exp (kWh)",

    actualGeneration:
        "220kV_Net_Exp (KWh)",

    dcCapacity:
        "Firm DC Capacity (MWp)"

};


/* =========================================================
   SHORTCUT FOR HTML ELEMENTS
   ========================================================= */

function $(id) {

    return document.getElementById(id);

}


/* =========================================================
   NUMBER FORMATTER
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


/* =========================================================
   PERCENTAGE FORMATTER
   ========================================================= */

function formatPercentage(value) {

    if (!Number.isFinite(value)) {

        return "—";

    }

    return (
        value * 100
    ).toLocaleString(
        "en-IN",
        {
            minimumFractionDigits: 2,
            maximumFractionDigits: 2
        }
    ) + "%";

}


/* =========================================================
   DATE CONVERSION
   ========================================================= */

function convertToDate(value) {

    if (value instanceof Date) {

        return value;

    }


    /*
       Excel serial date
    */

    if (typeof value === "number") {

        try {

            const parsed =
                XLSX.SSF.parse_date_code(value);

            return new Date(
                parsed.y,
                parsed.m - 1,
                parsed.d
            );

        }

        catch (error) {

            return null;

        }

    }


    /*
       Normal date string
    */

    const date =
        new Date(value);


    if (isNaN(date.getTime())) {

        return null;

    }


    return date;

}


/* =========================================================
   ISO DATE
   ========================================================= */

function getISODate(date) {

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


/* =========================================================
   DISPLAY DATE
   ========================================================= */

function displayDate(date) {

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
   NUMBER CLEANER
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

        return value;

    }


    const cleaned =
        String(value)
            .replace(/,/g, "")
            .replace(/%/g, "")
            .trim();


    const number =
        parseFloat(cleaned);


    if (
        Number.isFinite(number)
    ) {

        return number;

    }


    return NaN;

}


/* =========================================================
   NORMALISE DGR DATA
   ========================================================= */

function normaliseData(rawData) {

    const normalised = [];


    rawData.forEach(row => {

        const date =
            convertToDate(
                row[FIELD.date]
            );


        /*
           Ignore rows without valid dates.
        */

        if (!date) {

            return;

        }


        const capacity =
            cleanNumber(
                row[FIELD.dcCapacity]
            );


        normalised.push({

            date: date,

            pa:
                cleanNumber(
                    row[FIELD.plantAvailability]
                ),

            pr:
                cleanNumber(
                    row[FIELD.performanceRatio]
                ),

            ghi:
                cleanNumber(
                    row[FIELD.ghi]
                ),

            poaUp:
                cleanNumber(
                    row[FIELD.poaUp]
                ),

            poaDown:
                cleanNumber(
                    row[FIELD.poaDown]
                ),

            measured:
                cleanNumber(
                    row[FIELD.measuredGeneration]
                ),

            actual:
                cleanNumber(
                    row[FIELD.actualGeneration]
                ),

            capacity:
                capacity

        });

    });


    /*
       Sort chronologically.
    */

    normalised.sort(
        (a, b) =>
            a.date - b.date
    );


    return normalised;

}


/* =========================================================
   CREATE CHART
   ========================================================= */

function createChart(
    chartId,
    chartType,
    labels,
    datasets,
    options = {}
) {

    /*
       Destroy existing chart first.
    */

    if (charts[chartId]) {

        charts[chartId].destroy();

    }


    const canvas =
        $(chartId);


    if (!canvas) {

        return;

    }


    charts[chartId] =
        new Chart(
            canvas,
            {

                type: chartType,

                data: {

                    labels: labels,

                    datasets: datasets

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

                            display: true,

                            position: "top",

                            labels: {

                                usePointStyle: true,

                                boxWidth: 7,

                                padding: 16,

                                font: {

                                    size: 10

                                }

                            }

                        },

                        tooltip: {

                            enabled: true,

                            backgroundColor:
                                "#17252A",

                            titleColor:
                                "#FFFFFF",

                            bodyColor:
                                "#FFFFFF",

                            padding: 10,

                            cornerRadius: 8

                        }

                    },

                    scales: {

                        x: {

                            grid: {

                                display: false

                            },

                            ticks: {

                                font: {

                                    size: 9

                                },

                                maxRotation: 0

                            }

                        },

                        y: {

                            beginAtZero: false,

                            grid: {

                                color:
                                    "rgba(23,37,42,0.06)"

                            },

                            ticks: {

                                font: {

                                    size: 9

                                }

                            }

                        }

                    },

                    ...options

                }

            }

        );

}


/* =========================================================
   DATASET BUILDER
   ========================================================= */

function lineDataset(
    label,
    data
) {

    return {

        label: label,

        data: data,

        borderWidth: 2,

        pointRadius: 2,

        pointHoverRadius: 5,

        tension: 0.3,

        fill: false

    };

}


/* =========================================================
   GET FILTERED DATA
   ========================================================= */

function getFilteredRows() {

    if (!rows.length) {

        return [];

    }


    const fromValue =
        $("fromDate").value;


    const toValue =
        $("toDate").value;


    /*
       Custom date filter
    */

    if (
        fromValue ||
        toValue
    ) {

        const fromDate =
            fromValue
                ? new Date(
                    fromValue +
                    "T00:00:00"
                )
                : null;


        const toDate =
            toValue
                ? new Date(
                    toValue +
                    "T23:59:59"
                )
                : null;


        return rows.filter(row => {

            return (

                (!fromDate ||
                    row.date >= fromDate)

                &&

                (!toDate ||
                    row.date <= toDate)

            );

        });

    }


    /*
       Preset ranges
    */

    const selectedButton =
        document.querySelector(
            ".range-button.selected"
        );


    const selectedRange =
        selectedButton
            ? selectedButton.dataset.range
            : "30";


    if (
        selectedRange === "all"
    ) {

        return [...rows];

    }


    const numberOfDays =
        Number(selectedRange);


    return rows.slice(
        -numberOfDays
    );

}


/* =========================================================
   CALCULATE AEY
   =========================================================

   AEY =
   cumulative actual generation
   /
   installed DC capacity

   Generation = kWh

   Capacity = MWp

   MWp × 1000 = kWp
   ========================================================= */

function calculateAEY(data) {

    let cumulativeGeneration = 0;


    return data.map(row => {

        if (
            Number.isFinite(
                row.actual
            )
        ) {

            cumulativeGeneration +=
                row.actual;

        }


        const capacity =
            row.capacity ||
            data.find(
                r =>
                    Number.isFinite(
                        r.capacity
                    )
            )?.capacity;


        if (
            !capacity ||
            !Number.isFinite(capacity)
        ) {

            return NaN;

        }


        const capacityKWp =
            capacity * 1000;


        return (
            cumulativeGeneration /
            capacityKWp
        );

    });

}


/* =========================================================
   RENDER DASHBOARD
   ========================================================= */

function renderDashboard() {

    filteredRows =
        getFilteredRows();


    if (
        !filteredRows.length
    ) {

        return;

    }


    /*
       Labels
    */

    const labels =
        filteredRows.map(
            row =>
                displayDate(
                    row.date
                )
        );


    /* =====================================================
       PLANT AVAILABILITY
    ====================================================== */

    createChart(

        "paChart",

        "line",

        labels,

        [

            lineDataset(
                "Plant Availability (%)",

                filteredRows.map(
                    row =>
                        Number.isFinite(row.pa)
                            ? row.pa * 100
                            : null
                )

            )

        ],

        {

            scales: {

                y: {

                    min: 0,

                    max: 100,

                    title: {

                        display: true,

                        text: "Availability (%)"

                    }

                }

            }

        }

    );


    /* =====================================================
       PERFORMANCE RATIO
    ====================================================== */

    createChart(

        "prChart",

        "line",

        labels,

        [

            lineDataset(
                "Performance Ratio (%)",

                filteredRows.map(
                    row =>
                        Number.isFinite(row.pr)
                            ? row.pr * 100
                            : null
                )

            )

        ],

        {

            scales: {

                y: {

                    beginAtZero: false,

                    title: {

                        display: true,

                        text: "PR (%)"

                    }

                }

            }

        }

    );


    /* =====================================================
       GHI
    ====================================================== */

    createChart(

        "ghiChart",

        "line",

        labels,

        [

            lineDataset(
                "GHI (kWh/m²)",

                filteredRows.map(
                    row =>
                        Number.isFinite(row.ghi)
                            ? row.ghi
                            : null
                )

            )

        ],

        {

            scales: {

                y: {

                    beginAtZero: true,

                    title: {

                        display: true,

                        text: "kWh/m²"

                    }

                }

            }

        }

    );


    /* =====================================================
       GII / POA
    ====================================================== */

    createChart(

        "giiChart",

        "line",

        labels,

        [

            lineDataset(
                "POA Up (kWh/m²)",

                filteredRows.map(
                    row =>
                        Number.isFinite(row.poaUp)
                            ? row.poaUp
                            : null
                )

            ),

            lineDataset(
                "POA Down (kWh/m²)",

                filteredRows.map(
                    row =>
                        Number.isFinite(row.poaDown)
                            ? row.poaDown
                            : null
                )

            )

        ],

        {

            scales: {

                y: {

                    beginAtZero: true,

                    title: {

                        display: true,

                        text: "kWh/m²"

                    }

                }

            }

        }

    );


    /* =====================================================
       MEASURED VS ACTUAL GENERATION
    ====================================================== */

    createChart(

        "generationChart",

        "line",

        labels,

        [

            lineDataset(
                "Measured / Inverter Export",

                filteredRows.map(
                    row =>
                        Number.isFinite(row.measured)
                            ? row.measured
                            : null
                )

            ),

            lineDataset(
                "Actual / 220 kV Net Export",

                filteredRows.map(
                    row =>
                        Number.isFinite(row.actual)
                            ? row.actual
                            : null
                )

            )

        ],

        {

            scales: {

                y: {

                    beginAtZero: true,

                    title: {

                        display: true,

                        text: "Generation (kWh)"

                    }

                }

            }

        }

    );


    /* =====================================================
       AEY
    ====================================================== */

    const aeyValues =
        calculateAEY(
            filteredRows
        );


    createChart(

        "aeyChart",

        "line",

        labels,

        [

            lineDataset(
                "Cumulative AEY (kWh/kWp)",

                aeyValues

            )

        ],

        {

            scales: {

                y: {

                    beginAtZero: true,

                    title: {

                        display: true,

                        text: "kWh/kWp"

                    }

                }

            }

        }

    );


    /* =====================================================
       UPDATE KPI CARDS
    ====================================================== */

    updateKPIs(
        filteredRows,
        aeyValues
    );


    /* =====================================================
       UPDATE DATA STATUS
    ====================================================== */

    const latest =
        filteredRows[
            filteredRows.length - 1
        ];


    $("dataStatus").textContent =

        `${rows.length} DGR days loaded • ` +

        `Showing ${filteredRows.length} days • ` +

        `Latest: ${getISODate(latest.date)}`;

}


/* =========================================================
   UPDATE KPI CARDS
   ========================================================= */

function updateKPIs(
    data,
    aeyValues
) {

    if (!data.length) {

        return;

    }


    const latest =
        data[data.length - 1];


    /* =====================================================
       PA
    ====================================================== */

    $("kpiPA").textContent =
        formatPercentage(
            latest.pa
        );


    $("kpiPADetail").textContent =
        `Latest: ${getISODate(latest.date)}`;


    /* =====================================================
       PR
    ====================================================== */

    $("kpiPR").textContent =
        formatPercentage(
            latest.pr
        );


    $("kpiPRDetail").textContent =
        `Latest: ${getISODate(latest.date)}`;


    /* =====================================================
       ACTUAL GENERATION
    ====================================================== */

    if (
        Number.isFinite(
            latest.actual
        )
    ) {

        const generationMWh =
            latest.actual / 1000;


        $("kpiGen").textContent =
            formatNumber(
                generationMWh,
                2
            ) + " MWh";

    }

    else {

        $("kpiGen").textContent =
            "—";

    }


    $("kpiGenDetail").textContent =
        "220 kV net export";


    /* =====================================================
       AEY
    ====================================================== */

    const latestAEY =
        aeyValues[
            aeyValues.length - 1
        ];


    $("kpiAEY").textContent =
        formatNumber(
            latestAEY,
            2
        );


}


/* =========================================================
   FIND DGR HEADER ROW
   ========================================================= */

function findHeaderRow(
    worksheet
) {

    const matrix =
        XLSX.utils.sheet_to_json(
            worksheet,
            {
                header: 1,
                raw: true,
                defval: null
            }
        );


    /*
       Search first 20 rows.
    */

    for (
        let i = 0;
        i < Math.min(
            matrix.length,
            20
        );
        i++
    ) {

        const row =
            matrix[i] || [];


        const headers =
            row.map(
                value =>
                    String(
                        value ?? ""
                    ).trim()
            );


        if (

            headers.includes("Date")

            &&

            headers.includes("PA(%)")

            &&

            headers.includes("PR(%)")

        ) {

            return i;

        }

    }


    return -1;

}


/* =========================================================
   PARSE WORKBOOK
   ========================================================= */

function parseWorkbook(
    workbook
) {

    let worksheet =
        workbook.Sheets[
            "Daily_KPI"
        ];


    let headerRow = -1;


    /*
       First attempt:
       supplied DGR structure.
    */

    if (worksheet) {

        headerRow = 3;

    }


    /*
       Fallback:
       search all worksheets.
    */

    if (
        !worksheet ||
        headerRow < 0
    ) {

        for (
            const sheetName
            of workbook.SheetNames
        ) {

            const sheet =
                workbook.Sheets[
                    sheetName
                ];


            const found =
                findHeaderRow(
                    sheet
                );


            if (found >= 0) {

                worksheet = sheet;

                headerRow = found;

                break;

            }

        }

    }


    if (!worksheet) {

        throw new Error(

            "No suitable DGR sheet was found."

        );

    }


    if (
        headerRow < 0
    ) {

        throw new Error(

            "Could not find a header row containing Date, PA(%) and PR(%)."

        );

    }


    const matrix =
        XLSX.utils.sheet_to_json(
            worksheet,
            {
                header: 1,
                raw: true,
                defval: null
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


    const data = [];


    for (
        let i = headerRow + 1;
        i < matrix.length;
        i++
    ) {

        const row =
            matrix[i];


        if (
            !row ||
            !row.length
        ) {

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


        data.push(
            object
        );

    }


    return {

        data: data,

        headers: headers,

        sheetName:
            workbook.SheetNames.find(
                name =>
                    workbook.Sheets[name]
                    === worksheet
            ) || "Daily_KPI"

    };

}


/* =========================================================
   DISPLAY DATA MAPPING
   ========================================================= */

function showMappingInfo(
    parsed
) {

    $("mappingInfo").innerHTML = `

        <strong>
            Detected sheet:
        </strong>

        ${parsed.sheetName}

        <br>

        <strong>
            Date:
        </strong>

        ${FIELD.date}

        ·

        <strong>
            PA:
        </strong>

        ${FIELD.plantAvailability}

        ·

        <strong>
            PR:
        </strong>

        ${FIELD.performanceRatio}

        <br>

        <strong>
            GHI:
        </strong>

        ${FIELD.ghi}

        ·

        <strong>
            GII / POA:
        </strong>

        ${FIELD.poaUp}

        /

        ${FIELD.poaDown}

        <br>

        <strong>
            Measured Generation:
        </strong>

        ${FIELD.measuredGeneration}

        ·

        <strong>
            Actual Generation:
        </strong>

        ${FIELD.actualGeneration}

    `;

}


/* =========================================================
   UPDATE DATE PICKERS
   ========================================================= */

function updateDateLimits() {

    if (!rows.length) {

        return;

    }


    const firstDate =
        getISODate(
            rows[0].date
        );


    const lastDate =
        getISODate(
            rows[
                rows.length - 1
            ].date
        );


    $("fromDate").min =
        firstDate;

    $("fromDate").max =
        lastDate;


    $("toDate").min =
        firstDate;

    $("toDate").max =
        lastDate;

}


/* =========================================================
   LOAD DGR FILE
   ========================================================= */

async function loadDGRFile(
    file
) {

    if (!file) {

        return;

    }


    try {

        /*
           Read file in browser.
        */

        const buffer =
            await file.arrayBuffer();


        /*
           Read Excel workbook.
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
           Parse DGR.
        */

        const parsed =
            parseWorkbook(
                workbook
            );


        /*
           Convert raw data.
        */

        rows =
            normaliseData(
                parsed.data
            );


        if (!rows.length) {

            throw new Error(

                "No usable daily records were found in the DGR."

            );

        }


        /*
           Find DC capacity.
        */

        const capacityRow =
            rows.find(
                row =>
                    Number.isFinite(
                        row.capacity
                    )
            );


        if (
            capacityRow &&
            Number.isFinite(
                capacityRow.capacity
            )
        ) {

            $("plantCapacity").textContent =

                `${formatNumber(
                    capacityRow.capacity,
                    2
                )} MWp DC`;

        }


        /*
           Show mapping.
        */

        showMappingInfo(
            parsed
        );


        /*
           Update date limits.
        */

        updateDateLimits();


        /*
           Render dashboard.
        */

        renderDashboard();


        /*
           Update status.
        */

        $("dataStatus").textContent =

            `${rows.length} DGR days loaded successfully.`;


        /*
           Scroll to dashboard.
        */

        $("dashboard").scrollIntoView({

            behavior: "smooth",

            block: "start"

        });

    }

    catch (error) {

        console.error(
            "DGR Error:",
            error
        );


        alert(

            "DGR upload failed.\n\n" +
            error.message

        );

    }

}


/* =========================================================
   FILE INPUT
   ========================================================= */

$("dgrInput")
    .addEventListener(
        "change",
        event => {

            const file =
                event.target.files[0];


            loadDGRFile(
                file
            );

        }
    );


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

                    /*
                       Remove active state.
                    */

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


                    /*
                       Activate clicked button.
                    */

                    button.classList.add(
                        "selected"
                    );


                    /*
                       Clear custom dates.
                    */

                    $("fromDate").value =
                        "";

                    $("toDate").value =
                        "";


                    /*
                       Render.
                    */

                    renderDashboard();

                }
            );

        }
    );


/* =========================================================
   APPLY CUSTOM DATE
   ========================================================= */

$("applyDate")
    .addEventListener(
        "click",
        () => {

            const from =
                $("fromDate").value;

            const to =
                $("toDate").value;


            if (
                from &&
                to &&
                from > to
            ) {

                alert(
                    "The 'From' date cannot be after the 'To' date."
                );

                return;

            }


            /*
               Remove preset selection.
            */

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


            renderDashboard();

        }
    );


/* =========================================================
   RESET FILTER
   ========================================================= */

$("resetDate")
    .addEventListener(
        "click",
        () => {

            $("fromDate").value =
                "";

            $("toDate").value =
                "";


            /*
               Restore 30 days.
            */

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


            const thirtyDayButton =
                document.querySelector(
                    '.range-button[data-range="30"]'
                );


            if (
                thirtyDayButton
            ) {

                thirtyDayButton.classList.add(
                    "selected"
                );

            }


            renderDashboard();

        }
    );


/* =========================================================
   DRAG & DROP
   ========================================================= */

const dropzone =
    $("dropzone");


if (dropzone) {


    /*
       Click upload area.
    */

    dropzone.addEventListener(
        "click",
        () => {

            $("dgrInput").click();

        }
    );


    /*
       Drag enter.
    */

    dropzone.addEventListener(
        "dragenter",
        event => {

            event.preventDefault();

            dropzone.classList.add(
                "dragging"
            );

        }
    );


    /*
       Drag over.
    */

    dropzone.addEventListener(
        "dragover",
        event => {

            event.preventDefault();

            dropzone.classList.add(
                "dragging"
            );

        }
    );


    /*
       Drag leave.
    */

    dropzone.addEventListener(
        "dragleave",
        event => {

            event.preventDefault();

            dropzone.classList.remove(
                "dragging"
            );

        }
    );


    /*
       Drop.
    */

    dropzone.addEventListener(
        "drop",
        event => {

            event.preventDefault();


            dropzone.classList.remove(
                "dragging"
            );


            const file =
                event.dataTransfer.files[0];


            loadDGRFile(
                file
            );

        }
    );

}


/* =========================================================
   INITIAL STATE
   ========================================================= */

document.addEventListener(
    "DOMContentLoaded",
    () => {

        /*
           Dashboard starts empty.
           User uploads a DGR to populate it.
        */

        $("dataStatus").textContent =

            "No DGR loaded. Upload the Daily Generation Report to begin analysis.";

    }
);
