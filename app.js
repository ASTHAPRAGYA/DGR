```javascript
/* =========================================================
   SOLAR DGR ANALYTICS
   app.js

   EXACT DATA SOURCES

   DAILY_KPI
   ---------------------------------------------------------
   B  = Date
   I  = Operating Hours
   S  = PA (%)
   V  = PR (%)
   AD = System Losses (%)

   PA
   ---------------------------------------------------------
   B  = Date
   W  = Issue / Fault
   Z  = Fault Start Time
   AC = Work Completion Time
   AG = Breakdown Time
   AL = System Loss / Energy Loss

   CURTAILMENT RECORDS
   ---------------------------------------------------------
   C  = Date
   H  = Start Time
   I  = End Time
   R  = Loss of Generation MWh

   ANNUAL_KPI
   ---------------------------------------------------------
   H9 = Target PR (%)
   I9 = Measured PR (%)
   H10:I21 = Monthly PR data

   E10:E21 = Budgeted Energy
   F10:F21 = Measured Energy (Exp)

   ========================================================= */

"use strict";


/* =========================================================
   GLOBAL
========================================================= */

let workbook = null;

const charts = {};


/* =========================================================
   DOM HELPER
========================================================= */

function $(id) {
    return document.getElementById(id);
}


/* =========================================================
   INITIALISE
========================================================= */

document.addEventListener(
    "DOMContentLoaded",
    () => {

        installChartScrollStyles();

        setupNavigation();

        setupUpload();

        setupRemoveButton();

        hideAnalytics();

    }
);


/* =========================================================
   CHART SCROLL CSS
========================================================= */

function installChartScrollStyles() {

    if (
        document.getElementById(
            "dgr-scroll-styles"
        )
    ) {

        return;

    }


    const style =
        document.createElement(
            "style"
        );


    style.id =
        "dgr-scroll-styles";


    style.textContent = `

        .dgr-chart-scroll {
            width: 100%;
            height: 100%;
            overflow-x: auto;
            overflow-y: hidden;
            position: relative;
            box-sizing: border-box;
            scrollbar-width: thin;
            scrollbar-color: #b9cacc transparent;
        }

        .dgr-chart-scroll::-webkit-scrollbar {
            height: 8px;
        }

        .dgr-chart-scroll::-webkit-scrollbar-track {
            background: #eef3f4;
            border-radius: 10px;
        }

        .dgr-chart-scroll::-webkit-scrollbar-thumb {
            background: #b9cacc;
            border-radius: 10px;
        }

        .dgr-chart-scroll::-webkit-scrollbar-thumb:hover {
            background: #27a5ad;
        }

        .dgr-chart-scroll canvas {
            display: block !important;
            max-width: none !important;
        }

    `;


    document.head.appendChild(
        style
    );

}


/* =========================================================
   NAVIGATION
========================================================= */

function setupNavigation() {

    const items =
        document.querySelectorAll(
            ".nav-item"
        );


    items.forEach(
        item => {

            item.addEventListener(
                "click",
                () => {

                    items.forEach(
                        nav =>
                            nav.classList.remove(
                                "active"
                            )
                    );


                    item.classList.add(
                        "active"
                    );


                    const target =
                        $(
                            item.dataset.target
                        );


                    if (target) {

                        target.scrollIntoView(
                            {
                                behavior:
                                    "smooth",

                                block:
                                    "start"
                            }
                        );

                    }

                }
            );

        }
    );

}


/* =========================================================
   UPLOAD
========================================================= */

function setupUpload() {

    const input =
        $("dgrFile");


    const dropZone =
        $("dropZone");


    if (!input) {

        console.error(
            "dgrFile input not found."
        );

        return;

    }


    input.addEventListener(
        "change",
        event => {

            const file =
                event.target.files?.[0];


            if (file) {

                processFile(
                    file
                );

            }

        }
    );


    if (dropZone) {

        dropZone.addEventListener(
            "click",
            () => {

                input.click();

            }
        );


        dropZone.addEventListener(
            "dragover",
            event => {

                event.preventDefault();

                dropZone.classList.add(
                    "dragging"
                );

            }
        );


        dropZone.addEventListener(
            "dragleave",
            () => {

                dropZone.classList.remove(
                    "dragging"
                );

            }
        );


        dropZone.addEventListener(
            "drop",
            event => {

                event.preventDefault();

                dropZone.classList.remove(
                    "dragging"
                );


                const file =
                    event.dataTransfer
                        ?.files
                        ?. [0];


                if (file) {

                    processFile(
                        file
                    );

                }

            }
        );

    }

}


/* =========================================================
   REMOVE
========================================================= */

function setupRemoveButton() {

    const button =
        $("removeFile");


    if (!button) {
        return;
    }


    button.addEventListener(
        "click",
        resetDashboard
    );

}


/* =========================================================
   PROCESS FILE
========================================================= */

function processFile(
    file
) {

    if (
        !/\.(xlsx|xls|csv)$/i.test(
            file.name
        )
    ) {

        alert(
            "Please upload an Excel file (.xlsx/.xls) or CSV file."
        );

        return;

    }


    if (
        typeof XLSX === "undefined"
    ) {

        alert(
            "SheetJS is not loaded. Check index.html."
        );

        return;

    }


    setStatus(
        "Reading DGR workbook..."
    );


    const reader =
        new FileReader();


    reader.onload =
        event => {

            try {

                workbook =
                    XLSX.read(
                        new Uint8Array(
                            event.target.result
                        ),
                        {
                            type:
                                "array",

                            cellDates:
                                true,

                            cellNF:
                                true,

                            cellText:
                                true
                        }
                    );


                if (
                    !workbook.SheetNames ||
                    !workbook.SheetNames.length
                ) {

                    throw new Error(
                        "No worksheets found."
                    );

                }


                updateFileInformation(
                    file
                );


                showAnalytics();


                renderAll();


                setStatus(
                    `${file.name} loaded successfully.`
                );

            }

            catch (error) {

                console.error(
                    "DGR error:",
                    error
                );


                setStatus(
                    "Unable to read the DGR."
                );


                alert(
                    "Unable to read the uploaded DGR.\n\n" +
                    error.message
                );

            }

        };


    reader.onerror =
        () => {

            setStatus(
                "Unable to read the selected file."
            );

        };


    reader.readAsArrayBuffer(
        file
    );

}


/* =========================================================
   FILE INFORMATION
========================================================= */

function updateFileInformation(
    file
) {

    setText(
        "fileName",
        file.name
    );


    setText(
        "fileSheets",
        `${workbook.SheetNames.length} worksheets detected`
    );


    setText(
        "sidebarFileName",
        file.name
    );


    $("fileInfo")
        ?.classList
        .remove(
            "hidden"
        );


    $("workbookStatus")
        ?.classList
        .remove(
            "hidden"
        );


    $("emptyState")
        ?.classList
        .add(
            "hidden"
        );


    $("dropZone")
        ?.classList
        .add(
            "hidden"
        );


    renderSheetBadges();

}


/* =========================================================
   SHEET BADGES
========================================================= */

function renderSheetBadges() {

    const container =
        $("sheetBadges");


    if (
        !container ||
        !workbook
    ) {

        return;

    }


    container.innerHTML =
        "";


    const sheets = [
        "Dashboard",
        "Annual_KPI",
        "Daily_KPI",
        "PA",
        "Curtailment records"
    ];


    sheets.forEach(
        name => {

            const badge =
                document.createElement(
                    "span"
                );


            badge.className =
                "sheet-badge";


            const exists =
                workbook.SheetNames.some(
                    actual =>
                        normalizeSheetName(
                            actual
                        ) ===
                        normalizeSheetName(
                            name
                        )
                );


            badge.textContent =
                exists
                    ? `${name} ✓`
                    : `${name} — missing`;


            if (!exists) {

                badge.classList.add(
                    "missing"
                );

            }


            container.appendChild(
                badge
            );

        }
    );

}


/* =========================================================
   NORMALISE SHEET NAME
========================================================= */

function normalizeSheetName(
    name
) {

    return String(
        name || ""
    )
        .trim()
        .toLowerCase()
        .replace(
            /[\s_-]+/g,
            ""
        );

}


/* =========================================================
   GET SHEET
========================================================= */

function getSheet(
    requested
) {

    if (
        !workbook ||
        !workbook.Sheets
    ) {

        return null;

    }


    if (
        workbook.Sheets[
            requested
        ]
    ) {

        return workbook.Sheets[
            requested
        ];

    }


    const wanted =
        normalizeSheetName(
            requested
        );


    const actual =
        workbook.SheetNames.find(
            name =>
                normalizeSheetName(
                    name
                ) ===
                wanted
        );


    return actual
        ? workbook.Sheets[
            actual
        ]
        : null;

}


/* =========================================================
   DIRECT CELL ACCESS
========================================================= */

function readCell(
    sheetName,
    address
) {

    const sheet =
        getSheet(
            sheetName
        );


    if (!sheet) {
        return null;
    }


    const cell =
        sheet[
            address
        ];


    if (!cell) {
        return null;
    }


    /*
       Prefer the formatted/parsed value.
    */

    if (
        cell.v !== undefined &&
        cell.v !== null
    ) {

        return cell.v;

    }


    return null;

}


/* =========================================================
   NUMBER
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


    if (
        value instanceof Date
    ) {

        return null;

    }


    const text =
        String(
            value
        )
            .trim()
            .replace(
                /,/g,
                ""
            )
            .replace(
                /%/g,
                ""
            );


    if (
        !text ||
        text.startsWith("#")
    ) {

        return null;

    }


    const result =
        Number(
            text
        );


    return Number.isFinite(
        result
    )
        ? result
        : null;

}


/* =========================================================
   DATE
========================================================= */

function parseDate(
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
        value instanceof Date
    ) {

        return isNaN(
            value.getTime()
        )
            ? null
            : new Date(
                value.getTime()
            );

    }


    if (
        typeof value === "number"
    ) {

        try {

            const p =
                XLSX.SSF.parse_date_code(
                    value
                );


            if (p) {

                return new Date(
                    p.y,
                    p.m - 1,
                    p.d,
                    p.H || 0,
                    p.M || 0,
                    p.S || 0
                );

            }

        }

        catch (_) {}

    }


    const text =
        String(
            value
        )
            .trim();


    /*
       DD/MM/YYYY
    */

    let match =
        text.match(
            /^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})/
        );


    if (match) {

        const date =
            new Date(
                Number(
                    match[3]
                ),

                Number(
                    match[2]
                ) - 1,

                Number(
                    match[1]
                )
            );


        if (
            !isNaN(
                date.getTime()
            )
        ) {

            return date;

        }

    }


    /*
       DD-MMM-YYYY
    */

    match =
        text.match(
            /^(\d{1,2})[\/\-]([A-Za-z]{3,9})[\/\-](\d{2,4})/
        );


    if (match) {

        const months = [
            "jan",
            "feb",
            "mar",
            "apr",
            "may",
            "jun",
            "jul",
            "aug",
            "sep",
            "oct",
            "nov",
            "dec"
        ];


        const month =
            months.indexOf(
                match[2]
                    .substring(
                        0,
                        3
                    )
                    .toLowerCase()
            );


        let year =
            Number(
                match[3]
            );


        if (
            year < 100
        ) {

            year +=
                2000;

        }


        if (
            month >= 0
        ) {

            return new Date(
                year,
                month,
                Number(
                    match[1]
                )
            );

        }

    }


    const browserDate =
        new Date(
            text
        );


    return isNaN(
        browserDate.getTime()
    )
        ? null
        : browserDate;

}


/* =========================================================
   DATE FORMATTING
========================================================= */

function shortDate(
    date
) {

    if (
        !(date instanceof Date)
    ) {

        return "";

    }


    return date.toLocaleDateString(
        "en-IN",
        {
            day:
                "2-digit",

            month:
                "short"
        }
    );

}


function fullDate(
    date
) {

    if (
        !(date instanceof Date)
    ) {

        return "";

    }


    return date.toLocaleDateString(
        "en-IN",
        {
            day:
                "2-digit",

            month:
                "short",

            year:
                "numeric"
        }
    );

}


/* =========================================================
   TIME
========================================================= */

function parseTimeMinutes(
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
        value instanceof Date
    ) {

        return (
            value.getHours() * 60 +
            value.getMinutes()
        );

    }


    if (
        typeof value === "number"
    ) {

        if (
            value >= 0 &&
            value < 1
        ) {

            return Math.round(
                value * 1440
            );

        }


        try {

            const p =
                XLSX.SSF.parse_date_code(
                    value
                );


            if (p) {

                return (
                    p.H * 60 +
                    p.M
                );

            }

        }

        catch (_) {}

    }


    const text =
        String(
            value
        )
            .trim();


    const match =
        text.match(
            /^(\d{1,2}):(\d{2})(?::(\d{2}))?\s*(AM|PM)?$/i
        );


    if (!match) {

        return null;

    }


    let hours =
        Number(
            match[1]
        );


    const minutes =
        Number(
            match[2]
        );


    const ampm =
        match[4];


    if (ampm) {

        if (
            ampm.toUpperCase() ===
            "PM" &&
            hours < 12
        ) {

            hours +=
                12;

        }


        if (
            ampm.toUpperCase() ===
            "AM" &&
            hours === 12
        ) {

            hours =
                0;

        }

    }


    return (
        hours * 60 +
        minutes
    );

}


/* =========================================================
   TIME FORMAT
========================================================= */

function minutesToTime(
    value
) {

    const minutes =
        Math.max(
            0,
            Math.min(
                1439,
                Math.round(
                    value
                )
            )
        );


    const hour =
        Math.floor(
            minutes / 60
        );


    const minute =
        minutes % 60;


    return (
        String(
            hour
        ).padStart(
            2,
            "0"
        ) +
        ":" +
        String(
            minute
        ).padStart(
            2,
            "0"
        )
    );

}


/* =========================================================
   FORMAT
========================================================= */

function formatNumber(
    value,
    decimals = 2
) {

    if (
        value === null ||
        !Number.isFinite(
            value
        )
    ) {

        return "—";

    }


    return Number(
        value
    ).toLocaleString(
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
   SET TEXT
========================================================= */

function setText(
    id,
    value
) {

    const element =
        $(id);


    if (element) {

        element.textContent =
            value;

    }

}


/* =========================================================
   STATUS
========================================================= */

function setStatus(
    message
) {

    setText(
        "statusText",
        message
    );

}


/* =========================================================
   DESTROY CHART
========================================================= */

function destroyChart(
    id
) {

    if (
        charts[id] &&
        typeof charts[id].destroy ===
        "function"
    ) {

        try {

            charts[id].destroy();

        }

        catch (error) {

            console.warn(
                `Chart destroy failed: ${id}`,
                error
            );

        }

    }


    charts[id] =
        null;

}


/* =========================================================
   DESTROY ALL
========================================================= */

function destroyAllCharts() {

    Object.keys(
        charts
    )
        .forEach(
            destroyChart
        );

}


/* =========================================================
   SHOW / HIDE ANALYTICS
========================================================= */

function hideAnalytics() {

    [
        "dashboardSection",
        "paSection",
        "performanceSection",
        "curtailmentSection",
        "energySection"
    ]
        .forEach(
            id => {

                const element =
                    $(id);


                if (element) {

                    element.style.display =
                        "none";

                }

            }
        );

}


function showAnalytics() {

    [
        "dashboardSection",
        "paSection",
        "performanceSection",
        "curtailmentSection",
        "energySection"
    ]
        .forEach(
            id => {

                const element =
                    $(id);


                if (element) {

                    element.style.display =
                        "";

                }

            }
        );


    $("emptyState")
        ?.classList
        .add(
            "hidden"
        );

}


/* =========================================================
   SCROLL WRAPPER
========================================================= */

function makeChartScrollable(
    canvas,
    width
) {

    if (!canvas) {
        return null;
    }


    const parent =
        canvas.parentElement;


    if (!parent) {
        return null;
    }


    let wrapper =
        parent.querySelector(
            ".dgr-chart-scroll"
        );


    if (!wrapper) {

        wrapper =
            document.createElement(
                "div"
            );


        wrapper.className =
            "dgr-chart-scroll";


        parent.insertBefore(
            wrapper,
            canvas
        );


        wrapper.appendChild(
            canvas
        );

    }


    const actualWidth =
        Math.max(
            width || 800,
            parent.clientWidth || 800
        );


    canvas.style.setProperty(
        "width",
        `${actualWidth}px`,
        "important"
    );


    canvas.style.setProperty(
        "max-width",
        "none",
        "important"
    );


    canvas.style.height =
        "100%";


    return wrapper;

}


/* =========================================================
   DAILY KPI READER
   DIRECT CELL ACCESS
========================================================= */

function readDailyKPI() {

    const sheet =
        getSheet(
            "Daily_KPI"
        );


    if (!sheet) {

        console.warn(
            "Daily_KPI sheet not found."
        );

        return [];

    }


    const result = [];


    /*
       Read actual worksheet range.

       We intentionally inspect rows 5 onward.
       Row 4 contains headers.
    */

    const range =
        XLSX.utils.decode_range(
            sheet["!ref"]
        );


    for (
        let r = 4;
        r <= range.e.r;
        r++
    ) {

        const excelRow =
            r + 1;


        const dateCell =
            sheet[
                `B${excelRow}`
            ];


        const date =
            parseDate(
                dateCell?.v
            );


        if (!date) {
            continue;
        }


        const hours =
            parseNumber(
                sheet[
                    `I${excelRow}`
                ]?.v
            );


        const paRaw =
            parseNumber(
                sheet[
                    `S${excelRow}`
                ]?.v
            );


        const prRaw =
            parseNumber(
                sheet[
                    `V${excelRow}`
                ]?.v
            );


        const lossRaw =
            parseNumber(
                sheet[
                    `AD${excelRow}`
                ]?.v
            );


        /*
           Convert decimals to actual percentages.

           PA:
           1       -> 100%
           .9965   -> 99.65%

           PR:
           .7862   -> 78.62%

           System Loss:
           .0193   -> 1.93%
        */

        const pa =
            paRaw === null
                ? null
                : (
                    Math.abs(
                        paRaw
                    ) <= 1.5
                        ? paRaw * 100
                        : paRaw
                );


        const pr =
            prRaw === null
                ? null
                : (
                    Math.abs(
                        prRaw
                    ) <= 1.5
                        ? prRaw * 100
                        : prRaw
                );


        const loss =
            lossRaw === null
                ? null
                : (
                    Math.abs(
                        lossRaw
                    ) <= 1.5
                        ? lossRaw * 100
                        : lossRaw
                );


        result.push({

            date,

            pa,

            pr,

            hours,

            loss

        });

    }


    result.sort(
        (a, b) =>
            a.date -
            b.date
    );


    console.log(
        "Daily KPI records:",
        result
    );


    return result;

}


/* =========================================================
   DAILY LABELS
========================================================= */

function makeDailyLabels(
    rows
) {

    return rows.map(
        row => {

            const day =
                row.date.getDate();


            /*
               For a month-sized dataset:
               2 4 6 8 10...

               The data remains daily;
               only the displayed labels are
               reduced.
            */

            if (
                rows.length >= 20 &&
                rows.length <= 35
            ) {

                return day % 2 === 0
                    ? String(
                        day
                    )
                    : "";

            }


            return shortDate(
                row.date
            );

        }
    );

}


/* =========================================================
   DAILY CHARTS
========================================================= */

function renderDailyKPICharts(
    rows
) {

    if (!rows.length) {
        return;
    }


    const labels =
        makeDailyLabels(
            rows
        );


    /*
       PA
    */

    renderPAPercentage(
        rows
    );


    /*
       PR
    */

    createDailyLineChart(
        "prChart",
        labels,
        rows.map(
            row =>
                row.pr
        ),
        "Performance Ratio",
        "PR (%)"
    );


    /*
       Operating Hours
    */

    createDailyLineChart(
        "hoursChart",
        labels,
        rows.map(
            row =>
                row.hours
        ),
        "Operating Hours",
        "Operating Hours",
        true
    );


    /*
       System Loss
    */

    createDailyLineChart(
        "lossChart",
        labels,
        rows.map(
            row =>
                row.loss
        ),
        "System Losses",
        "System Loss (%)",
        true
    );


    /*
       Dashboard PR
    */

    createDailyLineChart(
        "dashboardPRChart",
        labels,
        rows.map(
            row =>
                row.pr
        ),
        "Performance Ratio",
        "PR (%)"
    );


    /*
       Dashboard System Loss
    */

    createDailyLineChart(
        "dashboardLossChart",
        labels,
        rows.map(
            row =>
                row.loss
        ),
        "System Losses",
        "System Loss (%)",
        true
    );


    /*
       KPI cards show latest DGR day.
    */

    const latest =
        rows[
            rows.length - 1
        ];


    setText(
        "dashboardPA",
        latest.pa === null
            ? "—"
            : `${latest.pa.toFixed(2)}%`
    );


    setText(
        "dashboardPR",
        latest.pr === null
            ? "—"
            : `${latest.pr.toFixed(2)}%`
    );


    setText(
        "dashboardLoss",
        latest.loss === null
            ? "—"
            : `${latest.loss.toFixed(2)}%`
    );


    setText(
        "dashboardHours",
        latest.hours === null
            ? "—"
            : `${latest.hours.toFixed(2)} h`
    );

}


/* =========================================================
   DAILY LINE CHART
========================================================= */

function createDailyLineChart(
    canvasId,
    labels,
    values,
    label,
    yTitle,
    beginAtZero = false
) {

    const canvas =
        $(canvasId);


    if (!canvas) {
        return;
    }


    destroyChart(
        canvasId
    );


    /*
       Give every daily chart enough
       internal width to keep dates readable.
    */

    const parent =
        canvas.parentElement;


    const width =
        Math.max(
            parent?.clientWidth || 700,
            labels.length * 80
        );


    makeChartScrollable(
        canvas,
        width
    );


    charts[canvasId] =
        new Chart(
            canvas.getContext(
                "2d"
            ),
            {

                type:
                    "line",

                data: {

                    labels,

                    datasets: [

                        {

                            label,

                            data:
                                values,

                            borderWidth:
                                2,

                            pointRadius:
                                3,

                            pointHoverRadius:
                                6,

                            tension:
                                0.20,

                            fill:
                                false

                        }

                    ]

                },


                options: {

                    responsive:
                        false,

                    maintainAspectRatio:
                        false,

                    animation:
                        false,

                    interaction: {

                        mode:
                            "index",

                        intersect:
                            false

                    },


                    plugins: {

                        legend: {

                            display:
                                false

                        }

                    },


                    scales: {

                        x: {

                            grid: {

                                display:
                                    false

                            },

                            ticks: {

                                autoSkip:
                                    false,

                                maxRotation:
                                    0,

                                minRotation:
                                    0,

                                font: {

                                    size:
                                        9

                                }

                            }

                        },


                        y: {

                            beginAtZero,

                            title: {

                                display:
                                    true,

                                text:
                                    yTitle

                            },

                            ticks: {

                                maxTicksLimit:
                                    7

                            }

                        }

                    }

                }

            }
        );

}


/* =========================================================
   PA PERCENTAGE CARD
========================================================= */

function ensurePAPercentageCard() {

    const section =
        $("paSection");


    if (!section) {
        return null;
    }


    let card =
        $("paPercentageCard");


    if (!card) {

        card =
            document.createElement(
                "div"
            );


        card.id =
            "paPercentageCard";


        card.className =
            "chart-card full-card";


        card.style.marginTop =
            "14px";


        card.innerHTML = `

            <div class="chart-heading">

                <div>

                    <h3>
                        Plant Availability
                    </h3>

                    <span>
                        Daily_KPI · Column S
                    </span>

                </div>

                <span class="chart-type">
                    PA %
                </span>

            </div>

            <div class="chart-container">

                <canvas id="paPercentageChart"></canvas>

            </div>

        `;


        const ganttCard =
            section.querySelector(
                ".chart-card.full-card"
            );


        if (ganttCard) {

            section.insertBefore(
                card,
                ganttCard
            );

        }

        else {

            section.appendChild(
                card
            );

        }

    }


    return card;

}


/* =========================================================
   PA PERCENTAGE GRAPH
========================================================= */

function renderPAPercentage(
    rows
) {

    const card =
        ensurePAPercentageCard();


    if (!card) {
        return;
    }


    const canvas =
        $("paPercentageChart");


    if (!canvas) {
        return;
    }


    destroyChart(
        "paPercentageChart"
    );


    const labels =
        makeDailyLabels(
            rows
        );


    const width =
        Math.max(
            card.clientWidth || 700,
            rows.length * 80
        );


    makeChartScrollable(
        canvas,
        width
    );


    charts.paPercentageChart =
        new Chart(
            canvas.getContext(
                "2d"
            ),
            {

                type:
                    "line",

                data: {

                    labels,

                    datasets: [

                        {

                            label:
                                "PA (%)",

                            data:
                                rows.map(
                                    row =>
                                        row.pa
                                ),

                            borderWidth:
                                2,

                            pointRadius:
                                3,

                            pointHoverRadius:
                                6,

                            tension:
                                0.20,

                            fill:
                                false

                        }

                    ]

                },


                options: {

                    responsive:
                        false,

                    maintainAspectRatio:
                        false,

                    animation:
                        false,


                    plugins: {

                        legend: {

                            display:
                                false

                        }

                    },


                    scales: {

                        x: {

                            grid: {

                                display:
                                    false

                            },

                            ticks: {

                                autoSkip:
                                    false,

                                maxRotation:
                                    0,

                                minRotation:
                                    0

                            }

                        },


                        y: {

                            min:
                                80,

                            max:
                                100,

                            title: {

                                display:
                                    true,

                                text:
                                    "Plant Availability (%)"

                            },

                            ticks: {

                                stepSize:
                                    5,

                                callback:
                                    value =>
                                        `${value}%`

                            }

                        }

                    }

                }

            }
        );

}


/* =========================================================
   PLANT UNAVAILABILITY
========================================================= */

function readPlantUnavailability() {

    const sheet =
        getSheet(
            "PA"
        );


    if (!sheet) {
        return [];
    }


    const range =
        XLSX.utils.decode_range(
            sheet["!ref"]
        );


    const records = [];


    for (
        let r = 4;
        r <= range.e.r;
        r++
    ) {

        const excelRow =
            r + 1;


        const issue =
            sheet[
                `W${excelRow}`
            ]?.v;


        const start =
            parseTimeMinutes(
                sheet[
                    `Z${excelRow}`
                ]?.v
            );


        const end =
            parseTimeMinutes(
                sheet[
                    `AC${excelRow}`
                ]?.v
            );


        if (
            issue === null ||
            issue === undefined ||
            String(
                issue
            ).trim() === ""
        ) {

            continue;

        }


        if (
            start === null ||
            end === null
        ) {

            continue;

        }


        let finish =
            end;


        if (
            finish < start
        ) {

            finish +=
                1440;

        }


        records.push({

            issue:
                String(
                    issue
                ).trim(),

            start,

            end:
                finish

        });

    }


    return records;

}


/* =========================================================
   PLANT UNAVAILABILITY GANTT
========================================================= */

function renderPlantUnavailability() {

    const canvas =
        $("paChart");


    if (!canvas) {
        return;
    }


    destroyChart(
        "paChart"
    );


    const records =
        readPlantUnavailability();


    if (!records.length) {

        showCanvasMessage(
            canvas,
            "No Plant Unavailability records found."
        );

        return;

    }


    const parent =
        canvas.parentElement;


    const width =
        Math.max(
            parent?.clientWidth || 800,
            2300
        );


    makeChartScrollable(
        canvas,
        width
    );


    const labels =
        records.map(
            record =>
                record.issue
        );


    const datasets =
        records.map(
            record => ({

                label:
                    record.issue,

                data: [

                    {

                        x: [

                            record.start,

                            record.end

                        ],

                        y:
                            record.issue

                    }

                ],

                backgroundColor:
                    "rgba(39,165,173,0.72)",

                borderColor:
                    "#27A5AD",

                borderWidth:
                    1,

                borderRadius:
                    4,

                barThickness:
                    22

            })
        );


    charts.paChart =
        new Chart(
            canvas.getContext(
                "2d"
            ),
            {

                type:
                    "bar",

                data: {

                    labels,

                    datasets

                },


                options: {

                    indexAxis:
                        "y",

                    responsive:
                        false,

                    maintainAspectRatio:
                        false,

                    animation:
                        false,

                    parsing:
                        false,


                    plugins: {

                        legend: {

                            display:
                                false

                        },


                        tooltip: {

                            callbacks: {

                                title:
                                    context =>
                                        records[
                                            context[0]
                                                .dataIndex
                                        ]?.issue ||
                                        "",


                                label:
                                    context => {

                                        const item =
                                            records[
                                                context
                                                    .dataIndex
                                            ];


                                        return [

                                            `Start: ${minutesToTime(item.start)}`,

                                            `End: ${minutesToTime(item.end)}`,

                                            `Duration: ${item.end - item.start} min`

                                        ];

                                    }

                            }

                        }

                    },


                    scales: {

                        x: {

                            type:
                                "linear",

                            min:
                                0,

                            max:
                                1439,

                            title: {

                                display:
                                    true,

                                text:
                                    "Time"

                            },

                            ticks: {

                                stepSize:
                                    30,

                                callback:
                                    value =>
                                        minutesToTime(
                                            value
                                        )

                            }

                        },


                        y: {

                            type:
                                "category",

                            labels,

                            title: {

                                display:
                                    true,

                                text:
                                    "Issue / Fault"

                            },

                            grid: {

                                display:
                                    false

                            }

                        }

                    }

                }

            }
        );

}


/* =========================================================
   BREAKDOWN
========================================================= */

function readBreakdownTimeline() {

    const sheet =
        getSheet(
            "PA"
        );


    if (!sheet) {
        return [];
    }


    const range =
        XLSX.utils.decode_range(
            sheet["!ref"]
        );


    const map =
        new Map();


    for (
        let r = 4;
        r <= range.e.r;
        r++
    ) {

        const excelRow =
            r + 1;


        const date =
            parseDate(
                sheet[
                    `B${excelRow}`
                ]?.v
            );


        const minutes =
            parseNumber(
                sheet[
                    `AG${excelRow}`
                ]?.v
            );


        if (
            !date ||
            minutes === null
        ) {

            continue;

        }


        const key =
            dateKey(
                date
            );


        if (
            !map.has(
                key
            )
        ) {

            map.set(
                key,
                {

                    date:
                        new Date(
                            date.getFullYear(),
                            date.getMonth(),
                            date.getDate()
                        ),

                    minutes:
                        0

                }
            );

        }


        map.get(
            key
        ).minutes +=
            minutes;

    }


    return Array.from(
        map.values()
    )
        .sort(
            (a, b) =>
                a.date -
                b.date
        );

}


/* =========================================================
   DATE KEY
========================================================= */

function dateKey(
    date
) {

    return [
        date.getFullYear(),

        String(
            date.getMonth() + 1
        ).padStart(
            2,
            "0"
        ),

        String(
            date.getDate()
        ).padStart(
            2,
            "0"
        )

    ].join(
        "-"
    );

}


/* =========================================================
   BREAKDOWN CARD
========================================================= */

function ensureBreakdownCard() {

    const section =
        $("paSection");


    if (!section) {
        return null;
    }


    let card =
        $("breakdownTimelineCard");


    if (!card) {

        card =
            document.createElement(
                "div"
            );


        card.id =
            "breakdownTimelineCard";


        card.className =
            "chart-card full-card";


        card.style.marginTop =
            "14px";


        card.innerHTML = `

            <div class="chart-heading">

                <div>

                    <h3>
                        Breakdown Timeline
                    </h3>

                    <span>
                        Daily breakdown time · PA Column AG
                    </span>

                </div>

                <span class="chart-type">
                    MINUTES
                </span>

            </div>

            <div class="chart-container">

                <canvas id="breakdownChart"></canvas>

            </div>

        `;


        const gantt =
            section.querySelector(
                ".chart-card.full-card"
            );


        if (gantt) {

            gantt.insertAdjacentElement(
                "afterend",
                card
            );

        }

        else {

            section.appendChild(
                card
            );

        }

    }


    return card;

}


/* =========================================================
   BREAKDOWN CHART
========================================================= */

function renderBreakdownTimeline() {

    const card =
        ensureBreakdownCard();


    if (!card) {
        return;
    }


    const canvas =
        $("breakdownChart");


    if (!canvas) {
        return;
    }


    destroyChart(
        "breakdownChart"
    );


    const records =
        readBreakdownTimeline();


    if (!records.length) {

        showCanvasMessage(
            canvas,
            "No breakdown records found."
        );

        return;

    }


    charts.breakdownChart =
        new Chart(
            canvas.getContext(
                "2d"
            ),
            {

                type:
                    "bar",

                data: {

                    labels:
                        records.map(
                            record =>
                                shortDate(
                                    record.date
                                )
                        ),

                    datasets: [

                        {

                            label:
                                "Breakdown Time",

                            data:
                                records.map(
                                    record =>
                                        record.minutes
                                ),

                            borderWidth:
                                1,

                            borderRadius:
                                3

                        }

                    ]

                },


                options: {

                    indexAxis:
                        "y",

                    responsive:
                        true,

                    maintainAspectRatio:
                        false,

                    animation:
                        false,


                    plugins: {

                        legend: {

                            display:
                                false

                        }

                    },


                    scales: {

                        x: {

                            min:
                                0,

                            max:
                                13,

                            ticks: {

                                stepSize:
                                    1

                            },

                            title: {

                                display:
                                    true,

                                text:
                                    "Breakdown Time (minutes)"

                            }

                        },


                        y: {

                            grid: {

                                display:
                                    false

                            }

                        }

                    }

                }

            }
        );

}


/* =========================================================
   SYSTEM LOSS MWh
========================================================= */

function readSystemLossMWh() {

    const sheet =
        getSheet(
            "PA"
        );


    if (!sheet) {
        return [];
    }


    const range =
        XLSX.utils.decode_range(
            sheet["!ref"]
        );


    const map =
        new Map();


    for (
        let r = 4;
        r <= range.e.r;
        r++
    ) {

        const excelRow =
            r + 1;


        const date =
            parseDate(
                sheet[
                    `B${excelRow}`
                ]?.v
            );


        const loss =
            parseNumber(
                sheet[
                    `AL${excelRow}`
                ]?.v
            );


        if (
            !date ||
            loss === null
        ) {

            continue;

        }


        const key =
            dateKey(
                date
            );


        if (
            !map.has(
                key
            )
        ) {

            map.set(
                key,
                {

                    date:
                        new Date(
                            date.getFullYear(),
                            date.getMonth(),
                            date.getDate()
                        ),

                    loss:
                        0

                }
            );

        }


        map.get(
            key
        ).loss +=
            loss;

    }


    return Array.from(
        map.values()
    )
        .sort(
            (a, b) =>
                a.date -
                b.date
        );

}


/* =========================================================
   SYSTEM LOSS MWh CARD
========================================================= */

function ensureSystemLossCard() {

    const section =
        $("paSection");


    if (!section) {
        return null;
    }


    let card =
        $("systemLossMwhCard");


    if (!card) {

        card =
            document.createElement(
                "div"
            );


        card.id =
            "systemLossMwhCard";


        card.className =
            "chart-card full-card";


        card.style.marginTop =
            "14px";


        card.innerHTML = `

            <div class="chart-heading">

                <div>

                    <h3>
                        System Loss
                    </h3>

                    <span>
                        Same-date system loss combined · PA Column AL
                    </span>

                </div>

                <span class="chart-type">
                    MWh
                </span>

            </div>

            <div class="large-chart-container">

                <canvas id="systemLossMwhChart"></canvas>

            </div>

        `;


        const breakdown =
            $("breakdownTimelineCard");


        if (breakdown) {

            breakdown.insertAdjacentElement(
                "afterend",
                card
            );

        }

        else {

            section.appendChild(
                card
            );

        }

    }


    return card;

}


/* =========================================================
   SYSTEM LOSS MWh CHART
========================================================= */

function renderSystemLossMWh() {

    const card =
        ensureSystemLossCard();


    if (!card) {
        return;
    }


    const canvas =
        $("systemLossMwhChart");


    if (!canvas) {
        return;
    }


    destroyChart(
        "systemLossMwhChart"
    );


    const records =
        readSystemLossMWh();


    if (!records.length) {

        showCanvasMessage(
            canvas,
            "No system loss MWh records found."
        );

        return;

    }


    const width =
        Math.max(
            card.clientWidth || 700,
            records.length * 80
        );


    makeChartScrollable(
        canvas,
        width
    );


    charts.systemLossMwhChart =
        new Chart(
            canvas.getContext(
                "2d"
            ),
            {

                type:
                    "bar",

                data: {

                    labels:
                        records.map(
                            record =>
                                shortDate(
                                    record.date
                                )
                        ),

                    datasets: [

                        {

                            label:
                                "System Loss (MWh)",

                            data:
                                records.map(
                                    record =>
                                        record.loss
                                ),

                            borderWidth:
                                1,

                            borderRadius:
                                4

                        }

                    ]

                },


                options: {

                    responsive:
                        false,

                    maintainAspectRatio:
                        false,

                    animation:
                        false,


                    plugins: {

                        legend: {

                            display:
                                false

                        }

                    },


                    scales: {

                        x: {

                            grid: {

                                display:
                                    false

                            }

                        },


                        y: {

                            beginAtZero:
                                true,

                            title: {

                                display:
                                    true,

                                text:
                                    "System Loss (MWh)"

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

function readCurtailment() {

    const sheet =
        getSheet(
            "Curtailment records"
        );


    if (!sheet) {
        return [];
    }


    const range =
        XLSX.utils.decode_range(
            sheet["!ref"]
        );


    const records = [];


    for (
        let r = 1;
        r <= range.e.r;
        r++
    ) {

        const excelRow =
            r + 1;


        const date =
            parseDate(
                sheet[
                    `C${excelRow}`
                ]?.v
            );


        const start =
            parseTimeMinutes(
                sheet[
                    `H${excelRow}`
                ]?.v
            );


        const end =
            parseTimeMinutes(
                sheet[
                    `I${excelRow}`
                ]?.v
            );


        const loss =
            parseNumber(
                sheet[
                    `R${excelRow}`
                ]?.v
            );


        if (
            !date ||
            start === null ||
            end === null
        ) {

            continue;

        }


        let finish =
            end;


        if (
            finish < start
        ) {

            finish +=
                1440;

        }


        records.push({

            date,

            key:
                dateKey(
                    date
                ),

            start,

            end:
                finish,

            loss:
                loss || 0

        });

    }


    return records;

}


/* =========================================================
   CURTAILMENT TABLE
========================================================= */

function ensureCurtailmentTable() {

    const section =
        $("curtailmentSection");


    if (!section) {
        return null;
    }


    let card =
        $("curtailmentTableCard");


    if (!card) {

        card =
            document.createElement(
                "div"
            );


        card.id =
            "curtailmentTableCard";


        card.className =
            "chart-card full-card";


        card.style.marginTop =
            "14px";


        card.innerHTML = `

            <div class="chart-heading">

                <div>

                    <h3>
                        Daily Curtailment Loss
                    </h3>

                    <span>
                        Loss of Generation merged by date · Column R
                    </span>

                </div>

                <span class="chart-type">
                    TABLE
                </span>

            </div>

            <div id="curtailmentTable"></div>

        `;


        const main =
            section.querySelector(
                ".chart-card.full-card"
            );


        if (main) {

            main.insertAdjacentElement(
                "afterend",
                card
            );

        }

        else {

            section.appendChild(
                card
            );

        }

    }


    return $("curtailmentTable");

}


/* =========================================================
   CURTAILMENT TABLE
========================================================= */

function renderCurtailmentTable(
    records
) {

    const target =
        ensureCurtailmentTable();


    if (!target) {
        return;
    }


    const map =
        new Map();


    records.forEach(
        record => {

            if (
                !map.has(
                    record.key
                )
            ) {

                map.set(
                    record.key,
                    {

                        date:
                            record.date,

                        loss:
                            0,

                        intervals:
                            0

                    }
                );

            }


            map.get(
                record.key
            ).loss +=
                record.loss;


            map.get(
                record.key
            ).intervals++;

        }
    );


    const daily =
        Array.from(
            map.values()
        )
        .sort(
            (a, b) =>
                a.date -
                b.date
        );


    if (!daily.length) {

        target.innerHTML = `

            <div style="
                padding:16px;
                color:#879397;
                font-size:10px;
            ">
                No curtailment records found.
            </div>

        `;

        return;

    }


    let html = `

        <table style="
            width:100%;
            border-collapse:collapse;
            font-size:10px;
        ">

            <thead>

                <tr>

                    <th style="
                        text-align:left;
                        padding:10px;
                        border-bottom:1px solid #e1ebed;
                    ">
                        Date
                    </th>

                    <th style="
                        text-align:right;
                        padding:10px;
                        border-bottom:1px solid #e1ebed;
                    ">
                        Loss of Generation (MWh)
                    </th>

                    <th style="
                        text-align:right;
                        padding:10px;
                        border-bottom:1px solid #e1ebed;
                    ">
                        Intervals
                    </th>

                </tr>

            </thead>

            <tbody>
    `;


    daily.forEach(
        item => {

            html += `

                <tr>

                    <td style="
                        padding:10px;
                        border-bottom:1px solid #edf2f3;
                    ">
                        ${fullDate(
                            item.date
                        )}
                    </td>

                    <td style="
                        padding:10px;
                        text-align:right;
                        border-bottom:1px solid #edf2f3;
                    ">
                        ${item.loss.toFixed(2)}
                    </td>

                    <td style="
                        padding:10px;
                        text-align:right;
                        border-bottom:1px solid #edf2f3;
                    ">
                        ${item.intervals}
                    </td>

                </tr>

            `;

        }
    );


    html += `

            </tbody>

        </table>

    `;


    target.innerHTML =
        html;


    const total =
        daily.reduce(
            (sum, item) =>
                sum +
                item.loss,
            0
        );


    setText(
        "curtailmentSummary",
        `${daily.length} date(s) · ${total.toFixed(2)} MWh total generation loss`
    );

}


/* =========================================================
   CURTAILMENT GANTT CARD
========================================================= */

function ensureCurtailmentGantt() {

    const section =
        $("curtailmentSection");


    if (!section) {
        return null;
    }


    let card =
        $("curtailmentGanttCard");


    if (!card) {

        card =
            document.createElement(
                "div"
            );


        card.id =
            "curtailmentGanttCard";


        card.className =
            "chart-card full-card";


        card.style.marginTop =
            "14px";


        card.innerHTML = `

            <div class="chart-heading">

                <div>

                    <h3>
                        Curtailment Duration
                    </h3>

                    <span>
                        06:00–18:00 · 15-minute intervals
                    </span>

                </div>

                <span class="chart-type">
                    GANTT
                </span>

            </div>

            <div class="timeline-wrapper">

                <canvas id="curtailmentGanttChart"></canvas>

            </div>

        `;


        const table =
            $("curtailmentTableCard");


        if (table) {

            table.insertAdjacentElement(
                "afterend",
                card
            );

        }

        else {

            section.appendChild(
                card
            );

        }

    }


    return card;

}


/* =========================================================
   CURTAILMENT GANTT
========================================================= */

function renderCurtailmentGantt(
    records
) {

    const card =
        ensureCurtailmentGantt();


    if (!card) {
        return;
    }


    const canvas =
        $("curtailmentGanttChart");


    if (!canvas) {
        return;
    }


    destroyChart(
        "curtailmentGanttChart"
    );


    if (!records.length) {

        showCanvasMessage(
            canvas,
            "No curtailment intervals found."
        );

        return;

    }


    const width =
        Math.max(
            card.clientWidth || 700,
            3120
        );


    makeChartScrollable(
        canvas,
        width
    );


    const labels = [];


    records.forEach(
        record => {

            const label =
                shortDate(
                    record.date
                );


            if (
                !labels.includes(
                    label
                )
            ) {

                labels.push(
                    label
                );

            }

        }
    );


    const datasets =
        records
            .map(
                record => {

                    const start =
                        Math.max(
                            360,
                            record.start
                        );


                    const end =
                        Math.min(
                            1080,
                            record.end
                        );


                    if (
                        end <= start
                    ) {

                        return null;

                    }


                    return {

                        label:
                            `${shortDate(record.date)} ${minutesToTime(start)}–${minutesToTime(end)}`,

                        data: [

                            {

                                x: [
                                    start,
                                    end
                                ],

                                y:
                                    shortDate(
                                        record.date
                                    )

                            }

                        ],

                        backgroundColor:
                            "rgba(39,165,173,0.72)",

                        borderColor:
                            "#27A5AD",

                        borderWidth:
                            1,

                        borderRadius:
                            4,

                        barThickness:
                            20

                    };

                }
            )
            .filter(
                Boolean
            );


    charts.curtailmentGanttChart =
        new Chart(
            canvas.getContext(
                "2d"
            ),
            {

                type:
                    "bar",

                data: {

                    labels,

                    datasets

                },


                options: {

                    indexAxis:
                        "y",

                    responsive:
                        false,

                    maintainAspectRatio:
                        false,

                    animation:
                        false,

                    parsing:
                        false,


                    plugins: {

                        legend: {

                            display:
                                false

                        }

                    },


                    scales: {

                        x: {

                            type:
                                "linear",

                            min:
                                360,

                            max:
                                1080,

                            title: {

                                display:
                                    true,

                                text:
                                    "Time"

                            },

                            ticks: {

                                stepSize:
                                    15,

                                callback:
                                    value =>
                                        minutesToTime(
                                            value
                                        )

                            }

                        },


                        y: {

                            type:
                                "category",

                            labels,

                            title: {

                                display:
                                    true,

                                text:
                                    "Date"

                            },

                            grid: {

                                display:
                                    false

                            }

                        }

                    }

                }

            }
        );

}


/* =========================================================
   MONTHLY PR
   Annual_KPI H9:I21

   H = Target PR
   I = Measured PR

   We graph MEASURED PR.
========================================================= */

function readMonthlyPR() {

    const sheet =
        getSheet(
            "Annual_KPI"
        );


    if (!sheet) {

        console.warn(
            "Annual_KPI sheet not found."
        );

        return [];

    }


    const months = [
        "April",
        "May",
        "June",
        "July",
        "August",
        "September",
        "October",
        "November",
        "December",
        "January",
        "February",
        "March"
    ];


    const rows = [];


    /*
       Row 9 = header.
       Row 10 = April.
       Row 11 = May.
       Row 12 = June.
       etc.
    */

    for (
        let row = 10;
        row <= 21;
        row++
    ) {

        const month =
            sheet[
                `A${row}`
            ]?.v;


        const targetRaw =
            sheet[
                `H${row}`
            ]?.v;


        const measuredRaw =
            sheet[
                `I${row}`
            ]?.v;


        const date =
            parseDate(
                month
            );


        const target =
            parseNumber(
                targetRaw
            );


        const measured =
            parseNumber(
                measuredRaw
            );


        if (
            target === null &&
            measured === null
        ) {

            continue;

        }


        /*
           Use worksheet month/date when valid,
           otherwise use the month sequence.
        */

        let monthName =
            months[
                rows.length
            ];


        if (
            date
        ) {

            monthName =
                date.toLocaleDateString(
                    "en-IN",
                    {
                        month:
                            "long"
                    }
                );

        }


        rows.push({

            month:
                monthName,

            target:

                target === null
                    ? null
                    : (
                        Math.abs(
                            target
                        ) <= 1.5
                            ? target * 100
                            : target
                    ),

            measured:

                measured === null
                    ? null
                    : (
                        Math.abs(
                            measured
                        ) <= 1.5
                            ? measured * 100
                            : measured
                    )

        });

    }


    return rows;

}


/* =========================================================
   MONTHLY PR CHART
   Horizontal bar:
   Y = April, May, June...
   X = PR (%)
========================================================= */

function renderMonthlyPRChart() {

    const records =
        readMonthlyPR();


    if (!records.length) {
        return;
    }


    /*
       We don't currently have a dedicated
       PR canvas in the HTML.

       Create one inside Performance section.
    */

    const section =
        $("performanceSection");


    if (!section) {
        return;
    }


    let card =
        $("monthlyPRCard");


    if (!card) {

        card =
            document.createElement(
                "div"
            );


        card.id =
            "monthlyPRCard";


        card.className =
            "chart-card full-card";


        card.style.marginTop =
            "14px";


        card.innerHTML = `

            <div class="chart-heading">

                <div>

                    <h3>
                        Monthly Performance Ratio
                    </h3>

                    <span>
                        Annual_KPI · Measured PR(%) · I column
                    </span>

                </div>

                <span class="chart-type">
                    MONTHLY PR
                </span>

            </div>

            <div class="chart-container">

                <canvas id="monthlyPRChart"></canvas>

            </div>

        `;


        section
            .querySelector(
                ".three-chart-grid"
            )
            ?.insertAdjacentElement(
                "afterend",
                card
            );

    }


    const canvas =
        $("monthlyPRChart");


    if (!canvas) {
        return;
    }


    destroyChart(
        "monthlyPRChart"
    );


    /*
       Horizontal bar chart.
       Y axis = months.
       X axis = percentages.
    */

    charts.monthlyPRChart =
        new Chart(
            canvas.getContext(
                "2d"
            ),
            {

                type:
                    "bar",

                data: {

                    labels:
                        records.map(
                            record =>
                                record.month
                        ),

                    datasets: [

                        {

                            label:
                                "Measured PR (%)",

                            data:
                                records.map(
                                    record =>
                                        record.measured
                                ),

                            borderWidth:
                                1,

                            borderRadius:
                                4

                        }

                    ]

                },


                options: {

                    indexAxis:
                        "y",

                    responsive:
                        true,

                    maintainAspectRatio:
                        false,

                    animation:
                        false,


                    plugins: {

                        legend: {

                            display:
                                false

                        }

                    },


                    scales: {

                        x: {

                            min:
                                0,

                            max:
                                100,

                            title: {

                                display:
                                    true,

                                text:
                                    "Performance Ratio (%)"

                            },

                            ticks: {

                                stepSize:
                                    10,

                                callback:
                                    value =>
                                        `${value}%`

                            }

                        },


                        y: {

                            title: {

                                display:
                                    true,

                                text:
                                    "Month"

                            },

                            grid: {

                                display:
                                    false

                            }

                        }

                    }

                }

            }
        );

}


/* =========================================================
   ENERGY
========================================================= */

function renderEnergyChart() {

    const sheet =
        getSheet(
            "Annual_KPI"
        );


    if (!sheet) {
        return;
    }


    const budget = [];
    const measured = [];


    const labels = [
        "April",
        "May",
        "June",
        "July",
        "August",
        "September",
        "October",
        "November",
        "December",
        "January",
        "February",
        "March"
    ];


    for (
        let row = 10;
        row <= 21;
        row++
    ) {

        budget.push(
            parseNumber(
                sheet[
                    `E${row}`
                ]?.v
            )
        );


        measured.push(
            parseNumber(
                sheet[
                    `F${row}`
                ]?.v
            )
        );

    }


    const totalBudget =
        budget.reduce(
            (sum, value) =>
                sum +
                (
                    value || 0
                ),
            0
        );


    const totalMeasured =
        measured.reduce(
            (sum, value) =>
                sum +
                (
                    value || 0
                ),
            0
        );


    const variance =
        totalMeasured -
        totalBudget;


    setText(
        "totalBudget",
        `${formatNumber(
            totalBudget
        )} MWh`
    );


    setText(
        "totalMeasured",
        `${formatNumber(
            totalMeasured
        )} MWh`
    );


    setText(
        "energyVariance",
        `${formatNumber(
            variance
        )} MWh`
    );


    const canvas =
        $("energyChart");


    if (!canvas) {
        return;
    }


    destroyChart(
        "energyChart"
    );


    charts.energyChart =
        new Chart(
            canvas.getContext(
                "2d"
            ),
            {

                type:
                    "bar",

                data: {

                    labels,

                    datasets: [

                        {

                            label:
                                "Budgeted Energy",

                            data:
                                budget,

                            borderWidth:
                                1,

                            borderRadius:
                                4

                        },

                        {

                            label:
                                "Measured Energy",

                            data:
                                measured,

                            borderWidth:
                                1,

                            borderRadius:
                                4

                        }

                    ]

                },


                options: {

                    responsive:
                        true,

                    maintainAspectRatio:
                        false,

                    animation:
                        false,


                    plugins: {

                        legend: {

                            display:
                                true,

                            position:
                                "top"

                        }

                    },


                    scales: {

                        x: {

                            grid: {

                                display:
                                    false

                            }

                        },


                        y: {

                            beginAtZero:
                                true,

                            title: {

                                display:
                                    true,

                                text:
                                    "Energy"

                            }

                        }

                    }

                }

            }
        );

}


/* =========================================================
   FULL RENDER
========================================================= */

function renderAll() {

    destroyAllCharts();

    /*
       Remove dynamic cards from previous upload.
    */

    [
        "paPercentageCard",
        "breakdownTimelineCard",
        "systemLossMwhCard",
        "monthlyPRCard",
        "curtailmentTableCard",
        "curtailmentGanttCard"
    ]
        .forEach(
            id => {

                const element =
                    $(id);


                if (element) {

                    element.remove();

                }

            }
        );


    /*
       DAILY KPI
    */

    const daily =
        readDailyKPI();


    if (
        daily.length
    ) {

        renderDailyKPICharts(
            daily
        );

    }


    /*
       PA
    */

    renderPlantUnavailability();

    renderBreakdownTimeline();

    renderSystemLossMWh();


    /*
       CURTAILMENT
    */

    const curtailment =
        readCurtailment();


    renderCurtailmentTable(
        curtailment
    );


    renderCurtailmentGantt(
        curtailment
    );


    /*
       MONTHLY PR
    */

    renderMonthlyPRChart();


    /*
       ENERGY
    */

    renderEnergyChart();

}


/* =========================================================
   RESET
========================================================= */

function resetDashboard() {

    workbook =
        null;


    destroyAllCharts();


    [
        "paPercentageCard",
        "breakdownTimelineCard",
        "systemLossMwhCard",
        "monthlyPRCard",
        "curtailmentTableCard",
        "curtailmentGanttCard"
    ]
        .forEach(
            id => {

                const element =
                    $(id);


                if (element) {

                    element.remove();

                }

            }
        );


    if (fileInput) {

        fileInput.value =
            "";

    }


    $("fileInfo")
        ?.classList
        .add(
            "hidden"
        );


    $("workbookStatus")
        ?.classList
        .add(
            "hidden"
        );


    $("dropZone")
        ?.classList
        .remove(
            "hidden"
        );


    $("emptyState")
        ?.classList
        .remove(
            "hidden"
        );


    hideAnalytics();


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


    setStatus(
        "Upload a DGR to generate the analytics."
    );


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
            id =>
                setText(
                    id,
                    "—"
                )
        );


    setText(
        "curtailmentSummary",
        "Waiting for DGR data"
    );

}


/* =========================================================
   END
========================================================= */
```
