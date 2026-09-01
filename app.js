```javascript
/* =========================================================
   SOLAR DGR ANALYTICS
   app.js
   =========================================================

   EXACT WORKBOOK MAPPING

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
   AL = System Loss

   CURTAILMENT RECORDS
   ---------------------------------------------------------
   C  = Date
   H  = Start Time
   I  = End Time
   R  = Loss of Generation MWh

   ANNUAL_KPI
   ---------------------------------------------------------
   H9  = Target PR (%)
   I9  = Measured PR (%)

   E10:E21 = Budgeted Energy
   F10:F21 = Measured Energy

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
    function () {

        installScrollStyles();

        setupNavigation();

        setupUpload();

        setupRemove();

        hideAnalytics();

    }
);


/* =========================================================
   SCROLL STYLES
========================================================= */

function installScrollStyles() {

    if (
        document.getElementById(
            "dgr-scroll-style"
        )
    ) {

        return;

    }


    const style =
        document.createElement(
            "style"
        );


    style.id =
        "dgr-scroll-style";


    style.textContent = `

        .scroll-chart-container {
            width: 100% !important;
            height: 100% !important;
            overflow-x: auto !important;
            overflow-y: hidden !important;
            position: relative !important;
            box-sizing: border-box !important;
            scrollbar-width: thin !important;
            scrollbar-color: #b9cacc transparent !important;
        }

        .scroll-chart-container::-webkit-scrollbar {
            height: 8px;
        }

        .scroll-chart-container::-webkit-scrollbar-track {
            background: #eef3f4;
            border-radius: 10px;
        }

        .scroll-chart-container::-webkit-scrollbar-thumb {
            background: #b9cacc;
            border-radius: 10px;
        }

        .scroll-chart-container::-webkit-scrollbar-thumb:hover {
            background: #27a5ad;
        }

        .scroll-chart-container > canvas {
            display: block !important;
            width: auto !important;
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

    const buttons =
        document.querySelectorAll(
            ".nav-item"
        );


    buttons.forEach(
        function (button) {

            button.addEventListener(
                "click",
                function () {

                    buttons.forEach(
                        function (item) {

                            item.classList.remove(
                                "active"
                            );

                        }
                    );


                    button.classList.add(
                        "active"
                    );


                    const target =
                        $(button.dataset.target);


                    if (target) {

                        target.scrollIntoView({
                            behavior:
                                "smooth",

                            block:
                                "start"
                        });

                    }

                }
            );

        }
    );

}


/* =========================================================
   FILE UPLOAD
========================================================= */

function setupUpload() {

    const input =
        $("dgrFile");


    if (!input) {

        console.error(
            "The #dgrFile input is missing."
        );

        return;

    }


    input.addEventListener(
        "change",
        function (event) {

            const file =
                event.target.files?.[0];


            if (file) {

                processDGR(
                    file
                );

            }

        }
    );


    const dropZone =
        $("dropZone");


    if (!dropZone) {
        return;
    }


    dropZone.addEventListener(
        "click",
        function () {

            input.click();

        }
    );


    dropZone.addEventListener(
        "dragover",
        function (event) {

            event.preventDefault();

            dropZone.classList.add(
                "dragging"
            );

        }
    );


    dropZone.addEventListener(
        "dragleave",
        function () {

            dropZone.classList.remove(
                "dragging"
            );

        }
    );


    dropZone.addEventListener(
        "drop",
        function (event) {

            event.preventDefault();

            dropZone.classList.remove(
                "dragging"
            );


            const file =
                event.dataTransfer?.files?.[0];


            if (file) {

                processDGR(
                    file
                );

            }

        }
    );

}


/* =========================================================
   REMOVE
========================================================= */

function setupRemove() {

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
   PROCESS DGR
========================================================= */

function processDGR(
    file
) {

    if (!file) {
        return;
    }


    if (
        !/\.(xlsx|xls|csv)$/i.test(
            file.name
        )
    ) {

        alert(
            "Please upload a valid Excel file."
        );

        return;

    }


    if (
        typeof XLSX === "undefined"
    ) {

        alert(
            "SheetJS is not loaded."
        );

        return;

    }


    setStatus(
        "Reading DGR workbook..."
    );


    const reader =
        new FileReader();


    reader.onload =
        function (event) {

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
                    !workbook ||
                    !workbook.SheetNames ||
                    !workbook.SheetNames.length
                ) {

                    throw new Error(
                        "No worksheets were found."
                    );

                }


                updateFileUI(
                    file
                );


                showAnalytics();


                renderDashboard();


                setStatus(
                    `${file.name} loaded successfully.`
                );

            }

            catch (error) {

                console.error(
                    "DGR loading error:",
                    error
                );


                setStatus(
                    "Unable to read the DGR."
                );


                alert(
                    "Unable to read the DGR.\n\n" +
                    error.message
                );

            }

        };


    reader.onerror =
        function () {

            setStatus(
                "Unable to read the selected file."
            );

        };


    reader.readAsArrayBuffer(
        file
    );

}


/* =========================================================
   FILE UI
========================================================= */

function updateFileUI(
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


    $("fileInfo")?.classList.remove(
        "hidden"
    );


    $("workbookStatus")?.classList.remove(
        "hidden"
    );


    $("emptyState")?.classList.add(
        "hidden"
    );


    $("dropZone")?.classList.add(
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


    const required = [

        "Dashboard",
        "Annual_KPI",
        "Daily_KPI",
        "PA",
        "Curtailment records"

    ];


    required.forEach(
        function (sheetName) {

            const badge =
                document.createElement(
                    "span"
                );


            badge.className =
                "sheet-badge";


            const exists =
                workbook.SheetNames.some(
                    function (actual) {

                        return (
                            normalizeSheet(
                                actual
                            ) ===
                            normalizeSheet(
                                sheetName
                            )
                        );

                    }
                );


            badge.textContent =
                exists
                    ? `${sheetName} ✓`
                    : `${sheetName} — missing`;


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
   SHEET NORMALISATION
========================================================= */

function normalizeSheet(
    name
) {

    return String(
        name || ""
    )
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
    requestedName
) {

    if (
        !workbook ||
        !workbook.Sheets
    ) {

        return null;

    }


    if (
        workbook.Sheets[
            requestedName
        ]
    ) {

        return workbook.Sheets[
            requestedName
        ];

    }


    const wanted =
        normalizeSheet(
            requestedName
        );


    const actual =
        workbook.SheetNames.find(
            function (name) {

                return (
                    normalizeSheet(
                        name
                    ) === wanted
                );

            }
        );


    return actual
        ? workbook.Sheets[
            actual
        ]
        : null;

}


/* =========================================================
   MATRIX
========================================================= */

function toMatrix(
    sheet
) {

    if (!sheet) {
        return [];
    }


    return XLSX.utils.sheet_to_json(
        sheet,
        {
            header:
                1,

            raw:
                true,

            defval:
                null,

            blankrows:
                false
        }
    );

}


/* =========================================================
   COLUMN INDEX
========================================================= */

function columnIndex(
    letter
) {

    let number =
        0;


    for (
        const char of letter.toUpperCase()
    ) {

        number =
            number * 26 +
            char.charCodeAt(0) -
            64;

    }


    return number - 1;

}


/* =========================================================
   GET CELL FROM ROW
========================================================= */

function getCell(
    row,
    column
) {

    if (!row) {
        return null;
    }


    return row[
        columnIndex(
            column
        )
    ];

}


/* =========================================================
   DIRECT WORKSHEET CELL
========================================================= */

function getWorksheetCell(
    sheet,
    address
) {

    if (
        !sheet ||
        !sheet[address]
    ) {

        return null;

    }


    return sheet[address];

}


/* =========================================================
   NUMBER PARSER
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
   READ NUMERIC WORKSHEET CELL
========================================================= */

function readNumericCell(
    sheet,
    address
) {

    const cell =
        getWorksheetCell(
            sheet,
            address
        );


    if (!cell) {
        return null;
    }


    if (
        typeof cell.v === "number" &&
        Number.isFinite(
            cell.v
        )
    ) {

        return cell.v;

    }


    if (
        cell.w !== undefined &&
        cell.w !== null
    ) {

        const parsed =
            parseNumber(
                cell.w
            );


        if (
            parsed !== null
        ) {

            return parsed;

        }

    }


    if (
        cell.v !== undefined
    ) {

        return parseNumber(
            cell.v
        );

    }


    return null;

}


/* =========================================================
   DATE PARSER
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

            const parsed =
                XLSX.SSF.parse_date_code(
                    value
                );


            if (
                parsed &&
                parsed.y
            ) {

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

        catch (_) {}

    }


    const text =
        String(
            value
        ).trim();


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
   DATE FORMAT
========================================================= */

function formatShortDate(
    value
) {

    const date =
        value instanceof Date
            ? value
            : parseDate(
                value
            );


    if (!date) {
        return "—";
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


function formatFullDate(
    value
) {

    const date =
        value instanceof Date
            ? value
            : parseDate(
                value
            );


    if (!date) {
        return "—";
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
   TIME PARSER
========================================================= */

function timeToMinutes(
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

            const parsed =
                XLSX.SSF.parse_date_code(
                    value
                );


            if (
                parsed
            ) {

                return (
                    parsed.H * 60 +
                    parsed.M
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


    const period =
        match[4]
            ? match[4].toUpperCase()
            : null;


    if (
        period === "PM" &&
        hours < 12
    ) {

        hours += 12;

    }


    if (
        period === "AM" &&
        hours === 12
    ) {

        hours = 0;

    }


    if (
        hours < 0 ||
        hours > 23 ||
        minutes < 0 ||
        minutes > 59
    ) {

        return null;

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
    minutes
) {

    let value =
        Number(
            minutes
        );


    if (
        !Number.isFinite(
            value
        )
    ) {

        return "—";

    }


    value =
        Math.max(
            0,
            Math.min(
                1439,
                Math.round(
                    value
                )
            )
        );


    const hours =
        Math.floor(
            value / 60
        );


    const mins =
        value % 60;


    return (
        String(
            hours
        ).padStart(
            2,
            "0"
        ) +
        ":" +
        String(
            mins
        ).padStart(
            2,
            "0"
        )
    );

}


/* =========================================================
   NUMBER FORMAT
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

        catch (_) {}

    }


    charts[id] =
        null;

}


/* =========================================================
   DESTROY ALL CHARTS
========================================================= */

function destroyAllCharts() {

    Object.keys(
        charts
    )
    .forEach(
        function (id) {

            destroyChart(
                id
            );

        }
    );

}


/* =========================================================
   SHOW / HIDE
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
        function (id) {

            const section =
                $(id);


            if (section) {

                section.style.display =
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
        function (id) {

            const section =
                $(id);


            if (section) {

                section.style.display =
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
   PREPARE SCROLLABLE CANVAS
========================================================= */

function prepareScrollableCanvas(
    canvas,
    width
) {

    if (!canvas) {
        return;
    }


    const parent =
        canvas.parentElement;


    if (!parent) {
        return;
    }


    let wrapper =
        parent.querySelector(
            ".scroll-chart-container"
        );


    if (!wrapper) {

        wrapper =
            document.createElement(
                "div"
            );


        wrapper.className =
            "scroll-chart-container";


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
            parent.clientWidth || 700,
            width || 700
        );


    canvas.style.setProperty(
        "width",
        `${actualWidth}px`,
        "important"
    );


    canvas.style.setProperty(
        "min-width",
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

}


/* =========================================================
   DAILY KPI
========================================================= */

function readDailyKPI() {

    const sheet =
        getSheet(
            "Daily_KPI"
        );


    if (!sheet || !sheet["!ref"]) {
        return [];
    }


    const range =
        XLSX.utils.decode_range(
            sheet["!ref"]
        );


    const records = [];

    let lastDate =
        null;


    /*
       IMPORTANT:

       Row 5 is the first data row.

       B contains dates.
       Some later B cells may be formula
       cells such as =B5+1.

       We therefore use the displayed
       cached result first and reconstruct
       sequential dates where necessary.
    */

    for (
        let r = 4;
        r <= range.e.r;
        r++
    ) {

        const excelRow =
            r + 1;


        let date =
            readDateCell(
                sheet,
                `B${excelRow}`
            );


        /*
           Formula-date fallback.
        */

        if (
            !date &&
            lastDate
        ) {

            date =
                new Date(
                    lastDate
                );


            date.setDate(
                date.getDate() +
                1
            );

        }


        if (!date) {
            continue;
        }


        /*
           If a formula cached a repeated
           date, use the expected sequential
           date instead.
        */

        if (
            lastDate &&
            date <= lastDate
        ) {

            const formula =
                sheet[
                    `B${excelRow}`
                ]?.f;


            if (
                formula
            ) {

                date =
                    new Date(
                        lastDate
                    );


                date.setDate(
                    date.getDate() +
                    1
                );

            }

        }


        lastDate =
            new Date(
                date
            );


        const hoursRaw =
            readNumericCell(
                sheet,
                `I${excelRow}`
            );


        const paRaw =
            readNumericCell(
                sheet,
                `S${excelRow}`
            );


        const prRaw =
            readNumericCell(
                sheet,
                `V${excelRow}`
            );


        const lossRaw =
            readNumericCell(
                sheet,
                `AD${excelRow}`
            );


        records.push({

            date,

            hours:
                hoursRaw,

            pa:
                convertPercentage(
                    paRaw
                ),

            pr:
                convertPercentage(
                    prRaw
                ),

            loss:
                convertPercentage(
                    lossRaw
                )

        });

    }


    /*
       Sort chronologically.
    */

    records.sort(
        (a, b) =>
            a.date -
            b.date
    );


    return records;

}


/* =========================================================
   READ DATE CELL
========================================================= */

function readDateCell(
    sheet,
    address
) {

    const cell =
        getWorksheetCell(
            sheet,
            address
        );


    if (!cell) {
        return null;
    }


    /*
       First try displayed value.
    */

    if (
        cell.w !== undefined &&
        cell.w !== null
    ) {

        const date =
            parseDate(
                cell.w
            );


        if (date) {
            return date;
        }

    }


    /*
       Then raw value.
    */

    if (
        cell.v !== undefined
    ) {

        return parseDate(
            cell.v
        );

    }


    return null;

}


/* =========================================================
   PERCENTAGE CONVERSION
========================================================= */

function convertPercentage(
    value
) {

    if (
        value === null ||
        value === undefined
    ) {

        return null;

    }


    const number =
        Number(
            value
        );


    if (
        !Number.isFinite(
            number
        )
    ) {

        return null;

    }


    /*
       Decimal:

       0.9965 = 99.65%
       0.8420 = 84.20%
       0.0193 = 1.93%

       Already formatted:

       99.65 stays 99.65
    */

    if (
        Math.abs(
            number
        ) <= 1.5
    ) {

        return number * 100;

    }


    return number;

}


/* =========================================================
   DAILY LABELS
========================================================= */

function dailyLabels(
    records
) {

    return records.map(
        record => {

            const day =
                record.date.getDate();


            /*
               Show every second day when
               there are approximately 20–35
               records in the selected month.

               Data remains daily.
            */

            if (
                records.length >= 20 &&
                records.length <= 35
            ) {

                return (
                    day % 2 === 0
                        ? String(day)
                        : ""
                );

            }


            return formatShortDate(
                record.date
            );

        }
    );

}


/* =========================================================
   DAILY LINE CHART
========================================================= */

function createDailyLineChart(
    canvasId,
    labels,
    values,
    datasetLabel,
    yTitle,
    options = {}
) {

    const canvas =
        $(canvasId);


    if (!canvas) {
        return;
    }


    destroyChart(
        canvasId
    );


    const parent =
        canvas.parentElement;


    const width =
        Math.max(
            parent?.clientWidth || 700,
            labels.length * 60
        );


    prepareScrollableCanvas(
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

                            label:
                                datasetLabel,

                            data:
                                values,

                            borderWidth:
                                2,

                            pointRadius:
                                3,

                            pointHoverRadius:
                                6,

                            tension:
                                0.22,

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
                                    0

                            }

                        },


                        y: {

                            beginAtZero:
                                options.beginAtZero ||
                                false,

                            min:
                                options.min,

                            max:
                                options.max,

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
   DAILY CHARTS
   CURRENT WORKING VERSION
========================================================= */

function renderDailyCharts(
    records
) {

    if (!records.length) {
        return;
    }


    const labels =
        dailyLabels(
            records
        );


    /*
       PR — KEEP CURRENT
    */

    createDailyLineChart(
        "prChart",
        labels,
        records.map(
            record =>
                record.pr
        ),
        "Performance Ratio",
        "PR (%)",
        {
            min:
                0,

            max:
                100
        }
    );


    createDailyLineChart(
        "dashboardPRChart",
        labels,
        records.map(
            record =>
                record.pr
        ),
        "Performance Ratio",
        "PR (%)",
        {
            min:
                0,

            max:
                100
        }
    );


    /*
       Operating Hours
    */

    createDailyLineChart(
        "hoursChart",
        labels,
        records.map(
            record =>
                record.hours
        ),
        "Operating Hours",
        "Operating Hours",
        {
            beginAtZero:
                true
        }
    );


    /*
       SYSTEM LOSS — KEEP CURRENT
    */

    createDailyLineChart(
        "lossChart",
        labels,
        records.map(
            record =>
                record.loss
        ),
        "System Loss",
        "System Loss (%)",
        {
            beginAtZero:
                true
        }
    );


    createDailyLineChart(
        "dashboardLossChart",
        labels,
        records.map(
            record =>
                record.loss
        ),
        "System Loss",
        "System Loss (%)",
        {
            beginAtZero:
                true
        }
    );


    /*
       PA — KEEP CURRENT
    */

    renderPAPercentageChart(
        records
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
                        Daily PA (%) from Daily_KPI · Column S
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


        const firstCard =
            section.querySelector(
                ".chart-card.full-card"
            );


        if (firstCard) {

            section.insertBefore(
                card,
                firstCard
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
   PA PERCENTAGE
========================================================= */

function renderPAPercentageChart(
    records
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
        dailyLabels(
            records
        );


    const width =
        Math.max(
            700,
            records.length * 60
        );


    prepareScrollableCanvas(
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
                                "Plant Availability",

                            data:
                                records.map(
                                    record =>
                                        record.pa
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

    const rows =
        toMatrix(
            getSheet(
                "PA"
            )
        );


    const records = [];


    rows.forEach(
        row => {

            const issue =
                getCell(
                    row,
                    "W"
                );


            const start =
                timeToMinutes(
                    getCell(
                        row,
                        "Z"
                    )
                );


            const end =
                timeToMinutes(
                    getCell(
                        row,
                        "AC"
                    )
                );


            if (
                issue === null ||
                issue === undefined
            ) {

                return;

            }


            const name =
                String(
                    issue
                ).trim();


            if (!name) {
                return;
            }


            if (
                start === null ||
                end === null
            ) {

                return;

            }


            let actualEnd =
                end;


            if (
                actualEnd < start
            ) {

                actualEnd +=
                    1440;

            }


            records.push({

                issue:
                    name,

                start,

                end:
                    actualEnd

            });

        }
    );


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
            "No plant unavailability records found."
        );

        return;

    }


    prepareScrollableCanvas(
        canvas,
        2304
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


                                        if (!item) {
                                            return "";
                                        }


                                        return [

                                            `Start: ${minutesToTime(item.start)}`,

                                            `End: ${minutesToTime(item.end)}`,

                                            `Duration: ${
                                                item.end -
                                                item.start
                                            } min`

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

                            ticks: {

                                stepSize:
                                    30,

                                callback:
                                    value =>
                                        minutesToTime(
                                            value
                                        )

                            },

                            title: {

                                display:
                                    true,

                                text:
                                    "Time"

                            }

                        },


                        y: {

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
   BREAKDOWN TIMELINE
   RESTORED FROM YOUR WORKING VERSION
========================================================= */

function readBreakdownTimeline() {

    const rows =
        toMatrix(
            getSheet(
                "PA"
            )
        );


    const grouped =
        new Map();


    rows.forEach(
        row => {

            const date =
                parseDate(
                    getCell(
                        row,
                        "B"
                    )
                );


            const minutes =
                parseNumber(
                    getCell(
                        row,
                        "AG"
                    )
                );


            if (
                !date ||
                minutes === null
            ) {

                return;

            }


            const key =
                makeDateKey(
                    date
                );


            if (
                !grouped.has(
                    key
                )
            ) {

                grouped.set(
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


            grouped.get(
                key
            ).minutes +=
                minutes;

        }
    );


    return Array.from(
        grouped.values()
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

function makeDateKey(
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
                        Same-date breakdown times combined from PA · Column AG
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


        section.appendChild(
            card
        );

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
            "No breakdown timeline records found."
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
                                formatShortDate(
                                    record.date
                                )
                        ),

                    datasets: [

                        {

                            label:
                                "Breakdown Time (min)",

                            data:
                                records.map(
                                    record =>
                                        record.minutes
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

    const rows =
        toMatrix(
            getSheet(
                "PA"
            )
        );


    const grouped =
        new Map();


    rows.forEach(
        row => {

            const date =
                parseDate(
                    getCell(
                        row,
                        "B"
                    )
                );


            const loss =
                parseNumber(
                    getCell(
                        row,
                        "AL"
                    )
                );


            if (
                !date ||
                loss === null
            ) {

                return;

            }


            const key =
                makeDateKey(
                    date
                );


            if (
                !grouped.has(
                    key
                )
            ) {

                grouped.set(
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


            grouped.get(
                key
            ).loss +=
                loss;

        }
    );


    return Array.from(
        grouped.values()
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
                        Same-date system losses combined from PA · Column AL
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


        section.appendChild(
            card
        );

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
            "No system loss records found."
        );

        return;

    }


    const width =
        Math.max(
            700,
            records.length * 60
        );


    prepareScrollableCanvas(
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
                                formatShortDate(
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
   RESTORED FROM YOUR WORKING VERSION
========================================================= */

function readCurtailment() {

    const rows =
        toMatrix(
            getSheet(
                "Curtailment records"
            )
        );


    const records = [];


    rows.forEach(
        row => {

            const date =
                parseDate(
                    getCell(
                        row,
                        "C"
                    )
                );


            const start =
                timeToMinutes(
                    getCell(
                        row,
                        "H"
                    )
                );


            const end =
                timeToMinutes(
                    getCell(
                        row,
                        "I"
                    )
                );


            const loss =
                parseNumber(
                    getCell(
                        row,
                        "R"
                    )
                );


            if (
                !date ||
                start === null ||
                end === null
            ) {

                return;

            }


            let actualEnd =
                end;


            if (
                actualEnd < start
            ) {

                actualEnd +=
                    1440;

            }


            records.push({

                date,

                key:
                    makeDateKey(
                        date
                    ),

                start,

                end:
                    actualEnd,

                loss:
                    loss === null
                        ? 0
                        : loss

            });

        }
    );


    records.sort(
        (a, b) =>
            a.date -
            b.date ||
            a.start -
            b.start
    );


    return records;

}


/* =========================================================
   CURTAILMENT LOSS CHART
========================================================= */

function renderCurtailmentLossChart(
    records
) {

    const canvas =
        $("curtailmentChart");


    if (!canvas) {
        return;
    }


    destroyChart(
        "curtailmentChart"
    );


    if (!records.length) {

        showCanvasMessage(
            canvas,
            "No curtailment records found."
        );

        return;

    }


    const grouped =
        new Map();


    records.forEach(
        record => {

            if (
                !grouped.has(
                    record.key
                )
            ) {

                grouped.set(
                    record.key,
                    {

                        date:
                            record.date,

                        loss:
                            0

                    }
                );

            }


            grouped.get(
                record.key
            ).loss +=
                record.loss;

        }
    );


    const daily =
        Array.from(
            grouped.values()
        )
        .sort(
            (a, b) =>
                a.date -
                b.date
        );


    const labels =
        daily.map(
            record =>
                formatShortDate(
                    record.date
                )
        );


    const values =
        daily.map(
            record =>
                record.loss
        );


    const width =
        Math.max(
            750,
            labels.length * 65
        );


    prepareScrollableCanvas(
        canvas,
        width
    );


    charts.curtailmentChart =
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
                                "Curtailment Loss (MWh)",

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
                                    true,

                                maxTicksLimit:
                                    15

                            },

                            title: {

                                display:
                                    true,

                                text:
                                    "Date"

                            }

                        },


                        y: {

                            beginAtZero:
                                true,

                            title: {

                                display:
                                    true,

                                text:
                                    "Curtailment Loss (MWh)"

                            }

                        }

                    }

                }

            }
        );


    const total =
        values.reduce(
            (sum, value) =>
                sum + value,
            0
        );


    setText(
        "curtailmentSummary",
        `${daily.length} date(s) · ${total.toFixed(2)} MWh total generation loss`
    );

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
                        Loss of generation merged across the same date · Column R
                    </span>

                </div>

                <span class="chart-type">
                    TABLE
                </span>

            </div>

            <div id="curtailmentTable"></div>

        `;


        section.appendChild(
            card
        );

    }


    return $("curtailmentTable");

}


/* =========================================================
   CURTAILMENT TABLE
========================================================= */

function renderCurtailmentTable(
    records
) {

    const container =
        ensureCurtailmentTable();


    if (!container) {
        return;
    }


    const grouped =
        new Map();


    records.forEach(
        record => {

            if (
                !grouped.has(
                    record.key
                )
            ) {

                grouped.set(
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


            grouped.get(
                record.key
            ).loss +=
                record.loss;


            grouped.get(
                record.key
            ).intervals++;

        }
    );


    const daily =
        Array.from(
            grouped.values()
        )
        .sort(
            (a, b) =>
                a.date -
                b.date
        );


    if (!daily.length) {

        container.innerHTML = `

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
                        padding:10px;
                        text-align:left;
                        border-bottom:1px solid #e1ebed;
                    ">
                        Date
                    </th>

                    <th style="
                        padding:10px;
                        text-align:right;
                        border-bottom:1px solid #e1ebed;
                    ">
                        Loss of Generation (MWh)
                    </th>

                    <th style="
                        padding:10px;
                        text-align:right;
                        border-bottom:1px solid #e1ebed;
                    ">
                        Number of Intervals
                    </th>

                </tr>

            </thead>

            <tbody>
    `;


    daily.forEach(
        record => {

            html += `

                <tr>

                    <td style="
                        padding:10px;
                        border-bottom:1px solid #edf2f3;
                    ">
                        ${formatFullDate(
                            record.date
                        )}
                    </td>

                    <td style="
                        padding:10px;
                        text-align:right;
                        border-bottom:1px solid #edf2f3;
                    ">
                        ${record.loss.toFixed(2)}
                    </td>

                    <td style="
                        padding:10px;
                        text-align:right;
                        border-bottom:1px solid #edf2f3;
                    ">
                        ${record.intervals}
                    </td>

                </tr>

            `;

        }
    );


    html += `

            </tbody>

        </table>

    `;


    container.innerHTML =
        html;

}


/* =========================================================
   CURTAILMENT GANTT CARD
========================================================= */

function ensureCurtailmentGanttCard() {

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
                        Date-wise duration from 06:00 to 18:00 · 15-minute intervals
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


        section.appendChild(
            card
        );

    }


    return $("curtailmentGanttChart");

}


/* =========================================================
   CURTAILMENT GANTT
========================================================= */

function renderCurtailmentGantt(
    records
) {

    const canvas =
        ensureCurtailmentGanttCard();


    if (!canvas) {
        return;
    }


    destroyChart(
        "curtailmentGanttChart"
    );


    if (!records.length) {

        showCanvasMessage(
            canvas,
            "No curtailment duration records found."
        );

        return;

    }


    prepareScrollableCanvas(
        canvas,
        2400
    );


    const uniqueDates = [];


    records.forEach(
        record => {

            const label =
                formatShortDate(
                    record.date
                );


            if (
                !uniqueDates.includes(
                    label
                )
            ) {

                uniqueDates.push(
                    label
                );

            }

        }
    );


    const datasets =
        records
            .filter(
                record =>
                    record.end > 360 &&
                    record.start < 1080
            )
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


                    return {

                        label:
                            `${formatShortDate(
                                record.date
                            )} ${minutesToTime(
                                start
                            )}`,

                        data: [

                            {

                                x: [
                                    start,
                                    end
                                ],

                                y:
                                    formatShortDate(
                                        record.date
                                    )

                            }

                        ],

                        borderWidth:
                            1,

                        borderRadius:
                            4,

                        barThickness:
                            20

                    };

                }
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

                    labels:
                        uniqueDates,

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
                                        context[0]
                                            ?.raw
                                            ?.y ||
                                        "",


                                label:
                                    context => {

                                        const raw =
                                            context.raw;


                                        if (!raw) {
                                            return "";
                                        }


                                        return (
                                            `Time: ${minutesToTime(
                                                raw.x[0]
                                            )} – ${minutesToTime(
                                                raw.x[1]
                                            )}`
                                        );

                                    }

                            }

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

                            labels:
                                uniqueDates,

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
========================================================= */

function readMonthlyPR() {

    const sheet =
        getSheet(
            "Annual_KPI"
        );


    if (!sheet) {
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


    const result = [];


    for (
        let row = 10;
        row <= 21;
        row++
    ) {

        const target =
            readNumericCell(
                sheet,
                `H${row}`
            );


        const measured =
            readNumericCell(
                sheet,
                `I${row}`
            );


        if (
            target === null &&
            measured === null
        ) {

            continue;

        }


        result.push({

            month:
                months[
                    row - 10
                ],

            target:
                convertPercentage(
                    target
                ),

            measured:
                convertPercentage(
                    measured
                )

        });

    }


    return result;

}


/* =========================================================
   MONTHLY PR CHART
========================================================= */

function renderMonthlyPRChart() {

    /*
       This is retained separately from
       the daily PR graph.

       It does NOT replace the current
       Daily_KPI PR graph.
    */

    const records =
        readMonthlyPR();


    if (!records.length) {
        return;
    }


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
                        Annual_KPI · Measured PR(%) · Column I
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


        section.appendChild(
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
   RESTORED FROM YOUR WORKING VERSION
========================================================= */

function renderEnergyChart() {

    const sheet =
        getSheet(
            "Annual_KPI"
        );


    if (!sheet) {
        return;
    }


    const rows =
        toMatrix(
            sheet
        );


    const budget = [];
    const measured = [];


    /*
       EXACT:
       E10:E21
       F10:F21
    */

    for (
        let excelRow = 10;
        excelRow <= 21;
        excelRow++
    ) {

        const row =
            rows[
                excelRow - 1
            ] ||
            [];


        budget.push(
            parseNumber(
                getCell(
                    row,
                    "E"
                )
            )
        );


        measured.push(
            parseNumber(
                getCell(
                    row,
                    "F"
                )
            )
        );

    }


    const labels = [

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
                                    "Energy (MWh)"

                            }

                        }

                    }

                }

            }
        );

}


/* =========================================================
   CANVAS EMPTY MESSAGE
========================================================= */

function showCanvasMessage(
    canvas,
    message
) {

    if (!canvas) {
        return;
    }


    const ctx =
        canvas.getContext(
            "2d"
        );


    if (!ctx) {
        return;
    }


    ctx.clearRect(
        0,
        0,
        canvas.width,
        canvas.height
    );


    ctx.save();


    ctx.textAlign =
        "center";


    ctx.textBaseline =
        "middle";


    ctx.font =
        "12px Inter, Arial";


    ctx.fillStyle =
        "#879397";


    ctx.fillText(
        message,
        canvas.width / 2,
        canvas.height / 2
    );


    ctx.restore();

}


/* =========================================================
   DASHBOARD KPI
========================================================= */

function renderDashboardKPI(
    records
) {

    if (!records.length) {
        return;
    }


    const latest =
        records[
            records.length - 1
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
   RENDER DASHBOARD
========================================================= */

function renderDashboard() {

    destroyAllCharts();

    removeDynamicCards();


    /*
       DAILY KPI
    */

    const daily =
        readDailyKPI();


    if (daily.length) {

        renderDashboardKPI(
            daily
        );


        renderDailyCharts(
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


    renderCurtailmentLossChart(
        curtailment
    );


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
   REMOVE DYNAMIC CARDS
========================================================= */

function removeDynamicCards() {

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

}


/* =========================================================
   RESET DASHBOARD
========================================================= */

function resetDashboard() {

    destroyAllCharts();

    removeDynamicCards();


    workbook =
        null;


    const input =
        $("dgrFile");


    if (input) {

        input.value =
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
        "sidebarFileName",
        "No DGR uploaded"
    );


    setText(
        "curtailmentSummary",
        "Waiting for DGR data"
    );


    setStatus(
        "Upload a DGR to generate the analytics."
    );


    hideAnalytics();

}


/* =========================================================
   END
========================================================= */
```
