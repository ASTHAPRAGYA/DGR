"use strict";

/* =========================================================
   SOLAR DGR ANALYTICS
   COMPLETE app.js

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
   AG = Breakdown Time (minutes)
   AL = System Loss (MWh)

   CURTAILMENT RECORDS
   ---------------------------------------------------------
   C  = Date
   H  = Start Time
   I  = End Time
   R  = Loss of Generation (MWh)

   ANNUAL_KPI
   ---------------------------------------------------------
   H9 = Target PR
   I9 = Measured PR

   E10:E21 = Budgeted Energy
   F10:F21 = Measured Energy

   ========================================================= */


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

        console.log(
            "SOLAR DGR APP.JS LOADED - FULL CHART VERSION"
        );

        installChartScrollStyles();

        setupNavigation();

        setupUpload();

        setupRemove();

        hideAnalytics();

    }
);


/* =========================================================
   CHART SCROLL STYLES
========================================================= */

function installChartScrollStyles() {

    if (
        document.getElementById(
            "solar-chart-scroll-style"
        )
    ) {
        return;
    }


    const style =
        document.createElement("style");


    style.id =
        "solar-chart-scroll-style";


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
            max-width: none !important;
        }

        .dgr-data-table {
            width: 100%;
            border-collapse: collapse;
            font-size: 10px;
        }

        .dgr-data-table th {
            padding: 10px;
            text-align: left;
            border-bottom: 1px solid #e1ebed;
            color: #637477;
            font-weight: 700;
        }

        .dgr-data-table td {
            padding: 10px;
            border-bottom: 1px solid #edf2f3;
            color: #44565a;
        }

        .dgr-data-table .number-cell {
            text-align: right;
        }

    `;


    document.head.appendChild(style);

}


/* =========================================================
   NAVIGATION
========================================================= */

function setupNavigation() {

    const buttons =
        document.querySelectorAll(".nav-item");


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
   UPLOAD
========================================================= */

function setupUpload() {

    const input =
        $("dgrFile");


    const dropZone =
        $("dropZone");


    if (!input) {

        console.error(
            "ERROR: #dgrFile not found."
        );

        return;

    }


    input.addEventListener(
        "change",
        function (event) {

            const file =
                event.target.files &&
                event.target.files[0];


            if (file) {

                processDGR(file);

            }

        }
    );


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
                event.dataTransfer &&
                event.dataTransfer.files &&
                event.dataTransfer.files[0];


            if (file) {

                processDGR(file);

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

function processDGR(file) {

    if (!file) {
        return;
    }


    if (
        !/\.(xlsx|xls|csv)$/i.test(
            file.name
        )
    ) {

        alert(
            "Please upload an Excel file (.xlsx, .xls or .csv)."
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


                console.log(
                    "Workbook sheets:",
                    workbook.SheetNames
                );


                updateFileUI(file);

                showAnalytics();

                renderDashboard();

                setStatus(
                    `${file.name} loaded successfully.`
                );

            }

            catch (error) {

                console.error(
                    "DGR PROCESSING ERROR:",
                    error
                );


                alert(
                    "Unable to process the DGR.\n\n" +
                    error.message
                );


                setStatus(
                    "Unable to process the DGR."
                );

            }

        };


    reader.onerror =
        function () {

            setStatus(
                "Unable to read the selected file."
            );

        };


    reader.readAsArrayBuffer(file);

}


/* =========================================================
   FILE UI
========================================================= */

function updateFileUI(file) {

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
        .remove("hidden");


    $("workbookStatus")
        ?.classList
        .remove("hidden");


    $("emptyState")
        ?.classList
        .add("hidden");


    $("dropZone")
        ?.classList
        .add("hidden");


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
                document.createElement("span");


            badge.className =
                "sheet-badge";


            const exists =
                workbook.SheetNames.some(
                    function (actual) {

                        return (
                            normalizeSheet(actual) ===
                            normalizeSheet(sheetName)
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
   NORMALIZE SHEET
========================================================= */

function normalizeSheet(name) {

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

function getSheet(requestedName) {

    if (
        !workbook ||
        !workbook.Sheets
    ) {
        return null;
    }


    if (
        workbook.Sheets[requestedName]
    ) {

        return workbook.Sheets[requestedName];

    }


    const wanted =
        normalizeSheet(
            requestedName
        );


    const actual =
        workbook.SheetNames.find(
            function (sheetName) {

                return (
                    normalizeSheet(
                        sheetName
                    ) === wanted
                );

            }
        );


    return actual
        ? workbook.Sheets[actual]
        : null;

}


/* =========================================================
   MATRIX
========================================================= */

function toMatrix(sheet) {

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
   COLUMN LETTER -> INDEX
========================================================= */

function columnIndex(letter) {

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
   GET CELL FROM MATRIX ROW
========================================================= */

function getCell(row, column) {

    if (!row) {
        return null;
    }


    return row[
        columnIndex(column)
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

function parseNumber(value) {

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

        return Number.isFinite(value)
            ? value
            : null;

    }


    if (
        value instanceof Date
    ) {

        return null;

    }


    const text =
        String(value)
            .trim()
            .replace(/,/g, "")
            .replace(/%/g, "");


    if (
        !text ||
        text.startsWith("#")
    ) {

        return null;

    }


    const number =
        Number(text);


    return Number.isFinite(number)
        ? number
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
        Number.isFinite(cell.v)
    ) {

        return cell.v;

    }


    if (
        cell.w !== undefined &&
        cell.w !== null
    ) {

        const parsed =
            parseNumber(cell.w);


        if (
            parsed !== null
        ) {

            return parsed;

        }

    }


    if (
        cell.v !== undefined &&
        cell.v !== null
    ) {

        return parseNumber(cell.v);

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


    if (
        value instanceof Date
    ) {

        return isNaN(value.getTime())
            ? null
            : new Date(value.getTime());

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
        String(value)
            .trim();


    let match =
        text.match(
            /^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})/
        );


    if (match) {

        const year =
            Number(match[3]);

        const month =
            Number(match[2]) - 1;

        const day =
            Number(match[1]);


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
                    .substring(0, 3)
                    .toLowerCase()
            );


        let year =
            Number(match[3]);


        if (
            year < 100
        ) {

            year += 2000;

        }


        if (
            month >= 0
        ) {

            return new Date(
                year,
                month,
                Number(match[1])
            );

        }

    }


    const browserDate =
        new Date(text);


    return isNaN(
        browserDate.getTime()
    )
        ? null
        : browserDate;

}


/* =========================================================
   READ DATE DIRECTLY FROM WORKSHEET
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
       If the cell contains a formula,
       the cached value can sometimes be
       repeated/stale. Formula dates are
       handled by the Daily_KPI reader.
    */

    if (
        cell.v instanceof Date
    ) {

        return parseDate(
            cell.v
        );

    }


    if (
        cell.v !== undefined &&
        cell.v !== null
    ) {

        const date =
            parseDate(
                cell.v
            );


        if (date) {
            return date;
        }

    }


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


    return null;

}


/* =========================================================
   PERCENTAGE CONVERSION
========================================================= */

function convertPercentage(value) {

    if (
        value === null ||
        value === undefined
    ) {
        return null;
    }


    const number =
        Number(value);


    if (
        !Number.isFinite(number)
    ) {
        return null;
    }


    if (
        Math.abs(number) <= 1.5
    ) {

        return number * 100;

    }


    return number;

}


/* =========================================================
   TIME TO MINUTES
========================================================= */

function timeToMinutes(value) {

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


            if (parsed) {

                return (
                    parsed.H * 60 +
                    parsed.M
                );

            }

        }

        catch (_) {}

    }


    const text =
        String(value)
            .trim();


    const match =
        text.match(
            /^(\d{1,2}):(\d{2})(?::(\d{2}))?\s*(AM|PM)?$/i
        );


    if (!match) {
        return null;
    }


    let hour =
        Number(match[1]);


    const minute =
        Number(match[2]);


    const period =
        match[4]
            ? match[4].toUpperCase()
            : null;


    if (
        period === "PM" &&
        hour < 12
    ) {

        hour += 12;

    }


    if (
        period === "AM" &&
        hour === 12
    ) {

        hour = 0;

    }


    if (
        hour < 0 ||
        hour > 23 ||
        minute < 0 ||
        minute > 59
    ) {

        return null;

    }


    return (
        hour * 60 +
        minute
    );

}


/* =========================================================
   TIME FORMAT
========================================================= */

function minutesToTime(value) {

    let minutes =
        Number(value);


    if (
        !Number.isFinite(minutes)
    ) {

        return "—";

    }


    minutes =
        Math.max(
            0,
            Math.min(
                1439,
                Math.round(minutes)
            )
        );


    const hours =
        Math.floor(
            minutes / 60
        );


    const mins =
        minutes % 60;


    return (
        String(hours).padStart(2, "0") +
        ":" +
        String(mins).padStart(2, "0")
    );

}


/* =========================================================
   DATE FORMATTING
========================================================= */

function formatShortDate(value) {

    const date =
        value instanceof Date
            ? value
            : parseDate(value);


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


function formatFullDate(value) {

    const date =
        value instanceof Date
            ? value
            : parseDate(value);


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
   NUMBER FORMAT
========================================================= */

function formatNumber(
    value,
    decimals = 2
) {

    if (
        value === null ||
        value === undefined ||
        !Number.isFinite(
            Number(value)
        )
    ) {

        return "—";

    }


    return Number(value)
        .toLocaleString(
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
   TEXT
========================================================= */

function setText(id, value) {

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

function setStatus(message) {

    setText(
        "statusText",
        message
    );

}


/* =========================================================
   DESTROY CHART
========================================================= */

function destroyChart(id) {

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


    /*
       Also clear any Chart.js instance
       attached directly to the canvas.
    */

    const canvas =
        $(id);


    if (canvas) {

        try {

            const existing =
                Chart.getChart(canvas);


            if (existing) {

                existing.destroy();

            }

        }

        catch (_) {}

    }

}


/* =========================================================
   DESTROY ALL
========================================================= */

function destroyAllCharts() {

    Object.keys(charts)
        .forEach(
            function (id) {

                destroyChart(id);

            }
        );


    document
        .querySelectorAll("canvas")
        .forEach(
            function (canvas) {

                try {

                    const existing =
                        Chart.getChart(canvas);


                    if (existing) {

                        existing.destroy();

                    }

                }

                catch (_) {}

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
        function (id) {

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
        .add("hidden");

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


    const finalWidth =
        Math.max(
            parent.clientWidth || 700,
            Number(width) || 700
        );


    canvas.style.setProperty(
        "width",
        `${finalWidth}px`,
        "important"
    );


    canvas.style.setProperty(
        "min-width",
        `${finalWidth}px`,
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
   GET DAILY KPI
========================================================= */

function readDailyKPI() {

    const sheet =
        getSheet("Daily_KPI");


    if (
        !sheet ||
        !sheet["!ref"]
    ) {

        console.warn(
            "Daily_KPI sheet unavailable."
        );

        return [];

    }


    const range =
        XLSX.utils.decode_range(
            sheet["!ref"]
        );


    const records = [];


    let previousDate =
        null;


    /*
       Data normally starts around row 5.
       We scan the entire worksheet and
       only retain rows with a valid date
       or a formula-based continuation.
    */

    for (
        let r = range.s.r;
        r <= range.e.r;
        r++
    ) {

        const rowNumber =
            r + 1;


        const dateCell =
            sheet[`B${rowNumber}`];


        let date =
            null;


        /*
           Formula date:
           e.g. =B5+1

           Use the previous actual date
           instead of trusting a repeated
           cached value.
        */

        if (
            dateCell &&
            dateCell.f &&
            previousDate
        ) {

            date =
                new Date(previousDate);


            date.setDate(
                date.getDate() + 1
            );

        }

        else {

            date =
                readDateCell(
                    sheet,
                    `B${rowNumber}`
                );

        }


        if (!date) {
            continue;
        }


        /*
           If a repeated/stale cached date
           appears after a valid date, advance
           sequentially when the cell is a
           formula.
        */

        if (
            previousDate &&
            date <= previousDate &&
            dateCell &&
            dateCell.f
        ) {

            date =
                new Date(previousDate);


            date.setDate(
                date.getDate() + 1
            );

        }


        previousDate =
            new Date(date);


        const hours =
            readNumericCell(
                sheet,
                `I${rowNumber}`
            );


        const paRaw =
            readNumericCell(
                sheet,
                `S${rowNumber}`
            );


        const prRaw =
            readNumericCell(
                sheet,
                `V${rowNumber}`
            );


        const lossRaw =
            readNumericCell(
                sheet,
                `AD${rowNumber}`
            );


        /*
           Skip obvious header rows.
        */

        if (
            hours === null &&
            paRaw === null &&
            prRaw === null &&
            lossRaw === null
        ) {

            continue;

        }


        records.push({

            date,

            hours,

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


    records.sort(
        function (a, b) {

            return (
                a.date -
                b.date
            );

        }
    );


    console.log(
        "Daily KPI records:",
        records
    );


    return records;

}


/* =========================================================
   DAILY X-AXIS LABELS
========================================================= */

function makeDailyLabels(
    records
) {

    return records.map(
        function (record, index) {

            const day =
                record.date.getDate();


            /*
               For approximately one month,
               show:
               2, 4, 6, 8, 10, ...

               Data itself remains daily.
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
   DAILY PR SCATTER
========================================================= */

function renderDailyPRScatter(
    records
) {

    const canvas =
        $("prChart");


    if (!canvas) {
        return;
    }


    destroyChart("prChart");


    const valid =
        records.filter(
            record =>
                Number.isFinite(
                    record.pr
                )
        );


    if (!valid.length) {

        showCanvasMessage(
            canvas,
            "No valid PR data found in Daily_KPI Column V."
        );

        return;

    }


    const width =
        Math.max(
            700,
            valid.length * 55
        );


    prepareScrollableCanvas(
        canvas,
        width
    );


    const data =
        valid.map(
            function (record) {

                return {

                    x:
                        record.date.getDate(),

                    y:
                        record.pr

                };

            }
        );


    charts.prChart =
        new Chart(
            canvas.getContext("2d"),
            {

                type:
                    "scatter",

                data: {

                    datasets: [

                        {

                            label:
                                "Performance Ratio",

                            data,

                            pointRadius:
                                5,

                            pointHoverRadius:
                                8,

                            borderWidth:
                                1

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
                            "nearest",

                        intersect:
                            false

                    },


                    plugins: {

                        legend: {

                            display:
                                false

                        },


                        tooltip: {

                            callbacks: {

                                title:
                                    function (items) {

                                        if (!items.length) {
                                            return "";
                                        }


                                        const day =
                                            items[0]
                                                .parsed
                                                .x;


                                        const item =
                                            valid.find(
                                                record =>
                                                    record.date.getDate() ===
                                                    day
                                            );


                                        return item
                                            ? formatFullDate(
                                                item.date
                                            )
                                            : `Day ${day}`;

                                    },


                                label:
                                    function (context) {

                                        return (
                                            "PR: " +
                                            Number(
                                                context.parsed.y
                                            ).toFixed(2) +
                                            "%"
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
                                1,

                            max:
                                Math.max(
                                    31,
                                    ...valid.map(
                                        record =>
                                            record.date.getDate()
                                    )
                                ),

                            title: {

                                display:
                                    true,

                                text:
                                    "Day of Month"

                            },

                            ticks: {

                                stepSize:
                                    2,

                                callback:
                                    function (value) {

                                        return value % 2 === 0
                                            ? value
                                            : "";

                                    }

                            },

                            grid: {

                                display:
                                    false

                            }

                        },


                        y: {

                            min:
                                0,

                            max:
                                100,

                            title: {

                                display:
                                    true,

                                text:
                                    "PR (%)"

                            },

                            ticks: {

                                stepSize:
                                    10,

                                callback:
                                    function (value) {

                                        return `${value}%`;

                                    }

                            }

                        }

                    }

                }

            }
        );


    /*
       Dashboard copy
    */

    renderDashboardPRScatter(
        valid
    );

}


/* =========================================================
   DASHBOARD PR
========================================================= */

function renderDashboardPRScatter(
    records
) {

    const canvas =
        $("dashboardPRChart");


    if (!canvas) {
        return;
    }


    destroyChart(
        "dashboardPRChart"
    );


    const width =
        Math.max(
            650,
            records.length * 45
        );


    prepareScrollableCanvas(
        canvas,
        width
    );


    charts.dashboardPRChart =
        new Chart(
            canvas.getContext("2d"),
            {

                type:
                    "scatter",

                data: {

                    datasets: [

                        {

                            label:
                                "PR",

                            data:
                                records.map(
                                    record => ({

                                        x:
                                            record.date.getDate(),

                                        y:
                                            record.pr

                                    })
                                ),

                            pointRadius:
                                4,

                            pointHoverRadius:
                                7

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

                            type:
                                "linear",

                            min:
                                1,

                            max:
                                31,

                            title: {

                                display:
                                    true,

                                text:
                                    "Day"

                            },

                            ticks: {

                                stepSize:
                                    2,

                                callback:
                                    function (value) {

                                        return value % 2 === 0
                                            ? value
                                            : "";

                                    }

                            }

                        },


                        y: {

                            min:
                                0,

                            max:
                                100,

                            title: {

                                display:
                                    true,

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

function renderOperatingHours(
    records
) {

    const canvas =
        $("hoursChart");


    if (!canvas) {
        return;
    }


    destroyChart("hoursChart");


    const labels =
        makeDailyLabels(records);


    prepareScrollableCanvas(
        canvas,
        Math.max(
            700,
            records.length * 55
        )
    );


    charts.hoursChart =
        new Chart(
            canvas.getContext("2d"),
            {

                type:
                    "line",

                data: {

                    labels,

                    datasets: [

                        {

                            label:
                                "Operating Hours",

                            data:
                                records.map(
                                    record =>
                                        record.hours
                                ),

                            borderWidth:
                                2,

                            pointRadius:
                                3,

                            pointHoverRadius:
                                6,

                            tension:
                                0.2,

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

                            },

                            title: {

                                display:
                                    true,

                                text:
                                    "Day of Month"

                            }

                        },


                        y: {

                            beginAtZero:
                                true,

                            title: {

                                display:
                                    true,

                                text:
                                    "Operating Hours"

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
   SYSTEM LOSS %
========================================================= */

function renderSystemLossPercent(
    records
) {

    const canvas =
        $("lossChart");


    if (!canvas) {
        return;
    }


    destroyChart("lossChart");


    const labels =
        makeDailyLabels(records);


    prepareScrollableCanvas(
        canvas,
        Math.max(
            700,
            records.length * 55
        )
    );


    charts.lossChart =
        new Chart(
            canvas.getContext("2d"),
            {

                type:
                    "line",

                data: {

                    labels,

                    datasets: [

                        {

                            label:
                                "System Loss",

                            data:
                                records.map(
                                    record =>
                                        record.loss
                                ),

                            borderWidth:
                                2,

                            pointRadius:
                                3,

                            pointHoverRadius:
                                6,

                            tension:
                                0.2,

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

                            },

                            title: {

                                display:
                                    true,

                                text:
                                    "Day of Month"

                            }

                        },


                        y: {

                            beginAtZero:
                                true,

                            title: {

                                display:
                                    true,

                                text:
                                    "System Loss (%)"

                            },

                            ticks: {

                                callback:
                                    function (value) {

                                        return `${value}%`;

                                    },

                                maxTicksLimit:
                                    7

                            }

                        }

                    }

                }

            }
        );


    renderDashboardSystemLoss(
        records
    );

}


/* =========================================================
   DASHBOARD SYSTEM LOSS
========================================================= */

function renderDashboardSystemLoss(
    records
) {

    const canvas =
        $("dashboardLossChart");


    if (!canvas) {
        return;
    }


    destroyChart(
        "dashboardLossChart"
    );


    prepareScrollableCanvas(
        canvas,
        Math.max(
            650,
            records.length * 45
        )
    );


    charts.dashboardLossChart =
        new Chart(
            canvas.getContext("2d"),
            {

                type:
                    "line",

                data: {

                    labels:
                        makeDailyLabels(
                            records
                        ),

                    datasets: [

                        {

                            label:
                                "System Loss",

                            data:
                                records.map(
                                    record =>
                                        record.loss
                                ),

                            borderWidth:
                                2,

                            pointRadius:
                                2,

                            tension:
                                0.2,

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
                                    "Loss (%)"

                            },

                            ticks: {

                                callback:
                                    function (value) {

                                        return `${value}%`;

                                    }

                            }

                        }

                    }

                }

            }
        );

}


/* =========================================================
   PA % CARD
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


        const first =
            section.querySelector(
                ".chart-card.full-card"
            );


        if (first) {

            section.insertBefore(
                card,
                first
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
   PA % CHART
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


    prepareScrollableCanvas(
        canvas,
        Math.max(
            700,
            records.length * 55
        )
    );


    charts.paPercentageChart =
        new Chart(
            canvas.getContext("2d"),
            {

                type:
                    "line",

                data: {

                    labels:
                        makeDailyLabels(
                            records
                        ),

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
                                0.2,

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

                            },

                            title: {

                                display:
                                    true,

                                text:
                                    "Day of Month"

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
                                    function (value) {

                                        return `${value}%`;

                                    }

                            }

                        }

                    }

                }

            }
        );

}


/* =========================================================
   PLANT UNAVAILABILITY DATA
========================================================= */

function readPlantUnavailability() {

    const rows =
        toMatrix(
            getSheet("PA")
        );


    const records = [];


    rows.forEach(
        function (row) {

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
                String(issue)
                    .trim();


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

                actualEnd += 1440;

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


    destroyChart("paChart");


    const records =
        readPlantUnavailability();


    if (!records.length) {

        showCanvasMessage(
            canvas,
            "No Plant Unavailability records found."
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
            function (record) {

                return {

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

                };

            }
        );


    charts.paChart =
        new Chart(
            canvas.getContext("2d"),
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
                                    function (context) {

                                        return (
                                            records[
                                                context[0]
                                                    .dataIndex
                                            ]?.issue || ""
                                        );

                                    },


                                label:
                                    function (context) {

                                        const item =
                                            records[
                                                context.dataIndex
                                            ];


                                        if (!item) {
                                            return "";
                                        }


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
                                1440,

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
                                    function (value) {

                                        return minutesToTime(
                                            value
                                        );

                                    }

                            },

                            grid: {

                                color:
                                    "rgba(23,37,42,0.08)"

                            }

                        },


                        y: {

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
   BREAKDOWN TIMELINE DATA
========================================================= */

function readBreakdownTimeline() {

    const rows =
        toMatrix(
            getSheet("PA")
        );


    const grouped =
        new Map();


    rows.forEach(
        function (row) {

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
                dateKey(date);


            if (
                !grouped.has(key)
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


            grouped.get(key).minutes +=
                minutes;

        }
    );


    return Array.from(
        grouped.values()
    )
    .sort(
        function (a, b) {

            return a.date - b.date;

        }
    );

}


/* =========================================================
   BREAKDOWN TIMELINE CARD
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


        section.appendChild(card);

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


    destroyChart("breakdownChart");


    const records =
        readBreakdownTimeline();


    if (!records.length) {

        showCanvasMessage(
            canvas,
            "No breakdown timeline records found."
        );

        return;

    }


    /*
       Keep requested axis:
       X = 0–13 minutes
       Y = dates
    */

    charts.breakdownChart =
        new Chart(
            canvas.getContext("2d"),
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

                        },

                        tooltip: {

                            callbacks: {

                                label:
                                    function (context) {

                                        return (
                                            "Breakdown: " +
                                            Number(
                                                context.parsed.x
                                            ).toFixed(2) +
                                            " min"
                                        );

                                    }

                            }

                        }

                    },


                    scales: {

                        x: {

                            min:
                                0,

                            max:
                                13,

                            title: {

                                display:
                                    true,

                                text:
                                    "Breakdown Time (minutes)"

                            },

                            ticks: {

                                stepSize:
                                    1

                            }

                        },


                        y: {

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
   SYSTEM LOSS MWh DATA
========================================================= */

function readSystemLossMWh() {

    const rows =
        toMatrix(
            getSheet("PA")
        );


    const grouped =
        new Map();


    rows.forEach(
        function (row) {

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
                dateKey(date);


            if (
                !grouped.has(key)
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


            grouped.get(key).loss +=
                loss;

        }
    );


    return Array.from(
        grouped.values()
    )
    .sort(
        function (a, b) {

            return a.date - b.date;

        }
    );

}


/* =========================================================
   SYSTEM LOSS CARD
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


        section.appendChild(card);

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


    prepareScrollableCanvas(
        canvas,
        Math.max(
            700,
            records.length * 60
        )
    );


    charts.systemLossMwhChart =
        new Chart(
            canvas.getContext("2d"),
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

                            },

                            ticks: {

                                autoSkip:
                                    false

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
                                    "System Loss (MWh)"

                            }

                        }

                    }

                }

            }
        );

}


/* =========================================================
   CURTAILMENT DATA
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
        function (row) {

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

                actualEnd += 1440;

            }


            records.push({

                date,

                key:
                    dateKey(date),

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
        function (a, b) {

            return (
                a.date - b.date ||
                a.start - b.start
            );

        }
    );


    return records;

}


/* =========================================================
   CURTAILMENT LOSS LINE CHART
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

        setText(
            "curtailmentSummary",
            "No curtailment records found"
        );

        return;

    }


    const grouped =
        new Map();


    records.forEach(
        function (record) {

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
            function (a, b) {

                return a.date - b.date;

            }
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


    prepareScrollableCanvas(
        canvas,
        Math.max(
            750,
            daily.length * 65
        )
    );


    charts.curtailmentChart =
        new Chart(
            canvas.getContext("2d"),
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
                                7,

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
            function (sum, value) {

                return sum + value;

            },
            0
        );


    setText(
        "curtailmentSummary",
        `${daily.length} date(s) · ${total.toFixed(2)} MWh total generation loss`
    );

}


/* =========================================================
   CURTAILMENT TABLE CARD
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


        section.appendChild(card);

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
        function (record) {

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
            function (a, b) {

                return a.date - b.date;

            }
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

        <table class="dgr-data-table">

            <thead>

                <tr>

                    <th>
                        Date
                    </th>

                    <th class="number-cell">
                        Loss of Generation (MWh)
                    </th>

                    <th class="number-cell">
                        Number of Intervals
                    </th>

                </tr>

            </thead>

            <tbody>

    `;


    daily.forEach(
        function (record) {

            html += `

                <tr>

                    <td>
                        ${formatFullDate(
                            record.date
                        )}
                    </td>

                    <td class="number-cell">
                        ${record.loss.toFixed(2)}
                    </td>

                    <td class="number-cell">
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


        section.appendChild(card);

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


    /*
       06:00 = 360 min
       18:00 = 1080 min
       15 min interval
    */

    prepareScrollableCanvas(
        canvas,
        2400
    );


    const uniqueDates = [];


    records.forEach(
        function (record) {

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
                function (record) {

                    return (
                        record.end > 360 &&
                        record.start < 1080
                    );

                }
            )
            .map(
                function (record) {

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
                            `${formatShortDate(
                                record.date
                            )} ${minutesToTime(
                                start
                            )}–${minutesToTime(
                                end
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
            )
            .filter(
                Boolean
            );


    charts.curtailmentGanttChart =
        new Chart(
            canvas.getContext("2d"),
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
                                    function (context) {

                                        return (
                                            context[0]
                                                ?.raw
                                                ?.y || ""
                                        );

                                    },


                                label:
                                    function (context) {

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
                                    function (value) {

                                        return minutesToTime(
                                            value
                                        );

                                    }

                            },

                            grid: {

                                color:
                                    "rgba(23,37,42,0.08)"

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
   MONTHLY PR DATA
   Annual_KPI H10:I21
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
   MONTHLY PR CARD
========================================================= */

function ensureMonthlyPRCard() {

    const section =
        $("performanceSection");


    if (!section) {
        return null;
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
                        Annual_KPI · Measured PR · Column I
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


        section.appendChild(card);

    }


    return card;

}


/* =========================================================
   MONTHLY PR CHART
   Y = months
   X = PR %
========================================================= */

function renderMonthlyPRChart() {

    const records =
        readMonthlyPR();


    if (!records.length) {
        return;
    }


    const card =
        ensureMonthlyPRCard();


    if (!card) {
        return;
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
            canvas.getContext("2d"),
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
                                    function (value) {

                                        return `${value}%`;

                                    }

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
   ENERGY CHART
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
        toMatrix(sheet);


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


    const budget = [];
    const measured = [];


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


    const totalBudget =
        budget.reduce(
            function (sum, value) {

                return (
                    sum +
                    (value || 0)
                );

            },
            0
        );


    const totalMeasured =
        measured.reduce(
            function (sum, value) {

                return (
                    sum +
                    (value || 0)
                );

            },
            0
        );


    const variance =
        totalMeasured -
        totalBudget;


    setText(
        "totalBudget",
        `${formatNumber(totalBudget)} MWh`
    );


    setText(
        "totalMeasured",
        `${formatNumber(totalMeasured)} MWh`
    );


    setText(
        "energyVariance",
        `${formatNumber(variance)} MWh`
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
            canvas.getContext("2d"),
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

                            },

                            title: {

                                display:
                                    true,

                                text:
                                    "Month"

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
   DATE KEY
========================================================= */

function dateKey(date) {

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

    ].join("-");

}


/* =========================================================
   CANVAS MESSAGE
========================================================= */

function showCanvasMessage(
    canvas,
    message
) {

    if (!canvas) {
        return;
    }


    try {

        const existing =
            Chart.getChart(canvas);


        if (existing) {

            existing.destroy();

        }

    }

    catch (_) {}


    const ctx =
        canvas.getContext("2d");


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
        function (id) {

            const element =
                $(id);


            if (element) {

                element.remove();

            }

        }
    );

}


/* =========================================================
   MAIN RENDER
========================================================= */

function renderDashboard() {

    console.log(
        "Rendering complete DGR dashboard..."
    );


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


        /*
           PR scatter
        */

        renderDailyPRScatter(
            daily
        );


        /*
           PA percentage
        */

        renderPAPercentageChart(
            daily
        );


        /*
           Operating hours
        */

        renderOperatingHours(
            daily
        );


        /*
           System loss %
        */

        renderSystemLossPercent(
            daily
        );

    }


    /*
       PA WORKSHEET
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
       ANNUAL KPI
    */

    renderMonthlyPRChart();

    renderEnergyChart();


    console.log(
        "DGR dashboard rendering complete."
    );

}


/* =========================================================
   RESET
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
        .add("hidden");


    $("workbookStatus")
        ?.classList
        .add("hidden");


    $("dropZone")
        ?.classList
        .remove("hidden");


    $("emptyState")
        ?.classList
        .remove("hidden");


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
        function (id) {

            setText(
                id,
                "—"
            );

        }
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
