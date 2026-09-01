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
   AG = Breakdown Time (minutes)
   AL = System Loss / Energy Loss

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
   H10:I21 = Monthly PR

   E10:E21 = Budgeted Energy
   F10:F21 = Measured Energy

   ========================================================= */

"use strict";


/* =========================================================
   GLOBAL VARIABLES
========================================================= */

let workbook = null;

const charts = {};

const fileInput =
    document.getElementById("dgrFile");


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

        setupRemoveButton();

        hideAnalytics();

    }
);


/* =========================================================
   SCROLL STYLES
========================================================= */

function installScrollStyles() {

    if (
        document.getElementById(
            "solar-dgr-scroll-style"
        )
    ) {
        return;
    }


    const style =
        document.createElement("style");


    style.id =
        "solar-dgr-scroll-style";


    style.textContent = `

        .dgr-scroll-wrapper {
            width: 100% !important;
            height: 100% !important;
            overflow-x: auto !important;
            overflow-y: hidden !important;
            position: relative !important;
            box-sizing: border-box !important;
            scrollbar-width: thin !important;
            scrollbar-color: #b9cacc transparent !important;
        }

        .dgr-scroll-wrapper::-webkit-scrollbar {
            height: 8px;
        }

        .dgr-scroll-wrapper::-webkit-scrollbar-track {
            background: #eef3f4;
            border-radius: 10px;
        }

        .dgr-scroll-wrapper::-webkit-scrollbar-thumb {
            background: #b9cacc;
            border-radius: 10px;
        }

        .dgr-scroll-wrapper::-webkit-scrollbar-thumb:hover {
            background: #27a5ad;
        }

        .dgr-scroll-wrapper > canvas {
            display: block !important;
            width: auto !important;
            max-width: none !important;
            min-width: 0 !important;
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
                function () {

                    items.forEach(
                        nav => {
                            nav.classList.remove(
                                "active"
                            );
                        }
                    );


                    item.classList.add(
                        "active"
                    );


                    const target =
                        $(item.dataset.target);


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
   UPLOAD SETUP
========================================================= */

function setupUpload() {

    const input =
        $("dgrFile");


    const dropZone =
        $("dropZone");


    if (!input) {

        console.error(
            "dgrFile input was not found."
        );

        return;

    }


    input.addEventListener(
        "change",
        function (event) {

            const file =
                event.target.files?.[0];


            if (file) {

                processFile(
                    file
                );

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
                event.dataTransfer?.files?.[0];


            if (file) {

                processFile(
                    file
                );

            }

        }
    );

}


/* =========================================================
   REMOVE FILE
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

    if (!file) {
        return;
    }


    if (
        !/\.(xlsx|xls|csv)$/i.test(
            file.name
        )
    ) {

        alert(
            "Please upload an Excel workbook (.xlsx/.xls) or CSV file."
        );

        return;

    }


    if (
        typeof XLSX === "undefined"
    ) {

        alert(
            "SheetJS is not loaded. Check the XLSX script in index.html."
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
                    !workbook.SheetNames ||
                    !workbook.SheetNames.length
                ) {

                    throw new Error(
                        "No worksheets were found."
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
                    "DGR processing error:",
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


    const requiredSheets = [
        "Dashboard",
        "Annual_KPI",
        "Daily_KPI",
        "PA",
        "Curtailment records"
    ];


    requiredSheets.forEach(
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
                        normalizeSheet(
                            actual
                        ) ===
                        normalizeSheet(
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
   NORMALIZE SHEET NAME
========================================================= */

function normalizeSheet(
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


    const normalized =
        normalizeSheet(
            requestedName
        );


    const actual =
        workbook.SheetNames.find(
            name =>
                normalizeSheet(
                    name
                ) === normalized
        );


    return actual
        ? workbook.Sheets[
            actual
        ]
        : null;

}


/* =========================================================
   GET USED RANGE
========================================================= */

function getSheetRange(
    sheet
) {

    if (
        !sheet ||
        !sheet["!ref"]
    ) {

        return null;

    }


    return XLSX.utils.decode_range(
        sheet["!ref"]
    );

}


/* =========================================================
   READ CELL OBJECT
========================================================= */

function getSheetCell(
    sheet,
    address
) {

    if (!sheet) {
        return null;
    }


    return sheet[address] || null;

}


/* =========================================================
   READ NUMERIC CELL

   Uses cached formula result first.
========================================================= */

function readNumericCell(
    sheet,
    address
) {

    const cell =
        getSheetCell(
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

        const text =
            String(
                cell.w
            )
                .replace(
                    /,/g,
                    ""
                )
                .replace(
                    /%/g,
                    ""
                )
                .trim();


        const value =
            Number(
                text
            );


        if (
            Number.isFinite(
                value
            )
        ) {

            return value;

        }

    }


    if (
        cell.v !== undefined &&
        cell.v !== null
    ) {

        const value =
            Number(
                cell.v
            );


        if (
            Number.isFinite(
                value
            )
        ) {

            return value;

        }

    }


    return null;

}


/* =========================================================
   READ DATE CELL

   For formula-driven dates, .w is preferred.
========================================================= */

function readDateCell(
    sheet,
    address
) {

    const cell =
        getSheetCell(
            sheet,
            address
        );


    if (!cell) {
        return null;
    }


    /*
       Displayed Excel value first.
       This is useful for formula dates.
    */

    if (
        cell.w !== undefined &&
        cell.w !== null &&
        String(
            cell.w
        ).trim() !== ""
    ) {

        const displayed =
            parseDate(
                cell.w
            );


        if (displayed) {

            return displayed;

        }

    }


    if (
        cell.v !== undefined &&
        cell.v !== null
    ) {

        const parsed =
            parseDate(
                cell.v
            );


        if (parsed) {

            return parsed;

        }

    }


    return null;

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


    let text =
        String(
            value
        )
            .trim()
            .replace(
                /,/g,
                ""
            );


    if (!text) {
        return null;
    }


    text =
        text.replace(
            /%/g,
            ""
        );


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
   CONVERT DECIMAL TO PERCENT
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
       Decimal Excel format:
       0.7547 -> 75.47
       0.0193 -> 1.93
       1      -> 100

       Already percentage:
       75.47 stays 75.47
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


    /*
       Excel serial date
    */

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
        )
            .trim();


    /*
       DD/MM/YYYY
       DD-MM-YYYY
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


    /*
       Month name by itself.
    */

    const monthsLong = [
        "january",
        "february",
        "march",
        "april",
        "may",
        "june",
        "july",
        "august",
        "september",
        "october",
        "november",
        "december"
    ];


    const monthOnly =
        monthsLong.indexOf(
            text.toLowerCase()
        );


    if (
        monthOnly >= 0
    ) {

        return new Date(
            2026,
            monthOnly,
            1
        );

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
   TIME PARSER
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


    /*
       Date/time object
    */

    if (
        value instanceof Date
    ) {

        if (
            isNaN(
                value.getTime()
            )
        ) {

            return null;

        }


        return (
            value.getHours() * 60 +
            value.getMinutes()
        );

    }


    /*
       Excel time fraction
    */

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

        const period =
            ampm.toUpperCase();


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

    const value =
        Math.max(
            0,
            Math.min(
                1439,
                Math.round(
                    minutes
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
   DATE FORMATTING
========================================================= */

function formatDate(
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


function formatFullDate(
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
   FORMAT NUMBER
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
    text
) {

    setText(
        "statusText",
        text
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
            id => {

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
            id => {

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
   DESTROY ONE CHART
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
                "Could not destroy chart:",
                id,
                error
            );

        }

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
            destroyChart
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
   CREATE SCROLL WRAPPER
========================================================= */

function makeChartScrollable(
    canvas,
    desiredWidth
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
            ".dgr-scroll-wrapper"
        );


    /*
       Remove accidental old wrappers
       if there are more than one.
    */

    const wrappers =
        parent.querySelectorAll(
            ".dgr-scroll-wrapper"
        );


    if (
        wrappers.length > 1
    ) {

        for (
            let i = 1;
            i < wrappers.length;
            i++
        ) {

            wrappers[i].parentNode
                ?.removeChild(
                    wrappers[i]
                );

        }

    }


    if (!wrapper) {

        wrapper =
            document.createElement(
                "div"
            );


        wrapper.className =
            "dgr-scroll-wrapper";


        parent.insertBefore(
            wrapper,
            canvas
        );


        wrapper.appendChild(
            canvas
        );

    }


    const visibleWidth =
        parent.clientWidth ||
        700;


    const finalWidth =
        Math.max(
            visibleWidth,
            desiredWidth || visibleWidth
        );


    wrapper.style.width =
        "100%";


    wrapper.style.height =
        "100%";


    wrapper.style.overflowX =
        "auto";


    wrapper.style.overflowY =
        "hidden";


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


    return wrapper;

}


/* =========================================================
   DAILY KPI READER
   IMPORTANT FIX FOR FORMULA DATES
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


    const range =
        getSheetRange(
            sheet
        );


    if (!range) {
        return [];
    }


    const rows = [];

    let previousDate = null;


    /*
       Row 4 = headers
       Row 5 = first data row
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
           If B contains a formula which
           isn't evaluated by SheetJS,
           continue the daily sequence.
        */

        if (
            !date &&
            previousDate
        ) {

            date =
                new Date(
                    previousDate
                );


            date.setDate(
                date.getDate() + 1
            );

        }


        if (!date) {
            continue;
        }


        /*
           Protect against a malformed
           repeated date.
        */

        if (
            previousDate &&
            date.getTime() <=
            previousDate.getTime()
        ) {

            const formula =
                sheet[
                    `B${excelRow}`
                ]?.f;


            /*
               If the row uses a sequential
               formula such as =B5+1,
               force sequential dates.
            */

            if (
                formula &&
                /B\d+\s*\+\s*1/i.test(
                    formula
                )
            ) {

                date =
                    new Date(
                        previousDate
                    );


                date.setDate(
                    date.getDate() + 1
                );

            }

        }


        previousDate =
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


        const pa =
            convertPercentage(
                paRaw
            );


        const pr =
            convertPercentage(
                prRaw
            );


        const loss =
            convertPercentage(
                lossRaw
            );


        rows.push({

            row:
                excelRow,

            date,

            hours:
                hoursRaw,

            pa,

            pr,

            loss

        });

    }


    /*
       Remove exact duplicate dates while
       retaining valid rows.

       This also protects against formula
       recalculation anomalies.
    */

    const unique =
        new Map();


    rows.forEach(
        row => {

            const key =
                dateKey(
                    row.date
                );


            /*
               Prefer a row that contains
               actual values.
            */

            if (
                !unique.has(key)
            ) {

                unique.set(
                    key,
                    row
                );

            }

            else {

                const existing =
                    unique.get(
                        key
                    );


                const existingScore =
                    Number(
                        existing.hours !== null
                    ) +
                    Number(
                        existing.pa !== null
                    ) +
                    Number(
                        existing.pr !== null
                    ) +
                    Number(
                        existing.loss !== null
                    );


                const newScore =
                    Number(
                        row.hours !== null
                    ) +
                    Number(
                        row.pa !== null
                    ) +
                    Number(
                        row.pr !== null
                    ) +
                    Number(
                        row.loss !== null
                    );


                if (
                    newScore >
                    existingScore
                ) {

                    unique.set(
                        key,
                        row
                    );

                }

            }

        }
    );


    const result =
        Array.from(
            unique.values()
        )
        .sort(
            (a, b) =>
                a.date -
                b.date
        );


    console.log(
        "Daily KPI records:",
        result
    );


    console.log(
        "PA values:",
        result.map(
            row => ({
                date:
                    formatFullDate(
                        row.date
                    ),
                pa:
                    row.pa
            })
        )
    );


    console.log(
        "System Loss values:",
        result.map(
            row => ({
                date:
                    formatFullDate(
                        row.date
                    ),
                loss:
                    row.loss
            })
        )
    );


    return result;

}


/* =========================================================
   DAILY LABELS
========================================================= */

function makeDailyLabels(
    rows
) {

    /*
       Example for July:
       2 4 6 8 10 12 14...

       Dates themselves remain daily.
       Only visible labels are reduced.
    */

    if (
        rows.length >= 20 &&
        rows.length <= 35
    ) {

        return rows.map(
            row => {

                const day =
                    row.date.getDate();


                return (
                    day % 2 === 0
                        ? String(
                            day
                        )
                        : ""
                );

            }
        );

    }


    return rows.map(
        row =>
            formatDate(
                row.date
            )
    );

}


/* =========================================================
   DAILY CHARTS
========================================================= */

function renderDailyCharts(
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
       Dashboard and Performance PR
    */

    createDailyLineChart(
        "prChart",
        labels,
        rows.map(
            row => row.pr
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
        rows.map(
            row => row.pr
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
        rows.map(
            row => row.hours
        ),
        "Operating Hours",
        "Operating Hours",
        {
            beginAtZero:
                true
        }
    );


    /*
       System Loss
    */

    createDailyLineChart(
        "lossChart",
        labels,
        rows.map(
            row => row.loss
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
        rows.map(
            row => row.loss
        ),
        "System Loss",
        "System Loss (%)",
        {
            beginAtZero:
                true
        }
    );


    /*
       Latest daily KPI card.
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
   GENERIC DAILY LINE CHART
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
            labels.length * 80
        );


    makeChartScrollable(
        canvas,
        width
    );


    const yValues =
        values.filter(
            value =>
                Number.isFinite(
                    value
                )
        );


    let yMin =
        options.min;


    let yMax =
        options.max;


    if (
        yMin === undefined &&
        yValues.length &&
        !options.beginAtZero
    ) {

        yMin =
            Math.floor(
                Math.min(
                    ...yValues
                ) * 0.95
            );

    }


    if (
        yMax === undefined &&
        yValues.length
    ) {

        yMax =
            Math.ceil(
                Math.max(
                    ...yValues
                ) * 1.05
            );

    }


    if (
        options.beginAtZero
    ) {

        yMin =
            0;

    }


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
                                    0,

                                font: {

                                    size:
                                        9

                                }

                            }

                        },


                        y: {

                            min:
                                yMin,

                            max:
                                yMax,

                            beginAtZero:
                                options.beginAtZero ||
                                false,

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


        const gantt =
            section.querySelector(
                ".chart-card.full-card"
            );


        if (gantt) {

            section.insertBefore(
                card,
                gantt
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
                                "Plant Availability",

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
        getSheetRange(
            sheet
        );


    if (!range) {
        return [];
    }


    const records = [];


    for (
        let r = 4;
        r <= range.e.r;
        r++
    ) {

        const row =
            r + 1;


        const issue =
            sheet[
                `W${row}`
            ]?.v;


        const start =
            parseTimeMinutes(
                sheet[
                    `Z${row}`
                ]?.v
            );


        const end =
            parseTimeMinutes(
                sheet[
                    `AC${row}`
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
   BREAKDOWN TIMELINE DATA
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
        getSheetRange(
            sheet
        );


    if (!range) {
        return [];
    }


    const byDate =
        new Map();


    for (
        let r = 4;
        r <= range.e.r;
        r++
    ) {

        const row =
            r + 1;


        const date =
            readDateCell(
                sheet,
                `B${row}`
            );


        const minutes =
            readNumericCell(
                sheet,
                `AG${row}`
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
            !byDate.has(
                key
            )
        ) {

            byDate.set(
                key,
                {

                    date:
                        new Date(
                            date
                        ),

                    minutes:
                        0

                }
            );

        }


        byDate.get(
            key
        ).minutes +=
            minutes;

    }


    return Array.from(
        byDate.values()
    )
        .sort(
            (a, b) =>
                a.date -
                b.date
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
                        Same-date breakdown time combined · PA Column AG
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


        const first =
            section.querySelector(
                ".chart-card.full-card"
            );


        if (first) {

            first.insertAdjacentElement(
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
   SYSTEM LOSS MWh DATA
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
        getSheetRange(
            sheet
        );


    if (!range) {
        return [];
    }


    const byDate =
        new Map();


    for (
        let r = 4;
        r <= range.e.r;
        r++
    ) {

        const row =
            r + 1;


        const date =
            readDateCell(
                sheet,
                `B${row}`
            );


        const loss =
            readNumericCell(
                sheet,
                `AL${row}`
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
            !byDate.has(
                key
            )
        ) {

            byDate.set(
                key,
                {

                    date:
                        new Date(
                            date
                        ),

                    loss:
                        0

                }
            );

        }


        byDate.get(
            key
        ).loss +=
            loss;

    }


    return Array.from(
        byDate.values()
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
            "No System Loss MWh records found."
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

    const sheet =
        getSheet(
            "Curtailment records"
        );


    if (!sheet) {
        return [];
    }


    const range =
        getSheetRange(
            sheet
        );


    if (!range) {
        return [];
    }


    const records = [];


    for (
        let r = 1;
        r <= range.e.r;
        r++
    ) {

        const row =
            r + 1;


        const date =
            readDateCell(
                sheet,
                `C${row}`
            );


        const start =
            parseTimeMinutes(
                sheet[
                    `H${row}`
                ]?.v
            );


        const end =
            parseTimeMinutes(
                sheet[
                    `I${row}`
                ]?.v
            );


        const loss =
            readNumericCell(
                sheet,
                `R${row}`
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
                loss === null
                    ? 0
                    : loss

        });

    }


    return records
        .sort(
            (a, b) =>
                a.date - b.date ||
                a.start - b.start
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
                        Loss of Generation merged by date · Column R
                    </span>

                </div>

                <span class="chart-type">
                    TABLE
                </span>

            </div>

            <div
                id="curtailmentTable"
                style="overflow-x:auto;"
            ></div>

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


            const item =
                map.get(
                    record.key
                );


            item.loss +=
                record.loss;


            item.intervals++;

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

        setText(
            "curtailmentSummary",
            "No curtailment records found"
        );

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
                        ${formatFullDate(
                            item.date
                        )}
                    </td>

                    <td style="
                        padding:10px;
                        text-align:right;
                        border-bottom:1px solid #edf2f3;
                    ">
                        ${formatNumber(
                            item.loss
                        )}
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
        `${daily.length} date(s) · ${formatNumber(total)} MWh total generation loss`
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
   Annual_KPI
   H10:H21 = Target
   I10:I21 = Measured
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


    console.log(
        "Monthly PR:",
        result
    );


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


        const grid =
            section.querySelector(
                ".three-chart-grid"
            );


        if (grid) {

            grid.insertAdjacentElement(
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
   MONTHLY PR GRAPH
========================================================= */

function renderMonthlyPRChart() {

    const records =
        readMonthlyPR();


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


    if (!records.length) {

        showCanvasMessage(
            canvas,
            "No monthly PR data found."
        );

        return;

    }


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


    const canvas =
        $("energyChart");


    if (
        !sheet ||
        !canvas
    ) {

        return;

    }


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


    const budget = [];
    const measured = [];


    for (
        let row = 10;
        row <= 21;
        row++
    ) {

        budget.push(
            readNumericCell(
                sheet,
                `E${row}`
            )
        );


        measured.push(
            readNumericCell(
                sheet,
                `F${row}`
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


    setText(
        "totalBudget",
        `${formatNumber(totalBudget)}`
    );


    setText(
        "totalMeasured",
        `${formatNumber(totalMeasured)}`
    );


    setText(
        "energyVariance",
        `${formatNumber(
            totalMeasured -
            totalBudget
        )}`
    );


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
   CANVAS MESSAGE
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
   RESET
========================================================= */

function resetDashboard() {

    workbook =
        null;


    destroyAllCharts();


    removeDynamicCards();


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


    setStatus(
        "Upload a DGR to generate the analytics."
    );

}


/* =========================================================
   RENDER EVERYTHING
========================================================= */

function renderAll() {

    if (!workbook) {
        return;
    }


    destroyAllCharts();


    removeDynamicCards();


    /*
       DAILY KPI
    */

    const daily =
        readDailyKPI();


    if (
        daily.length
    ) {

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
   END
========================================================= */
