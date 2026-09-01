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
   AG = Breakdown Time (minutes)
   AL = System Loss (MWh)

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
   GLOBALS
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
            max-width: none !important;
            box-sizing: border-box !important;
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
        function (item) {

            item.addEventListener(
                "click",
                function () {

                    items.forEach(
                        function (nav) {

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
   UPLOAD SETUP
========================================================= */

function setupUpload() {

    const input =
        $("dgrFile");


    const dropZone =
        $("dropZone");


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
                event.target.files &&
                event.target.files[0];


            if (file) {

                processDGR(
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
                event.dataTransfer &&
                event.dataTransfer.files &&
                event.dataTransfer.files[0];


            if (file) {

                processDGR(
                    file
                );

            }

        }
    );

}


/* =========================================================
   REMOVE BUTTON
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
            "Please upload a valid Excel workbook (.xlsx/.xls) or CSV file."
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
                    !workbook ||
                    !workbook.SheetNames ||
                    !workbook.SheetNames.length
                ) {

                    throw new Error(
                        "No worksheets were found in the workbook."
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


    const requiredSheets = [

        "Dashboard",
        "Annual_KPI",
        "Daily_KPI",
        "PA",
        "Curtailment records"

    ];


    container.innerHTML =
        "";


    requiredSheets.forEach(
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


    const requested =
        normalizeSheet(
            requestedName
        );


    const actual =
        workbook.SheetNames.find(
            function (sheetName) {

                return (
                    normalizeSheet(
                        sheetName
                    ) === requested
                );

            }
        );


    if (!actual) {
        return null;
    }


    return workbook.Sheets[
        actual
    ];

}


/* =========================================================
   SHEET RANGE
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


    try {

        return XLSX.utils.decode_range(
            sheet["!ref"]
        );

    }

    catch (_) {

        return null;

    }

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
        const char of String(
            letter
        ).toUpperCase()
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
   GET WORKSHEET CELL
========================================================= */

function getWorksheetCell(
    sheet,
    address
) {

    if (
        !sheet ||
        !address
    ) {

        return null;

    }


    return sheet[address] || null;

}


/* =========================================================
   READ NUMERIC CELL
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

        const displayed =
            parseNumber(
                cell.w
            );


        if (
            displayed !== null
        ) {

            return displayed;

        }

    }


    if (
        cell.v !== undefined &&
        cell.v !== null
    ) {

        return parseNumber(
            cell.v
        );

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


    const number =
        Number(
            text
        );


    return Number.isFinite(
        number
    )
        ? number
        : null;

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
       Excel decimal percentages:

       0.8420 -> 84.20
       0.0193 -> 1.93
       1      -> 100

       Already percentage values are
       left unchanged.
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
       DD/MM/YYYY or DD-MM-YYYY
    */

    let match =
        text.match(
            /^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})/
        );


    if (match) {

        const day =
            Number(
                match[1]
            );


        const month =
            Number(
                match[2]
            ) - 1;


        const year =
            Number(
                match[3]
            );


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
       Month name only
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
       For formula-generated dates,
       the displayed cached value is
       often the useful value.
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

        return parseDate(
            cell.v
        );

    }


    return null;

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


    /*
       Excel time represented as Date
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
   MINUTES TO HH:MM
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
   DATE FORMAT
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


/*
   Alias used by the restored chart sections.
*/

function shortDate(
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


    return formatDate(
        date
    );

}


/* =========================================================
   FULL DATE
========================================================= */

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
   DATE KEY
========================================================= */

function dateKey(
    date
) {

    if (
        !(date instanceof Date)
    ) {

        return "";

    }


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
   TEXT
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
            function (id) {

                destroyChart(
                    id
                );

            }
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
   SCROLLABLE CHART
========================================================= */

function prepareScrollableCanvas(
    canvas,
    desiredWidth
) {

    if (!canvas) {
        return null;
    }


    let parent =
        canvas.parentElement;


    if (!parent) {
        return null;
    }


    /*
       If canvas was previously wrapped,
       locate the existing wrapper.
    */

    let wrapper =
        parent.querySelector(
            ".dgr-scroll-wrapper"
        );


    /*
       If no wrapper exists, create one.
    */

    if (!wrapper) {

        wrapper =
            document.createElement(
                "div"
            );


        wrapper.className =
            "dgr-scroll-wrapper";


        wrapper.style.width =
            "100%";


        wrapper.style.height =
            "100%";


        wrapper.style.overflowX =
            "auto";


        wrapper.style.overflowY =
            "hidden";


        wrapper.style.position =
            "relative";


        wrapper.style.paddingBottom =
            "3px";


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


    canvas.style.display =
        "block";


    return wrapper;

}


/* =========================================================
   DAILY KPI READER
   CURRENT WORKING LOGIC
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

    let previousDate =
        null;


    /*
       Row 5 is the first data row.
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
           Formula/date fallback.
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
           Some Excel workbooks store
           formula dates as the first
           cached date repeatedly.

           If a formula exists and the
           date doesn't move forward,
           construct the expected date.
        */

        if (
            previousDate &&
            date.getTime() <=
            previousDate.getTime()
        ) {

            const cell =
                getWorksheetCell(
                    sheet,
                    `B${excelRow}`
                );


            const formula =
                cell &&
                cell.f
                    ? String(
                        cell.f
                    )
                    : "";


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


        rows.push({

            row:
                excelRow,

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
       Remove duplicate dates and retain
       the row containing the most useful data.
    */

    const unique =
        new Map();


    rows.forEach(
        function (row) {

            const key =
                dateKey(
                    row.date
                );


            if (
                !unique.has(key)
            ) {

                unique.set(
                    key,
                    row
                );

                return;

            }


            const current =
                unique.get(
                    key
                );


            const currentScore =
                Number(
                    current.hours !== null
                ) +
                Number(
                    current.pa !== null
                ) +
                Number(
                    current.pr !== null
                ) +
                Number(
                    current.loss !== null
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
                currentScore
            ) {

                unique.set(
                    key,
                    row
                );

            }

        }
    );


    const result =
        Array.from(
            unique.values()
        )
        .sort(
            function (a, b) {

                return (
                    a.date -
                    b.date
                );

            }
        );


    console.log(
        "Daily KPI:",
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

    if (
        rows.length >= 20 &&
        rows.length <= 35
    ) {

        return rows.map(
            function (row) {

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
        function (row) {

            return formatDate(
                row.date
            );

        }
    );

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
            labels.length * 80
        );


    prepareScrollableCanvas(
        canvas,
        width
    );


    const validValues =
        values.filter(
            function (value) {

                return Number.isFinite(
                    value
                );

            }
        );


    let yMin =
        options.min;


    let yMax =
        options.max;


    if (
        yMin === undefined &&
        validValues.length &&
        !options.beginAtZero
    ) {

        yMin =
            Math.floor(
                Math.min(
                    ...validValues
                ) * 0.95
            );

    }


    if (
        yMax === undefined &&
        validValues.length
    ) {

        yMax =
            Math.ceil(
                Math.max(
                    ...validValues
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
   DAILY CHARTS
========================================================= */

function renderDailyCharts(
    records
) {

    if (!records.length) {
        return;
    }


    const labels =
        makeDailyLabels(
            records
        );


    /*
       PR
    */

    createDailyLineChart(
        "prChart",
        labels,
        records.map(
            function (record) {

                return record.pr;

            }
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
            function (record) {

                return record.pr;

            }
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
            function (record) {

                return record.hours;

            }
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
        records.map(
            function (record) {

                return record.loss;

            }
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
            function (record) {

                return record.loss;

            }
        ),
        "System Loss",
        "System Loss (%)",
        {
            beginAtZero:
                true
        }
    );


    /*
       PA
    */

    renderPAPercentageChart(
        records
    );


    /*
       Latest KPI cards
    */

    renderDashboardKPI(
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
   PA PERCENTAGE CHART
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
        makeDailyLabels(
            records
        );


    const width =
        Math.max(
            card.clientWidth || 700,
            records.length * 80
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
                                    function (record) {

                                        return record.pa;

                                    }
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
            timeToMinutes(
                sheet[
                    `Z${row}`
                ]?.v
            );


        const end =
            timeToMinutes(
                sheet[
                    `AC${row}`
                ]?.v
            );


        if (
            issue === null ||
            issue === undefined
        ) {

            continue;

        }


        const name =
            String(
                issue
            )
                .trim();


        if (
            !name ||
            name === "#REF!"
        ) {

            continue;

        }


        if (
            start === null ||
            end === null
        ) {

            continue;

        }


        let actualEnd =
            end;


        if (
            actualEnd <
            start
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


    const parent =
        canvas.parentElement;


    const width =
        Math.max(
            parent?.clientWidth || 700,
            2304
        );


    prepareScrollableCanvas(
        canvas,
        width
    );


    const labels =
        records.map(
            function (record) {

                return record.issue;

            }
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

                };

            }
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
                                    function (context) {

                                        const index =
                                            context[0]
                                                ?.dataIndex;


                                        return (
                                            records[index]
                                                ?.issue ||
                                            ""
                                        );

                                    },

                                label:
                                    function (context) {

                                        const item =
                                            records[
                                                context
                                                    .dataIndex
                                            ];


                                        if (!item) {
                                            return "";
                                        }


                                        return [

                                            `Start: ${minutesToTime(
                                                item.start
                                            )}`,

                                            `End: ${minutesToTime(
                                                item.end
                                            )}`,

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
   WORKING VERSION RETAINED
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
                dateKey(
                    date
                );


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
            function (a, b) {

                return (
                    a.date -
                    b.date
                );

            }
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
                            function (record) {

                                return shortDate(
                                    record.date
                                );

                            }
                        ),

                    datasets: [

                        {

                            label:
                                "Breakdown Time (min)",

                            data:
                                records.map(
                                    function (record) {

                                        return record.minutes;

                                    }
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
                dateKey(
                    date
                );


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
            function (a, b) {

                return (
                    a.date -
                    b.date
                );

            }
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
                        Same-date system losses combined · PA Column AL
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
                            function (record) {

                                return shortDate(
                                    record.date
                                );

                            }
                        ),

                    datasets: [

                        {

                            label:
                                "System Loss (MWh)",

                            data:
                                records.map(
                                    function (record) {

                                        return record.loss;

                                    }
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
                actualEnd <
                start
            ) {

                actualEnd +=
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
                    actualEnd,

                loss:
                    loss === null
                        ? 0
                        : loss

            });

        }
    );


    return records.sort(
        function (a, b) {

            return (
                a.date - b.date ||
                a.start - b.start
            );

        }
    );

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

                return (
                    a.date -
                    b.date
                );

            }
        );


    const labels =
        daily.map(
            function (record) {

                return shortDate(
                    record.date
                );

            }
        );


    const values =
        daily.map(
            function (record) {

                return record.loss;

            }
        );


    const parent =
        canvas.parentElement;


    const width =
        Math.max(
            parent?.clientWidth || 750,
            labels.length * 70
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
                                    15,

                                maxRotation:
                                    0,

                                minRotation:
                                    0

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
        `${daily.length} date(s) · ${formatNumber(total)} MWh total generation loss`
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


            const item =
                grouped.get(
                    record.key
                );


            item.loss +=
                record.loss;


            item.intervals++;

        }
    );


    const daily =
        Array.from(
            grouped.values()
        )
        .sort(
            function (a, b) {

                return (
                    a.date -
                    b.date
                );

            }
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
        function (record) {

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
                        ${formatNumber(
                            record.loss
                        )}
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


    target.innerHTML =
        html;


    const total =
        daily.reduce(
            function (sum, item) {

                return sum + item.loss;

            },
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
            2400
        );


    prepareScrollableCanvas(
        canvas,
        width
    );


    const uniqueDates = [];


    records.forEach(
        function (record) {

            const label =
                shortDate(
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
                            `${shortDate(
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
                                                ?.y ||
                                            ""
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
   MONTHLY PR CHART
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


    const width =
        Math.max(
            card.clientWidth || 700,
            records.length * 120
        );


    prepareScrollableCanvas(
        canvas,
        width
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
                            function (record) {

                                return record.month;

                            }
                        ),

                    datasets: [

                        {

                            label:
                                "Measured PR (%)",

                            data:
                                records.map(
                                    function (record) {

                                        return record.measured;

                                    }
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
   ENERGY
   WORKING VERSION RETAINED
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

        budget.push(
            readNumericCell(
                sheet,
                `E${excelRow}`
            )
        );


        measured.push(
            readNumericCell(
                sheet,
                `F${excelRow}`
            )
        );

    }


    const totalBudget =
        budget.reduce(
            function (sum, value) {

                return (
                    sum +
                    (
                        value || 0
                    )
                );

            },
            0
        );


    const totalMeasured =
        measured.reduce(
            function (sum, value) {

                return (
                    sum +
                    (
                        value || 0
                    )
                );

            },
            0
        );


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
            totalMeasured -
            totalBudget
        )} MWh`
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
   EMPTY CANVAS MESSAGE
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

    try {

        const daily =
            readDailyKPI();


        if (
            daily.length
        ) {

            renderDailyCharts(
                daily
            );

        }

    }

    catch (error) {

        console.error(
            "Daily KPI rendering failed:",
            error
        );

    }


    /*
       PA
    */

    try {

        renderPlantUnavailability();

    }

    catch (error) {

        console.error(
            "Plant Unavailability rendering failed:",
            error
        );

    }


    try {

        renderBreakdownTimeline();

    }

    catch (error) {

        console.error(
            "Breakdown Timeline rendering failed:",
            error
        );

    }


    try {

        renderSystemLossMWh();

    }

    catch (error) {

        console.error(
            "System Loss MWh rendering failed:",
            error
        );

    }


    /*
       CURTAILMENT
    */

    try {

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

    }

    catch (error) {

        console.error(
            "Curtailment rendering failed:",
            error
        );

    }


    /*
       MONTHLY PR
    */

    try {

        renderMonthlyPRChart();

    }

    catch (error) {

        console.error(
            "Monthly PR rendering failed:",
            error
        );

    }


    /*
       ENERGY
    */

    try {

        renderEnergyChart();

    }

    catch (error) {

        console.error(
            "Energy rendering failed:",
            error
        );

    }

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
   END
========================================================= */
```
