/* =========================================================
   SOLAR DGR ANALYTICS
   app.js

   EXACT DGR MAPPINGS

   DAILY_KPI
     B  = Date
     I  = Operating Hours
     S  = PA (%)
     V  = PR (%)
     AD = System Losses (%)

   PA
     B  = Date
     W  = Issue / Fault
     Z  = Fault Start Time
     AC = Work Completion Time
     AG = Breakdown Time (minutes)
     AL = System Loss (MWh)

   CURTAILMENT RECORDS
     C  = Date
     H  = Start Time
     I  = End Time
     R  = Loss of Generation MWh

   ANNUAL_KPI
     E10:E21 = Budgeted Energy
     F10:F21 = Measured Energy

   ========================================================= */

"use strict";


/* =========================================================
   GLOBAL
========================================================= */

let workbook = null;

const chartRegistry = {};


/* =========================================================
   DOM HELPER
========================================================= */

function $(id) {
    return document.getElementById(id);
}


/* =========================================================
   PAGE INITIALISATION
========================================================= */

document.addEventListener(
    "DOMContentLoaded",
    function () {

        setupNavigation();

        setupUpload();

        setupRemoveButton();

        hideAnalytics();

    }
);


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
   UPLOAD
========================================================= */

function setupUpload() {

    const input =
        $("dgrFile");


    const dropZone =
        $("dropZone");


    if (!input) {

        console.error(
            "File input #dgrFile was not found."
        );

        return;

    }


    input.addEventListener(
        "change",
        function (event) {

            const files =
                event.target.files;


            if (
                files &&
                files.length > 0
            ) {

                processFile(
                    files[0]
                );

            }

        }
    );


    /*
       Drag and drop support.
       Works only when the old drop-zone
       is still present in your HTML.
    */

    if (dropZone) {

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


                const files =
                    event.dataTransfer.files;


                if (
                    files &&
                    files.length > 0
                ) {

                    processFile(
                        files[0]
                    );

                }

            }
        );

    }

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
            "Please upload an Excel file (.xlsx/.xls) or CSV file."
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
                            type: "array",
                            cellDates: true,
                            cellNF: true
                        }
                    );


                if (
                    !workbook.SheetNames ||
                    workbook.SheetNames.length === 0
                ) {

                    throw new Error(
                        "No worksheets were found in this file."
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
                "Could not read the selected file."
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


    if ($("fileInfo")) {

        $("fileInfo").classList.remove(
            "hidden"
        );

    }


    if ($("workbookStatus")) {

        $("workbookStatus").classList.remove(
            "hidden"
        );

    }


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
        function (name) {

            const badge =
                document.createElement(
                    "span"
                );


            badge.className =
                "sheet-badge";


            const exists =
                workbook.SheetNames.includes(
                    name
                );


            if (exists) {

                badge.textContent =
                    `${name} ✓`;

            }

            else {

                badge.textContent =
                    `${name} — missing`;

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
   GET SHEET
========================================================= */

function getSheet(
    name
) {

    if (
        !workbook ||
        !workbook.Sheets
    ) {

        return null;

    }


    if (
        workbook.Sheets[name]
    ) {

        return workbook.Sheets[name];

    }


    const target =
        name
            .toLowerCase()
            .replace(
                /[\s_-]+/g,
                ""
            );


    const actual =
        workbook.SheetNames.find(
            function (sheetName) {

                return (
                    sheetName
                        .toLowerCase()
                        .replace(
                            /[\s_-]+/g,
                            ""
                        ) ===
                    target
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
   SHEET MATRIX
========================================================= */

function sheetMatrix(
    sheet
) {

    if (!sheet) {
        return [];
    }


    return XLSX.utils.sheet_to_json(
        sheet,
        {
            header: 1,
            raw: true,
            defval: null,
            blankrows: false
        }
    );

}


/* =========================================================
   COLUMN INDEX
========================================================= */

function columnIndex(
    column
) {

    let result = 0;


    for (
        const character
        of column.toUpperCase()
    ) {

        result =
            result * 26 +
            character.charCodeAt(0) -
            64;

    }


    return result - 1;

}


/* =========================================================
   GET CELL
========================================================= */

function cell(
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
   NUMBER PARSER
========================================================= */

function numberValue(
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
        String(value)
            .trim()
            .replace(
                /,/g,
                ""
            )
            .replace(
                /%/g,
                ""
            );


    if (!text) {
        return null;
    }


    if (
        text.startsWith("#")
    ) {

        return null;

    }


    const number =
        Number(text);


    return Number.isFinite(
        number
    )
        ? number
        : null;

}


/* =========================================================
   DATE PARSER
========================================================= */

function dateValue(
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
                parsed
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

        catch (_) {

            return null;

        }

    }


    if (
        typeof value === "string"
    ) {

        const text =
            value.trim();


        /*
           DD/MM/YYYY
        */

        let match =
            text.match(
                /^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})/
            );


        if (
            match
        ) {

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


            const result =
                new Date(
                    year,
                    month,
                    day
                );


            if (
                result.getFullYear() === year &&
                result.getMonth() === month &&
                result.getDate() === day
            ) {

                return result;

            }

        }


        /*
           DD-MMM-YYYY
        */

        match =
            text.match(
                /^(\d{1,2})[\/\-]([A-Za-z]{3,9})[\/\-](\d{2,4})/
            );


        if (
            match
        ) {

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
                        .toLowerCase()
                        .substring(
                            0,
                            3
                        )
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


        const parsed =
            new Date(
                text
            );


        if (
            !isNaN(
                parsed.getTime()
            )
        ) {

            return parsed;

        }

    }


    return null;

}


/* =========================================================
   DAY KEY
========================================================= */

function dayKey(
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
   DATE FORMATTING
   ALL REQUIRED DATE FUNCTIONS LIVE HERE
========================================================= */

function formatShortDate(
    value
) {

    const date =
        value instanceof Date
            ? value
            : dateValue(
                value
            );


    if (!date) {
        return "—";
    }


    return date.toLocaleDateString(
        "en-IN",
        {
            day: "2-digit",
            month: "short"
        }
    );

}


function formatFullDate(
    value
) {

    const date =
        value instanceof Date
            ? value
            : dateValue(
                value
            );


    if (!date) {
        return "—";
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


function formatTime(
    value
) {

    if (
        value instanceof Date
    ) {

        return value.toLocaleTimeString(
            "en-IN",
            {
                hour: "2-digit",
                minute: "2-digit",
                hour12: false
            }
        );

    }


    const minutes =
        timeMinutes(
            value
        );


    if (
        minutes === null
    ) {

        return "—";

    }


    return minutesToTime(
        minutes
    );

}


/* =========================================================
   TIME PARSER
========================================================= */

function timeMinutes(
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

        /*
           Excel time fraction.
        */

        if (
            value >= 0 &&
            value < 1
        ) {

            return Math.round(
                value * 1440
            );

        }


        /*
           Excel datetime serial.
        */

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
        String(value)
            .trim();


    const match =
        text.match(
            /^(\d{1,2}):(\d{2})(?::(\d{2}))?\s*(AM|PM)?$/i
        );


    if (
        match
    ) {

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
                ?.toUpperCase();


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
            hours >= 0 &&
            hours <= 23 &&
            minutes >= 0 &&
            minutes <= 59
        ) {

            return (
                hours * 60 +
                minutes
            );

        }

    }


    return null;

}


/* =========================================================
   MINUTES TO TIME
========================================================= */

function minutesToTime(
    minutes
) {

    let value =
        Math.round(
            minutes
        );


    value =
        Math.max(
            0,
            Math.min(
                1439,
                value
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
   NUMBER DISPLAY
========================================================= */

function formatNumber(
    value,
    decimals = 2
) {

    if (
        value === null ||
        value === undefined ||
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
   PERCENT DISPLAY
========================================================= */

function displayPercent(
    value
) {

    if (
        value === null ||
        value === undefined ||
        !Number.isFinite(
            value
        )
    ) {

        return "—";

    }


    return (
        Number(value).toFixed(
            2
        ) +
        "%"
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
   DAILY KPI READER
========================================================= */

function readDailyKPI() {

    const rows =
        sheetMatrix(
            getSheet(
                "Daily_KPI"
            )
        );


    const result = [];


    rows.forEach(
        row => {

            const date =
                dateValue(
                    cell(
                        row,
                        "B"
                    )
                );


            if (!date) {
                return;
            }


            const hours =
                numberValue(
                    cell(
                        row,
                        "I"
                    )
                );


            const paRaw =
                numberValue(
                    cell(
                        row,
                        "S"
                    )
                );


            const prRaw =
                numberValue(
                    cell(
                        row,
                        "V"
                    )
                );


            const lossRaw =
                numberValue(
                    cell(
                        row,
                        "AD"
                    )
                );


            /*
               DGR stores these percentages
               as decimals.

               0.84 → 84%
               0.019 → 1.9%
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
    );


    result.sort(
        (a, b) =>
            a.date -
            b.date
    );


    return result;

}


/* =========================================================
   KPI CARDS
========================================================= */

function renderKPI(
    rows
) {

    if (!rows.length) {
        return;
    }


    const latest =
        rows[
            rows.length - 1
        ];


    setText(
        "dashboardPA",
        displayPercent(
            latest.pa
        )
    );


    setText(
        "dashboardPR",
        displayPercent(
            latest.pr
        )
    );


    setText(
        "dashboardLoss",
        displayPercent(
            latest.loss
        )
    );


    setText(
        "dashboardHours",
        latest.hours === null
            ? "—"
            : `${latest.hours.toFixed(2)} h`
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
        rows.map(
            row =>
                formatShortDate(
                    row.date
                )
        );


    /*
       PA percentage
    */

    renderPAPercentageChart(
        labels,
        rows.map(
            row =>
                row.pa
        )
    );


    /*
       PR
    */

    makeScrollableLineChart(
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

    makeScrollableLineChart(
        "hoursChart",
        labels,
        rows.map(
            row =>
                row.hours
        ),
        "Operating Hours",
        "Operating Hours"
    );


    /*
       System Loss %
    */

    makeScrollableLineChart(
        "lossChart",
        labels,
        rows.map(
            row =>
                row.loss
        ),
        "System Losses",
        "System Loss (%)"
    );


    /*
       Dashboard PR
    */

    makeScrollableLineChart(
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

    makeScrollableLineChart(
        "dashboardLossChart",
        labels,
        rows.map(
            row =>
                row.loss
        ),
        "System Losses",
        "System Loss (%)"
    );

}


/* =========================================================
   SCROLLABLE LINE CHART
========================================================= */

function makeScrollableLineChart(
    canvasId,
    labels,
    values,
    datasetLabel,
    yTitle
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


        wrapper.style.width =
            "100%";


        wrapper.style.overflowX =
            "auto";


        wrapper.style.overflowY =
            "hidden";


        wrapper.style.position =
            "relative";


        parent.insertBefore(
            wrapper,
            canvas
        );


        wrapper.appendChild(
            canvas
        );

    }


    const internalWidth =
        Math.max(
            600,
            labels.length * 58
        );


    canvas.style.width =
        `${internalWidth}px`;


    canvas.style.height =
        "100%";


    canvas.style.maxWidth =
        "none";


    chartRegistry[canvasId] =
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
                                0.25,

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

                        },


                        tooltip: {

                            callbacks: {

                                title:
                                    items =>
                                        items[0]
                                            ?.label ||
                                        ""

                            }

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
                                    12,

                                maxRotation:
                                    0,

                                minRotation:
                                    0

                            }

                        },


                        y: {

                            title: {

                                display:
                                    true,

                                text:
                                    yTitle

                            },

                            beginAtZero:
                                false,

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
   PA PERCENTAGE CHART
========================================================= */

function renderPAPercentageChart(
    labels,
    values
) {

    const section =
        $("paSection");


    if (!section) {
        return;
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
                        Daily PA (%) from Daily_KPI
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


    const canvas =
        $("paPercentageChart");


    if (!canvas) {
        return;
    }


    destroyChart(
        "paPercentageChart"
    );


    const parent =
        canvas.parentElement;


    const width =
        Math.max(
            600,
            labels.length * 58
        );


    prepareScrollableCanvas(
        canvas,
        width
    );


    chartRegistry.paPercentageChart =
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
                                    12,

                                maxRotation:
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
   PREPARE SCROLLABLE CANVAS
========================================================= */

function prepareScrollableCanvas(
    canvas,
    width
) {

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


        wrapper.style.width =
            "100%";


        wrapper.style.overflowX =
            "auto";


        wrapper.style.overflowY =
            "hidden";


        parent.insertBefore(
            wrapper,
            canvas
        );


        wrapper.appendChild(
            canvas
        );

    }


    canvas.style.width =
        `${width}px`;


    canvas.style.height =
        "100%";


    canvas.style.maxWidth =
        "none";

}


/* =========================================================
   PLANT UNAVAILABILITY
========================================================= */

function readPlantUnavailability() {

    const rows =
        sheetMatrix(
            getSheet(
                "PA"
            )
        );


    const records = [];


    rows.forEach(
        row => {

            const issue =
                cell(
                    row,
                    "W"
                );


            const start =
                timeMinutes(
                    cell(
                        row,
                        "Z"
                    )
                );


            const end =
                timeMinutes(
                    cell(
                        row,
                        "AC"
                    )
                );


            if (
                issue === null ||
                issue === undefined ||
                String(
                    issue
                ).trim() === ""
            ) {

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
                actualEnd <
                start
            ) {

                actualEnd +=
                    1440;

            }


            records.push({

                issue:
                    String(
                        issue
                    ).trim(),

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


    const width =
        Math.max(
            2304,
            records.length * 70
        );


    prepareScrollableCanvas(
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


    chartRegistry.paChart =
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

                            },

                            grid: {

                                color:
                                    "rgba(23,37,42,0.08)"

                            }

                        },


                        y: {

                            labels,

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
========================================================= */

function readBreakdownTimeline() {

    const rows =
        sheetMatrix(
            getSheet(
                "PA"
            )
        );


    const grouped =
        new Map();


    rows.forEach(
        row => {

            const date =
                dateValue(
                    cell(
                        row,
                        "B"
                    )
                );


            const breakdown =
                numberValue(
                    cell(
                        row,
                        "AG"
                    )
                );


            if (
                !date ||
                breakdown === null
            ) {

                return;

            }


            const key =
                dayKey(
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

                        total:
                            0

                    }
                );

            }


            grouped.get(
                key
            ).total +=
                breakdown;

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
            "No breakdown timeline data found."
        );

        return;

    }


    chartRegistry.breakdownChart =
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
                                        record.total
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

    const rows =
        sheetMatrix(
            getSheet(
                "PA"
            )
        );


    const grouped =
        new Map();


    rows.forEach(
        row => {

            const date =
                dateValue(
                    cell(
                        row,
                        "B"
                    )
                );


            const loss =
                numberValue(
                    cell(
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
                dayKey(
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
                        Same-date losses combined from PA · Column AL
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
            "No system loss MWh records found."
        );

        return;

    }


    const labels =
        records.map(
            record =>
                formatShortDate(
                    record.date
                )
        );


    chartRegistry.systemLossMwhChart =
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
                                "System Loss (MWh)",

                            data:
                                records.map(
                                    record =>
                                        record.loss
                                ),

                            borderWidth:
                                1,

                            borderRadius:
                                3

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

    const rows =
        sheetMatrix(
            getSheet(
                "Curtailment records"
            )
        );


    const result = [];


    rows.forEach(
        row => {

            const date =
                dateValue(
                    cell(
                        row,
                        "C"
                    )
                );


            const start =
                timeMinutes(
                    cell(
                        row,
                        "H"
                    )
                );


            const end =
                timeMinutes(
                    cell(
                        row,
                        "I"
                    )
                );


            const loss =
                numberValue(
                    cell(
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


            result.push({

                date,

                key:
                    dayKey(
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


    result.sort(
        (a, b) =>
            a.date -
            b.date ||
            a.start -
            b.start
    );


    return result;

}


/* =========================================================
   CURTAILMENT TABLE CARD
========================================================= */

function ensureCurtailmentTableCard() {

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
                        Loss of generation merged by date · Column R
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
        ensureCurtailmentTableCard();


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
                        Intervals
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


    const total =
        daily.reduce(
            (sum, record) =>
                sum +
                record.loss,
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
            "No curtailment intervals found."
        );

        return;

    }


    const parent =
        canvas.parentElement;


    const width =
        Math.max(
            2400,
            records.length * 55
        );


    prepareScrollableCanvas(
        canvas,
        width
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

                    return {

                        label:
                            `${formatShortDate(
                                record.date
                            )} ${minutesToTime(
                                record.start
                            )}`,

                        data: [

                            {

                                x: [

                                    Math.max(
                                        360,
                                        record.start
                                    ),

                                    Math.min(
                                        1080,
                                        record.end
                                    )

                                ],

                                y:
                                    formatShortDate(
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
            );


    chartRegistry.curtailmentGanttChart =
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

                            },

                            grid: {

                                color:
                                    "rgba(23,37,42,0.08)"

                            }

                        },


                        y: {

                            labels:
                                uniqueDates,

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
   ANNUAL ENERGY
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
        sheetMatrix(
            sheet
        );


    const budget = [];
    const measured = [];


    /*
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
            numberValue(
                cell(
                    row,
                    "E"
                )
            )
        );


        measured.push(
            numberValue(
                cell(
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


    setText(
        "totalBudget",
        `${formatNumber(
            budget.reduce(
                (sum, value) =>
                    sum +
                    (value || 0),
                0
            )
        )} MWh`
    );


    setText(
        "totalMeasured",
        `${formatNumber(
            measured.reduce(
                (sum, value) =>
                    sum +
                    (value || 0),
                0
            )
        )} MWh`
    );


    setText(
        "energyVariance",
        `${formatNumber(
            measured.reduce(
                (sum, value) =>
                    sum +
                    (value || 0),
                0
            ) -
            budget.reduce(
                (sum, value) =>
                    sum +
                    (value || 0),
                0
            )
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


    chartRegistry.energyChart =
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
   SHOW CANVAS MESSAGE
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


    ctx.fillStyle =
        "#879397";


    ctx.font =
        "12px Inter, Arial";


    ctx.textAlign =
        "center";


    ctx.textBaseline =
        "middle";


    ctx.fillText(
        message,
        canvas.width / 2,
        canvas.height / 2
    );


    ctx.restore();

}


/* =========================================================
   FULL RENDER
========================================================= */

function renderAll() {

    destroyAllCharts();


    /*
       Remove old dynamically generated cards.
    */

    removeDynamicCards();


    /*
       DAILY KPI
    */

    const daily =
        readDailyKPI();


    if (daily.length) {

        renderKPI(
            daily
        );


        renderDailyCharts(
            daily
        );

    }


    /*
       PA
    */

    renderPAPercentage(
        daily
    );

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
       ENERGY
    */

    renderEnergyChart();

}


/* =========================================================
   PA PERCENTAGE WRAPPER
========================================================= */

function renderPAPercentage(
    rows
) {

    if (!rows || !rows.length) {
        return;
    }


    const labels =
        rows.map(
            row =>
                formatShortDate(
                    row.date
                )
        );


    const values =
        rows.map(
            row =>
                row.pa
        );


    renderPAPercentageChart(
        labels,
        values
    );

}


/* =========================================================
   REMOVE DYNAMIC CARDS
========================================================= */

function removeDynamicCards() {

    dynamicCardIds().forEach(
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
   DYNAMIC CARD IDS
========================================================= */

function dynamicCardIds() {

    return [

        "paPercentageCard",

        "breakdownTimelineCard",

        "systemLossMwhCard",

        "curtailmentTableCard",

        "curtailmentGanttCard"

    ];

}


/* =========================================================
   DESTROY ALL CHARTS
========================================================= */

function destroyAllCharts() {

    Object.keys(
        chartRegistry
    )
    .forEach(
        id =>
            destroyChart(
                id
            )
    );

}


/* =========================================================
   RESET
========================================================= */

function resetDashboard() {

    destroyAllCharts();

    removeDynamicCards();


    workbook = null;


    if (fileInput) {

        fileInput.value =
            "";

    }


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


    hideAnalytics();


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


    if ($("sidebarFileName")) {

        $("sidebarFileName").textContent =
            "No DGR uploaded";

    }


    setStatus(
        "Upload a DGR to generate the analytics."
    );

}


/* =========================================================
   END
========================================================= */
