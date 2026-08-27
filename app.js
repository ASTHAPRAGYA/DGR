/* =========================================================
   SOLAR DGR ANALYTICS
   app.js

   ANALYSED WORKSHEETS ONLY:

   Dashboard
   Annual_KPI
   Daily_KPI
   PA
   Curtailment records

   ========================================================= */


/* =========================================================
   GLOBAL VARIABLES
========================================================= */

let workbook = null;
let sheetData = {};
let charts = {};
let uploadedFile = null;


/* =========================================================
   REQUIRED WORKSHEETS
========================================================= */

const REQUIRED_SHEETS = [
    "Dashboard",
    "Annual_KPI",
    "Daily_KPI",
    "PA",
    "Curtailment records"
];


/* =========================================================
   DOM ELEMENTS
========================================================= */

const fileInput =
    document.getElementById("dgrFile");

const dropZone =
    document.getElementById("dropZone");

const fileInfo =
    document.getElementById("fileInfo");

const fileName =
    document.getElementById("fileName");

const fileSheets =
    document.getElementById("fileSheets");

const removeFile =
    document.getElementById("removeFile");

const workbookStatus =
    document.getElementById("workbookStatus");

const sheetBadges =
    document.getElementById("sheetBadges");

const statusText =
    document.getElementById("statusText");

const sidebarFileName =
    document.getElementById("sidebarFileName");

const emptyState =
    document.getElementById("emptyState");


/* =========================================================
   INITIALISE
========================================================= */

hideAnalytics();
setupNavigation();
setupFileUpload();


/* =========================================================
   CHART DEFAULTS
========================================================= */

Chart.defaults.font.family =
    "Inter, Arial, sans-serif";

Chart.defaults.font.size = 10;

Chart.defaults.color = "#728286";


/* =========================================================
   NAVIGATION
========================================================= */

function setupNavigation() {

    const navItems =
        document.querySelectorAll(".nav-item");

    navItems.forEach(button => {

        button.addEventListener("click", () => {

            navItems.forEach(item => {
                item.classList.remove("active");
            });

            button.classList.add("active");

            const targetId =
                button.dataset.target;

            const target =
                document.getElementById(targetId);

            if (target) {

                target.scrollIntoView({
                    behavior: "smooth",
                    block: "start"
                });

            }

        });

    });

}


/* =========================================================
   FILE UPLOAD
========================================================= */

function setupFileUpload() {

    if (!fileInput || !dropZone) {
        return;
    }


    fileInput.addEventListener(
        "change",
        event => {

            const file =
                event.target.files[0];

            if (file) {
                processFile(file);
            }

        }
    );


    dropZone.addEventListener(
        "click",
        () => {

            fileInput.click();

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
                event.dataTransfer.files[0];

            if (!file) {
                return;
            }

            const extension =
                file.name
                    .split(".")
                    .pop()
                    .toLowerCase();

            if (
                !["xlsx", "xls", "csv"]
                    .includes(extension)
            ) {

                alert(
                    "Please upload an Excel or CSV file."
                );

                return;

            }

            processFile(file);

        }
    );


    if (removeFile) {

        removeFile.addEventListener(
            "click",
            resetDashboard
        );

    }

}


/* =========================================================
   PROCESS FILE
========================================================= */

function processFile(file) {

    uploadedFile = file;

    statusText.textContent =
        "Reading DGR workbook...";


    const reader =
        new FileReader();


    reader.onload =
        function(event) {

            try {

                const data =
                    new Uint8Array(
                        event.target.result
                    );


                workbook =
                    XLSX.read(
                        data,
                        {
                            type: "array",
                            cellDates: true
                        }
                    );


                readRequiredSheets();

                updateFileInformation();

                buildDashboard();

            }

            catch (error) {

                console.error(
                    "DGR ERROR:",
                    error
                );

                alert(
                    "Unable to read this DGR. Please check the Excel file."
                );

                statusText.textContent =
                    "Error reading DGR.";

            }

        };


    reader.readAsArrayBuffer(file);

}


/* =========================================================
   READ REQUIRED SHEETS
========================================================= */

function readRequiredSheets() {

    sheetData = {};


    REQUIRED_SHEETS.forEach(
        sheetName => {

            const actualSheetName =
                findSheetName(sheetName);


            if (!actualSheetName) {

                sheetData[sheetName] =
                    null;

                return;

            }


            const worksheet =
                workbook.Sheets[
                    actualSheetName
                ];


            sheetData[sheetName] =
                XLSX.utils.sheet_to_json(
                    worksheet,
                    {
                        header: 1,
                        defval: null,
                        raw: true
                    }
                );

        }
    );

}


/* =========================================================
   FIND SHEET
========================================================= */

function findSheetName(requiredName) {

    if (
        workbook.SheetNames.includes(
            requiredName
        )
    ) {

        return requiredName;

    }


    const normalized =
        normalizeText(
            requiredName
        );


    return workbook.SheetNames.find(
        name =>
            normalizeText(name) ===
            normalized
    ) || null;

}


/* =========================================================
   NORMALIZE TEXT
========================================================= */

function normalizeText(value) {

    return String(value ?? "")
        .trim()
        .toLowerCase()
        .replace(
            /[\s_-]+/g,
            ""
        );

}


/* =========================================================
   FILE INFORMATION
========================================================= */

function updateFileInformation() {

    if (fileName) {

        fileName.textContent =
            uploadedFile.name;

    }


    if (fileSheets) {

        fileSheets.textContent =
            workbook.SheetNames.length +
            " worksheets detected";

    }


    if (sidebarFileName) {

        sidebarFileName.textContent =
            uploadedFile.name;

    }


    if (fileInfo) {

        fileInfo.classList.remove(
            "hidden"
        );

    }


    if (workbookStatus) {

        workbookStatus.classList.remove(
            "hidden"
        );

    }


    if (sheetBadges) {

        sheetBadges.innerHTML = "";


        REQUIRED_SHEETS.forEach(
            sheetName => {

                const badge =
                    document.createElement(
                        "span"
                    );


                badge.className =
                    "sheet-badge";


                if (
                    sheetData[sheetName]
                ) {

                    badge.textContent =
                        sheetName;

                }

                else {

                    badge.textContent =
                        sheetName +
                        " — not found";

                    badge.classList.add(
                        "missing"
                    );

                }


                sheetBadges.appendChild(
                    badge
                );

            }
        );

    }

}


/* =========================================================
   BUILD DASHBOARD
========================================================= */

function buildDashboard() {

    emptyState.classList.add(
        "hidden"
    );


    statusText.textContent =
        "DGR loaded successfully. Analytics generated from the selected worksheets.";


    showAnalytics();


    buildPAChart();

    buildPRChart();

    buildOperatingHoursChart();

    buildSystemLossChart();

    buildCurtailmentChart();

    buildEnergyChart();

    updateKPIs();

}


/* =========================================================
   SHOW / HIDE ANALYTICS
========================================================= */

function hideAnalytics() {

    document
        .querySelectorAll(
            ".dashboard-section"
        )
        .forEach(section => {

            section.style.display =
                "none";

        });

}


function showAnalytics() {

    document
        .querySelectorAll(
            ".dashboard-section"
        )
        .forEach(section => {

            section.style.display =
                "block";

        });

}


/* =========================================================
   DESTROY CHART
========================================================= */

function destroyChart(name) {

    if (charts[name]) {

        charts[name].destroy();

        charts[name] = null;

    }

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
        typeof value === "number" &&
        Number.isFinite(value)
    ) {

        return value;

    }


    let text =
        String(value)
            .trim()
            .replace(/,/g, "");


    if (!text) {
        return null;
    }


    text =
        text.replace(
            /[^0-9.\-+]/g,
            ""
        );


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
   DATE CONVERSION
========================================================= */

function convertExcelDate(value) {

    if (
        value instanceof Date &&
        !isNaN(value)
    ) {

        return value;

    }


    if (
        typeof value === "number" &&
        value > 20000 &&
        value < 60000
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


    if (
        typeof value === "string"
    ) {

        const parsed =
            new Date(value);


        if (!isNaN(parsed)) {
            return parsed;
        }

    }


    return null;

}


/* =========================================================
   COMBINE EXCEL DATE + TIME
========================================================= */

function combineDateTime(
    dateValue,
    timeValue
) {

    const date =
        convertExcelDate(
            dateValue
        );


    if (!date) {
        return null;
    }


    /*
       Excel may return time as
       a JavaScript Date, a number,
       or a string.
    */

    let hours = 0;
    let minutes = 0;
    let seconds = 0;


    if (
        timeValue instanceof Date
    ) {

        hours =
            timeValue.getHours();

        minutes =
            timeValue.getMinutes();

        seconds =
            timeValue.getSeconds();

    }

    else if (
        typeof timeValue === "number"
    ) {

        /*
           Excel time fraction.
        */

        const totalSeconds =
            Math.round(
                timeValue * 86400
            );

        hours =
            Math.floor(
                totalSeconds / 3600
            );

        minutes =
            Math.floor(
                (
                    totalSeconds % 3600
                ) / 60
            );

        seconds =
            totalSeconds % 60;

    }

    else if (
        typeof timeValue === "string"
    ) {

        const match =
            timeValue.match(
                /(\d{1,2}):(\d{2})(?::(\d{2}))?/
            );


        if (match) {

            hours =
                Number(match[1]);

            minutes =
                Number(match[2]);

            seconds =
                Number(match[3] || 0);

        }

    }


    return new Date(
        date.getFullYear(),
        date.getMonth(),
        date.getDate(),
        hours,
        minutes,
        seconds
    );

}


/* =========================================================
   DATE FORMAT
========================================================= */

function formatDate(value) {

    if (
        !(value instanceof Date) ||
        isNaN(value)
    ) {

        return String(value ?? "");

    }


    return value.toLocaleDateString(
        "en-IN",
        {
            day: "2-digit",
            month: "short",
            year: "numeric"
        }
    );

}


/* =========================================================
   DATETIME FORMAT
========================================================= */

function formatDateTime(value) {

    if (
        !(value instanceof Date) ||
        isNaN(value)
    ) {

        return "—";

    }


    return value.toLocaleString(
        "en-IN",
        {
            day: "2-digit",
            month: "short",
            hour: "2-digit",
            minute: "2-digit"
        }
    );

}


/* =========================================================
   DAILY KPI DATA

   ACTUAL DGR:

   B = Date
   I = Operating Hours
   V = PR
   AD = System Loss

========================================================= */

function getDailyKPIData() {

    const rows =
        sheetData["Daily_KPI"];


    if (
        !rows ||
        rows.length < 5
    ) {

        return [];

    }


    const result = [];


    /*
       Actual DGR header is row 4.
       Excel row 4 = JS index 3.
       Data begins row 5.
    */

    for (
        let r = 4;
        r < rows.length;
        r++
    ) {

        const row =
            rows[r];


        if (!row) {
            continue;
        }


        const date =
            convertExcelDate(
                row[1]
            );


        const operatingHours =
            parseNumber(
                row[8]
            );


        const pr =
            parseNumber(
                row[21]
            );


        const systemLoss =
            parseNumber(
                row[29]
            );


        /*
           A row is useful if it has
           a date and at least one metric.
        */

        if (
            !date ||
            (
                operatingHours === null &&
                pr === null &&
                systemLoss === null
            )
        ) {

            continue;

        }


        result.push({

            date,

            operatingHours,

            pr,

            systemLoss

        });

    }


    return result;

}


/* =========================================================
   PR CHART
   Daily_KPI Column V
========================================================= */

function buildPRChart() {

    const data =
        getDailyKPIData();


    const valid =
        data.filter(
            item =>
                item.pr !== null
        );


    destroyChart("pr");


    const canvas =
        document.getElementById(
            "prChart"
        );


    if (!canvas) {
        return;
    }


    charts.pr =
        new Chart(
            canvas,
            {

                type: "scatter",

                data: {

                    datasets: [

                        {

                            label:
                                "Performance Ratio",

                            data:
                                valid.map(
                                    item => ({

                                        x:
                                            item.date.getTime(),

                                        y:
                                            item.pr * 100

                                    })
                                ),

                            pointRadius: 5,

                            pointHoverRadius: 8,

                            backgroundColor:
                                "#27A5AD",

                            borderColor:
                                "#27A5AD"

                        }

                    ]

                },


                options: {

                    responsive: true,

                    maintainAspectRatio:
                        false,

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

                                title:
                                    context => {

                                        if (
                                            !context.length
                                        ) {
                                            return "";
                                        }

                                        const date =
                                            new Date(
                                                context[0]
                                                    .parsed
                                                    .x
                                            );

                                        return formatDate(
                                            date
                                        );

                                    },


                                label:
                                    context => {

                                        return (
                                            "PR: " +
                                            context.parsed.y
                                                .toFixed(2) +
                                            "%"
                                        );

                                    }

                            }

                        }

                    },


                    scales: {

                        x: {

                            type: "linear",

                            title: {

                                display: true,

                                text:
                                    "Date"

                            },

                            ticks: {

                                callback:
                                    value =>
                                        formatDate(
                                            new Date(
                                                value
                                            )
                                        )

                            }

                        },


                        y: {

                            title: {

                                display: true,

                                text:
                                    "Performance Ratio (%)"

                            },

                            beginAtZero: true

                        }

                    }

                }

            }
        );


    buildDashboardPRChart(
        valid
    );

}


/* =========================================================
   DASHBOARD PR
========================================================= */

function buildDashboardPRChart(data) {

    destroyChart(
        "dashboardPR"
    );


    const canvas =
        document.getElementById(
            "dashboardPRChart"
        );


    if (!canvas) {
        return;
    }


    charts.dashboardPR =
        new Chart(
            canvas,
            {

                type: "scatter",

                data: {

                    datasets: [

                        {

                            label: "PR",

                            data:
                                data.map(
                                    item => ({

                                        x:
                                            item.date.getTime(),

                                        y:
                                            item.pr * 100

                                    })
                                ),

                            pointRadius: 4,

                            pointHoverRadius: 7,

                            backgroundColor:
                                "#27A5AD"

                        }

                    ]

                },


                options: {

                    responsive: true,

                    maintainAspectRatio:
                        false,

                    animation: false,

                    plugins: {

                        legend: {

                            display: false

                        }

                    },


                    scales: {

                        x: {

                            type: "linear",

                            title: {

                                display: true,

                                text:
                                    "Date"

                            },

                            ticks: {

                                callback:
                                    value =>
                                        formatDate(
                                            new Date(
                                                value
                                            )
                                        )

                            }

                        },


                        y: {

                            title: {

                                display: true,

                                text:
                                    "PR (%)"

                            },

                            beginAtZero: true

                        }

                    }

                }

            }

        );

}


/* =========================================================
   OPERATING HOURS
   Daily_KPI Column I
========================================================= */

function buildOperatingHoursChart() {

    const data =
        getDailyKPIData();


    const valid =
        data.filter(
            item =>
                item.operatingHours !== null
        );


    destroyChart("hours");


    const canvas =
        document.getElementById(
            "hoursChart"
        );


    if (!canvas) {
        return;
    }


    charts.hours =
        new Chart(
            canvas,
            {

                type: "line",

                data: {

                    datasets: [

                        {

                            label:
                                "Operating Hours",

                            data:
                                valid.map(
                                    item => ({

                                        x:
                                            item.date.getTime(),

                                        y:
                                            item.operatingHours

                                    })
                                ),

                            borderColor:
                                "#27A5AD",

                            backgroundColor:
                                "rgba(39,165,173,0.10)",

                            fill: true,

                            tension: 0.3,

                            pointRadius: 3,

                            pointHoverRadius: 6

                        }

                    ]

                },


                options: {

                    responsive: true,

                    maintainAspectRatio:
                        false,

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

                                title:
                                    context => {

                                        return formatDate(
                                            new Date(
                                                context[0]
                                                    .parsed
                                                    .x
                                            )
                                        );

                                    },

                                label:
                                    context =>
                                        "Operating Hours: " +
                                        context.parsed.y
                                            .toFixed(2) +
                                        " h"

                            }

                        }

                    },


                    scales: {

                        x: {

                            type: "linear",

                            title: {

                                display: true,

                                text:
                                    "Date"

                            },

                            ticks: {

                                callback:
                                    value =>
                                        formatDate(
                                            new Date(
                                                value
                                            )
                                        )

                            }

                        },


                        y: {

                            title: {

                                display: true,

                                text:
                                    "Operating Hours"

                            },

                            beginAtZero: true

                        }

                    }

                }

            }
        );

}


/* =========================================================
   SYSTEM LOSS
   Daily_KPI Column AD
========================================================= */

function buildSystemLossChart() {

    const data =
        getDailyKPIData();


    const valid =
        data.filter(
            item =>
                item.systemLoss !== null
        );


    destroyChart("loss");


    const canvas =
        document.getElementById(
            "lossChart"
        );


    if (!canvas) {
        return;
    }


    charts.loss =
        new Chart(
            canvas,
            {

                type: "line",

                data: {

                    datasets: [

                        {

                            label:
                                "System Losses",

                            data:
                                valid.map(
                                    item => ({

                                        x:
                                            item.date.getTime(),

                                        y:
                                            item.systemLoss * 100

                                    })
                                ),

                            borderColor:
                                "#27A5AD",

                            backgroundColor:
                                "rgba(39,165,173,0.12)",

                            fill: true,

                            tension: 0.25,

                            pointRadius: 3,

                            pointHoverRadius: 6

                        }

                    ]

                },


                options: {

                    responsive: true,

                    maintainAspectRatio:
                        false,

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

                                title:
                                    context =>
                                        formatDate(
                                            new Date(
                                                context[0]
                                                    .parsed
                                                    .x
                                            )
                                        ),

                                label:
                                    context =>
                                        "System Loss: " +
                                        context.parsed.y
                                            .toFixed(2) +
                                        "%"

                            }

                        }

                    },


                    scales: {

                        x: {

                            type: "linear",

                            title: {

                                display: true,

                                text:
                                    "Date"

                            },

                            ticks: {

                                callback:
                                    value =>
                                        formatDate(
                                            new Date(
                                                value
                                            )
                                        )

                            }

                        },


                        y: {

                            title: {

                                display: true,

                                text:
                                    "System Loss (%)"

                            },

                            beginAtZero: true

                        }

                    }

                }

            }
        );


    buildDashboardLossChart(
        valid
    );

}


/* =========================================================
   DASHBOARD SYSTEM LOSS
========================================================= */

function buildDashboardLossChart(data) {

    destroyChart(
        "dashboardLoss"
    );


    const canvas =
        document.getElementById(
            "dashboardLossChart"
        );


    if (!canvas) {
        return;
    }


    charts.dashboardLoss =
        new Chart(
            canvas,
            {

                type: "line",

                data: {

                    datasets: [

                        {

                            label:
                                "System Loss",

                            data:
                                data.map(
                                    item => ({

                                        x:
                                            item.date.getTime(),

                                        y:
                                            item.systemLoss * 100

                                    })
                                ),

                            borderColor:
                                "#27A5AD",

                            backgroundColor:
                                "rgba(39,165,173,0.10)",

                            fill: true,

                            tension: 0.25,

                            pointRadius: 2

                        }

                    ]

                },


                options: {

                    responsive: true,

                    maintainAspectRatio:
                        false,

                    animation: false,


                    plugins: {

                        legend: {

                            display: false

                        }

                    },


                    scales: {

                        x: {

                            type: "linear",

                            title: {

                                display: true,

                                text:
                                    "Date"

                            },

                            ticks: {

                                callback:
                                    value =>
                                        formatDate(
                                            new Date(
                                                value
                                            )
                                        )

                            }

                        },


                        y: {

                            title: {

                                display: true,

                                text:
                                    "Loss (%)"

                            },

                            beginAtZero: true

                        }

                    }

                }

            }
        );

}


/* =========================================================
   CURTAILMENT DATA

   ACTUAL DGR:

   C = Date
   H = From
   I = To
   R = Loss of Generation MWh

   We use R for actual curtailment
   generation loss.

========================================================= */

function getCurtailmentData() {

    const rows =
        sheetData[
            "Curtailment records"
        ];


    if (
        !rows ||
        rows.length < 2
    ) {

        return [];

    }


    const result = [];


    /*
       Header row = row 1.
       Data starts row 2.
    */

    for (
        let r = 1;
        r < rows.length;
        r++
    ) {

        const row =
            rows[r];


        if (!row) {
            continue;
        }


        const date =
            convertExcelDate(
                row[2]
            );


        const from =
            combineDateTime(
                row[2],
                row[7]
            );


        const to =
            combineDateTime(
                row[2],
                row[8]
            );


        const loss =
            parseNumber(
                row[17]
            );


        const category =
            String(
                row[4] ?? ""
            )
            .trim();


        /*
           Only LOAD CURTAILMENT
           records are relevant.
        */

        if (
            !category
                .toLowerCase()
                .includes(
                    "curtail"
                )
        ) {

            continue;

        }


        if (
            !date ||
            loss === null
        ) {

            continue;

        }


        result.push({

            date,

            from,

            to,

            loss,

            category,

            setPoint:
                parseNumber(
                    row[19]
                ),

            remark:
                row[18] ??
                ""

        });

    }


    /*
       Sort chronologically.
    */

    result.sort(
        (a, b) => {

            const timeA =
                a.from
                    ? a.from.getTime()
                    : a.date.getTime();

            const timeB =
                b.from
                    ? b.from.getTime()
                    : b.date.getTime();

            return timeA - timeB;

        }
    );


    return result;

}


/* =========================================================
   CURTAILMENT CHART
========================================================= */

function buildCurtailmentChart() {

    const data =
        getCurtailmentData();


    destroyChart(
        "curtailment"
    );


    const canvas =
        document.getElementById(
            "curtailmentChart"
        );


    if (!canvas) {
        return;
    }


    charts.curtailment =
        new Chart(
            canvas,
            {

                type: "line",

                data: {

                    datasets: [

                        {

                            label:
                                "Curtailment Loss",

                            data:
                                data.map(
                                    item => ({

                                        x:
                                            (
                                                item.from ||
                                                item.date
                                            ).getTime(),

                                        y:
                                            item.loss

                                    })
                                ),

                            borderColor:
                                "#27A5AD",

                            backgroundColor:
                                "rgba(39,165,173,0.10)",

                            fill: true,

                            tension: 0.15,

                            pointRadius: 2,

                            pointHoverRadius: 6

                        }

                    ]

                },


                options: {

                    responsive: true,

                    maintainAspectRatio:
                        false,

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

                                title:
                                    context => {

                                        const item =
                                            data[
                                                context[0]
                                                    .dataIndex
                                            ];

                                        return item &&
                                            item.from
                                            ? formatDateTime(
                                                item.from
                                            )
                                            : "";

                                    },


                                label:
                                    context => {

                                        const item =
                                            data[
                                                context.dataIndex
                                            ];

                                        return [
                                            "Loss of Generation: " +
                                            context.parsed.y
                                                .toFixed(2) +
                                            " MWh",

                                            item?.to
                                                ? "To: " +
                                                  formatDateTime(
                                                      item.to
                                                  )
                                                : "",

                                            item?.setPoint !==
                                                null
                                                ? "Load Shedding Set Point: " +
                                                  item.setPoint +
                                                  " MW"
                                                : ""

                                        ].filter(
                                            Boolean
                                        );

                                    }

                            }

                        }

                    },


                    scales: {

                        x: {

                            type: "linear",

                            title: {

                                display: true,

                                text:
                                    "Date / Time"

                            },

                            ticks: {

                                callback:
                                    value =>
                                        formatDateTime(
                                            new Date(
                                                value
                                            )
                                        )

                            }

                        },


                        y: {

                            title: {

                                display: true,

                                text:
                                    "Curtailment Loss (MWh)"

                            },

                            beginAtZero: true

                        }

                    }

                }

            }
        );


    const total =
        data.reduce(
            (
                sum,
                item
            ) =>
                sum + item.loss,
            0
        );


    const summary =
        document.getElementById(
            "curtailmentSummary"
        );


    if (summary) {

        summary.textContent =
            data.length +
            " curtailment records · Total generation loss: " +
            total.toFixed(2) +
            " MWh";

    }

}


/* =========================================================
   ENERGY DATA

   Annual_KPI monthwise section:

   Row 10 onward

   E = Budgeted Energy
   F = Measured Energy (MWh) (Exp)

   Actual workbook values are in MWh.

   We convert both to GWh for the chart
   so both datasets use the same unit.

========================================================= */

function getEnergyData() {

    const rows =
        sheetData[
            "Annual_KPI"
        ];


    if (
        !rows ||
        rows.length < 21
    ) {

        return [];

    }


    const result = [];


    /*
       Excel rows 10–21.
       JavaScript indexes 9–20.
    */

    for (
        let excelRow = 10;
        excelRow <= 21;
        excelRow++
    ) {

        const row =
            rows[
                excelRow - 1
            ];


        if (!row) {
            continue;
        }


        const monthDate =
            convertExcelDate(
                row[0]
            );


        const budgetMWh =
            parseNumber(
                row[4]
            );


        const measuredMWh =
            parseNumber(
                row[5]
            );


        /*
           Skip completely empty
           future rows.
        */

        if (
            budgetMWh === null &&
            measuredMWh === null
        ) {

            continue;

        }


        let month;


        if (monthDate) {

            month =
                monthDate.toLocaleDateString(
                    "en-IN",
                    {
                        month: "short",
                        year: "numeric"
                    }
                );

        }

        else {

            month =
                "Month " +
                (
                    excelRow - 9
                );

        }


        result.push({

            month,

            date:
                monthDate,

            budgetMWh,

            measuredMWh,

            budgetGWh:
                budgetMWh !== null
                    ? budgetMWh / 1000
                    : null,

            measuredGWh:
                measuredMWh !== null
                    ? measuredMWh / 1000
                    : null

        });

    }


    return result;

}


/* =========================================================
   ENERGY CHART
========================================================= */

function buildEnergyChart() {

    const data =
        getEnergyData();


    destroyChart(
        "energy"
    );


    const canvas =
        document.getElementById(
            "energyChart"
        );


    if (!canvas) {
        return;
    }


    charts.energy =
        new Chart(
            canvas,
            {

                type: "bar",

                data: {

                    labels:
                        data.map(
                            item =>
                                item.month
                        ),

                    datasets: [

                        {

                            label:
                                "Budgeted Energy",

                            data:
                                data.map(
                                    item =>
                                        item.budgetGWh
                                ),

                            backgroundColor:
                                "#17252A",

                            borderRadius: 4,

                            barPercentage:
                                0.7,

                            categoryPercentage:
                                0.75

                        },


                        {

                            label:
                                "Measured Energy",

                            data:
                                data.map(
                                    item =>
                                        item.measuredGWh
                                ),

                            backgroundColor:
                                "#27A5AD",

                            borderRadius: 4,

                            barPercentage:
                                0.7,

                            categoryPercentage:
                                0.75

                        }

                    ]

                },


                options: {

                    responsive: true,

                    maintainAspectRatio:
                        false,

                    animation: false,

                    interaction: {

                        mode: "index",

                        intersect: false

                    },


                    plugins: {

                        legend: {

                            position:
                                "top"

                        },


                        tooltip: {

                            callbacks: {

                                label:
                                    context => {

                                        return (
                                            context.dataset.label +
                                            ": " +
                                            context.parsed.y
                                                .toFixed(2) +
                                            " GWh"
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
                                    "Month"

                            },

                            grid: {

                                display: false

                            }

                        },


                        y: {

                            title: {

                                display: true,

                                text:
                                    "Energy (GWh)"

                            },

                            beginAtZero: true

                        }

                    }

                }

            }
        );


    updateEnergySummary(
        data
    );

}


/* =========================================================
   ENERGY SUMMARY
========================================================= */

function updateEnergySummary(
    data
) {

    const budget =
        data.reduce(
            (
                sum,
                item
            ) =>
                sum +
                (
                    item.budgetGWh ||
                    0
                ),
            0
        );


    const measured =
        data.reduce(
            (
                sum,
                item
            ) =>
                sum +
                (
                    item.measuredGWh ||
                    0
                ),
            0
        );


    const variance =
        measured - budget;


    const budgetElement =
        document.getElementById(
            "totalBudget"
        );


    const measuredElement =
        document.getElementById(
            "totalMeasured"
        );


    const varianceElement =
        document.getElementById(
            "energyVariance"
        );


    if (budgetElement) {

        budgetElement.textContent =
            budget.toFixed(2) +
            " GWh";

    }


    if (measuredElement) {

        measuredElement.textContent =
            measured.toFixed(2) +
            " GWh";

    }


    if (varianceElement) {

        varianceElement.textContent =
            (
                variance >= 0
                    ? "+"
                    : ""
            ) +
            variance.toFixed(2) +
            " GWh";

    }

}


/* =========================================================
   PA DATA

   ACTUAL DGR:

   B  = Date
   Z  = Fault Time
   AC = Work Completion time on fault
   AG = Breakdown Time
   AH = Action taken
   AK = Approximate Energy Loss

   We construct:

   Breakdown Start =
   Date + Fault Time

   Breakdown End =
   Date + Work Completion Time

   Duration =
   Breakdown Time (AG)

========================================================= */

function getPAData() {

    const rows =
        sheetData["PA"];


    if (
        !rows ||
        rows.length < 5
    ) {

        return [];

    }


    const result = [];


    /*
       Actual PA header is row 4.
       Data begins row 5.
    */

    for (
        let r = 4;
        r < rows.length;
        r++
    ) {

        const row =
            rows[r];


        if (!row) {
            continue;
        }


        const date =
            convertExcelDate(
                row[1]
            );


        if (!date) {
            continue;
        }


        /*
           Z = column 26
           JS index = 25
        */

        const faultTime =
            row[25];


        /*
           AC = column 29
           JS index = 28
        */

        const completionTime =
            row[28];


        /*
           AG = column 33
           JS index = 32
        */

        const breakdownTime =
            parseDurationHours(
                row[32]
            );


        const start =
            combineDateTime(
                date,
                faultTime
            );


        const end =
            combineDateTime(
                date,
                completionTime
            );


        /*
           Only include actual breakdowns
           with valid start and end.
        */

        if (
            !start ||
            !end
        ) {

            continue;

        }


        let duration =
            breakdownTime;


        /*
           If AG is unavailable,
           calculate from start/end.
        */

        if (
            duration === null
        ) {

            duration =
                (
                    end.getTime() -
                    start.getTime()
                ) /
                3600000;

        }


        /*
           Equipment hierarchy from
           actual DGR columns.
        */

        const station =
            row[11] ??
            "";


        const inverter =
            row[12] ??
            "";


        const module =
            row[13] ??
            "";


        const scb =
            row[14] ??
            "";


        const equipment =
            row[16] ??
            "";


        const affected =
            row[19] ??
            "";


        const faultCategory =
            row[22] ??
            "";


        const description =
            row[24] ??
            "";


        const action =
            row[33] ??
            "";


        const energyLoss =
            parseNumber(
                row[37]
            );


        /*
           Best display name.
        */

        let label = "";


        if (
            String(affected).trim()
        ) {

            label =
                String(affected);

        }

        else if (
            String(equipment).trim()
        ) {

            label =
                String(equipment);

        }

        else if (
            String(scb).trim()
        ) {

            label =
                String(scb);

        }

        else if (
            String(inverter).trim()
        ) {

            label =
                String(inverter);

        }

        else {

            label =
                "Breakdown";

        }


        result.push({

            date,

            start,

            end,

            duration,

            label,

            station:
                String(station),

            inverter:
                String(inverter),

            module:
                String(module),

            scb:
                String(scb),

            faultCategory:
                String(faultCategory),

            description:
                String(description),

            action:
                String(action),

            energyLoss

        });

    }


    /*
       Sort by breakdown start.
    */

    result.sort(
        (a, b) =>
            a.start.getTime() -
            b.start.getTime()
    );


    return result;

}


/* =========================================================
   PARSE PA DURATION
========================================================= */

function parseDurationHours(
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
       Excel time value.
    */

    if (
        typeof value === "number"
    ) {

        /*
           Fraction of day.
        */

        if (
            value >= 0 &&
            value < 1
        ) {

            return value * 24;

        }


        return value;

    }


    /*
       JavaScript Date containing
       a time.
    */

    if (
        value instanceof Date
    ) {

        return (
            value.getHours() * 3600 +
            value.getMinutes() * 60 +
            value.getSeconds()
        ) / 3600;

    }


    const text =
        String(value)
            .trim();


    const match =
        text.match(
            /^(\d+):(\d{1,2})(?::(\d{1,2}))?$/
        );


    if (match) {

        return (
            Number(match[1]) +
            Number(match[2]) / 60 +
            Number(match[3] || 0) / 3600
        );

    }


    return parseNumber(text);

}


/* =========================================================
   PA GANTT TIMELINE
========================================================= */

function buildPAChart() {

    const data =
        getPAData();


    destroyChart("pa");


    const canvas =
        document.getElementById(
            "paChart"
        );


    if (!canvas) {
        return;
    }


    if (
        data.length === 0
    ) {

        showPAEmptyMessage();

        return;

    }


    /*
       Give every breakdown its own
       horizontal row.
    */

    const labels =
        data.map(
            (
                item,
                index
            ) =>
                (
                    index + 1
                ) +
                ". " +
                item.label
        );


    const ganttData =
        data.map(
            item => ({

                y:
                    item.label,

                x: [

                    item.start.getTime(),

                    item.end.getTime()

                ]

            })
        );


    const minTime =
        Math.min(
            ...data.map(
                item =>
                    item.start.getTime()
            )
        );


    const maxTime =
        Math.max(
            ...data.map(
                item =>
                    item.end.getTime()
            )
        );


    /*
       Add padding around the timeline.
    */

    const padding =
        30 * 60 * 1000;


    charts.pa =
        new Chart(
            canvas,
            {

                type: "bar",

                data: {

                    labels,

                    datasets: [

                        {

                            label:
                                "Breakdown",

                            data:
                                ganttData,

                            backgroundColor:
                                "#27A5AD",

                            borderRadius: 5,

                            barThickness: 24,

                            borderSkipped: false

                        }

                    ]

                },


                options: {

                    indexAxis: "y",

                    responsive: true,

                    maintainAspectRatio:
                        false,

                    animation: false,


                    interaction: {

                        mode: "nearest",

                        intersect: true

                    },


                    scales: {

                        x: {

                            type: "linear",

                            min:
                                minTime -
                                padding,

                            max:
                                maxTime +
                                padding,

                            title: {

                                display: true,

                                text:
                                    "Breakdown Date / Time"

                            },


                            ticks: {

                                maxTicksLimit: 12,

                                callback:
                                    value =>
                                        formatDateTime(
                                            new Date(
                                                value
                                            )
                                        )

                            },


                            grid: {

                                color:
                                    "#edf2f3"

                            }

                        },


                        y: {

                            type: "category",

                            labels,

                            title: {

                                display: true,

                                text:
                                    "Affected Equipment"

                            },


                            grid: {

                                display: false

                            }

                        }

                    },


                    plugins: {

                        legend: {

                            display: false

                        },


                        tooltip: {

                            callbacks: {

                                title:
                                    context => {

                                        const item =
                                            data[
                                                context[0]
                                                    .dataIndex
                                            ];

                                        return item
                                            ? item.label
                                            : "";

                                    },


                                label:
                                    context => {

                                        const item =
                                            data[
                                                context.dataIndex
                                            ];


                                        if (!item) {
                                            return "";
                                        }


                                        const lines = [];


                                        lines.push(
                                            "Start: " +
                                            formatDateTime(
                                                item.start
                                            )
                                        );


                                        lines.push(
                                            "End: " +
                                            formatDateTime(
                                                item.end
                                            )
                                        );


                                        lines.push(
                                            "Duration: " +
                                            formatDuration(
                                                item.duration
                                            )
                                        );


                                        if (
                                            item.faultCategory
                                        ) {

                                            lines.push(
                                                "Fault: " +
                                                item.faultCategory
                                            );

                                        }


                                        if (
                                            item.energyLoss !==
                                            null
                                        ) {

                                            lines.push(
                                                "Energy Loss: " +
                                                item.energyLoss
                                                    .toFixed(2) +
                                                " kWh"
                                            );

                                        }


                                        return lines;

                                    }

                            }

                        }

                    }

                }

            }
        );


    /*
       Dynamically increase chart height
       when there are many breakdowns.
    */

    const wrapper =
        canvas.parentElement;


    if (wrapper) {

        wrapper.style.height =
            Math.max(
                420,
                data.length * 42
            ) +
            "px";

    }

}


/* =========================================================
   PA EMPTY STATE
========================================================= */

function showPAEmptyMessage() {

    const canvas =
        document.getElementById(
            "paChart"
        );


    if (!canvas) {
        return;
    }


    const ctx =
        canvas.getContext(
            "2d"
        );


    ctx.clearRect(
        0,
        0,
        canvas.width,
        canvas.height
    );


    ctx.font =
        "12px Inter";


    ctx.fillStyle =
        "#8b999c";


    ctx.textAlign =
        "center";


    ctx.fillText(
        "No valid breakdown records found.",
        canvas.width / 2,
        canvas.height / 2
    );

}


/* =========================================================
   FORMAT DURATION
========================================================= */

function formatDuration(
    hours
) {

    if (
        hours === null ||
        !Number.isFinite(hours)
    ) {

        return "—";

    }


    const minutes =
        Math.round(
            hours * 60
        );


    const h =
        Math.floor(
            minutes / 60
        );


    const m =
        minutes % 60;


    return (
        h +
        "h " +
        m +
        "m"
    );

}


/* =========================================================
   UPDATE KPI CARDS
========================================================= */

function updateKPIs() {

    const daily =
        getDailyKPIData();


    const latestPR =
        getLatest(
            daily,
            "pr"
        );


    const latestHours =
        getLatest(
            daily,
            "operatingHours"
        );


    const latestLoss =
        getLatest(
            daily,
            "systemLoss"
        );


    const prElement =
        document.getElementById(
            "dashboardPR"
        );


    const hoursElement =
        document.getElementById(
            "dashboardHours"
        );


    const lossElement =
        document.getElementById(
            "dashboardLoss"
        );


    if (prElement) {

        prElement.textContent =
            latestPR !== null
                ? (
                    latestPR * 100
                ).toFixed(2) +
                  "%"
                : "—";

    }


    if (hoursElement) {

        hoursElement.textContent =
            latestHours !== null
                ? latestHours.toFixed(2) +
                  " h"
                : "—";

    }


    if (lossElement) {

        lossElement.textContent =
            latestLoss !== null
                ? (
                    latestLoss * 100
                ).toFixed(2) +
                  "%"
                : "—";

    }


    /*
       PA card:
       use the PA breakdown records to
       calculate downtime percentage
       across the reporting period.
    */

    const pa =
        getPAData();


    const paValue =
        calculatePA(
            pa
        );


    const paElement =
        document.getElementById(
            "dashboardPA"
        );


    if (paElement) {

        paElement.textContent =
            paValue !== null
                ? paValue.toFixed(2) +
                  "%"
                : "—";

    }

}


/* =========================================================
   GET LATEST
========================================================= */

function getLatest(
    data,
    property
) {

    for (
        let i = data.length - 1;
        i >= 0;
        i--
    ) {

        const value =
            data[i][property];


        if (
            value !== null &&
            Number.isFinite(value)
        ) {

            return value;

        }

    }


    return null;

}


/* =========================================================
   CALCULATE PA

   Uses the reporting period represented
   by the PA breakdown records.

========================================================= */

function calculatePA(
    pa
) {

    if (
        !pa ||
        pa.length === 0
    ) {

        return null;

    }


    const earliest =
        Math.min(
            ...pa.map(
                item =>
                    item.start.getTime()
            )
        );


    const latest =
        Math.max(
            ...pa.map(
                item =>
                    item.end.getTime()
            )
        );


    const totalHours =
        (
            latest -
            earliest
        ) /
        3600000;


    if (
        totalHours <= 0
    ) {

        return null;

    }


    const downtime =
        pa.reduce(
            (
                sum,
                item
            ) =>
                sum +
                (
                    item.duration || 0
                ),
            0
        );


    return Math.max(
        0,
        Math.min(
            100,
            (
                (
                    totalHours -
                    downtime
                ) /
                totalHours
            ) *
            100
        )
    );

}


/* =========================================================
   RESET
========================================================= */

function resetDashboard() {

    workbook = null;

    sheetData = {};

    uploadedFile = null;


    Object.keys(charts)
        .forEach(
            key =>
                destroyChart(key)
        );


    if (fileInput) {
        fileInput.value = "";
    }


    if (fileInfo) {

        fileInfo.classList.add(
            "hidden"
        );

    }


    if (workbookStatus) {

        workbookStatus.classList.add(
            "hidden"
        );

    }


    if (emptyState) {

        emptyState.classList.remove(
            "hidden"
        );

    }


    hideAnalytics();


    if (sidebarFileName) {

        sidebarFileName.textContent =
            "No DGR uploaded";

    }


    statusText.textContent =
        "Upload a DGR to generate the analytics.";

}


/* =========================================================
   END
========================================================= */
