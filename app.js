/* =========================================================
   SOLAR DGR ANALYTICS
   app.js

   USES ONLY:

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
   INITIAL STATE
========================================================= */

hideAnalytics();

setupNavigation();

setupFileUpload();


/* =========================================================
   NAVIGATION
========================================================= */

function setupNavigation() {

    const navItems =
        document.querySelectorAll(".nav-item");


    navItems.forEach(button => {

        button.addEventListener(
            "click",
            () => {

                navItems.forEach(item => {

                    item.classList.remove("active");

                });


                button.classList.add("active");


                const targetId =
                    button.dataset.target;


                const target =
                    document.getElementById(
                        targetId
                    );


                if (target) {

                    target.scrollIntoView({
                        behavior: "smooth",
                        block: "start"
                    });

                }

            }
        );

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
                extension !== "xlsx" &&
                extension !== "xls" &&
                extension !== "csv"
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


    if (statusText) {

        statusText.textContent =
            "Reading DGR workbook...";

    }


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
                    "DGR processing error:",
                    error
                );


                alert(
                    "Unable to read this DGR. Please check that it is a valid Excel workbook."
                );


                if (statusText) {

                    statusText.textContent =
                        "Error reading DGR.";

                }

            }

        };


    reader.readAsArrayBuffer(file);

}


/* =========================================================
   READ ONLY REQUIRED SHEETS
========================================================= */

function readRequiredSheets() {

    sheetData = {};


    REQUIRED_SHEETS.forEach(
        sheetName => {

            const actualSheetName =
                findSheetName(
                    sheetName
                );


            if (!actualSheetName) {

                sheetData[sheetName] = null;

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
   FIND SHEET NAME
========================================================= */

function findSheetName(requiredName) {

    if (
        workbook &&
        workbook.SheetNames.includes(
            requiredName
        )
    ) {

        return requiredName;

    }


    const normalizedRequired =
        normalizeText(
            requiredName
        );


    if (!workbook) {
        return null;
    }


    return workbook.SheetNames.find(
        sheetName => {

            return normalizeText(
                sheetName
            ) === normalizedRequired;

        }
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

    if (!uploadedFile || !workbook) {
        return;
    }


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

    if (emptyState) {

        emptyState.classList.add(
            "hidden"
        );

    }


    if (statusText) {

        statusText.textContent =
            "DGR loaded successfully. Analytics generated from the selected worksheets.";

    }


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
   CHART DEFAULTS
========================================================= */

if (typeof Chart !== "undefined") {

    Chart.defaults.font.family =
        "Inter, Arial, sans-serif";


    Chart.defaults.font.size =
        9;


    Chart.defaults.color =
        "#728286";

}


/* =========================================================
   GENERIC CHART DESTROY
========================================================= */

function destroyChart(name) {

    if (charts[name]) {

        charts[name].destroy();

        charts[name] = null;

    }

}


/* =========================================================
   DAILY KPI DATA
=========================================================

   Daily_KPI:

   I  = operating hours
   V  = PR
   AD = system loss

   JS indexes:

   I  = 8
   V  = 21
   AD = 29

========================================================= */

function getDailyKPIData() {

    const rows =
        sheetData["Daily_KPI"];


    if (
        !rows ||
        rows.length === 0
    ) {

        return [];

    }


    const result = [];


    for (
        let i = 0;
        i < rows.length;
        i++
    ) {

        const row =
            rows[i];


        if (
            !row ||
            row.length === 0
        ) {

            continue;

        }


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


        const date =
            findDateInRow(
                row,
                i
            );


        if (
            operatingHours === null &&
            pr === null &&
            systemLoss === null
        ) {

            continue;

        }


        result.push({

            index:
                result.length + 1,

            date,

            operatingHours,

            pr,

            systemLoss

        });

    }


    /*
       Sort by actual date when dates
       are available.

       This prevents charts from
       jumping around if Excel rows
       are not perfectly ordered.
    */

    result.sort(
        (a, b) => {

            const aTime =
                a.date instanceof Date
                    ? a.date.getTime()
                    : Infinity;


            const bTime =
                b.date instanceof Date
                    ? b.date.getTime()
                    : Infinity;


            return aTime - bTime;

        }
    );


    /*
       Rebuild sequential index after sorting.
    */

    result.forEach(
        (item, index) => {

            item.index =
                index + 1;

        }
    );


    return result;

}


/* =========================================================
   FIND DATE IN ROW
========================================================= */

function findDateInRow(
    row,
    index
) {

    if (!row) {

        return null;

    }


    /*
       1. Search for JavaScript Date.
    */

    for (
        let i = 0;
        i < row.length;
        i++
    ) {

        const value =
            row[i];


        if (
            value instanceof Date &&
            !isNaN(value.getTime())
        ) {

            return value;

        }

    }


    /*
       2. Search for Excel serial date.
    */

    for (
        let i = 0;
        i < row.length;
        i++
    ) {

        const value =
            row[i];


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

    }


    /*
       3. Search for date-like strings.
    */

    for (
        let i = 0;
        i < row.length;
        i++
    ) {

        const value =
            row[i];


        if (
            typeof value === "string" &&
            looksLikeDate(value)
        ) {

            const parsed =
                parseDateString(value);


            if (parsed) {

                return parsed;

            }

        }

    }


    /*
       No valid date found.

       Return null instead of a fake
       day number.

       This is important because a fake
       "29, 31, 31, 29" axis can occur
       when invalid dates are treated
       as actual dates.
    */

    return null;

}


/* =========================================================
   DATE CHECK
========================================================= */

function looksLikeDate(value) {

    const text =
        String(value).trim();


    return (

        /^\d{1,2}[\/\-]\d{1,2}[\/\-]\d{2,4}$/
            .test(text)

        ||

        /^\d{1,2}[\/\-]\w{3,9}[\/\-]\d{2,4}$/
            .test(text)

    );

}


/* =========================================================
   PARSE DATE STRING
========================================================= */

function parseDateString(value) {

    const text =
        String(value).trim();


    /*
       DD/MM/YYYY
       DD-MM-YYYY
    */

    let match =
        text.match(
            /^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})$/
        );


    if (match) {

        const day =
            Number(match[1]);


        const month =
            Number(match[2]) - 1;


        let year =
            Number(match[3]);


        if (year < 100) {

            year += 2000;

        }


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
       DD/MMM/YYYY
    */

    match =
        text.match(
            /^(\d{1,2})[\/\-]([A-Za-z]{3,9})[\/\-](\d{2,4})$/
        );


    if (match) {

        const day =
            Number(match[1]);


        const monthNames = [

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
            monthNames.indexOf(
                match[2]
                    .toLowerCase()
                    .substring(0, 3)
            );


        let year =
            Number(match[3]);


        if (year < 100) {

            year += 2000;

        }


        if (month >= 0) {

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

    }


    return null;

}


/* =========================================================
   PARSE NUMBER
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
            .replace(
                /,/g,
                ""
            );


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


    if (
        !Number.isFinite(number)
    ) {

        return null;

    }


    return number;

}


/* =========================================================
   FORMAT DATE
========================================================= */

function formatDate(value) {

    if (
        value instanceof Date &&
        !isNaN(value.getTime())
    ) {

        return value.toLocaleDateString(
            "en-IN",
            {
                day: "2-digit",
                month: "short"
            }
        );

    }


    return String(value ?? "—");

}


/* =========================================================
   FORMAT DATETIME
========================================================= */

function formatDateTime(value) {

    if (
        !(value instanceof Date) ||
        isNaN(value.getTime())
    ) {

        return String(value ?? "—");

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
   NICE NUMBER STEP
========================================================= */

function getNiceStep(
    min,
    max,
    targetTicks = 6
) {

    if (
        !Number.isFinite(min) ||
        !Number.isFinite(max)
    ) {

        return 1;

    }


    if (max <= min) {

        return 1;

    }


    const range =
        max - min;


    const roughStep =
        range /
        targetTicks;


    const magnitude =
        Math.pow(
            10,
            Math.floor(
                Math.log10(
                    roughStep
                )
            )
        );


    const normalized =
        roughStep /
        magnitude;


    let niceNormalized;


    if (normalized <= 1) {

        niceNormalized = 1;

    }

    else if (normalized <= 2) {

        niceNormalized = 2;

    }

    else if (normalized <= 5) {

        niceNormalized = 5;

    }

    else {

        niceNormalized = 10;

    }


    return (
        niceNormalized *
        magnitude
    );

}


/* =========================================================
   NUMERIC Y-AXIS OPTIONS
========================================================= */

function getNumericAxisOptions(
    values,
    options = {}
) {

    const valid =
        values.filter(
            value =>
                Number.isFinite(value)
        );


    if (!valid.length) {

        return {};

    }


    let min =
        Math.min(
            ...valid
        );


    let max =
        Math.max(
            ...valid
        );


    /*
       Handle identical values.
    */

    if (min === max) {

        const padding =
            Math.max(
                Math.abs(min) * 0.1,
                1
            );


        min -= padding;

        max += padding;

    }

    else {

        const padding =
            (
                max - min
            ) * (
                options.paddingPercent ??
                0.10
            );


        min -= padding;

        max += padding;

    }


    if (options.beginAtZero) {

        min = 0;

    }


    min =
        Math.max(
            0,
            min
        );


    const step =
        getNiceStep(
            min,
            max,
            options.targetTicks || 6
        );


    const axisMin =
        options.beginAtZero
            ? 0
            : Math.floor(
                min / step
            ) * step;


    const axisMax =
        Math.ceil(
            max / step
        ) * step;


    return {

        min:
            axisMin,

        max:
            axisMax,

        ticks: {

            stepSize:
                step,

            maxTicksLimit:
                options.targetTicks || 7,

            callback:
                value => {

                    return Number(
                        value
                    ).toLocaleString(
                        "en-IN",
                        {
                            maximumFractionDigits:
                                options.decimals ?? 2
                        }
                    );

                }

        }

    };

}


/* =========================================================
   DATE TICK INTERVAL
========================================================= */

function getDateTickStep(
    days
) {

    /*
       Short ranges.
    */

    if (days <= 7) {

        return 1;

    }


    if (days <= 15) {

        return 2;

    }


    /*
       Normal monthly data.

       July 1–31:
       2,4,6,8,10...
    */

    if (days <= 31) {

        return 2;

    }


    if (days <= 62) {

        return 5;

    }


    if (days <= 120) {

        return 10;

    }


    if (days <= 240) {

        return 15;

    }


    return 30;

}


/* =========================================================
   DATE AXIS OPTIONS
========================================================= */

function getDateAxisOptions(
    dates
) {

    const valid =
        dates
            .filter(
                date =>
                    date instanceof Date &&
                    !isNaN(
                        date.getTime()
                    )
            )
            .sort(
                (a, b) =>
                    a.getTime() -
                    b.getTime()
            );


    if (!valid.length) {

        return {

            title: {

                display: true,

                text: "Date"

            }

        };

    }


    const first =
        valid[0];


    const last =
        valid[valid.length - 1];


    const firstDay =
        new Date(
            first.getFullYear(),
            first.getMonth(),
            first.getDate()
        );


    const lastDay =
        new Date(
            last.getFullYear(),
            last.getMonth(),
            last.getDate()
        );


    const days =
        Math.round(
            (
                lastDay.getTime() -
                firstDay.getTime()
            ) /
            86400000
        ) + 1;


    const stepDays =
        getDateTickStep(
            days
        );


    const startTime =
        firstDay.getTime();


    const endTime =
        lastDay.getTime();


    return {

        type: "linear",

        min:
            startTime,

        max:
            endTime,

        title: {

            display: true,

            text: "Date"

        },

        ticks: {

            autoSkip: false,

            maxRotation: 0,

            minRotation: 0,

            callback:
                function(value) {

                    const date =
                        new Date(value);


                    /*
                       Only show ticks on the
                       requested interval.

                       The first day can also
                       appear when appropriate.
                    */

                    const day =
                        date.getDate();


                    const dayFromStart =
                        Math.round(
                            (
                                date.getTime() -
                                firstDay.getTime()
                            ) /
                            86400000
                        );


                    if (
                        dayFromStart %
                        stepDays === 0
                    ) {

                        return String(
                            day
                        );

                    }


                    return "";

                }

        },

        grid: {

            color:
                "#edf2f3"

        }

    };

}


/* =========================================================
   BUILD PR CHART
========================================================= */

function buildPRChart() {

    const data =
        getDailyKPIData();


    const valid =
        data.filter(
            item =>
                item.pr !== null &&
                item.date instanceof Date
        );


    destroyChart("pr");


    const canvas =
        document.getElementById(
            "prChart"
        );


    if (
        !canvas ||
        typeof Chart === "undefined"
    ) {

        return;

    }


    const dates =
        valid.map(
            item =>
                item.date
        );


    const values =
        valid.map(
            item =>
                item.pr
        );


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
                                            item.pr

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
                                    tooltipPRTitle,

                                label:
                                    tooltipPRLabel

                            }

                        }

                    },


                    scales: {

                        x:
                            getDateAxisOptions(
                                dates
                            ),

                        y: {

                            title: {

                                display: true,

                                text:
                                    "PR (%)"

                            },

                            ...getNumericAxisOptions(
                                values,
                                {
                                    targetTicks: 6,
                                    decimals: 1
                                }
                            )

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
   PR TOOLTIP
========================================================= */

function tooltipPRTitle(items) {

    if (
        !items ||
        !items.length
    ) {

        return "";

    }


    const date =
        new Date(
            items[0].parsed.x
        );


    return (
        "Date — " +
        formatDate(date)
    );

}


function tooltipPRLabel(context) {

    return (
        "PR: " +
        Number(
            context.parsed.y
        ).toFixed(2) +
        "%"
    );

}


/* =========================================================
   DASHBOARD PR CHART
========================================================= */

function buildDashboardPRChart(
    data
) {

    destroyChart(
        "dashboardPR"
    );


    const canvas =
        document.getElementById(
            "dashboardPRChart"
        );


    if (
        !canvas ||
        typeof Chart === "undefined"
    ) {

        return;

    }


    const valid =
        data.filter(
            item =>
                item.pr !== null &&
                item.date instanceof Date
        );


    const values =
        valid.map(
            item =>
                item.pr
        );


    charts.dashboardPR =
        new Chart(
            canvas,
            {

                type: "scatter",

                data: {

                    datasets: [

                        {

                            label:
                                "PR",

                            data:
                                valid.map(
                                    item => ({

                                        x:
                                            item.date.getTime(),

                                        y:
                                            item.pr

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

                        x:
                            getDateAxisOptions(
                                valid.map(
                                    item =>
                                        item.date
                                )
                            ),

                        y: {

                            title: {

                                display: true,

                                text:
                                    "PR (%)"

                            },

                            ...getNumericAxisOptions(
                                values,
                                {
                                    targetTicks: 5,
                                    decimals: 1
                                }
                            )

                        }

                    }

                }

            }
        );

}


/* =========================================================
   OPERATING HOURS
========================================================= */

function buildOperatingHoursChart() {

    const data =
        getDailyKPIData();


    const valid =
        data.filter(
            item =>
                item.operatingHours !== null &&
                item.date instanceof Date
        );


    destroyChart("hours");


    const canvas =
        document.getElementById(
            "hoursChart"
        );


    if (
        !canvas ||
        typeof Chart === "undefined"
    ) {

        return;

    }


    const values =
        valid.map(
            item =>
                item.operatingHours
        );


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

                        mode: "index",

                        intersect: false

                    },


                    plugins: {

                        legend: {

                            display: false

                        }

                    },


                    scales: {

                        x:
                            getDateAxisOptions(
                                valid.map(
                                    item =>
                                        item.date
                                )
                            ),

                        y: {

                            title: {

                                display: true,

                                text:
                                    "Operating Hours"

                            },

                            ...getNumericAxisOptions(
                                values,
                                {
                                    beginAtZero: true,
                                    targetTicks: 6,
                                    decimals: 1
                                }
                            )

                        }

                    }

                }

            }
        );

}


/* =========================================================
   SYSTEM LOSS
========================================================= */

function buildSystemLossChart() {

    const data =
        getDailyKPIData();


    const valid =
        data.filter(
            item =>
                item.systemLoss !== null &&
                item.date instanceof Date
        );


    destroyChart("loss");


    const canvas =
        document.getElementById(
            "lossChart"
        );


    if (
        !canvas ||
        typeof Chart === "undefined"
    ) {

        return;

    }


    const values =
        valid.map(
            item =>
                item.systemLoss
        );


    charts.loss =
        new Chart(
            canvas,
            {

                type: "line",

                data: {

                    datasets: [

                        {

                            label:
                                "System Losses (%)",

                            data:
                                valid.map(
                                    item => ({

                                        x:
                                            item.date.getTime(),

                                        y:
                                            item.systemLoss

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

                        }

                    },


                    scales: {

                        x:
                            getDateAxisOptions(
                                valid.map(
                                    item =>
                                        item.date
                                )
                            ),

                        y: {

                            title: {

                                display: true,

                                text:
                                    "System Loss (%)"

                            },

                            ...getNumericAxisOptions(
                                values,
                                {
                                    beginAtZero: true,
                                    targetTicks: 6,
                                    decimals: 1
                                }
                            )

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

function buildDashboardLossChart(
    data
) {

    destroyChart(
        "dashboardLoss"
    );


    const canvas =
        document.getElementById(
            "dashboardLossChart"
        );


    if (
        !canvas ||
        typeof Chart === "undefined"
    ) {

        return;

    }


    const valid =
        data.filter(
            item =>
                item.systemLoss !== null &&
                item.date instanceof Date
        );


    const values =
        valid.map(
            item =>
                item.systemLoss
        );


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
                                valid.map(
                                    item => ({

                                        x:
                                            item.date.getTime(),

                                        y:
                                            item.systemLoss

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

                        x:
                            getDateAxisOptions(
                                valid.map(
                                    item =>
                                        item.date
                                )
                            ),

                        y: {

                            title: {

                                display: true,

                                text:
                                    "Loss (%)"

                            },

                            ...getNumericAxisOptions(
                                values,
                                {
                                    beginAtZero: true,
                                    targetTicks: 5,
                                    decimals: 1
                                }
                            )

                        }

                    }

                }

            }
        );

}


/* =========================================================
   CURTAILMENT DATA
========================================================= */

function getCurtailmentData() {

    const rows =
        sheetData[
            "Curtailment records"
        ];


    if (
        !rows ||
        rows.length === 0
    ) {

        return [];

    }


    const headers =
        rows[0] || [];


    const dateIndex =
        findColumn(
            headers,
            [
                "date",
                "day",
                "timestamp",
                "time",
                "start time"
            ]
        );


    let lossIndex =
        findColumn(
            headers,
            [
                "curtailment loss",
                "curtailment losses",
                "loss",
                "losses",
                "curtailed energy",
                "energy loss",
                "mwh"
            ]
        );


    if (
        lossIndex === -1
    ) {

        lossIndex =
            findLikelyNumericColumn(
                rows,
                1
            );

    }


    const result = [];


    for (
        let i = 1;
        i < rows.length;
        i++
    ) {

        const row =
            rows[i];


        if (!row) {
            continue;
        }


        const loss =
            lossIndex >= 0
                ? parseNumber(
                    row[lossIndex]
                )
                : null;


        if (
            loss === null
        ) {

            continue;

        }


        let date =
            dateIndex >= 0
                ? row[dateIndex]
                : null;


        date =
            convertExcelDate(
                date
            );


        /*
           Only retain actual dates.
           Do not manufacture dates from row numbers.
        */

        if (
            !(date instanceof Date) ||
            isNaN(date.getTime())
        ) {

            continue;

        }


        result.push({

            date,

            loss

        });

    }


    result.sort(
        (a, b) =>
            a.date.getTime() -
            b.date.getTime()
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


    if (
        !canvas ||
        typeof Chart === "undefined"
    ) {

        return;

    }


    const dates =
        data.map(
            item =>
                item.date
        );


    const values =
        data.map(
            item =>
                item.loss
        );


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
                                            item.date.getTime(),

                                        y:
                                            item.loss

                                    })
                                ),

                            borderColor:
                                "#27A5AD",

                            backgroundColor:
                                "rgba(39,165,173,0.10)",

                            fill: true,

                            tension: 0.25,

                            pointRadius: 3,

                            pointHoverRadius: 7

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

                                label:
                                    context =>
                                        "Curtailment Loss: " +
                                        Number(
                                            context.parsed.y
                                        ).toFixed(2)

                            }

                        }

                    },


                    scales: {

                        x:
                            getDateAxisOptions(
                                dates
                            ),

                        y: {

                            title: {

                                display: true,

                                text:
                                    "Curtailment Loss"

                            },

                            ...getNumericAxisOptions(
                                values,
                                {
                                    beginAtZero: true,
                                    targetTicks: 6,
                                    decimals: 1
                                }
                            )

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
                sum +
                item.loss,
            0
        );


    const summary =
        document.getElementById(
            "curtailmentSummary"
        );


    if (summary) {

        summary.textContent =
            data.length +
            " records · Total loss: " +
            total.toFixed(2);

    }

}


/* =========================================================
   ENERGY DATA
=========================================================

   Annual_KPI:

   Budgeted Energy = E10:E21
   Measured Energy = F10:F21

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


        const budget =
            parseNumber(
                row[4]
            );


        const measured =
            parseNumber(
                row[5]
            );


        let month =
            null;


        /*
           Try first four columns
           for month name.
        */

        for (
            let c = 0;
            c < 4;
            c++
        ) {

            const value =
                row[c];


            if (
                value !== null &&
                value !== undefined &&
                value !== ""
            ) {

                const text =
                    String(
                        value
                    ).trim();


                if (
                    text.length <= 15
                ) {

                    month =
                        text;

                    break;

                }

            }

        }


        if (!month) {

            month =
                "Month " +
                (
                    excelRow - 9
                );

        }


        result.push({

            month,

            budget,

            measured

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


    if (
        !canvas ||
        typeof Chart === "undefined"
    ) {

        return;

    }


    const budgetValues =
        data.map(
            item =>
                item.budget
        );


    const measuredValues =
        data.map(
            item =>
                item.measured
        );


    const allValues =
        budgetValues
            .concat(
                measuredValues
            )
            .filter(
                value =>
                    Number.isFinite(
                        value
                    )
            );


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
                                budgetValues,

                            backgroundColor:
                                "#17252A",

                            borderRadius: 4,

                            barPercentage:
                                0.65,

                            categoryPercentage:
                                0.7

                        },


                        {

                            label:
                                "Measured Energy",

                            data:
                                measuredValues,

                            backgroundColor:
                                "#27A5AD",

                            borderRadius: 4,

                            barPercentage:
                                0.65,

                            categoryPercentage:
                                0.7

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
                                "top",

                            labels: {

                                boxWidth: 10,

                                font: {

                                    size: 9

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

                                display:
                                    false

                            }

                        },

                        y: {

                            title: {

                                display: true,

                                text:
                                    "Energy"

                            },

                            ...getNumericAxisOptions(
                                allValues,
                                {
                                    beginAtZero: true,
                                    targetTicks: 6,
                                    decimals: 0
                                }
                            )

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
                    Number.isFinite(
                        item.budget
                    )
                        ? item.budget
                        : 0
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
                    Number.isFinite(
                        item.measured
                    )
                        ? item.measured
                        : 0
                ),
            0
        );


    const variance =
        measured -
        budget;


    const totalBudget =
        document.getElementById(
            "totalBudget"
        );


    const totalMeasured =
        document.getElementById(
            "totalMeasured"
        );


    const energyVariance =
        document.getElementById(
            "energyVariance"
        );


    if (totalBudget) {

        totalBudget.textContent =
            budget.toFixed(2);

    }


    if (totalMeasured) {

        totalMeasured.textContent =
            measured.toFixed(2);

    }


    if (energyVariance) {

        energyVariance.textContent =
            (
                variance >= 0
                    ? "+"
                    : ""
            ) +
            variance.toFixed(2);

    }

}


/* =========================================================
   PA DATA
========================================================= */

function getPAData() {

    const rows =
        sheetData["PA"];


    if (
        !rows ||
        rows.length < 2
    ) {

        return [];

    }


    const headers =
        rows[0] || [];


    const startIndex =
        findColumn(
            headers,
            [
                "breakdown start",
                "start time",
                "start",
                "from",
                "outage start",
                "failure start"
            ]
        );


    const endIndex =
        findColumn(
            headers,
            [
                "breakdown end",
                "end time",
                "end",
                "to",
                "outage end",
                "failure end"
            ]
        );


    const durationIndex =
        findColumn(
            headers,
            [
                "duration",
                "breakdown duration",
                "downtime",
                "outage duration"
            ]
        );


    const equipmentIndex =
        findColumn(
            headers,
            [
                "equipment",
                "asset",
                "plant",
                "device",
                "element"
            ]
        );


    const result = [];


    for (
        let i = 1;
        i < rows.length;
        i++
    ) {

        const row =
            rows[i];


        if (!row) {
            continue;
        }


        let start =
            startIndex >= 0
                ? convertExcelDate(
                    row[startIndex]
                )
                : null;


        let end =
            endIndex >= 0
                ? convertExcelDate(
                    row[endIndex]
                )
                : null;


        let duration =
            durationIndex >= 0
                ? parseDuration(
                    row[durationIndex]
                )
                : null;


        /*
           Calculate duration from
           start/end when possible.
        */

        if (
            start instanceof Date &&
            end instanceof Date
        ) {

            const difference =
                (
                    end.getTime() -
                    start.getTime()
                ) /
                3600000;


            if (difference >= 0) {

                duration =
                    difference;

            }

        }


        if (
            !(start instanceof Date) &&
            !(end instanceof Date)
        ) {

            continue;

        }


        result.push({

            equipment:
                equipmentIndex >= 0
                    ? String(
                        row[equipmentIndex] ??
                        "Breakdown"
                    )
                    : "Breakdown " +
                      i,

            start,

            end,

            duration

        });

    }


    return result;

}


/* =========================================================
   PA TIMELINE CHART
========================================================= */

function buildPAChart() {

    const data =
        getPAData();


    destroyChart(
        "pa"
    );


    const canvas =
        document.getElementById(
            "paChart"
        );


    if (
        !canvas ||
        typeof Chart === "undefined"
    ) {

        return;

    }


    const valid =
        data.filter(
            item =>
                item.start instanceof Date
        );


    if (
        valid.length === 0
    ) {

        showPAEmptyMessage();

        return;

    }


    const minTime =
        Math.min(
            ...valid.map(
                item =>
                    item.start.getTime()
            )
        );


    const maxTime =
        Math.max(
            ...valid.map(
                item => {

                    return (
                        item.end instanceof Date
                            ? item.end.getTime()
                            : item.start.getTime()
                    );

                }
            )
        );


    charts.pa =
        new Chart(
            canvas,
            {

                type: "bar",

                data: {

                    labels:
                        valid.map(
                            item =>
                                item.equipment
                        ),

                    datasets: [

                        {

                            label:
                                "Breakdown Duration",

                            data:
                                valid.map(
                                    item => {

                                        const start =
                                            item.start.getTime();


                                        const end =
                                            item.end instanceof Date
                                                ? item.end.getTime()
                                                : start;


                                        return [
                                            start,
                                            end
                                        ];

                                    }
                                ),

                            backgroundColor:
                                "#27A5AD",

                            borderRadius: 4,

                            barThickness: 22

                        }

                    ]

                },


                options: {

                    indexAxis: "y",

                    responsive: true,

                    maintainAspectRatio:
                        false,

                    animation: false,


                    scales: {

                        x: {

                            type: "linear",

                            min:
                                minTime,

                            max:
                                maxTime,

                            ticks: {

                                callback:
                                    value =>
                                        formatDateTime(
                                            new Date(
                                                value
                                            )
                                        ),

                                maxRotation: 0,

                                minRotation: 0

                            },

                            title: {

                                display: true,

                                text:
                                    "Breakdown Timeline"

                            },

                            grid: {

                                color:
                                    "#edf2f3"

                            }

                        },


                        y: {

                            title: {

                                display: true,

                                text:
                                    "Equipment"

                            },

                            grid: {

                                display:
                                    false

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
                                    tooltipPATitle,

                                label:
                                    tooltipPALabel

                            }

                        }

                    }

                }

            }
        );

}


/* =========================================================
   PA TOOLTIP
========================================================= */

function tooltipPATitle(
    context
) {

    if (
        !context ||
        !context.length
    ) {

        return "";

    }


    const index =
        context[0].dataIndex;


    const data =
        getPAData();


    const item =
        data[index];


    if (!item) {

        return "";

    }


    return item.equipment;

}


function tooltipPALabel(
    context
) {

    const index =
        context.dataIndex;


    const data =
        getPAData();


    const item =
        data[index];


    if (!item) {

        return "";

    }


    const lines = [];


    if (
        item.start instanceof Date
    ) {

        lines.push(
            "Start: " +
            formatDateTime(
                item.start
            )
        );

    }


    if (
        item.end instanceof Date
    ) {

        lines.push(
            "End: " +
            formatDateTime(
                item.end
            )
        );

    }


    if (
        item.duration !== null &&
        Number.isFinite(
            item.duration
        )
    ) {

        lines.push(
            "Duration: " +
            formatDuration(
                item.duration
            )
        );

    }


    return lines;

}


/* =========================================================
   PA EMPTY MESSAGE
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


    if (!ctx) {
        return;
    }


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
        "No valid breakdown start/end times found in PA worksheet.",
        canvas.width / 2,
        canvas.height / 2
    );

}


/* =========================================================
   FIND COLUMN
========================================================= */

function findColumn(
    headers,
    possibleNames
) {

    if (!headers) {

        return -1;

    }


    const normalizedHeaders =
        headers.map(
            header =>
                normalizeText(
                    header
                )
        );


    const normalizedNames =
        possibleNames.map(
            name =>
                normalizeText(
                    name
                )
        );


    /*
       Exact match first.
    */

    for (
        let i = 0;
        i < normalizedHeaders.length;
        i++
    ) {

        if (
            normalizedNames.includes(
                normalizedHeaders[i]
            )
        ) {

            return i;

        }

    }


    /*
       Partial match second.
    */

    for (
        let i = 0;
        i < normalizedHeaders.length;
        i++
    ) {

        for (
            const name of normalizedNames
        ) {

            if (
                normalizedHeaders[i]
                    .includes(name)
                ||
                name.includes(
                    normalizedHeaders[i]
                )
            ) {

                return i;

            }

        }

    }


    return -1;

}


/* =========================================================
   FIND LIKELY NUMERIC COLUMN
========================================================= */

function findLikelyNumericColumn(
    rows,
    startRow
) {

    if (
        !rows ||
        !rows.length
    ) {

        return -1;

    }


    const maxColumns =
        Math.max(
            ...rows.map(
                row =>
                    row
                        ? row.length
                        : 0
            )
        );


    let bestIndex =
        -1;


    let bestScore =
        0;


    for (
        let c = 0;
        c < maxColumns;
        c++
    ) {

        let numericCount =
            0;


        for (
            let r = startRow;
            r < rows.length;
            r++
        ) {

            if (
                parseNumber(
                    rows[r]?.[c]
                ) !== null
            ) {

                numericCount++;

            }

        }


        if (
            numericCount >
            bestScore
        ) {

            bestScore =
                numericCount;


            bestIndex =
                c;

        }

    }


    return bestIndex;

}


/* =========================================================
   CONVERT EXCEL DATE
========================================================= */

function convertExcelDate(
    value
) {

    if (
        value instanceof Date &&
        !isNaN(value.getTime())
    ) {

        return value;

    }


    if (
        typeof value === "number"
    ) {

        if (
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

    }


    if (
        typeof value === "string"
    ) {

        if (
            looksLikeDate(value)
        ) {

            const parsed =
                parseDateString(
                    value
                );


            if (parsed) {

                return parsed;

            }

        }


        const parsed =
            new Date(value);


        if (
            !isNaN(parsed.getTime())
        ) {

            return parsed;

        }

    }


    return null;

}


/* =========================================================
   PARSE DURATION
========================================================= */

function parseDuration(
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

        /*
           Excel time duration can be
           stored as fraction of a day.
        */

        if (
            value > 0 &&
            value < 1
        ) {

            return value * 24;

        }


        return value;

    }


    const text =
        String(value)
            .trim();


    /*
       HH:MM
       HH:MM:SS
    */

    const match =
        text.match(
            /^(\d+):(\d{1,2})(?::(\d{1,2}))?$/
        );


    if (match) {

        const hours =
            Number(
                match[1]
            );


        const minutes =
            Number(
                match[2]
            );


        const seconds =
            Number(
                match[3] || 0
            );


        return (
            hours +
            minutes / 60 +
            seconds / 3600
        );

    }


    return parseNumber(
        text
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


    const totalMinutes =
        Math.round(
            hours * 60
        );


    const h =
        Math.floor(
            totalMinutes / 60
        );


    const m =
        totalMinutes % 60;


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


    const dashboardPR =
        document.getElementById(
            "dashboardPR"
        );


    const dashboardHours =
        document.getElementById(
            "dashboardHours"
        );


    const dashboardLoss =
        document.getElementById(
            "dashboardLoss"
        );


    if (dashboardPR) {

        dashboardPR.textContent =
            latestPR !== null
                ? latestPR.toFixed(2) + "%"
                : "—";

    }


    if (dashboardHours) {

        dashboardHours.textContent =
            latestHours !== null
                ? latestHours.toFixed(2) + " h"
                : "—";

    }


    if (dashboardLoss) {

        dashboardLoss.textContent =
            latestLoss !== null
                ? latestLoss.toFixed(2) + "%"
                : "—";

    }


    /*
       PA.
    */

    const pa =
        getPAData();


    const paValue =
        calculatePA(
            pa
        );


    const dashboardPA =
        document.getElementById(
            "dashboardPA"
        );


    if (dashboardPA) {

        dashboardPA.textContent =
            paValue !== null
                ? paValue.toFixed(2) + "%"
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
========================================================= */

function calculatePA(
    pa
) {

    const valid =
        pa.filter(
            item =>
                item.start instanceof Date &&
                item.end instanceof Date &&
                item.end.getTime() >=
                    item.start.getTime()
        );


    if (
        valid.length === 0
    ) {

        return null;

    }


    const earliest =
        Math.min(
            ...valid.map(
                item =>
                    item.start.getTime()
            )
        );


    const latest =
        Math.max(
            ...valid.map(
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
        valid.reduce(
            (
                sum,
                item
            ) =>
                sum +
                (
                    Number.isFinite(
                        item.duration
                    )
                        ? item.duration
                        : 0
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
   RESET DASHBOARD
========================================================= */

function resetDashboard() {

    workbook = null;

    sheetData = {};

    uploadedFile = null;


    Object.keys(charts)
        .forEach(
            key =>
                destroyChart(
                    key
                )
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


    if (statusText) {

        statusText.textContent =
            "Upload a DGR to generate the analytics.";

    }

}


/* =========================================================
   END
========================================================= */
