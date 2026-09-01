"use strict";

console.log("SOLAR DGR APP.JS LOADED");

let workbook = null;


/* =========================================================
   DOM HELPER
========================================================= */

function $(id) {
    return document.getElementById(id);
}


/* =========================================================
   INITIALISE
========================================================= */

document.addEventListener("DOMContentLoaded", function () {

    console.log("DOM LOADED");

    setupUpload();

    setupRemove();

    setupNavigation();

    hideAnalytics();

});


/* =========================================================
   UPLOAD
========================================================= */

function setupUpload() {

    console.log("setupUpload() running");

    const input = $("dgrFile");
    const dropZone = $("dropZone");


    if (!input) {

        console.error(
            "ERROR: #dgrFile was not found in index.html"
        );

        return;
    }


    console.log("#dgrFile found");


    input.addEventListener("change", function (event) {

        console.log("FILE CHANGE EVENT FIRED");


        const file =
            event.target.files &&
            event.target.files[0];


        if (!file) {

            console.warn("No file selected.");

            return;
        }


        console.log(
            "Selected file:",
            file.name
        );


        processDGR(file);

    });


    if (!dropZone) {

        console.log("No drop zone found.");

        return;
    }


    dropZone.addEventListener("click", function () {

        input.click();

    });


    dropZone.addEventListener("dragover", function (event) {

        event.preventDefault();

        dropZone.classList.add("dragging");

    });


    dropZone.addEventListener("dragleave", function () {

        dropZone.classList.remove("dragging");

    });


    dropZone.addEventListener("drop", function (event) {

        event.preventDefault();

        dropZone.classList.remove("dragging");


        const file =
            event.dataTransfer &&
            event.dataTransfer.files &&
            event.dataTransfer.files[0];


        if (file) {

            console.log(
                "Dropped file:",
                file.name
            );


            processDGR(file);

        }

    });

}


/* =========================================================
   PROCESS DGR
========================================================= */

function processDGR(file) {

    console.log("processDGR() called:", file);


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
            "Please upload an Excel file (.xlsx, .xls or .csv)."
        );

        return;
    }


    if (typeof XLSX === "undefined") {

        console.error(
            "SheetJS (XLSX) is not loaded."
        );


        alert(
            "SheetJS is not loaded. Check index.html."
        );

        return;
    }


    setStatus(
        "Reading DGR workbook..."
    );


    const reader = new FileReader();


    reader.onload = function (event) {

        console.log("File successfully read.");


        try {

            workbook =
                XLSX.read(
                    new Uint8Array(
                        event.target.result
                    ),
                    {
                        type: "array",
                        cellDates: true,
                        cellNF: true,
                        cellText: true
                    }
                );


            console.log(
                "WORKBOOK LOADED:",
                workbook.SheetNames
            );


            if (
                !workbook ||
                !workbook.SheetNames ||
                workbook.SheetNames.length === 0
            ) {

                throw new Error(
                    "No worksheets were found in the workbook."
                );
            }


            updateFileUI(file);

            showAnalytics();


            setStatus(
                file.name +
                " loaded successfully."
            );

        }
        catch (error) {

            console.error(
                "WORKBOOK ERROR:",
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


    reader.onerror = function (error) {

        console.error(
            "FILEREADER ERROR:",
            error
        );


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
        workbook.SheetNames.length +
        " worksheets detected"
    );


    setText(
        "sidebarFileName",
        file.name
    );


    const fileInfo =
        $("fileInfo");


    if (fileInfo) {

        fileInfo.classList.remove(
            "hidden"
        );

    }


    const workbookStatus =
        $("workbookStatus");


    if (workbookStatus) {

        workbookStatus.classList.remove(
            "hidden"
        );

    }


    const emptyState =
        $("emptyState");


    if (emptyState) {

        emptyState.classList.add(
            "hidden"
        );

    }


    const dropZone =
        $("dropZone");


    if (dropZone) {

        dropZone.classList.add(
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


    container.innerHTML = "";


    const requiredSheets = [
        "Dashboard",
        "Annual_KPI",
        "Daily_KPI",
        "PA",
        "Curtailment records"
    ];


    requiredSheets.forEach(function (sheetName) {

        const badge =
            document.createElement("span");


        badge.className =
            "sheet-badge";


        const found =
            workbook.SheetNames.some(
                function (actualName) {

                    return normalizeSheet(
                        actualName
                    ) === normalizeSheet(
                        sheetName
                    );

                }
            );


        badge.textContent =
            found
                ? sheetName + " ✓"
                : sheetName + " — missing";


        if (!found) {

            badge.classList.add(
                "missing"
            );

        }


        container.appendChild(badge);

    });

}


/* =========================================================
   NORMALIZE SHEET NAME
========================================================= */

function normalizeSheet(name) {

    return String(name || "")
        .trim()
        .toLowerCase()
        .replace(/[\s_-]+/g, "");

}


/* =========================================================
   NAVIGATION
========================================================= */

function setupNavigation() {

    const buttons =
        document.querySelectorAll(
            ".nav-item"
        );


    buttons.forEach(function (button) {

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
                        behavior: "smooth",
                        block: "start"
                    });

                }

            }
        );

    });

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
        function () {

            workbook = null;


            const input =
                $("dgrFile");


            if (input) {
                input.value = "";
            }


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


            const fileInfo =
                $("fileInfo");


            if (fileInfo) {

                fileInfo.classList.add(
                    "hidden"
                );

            }


            const workbookStatus =
                $("workbookStatus");


            if (workbookStatus) {

                workbookStatus.classList.add(
                    "hidden"
                );

            }


            const dropZone =
                $("dropZone");


            if (dropZone) {

                dropZone.classList.remove(
                    "hidden"
                );

            }


            const emptyState =
                $("emptyState");


            if (emptyState) {

                emptyState.classList.remove(
                    "hidden"
                );

            }


            hideAnalytics();


            setStatus(
                "Upload a DGR to generate the analytics."
            );

        }
    );

}


/* =========================================================
   SHOW / HIDE ANALYTICS
========================================================= */

function hideAnalytics() {

    const sections = [
        "dashboardSection",
        "paSection",
        "performanceSection",
        "curtailmentSection",
        "energySection"
    ];


    sections.forEach(function (id) {

        const element =
            $(id);


        if (element) {

            element.style.display =
                "none";

        }

    });

}


function showAnalytics() {

    const sections = [
        "dashboardSection",
        "paSection",
        "performanceSection",
        "curtailmentSection",
        "energySection"
    ];


    sections.forEach(function (id) {

        const element =
            $(id);


        if (element) {

            element.style.display =
                "";

        }

    });


    const emptyState =
        $("emptyState");


    if (emptyState) {

        emptyState.classList.add(
            "hidden"
        );

    }

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
