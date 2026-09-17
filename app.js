const SUPABASE_FUNCTION_URL =
    "https://ctihdqybwmymalpkjkov.supabase.co/functions/v1/tableia-ai";

const SUPABASE_PUBLISHABLE_KEY =
    "sb_publishable_LU68bB5ip1OEPr8DJnjJJg_G36GmXoY";

const textPrompt =
    document.getElementById("textPrompt");

const voiceButton =
    document.getElementById("voiceButton");

const textAnalyzeButton =
    document.getElementById("textAnalyzeButton");

const clearPromptButton =
    document.getElementById("clearPromptButton");

const cameraInput =
    document.getElementById("cameraInput");

const galleryInput =
    document.getElementById("galleryInput");;

const photoPreview =
    document.getElementById("photoPreview");

const statusText =
    document.getElementById("status");

const resultSection =
    document.getElementById("resultSection");

const tableContainer =
    document.getElementById("tableContainer");

const undoTableButton =
    document.getElementById("undoTableButton");

const deleteTableButton =
    document.getElementById("deleteTableButton");

const insertExcelButton =
    document.getElementById("insertExcelButton");

clearPromptButton.style.display =
    "none";

textPrompt.addEventListener(
    "input",
    function () {
clearPromptButton.style.display =
    textPrompt.value.trim() || selectedPhoto
        ? "block"
        : "none";
    }
);

let selectedPhoto = null;
let currentTableData = null;
let previousTableData = null;

// ==========================================
// DICTÉE VOCALE
// ==========================================

const SpeechRecognition =
    window.SpeechRecognition ||
    window.webkitSpeechRecognition;

if (SpeechRecognition) {

    const recognition =
        new SpeechRecognition();

    recognition.lang =
        "fr-FR";

    recognition.interimResults =
        false;

    recognition.continuous =
        false;


    voiceButton.addEventListener(
        "click",
        function () {

            voiceButton.textContent =
                "🎙 ÉCOUTE EN COURS...";

            recognition.start();
        }
    );


    recognition.addEventListener(
        "result",
        function (event) {

            const transcript =
                event.results[0][0].transcript;

            const currentText =
                textPrompt.value.trim();

            textPrompt.value =
                currentText
                    ? currentText + " " + transcript
                    : transcript;
            clearPromptButton.style.display =
    "block";

            try {
    recognition.stop();
} catch (error) {
    console.log("Micro déjà arrêté");
            }
            
            setTimeout(
    function () {

        try {
            recognition.abort();
        } catch (error) {
            console.log("Reconnaissance déjà terminée");
        }

        voiceButton.textContent =
            "🎙 DICTER MA DEMANDE";
    },
    1500
);
        }
    );


    recognition.addEventListener(
        "end",
        function () {

            voiceButton.textContent =
                "🎙 DICTER MA DEMANDE";
        }
    );


    recognition.addEventListener(
        "error",
        function () {

            voiceButton.textContent =
                "🎙 DICTER MA DEMANDE";
        }
    );

} else {

    voiceButton.disabled =
        true;

    voiceButton.textContent =
        "🎙 VOIX NON DISPONIBLE";
}

// ==========================================
// CREATION PAR TEXTE
// ==========================================

        textAnalyzeButton.addEventListener(
    "click",
    async function () {

        const prompt =
            textPrompt.value.trim();

if (currentTableData) {

    await modifyCurrentTable();

    return;
}
        
        if (!prompt && !selectedPhoto) {
            statusText.textContent =
                "Écris, dicte ou ajoute une photo.";
            return;
        }

        textAnalyzeButton.disabled =
            true;

        statusText.textContent =
            selectedPhoto
                ? "✨ Analyse du document par l'IA..."
                : "✨ Création du tableau par l'IA...";

        try {

            const body = {};

            if (prompt) {
                body.prompt =
                    prompt;
            }

            if (selectedPhoto) {

                const imageBase64 =
                    await prepareImage(
                        selectedPhoto
                    );

                body.imageBase64 =
                    imageBase64;

                body.imageMimeType =
                    "image/jpeg";

                if (!prompt) {
                    body.prompt =
                        "Analyse ce document et transforme les informations visibles en tableau TableIA structuré.";
                }
            }

            const response =
                await fetch(
                    SUPABASE_FUNCTION_URL,
                    {
                        method: "POST",

                        headers: {
                            "Content-Type":
                                "application/json",

                            "apikey":
                                SUPABASE_PUBLISHABLE_KEY,

                            "Authorization":
                                `Bearer ${SUPABASE_PUBLISHABLE_KEY}`
                        },

                        body:
                            JSON.stringify(body)
                    }
                );

            const responseText =
                await response.text();

            if (!response.ok) {
                throw new Error(
                    responseText
                );
            }

            const data =
                JSON.parse(
                    responseText
                );

            displayAiTable(
                data
            );

            statusText.textContent =
                "✅ Tableau généré avec succès.";

        } catch (error) {

            console.error(error);

            statusText.textContent =
                "❌ Impossible de générer le tableau.";
        }

        textAnalyzeButton.disabled =
            false;
    }
);

clearPromptButton.addEventListener(
    "click",
    function () {

        textPrompt.value =
            "";

selectedPhoto =
    null;

cameraInput.value =
    "";

galleryInput.value =
    "";

photoPreview.src =
    "";

photoPreview.style.display =
    "none";
        
clearPromptButton.style.display =
    "none";
        
        statusText.textContent =
            "🧹 Demande effacée.";
        setTimeout(
    function () {
        if (
            statusText.textContent ===
            "🧹 Demande effacée."
        ) {
            statusText.textContent =
                "";
        }
    },
    2000
);
    }
);

// ==========================================
// PHOTO
// ==========================================

function handleSelectedPhoto(file) {

    if (!file) {
        return;
    }

    selectedPhoto =
        file;

    const imageUrl =
        URL.createObjectURL(file);

    photoPreview.src =
        imageUrl;

    photoPreview.style.display =
        "block";

clearPromptButton.style.display =
    "block";
    
    statusText.textContent =
        "📷 Photo prise en compte. Prête pour l'analyse.";
}


cameraInput.addEventListener(
    "change",
    function () {

        handleSelectedPhoto(
            cameraInput.files[0]
        );
    }
);


galleryInput.addEventListener(
    "change",
    function () {

        handleSelectedPhoto(
            galleryInput.files[0]
        );
    }
);

deleteTableButton.addEventListener(
    "click",
    function () {

        tableContainer.innerHTML =
            "";
        
currentTableData = null;
        textAnalyzeButton.textContent =
    "✨ CRÉER LE TABLEAU";
previousTableData =
    null;
        
        resultSection.style.display =
            "none";

        insertExcelButton.disabled =
            true;

        statusText.textContent =
            "🗑 Tableau supprimé.";
setTimeout(
    function () {

        if (
            statusText.textContent ===
            "🗑 Tableau supprimé."
        ) {
            statusText.textContent =
                "";
        }
    },
    2000
);
        
    }
);
async function modifyCurrentTable() {

        if (!currentTableData) {
            statusText.textContent =
                "Aucun tableau à modifier.";
            return;
        }

previousTableData =
    JSON.parse(
        JSON.stringify(currentTableData)
    );
        
        const instruction =
    textPrompt.value.trim();

if (!instruction) {
    statusText.textContent =
        "Écris ou dicte la modification à apporter.";
    return;
}

        textAnalyzeButton.disabled =
    true;

        statusText.textContent =
            "✨ Modification du tableau par l'IA...";

        try {

            const response =
                await fetch(
                    SUPABASE_FUNCTION_URL,
                    {
                        method: "POST",

                        headers: {
                            "Content-Type":
                                "application/json",

                            "apikey":
                                SUPABASE_PUBLISHABLE_KEY,

                            "Authorization":
                                `Bearer ${SUPABASE_PUBLISHABLE_KEY}`
                        },

                        body: JSON.stringify({
                            prompt:
                                "Modifie ce tableau existant selon la demande de l'utilisateur. " +
                                "Conserve toutes les données qui ne sont pas concernées. " +
                                "Retourne le tableau complet au format TableIA.\n\n" +
                                "TABLEAU ACTUEL :\n" +
                                JSON.stringify(currentTableData) +
                                "\n\nMODIFICATION DEMANDÉE :\n" +
                                instruction
                        })
                    }
                );

            const responseText =
                await response.text();

            if (!response.ok) {
                throw new Error(
                    responseText
                );
            }

            const data =
                JSON.parse(responseText);

            displayAiTable(data);

            statusText.textContent =
                "✅ Tableau modifié avec succès.";

        } catch (error) {

            console.error(error);

            statusText.textContent =
                "❌ Impossible de modifier le tableau.";
        }

textAnalyzeButton.disabled =
    false;
    
    }

undoTableButton.addEventListener(
    "click",
    function () {

        if (!previousTableData) {
            statusText.textContent =
                "Aucune modification à annuler.";
            return;
        }

        currentTableData =
            JSON.parse(
                JSON.stringify(previousTableData)
            );

        displayAiTable(
            currentTableData
        );

        previousTableData =
            null;

        statusText.textContent =
            "↩ Dernière modification annulée.";
    }
);

// ==========================================
// INSERTION DANS EXCEL
// ==========================================

insertExcelButton.addEventListener(
    "click",
    async function () {

        if (!currentTableData) {
            statusText.textContent =
                "Aucun tableau à insérer.";
            return;
        }

        if (typeof Excel === "undefined") {
            statusText.textContent =
                "📊 TableIA doit être ouvert depuis Excel pour insérer le tableau.";
            return;
        }

        insertExcelButton.disabled =
            true;

        statusText.textContent =
            "📊 Insertion du tableau dans Excel...";

        try {

            await Excel.run(
                async function (context) {

                    const sheet =
                        context.workbook.worksheets
                            .getActiveWorksheet();

                    const headers =
                        currentTableData.columns.map(
                            function (column) {
                                return column.name || "";
                            }
                        );

                    const rows =
                        Array.isArray(currentTableData.rows)
                            ? currentTableData.rows
                            : [];

                    const values = [
                        headers,
                        ...rows.map(
                            function (row) {
                                return headers.map(
                                    function (_, index) {
                                        return row[index] ?? "";
                                    }
                                );
                            }
                        )
                    ];

                    const range =
                        sheet.getRangeByIndexes(
                            0,
                            0,
                            values.length,
                            headers.length
                        );

                    range.values =
                        values;

                    range.format.autofitColumns();
                    range.format.autofitRows();

                    await context.sync();
                }
            );

            statusText.textContent =
                "✅ Tableau inséré dans Excel.";

        } catch (error) {

            console.error(error);

            statusText.textContent =
                "❌ Impossible d'insérer le tableau dans Excel.";
        }

        insertExcelButton.disabled =
            false;
    }
);


// ==========================================
// PREPARATION IMAGE
// ==========================================

function prepareImage(
    file
) {

    return new Promise(
        function (
            resolve,
            reject
        ) {

            const reader =
                new FileReader();

            reader.onload =
                function (event) {

                    const image =
                        new Image();

                    image.onload =
                        function () {

                            const maxSize =
                                1600;

                            let width =
                                image.width;

                            let height =
                                image.height;

                            if (
                                width > maxSize ||
                                height > maxSize
                            ) {

                                const ratio =
                                    Math.min(
                                        maxSize / width,
                                        maxSize / height
                                    );

                                width =
                                    Math.round(
                                        width * ratio
                                    );

                                height =
                                    Math.round(
                                        height * ratio
                                    );
                            }

                            const canvas =
                                document.createElement(
                                    "canvas"
                                );

                            canvas.width =
                                width;

                            canvas.height =
                                height;

                            const context =
                                canvas.getContext(
                                    "2d"
                                );

                            context.drawImage(
                                image,
                                0,
                                0,
                                width,
                                height
                            );

                            const compressedImage =
                                canvas.toDataURL(
                                    "image/jpeg",
                                    0.82
                                );

                            const base64 =
                                compressedImage.split(
                                    ","
                                )[1];

                            resolve(
                                base64
                            );
                        };

                    image.onerror =
                        reject;

                    image.src =
                        event.target.result;
                };

            reader.onerror =
                reject;

            reader.readAsDataURL(
                file
            );
        }
    );
              }
// ==========================================
// AFFICHAGE TABLEAU IA
// ==========================================

function displayAiTable(data) {

    if (
        !data ||
        !Array.isArray(data.columns)
    ) {
        throw new Error(
            "Réponse IA invalide."
        );
    }

    currentTableData = data;
    textAnalyzeButton.textContent =
    "✨ MODIFIER LE TABLEAU";
    
    const table =
        document.createElement("table");

    const headerRow =
        document.createElement("tr");

    data.columns.forEach(
        function (column) {

            const th =
                document.createElement("th");

            th.textContent =
                column.name || "";

            headerRow.appendChild(th);
        }
    );

    table.appendChild(headerRow);


    const rows =
        Array.isArray(data.rows)
            ? data.rows
            : [];

    rows.forEach(
        function (row) {

            const tr =
                document.createElement("tr");

            data.columns.forEach(
                function (_, index) {

                    const td =
                        document.createElement("td");

                    td.textContent =
                        row[index] ?? "";

                    tr.appendChild(td);
                }
            );

            table.appendChild(tr);
        }
    );


    tableContainer.innerHTML =
        "";

    tableContainer.appendChild(
        table
    );

    resultSection.style.display =
        "block";

    insertExcelButton.disabled =
        false;
        }
