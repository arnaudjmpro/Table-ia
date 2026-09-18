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

const isEmbeddedInOffice =
    window.self !== window.top;

let officeReadyInfo = null;
let officeReadyError = null;

let officeInitialization =
    Promise.resolve(null);

// Office est initialisé dès le chargement de la page. Ainsi, au moment où
// l'utilisateur appuie sur le bouton, la demande de permission est appelée
// directement depuis le clic, sans attente préalable qui invaliderait le geste.
if (
    isEmbeddedInOffice &&
    typeof Office !== "undefined"
) {
    officeInitialization =
        Office.onReady()
            .then(function (info) {
                officeReadyInfo =
                    info;

                return info;
            })
            .catch(function (error) {
                officeReadyError =
                    error;

                return null;
            });
}

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
    let recognition = null;
    let recognitionState =
        "idle";

    function resetVoiceButton() {
        recognitionState =
            "idle";

        recognition =
            null;

        voiceButton.disabled =
            false;

        voiceButton.textContent =
            "🎙 DICTER MA DEMANDE";
    }

    function readableError(error) {
        const details = [
            error?.name,
            error?.code,
            error?.message
        ].filter(function (value, index, values) {
            return value && values.indexOf(value) === index;
        });

        return details.length
            ? details.join(" — ")
            : String(error || "erreur inconnue");
    }

    async function authorizeMicrophoneInExcel() {
        // Hors du volet Excel, Chrome gère directement l'autorisation.
        if (window.self === window.top) {
            return true;
        }

        if (typeof Office === "undefined") {
            throw new Error(
                "La bibliothèque Office n'est pas chargée."
            );
        }

        if (
            !officeReadyInfo
        ) {
            throw new Error(
                "Excel n'a pas terminé l'initialisation de TableIA. Patientez une seconde puis réessayez. " +
                readableError(officeReadyError)
            );
        }

        if (
            officeReadyInfo.host !== Office.HostType.Excel
        ) {
            throw new Error(
                "TableIA n'est pas reconnu comme complément Excel."
            );
        }

        if (
            Office.context.platform ===
            Office.PlatformType.OfficeOnline
        ) {
            if (
                Office.context.requirements &&
                typeof Office.context.requirements.isSetSupported ===
                    "function" &&
                !Office.context.requirements.isSetSupported(
                    "DevicePermissionService",
                    "1.1"
                )
            ) {
                throw new Error(
                    "Cette version d'Excel ne prend pas en charge DevicePermissionService 1.1."
                );
            }

            if (
                !Office.devicePermission ||
                typeof Office.devicePermission.requestPermissions !==
                    "function"
            ) {
                throw new Error(
                    "L'autorisation du microphone Excel n'est pas disponible dans ce navigateur. Utilisez Excel Web dans Google Chrome ou Microsoft Edge."
                );
            }

            let permissionGrantedNow;

            try {
                permissionGrantedNow =
                    await Office.devicePermission
                        .requestPermissions([
                            Office.DevicePermissionType.microphone
                        ]);
            } catch (permissionError) {
                throw new Error(
                    "Autorisation refusée par Excel : " +
                    readableError(permissionError)
                );
            }

            if (permissionGrantedNow) {
                statusText.textContent =
                    "🎙 Micro autorisé par Excel. Rechargement de TableIA...";

                window.setTimeout(
                    function () {
                        window.location.reload();
                    },
                    500
                );

                return false;
            }
        }

        // Vérifie que le cadre Excel peut réellement ouvrir le flux audio.
        if (
            navigator.mediaDevices &&
            typeof navigator.mediaDevices.getUserMedia ===
                "function"
        ) {
            let stream;

            try {
                stream =
                    await navigator.mediaDevices.getUserMedia({
                        audio: true
                    });
            } catch (mediaError) {
                throw new Error(
                    "Excel bloque encore le microphone : " +
                    readableError(mediaError)
                );
            } finally {
                if (stream) {
                    stream.getTracks().forEach(
                        function (track) {
                            track.stop();
                        }
                    );
                }
            }
        }

        return true;
    }

    function createRecognition() {
        const newRecognition =
            new SpeechRecognition();

        newRecognition.lang =
            "fr-FR";

        newRecognition.interimResults =
            false;

        newRecognition.continuous =
            false;

        newRecognition.addEventListener(
            "start",
            function () {
                recognitionState =
                    "listening";

                voiceButton.disabled =
                    false;

                voiceButton.textContent =
                    "⏹ ARRÊTER LA DICTÉE";

                statusText.textContent =
                    "🎙 Je vous écoute...";
            }
        );

        newRecognition.addEventListener(
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

                statusText.textContent =
                    "✅ Dictée ajoutée à votre demande.";
            }
        );

        newRecognition.addEventListener(
            "end",
            function () {
                resetVoiceButton();
            }
        );

        newRecognition.addEventListener(
            "error",
            function (event) {
                if (event.error === "no-speech") {
                    statusText.textContent =
                        "Aucune parole détectée. Vous pouvez recommencer.";
                } else if (event.error !== "aborted") {
                    statusText.textContent =
                        "❌ Erreur du micro : " +
                        event.error;
                }

                recognitionState =
                    "stopping";

                voiceButton.disabled =
                    true;

                try {
                    newRecognition.abort();
                } catch (abortError) {
                    console.warn(
                        "Arrêt du micro impossible.",
                        abortError
                    );
                }

                window.setTimeout(
                    function () {
                        if (recognition === newRecognition) {
                            resetVoiceButton();
                        }
                    },
                    500
                );
            }
        );

        return newRecognition;
    }

    voiceButton.addEventListener(
        "click",
        async function () {
            if (
                recognitionState === "starting" ||
                recognitionState === "stopping"
            ) {
                return;
            }

            if (
                recognitionState === "listening" &&
                recognition
            ) {
                recognitionState =
                    "stopping";

                voiceButton.disabled =
                    true;

                voiceButton.textContent =
                    "⏳ ARRÊT EN COURS...";

                try {
                    recognition.stop();
                } catch (stopError) {
                    console.warn(
                        "Arrêt du micro impossible.",
                        stopError
                    );

                    resetVoiceButton();
                }

                return;
            }

            recognitionState =
                "starting";

            voiceButton.disabled =
                true;

            voiceButton.textContent =
                "⏳ AUTORISATION DU MICRO...";

            statusText.textContent =
                "Vérification de l'autorisation Excel...";

            try {
                const canStart =
                    await authorizeMicrophoneInExcel();

                if (!canStart) {
                    return;
                }

                recognition =
                    createRecognition();

                recognition.start();
            } catch (error) {
                console.error(error);

                statusText.textContent =
                    "❌ Micro : " +
                    readableError(error);

                if (recognition) {
                    try {
                        recognition.abort();
                    } catch (abortError) {
                        console.warn(
                            "Réinitialisation du micro impossible.",
                            abortError
                        );
                    }
                }

                resetVoiceButton();
            }
        }
    );

    if (isEmbeddedInOffice) {
        voiceButton.disabled =
            true;

        voiceButton.textContent =
            "⏳ INITIALISATION EXCEL...";

        officeInitialization.then(
            function (info) {
                if (info) {
                    resetVoiceButton();
                } else {
                    voiceButton.disabled =
                        true;

                    voiceButton.textContent =
                        "🎙 MICROPHONE INDISPONIBLE";

                    statusText.textContent =
                        "❌ Initialisation Excel impossible : " +
                        readableError(officeReadyError);
                }
            }
        );
    }

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
