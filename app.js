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

const cameraButton =
    document.getElementById("cameraButton");

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

const isMobileCaptureDevice =
    /Android|iPhone|iPad|iPod/i.test(
        navigator.userAgent
    ) ||
    (
        navigator.platform === "MacIntel" &&
        navigator.maxTouchPoints > 1
    );

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
let cameraDialog = null;

// Le lien vise un tableau précis du classeur, même après son renommage.
const EXCEL_TABLE_SETTING = "tableia.linkedTable.v1";
let linkedExcelTable = null;
let excelSyncTimer = null;
let excelSyncBusy = false;
let excelSyncGeneration = 0;
let excelPreviewIsDraft = false;
let excelDraftBase = null;
let lastExcelSnapshot = null;
let excelWriteBusy = false;
let excelAppliedFingerprint = null;

// ==========================================
// DICTÉE VOCALE
// ==========================================

const SpeechRecognition =
    window.SpeechRecognition ||
    window.webkitSpeechRecognition;

if (SpeechRecognition || isEmbeddedInOffice) {
    let recognition = null;
    let voiceDialog = null;
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

    function showTemporaryStatus(message, delay) {
        statusText.textContent =
            message;

        window.setTimeout(
            function () {
                if (statusText.textContent === message) {
                    statusText.textContent =
                        "";
                }
            },
            delay || 2000
        );
    }

    function appendTranscript(transcript) {
        const cleanTranscript =
            String(transcript || "").trim();

        if (!cleanTranscript) {
            return;
        }

        const currentText =
            textPrompt.value.trim();

        textPrompt.value =
            currentText
                ? currentText + " " + cleanTranscript
                : cleanTranscript;

        clearPromptButton.style.display =
            "block";
    }

    function openOfficeVoiceDialog() {
        if (
            typeof Office === "undefined" ||
            !Office.context ||
            !Office.context.ui ||
            typeof Office.context.ui.displayDialogAsync !== "function"
        ) {
            throw new Error(
                "La fenêtre officielle de dictée n'est pas disponible dans cette version d'Excel."
            );
        }

        const dialogUrl =
            new URL(
                "dialog-micro.html",
                window.location.href
            ).href;

        recognitionState =
            "starting";

        voiceButton.disabled =
            true;

        voiceButton.textContent =
            "⏳ OUVERTURE DE LA DICTÉE...";

        statusText.textContent =
            "Ouverture de la fenêtre sécurisée de dictée...";

        Office.context.ui.displayDialogAsync(
            dialogUrl,
            {
                height: 35,
                width: 30,
                displayInIframe: false
            },
            function (asyncResult) {
                if (
                    asyncResult.status !==
                    Office.AsyncResultStatus.Succeeded
                ) {
                    statusText.textContent =
                        "❌ Ouverture de la dictée impossible : " +
                        readableError(asyncResult.error);

                    resetVoiceButton();
                    return;
                }

                voiceDialog =
                    asyncResult.value;

                let dialogFinished =
                    false;

                recognitionState =
                    "listening";

                voiceButton.disabled =
                    true;

                voiceButton.textContent =
                    "🎙 DICTÉE OUVERTE";

                showTemporaryStatus(
                    "🎙 Dictez votre demande dans la petite fenêtre.",
                    2000
                );

                voiceDialog.addEventHandler(
                    Office.EventType.DialogMessageReceived,
                    function (event) {
                        let payload;

                        try {
                            payload =
                                JSON.parse(event.message);
                        } catch (parseError) {
                            payload = {
                                type: "dialog-error",
                                message: "Réponse de dictée invalide."
                            };
                        }

                        if (
                            payload.type ===
                            "tableia-transcript"
                        ) {
                            dialogFinished =
                                true;

                            appendTranscript(
                                payload.text
                            );

                            showTemporaryStatus(
                                "✅ Dictée ajoutée à votre demande.",
                                2000
                            );

                            voiceDialog.close();
                            voiceDialog = null;
                            resetVoiceButton();
                            return;
                        }

                        statusText.textContent =
                            "❌ Dictée : " +
                            (payload.message ||
                                "Une erreur inconnue s'est produite.");
                    }
                );

                voiceDialog.addEventHandler(
                    Office.EventType.DialogEventReceived,
                    function () {
                        voiceDialog =
                            null;

                        if (!dialogFinished) {
                            statusText.textContent =
                                "Dictée fermée sans ajouter de texte.";
                        }

                        resetVoiceButton();
                    }
                );
            }
        );
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

                showTemporaryStatus(
                    "🎙 Je vous écoute...",
                    2000
                );
            }
        );

        newRecognition.addEventListener(
            "result",
            function (event) {
                const transcript =
                    event.results[0][0].transcript;

                appendTranscript(
                    transcript
                );

                showTemporaryStatus(
                    "✅ Dictée ajoutée à votre demande.",
                    2000
                );
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

            if (isEmbeddedInOffice) {
                if (voiceDialog) {
                    return;
                }

                try {
                    openOfficeVoiceDialog();
                } catch (dialogError) {
                    console.error(dialogError);

                    statusText.textContent =
                        "❌ Dictée : " +
                        readableError(dialogError);

                    resetVoiceButton();
                }

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
                "⏳ DÉMARRAGE DU MICRO...";

            statusText.textContent =
                "Démarrage de la dictée...";

            try {
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

function dataUrlToPhotoFile(dataUrl) {
    const parts =
        String(dataUrl || "").split(",");

    if (
        parts.length !== 2 ||
        !parts[0].startsWith("data:image/")
    ) {
        throw new Error(
            "La photo reçue n'est pas valide."
        );
    }

    const mimeMatch =
        parts[0].match(/^data:(image\/[a-zA-Z0-9.+-]+);base64$/);

    if (!mimeMatch) {
        throw new Error(
            "Le format de la photo n'est pas reconnu."
        );
    }

    const binary =
        atob(parts[1]);

    const bytes =
        new Uint8Array(binary.length);

    for (
        let index = 0;
        index < binary.length;
        index += 1
    ) {
        bytes[index] =
            binary.charCodeAt(index);
    }

    return new File(
        [bytes],
        "photo-tableia.jpg",
        {
            type: mimeMatch[1],
            lastModified: Date.now()
        }
    );
}

function openOfficeCameraDialog() {
    if (
        typeof Office === "undefined" ||
        !Office.context ||
        !Office.context.ui ||
        typeof Office.context.ui.displayDialogAsync !== "function"
    ) {
        throw new Error(
            "La fenêtre sécurisée de prise de photo n'est pas disponible dans cette version d'Excel."
        );
    }

    if (cameraDialog) {
        return;
    }

    const dialogUrl =
        new URL(
            "dialog-camera.html",
            window.location.href
        ).href;

    statusText.textContent =
        "Ouverture de la caméra du Mac...";

    Office.context.ui.displayDialogAsync(
        dialogUrl,
        {
            height: 48,
            width: 38,
            displayInIframe: false
        },
        function (asyncResult) {
            if (
                asyncResult.status !==
                Office.AsyncResultStatus.Succeeded
            ) {
                statusText.textContent =
                    "❌ Ouverture de la caméra impossible : " +
                    (
                        asyncResult.error?.message ||
                        "erreur inconnue"
                    );
                return;
            }

            const activeDialog =
                asyncResult.value;

            cameraDialog =
                activeDialog;

            let photoReceived =
                false;

            statusText.textContent =
                "📷 Activez la caméra dans la petite fenêtre.";

            activeDialog.addEventHandler(
                Office.EventType.DialogMessageReceived,
                function (event) {
                    try {
                        const payload =
                            JSON.parse(event.message);

                        if (
                            payload.type ===
                            "tableia-camera-error"
                        ) {
                            throw new Error(
                                payload.message ||
                                "La caméra n'est pas disponible."
                            );
                        }

                        if (
                            payload.type !==
                            "tableia-camera-photo"
                        ) {
                            throw new Error(
                                "Réponse de la caméra non reconnue."
                            );
                        }

                        const photoFile =
                            dataUrlToPhotoFile(
                                payload.dataUrl
                            );

                        photoReceived =
                            true;

                        handleSelectedPhoto(
                            photoFile
                        );

                        activeDialog.close();

                        if (
                            cameraDialog ===
                            activeDialog
                        ) {
                            cameraDialog =
                                null;
                        }
                    } catch (error) {
                        statusText.textContent =
                            "❌ Photo : " +
                            (
                                error?.message ||
                                "réception impossible"
                            );
                    }
                }
            );

            activeDialog.addEventHandler(
                Office.EventType.DialogEventReceived,
                function () {
                    if (
                        cameraDialog ===
                        activeDialog
                    ) {
                        cameraDialog =
                            null;
                    }

                    if (!photoReceived) {
                        statusText.textContent =
                            "Prise de photo fermée sans ajouter d'image.";
                    }
                }
            );
        }
    );
}

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


cameraButton.addEventListener(
    "click",
    function (event) {
        // Sur téléphone et tablette, le champ capture natif reste
        // la solution la plus directe et déjà validée.
        if (
            isMobileCaptureDevice ||
            !isEmbeddedInOffice
        ) {
            return;
        }

        // Sur Mac et PC, l'attribut capture ouvre généralement le
        // sélecteur de fichiers. Une vraie fenêtre webcam est donc utilisée.
        event.preventDefault();

        try {
            openOfficeCameraDialog();
        } catch (error) {
            console.error(error);

            statusText.textContent =
                "❌ Photo : " +
                (
                    error?.message ||
                    "ouverture impossible"
                );
        }
    }
);


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
        detachExcelTable();
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

            // Toujours relire Excel avant de préparer la demande IA.
            // Une erreur de lecture interrompt la demande : aucune ancienne
            // copie ne doit être envoyée comme s'il s'agissait du classeur.
            let sourceFingerprint = null;
            const sourceGeneration = excelSyncGeneration;
            if (linkedExcelTable) {
                const snapshot = await readLinkedExcelTable();
                sourceFingerprint = JSON.stringify(snapshot.data);
                if (excelPreviewIsDraft && sourceFingerprint !== excelDraftBase) {
                    throw new Error(
                        "Excel a changé depuis la proposition IA. " +
                        "Clique sur Annuler pour reprendre les données d'Excel."
                    );
                }
                if (!excelPreviewIsDraft) {
                    displayAiTable(snapshot.data, snapshot.displayRows);
                    lastExcelSnapshot = JSON.stringify(snapshot);
                }
            }
            const tableBeforeModification =
                JSON.parse(JSON.stringify(currentTableData));

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
                                "Les formules Excel présentes dans les cellules doivent être conservées " +
                                "sauf si la demande exige leur modification. " +
                                "Retourne le tableau complet au format TableIA.\n\n" +
                                "TABLEAU ACTUEL :\n" +
                                JSON.stringify(tableBeforeModification) +
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

            if (sourceGeneration !== excelSyncGeneration) {
                throw new Error("Le tableau a changé. Relance ta demande.");
            }
            if (linkedExcelTable) {
                const latest = await readLinkedExcelTable();
                if (JSON.stringify(latest.data) !== sourceFingerprint) {
                    throw new Error(
                        "Le tableau Excel a changé pendant l'analyse. " +
                        "Relance ta demande pour utiliser les dernières données."
                    );
                }
            }
            if (linkedExcelTable) {
                statusText.textContent = "📊 Mise à jour du tableau dans Excel...";
                await updateLinkedExcelTable(data, sourceFingerprint);
                previousTableData = tableBeforeModification;
                statusText.textContent = "✅ Modification appliquée directement dans ton tableau Excel.";
            } else {
                displayAiTable(data);
                previousTableData = tableBeforeModification;
                statusText.textContent = "✅ Aperçu modifié. Insère le tableau une première fois pour le lier à Excel.";
            }

        } catch (error) {

            console.error(error);

            statusText.textContent =
                "❌ " + (error?.message || "Impossible de modifier le tableau.");
        }

textAnalyzeButton.disabled =
    false;
    
    }

undoTableButton.addEventListener(
    "click",
    async function () {

        if (linkedExcelTable) {
            try {
                const snapshot = await readLinkedExcelTable();
                if (!excelPreviewIsDraft && previousTableData) {
                    const fingerprint = JSON.stringify(snapshot.data);
                    if (fingerprint !== excelAppliedFingerprint) {
                        throw new Error("Excel a changé depuis la dernière modification. Utilise Annuler dans Excel pour conserver ces changements.");
                    }
                    const restoreData = JSON.parse(JSON.stringify(previousTableData));
                    undoTableButton.disabled = true;
                    await updateLinkedExcelTable(restoreData, fingerprint);
                    previousTableData = null;
                    statusText.textContent = "↩ Dernière modification annulée dans Excel.";
                    return;
                }
                excelPreviewIsDraft = false;
                excelDraftBase = null;
                previousTableData = null;
                displayAiTable(snapshot.data, snapshot.displayRows);
                lastExcelSnapshot = JSON.stringify(snapshot);
                statusText.textContent = "↩ Aperçu rétabli depuis Excel.";
            } catch (error) {
                statusText.textContent = "❌ " + error.message;
            } finally {
                undoTableButton.disabled = false;
            }
            return;
        }
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

function createExcelWorksheetName(title) {
    const suffix =
        Date.now().toString().slice(-6);

    const cleanTitle =
        String(title || "TableIA")
            .replace(/[\\/:?*\[\]]/g, " ")
            .replace(/\s+/g, " ")
            .trim()
            .slice(0, 23) ||
        "TableIA";

    return (
        cleanTitle + " " + suffix
    ).slice(0, 31);
}

function normalizeExcelFormula(
    formula,
    replacements
) {
    const cleanFormula =
        String(formula || "")
            .trim()
            .replace(/^=/, "");

    if (!cleanFormula) {
        return "";
    }

    let result =
        cleanFormula;

    Object.keys(replacements).forEach(
        function (placeholder) {
            result = result.replace(
                new RegExp(
                    "\\{" + placeholder + "\\}",
                    "g"
                ),
                String(replacements[placeholder])
            );
        }
    );

    return "=" + result;
}

function coerceExcelValue(value, type) {
    // Conserve aussi les numéros de série des dates relues dans ce classeur.
    if (typeof value === "number" || typeof value === "boolean") return value;
    if (
        value === null ||
        value === undefined ||
        String(value).trim() === ""
    ) {
        return "";
    }

    const text =
        String(value).trim();

    if (
        type === "number" ||
        type === "currency" ||
        type === "percentage"
    ) {
        const normalizedNumber =
            text
                .replace(/\s/g, "")
                .replace(/[€$£%]/g, "")
                .replace(",", ".");

        const numericValue =
            Number(normalizedNumber);

        if (!Number.isNaN(numericValue)) {
            if (
                type === "percentage" &&
                text.includes("%")
            ) {
                return numericValue / 100;
            }

            return numericValue;
        }
    }

    if (type === "boolean") {
        const normalizedBoolean =
            text.toLowerCase();

        if (
            normalizedBoolean === "true" ||
            normalizedBoolean === "vrai" ||
            normalizedBoolean === "oui"
        ) {
            return true;
        }

        if (
            normalizedBoolean === "false" ||
            normalizedBoolean === "faux" ||
            normalizedBoolean === "non"
        ) {
            return false;
        }
    }

    if (type === "date") {
        let day;
        let month;
        let year;

        const frenchDate =
            text.match(
                /^(\d{1,2})[\/-](\d{1,2})[\/-](\d{4})$/
            );

        const isoDate =
            text.match(
                /^(\d{4})-(\d{1,2})-(\d{1,2})$/
            );

        if (frenchDate) {
            day = Number(frenchDate[1]);
            month = Number(frenchDate[2]);
            year = Number(frenchDate[3]);
        } else if (isoDate) {
            year = Number(isoDate[1]);
            month = Number(isoDate[2]);
            day = Number(isoDate[3]);
        }

        if (day && month && year) {
            return (
                Date.UTC(
                    year,
                    month - 1,
                    day
                ) /
                86400000
            ) + 25569;
        }
    }

    return text;
}

function excelNumberFormatForType(type) {
    switch (String(type || "text")) {
        case "currency":
            return "#,##0.00 €";

        case "percentage":
            return "0.00%";

        case "number":
            return "0.##";

        case "date":
            return "dd/mm/yyyy";

        case "text":
        case "status":
            return "@";

        default:
            return "General";
    }
}

// ==========================================
// LECTURE DES MODIFICATIONS NATIVES D'EXCEL
// ==========================================

function excelDocumentSettings() {
    return typeof Office !== "undefined"
        ? Office.context?.document?.settings
        : null;
}

function rememberExcelTable() {
    const settings = excelDocumentSettings();
    if (!settings) return;
    try {
        if (linkedExcelTable) {
            settings.set(EXCEL_TABLE_SETTING, linkedExcelTable);
        } else {
            settings.remove(EXCEL_TABLE_SETTING);
        }
        settings.saveAsync(function (result) {
            if (result.status === Office.AsyncResultStatus.Failed) {
                console.warn("Le lien Excel n'a pas pu être mémorisé.", result.error);
            }
        });
    } catch (error) {
        console.warn("Mémorisation du lien Excel impossible.", error);
    }
}

function detachExcelTable() {
    linkedExcelTable = null;
    excelSyncGeneration += 1;
    clearTimeout(excelSyncTimer);
    excelPreviewIsDraft = false;
    excelDraftBase = null;
    lastExcelSnapshot = null;
    excelAppliedFingerprint = null;
    rememberExcelTable();
}

function inferExcelColumnType(format, values, previousType) {
    // Le format actuel prime sur l'ancien type : l'utilisateur peut le changer.
    const cleanFormat = String(format || "General")
        .replace(/"[^"]*"/g, "");
    if (/%/.test(cleanFormat)) return "percentage";
    if (/[€$£¥]/.test(String(format))) return "currency";
    if (/[dy]/i.test(cleanFormat) && /[m/\-]/i.test(cleanFormat)) return "date";
    if (cleanFormat === "@") return "text";
    const nonEmpty = values.filter(value => value !== "" && value != null);
    if (nonEmpty.length && nonEmpty.every(value => typeof value === "boolean")) {
        return "boolean";
    }
    if (nonEmpty.length && nonEmpty.every(value => typeof value === "number")) {
        return "number";
    }
    return previousType || "text";
}

async function readLinkedExcelTable() {
    if (!linkedExcelTable || typeof Excel === "undefined") {
        throw new Error("Aucun tableau Excel lié.");
    }
    const reference = linkedExcelTable;
    return Excel.run(async function (context) {
        // getItem accepte le nom ou l'identifiant. L'identifiant résiste au renommage.
        const tables = context.workbook.tables;
        tables.load("items/id,items/name");
        await context.sync();
        const table = tables.items.find(item => item.id === reference.id);
        if (!table) {
            throw new Error(
                "Le tableau lié a été supprimé ou converti en plage. " +
                "Crée un nouveau tableau pour poursuivre."
            );
        }
        table.load("showHeaders,showTotals");
        const headers = table.getHeaderRowRange();
        const range = table.getRange();
        headers.load("values");
        range.load("values,text,formulas,numberFormat");
        await context.sync();

        return excelSnapshotFromRanges(reference, table, headers.values, range);
    });
}

function excelSnapshotFromRanges(reference, table, headerValues, range) {
    const firstDataRow = table.showHeaders ? 1 : 0;
    const endDataRow = range.values.length - (table.showTotals ? 1 : 0);
    const values = range.values.slice(firstDataRow, endDataRow);
    const displayRows = range.text.slice(firstDataRow, endDataRow);
    const formulas = range.formulas.slice(firstDataRow, endDataRow);
    const formats = range.numberFormat.slice(firstDataRow, endDataRow);
    const oldColumns = reference.columns || [];
    const names = headerValues[0] || [];
    const columns = names.map(function (name, columnIndex) {
        const oldColumn = oldColumns.find(column => column.name === String(name));
        return {
            name: String(name),
            type: inferExcelColumnType(
                formats[0]?.[columnIndex],
                values.map(row => row[columnIndex]),
                oldColumn?.type
            )
            // Pas de formule de colonne ancienne : chaque cellule Excel
            // est désormais la référence, y compris les exceptions manuelles.
        };
    });
    const rows = values.map(function (row, rowIndex) {
        return columns.map(function (column, columnIndex) {
            const formula = formulas[rowIndex]?.[columnIndex];
            if (typeof formula === "string" && formula.startsWith("=")) {
                return formula;
            }
            return row[columnIndex] ?? "";
        });
    });
    return {
        data: {
            title: reference.title || table.name,
            columns: columns,
            rows: rows,
            summaryFormulas: reference.summaryFormulas || []
        },
        displayRows: displayRows
    };
}

function validateExcelTableData(data) {
    if (!data || !Array.isArray(data.columns) || !data.columns.length ||
        !Array.isArray(data.rows) || data.rows.some(row => !Array.isArray(row))) {
        throw new Error("La réponse IA ne contient pas un tableau valide.");
    }
    const names = data.columns.map(column => String(column?.name || "").trim());
    if (names.some(name => !name) ||
        new Set(names.map(name => name.toLowerCase())).size !== names.length) {
        throw new Error("Chaque colonne doit avoir un nom unique et non vide.");
    }
    return names;
}

function excelCellIsEmpty(value) {
    return value === "" || value === null || value === undefined;
}

function excelRectangleContains(rect, row, column) {
    return row >= rect.row && row < rect.row + rect.height &&
        column >= rect.column && column < rect.column + rect.width;
}

async function updateLinkedExcelTable(data, expectedFingerprint) {
    const names = validateExcelTableData(data);
    if (!linkedExcelTable || excelWriteBusy) {
        throw new Error("Le tableau Excel n'est pas disponible pour la mise à jour.");
    }
    const reference = linkedExcelTable;
    const generation = ++excelSyncGeneration;
    excelWriteBusy = true;
    clearTimeout(excelSyncTimer);
    try {
        let nextReference;
        await Excel.run(async function (context) {
            const table = context.workbook.tables.getItem(reference.id);
            table.load("id,name,showHeaders,showTotals");
            const sheet = table.worksheet;
            const oldRange = table.getRange();
            const header = table.getHeaderRowRange();
            header.load("values");
            oldRange.load("rowIndex,columnIndex,rowCount,columnCount,values,text,formulas,numberFormat");
            await context.sync();
            const current = excelSnapshotFromRanges(reference, table, header.values, oldRange);
            if (JSON.stringify(current.data) !== expectedFingerprint ||
                generation !== excelSyncGeneration || linkedExcelTable?.id !== reference.id) {
                throw new Error("Excel a changé avant l'intégration. Relance ta demande.");
            }
            if (!table.showHeaders) {
                throw new Error("Affiche la ligne d'en-têtes du tableau Excel avant de le modifier avec TableIA.");
            }
            const oldRect = {
                row: oldRange.rowIndex, column: oldRange.columnIndex,
                height: oldRange.rowCount, width: oldRange.columnCount
            };
            const rows = data.rows.length ? data.rows : [names.map(() => "")];
            const nextRect = {
                row: oldRect.row, column: oldRect.column,
                height: rows.length + 1 + (table.showTotals ? 1 : 0), width: names.length
            };
            const resizing = nextRect.height !== oldRect.height || nextRect.width !== oldRect.width;
            if (resizing && !Office.context.requirements.isSetSupported("ExcelApi", "1.13")) {
                throw new Error("Cette version d'Excel ne permet pas le redimensionnement automatique. Mets Excel à jour pour ajouter ou supprimer des lignes et colonnes avec TableIA.");
            }

            // Les indicateurs créés sous le tableau sont déplacés avec lui.
            // Seules les cellules encore identiques à celles de TableIA sont
            // reconnues comme appartenant au complément.
            const oldNames = header.values[0];
            const oldDataCount = oldRect.height - 1 - (table.showTotals ? 1 : 0);
            const oldSummaries = Array.isArray(reference.summaryCells)
                ? reference.summaryCells
                : (reference.summaryFormulas || []).flatMap(function (summary, index) {
                    const column = oldNames.indexOf(summary.targetColumn);
                    if (column < 0) return [];
                    const row = oldRect.row + oldRect.height + 1 + index;
                    return [{
                        row: row, column: oldRect.column + column,
                        labelColumn: column === 0 ? null : oldRect.column,
                        label: summary.label || "Indicateur",
                        formula: normalizeExcelFormula(summary.formula, {
                            lastRow: oldRect.row + oldDataCount + 1
                        })
                    }];
                });
            const oldSummaryRanges = oldSummaries.map(function (summary) {
                const value = sheet.getCell(summary.row, summary.column);
                value.load("values,formulas");
                const label = summary.labelColumn == null ? null : sheet.getCell(summary.row, summary.labelColumn);
                if (label) label.load("values");
                return {summary: summary, value: value, label: label};
            });
            const summaries = (Array.isArray(data.summaryFormulas) ? data.summaryFormulas : [])
                .flatMap(function (summary, index) {
                    const column = names.indexOf(summary.targetColumn);
                    if (column < 0) return [];
                    return [{
                        row: nextRect.row + nextRect.height + 1 + index,
                        column: nextRect.column + column,
                        labelColumn: nextRect.column + (column === 0 ? 1 : 0),
                        label: summary.label || "Indicateur",
                        formula: normalizeExcelFormula(summary.formula, {
                            lastRow: nextRect.row + rows.length + 1
                        }),
                        type: data.columns[column].type
                    }];
                });
            const summaryDestinations = summaries.map(function (summary) {
                const value = sheet.getCell(summary.row, summary.column);
                const label = sheet.getCell(summary.row, summary.labelColumn);
                value.load("values,formulas");
                label.load("values,formulas");
                return {summary: summary, value: value, label: label};
            });
            const target = sheet.getRangeByIndexes(nextRect.row, nextRect.column, nextRect.height, nextRect.width);
            target.load("values,formulas");
            const otherTables = sheet.tables;
            otherTables.load("items/id");
            await context.sync();
            const otherRanges = otherTables.items.filter(item => item.id !== table.id).map(function (item) {
                const range = item.getRange();
                range.load("rowIndex,columnIndex,rowCount,columnCount");
                return range;
            });
            await context.sync();
            if (otherRanges.some(range =>
                nextRect.row < range.rowIndex + range.rowCount &&
                nextRect.row + nextRect.height > range.rowIndex &&
                nextRect.column < range.columnIndex + range.columnCount &&
                nextRect.column + nextRect.width > range.columnIndex)) {
                throw new Error("Un autre tableau occupe l'espace nécessaire à l'intégration.");
            }
            const ownedCells = new Set();
            const ownedRanges = [];
            oldSummaryRanges.forEach(function (item) {
                if (item.value.formulas[0][0] !== item.summary.formula ||
                    (item.label && item.label.values[0][0] !== item.summary.label)) return;
                ownedCells.add(item.summary.row + ":" + item.summary.column);
                ownedRanges.push(item.value);
                if (item.label) {
                    ownedCells.add(item.summary.row + ":" + item.summary.labelColumn);
                    ownedRanges.push(item.label);
                }
            });
            function canReplace(row, column, value, formula) {
                return excelRectangleContains(oldRect, row, column) ||
                    ownedCells.has(row + ":" + column) ||
                    (excelCellIsEmpty(value) && excelCellIsEmpty(formula));
            }
            target.values.forEach(function (row, rowIndex) {
                row.forEach(function (value, columnIndex) {
                    if (!canReplace(nextRect.row + rowIndex, nextRect.column + columnIndex,
                        value, target.formulas[rowIndex][columnIndex])) {
                        throw new Error("Des données à côté du tableau empêchent son agrandissement. Déplace-les puis relance ta demande.");
                    }
                });
            });
            summaryDestinations.forEach(function (item) {
                for (const [column, cell] of [[item.summary.column, item.value], [item.summary.labelColumn, item.label]]) {
                    if (!canReplace(item.summary.row, column, cell.values[0][0], cell.formulas[0][0]) ||
                        otherRanges.some(range => excelRectangleContains({row:range.rowIndex,column:range.columnIndex,height:range.rowCount,width:range.columnCount},item.summary.row,column))) {
                        throw new Error("Des données occupent la place des indicateurs. Déplace-les puis relance ta demande.");
                    }
                }
            });
            // Toutes les vérifications précèdent les écritures dans Excel.
            oldRange.load("values,text,formulas,numberFormat");
            header.load("values");
            await context.sync();
            if (JSON.stringify(excelSnapshotFromRanges(reference, table, header.values, oldRange).data) !== expectedFingerprint) {
                throw new Error("Excel a changé pendant la préparation. Relance ta demande.");
            }
            if (generation !== excelSyncGeneration) throw new Error("Le tableau a changé.");
            ownedRanges.forEach(range => range.clear("All"));
            if (resizing) table.resize(target);
            if (oldRect.height > nextRect.height) {
                sheet.getRangeByIndexes(oldRect.row + nextRect.height, oldRect.column,
                    oldRect.height - nextRect.height, oldRect.width).clear("All");
            }
            if (oldRect.width > nextRect.width) {
                sheet.getRangeByIndexes(oldRect.row, oldRect.column + nextRect.width,
                    Math.min(oldRect.height, nextRect.height), oldRect.width - nextRect.width).clear("All");
            }
            const matrix = [names, ...rows.map(row => names.map((_, index) =>
                coerceExcelValue(row[index], data.columns[index].type)))];
            sheet.getRangeByIndexes(nextRect.row, nextRect.column, matrix.length, names.length).values = matrix;
            const formulaMatrix = matrix.slice(1).map(row => row.slice());
            data.columns.forEach(function (column, columnIndex) {
                const template = String(column.formula || "").trim();
                const oldColumnIndex = oldNames.indexOf(names[columnIndex]);
                rows.forEach(function (row, rowIndex) {
                    const cell = sheet.getCell(nextRect.row + rowIndex + 1, nextRect.column + columnIndex);
                    const value = row[columnIndex];
                    const formula = template || (typeof value === "string" && value.startsWith("=") ? value : "");
                    const originalFormat = oldColumnIndex < 0 ? null : oldRange.numberFormat[rowIndex + 1]?.[oldColumnIndex];
                    const oldType = oldColumnIndex < 0 ? null : current.data.columns[oldColumnIndex]?.type;
                    const keepFormat = !column.type || column.type === oldType;
                    let format = (keepFormat && originalFormat) || excelNumberFormatForType(column.type);
                    if (formula && format === "@") format = "General";
                    cell.numberFormat = [[format]];
                    if (formula) formulaMatrix[rowIndex][columnIndex] = normalizeExcelFormula(formula, {row: nextRect.row + rowIndex + 2});
                });
            });
            // Une seule écriture du corps, avec ses formules ET ses constantes,
            // évite qu'une formule isolée propage un calcul sur les exceptions.
            sheet.getRangeByIndexes(nextRect.row + 1, nextRect.column, rows.length, names.length)
                .formulas = formulaMatrix;
            if (table.showTotals) {
                const oldTotals = oldRange.formulas[oldRange.formulas.length - 1];
                names.forEach(function (name, index) {
                    const oldIndex = oldNames.indexOf(name);
                    if (oldIndex >= 0) {
                        sheet.getCell(nextRect.row + nextRect.height - 1, nextRect.column + index)
                            .formulas = [[oldTotals[oldIndex]]];
                    }
                });
            }
            summaryDestinations.forEach(function (item) {
                item.label.values = [[item.summary.label]];
                item.value.formulas = [[item.summary.formula]];
                item.value.numberFormat = [[excelNumberFormatForType(item.summary.type)]];
                [item.label, item.value].forEach(function (cell) {
                    cell.format.font.bold = true;
                    cell.format.fill.color = "#E2F0D9";
                });
            });
            await context.sync();
            nextReference = {
                ...reference, title: data.title || reference.title,
                columns: data.columns.map((column, index) => ({name: names[index], type: column.type})),
                summaryFormulas: data.summaryFormulas || [], summaryCells: summaries
            };
        });
        linkedExcelTable = nextReference;
        excelPreviewIsDraft = false;
        excelDraftBase = null;
        rememberExcelTable();
        const snapshot = await readLinkedExcelTable();
        displayAiTable(snapshot.data, snapshot.displayRows);
        lastExcelSnapshot = JSON.stringify(snapshot);
        excelAppliedFingerprint = JSON.stringify(snapshot.data);
    } finally {
        excelWriteBusy = false;
        scheduleExcelRefresh();
    }
}

function scheduleExcelRefresh(event) {
    if (!linkedExcelTable) return;
    if (event?.tableId && event.tableId !== linkedExcelTable.id) return;
    clearTimeout(excelSyncTimer);
    excelSyncTimer = setTimeout(refreshExcelPreview, 350);
}

async function refreshExcelPreview() {
    if (!linkedExcelTable) return;
    if (excelSyncBusy || excelWriteBusy) {
        scheduleExcelRefresh();
        return;
    }
    excelSyncBusy = true;
    const generation = excelSyncGeneration;
    try {
        const snapshot = await readLinkedExcelTable();
        if (generation !== excelSyncGeneration) return;
        const fingerprint = JSON.stringify(snapshot.data);
        if (excelPreviewIsDraft) {
            if (fingerprint !== excelDraftBase) {
                statusText.textContent =
                    "Excel a changé. Clique sur Annuler pour reprendre ses données dans l'aperçu.";
            }
            return;
        }
        const previewFingerprint = JSON.stringify(snapshot);
        if (previewFingerprint !== lastExcelSnapshot) {
            displayAiTable(snapshot.data, snapshot.displayRows);
            if (fingerprint !== excelAppliedFingerprint) previousTableData = null;
            lastExcelSnapshot = previewFingerprint;
        }
    } catch (error) {
        if (generation === excelSyncGeneration) {
            console.warn("Lecture du tableau Excel impossible.", error);
            statusText.textContent = "⚠️ " + error.message;
        }
    } finally {
        excelSyncBusy = false;
    }
}

async function initializeExcelTableSync() {
    const info = await officeInitialization;
    if (!info || typeof Excel === "undefined") return;
    const generation = excelSyncGeneration;
    try {
        const saved = excelDocumentSettings()?.get(EXCEL_TABLE_SETTING);
        // Aucun rattachement arbitraire aux anciens tableaux : réinsérer une
        // première fois avec cette version établit le lien sans ambiguïté.
        if (saved && typeof saved.id === "string" && !currentTableData &&
            generation === excelSyncGeneration) {
            linkedExcelTable = saved;
            await refreshExcelPreview();
        }
        if (Office.context.requirements.isSetSupported("ExcelApi", "1.7")) {
            await Excel.run(async function (context) {
                context.workbook.tables.onChanged.add(scheduleExcelRefresh);
                await context.sync();
            });
        }
    } catch (error) {
        console.warn("Initialisation du suivi Excel impossible.", error);
    }
}

// Un retour au volet permet aussi de reprendre une suppression
// de tableau et les recalculs, même si aucun événement de cellule n'est reçu.
window.addEventListener("focus", function () { scheduleExcelRefresh(); });
document.addEventListener("visibilitychange", function () {
    if (!document.hidden) scheduleExcelRefresh();
});
initializeExcelTableSync();

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

            if (linkedExcelTable) {
                if (excelPreviewIsDraft) {
                    const before = await readLinkedExcelTable();
                    await updateLinkedExcelTable(currentTableData, excelDraftBase);
                    previousTableData = before.data;
                    statusText.textContent = "✅ Modification appliquée dans le tableau Excel existant.";
                } else {
                    const snapshot = await readLinkedExcelTable();
                    displayAiTable(snapshot.data, snapshot.displayRows);
                    lastExcelSnapshot = JSON.stringify(snapshot);
                    statusText.textContent = "✅ Tableau déjà lié. Modifie-le directement dans Excel.";
                }
                return;
            }
            const insertionData = JSON.parse(JSON.stringify(currentTableData));
            let insertedReference = null;

            await Excel.run(
                async function (context) {

                    let sheet =
                        context.workbook.worksheets
                            .getActiveWorksheet();

                    const usedRange =
                        sheet.getUsedRangeOrNullObject();

                    usedRange.load("values");

                    await context.sync();

                    const activeSheetContainsData =
                        !usedRange.isNullObject && usedRange.values.some(
                            function (row) {
                                return row.some(
                                    function (value) {
                                        return value !== "" &&
                                            value !== null;
                                    }
                                );
                            }
                        );

                    if (activeSheetContainsData) {
                        sheet =
                            context.workbook.worksheets.add(
                                createExcelWorksheetName(
                                    insertionData.title
                                )
                            );

                        sheet.activate();
                    }

                    const headers =
                        insertionData.columns.map(
                            function (column) {
                                return column.name || "";
                            }
                        );

                    const rows =
                        Array.isArray(insertionData.rows) &&
                        insertionData.rows.length
                            ? insertionData.rows
                            : [[], [], []];

                    const values = [
                        headers,
                        ...rows.map(
                            function (row) {
                                return headers.map(
                                    function (_, index) {
                                        return coerceExcelValue(
                                            row[index],
                                            insertionData.columns[
                                                index
                                            ]?.type
                                        );
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

                    const table =
                        sheet.tables.add(
                            range,
                            true
                        );

                    table.name =
                        "TableIA_" +
                        Date.now().toString().slice(-8);

                    table.style =
                        "TableStyleMedium7";

                    table.showHeaders =
                        true;

                    table.showTotals =
                        false;

                    const dataRowCount =
                        rows.length;

                    insertionData.columns.forEach(
                        function (column, columnIndex) {
                            const columnRange =
                                sheet.getRangeByIndexes(
                                    1,
                                    columnIndex,
                                    dataRowCount,
                                    1
                                );

                            const columnFormula =
                                String(
                                    column.formula || ""
                                ).trim();

                            // Une cellule calculée ne doit jamais être forcée
                            // au format texte, sinon Excel affiche la formule
                            // au lieu de l'exécuter (cas des colonnes "status").
                            const formulaReturnsText =
                                columnFormula &&
                                [
                                    "text",
                                    "status",
                                    "boolean"
                                ].includes(column.type);

                            const numberFormat =
                                formulaReturnsText
                                    ? "General"
                                    : excelNumberFormatForType(
                                        column.type
                                    );

                            columnRange.numberFormat =
                                Array.from(
                                    {
                                        length: dataRowCount
                                    },
                                    function () {
                                        return [
                                            numberFormat
                                        ];
                                    }
                                );

                            if (columnFormula) {
                                columnRange.formulas =
                                    Array.from(
                                        {
                                            length: dataRowCount
                                        },
                                        function (_, rowIndex) {
                                            return [
                                                normalizeExcelFormula(
                                                    columnFormula,
                                                    {
                                                        row:
                                                            rowIndex + 2
                                                    }
                                                )
                                            ];
                                        }
                                    );
                            } else {
                                // Les formules relues d'Excel sont portées par
                                // les cellules, pour respecter les exceptions.
                                const hasCellFormula = rows.some(function (row) {
                                    return typeof row[columnIndex] === "string" && row[columnIndex].startsWith("=");
                                });
                                if (hasCellFormula) {
                                rows.forEach(function (row, rowIndex) {
                                    const value = row[columnIndex];
                                    if (typeof value === "string" && value.startsWith("=")) {
                                        const cell = sheet.getCell(rowIndex + 1, columnIndex);
                                        cell.numberFormat = [[numberFormat === "@" ? "General" : numberFormat]];
                                    }
                                });
                                columnRange.formulas = rows.map(function (row) {
                                    return [coerceExcelValue(row[columnIndex], column.type)];
                                });
                                }
                            }
                        }
                    );

                    const summaries =
                        Array.isArray(
                            insertionData.summaryFormulas
                        )
                            ? insertionData.summaryFormulas
                            : [];

                    summaries.forEach(
                        function (summary, summaryIndex) {
                            const targetColumnIndex =
                                headers.indexOf(
                                    summary.targetColumn
                                );

                            if (targetColumnIndex === -1) {
                                return;
                            }

                            const summaryRowIndex =
                                dataRowCount +
                                2 +
                                summaryIndex;

                            const labelCell =
                                sheet.getCell(
                                    summaryRowIndex,
                                    0
                                );

                            const valueCell =
                                sheet.getCell(
                                    summaryRowIndex,
                                    targetColumnIndex
                                );

                            labelCell.values = [[
                                summary.label ||
                                "Indicateur"
                            ]];

                            labelCell.format.font.bold =
                                true;

                            labelCell.format.fill.color =
                                "#E2F0D9";

                            valueCell.formulas = [[
                                normalizeExcelFormula(
                                    summary.formula,
                                    {
                                        lastRow:
                                            dataRowCount + 1
                                    }
                                )
                            ]];

                            valueCell.format.font.bold =
                                true;

                            valueCell.format.fill.color =
                                "#E2F0D9";

                            valueCell.numberFormat = [[
                                excelNumberFormatForType(
                                    insertionData.columns[
                                        targetColumnIndex
                                    ]?.type
                                )
                            ]];
                        }
                    );

                    const finalRange =
                        sheet.getUsedRange();

                    finalRange.format.autofitColumns();
                    finalRange.format.autofitRows();

                    table.load("id");
                    await context.sync();
                    insertedReference = {
                        id: table.id,
                        title: insertionData.title || table.name,
                        columns: insertionData.columns.map(column => ({name: column.name, type: column.type})),
                        summaryFormulas: insertionData.summaryFormulas || [],
                        summaryCells: summaries.flatMap(function (summary, index) {
                            const column = headers.indexOf(summary.targetColumn);
                            if (column < 0) return [];
                            return [{row: dataRowCount + 2 + index, column: column,
                                labelColumn: column === 0 ? null : 0,
                                label: summary.label || "Indicateur",
                                formula: normalizeExcelFormula(summary.formula, {lastRow: dataRowCount + 1})}];
                        })
                    };
                }
            );

            linkedExcelTable = insertedReference;
            excelSyncGeneration += 1;
            excelPreviewIsDraft = false;
            excelDraftBase = null;
            lastExcelSnapshot = null;
            rememberExcelTable();
            await refreshExcelPreview();

            statusText.textContent =
                "✅ Tableau lié. Tu peux modifier ses cellules, lignes et colonnes directement dans Excel.";

        } catch (error) {

            console.error(error);

            statusText.textContent =
                "❌ Impossible de créer le tableau Excel : " +
                (error?.message || "erreur inconnue");
        } finally {
            insertExcelButton.disabled = false;
        }
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

function displayAiTable(data, previewRows) {

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
        function (row, rowIndex) {

            const tr =
                document.createElement("tr");

            data.columns.forEach(
                function (_, index) {

                    const td =
                        document.createElement("td");

                    td.textContent =
                        previewRows?.[rowIndex]?.[index] ?? row[index] ?? "";

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
