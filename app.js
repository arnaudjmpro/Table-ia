const SUPABASE_FUNCTION_URL =
    "https://ctihdqybwmymalpkjkov.supabase.co/functions/v1/tableia-ai";

const SUPABASE_PUBLISHABLE_KEY =
    "sb_publishable_LU68bB5ip1OEPr8DJnjJJg_G36GmXoY";

const PHOTO_TABLE_INSTRUCTION =
    "Lis la photo et crée un tableau TableIA fidèle aux données visibles. " +
    "Identifie les en-têtes utiles : par exemple Nom et Montant si la photo contient des noms et des sommes. " +
    "Mets une personne ou une entrée par ligne, avec chaque valeur dans sa colonne. " +
    "Pour les sommes, utilise le type currency et conserve les montants lisibles. " +
    "Si quantité et prix unitaire sont présents, ajoute une colonne Total calculée ; sinon ne crée pas de colonne Total redondante. " +
    "Ne devine pas les valeurs illisibles : laisse ces cellules vides. " +
    "Place les totaux dans summaryFormulas, hors des lignes de données, pour qu'Excel puisse recalculer la somme. ";

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

// La Preview Replit utilise elle aussi une iframe : seule la présence réelle
// de l'hôte Excel doit désactiver la conversation destinée au navigateur.
const isExcelHost = isEmbeddedInOffice &&
    typeof Office !== "undefined" &&
    Office.context && Office.context.host === "Excel";
const isWebConversation = !isExcelHost;
const conversationPanel = document.getElementById("conversationPanel");
const conversationMessagesElement = document.getElementById("conversationMessages");
const conversationHint = document.getElementById("conversationHint");
const newConversationButton = document.getElementById("newConversationButton");
const sendHint = document.getElementById("sendHint");
const CONVERSATION_STORAGE_KEY = "tableia.webConversation.v1";

let conversationMessages = [];

if (isWebConversation) {
    document.body.classList.add("web-conversation");
    conversationPanel.hidden = false;
    sendHint.hidden = false;
    insertExcelButton.textContent = "📥 TÉLÉCHARGER LE FICHIER EXCEL (.XLSX)";
}

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
let photoPreviewUrl = null;

function releasePhotoPreview() {
    if (photoPreviewUrl) {
        URL.revokeObjectURL(photoPreviewUrl);
        photoPreviewUrl = null;
    }
    photoPreview.removeAttribute("src");
    photoPreview.style.display = "none";
}

function saveConversation() {
    if (!isWebConversation) return;
    try {
        sessionStorage.setItem(CONVERSATION_STORAGE_KEY, JSON.stringify({
            messages: conversationMessages.slice(-40),
            table: currentTableData,
            previousTable: previousTableData
        }));
    } catch (error) {
        console.warn("Conversation non sauvegardée dans cet onglet.", error);
    }
}

function appendConversationMessage(role, text, persist = true) {
    if (!isWebConversation) return;
    const message = { role, text: String(text || "") };
    conversationMessages.push(message);
    const bubble = document.createElement("div");
    bubble.className = "conversation-message " + role;
    bubble.textContent = message.text;
    conversationMessagesElement.appendChild(bubble);
    conversationHint.hidden = true;
    conversationMessagesElement.scrollTop = conversationMessagesElement.scrollHeight;
    if (persist) saveConversation();
}

function clearComposer() {
    textPrompt.value = "";
    selectedPhoto = null;
    cameraInput.value = "";
    galleryInput.value = "";
    releasePhotoPreview();
    clearPromptButton.style.display = "none";
}

function conversationContext() {
    return conversationMessages.slice(-12).map(function (message) {
        return (message.role === "user" ? "Utilisateur : " : "TableIA : ") + message.text;
    }).join("\n");
}

if (isWebConversation) {
    newConversationButton.addEventListener("click", function () {
        conversationMessages = [];
        conversationMessagesElement.replaceChildren();
        conversationHint.hidden = false;
        currentTableData = null;
        previousTableData = null;
        tableContainer.replaceChildren();
        resultSection.style.display = "none";
        insertExcelButton.disabled = true;
        textAnalyzeButton.textContent = "✨ CRÉER LE TABLEAU";
        statusText.textContent = "";
        clearComposer();
        saveConversation();
        textPrompt.focus();
    });

    textPrompt.addEventListener("keydown", function (event) {
        if (event.key === "Enter" && !event.shiftKey && !event.isComposing) {
            event.preventDefault();
            if (!textAnalyzeButton.disabled) textAnalyzeButton.click();
        }
    });
}

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

            if (prompt) body.prompt = prompt;

            if (selectedPhoto) {

                const imageBase64 =
                    await prepareImage(
                        selectedPhoto
                    );

                body.imageBase64 =
                    imageBase64;

                body.imageMimeType =
                    "image/jpeg";

                body.prompt = PHOTO_TABLE_INSTRUCTION +
                    (prompt ? "\n\nDemande de l'utilisateur : " + prompt : "");
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

            const data = selectedPhoto
                ? structurePhotoTable(JSON.parse(responseText))
                : JSON.parse(responseText);

            displayAiTable(
                data
            );

            appendConversationMessage("user", prompt || "Analyse cette photo et crée un tableau.");
            appendConversationMessage("assistant", "Le tableau « " +
                (data.title || "TableIA") + " » est prêt. Vous pouvez me demander de le modifier.");
            if (isWebConversation) clearComposer();

            statusText.textContent =
                "✅ Tableau généré avec succès.";

        } catch (error) {

            console.error(error);

            appendConversationMessage("assistant", "Je n'ai pas pu créer le tableau. Vérifiez votre demande et réessayez.");

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

releasePhotoPreview();
        
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

    releasePhotoPreview();
    photoPreviewUrl = URL.createObjectURL(file);

    photoPreview.src =
        photoPreviewUrl;

    photoPreview.style.display =
        "block";

clearPromptButton.style.display =
    "block";
    
    statusText.textContent =
        "📷 Photo prise en compte. Prête pour l'analyse.";
}

const webCameraDialog = document.getElementById("webCameraDialog");
const webCameraVideo = document.getElementById("webCameraVideo");
const webCameraStatus = document.getElementById("webCameraStatus");
const capturePhotoButton = document.getElementById("capturePhotoButton");
const cancelPhotoButton = document.getElementById("cancelPhotoButton");
let webCameraStream = null;

function stopWebCamera() {
    if (webCameraStream) {
        webCameraStream.getTracks().forEach(track => track.stop());
        webCameraStream = null;
    }
    webCameraVideo.srcObject = null;
    capturePhotoButton.disabled = true;
}

async function openWebCamera() {
    if (!navigator.mediaDevices?.getUserMedia) {
        throw new Error("Caméra indisponible dans ce navigateur. Vérifiez que la page est ouverte en HTTPS.");
    }
    webCameraStatus.textContent = "Autorisez l'accès à la caméra dans Safari…";
    capturePhotoButton.disabled = true;
    webCameraDialog.showModal();
    try {
        const stream = await navigator.mediaDevices.getUserMedia({
            audio: false,
            video: true
        });
        if (!webCameraDialog.open) {
            stream.getTracks().forEach(track => track.stop());
            return;
        }
        webCameraStream = stream;
        webCameraVideo.srcObject = stream;
        await webCameraVideo.play();
        if (!webCameraDialog.open || webCameraStream !== stream) return;
        webCameraStatus.textContent = "Cadrez le document puis prenez la photo.";
        capturePhotoButton.disabled = false;
    } catch (error) {
        const stillOpen = webCameraDialog.open;
        stopWebCamera();
        if (stillOpen) webCameraDialog.close();
        if (!stillOpen) return;
        if (error?.name === "NotAllowedError" || error?.name === "PermissionDeniedError") {
            throw new Error("Accès à la caméra refusé. Autorisez la caméra pour ce site dans Safari.");
        }
        if (error?.name === "NotFoundError" || error?.name === "DevicesNotFoundError") {
            throw new Error("Aucune caméra n'a été détectée sur cet appareil.");
        }
        throw new Error("La caméra n'a pas pu démarrer : " + (error?.message || "erreur inconnue"));
    }
}

capturePhotoButton.addEventListener("click", function () {
    if (!webCameraStream || !webCameraVideo.videoWidth || !webCameraVideo.videoHeight) {
        webCameraStatus.textContent = "Attendez que l'image de la caméra apparaisse.";
        return;
    }
    try {
        const canvas = document.createElement("canvas");
        const ratio = Math.min(1, 1600 / Math.max(webCameraVideo.videoWidth, webCameraVideo.videoHeight));
        canvas.width = Math.round(webCameraVideo.videoWidth * ratio);
        canvas.height = Math.round(webCameraVideo.videoHeight * ratio);
        const context = canvas.getContext("2d");
        if (!context) throw new Error("Capture d'image indisponible.");
        context.drawImage(webCameraVideo, 0, 0, canvas.width, canvas.height);
        handleSelectedPhoto(dataUrlToPhotoFile(canvas.toDataURL("image/jpeg", 0.88)));
        webCameraDialog.close();
    } catch (error) {
        console.error(error);
        webCameraStatus.textContent = "Impossible de prendre la photo : " +
            (error?.message || "erreur inconnue");
    }
});

cancelPhotoButton.addEventListener("click", function () {
    webCameraDialog.close();
});
webCameraDialog.addEventListener("close", stopWebCamera);
window.addEventListener("pagehide", stopWebCamera);


cameraButton.addEventListener(
    "click",
    function (event) {
        // Sur téléphone et tablette, le champ capture natif reste
        // la solution la plus directe et déjà validée.
        if (isMobileCaptureDevice) {
            return;
        }

        // Sur Mac et PC, l'attribut capture ouvre généralement le
        // sélecteur de fichiers. Une vraie fenêtre webcam est donc utilisée.
        event.preventDefault();

        if (isExcelHost) {
            try {
                openOfficeCameraDialog();
            } catch (error) {
                console.error(error);
                statusText.textContent = "❌ Photo : " +
                    (error?.message || "ouverture impossible");
            }
            return;
        }
        openWebCamera().catch(function (error) {
            console.error(error);
            statusText.textContent = "❌ Photo : " +
                (error?.message || "ouverture impossible") +
                " Vous pouvez aussi choisir une image dans la galerie.";
        });
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
        textAnalyzeButton.textContent =
    "✨ CRÉER LE TABLEAU";
previousTableData =
    null;
        appendConversationMessage("assistant", "Le tableau a été supprimé. Décrivez le nouveau tableau à créer.");
        saveConversation();
        
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

        const instruction = textPrompt.value.trim() ||
            (selectedPhoto ? "Ajoute au tableau actuel les informations visibles sur cette photo, " +
                "en les plaçant dans les colonnes adaptées. Conserve les lignes existantes." : "");

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

            const previousSnapshot = JSON.parse(JSON.stringify(currentTableData));
            const body = {
                prompt:
                    (selectedPhoto ? PHOTO_TABLE_INSTRUCTION + "\n\n" : "") +
                    "Modifie ce tableau existant selon la demande de l'utilisateur. " +
                    "Conserve toutes les données qui ne sont pas concernées. " +
                    "Retourne le tableau complet au format TableIA.\n\n" +
                    (isWebConversation && conversationMessages.length
                        ? "CONVERSATION PRÉCÉDENTE :\n" + conversationContext() + "\n\n"
                        : "") +
                    "TABLEAU ACTUEL :\n" + JSON.stringify(currentTableData) +
                    "\n\nMODIFICATION DEMANDÉE :\n" + instruction
            };
            if (selectedPhoto) {
                body.imageBase64 = await prepareImage(selectedPhoto);
                body.imageMimeType = "image/jpeg";
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

                        body: JSON.stringify(body)
                    }
                );

            const responseText =
                await response.text();

            if (!response.ok) {
                throw new Error(
                    responseText
                );
            }

            const data = selectedPhoto
                ? structurePhotoTable(JSON.parse(responseText))
                : JSON.parse(responseText);

            previousTableData = previousSnapshot;
            displayAiTable(data);
            appendConversationMessage("user", instruction + (selectedPhoto ? " [photo jointe]" : ""));
            appendConversationMessage("assistant", "Le tableau a été mis à jour. Vous pouvez poursuivre la discussion.");
            if (isWebConversation) clearComposer();

            statusText.textContent =
                "✅ Tableau modifié avec succès.";

        } catch (error) {

            console.error(error);

            appendConversationMessage("assistant", "Je n'ai pas pu modifier le tableau. Réessayez sans perdre le tableau actuel.");

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

        appendConversationMessage("assistant", "La dernière modification du tableau a été annulée.");
        saveConversation();

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
        let normalizedNumber = text
            .replace(/\s/g, "")
            .replace(/[€$£%]/g, "");

        // Les photos de documents français contiennent souvent 1.234,56 €.
        if (normalizedNumber.includes(",") && normalizedNumber.includes(".")) {
            normalizedNumber = normalizedNumber.lastIndexOf(",") > normalizedNumber.lastIndexOf(".")
                ? normalizedNumber.replace(/\./g, "").replace(",", ".")
                : normalizedNumber.replace(/,/g, "");
        } else if (normalizedNumber.includes(",")) {
            normalizedNumber = normalizedNumber.replace(",", ".");
        } else if (type === "currency" && /^-?\d{1,3}(?:\.\d{3})+$/.test(normalizedNumber)) {
            normalizedNumber = normalizedNumber.replace(/\./g, "");
        }

        const numericValue =
            Number(normalizedNumber);

        if (normalizedNumber && Number.isFinite(numericValue)) {
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

function excelColumnLetter(index) {
    let result = "";
    for (let n = index + 1; n > 0; n = Math.floor((n - 1) / 26)) {
        result = String.fromCharCode(65 + (n - 1) % 26) + result;
    }
    return result;
}

function structurePhotoTable(data) {
    if (!data || !Array.isArray(data.columns) || !Array.isArray(data.rows)) return data;

    const columns = data.columns.map(column => ({ ...column }));
    const rows = data.rows.map(row => Array.isArray(row) ? row.slice() : row);
    const monetaryHeader = /montant|prix|co[uû]t|somme|solde/i;
    const hasMonetaryHeader = columns.some(column => monetaryHeader.test(column.name || ""));

    const amountColumns = columns.flatMap((column, index) => {
        const header = String(column.name || "");
        if (!monetaryHeader.test(header) && !(hasMonetaryHeader && /\btotal\b/i.test(header))) return [];
        const hasAmount = rows.some(row => Array.isArray(row) &&
            typeof coerceExcelValue(row[index], "currency") === "number");
        if (!hasAmount) return [];
        if (!column.type || column.type === "text" || column.type === "number") column.type = "currency";
        return [index];
    });

    const totalIndex = amountColumns.slice().reverse().find(index => /\btotal\b/i.test(columns[index].name || "")) ??
        amountColumns[amountColumns.length - 1];
    if (totalIndex === undefined) return { ...data, columns, rows };

    const lastRow = rows[rows.length - 1];
    if (rows.length > 1 && Array.isArray(lastRow) &&
        /^total(?:\s+g[ée]n[ée]ral)?$/i.test(String(lastRow[0] || "").trim()) &&
        typeof coerceExcelValue(lastRow[totalIndex], "currency") === "number") {
        rows.pop(); // Le total imprimé n'est pas une ligne de données à recompter.
    }

    const summaries = Array.isArray(data.summaryFormulas) ? data.summaryFormulas.slice() : [];
    if (rows.length) {
        const letter = excelColumnLetter(totalIndex);
        const formula = `=SUM(${letter}2:${letter}{lastRow})`;
        const existingTotalIndex = summaries.findIndex(summary =>
            summary.targetColumn === columns[totalIndex].name && /^=?\s*SUM\(/i.test(summary.formula || ""));
        if (existingTotalIndex >= 0) {
            summaries[existingTotalIndex] = { ...summaries[existingTotalIndex], formula };
        } else {
            summaries.push({
                label: "Total",
                targetColumn: columns[totalIndex].name,
                formula
            });
        }
    }
    return { ...data, columns, rows, summaryFormulas: summaries };
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

insertExcelButton.addEventListener(
    "click",
    async function () {

        if (!currentTableData) {
            statusText.textContent =
                isWebConversation
                    ? "Aucun tableau à télécharger."
                    : "Aucun tableau à insérer.";
            return;
        }

        if (isWebConversation) {
            insertExcelButton.disabled = true;
            try {
                if (!window.TableIAXlsx) {
                    throw new Error("Le module d'export Excel n'a pas été chargé.");
                }
                const workbook = window.TableIAXlsx.createWorkbookBlob(
                    currentTableData,
                    { coerceExcelValue, normalizeExcelFormula }
                );
                const filename = String(currentTableData.title || "TableIA")
                    .replace(/[\\/:*?"<>|]/g, " ")
                    .replace(/\s+/g, " ")
                    .trim()
                    .slice(0, 80) || "TableIA";
                const url = URL.createObjectURL(workbook);
                const download = document.createElement("a");
                download.href = url;
                download.download = filename + ".xlsx";
                document.body.appendChild(download);
                download.click();
                download.remove();
                setTimeout(() => URL.revokeObjectURL(url), 60000);
                statusText.textContent =
                    "✅ Fichier .xlsx téléchargé. Ouvrez-le depuis Téléchargements.";
            } catch (error) {
                console.error(error);
                statusText.textContent = "❌ Export Excel impossible : " +
                    (error?.message || "erreur inconnue");
            } finally {
                insertExcelButton.disabled = false;
            }
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

                    let sheet =
                        context.workbook.worksheets
                            .getActiveWorksheet();

                    const usedRange =
                        sheet.getUsedRange();

                    usedRange.load("values");

                    await context.sync();

                    const activeSheetContainsData =
                        usedRange.values.some(
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
                                    currentTableData.title
                                )
                            );

                        sheet.activate();
                    }

                    const headers =
                        currentTableData.columns.map(
                            function (column) {
                                return column.name || "";
                            }
                        );

                    const rows =
                        Array.isArray(currentTableData.rows) &&
                        currentTableData.rows.length
                            ? currentTableData.rows
                            : [[], [], []];

                    const values = [
                        headers,
                        ...rows.map(
                            function (row) {
                                return headers.map(
                                    function (_, index) {
                                        return coerceExcelValue(
                                            row[index],
                                            currentTableData.columns[
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

                    currentTableData.columns.forEach(
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
                            }
                        }
                    );

                    const summaries =
                        Array.isArray(
                            currentTableData.summaryFormulas
                        )
                            ? currentTableData.summaryFormulas
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
                                1 +
                                summaryIndex;

                            const labelCell =
                                sheet.getCell(
                                    summaryRowIndex,
                                    targetColumnIndex === 0 ? 1 : 0
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
                                    currentTableData.columns[
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

                    await context.sync();
                }
            );

            statusText.textContent =
                "✅ Véritable tableau Excel créé avec filtres, style et formules.";

        } catch (error) {

            console.error(error);

            statusText.textContent =
                "❌ Impossible de créer le tableau Excel : " +
                (error?.message || "erreur inconnue");
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

    const summaries = Array.isArray(data.summaryFormulas) ? data.summaryFormulas : [];
    summaries.forEach(function (summary) {
        const index = data.columns.findIndex(column => column.name === summary.targetColumn);
        if (index < 0 || !rows.length) return;
        const letter = excelColumnLetter(index);
        const formula = String(summary.formula || "").replace(/\s/g, "").toUpperCase();
        const lastRow = rows.length + 1;
        if (formula !== `=SUM(${letter}2:${letter}{LASTROW})` &&
            formula !== `=SUM(${letter}2:${letter}${lastRow})`) return;

        const values = rows.map(row => coerceExcelValue(Array.isArray(row) ? row[index] : "", data.columns[index].type));
        if (!values.some(value => typeof value === "number")) return;
        const total = values.reduce((sum, value) => sum + (typeof value === "number" ? value : 0), 0);
        const tr = document.createElement("tr");
        tr.className = "summary-row";
        data.columns.forEach(function (column, columnIndex) {
            const cell = document.createElement("td");
            if (columnIndex === index) {
                cell.textContent = new Intl.NumberFormat("fr-FR", column.type === "currency"
                    ? { style: "currency", currency: "EUR" }
                    : { maximumFractionDigits: 2 }).format(total);
            } else if (columnIndex === (index === 0 ? 1 : 0)) {
                cell.textContent = summary.label || "Total";
            }
            tr.appendChild(cell);
        });
        table.appendChild(tr);
    });


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

if (isWebConversation) {
    try {
        const saved = JSON.parse(sessionStorage.getItem(CONVERSATION_STORAGE_KEY) || "null");
        if (saved && Array.isArray(saved.messages)) {
            saved.messages.slice(-40).forEach(function (message) {
                if (message && (message.role === "user" || message.role === "assistant") &&
                    typeof message.text === "string") {
                    appendConversationMessage(message.role, message.text, false);
                }
            });
        }
        if (saved && saved.table && Array.isArray(saved.table.columns)) {
            displayAiTable(saved.table);
            if (saved.previousTable && Array.isArray(saved.previousTable.columns)) {
                previousTableData = saved.previousTable;
            }
        }
    } catch (error) {
        console.warn("La conversation précédente n'a pas pu être restaurée.", error);
    }
}
