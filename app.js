const SUPABASE_FUNCTION_URL =
    "https://ctihdqybwmymalpkjkov.supabase.co/functions/v1/tableia-ai";

const SUPABASE_PUBLISHABLE_KEY =
    "sb_publishable_LU68bB5ip1OEPr8DJnjJJg_G36GmXoY";

const cameraInput =
    document.getElementById("cameraInput");

const galleryInput =
    document.getElementById("galleryInput");;

const photoPreview =
    document.getElementById("photoPreview");

const analyzeButton =
    document.getElementById("analyzeButton");

const statusText =
    document.getElementById("status");

const resultSection =
    document.getElementById("resultSection");

const tableContainer =
    document.getElementById("tableContainer");

const modifyTableButton =
    document.getElementById("modifyTableButton");

const deleteTableButton =
    document.getElementById("deleteTableButton");

const insertExcelButton =
    document.getElementById("insertExcelButton");


let selectedPhoto = null;
let currentTableData = null;


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

    analyzeButton.disabled =
        false;

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
        
        resultSection.style.display =
            "none";

        insertExcelButton.disabled =
            true;

        statusText.textContent =
            "🗑 Tableau supprimé.";
    }
);
modifyTableButton.addEventListener(
    "click",
    async function () {

        if (!currentTableData) {
            statusText.textContent =
                "Aucun tableau à modifier.";
            return;
        }

        const instruction =
            window.prompt(
                "Quelle modification veux-tu apporter au tableau ?"
            );

        if (!instruction) {
            return;
        }

        modifyTableButton.disabled =
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

        modifyTableButton.disabled =
            false;
    }
);
// ==========================================
// ANALYSE
// ==========================================

analyzeButton.addEventListener(
    "click",
    async function () {

        if (!selectedPhoto) {

            statusText.textContent =
                "Veuillez sélectionner une photo.";

            return;
        }

        analyzeButton.disabled =
            true;

        statusText.textContent =
            "✨ Préparation de la photo...";

        try {

            const imageBase64 =
                await prepareImage(
                    selectedPhoto
                );

            statusText.textContent =
    "✨ Analyse de la photo par l'IA...";

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
                    "Analyse cette image et transforme les informations visibles en un tableau TableIA structuré.",

                imageBase64:
                    imageBase64,

                imageMimeType:
                    "image/jpeg"
            })
        }
    );

const responseText =
    await response.text();

if (!response.ok) {

    throw new Error(
        `Erreur ${response.status} : ${responseText}`
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
    "✅ Photo analysée — tableau généré.";

        } catch (error) {

            console.error(
                error
            );

            statusText.textContent =
                "❌ Impossible de préparer la photo.";
        }

        analyzeButton.disabled =
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
