const SUPABASE_FUNCTION_URL =
    "https://ctihdqybwmymalpkjkov.supabase.co/functions/v1/tableia-ai";

const SUPABASE_PUBLISHABLE_KEY =
    "sb_publishable_LU68bB5ip1OEPr8DJnjJJg_G36GmXoY";

const photoInput =
    document.getElementById("photoInput");

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

const insertExcelButton =
    document.getElementById("insertExcelButton");


let selectedPhoto = null;


// ==========================================
// PHOTO
// ==========================================

photoInput.addEventListener(
    "change",
    function () {

        const file =
            photoInput.files[0];

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
                "✅ Photo prête pour l'analyse IA.";

            /*
                L'étape suivante connectera ici
                notre fonction Supabase TableIA.
            */

            console.log(
                "Image prête :",
                imageBase64.length,
                "caractères"
            );

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
