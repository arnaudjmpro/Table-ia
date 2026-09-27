const corsHeaders = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Headers":
    "authorization, x-client-info, apikey, content-type",
  "Access-Control-Allow-Methods": "POST, OPTIONS",
};

function jsonResponse(data: unknown, status = 200) {
  return new Response(JSON.stringify(data), {
    status,
    headers: {
      ...corsHeaders,
      "Content-Type": "application/json",
    },
  });
}

const tableSchema = {
  type: "object",
  additionalProperties: false,
  required: [
    "title",
    "columns",
    "rows",
    "summaryFormulas",
  ],
  properties: {
    title: {
      type: "string",
    },

    columns: {
      type: "array",
      items: {
        type: "object",
        additionalProperties: false,
        required: [
          "name",
          "type",
          "formula",
        ],
        properties: {
          name: {
            type: "string",
          },

          type: {
            type: "string",
            enum: [
              "text",
              "number",
              "currency",
              "percentage",
              "date",
              "status",
              "boolean",
            ],
          },

          formula: {
            type: "string",
          },
        },
      },
    },

    rows: {
      type: "array",
      items: {
        type: "array",
        items: {
          type: "string",
        },
      },
    },

    summaryFormulas: {
      type: "array",
      items: {
        type: "object",
        additionalProperties: false,
        required: [
          "label",
          "targetColumn",
          "formula",
        ],
        properties: {
          label: {
            type: "string",
          },

          targetColumn: {
            type: "string",
          },

          formula: {
            type: "string",
          },
        },
      },
    },
  },
};

Deno.serve(async (req) => {
  if (req.method === "OPTIONS") {
    return new Response("ok", {
      headers: corsHeaders,
    });
  }

  if (req.method !== "POST") {
    return jsonResponse(
      {
        error: "Méthode non autorisée",
      },
      405,
    );
  }

  try {
    const openAiKey =
      Deno.env.get("OPENAI_API_KEY");

    if (!openAiKey) {
      return jsonResponse(
        {
          error:
            "OPENAI_API_KEY absente dans les secrets Supabase",
        },
        500,
      );
    }

    let body: any;

    try {
      body = await req.json();
    } catch {
      return jsonResponse(
        {
          error:
            "Le corps de la requête doit être un JSON valide",
        },
        400,
      );
    }

    const prompt =
      String(body?.prompt ?? "").trim();
const imageBase64 =
  String(body?.imageBase64 ?? "").trim();

  const imageMimeType =
    String(body?.imageMimeType ?? "image/jpeg").trim();
    if (!prompt && !imageBase64) {
      return jsonResponse(
        {
          error:
            "Le champ prompt ou une image est obligatoire",
        },
        400,
      );
    }

    const instructions = `
Tu es le moteur IA de l'application TableIA.

Ta mission est de transformer la demande de l'utilisateur
en un véritable outil de gestion de type Excel.

Tu dois créer un tableau professionnel, cohérent,
simple à utiliser et adapté à la demande.

RÈGLES POUR LES COLONNES

Chaque colonne contient obligatoirement :

- name
- type
- formula

Types autorisés :

text
number
currency
percentage
date
status
boolean

Si une colonne n'est pas calculée :

formula doit être une chaîne vide "".

Si une colonne est calculée automatiquement,
formula doit contenir une vraie formule Excel.

IMPORTANT :

Les formules Excel doivent être écrites
SANS le signe = au début.

Utilise les noms de fonctions Excel internes
en anglais.

Utilise des virgules comme séparateurs
dans les fonctions Excel.

Pour une formule appliquée à toutes les lignes,
utilise obligatoirement :

{row}

Exemples :

B{row}-C{row}

IFERROR(C{row}/B{row},0)

IF(D{row}>0,"Payé","À payer")

DATEDIF(B{row},TODAY(),"Y")

ROUND(B{row}*C{row},2)

SUM(B{row}:D{row})

Tu peux utiliser les fonctions Excel standards,
notamment :

SUM
AVERAGE
MIN
MAX
COUNT
COUNTA
IF
IFS
AND
OR
NOT
ROUND
ROUNDUP
ROUNDDOWN
SUMIF
SUMIFS
COUNTIF
COUNTIFS
AVERAGEIF
AVERAGEIFS
IFERROR
INDEX
MATCH
XMATCH
VLOOKUP
HLOOKUP
XLOOKUP
LEFT
RIGHT
MID
LEN
TRIM
CONCAT
CONCATENATE
TEXTJOIN
UPPER
LOWER
PROPER
SUBSTITUTE
REPLACE
FIND
SEARCH
TODAY
NOW
DATE
YEAR
MONTH
DAY
DATEDIF
NETWORKDAYS
WORKDAY
WEEKDAY
EOMONTH
ABS
MOD
POWER
SQRT

et les autres fonctions Excel classiques
nécessaires à la demande.

N'utilise jamais :

- macro VBA
- script
- DDE
- WEBSERVICE
- RTD
- lien Internet
- lien externe vers un autre fichier Excel.

RÈGLES POUR LES LIGNES

Chaque ligne doit posséder exactement
le même nombre de valeurs que le nombre de colonnes.

Pour une colonne calculée automatiquement,
mets une chaîne vide "" dans rows.

Si l'utilisateur fournit des données,
intègre-les.

S'il demande simplement de créer un outil
sans fournir de données,
crée trois lignes vides prêtes à être remplies.

RÈGLES POUR LES TOTAUX ET INDICATEURS

Utilise summaryFormulas pour créer
les totaux, moyennes, compteurs ou indicateurs
utiles au tableau.

Chaque élément contient exactement :

label
targetColumn
formula

targetColumn doit correspondre exactement
au nom d'une colonne existante.

Pour désigner la dernière ligne de données,
utilise obligatoirement :

{lastRow}

Exemples :

SUM(B2:B{lastRow})

AVERAGE(E2:E{lastRow})

COUNTIF(D2:D{lastRow},"Payé")

SUMIFS(C2:C{lastRow},D2:D{lastRow},"Payé")

Si aucun résumé n'est pertinent,
summaryFormulas doit être un tableau vide [].

Le titre et les noms des colonnes
doivent être en français.

Ne renvoie aucune explication.
Ne renvoie aucun Markdown.
Ne renvoie aucun bloc de code.
Renvoie uniquement la structure JSON demandée.
`;
const userContent: any[] = [];

if (prompt) {
  userContent.push({
      type: "input_text",
          text: prompt,
            });
            } else {
              userContent.push({
                  type: "input_text",
                      text: "Analyse cette image et transforme les informations visibles en un tableau TableIA structuré.",
                        });
                        }

                        if (imageBase64) {
                          userContent.push({
                              type: "input_image",
                                  image_url: `data:${imageMimeType};base64,${imageBase64}`,
                                      detail: "high",
                                        });
                                        }
    const openAiResponse =
      await fetch(
        "https://api.openai.com/v1/responses",
        {
          method: "POST",

          headers: {
            "Authorization":
              `Bearer ${openAiKey}`,
            "Content-Type":
              "application/json",
          },

          body: JSON.stringify({
            model: "gpt-5.6-sol",

            store: false,

            reasoning: {
              effort: "low",
            },

            max_output_tokens: 5000,

input: [
    {
        role: "developer",
            content: [
                  {
                          type: "input_text",
                                  text: instructions,
                                        },
                                            ],
                                              },
                                                {
                                                    role: "user",
                                                        content: userContent,
                                                          },
                                                          ],

            text: {
              format: {
                type: "json_schema",
                name: "tableia_tool",
                description:
                  "Structure d'un outil Excel généré par TableIA",
                strict: true,
                schema: tableSchema,
              },
            },
          }),
        },
      );

    const rawResponse =
      await openAiResponse.text();

    if (!openAiResponse.ok) {
      return jsonResponse(
        {
          error: "Erreur OpenAI",
          status: openAiResponse.status,
          details: rawResponse,
        },
        502,
      );
    }

    let openAiData: any;

    try {
      openAiData =
        JSON.parse(rawResponse);
    } catch {
      return jsonResponse(
        {
          error:
            "Réponse OpenAI illisible",
          details: rawResponse,
        },
        502,
      );
    }

    let outputText = "";

    for (
      const outputItem of
      openAiData.output ?? []
    ) {
      for (
        const contentItem of
        outputItem.content ?? []
      ) {
        if (
          contentItem.type ===
            "output_text" &&
          typeof contentItem.text ===
            "string"
        ) {
          outputText +=
            contentItem.text;
        }
      }
    }

    if (!outputText.trim()) {
      return jsonResponse(
        {
          error:
            "OpenAI n'a renvoyé aucun résultat exploitable",
          details:
            openAiData.error ??
            openAiData.status ??
            "Aucun output_text",
        },
        502,
      );
    }

    let result: any;

    try {
      result =
        JSON.parse(outputText);
    } catch {
      const cleaned =
        outputText
          .replace(/```json/gi, "")
          .replace(/```/g, "")
          .trim();

      const firstBrace =
        cleaned.indexOf("{");

      const lastBrace =
        cleaned.lastIndexOf("}");

      if (
        firstBrace === -1 ||
        lastBrace === -1 ||
        lastBrace <= firstBrace
      ) {
        return jsonResponse(
          {
            error:
              "Le résultat IA ne contient pas de JSON valide",
            details: outputText,
          },
          502,
        );
      }

      try {
        result =
          JSON.parse(
            cleaned.slice(
              firstBrace,
              lastBrace + 1,
            ),
          );
      } catch {
        return jsonResponse(
          {
            error:
              "Impossible de décoder le JSON généré par l'IA",
            details: outputText,
          },
          502,
        );
      }
    }

    if (
      Array.isArray(result.columns)
    ) {
      result.columns =
        result.columns.map(
          (column: any) => ({
            ...column,
            formula:
              String(
                column.formula ?? "",
              )
                .trim()
                .replace(/^=/, ""),
          }),
        );
    }

    if (
      Array.isArray(
        result.summaryFormulas,
      )
    ) {
      result.summaryFormulas =
        result.summaryFormulas.map(
          (summary: any) => ({
            ...summary,
            formula:
              String(
                summary.formula ?? "",
              )
                .trim()
                .replace(/^=/, ""),
          }),
        );
    }

    return jsonResponse(result);
  } catch (error) {
    return jsonResponse(
      {
        error:
          "Erreur interne TableIA",
        details:
          error instanceof Error
            ? error.message
            : String(error),
      },
      500,
    );
  }
});