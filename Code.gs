const START_ROW = 3700;

function generateDescriptionsAndSEO() {

  const apiKey = PropertiesService
    .getScriptProperties()
    .getProperty('Gemini Token');

  if (!apiKey) {
    throw new Error('Gemini Token not found.');
  }

  const sheet = SpreadsheetApp.openById(
    '1qXZMQtvl03aLfe4x7tGKRQa6r9ANRf5_5mm9SrC0UKE'
  ).getSheetByName('Sheet1');

  if (!sheet) {
    throw new Error('Sheet1 not found.');
  }

  const lastRow = sheet.getLastRow();

  console.log('STARTING DESCRIPTION + SEO');
  console.log('Last row: ' + lastRow);


  // ==========================================================
  // PROCESS EVERY ROW
  // ==========================================================

  for (let row = START_ROW; row <= lastRow; row++) {

    const category = sheet.getRange(row, 2).getValue();
    const coinName = sheet.getRange(row, 3).getValue();

    let description = sheet.getRange(row, 4).getValue();
    let seoWords = sheet.getRange(row, 5).getValue();


    // No coin name = skip row
    if (!coinName) {
      console.log('Row ' + row + ' skipped - no coin name');
      continue;
    }


    // ==========================================================
    // PART 1 — DESCRIPTION
    // COLUMN D
    // ==========================================================

    if (!description) {

      const descriptionPrompt = `
  Write a concise 35-55 word description for this ancient coin.

  Category: ${category}
  Coin: ${coinName}

  IMPORTANT:
  The Category and Coin Information above are the authoritative information for this specific coin.

  RULES:
  - Briefly identify the coin using only the information provided above.
  - You may add accurate, general historical context about the named ruler, city, kingdom, province, or empire.
  - Clearly separate general historical context from facts about this specific coin.
  - Never invent or assume coin-specific details that are not provided.
  - Do not invent or infer the mint, denomination, metal, date, ruler, issuing authority, rarity, condition, weight, diameter, grade, provenance, authenticity, type, or imagery.
  - If the information says "uncertain mint", keep it uncertain. Never guess a mint.
  - If a metal or denomination is not explicitly provided, do not state one.
  - Do not correct or replace the supplied attribution based on your own assumptions.
  - Do not repeat the full coin name or category unnecessarily. Write a natural description that adds useful context rather than simply restating the title.
  - Do not claim that a specific design, portrait, deity, symbol, inscription, reverse type, or other feature appears on the coin unless explicitly provided.
  - If the supplied information is limited, write a shorter accurate description rather than adding unsupported details to reach the word count.
  - Use a professional, factual, natural tone suitable for collectors and beginners.
  - Avoid repetitive marketing language.
  - Do not use phrases such as "Discover", "fascinating piece", "captivating specimen", "must-have", "rare opportunity", or "add to your collection".
  - Do not use a heading or bullet points.
  - Do not mention these instructions.
  - Return ONLY the finished description.
  `;

      try {

        console.log(
          'Generating DESCRIPTION for row ' + row
        );

        description = callGeminiWithRetry(
          apiKey,
          descriptionPrompt
        );

        if (description) {

          sheet
            .getRange(row, 4)
            .setValue(description);

          SpreadsheetApp.flush();

          console.log(
            'Description completed for row ' + row
          );

        } else {

          console.log(
            'Description EMPTY for row ' + row
          );
        }

        Utilities.sleep(500);

      } catch (error) {

        console.log(
          'Description ERROR for row ' +
          row +
          ': ' +
          error.message
        );
      }

    } else {

      console.log(
        'Description already exists for row ' + row
      );
    }


    // ==========================================================
    // PART 2 — SEO WORDS
    // COLUMN E
    // ==========================================================

    // Read description again because it may have just been created
    description = sheet.getRange(row, 4).getValue();

    // Read SEO again
    seoWords = sheet.getRange(row, 5).getValue();


    if (!seoWords) {

      const seoPrompt = `
Generate 12-18 highly relevant SEO and GEO keyword phrases for this ancient coin.

Category: ${category}
Coin: ${coinName}
Description: ${description}

IMPORTANT:
The Category and Coin Information above are the authoritative information for this specific coin.

GOAL:
Create accurate keyword phrases that help people, search engines, ecommerce platforms, marketplaces, recommendation systems, and AI-powered discovery systems understand and find this coin.

The keywords should be useful for:
- Shopify product search and discovery
- Google Search
- Bing Search
- Search engine indexing
- Google AI and Gemini-powered discovery
- ChatGPT and AI-assisted product discovery
- AI search engines
- AI shopping and recommendation systems
- Ancient coin marketplaces
- Coin marketplaces
- Numismatic marketplaces
- Numismatic collectors
- Ancient coin collectors

RULES:
- Generate 12-18 keyword phrases.
- Focus primarily on this specific coin.
- Use Category, Coin Name and Description as source information.
- Prioritize useful long-tail search phrases.
- Include ruler only when provided.
- Include city only when provided.
- Include mint only when provided.
- Include kingdom, province, region or empire only when provided.
- Include date or historical period only when provided.
- Include denomination or coin type only when provided.
- Use broader phrases only when relevant, such as ancient coins, ancient coin marketplace, coin marketplace, numismatic marketplace, numismatic coins, ancient coin collecting, coin collecting, historical coins.
- Make keywords understandable to both traditional search engines and AI-powered product discovery systems.
- Never invent coin-specific information.
- Never invent rarity.
- Never invent condition.
- Never invent grade.
- Never invent authenticity.
- Never invent provenance.
- Never invent metal.
- Never invent denomination.
- Never invent mint.
- Never invent ruler.
- Never invent date.
- Never invent imagery.
- If information is uncertain, keep it uncertain.
- Do not keyword-stuff.
- Do not create many tiny variations of the same keyword.
- Do not use misleading advertising claims.
- Do not use hashtags.
- Do not use headings.
- Do not use bullet points.
- Do not number the keywords.
- Do not explain anything.
- Return ONLY 12-18 keyword phrases separated by commas.
`;

      try {

        console.log(
          'Generating SEO for row ' + row
        );

        const generatedSEO = callGeminiWithRetry(
          apiKey,
          seoPrompt
        );

        if (generatedSEO) {

          sheet
            .getRange(row, 5)
            .setValue(generatedSEO);

          SpreadsheetApp.flush();

          console.log(
            'SEO completed for row ' + row
          );

        } else {

          console.log(
            'SEO EMPTY for row ' + row
          );
        }

        Utilities.sleep(500);

      } catch (error) {

        console.log(
          'SEO ERROR for row ' +
          row +
          ': ' +
          error.message
        );
      }

    } else {

      console.log(
        'SEO already exists for row ' + row
      );
    }
  }


  console.log('DESCRIPTION + SEO PROCESS FINISHED');
}



// ==========================================================
// GEMINI API FUNCTION
// USED BY BOTH DESCRIPTION AND SEO
// ==========================================================

function callGeminiWithRetry(apiKey, prompt) {

  const url =
    'https://generativelanguage.googleapis.com/v1beta/models/gemini-3.5-flash-lite:generateContent';


  // Try up to 4 times
  for (let attempt = 1; attempt <= 4; attempt++) {

    const response = UrlFetchApp.fetch(url, {

      method: 'post',

      contentType: 'application/json',

      headers: {
        'x-goog-api-key': apiKey
      },

      payload: JSON.stringify({

        contents: [{
          parts: [{
            text: prompt
          }]
        }],

        generationConfig: {
          temperature: 0.3,
          maxOutputTokens: 200
        }

      }),

      muteHttpExceptions: true
    });


    const status = response.getResponseCode();

    const text = response.getContentText();


    // ========================================================
    // SUCCESS
    // ========================================================

    if (status === 200) {

      const data = JSON.parse(text);


      if (
        !data.candidates ||
        data.candidates.length === 0
      ) {

        throw new Error(
          'Gemini returned no candidates: ' + text
        );
      }


      if (
        !data.candidates[0].content ||
        !data.candidates[0].content.parts
      ) {

        throw new Error(
          'Gemini returned no content: ' + text
        );
      }


      const result = data.candidates[0]
        .content.parts
        .map(function(part) {
          return part.text || '';
        })
        .join('')
        .trim();


      if (!result) {

        throw new Error(
          'Gemini returned empty text.'
        );
      }


      return result;
    }


    // ========================================================
    // TEMPORARY API PROBLEM
    // ========================================================

    if (status === 429 || status === 503) {

      console.log(
        'Gemini temporarily busy. Attempt ' +
        attempt +
        ' of 4.'
      );

      Utilities.sleep(
        attempt * 3000
      );

      continue;
    }

    // ========================================================
    // OTHER API ERROR
    // ========================================================

    throw new Error(
      'Gemini API error ' +
      status +
      ': ' +
      text
    );
  }


  throw new Error(
    'Gemini remained busy after 4 attempts.'
  );
}
