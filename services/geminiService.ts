

import { GoogleGenAI, Type } from "@google/genai";

if (!process.env.API_KEY) {
    throw new Error("API_KEY environment variable is not set.");
}

const ai = new GoogleGenAI({ apiKey: process.env.API_KEY });

const responseSchema = {
    type: Type.OBJECT,
    properties: {
        sheets: {
            type: Type.ARRAY,
            description: "An array of sheets found in the document. Each sheet contains tabular data.",
            items: {
                type: Type.OBJECT,
                properties: {
                    sheetName: {
                        type: Type.STRING,
                        description: "A descriptive name for the sheet, e.g., 'Summary Q1' or 'Page 3 Table 1'."
                    },
                    data: {
                        type: Type.ARRAY,
                        description: "The tabular data, represented as an array of arrays. The first inner array MUST be the headers. Subsequent arrays are data rows.",
                        items: {
                            type: Type.ARRAY,
                            items: {
                                type: Type.STRING,
                                description: "A single cell value, represented as a string."
                            }
                        }
                    }
                },
                required: ["sheetName", "data"],
            }
        }
    },
    required: ["sheets"],
};

export const extractDataFromPdfImages = async (images: string[]): Promise<{sheetName: string; data: string[][]}[]> => {
    const imageParts = images.map(imgBase64 => ({
        inlineData: {
            mimeType: 'image/jpeg',
            data: imgBase64,
        },
    }));

    const textPart = {
        text: `Analyze the following document pages. Extract all tabular data you can find.
        If a single table spans multiple pages, consolidate it into one sheet.
        Return the result as a JSON object that strictly follows the provided schema.
        Ensure the first inner array in the 'data' field for each sheet contains the column headers.
        If there are no clear headers, infer them from the data content (e.g., 'Column 1', 'Product_ID', 'Date').
        Do not return empty sheets or sheets without data rows.`
    };

    try {
        const response = await ai.models.generateContent({
            model: 'gemini-2.5-flash',
            contents: { parts: [textPart, ...imageParts] },
            config: {
                responseMimeType: "application/json",
                responseSchema: responseSchema,
                temperature: 0.1,
            },
        });

        if (response.candidates?.[0]?.finishReason === 'SAFETY') {
            throw new Error("The response was blocked for safety reasons. The document may contain sensitive content.");
        }

        const jsonString = response.text;
        if (!jsonString) {
             throw new Error("The AI returned an empty response. No data could be extracted.");
        }

        let result;
        try {
            result = JSON.parse(jsonString);
        } catch (parseError) {
            throw new Error("The AI returned a malformed data structure that could not be read.");
        }
        
        if (!result.sheets || !Array.isArray(result.sheets)) {
            throw new Error("The AI response was missing the expected 'sheets' data structure.");
        }
        
        if (result.sheets.length === 0) {
            throw new Error("No tables were detected in the document.");
        }
        
        const validSheets = result.sheets.filter((sheet: {data?: any[][]}) => 
            sheet.data && sheet.data.length > 1 && sheet.data[0] && sheet.data[0].length > 0
        );

        if (validSheets.length === 0) {
            throw new Error("The detected tables appear to be empty or contain only headers.");
        }

        return validSheets;

    } catch (error) {
        console.error("Error extracting data from PDF:", error);
        if (error instanceof Error) {
            throw error;
        }
        throw new Error("An unknown error occurred during data extraction.");
    }
};
