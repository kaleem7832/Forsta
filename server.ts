import express from "express";
import path from "path";
import multer from "multer";
import mammoth from "mammoth";
import { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType } from "docx";
import { GoogleGenAI, Type } from "@google/genai";
import dotenv from "dotenv";

dotenv.config();

const app = express();
const PORT = 3000;

app.use(express.json({ limit: "50mb" }));
app.use(express.urlencoded({ limit: "50mb", extended: true }));

// Configure multer for memory storage
const upload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 20 * 1024 * 1024 }, // 20MB limit
});

// Initialize Gemini Client
const ai = new GoogleGenAI({
  apiKey: process.env.GEMINI_API_KEY,
  httpOptions: {
    headers: {
      "User-Agent": "aistudio-build",
    },
  },
});

// API endpoint to analyze questionnaire
app.post("/api/parse-questionnaire", upload.single("file"), async (req, res): Promise<any> => {
  try {
    if (!req.file) {
      return res.status(400).json({ error: "No file uploaded" });
    }

    const originalName = req.file.originalname;

    // Extract text from the docx file
    let extractedText = "";
    try {
      const result = await mammoth.extractRawText({ buffer: req.file.buffer });
      extractedText = result.value;
    } catch (err: any) {
      console.error("Mammoth error:", err);
      return res.status(400).json({ error: `Failed to extract text from Word document. Please ensure it is a valid .docx file.` });
    }

    if (!extractedText || extractedText.trim().length === 0) {
      return res.status(400).json({ error: "The uploaded document is empty or could not be read." });
    }

    // Prepare system instructions and prompt for Gemini
    const systemInstruction = `You are an expert Forsta survey programming assistant.
Your job is to analyze questionnaire documents, identify questions and their options, detect their types, and convert them to Forsta-compatible survey tagged formats.

Follow these strict rules:
1. Identify Screeners (starting from the beginning) and any Main questions. Ignore any cover letters, descriptions, or introductions that aren't questions.
2. For each question, extract:
  - id: Create an ID.
    - If question numbers are given in numeric value, prepend "S" for screener questions and "Q" for main questions.
    - If question numbers have decimals, replace decimals with "x" (e.g. 1.2 becomes Q1x2).
    - If missing entirely, assign sequential IDs starting from S1 for screeners and Q1 for main questions.
  - type: Must be one of:
    - "single" (Single choose)
    - "multi" (Multiple response)
    - "grid" (Grid question)
    - "numeric" (Numeric response)
    - "numeric list" (Numeric list response)
    - "text" (Open text)
    - "text list" (Open text list)
    - "date" (Date)
    - "ranking" (Ranking)
    - "info" (Informational/Note text)
    - "unknown" (Any unidentified segment)
  - questionText: Clean, human-readable question text. Strip any type labels or tags but keep standard tags if desired.
  - title: The title tag label, e.g. "S1" or "Q1".
  - instruction: The appropriate instruction based on the question type:
    - single: "Please select one response."
    - multi: "Please select all that apply."
    - grid: "Please select one response per row."
    - ranking: "Please rank all that apply." or "Please rank max N, depending on question"
    - numeric: "Please provide the answer."
    - numeric list: "Please provide the answers."
    - text: "Please provide the answer."
    - text list: "Please provide the answers"
    - If instructions are explicitly different in the source, keep the clean version or use these exact templates.
  - isIgnored: Set to true if the question mentions "answer per column". Ignore completely by setting isIgnored to true.
  - options: List of answer options (choices) for single, multi, ranking.
    Apply these Special Option Codes regardless of codes in source document:
    - Other / Other (please specify) must get code 97: "Other (please specify) [[97, other]]"
    - None of the above / None of these / Not applicable / NA must get code 98: "None of the above [[#98, exclusive]]"
    - Don't know / I don't know / DK must get code 99: "Don't know [[#99, exclusive]]"
    - Standard options should just have their text or optionally labels like "* option [[#1]]", etc.
  - gridHeaders: (Only for "grid") Array of scale headings. E.g. ["Scale 1", "Scale 2", "Scale 3"]
  - gridRows: (Only for "grid") Array of row prompt items. E.g. ["Prompt 1", "Prompt 2", "Prompt 3"]
  - tags: List of Forsta tags applied to this question (e.g. ["single", "randomize", "required"]).
    - Apply "randomize" if randomization is mentioned in source.
    - Setup min/max properties based on question (e.g. "min=1", "max=3") if instructions mandate constraint.
  - originalText: A small segment of the original raw text to show as reference.

3. Grid Question Consolidation & Ignore Row Sub-Headings:
  - Do NOT split a grid question (like Q14) into multiple questions. It must remain a single unified question. Discard and ignore any visual row sub-headings, sub-category labels, or program instruction sections inside the grid's list of rows (such as block headings like 'SHOW TO C-SUITE (CODES 1-3 AT S2)', 'DATA/ AI AND SOFTWARE ROLES (CODES 4-13 AT S2)', 'ASK ALL', or logical criteria blocks).
  - Extract only the actual, user-facing prompt rows (e.g. 'A project being abandoned because of...', 'The need to step in to mediate...').
  - Do NOT include those subheading/instruction headers as choices or grid rows.
4. Avoid Headings or Instructions in Options/Rows:
  - In general, never include any question-level instructions, routing conditions, logic comments, or headers in the list of options or rows. Thoroughly clean and discard them. Only keep genuine user-facing response choices.

Analyze the entire text below and output a strictly valid JSON object conforming to the response schema.`;

    const userPrompt = `Here is the raw text of the questionnaire document:
---
${extractedText}
---

Please parse and return a structured JSON list of questions according to the rules. Make sure you tag and structure everything correctly. Do not skip any real questions unless they contain "answer per column".`;

    const response = await ai.models.generateContent({
      model: "gemini-3.5-flash",
      contents: userPrompt,
      config: {
        systemInstruction: systemInstruction,
        responseMimeType: "application/json",
        responseSchema: {
          type: Type.OBJECT,
          required: ["questions"],
          properties: {
            questions: {
              type: Type.ARRAY,
              description: "The list of parsed and tagged questions",
              items: {
                type: Type.OBJECT,
                required: ["id", "type", "questionText", "title", "instruction", "options", "isIgnored"],
                properties: {
                  id: { type: Type.STRING },
                  type: { type: Type.STRING },
                  questionText: { type: Type.STRING },
                  title: { type: Type.STRING },
                  instruction: { type: Type.STRING },
                  options: {
                    type: Type.ARRAY,
                    items: { type: Type.STRING },
                  },
                  gridHeaders: {
                    type: Type.ARRAY,
                    items: { type: Type.STRING },
                  },
                  gridRows: {
                    type: Type.ARRAY,
                    items: { type: Type.STRING },
                  },
                  tags: {
                    type: Type.ARRAY,
                    items: { type: Type.STRING },
                  },
                  isIgnored: { type: Type.BOOLEAN },
                  originalText: { type: Type.STRING },
                },
              },
            },
          },
        },
      },
    });

    const resultText = response.text || "{}";
    const parsedData = JSON.parse(resultText);

    // Helper to identify header rows or programming instructions that are not actual choices
    const isHeadingOrInstruction = (text: string) => {
      if (!text) return false;
      const trimmed = text.trim();
      const upper = trimmed.toUpperCase();

      // Common instruction, logic headers or subheader titles
      if (
        upper.startsWith("SHOW TO ") ||
        upper.startsWith("ASK TO ") ||
        upper.startsWith("ASK ONLY ") ||
        upper.startsWith("ONLY ASK ") ||
        upper.startsWith("ASK IF ") ||
        upper.startsWith("ONLY IF ") ||
        upper.startsWith("ASK ALL ") ||
        upper.startsWith("DATA/ ") ||
        upper === "ASK ALL" ||
        upper === "ASK ONLY" ||
        upper.startsWith("FILTER:") ||
        upper.startsWith("NOTE:") ||
        upper.startsWith("INSTRUCTION:") ||
        upper.startsWith("ROUTING:") ||
        upper.startsWith("PROGRAM ") ||
        upper.includes("SINGLE CODE PER ROW") ||
        upper.includes("SELECT ONE PER ROW") ||
        upper.includes("SELECT ALL THAT APPLY")
      ) {
        return true;
      }

      // Check for code-scoping or targeting indicator keywords typically found in headers (e.g., "(CODES 1-3 AT S2)", "(CODES 4-13)")
      if (
        upper.match(/\(CODES?\s+\d+/) ||
        upper.match(/CODES?\s+\d+\s+AT/) ||
        upper.match(/CODES?\s+\d+-\d+/)
      ) {
        return true;
      }

      // Check for entirely uppercase text lines that look like informational divides (more than 2 words, all caps)
      const words = trimmed.split(/\s+/);
      if (words.length >= 2) {
        const isAllCaps = words.every(w => {
          // Remove non-alphabetic characters
          const letters = w.replace(/[^A-Za-z]/g, "");
          return letters.length === 0 || letters === letters.toUpperCase();
        });
        if (isAllCaps && (upper.includes("ROLE") || upper.includes("CODE") || upper.includes("S2") || upper.includes("S1") || upper.includes("Q"))) {
          return true;
        }
      }

      return false;
    };

    // First merge consecutive grid questions if they are part of a split grid
    const rawQuestions = parsedData.questions || [];
    const mergedQuestions: any[] = [];

    for (let i = 0; i < rawQuestions.length; i++) {
      const current = rawQuestions[i];
      if (!current) continue;

      let mergedAny = false;
      // If of type grid and we have a previous element of type grid, check if we should merge them
      if (current.type === "grid" && mergedQuestions.length > 0) {
        const last = mergedQuestions[mergedQuestions.length - 1];
        if (last.type === "grid") {
          const cleanText = (t: string) => (t || "").toLowerCase().replace(/[^a-z0-9]/g, "");
          const isSameTitle = (last.title && current.title) && (last.title.trim() === current.title.trim() || last.id === current.id);
          const isVerySimilarText = cleanText(last.questionText) === cleanText(current.questionText);

          if (isSameTitle || isVerySimilarText) {
            // Merge their gridRows
            if (Array.isArray(current.gridRows)) {
              if (!Array.isArray(last.gridRows)) last.gridRows = [];
              last.gridRows = [...last.gridRows, ...current.gridRows];
            }
            // Combine headers
            if (Array.isArray(current.gridHeaders) && Array.isArray(last.gridHeaders)) {
              const headerSet = new Set([...last.gridHeaders, ...current.gridHeaders]);
              last.gridHeaders = Array.from(headerSet);
            }
            mergedAny = true;
          }
        }
      }

      if (!mergedAny) {
        mergedQuestions.push(current);
      }
    }

    // Now post-process each of the merged questions (filtering option headings and identifying randomizations)
    const processedQuestions = mergedQuestions.map((q: any) => {
      if (!q.tags) q.tags = [];

      // Filter out headings and program instructions from options and gridRows
      if (Array.isArray(q.options)) {
        q.options = q.options.filter((opt: string) => !isHeadingOrInstruction(opt));
      }
      if (Array.isArray(q.gridRows)) {
        q.gridRows = q.gridRows.filter((row: string) => !isHeadingOrInstruction(row));
      }

      const containsRandomizationKeyword = (text: string) => {
        if (!text) return false;
        const lower = text.toLowerCase();
        return lower.includes("random") || 
               lower.includes("rotate") || 
               lower.includes("rotation") || 
               lower.includes("shuffle");
      };

      let needsRandomize = false;
      if (containsRandomizationKeyword(q.questionText)) needsRandomize = true;
      if (containsRandomizationKeyword(q.instruction)) needsRandomize = true;
      if (containsRandomizationKeyword(q.originalText)) needsRandomize = true;
      if (Array.isArray(q.options) && q.options.some((opt: string) => containsRandomizationKeyword(opt))) needsRandomize = true;
      if (Array.isArray(q.gridRows) && q.gridRows.some((row: string) => containsRandomizationKeyword(row))) needsRandomize = true;

      if (needsRandomize && !q.tags.includes("randomize")) {
        q.tags.push("randomize");
      }
      return q;
    });

    res.json({
      success: true,
      originalName,
      questions: processedQuestions,
      rawText: extractedText,
    });
  } catch (err: any) {
    console.error("API Error in /api/parse-questionnaire:", err);
    res.status(500).json({ error: err.message || "An unexpected error occurred during document parsing." });
  }
});

// Helper to format timestamps as HHMMDDMMYYYY
function getFormattedTimestamp(date: Date) {
  const pad = (n: number) => n.toString().padStart(2, "0");
  const hh = pad(date.getHours());
  const min = pad(date.getMinutes());
  const dd = pad(date.getDate());
  const mm = pad(date.getMonth() + 1); // 0-indexed
  const yyyy = date.getFullYear().toString();
  return `${hh}${min}${dd}${mm}${yyyy}`;
}

// API endpoint to generate docx
app.post("/api/generate-docx", async (req, res): Promise<any> => {
  try {
    const { originalName, questions } = req.body;

    if (!questions || !Array.isArray(questions)) {
      return res.status(400).json({ error: "Invalid questions list provided." });
    }

    const docChildren: any[] = [];

    // Add a title page or brief header
    docChildren.push(
      new Paragraph({
        children: [
          new TextRun({
            text: "FORSTA SURVEY PROGRAMMING QUESTIONNAIRE - AUTO TAGGED",
            bold: true,
            size: 28,
            color: "1F2937",
          }),
        ],
        spacing: { after: 400 },
      })
    );

    for (const q of questions) {
      if (q.isIgnored) continue;

      // Construct Question Header
      // QN. Question text [[type, randomize]]
      // Apply any general purpose and property tags
      let hasRandomize = false;
      const combinedTags: string[] = [];
      if (q.tags && Array.isArray(q.tags)) {
        q.tags.forEach((t: string) => {
          if (t === "randomize") {
            hasRandomize = true;
          } else if (t !== q.type && t !== "required" && t !== "not required") {
            combinedTags.push(t);
          }
        });
      }

      let typeTag = q.type;
      if (q.type === "numeric list") typeTag = "numeric list";
      if (q.type === "text list") typeTag = "text list";

      if (hasRandomize) {
        typeTag = `${typeTag}, randomize`;
      }

      const formattedID = q.id.toLowerCase().endsWith(".") ? q.id : `${q.id}.`;
      const questionLabel = `${formattedID} ${q.questionText} [[${typeTag}]]`;

      docChildren.push(
        new Paragraph({
          children: [
            new TextRun({
              text: questionLabel,
              bold: true,
              size: 24,
            }),
          ],
          spacing: { before: 300, after: 100 },
        })
      );

      // QN [[title]]
      const titleLabel = `${q.id} [[title]] ${combinedTags.map(t => `[[${t}]]`).join(" ")}`.trim();
      docChildren.push(
        new Paragraph({
          children: [
            new TextRun({
              text: titleLabel,
              color: "4B5563",
              size: 20,
            }),
          ],
          spacing: { after: 100 },
        })
      );

      // Instruction based on type [[instruction]]
      if (q.instruction) {
        const instrText = q.instruction.includes("[[instruction]]")
          ? q.instruction
          : `${q.instruction} [[instruction]]`;
        docChildren.push(
          new Paragraph({
            children: [
              new TextRun({
                text: instrText,
                italics: true,
                color: "6B7280",
                size: 20,
              }),
            ],
            spacing: { after: 200 },
          })
        );
      }

      // Format answers or Grid tables
      if (q.type === "grid") {
        const headers = q.gridHeaders || [];
        const rows = q.gridRows || [];

        // Header cells
        const headerCells = [
          new TableCell({
            width: {
              size: 30,
              type: WidthType.PERCENTAGE,
            },
            children: [
              new Paragraph({
                children: [
                  new TextRun({
                    text: "",
                    bold: true,
                    size: 20,
                  }),
                ],
              }),
            ],
          }),
          ...headers.map((hdr: string) => {
            return new TableCell({
              width: {
                size: 70 / Math.max(1, headers.length),
                type: WidthType.PERCENTAGE,
              },
              children: [
                new Paragraph({
                  children: [
                    new TextRun({
                      text: hdr,
                      bold: true,
                      size: 20,
                    }),
                  ],
                }),
              ],
            });
          }),
        ];

        const tableRows = [
          new TableRow({
            children: headerCells,
          }),
        ];

        // Body rows
        for (const rowVal of rows) {
          const cells = [
            new TableCell({
              width: {
                size: 30,
                type: WidthType.PERCENTAGE,
              },
              children: [
                new Paragraph({
                  children: [
                    new TextRun({
                      text: rowVal,
                      size: 20,
                    }),
                  ],
                }),
              ],
            }),
            ...headers.map(() => {
              return new TableCell({
                width: {
                  size: 70 / Math.max(1, headers.length),
                  type: WidthType.PERCENTAGE,
                },
                children: [
                  new Paragraph({
                    children: [],
                  }),
                ],
              });
            }),
          ];

          tableRows.push(
            new TableRow({
              children: cells,
            })
          );
        }

        const gridTable = new Table({
          width: {
            size: 100,
            type: WidthType.PERCENTAGE,
          },
          rows: tableRows,
        });

        docChildren.push(gridTable);

        // spacing after grid
        docChildren.push(
          new Paragraph({
            children: [new TextRun("")],
            spacing: { after: 200 },
          })
        );
      } else {
        // List choices
        const choices = q.options || [];
        for (const choice of choices) {
          // Clean asterisk or hyphen prefix if present, as word bullet lists handle render natively
          let cleanChoice = choice.trim();
          if (cleanChoice.startsWith("* ")) {
            cleanChoice = cleanChoice.substring(2).trim();
          } else if (cleanChoice.startsWith("*")) {
            cleanChoice = cleanChoice.substring(1).trim();
          } else if (cleanChoice.startsWith("- ")) {
            cleanChoice = cleanChoice.substring(2).trim();
          } else if (cleanChoice.startsWith("-")) {
            cleanChoice = cleanChoice.substring(1).trim();
          }

          docChildren.push(
            new Paragraph({
              children: [
                new TextRun({
                  text: cleanChoice,
                  size: 20,
                }),
              ],
              bullet: {
                level: 0,
              },
              spacing: { after: 80 },
            })
          );
        }
      }

      // Single line space split between tagged questions (no divider lines as requested)
      docChildren.push(
        new Paragraph({
          children: [],
          spacing: { after: 180 },
        })
      );
    }

    // Build Word Doc
    const doc = new Document({
      sections: [
        {
          properties: {},
          children: docChildren,
        },
      ],
    });

    const buffer = await Packer.toBuffer(doc);

    // Formulate name: original_name_tagged_HHMMDDMMYYYY
    let baseName = "questionnaire";
    if (originalName) {
      baseName = originalName.replace(/\.[^/.]+$/, ""); // strip extension
    }

    // Use current local time zone timestamp
    const now = new Date();
    const ts = getFormattedTimestamp(now);
    const finalFilename = `${baseName}_tagged_${ts}.docx`;

    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.wordprocessingml.document");
    res.setHeader("Content-Disposition", `attachment; filename="${finalFilename}"`);
    res.setHeader("X-Filename", finalFilename);
    res.send(buffer);
  } catch (err: any) {
    console.error("docx generator error:", err);
    res.status(500).json({ error: err.message || "Failed to generate Word document." });
  }
});

// Global JSON-friendly error Handler
app.use((err: any, req: express.Request, res: express.Response, next: express.NextFunction) => {
  console.error("Global server error caught:", err);
  const status = err.status || err.statusCode || 500;
  res.status(status).json({
    success: false,
    error: err.message || "A server error occurred during processing."
  });
});

// Vite & express logic
async function startServer() {
  if (process.env.NODE_ENV !== "production") {
    const { createServer: createViteServer } = await import("vite");
    const vite = await createViteServer({
      server: { middlewareMode: true },
      appType: "spa",
    });
    app.use(vite.middlewares);
  } else {
    const distPath = path.join(process.cwd(), "dist");
    app.use(express.static(distPath));
    app.get("*", (req, res) => {
      res.sendFile(path.join(distPath, "index.html"));
    });
  }

  app.listen(PORT, "0.0.0.0", () => {
    console.log(`Server running on http://localhost:${PORT}`);
  });
}

startServer();
