
import { GoogleGenAI, Type, HarmCategory, HarmBlockThreshold } from "@google/genai";
import { Student, Grade, Anecdote, AIAnalysisResult, ExtractedGrade, ExtractedStudentData, DlpContent, GeneratedQuiz, QuizType, DlpRubricItem, DllContent, AttendanceStatus, DlpProcedure, LearningActivitySheet, CotLessonPlan, CotProcedureStep, ExamObjective, GeneratedExam, GeneratedQuizSection } from '../types';
import { toast } from "react-hot-toast";

const ai = new GoogleGenAI({ apiKey: process.env.API_KEY });

const safetySettings = [
  { category: HarmCategory.HARM_CATEGORY_HARASSMENT, threshold: HarmBlockThreshold.BLOCK_MEDIUM_AND_ABOVE },
  { category: HarmCategory.HARM_CATEGORY_HATE_SPEECH, threshold: HarmBlockThreshold.BLOCK_MEDIUM_AND_ABOVE },
  { category: HarmCategory.HARM_CATEGORY_SEXUALLY_EXPLICIT, threshold: HarmBlockThreshold.BLOCK_MEDIUM_AND_ABOVE },
  { category: HarmCategory.HARM_CATEGORY_DANGEROUS_CONTENT, threshold: HarmBlockThreshold.BLOCK_MEDIUM_AND_ABOVE },
];

const efficientGenerationSystemInstruction = "You are a helpful AI assistant for teachers. Provide concise, accurate, and JSON-formatted responses where requested.";

const handleGeminiError = (error: any, context: string) => {
    console.error(`Error in ${context}:`, error);
    return new Error(`AI Error (${context}): ${error.message || 'Unknown error'}`);
};

const parseJsonFromAiResponse = <T>(text: string): T => {
    try {
        let cleanText = text.replace(/```json/g, '').replace(/```/g, '').trim();
        return JSON.parse(cleanText) as T;
    } catch (e) {
        console.error("JSON Parse Error:", e, text);
        throw new Error("Failed to parse AI response as JSON.");
    }
};

const callApiProxy = async (params: { model: string, contents: any, config?: any, systemInstruction?: string }) => {
    const modelId = params.model;
    const config = {
        ...params.config,
        systemInstruction: params.systemInstruction,
        safetySettings: safetySettings,
    };
    
    try {
        const response = await ai.models.generateContent({
            model: modelId,
            contents: params.contents,
            config: config
        });
        return response;
    } catch (e) {
        throw e;
    }
};

// ... (Existing exported functions like getInspirationalQuote, etc., assuming they remain unchanged)
export const getInspirationalQuote = async (): Promise<{ quote: string; author: string }> => {
    // ... implementation ...
    const model = "gemini-2.5-flash";
    const prompt = "Give me an inspirational quote for a teacher. Return JSON: { \"quote\": \"...\", \"author\": \"...\" }";
    try { const response = await callApiProxy({ model, contents: prompt, config: { responseMimeType: "application/json" } }); return parseJsonFromAiResponse<{ quote: string; author: string }>(response.text || "{}"); } catch (error) { throw handleGeminiError(error, 'getInspirationalQuote'); }
};

export const processAttendanceCommand = async (command: string, students: Student[]): Promise<{ studentIds: string[], status: AttendanceStatus } | null> => {
    // ... implementation ...
    const model = "gemini-2.5-flash";
    const studentList = students.map(s => `${s.id}: ${s.firstName} ${s.lastName}`).join('\n');
    const prompt = `Students:\n${studentList}\n\nCommand: "${command}"\n\nIdentify which students are mentioned or implied in the command and the attendance status (present, absent, late).\nIf "everyone" or similar is mentioned, include all IDs.\nReturn JSON: { "studentIds": ["id1", "id2"], "status": "present" | "absent" | "late" }`;
    try { const response = await callApiProxy({ model, contents: prompt, config: { responseMimeType: "application/json" } }); return parseJsonFromAiResponse(response.text || "null"); } catch (error) { throw handleGeminiError(error, 'processAttendanceCommand'); }
};

export const analyzeStudentPerformance = async (students: Student[], grades: Grade[], anecdotes: Anecdote[]): Promise<AIAnalysisResult[]> => {
    // ... implementation ...
    const model = "gemini-2.5-flash"; const dataSummary = students.map(s => { const sGrades = grades.filter(g => g.studentId === s.id).map(g => `${g.subject} (${g.type}): ${g.score}/${g.maxScore}`); const sAnecdotes = anecdotes.filter(a => a.studentId === s.id).map(a => a.observation); return { name: `${s.firstName} ${s.lastName}`, grades: sGrades, anecdotes: sAnecdotes }; }); const prompt = `Analyze the performance of the following students based on their grades and anecdotal records.\nIdentify meaningful trends (improving, declining, struggling in specific areas, excelling).\nProvide a short summary and a specific recommendation for each student who has notable data.\n\nData: ${JSON.stringify(dataSummary)}\n\nReturn JSON array:\n[\n  { "studentName": "Name", "trendSummary": "...", "recommendation": "..." }\n]`; try { const response = await callApiProxy({ model, contents: prompt, config: { responseMimeType: "application/json" } }); return parseJsonFromAiResponse(response.text || "[]"); } catch (error) { throw handleGeminiError(error, 'analyzeStudentPerformance'); }
};

export const extractGradesFromImage = async (base64Image: string, students: Student[]): Promise<ExtractedGrade[]> => {
    // ... implementation ...
    const model = "gemini-2.5-flash"; const studentNames = students.map(s => `${s.firstName} ${s.lastName}`).join(', '); const prompt = `Extract student names and their scores from this image of a grade sheet.\nThe known students in this class are: ${studentNames}.\nTry to match extracted names to the known list.\nReturn JSON array: [{ "studentName": "matched name", "score": number, "maxScore": number }]\nIf max score is not visible, guess based on the highest possible score or typical values (e.g. 10, 20, 50, 100).`; try { const response = await callApiProxy({ model, contents: { parts: [ { inlineData: { mimeType: "image/jpeg", data: base64Image } }, { text: prompt } ] }, config: { responseMimeType: "application/json" } }); return parseJsonFromAiResponse(response.text || "[]"); } catch (error) { throw handleGeminiError(error, 'extractGradesFromImage'); }
};

export const rephraseAnecdote = async (text: string, mode: 'correct' | 'rephrase'): Promise<string> => {
    // ... implementation ...
    const model = "gemini-2.5-flash"; const prompt = mode === 'correct' ? `Correct the grammar and spelling of this teacher's observation: "${text}"` : `Rephrase this teacher's observation to be more professional, objective, and constructive: "${text}"`; try { const response = await callApiProxy({ model, contents: prompt }); return response.text || text; } catch (error) { throw handleGeminiError(error, 'rephraseAnecdote'); }
};

export const generateReportCardComment = async (student: Student, grades: Grade[], anecdotes: Anecdote[]): Promise<{ strengths: string, areasForImprovement: string, closingStatement: string }> => {
    // ... implementation ...
    const model = "gemini-2.5-flash"; const prompt = `Generate a report card comment for ${student.firstName} ${student.lastName}.\nGrades: ${JSON.stringify(grades)}\nAnecdotes: ${JSON.stringify(anecdotes)}\n\nReturn JSON: { "strengths": "...", "areasForImprovement": "...", "closingStatement": "..." }`; try { const response = await callApiProxy({ model, contents: prompt, config: { responseMimeType: "application/json" } }); return parseJsonFromAiResponse(response.text || "{}"); } catch (error) { throw handleGeminiError(error, 'generateReportCardComment'); }
};

export const generateCertificateContent = async (formData: { awardTitle: string; tone: string; achievements: string }): Promise<string> => {
    // ... implementation ...
    const model = "gemini-2.5-flash"; const prompt = `Write the content body for a certificate.\nAward: ${formData.awardTitle}\nTone: ${formData.tone}\nAchievements to mention: ${formData.achievements}\n\nUse placeholders like [STUDENT_NAME], [DATE], etc.\nFormat with markdown for bolding key phrases.\nJust return the body text.`; try { const response = await callApiProxy({ model, contents: prompt }); return response.text || ""; } catch (error) { throw handleGeminiError(error, 'generateCertificateContent'); }
};

export const generateDlpContent = async (options: any): Promise<DlpContent> => {
    // ... implementation ...
    const model = "gemini-2.5-flash"; const prompt = `Generate a detailed Daily Lesson Plan (DLP) for:\nSubject: ${options.subject}\nGrade: ${options.gradeLevel}\nCompetency: ${options.learningCompetency}\nObjective: ${options.lessonObjective}\nFormat: ${options.dlpFormat}\nLanguage: ${options.language}\n\nReturn JSON matching this schema:\n{\n    "contentStandard": "string",\n    "performanceStandard": "string",\n    "topic": "string",\n    "learningReferences": "string",\n    "learningMaterials": "string",\n    "procedures": [ { "title": "string", "content": "string", "ppst": "string" } ],\n    "evaluationQuestions": [ { "question": "string", "options": ["string"], "answer": "string" } ],\n    "remarksContent": "string"\n}`; try { const response = await callApiProxy({ model, contents: prompt, config: { responseMimeType: "application/json" } }); return parseJsonFromAiResponse(response.text || "{}"); } catch (error) { throw handleGeminiError(error, 'generateDlpContent'); }
};

export const generateQuizContent = async (options: any): Promise<GeneratedQuiz> => {
    // ... implementation ...
    const model = "gemini-2.5-flash"; const prompt = `Generate a quiz.\nTopic: ${options.topic}\nSubject: ${options.subject}\nGrade: ${options.gradeLevel}\nQuestions per type: ${options.numQuestions}\nTypes: ${options.quizTypes.join(', ')}\n\nReturn JSON:\n{\n    "quizTitle": "string",\n    "questionsByType": {\n        "Multiple Choice": { "instructions": "...", "questions": [{ "questionText": "...", "options": ["..."], "correctAnswer": "..." }] },\n        // ... other types\n    },\n    "activities": [\n        { "activityName": "string", "activityInstructions": "string" } \n    ],\n    "tableOfSpecifications": [\n         { "objective": "string", "cognitiveLevel": "string", "itemNumbers": "string" }\n    ]\n}`; try { const response = await callApiProxy({ model, contents: prompt, config: { responseMimeType: "application/json" } }); return parseJsonFromAiResponse(response.text || "{}"); } catch (error) { throw handleGeminiError(error, 'generateQuizContent'); }
};

export const generateRubricForActivity = async (options: { activityName: string, activityInstructions: string, totalPoints: number }): Promise<DlpRubricItem[]> => {
    // ... implementation ...
    const model = "gemini-2.5-flash"; const prompt = `Create a rubric for activity: "${options.activityName}"\nInstructions: "${options.activityInstructions}"\nTotal Points: ${options.totalPoints}\n\nReturn JSON array: [{ "criteria": "string", "points": number }]`; try { const response = await callApiProxy({ model, contents: prompt, config: { responseMimeType: "application/json" } }); return parseJsonFromAiResponse(response.text || "[]"); } catch (error) { throw handleGeminiError(error, 'generateRubricForActivity'); }
};

export const generateDllContent = async (options: any): Promise<DllContent> => {
    // ... implementation ...
    const model = "gemini-2.5-flash"; const prompt = `Generate a Daily Lesson Log (DLL) for one week.\nSubject: ${options.subject}\nGrade: ${options.gradeLevel}\nTopic: ${options.weeklyTopic}\nLanguage: ${options.language}\n\nReturn JSON structure matching DllContent interface.`; const schemaHint = `\n    {\n        "contentStandard": "", "performanceStandard": "",\n        "learningCompetencies": { "monday": "", "tuesday": "", "wednesday": "", "thursday": "", "friday": "" },\n        "content": "",\n        "learningResources": { ... },\n        "procedures": [ { "procedure": "Review...", "monday": "", "tuesday": "", ... } ],\n        "remarks": "", "reflection": []\n    }\n    `; try { const response = await callApiProxy({ model, contents: prompt + "\nJSON Schema:" + schemaHint, config: { responseMimeType: "application/json" } }); return parseJsonFromAiResponse(response.text || "{}"); } catch (error) { throw handleGeminiError(error, 'generateDllContent'); }
};

// UPDATED LAS GENERATOR FUNCTION
export const generateLearningActivitySheet = async (options: { subject: string; gradeLevel: string; learningCompetency: string; lessonObjective: string; activityType: string; language: 'English' | 'Filipino' }): Promise<LearningActivitySheet> => {
    const { subject, gradeLevel, learningCompetency, lessonObjective, activityType, language } = options;
    const model = "gemini-3-pro-preview"; // Use stronger model for complex structure
    const prompt = `
        You are an expert curriculum designer for the Department of Education in the Philippines. Create a comprehensive, 5-day learning packet following the DLP-style Learning Activity Sheet (LAS) format. 
        The entire week must focus on scaffolding the single provided learning competency, with each day building upon the previous one.

        **Main Learning Competency:** ${learningCompetency}
        **Main Lesson Objective:** ${lessonObjective}
        **Activity Focus for the week:** ${activityType}
        **Language:** ${language}

        **CRITICAL INSTRUCTIONS FOR EACH OF THE 5 DAYS:**
        1.  **days Array:** The output JSON must contain a root key "days" which is an array of exactly 5 objects.
        2.  **dayTitle:** Create a clear title (e.g., "Day 1: Understanding Bias").
        3.  **activityTitle:** Specific title for the activity sheet.
        4.  **learningTarget:** Specific sub-objective for the day.
        5.  **references:** Suggest learning materials.
        6.  **conceptNotes:** An ARRAY of objects, each with "title" and "content". "title" must be "CONCEPT NOTES". Content uses markdown (**bold**, *italics*) for emphasis.
        7.  **activities:** An ARRAY of objects. "title" must start with "ACTIVITY:". "instructions" contains directions. 
            - **IMPORTANT:** For matching activities or tables, separate columns with " || " in the instructions string (e.g., "Item 1 || Answer A").
        8.  **reflection:** A single thought-provoking question string.

        **Output Schema (Strict JSON):**
        {
          "days": [
            {
              "dayTitle": "Day 1: Title",
              "activityTitle": "Activity Title",
              "learningTarget": "Target",
              "references": "Ref",
              "conceptNotes": [ { "title": "CONCEPT NOTES", "content": "Content here..." } ],
              "activities": [ 
                  { 
                    "title": "ACTIVITY: NAME", 
                    "instructions": "Instructions here...", 
                    "questions": [ { "questionText": "Q1", "options": ["A","B"] } ] 
                  } 
              ],
              "reflection": "Reflection question..."
            }
            // ... 5 days total
          ]
        }
    `;
    try {
        const response = await callApiProxy({
            model, contents: prompt,
            config: { responseMimeType: "application/json" },
            systemInstruction: efficientGenerationSystemInstruction
        });
        return parseJsonFromAiResponse<LearningActivitySheet>(response.text || "{}");
    } catch (error) {
        throw handleGeminiError(error, 'generateLearningActivitySheet');
    }
};

export const generateExam = async (options: { objectives: { text: string, days: string }[], subject: string, gradeLevel: string, quarter: string }): Promise<GeneratedExam> => {
    // ... implementation ...
    const model = "gemini-3-pro-preview"; const prompt = `Create a 50-item Periodical Exam with Table of Specifications (TOS).\nSubject: ${options.subject}\nGrade: ${options.gradeLevel}\nQuarter: ${options.quarter}\nObjectives and Days Taught: ${JSON.stringify(options.objectives)}\n\nReturn JSON:\n{\n    "title": "First Periodical Examination",\n    "tableOfSpecifications": [\n        { "objective": "...", "daysTaught": 5, "percentage": "10%", "numItems": 5, "itemPlacement": "1-5", "remembering": "1", "understanding": "1", "applying": "1", "analyzing": "1", "evaluating": "1", "creating": "0" }\n    ],\n    "questions": [\n        { "questionText": "...", "options": ["A", "B", "C", "D"], "correctAnswer": "A. ..." }\n    ],\n    "subject": "${options.subject}",\n    "gradeLevel": "${options.gradeLevel}",\n    "quarter": "${options.quarter}"\n}`; try { const response = await callApiProxy({ model, contents: prompt, config: { responseMimeType: "application/json" } }); return parseJsonFromAiResponse(response.text || "{}"); } catch (error) { throw handleGeminiError(error, 'generateExam'); }
};
