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

const dlpRubricItemSchema = {
    type: Type.OBJECT,
    properties: {
        criteria: { type: Type.STRING },
        points: { type: Type.NUMBER },
    },
    required: ["criteria", "points"]
};

export const getInspirationalQuote = async (): Promise<{ quote: string; author: string }> => {
    const model = "gemini-2.5-flash";
    const prompt = "Give me an inspirational quote for a teacher. Return JSON: { \"quote\": \"...\", \"author\": \"...\" }";
    
    try {
        const response = await callApiProxy({
             model, 
             contents: prompt,
             config: { responseMimeType: "application/json" } 
        });
        return parseJsonFromAiResponse<{ quote: string; author: string }>(response.text || "{}");
    } catch (error) {
        throw handleGeminiError(error, 'getInspirationalQuote');
    }
};

export const processAttendanceCommand = async (command: string, students: Student[]): Promise<{ studentIds: string[], status: AttendanceStatus } | null> => {
    const model = "gemini-2.5-flash";
    const studentList = students.map(s => `${s.id}: ${s.firstName} ${s.lastName}`).join('\n');
    const prompt = `
    Students:
    ${studentList}

    Command: "${command}"

    Identify which students are mentioned or implied in the command and the attendance status (present, absent, late).
    If "everyone" or similar is mentioned, include all IDs.
    Return JSON: { "studentIds": ["id1", "id2"], "status": "present" | "absent" | "late" }
    `;

    try {
        const response = await callApiProxy({
             model, 
             contents: prompt,
             config: { responseMimeType: "application/json" }
        });
        return parseJsonFromAiResponse(response.text || "null");
    } catch (error) {
        throw handleGeminiError(error, 'processAttendanceCommand');
    }
};

export const analyzeStudentPerformance = async (students: Student[], grades: Grade[], anecdotes: Anecdote[]): Promise<AIAnalysisResult[]> => {
    const model = "gemini-2.5-flash";
    // Simplify data for prompt context window
    const dataSummary = students.map(s => {
        const sGrades = grades.filter(g => g.studentId === s.id).map(g => `${g.subject} (${g.type}): ${g.score}/${g.maxScore}`);
        const sAnecdotes = anecdotes.filter(a => a.studentId === s.id).map(a => a.observation);
        return { name: `${s.firstName} ${s.lastName}`, grades: sGrades, anecdotes: sAnecdotes };
    });

    const prompt = `
    Analyze the performance of the following students based on their grades and anecdotal records.
    Identify meaningful trends (improving, declining, struggling in specific areas, excelling).
    Provide a short summary and a specific recommendation for each student who has notable data.
    
    Data: ${JSON.stringify(dataSummary)}

    Return JSON array:
    [
      { "studentName": "Name", "trendSummary": "...", "recommendation": "..." }
    ]
    `;

    try {
        const response = await callApiProxy({
             model, 
             contents: prompt,
             config: { responseMimeType: "application/json" }
        });
        return parseJsonFromAiResponse(response.text || "[]");
    } catch (error) {
        throw handleGeminiError(error, 'analyzeStudentPerformance');
    }
};

export const extractGradesFromImage = async (base64Image: string, students: Student[]): Promise<ExtractedGrade[]> => {
    const model = "gemini-2.5-flash";
    const studentNames = students.map(s => `${s.firstName} ${s.lastName}`).join(', ');
    
    const prompt = `
    Extract student names and their scores from this image of a grade sheet.
    The known students in this class are: ${studentNames}.
    Try to match extracted names to the known list.
    Return JSON array: [{ "studentName": "matched name", "score": number, "maxScore": number }]
    If max score is not visible, guess based on the highest possible score or typical values (e.g. 10, 20, 50, 100).
    `;

    try {
        const response = await callApiProxy({
            model,
            contents: {
                parts: [
                    { inlineData: { mimeType: "image/jpeg", data: base64Image } },
                    { text: prompt }
                ]
            },
            config: { responseMimeType: "application/json" }
        });
        return parseJsonFromAiResponse(response.text || "[]");
    } catch (error) {
        throw handleGeminiError(error, 'extractGradesFromImage');
    }
};

export const rephraseAnecdote = async (text: string, mode: 'correct' | 'rephrase'): Promise<string> => {
    const model = "gemini-2.5-flash";
    const prompt = mode === 'correct' 
        ? `Correct the grammar and spelling of this teacher's observation: "${text}"`
        : `Rephrase this teacher's observation to be more professional, objective, and constructive: "${text}"`;

    try {
        const response = await callApiProxy({ model, contents: prompt });
        return response.text || text;
    } catch (error) {
        throw handleGeminiError(error, 'rephraseAnecdote');
    }
};

export const generateReportCardComment = async (student: Student, grades: Grade[], anecdotes: Anecdote[]): Promise<{ strengths: string, areasForImprovement: string, closingStatement: string }> => {
    const model = "gemini-2.5-flash";
    const prompt = `
    Generate a report card comment for ${student.firstName} ${student.lastName}.
    Grades: ${JSON.stringify(grades)}
    Anecdotes: ${JSON.stringify(anecdotes)}
    
    Return JSON: { "strengths": "...", "areasForImprovement": "...", "closingStatement": "..." }
    `;

    try {
        const response = await callApiProxy({
             model, 
             contents: prompt,
             config: { responseMimeType: "application/json" }
        });
        return parseJsonFromAiResponse(response.text || "{}");
    } catch (error) {
        throw handleGeminiError(error, 'generateReportCardComment');
    }
};

export const generateCertificateContent = async (formData: { awardTitle: string; tone: string; achievements: string }): Promise<string> => {
    const model = "gemini-2.5-flash";
    const prompt = `
    Write the content body for a certificate.
    Award: ${formData.awardTitle}
    Tone: ${formData.tone}
    Achievements to mention: ${formData.achievements}
    
    Use placeholders like [STUDENT_NAME], [DATE], etc.
    Format with markdown for bolding key phrases.
    Just return the body text.
    `;

    try {
        const response = await callApiProxy({ model, contents: prompt });
        return response.text || "";
    } catch (error) {
        throw handleGeminiError(error, 'generateCertificateContent');
    }
};

export const generateDlpContent = async (options: any): Promise<DlpContent> => {
    const model = "gemini-2.5-flash"; 
    const prompt = `
    Generate a detailed Daily Lesson Plan (DLP) for:
    Subject: ${options.subject}
    Grade: ${options.gradeLevel}
    Competency: ${options.learningCompetency}
    Objective: ${options.lessonObjective}
    Format: ${options.dlpFormat}
    Language: ${options.language}

    Return JSON matching this schema:
    {
        "contentStandard": "string",
        "performanceStandard": "string",
        "topic": "string",
        "learningReferences": "string",
        "learningMaterials": "string",
        "procedures": [ { "title": "string", "content": "string", "ppst": "string" } ],
        "evaluationQuestions": [ { "question": "string", "options": ["string"], "answer": "string" } ],
        "remarksContent": "string"
    }
    `;

    try {
        const response = await callApiProxy({
             model, 
             contents: prompt,
             config: { responseMimeType: "application/json" }
        });
        return parseJsonFromAiResponse(response.text || "{}");
    } catch (error) {
        throw handleGeminiError(error, 'generateDlpContent');
    }
};

export const generateQuizContent = async (options: any): Promise<GeneratedQuiz> => {
    const model = "gemini-2.5-flash";
    const prompt = `
    Generate a quiz.
    Topic: ${options.topic}
    Subject: ${options.subject}
    Grade: ${options.gradeLevel}
    Questions per type: ${options.numQuestions}
    Types: ${options.quizTypes.join(', ')}

    Return JSON:
    {
        "quizTitle": "string",
        "questionsByType": {
            "Multiple Choice": { "instructions": "...", "questions": [{ "questionText": "...", "options": ["..."], "correctAnswer": "..." }] },
            // ... other types
        },
        "activities": [
            { "activityName": "string", "activityInstructions": "string" } 
        ],
        "tableOfSpecifications": [
             { "objective": "string", "cognitiveLevel": "string", "itemNumbers": "string" }
        ]
    }
    `;

    try {
        const response = await callApiProxy({
             model, 
             contents: prompt,
             config: { responseMimeType: "application/json" }
        });
        return parseJsonFromAiResponse(response.text || "{}");
    } catch (error) {
        throw handleGeminiError(error, 'generateQuizContent');
    }
};

export const generateRubricForActivity = async (options: { activityName: string, activityInstructions: string, totalPoints: number }): Promise<DlpRubricItem[]> => {
    const model = "gemini-2.5-flash";
    const prompt = `
    Create a rubric for activity: "${options.activityName}"
    Instructions: "${options.activityInstructions}"
    Total Points: ${options.totalPoints}

    Return JSON array: [{ "criteria": "string", "points": number }]
    `;

    try {
        const response = await callApiProxy({
             model, 
             contents: prompt,
             config: { responseMimeType: "application/json" }
        });
        return parseJsonFromAiResponse(response.text || "[]");
    } catch (error) {
        throw handleGeminiError(error, 'generateRubricForActivity');
    }
};

export const generateDllContent = async (options: any): Promise<DllContent> => {
    const model = "gemini-2.5-flash";
    const prompt = `
    Generate a Daily Lesson Log (DLL) for one week.
    Subject: ${options.subject}
    Grade: ${options.gradeLevel}
    Topic: ${options.weeklyTopic}
    Language: ${options.language}

    Return JSON structure matching DllContent interface.
    `;
    const schemaHint = `
    {
        "contentStandard": "", "performanceStandard": "",
        "learningCompetencies": { "monday": "", "tuesday": "", "wednesday": "", "thursday": "", "friday": "" },
        "content": "",
        "learningResources": { ... },
        "procedures": [ { "procedure": "Review...", "monday": "", "tuesday": "", ... } ],
        "remarks": "", "reflection": []
    }
    `;

    try {
        const response = await callApiProxy({
             model, 
             contents: prompt + "\nJSON Schema:" + schemaHint,
             config: { responseMimeType: "application/json" }
        });
        return parseJsonFromAiResponse(response.text || "{}");
    } catch (error) {
        throw handleGeminiError(error, 'generateDllContent');
    }
};

export const generateLearningActivitySheet = async (options: { subject: string; gradeLevel: string; learningCompetency: string; lessonObjective: string; activityType: string; language: 'English' | 'Filipino' }): Promise<LearningActivitySheet> => {
    const { subject, gradeLevel, learningCompetency, lessonObjective, activityType, language } = options;
    const model = "gemini-3-pro-preview";
    const prompt = `
        You are an expert curriculum designer for the Department of Education in the Philippines. Create a comprehensive, 5-day learning packet following the DLP-style Learning Activity Sheet (LAS) format. The entire week must focus on scaffolding the single provided learning competency, with each day building upon the previous one.

        **Main Learning Competency:** ${learningCompetency}
        **Main Lesson Objective:** ${lessonObjective}
        **Activity Focus for the week:** ${activityType}
        **Language:** ${language}

        **CRITICAL INSTRUCTIONS FOR EACH OF THE 5 DAYS:**
        1.  **Structure:** Each day's content MUST be self-contained and follow this exact structure: **CONCEPT NOTES**, **ACTIVITY**, and **REFLECTION**.
        2.  **dayTitle:** Create a clear, concise title for the day's lesson (e.g., "Day 1: Understanding Bias and Prejudice").
        3.  **activityTitle:** Create a specific title for the activity sheet itself.
        4.  **learningTarget:** Write a specific, measurable sub-objective for the day that scaffolds towards the main weekly objective.
        5.  **references:** Suggest relevant learning materials or sources.
        6.  **conceptNotes Section:** This MUST be an array containing objects with a title and content.
            - The object's \`title\` property must be exactly "CONCEPT NOTES".
            - The object's \`content\` property should be a clear, detailed explanation of the day's core concept, including definitions and examples. Use markdown for emphasis: **bold text** for key terms, *italic text* for examples, and bullet points starting with '• ' for lists.
        7.  **activities Section:** This MUST be an array containing activity objects.
            - The object's \`title\` property must start with "ACTIVITY:", followed by a descriptive name (e.g., "ACTIVITY: DEFINE AND MATCH").
            - The object's \`instructions\` property must contain clear directions for the student, followed by all activity content.
            - **IMPORTANT:** All activity content, including questions, scenarios, or tables, must be included within the \`instructions\` string. For matching type activities or simple two-column tables, format each row on a new line using " || " as a separator between the columns. For example: \`1. Judging a person before knowing them. || A. Bias\`. 
        8.  **Day 5 Task:** Ensure the activity for Day 5 is a culminating Performance Task.
        9.  **reflection Section:** This must be a single, thought-provoking question or sentence-completion task related to the day's lesson.

        **Output Format:**
        Strictly return a single JSON object that adheres to the provided schema. The root object must have a 'days' key, which is an array of 5 day objects, each following the structure detailed above. Do not add any extra text, conversation, or explanations outside the JSON structure.
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
    const model = "gemini-3-pro-preview"; 
    const prompt = `
    Create a 50-item Periodical Exam with Table of Specifications (TOS).
    Subject: ${options.subject}
    Grade: ${options.gradeLevel}
    Quarter: ${options.quarter}
    Objectives and Days Taught: ${JSON.stringify(options.objectives)}

    Return JSON:
    {
        "title": "First Periodical Examination",
        "tableOfSpecifications": [
            { "objective": "...", "daysTaught": 5, "percentage": "10%", "numItems": 5, "itemPlacement": "1-5", "remembering": "1", "understanding": "1", "applying": "1", "analyzing": "1", "evaluating": "1", "creating": "0" }
        ],
        "questions": [
            { "questionText": "...", "options": ["A", "B", "C", "D"], "correctAnswer": "A. ..." }
        ],
        "subject": "${options.subject}",
        "gradeLevel": "${options.gradeLevel}",
        "quarter": "${options.quarter}"
    }
    `;

    try {
        const response = await callApiProxy({
             model, 
             contents: prompt,
             config: { responseMimeType: "application/json" }
        });
        return parseJsonFromAiResponse(response.text || "{}");
    } catch (error) {
        throw handleGeminiError(error, 'generateExam');
    }
};
