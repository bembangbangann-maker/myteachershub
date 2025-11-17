import { Type, HarmCategory, HarmBlockThreshold } from "@google/genai";
import { Student, Grade, Anecdote, AIAnalysisResult, ExtractedGrade, ExtractedStudentData, DlpContent, GeneratedQuiz, QuizType, DlpRubricItem, DllContent, AttendanceStatus, DlpProcedure, LearningActivitySheet, CotLessonPlan, CotProcedureStep, ExamObjective, GeneratedExam } from '../types';
// Fix: Import toast to show error messages to the user.
import { toast } from "react-hot-toast";

// --- UTILITY FUNCTIONS ---

/**
 * A robust JSON parser that handles AI responses which might be wrapped in markdown code blocks.
 * @param jsonString The raw string response from the AI model.
 * @returns The parsed JSON object.
 */
const parseJsonFromAiResponse = <T>(jsonString: string | undefined | null): T => {
    // FIX: Guard against null, undefined, or empty string responses from the AI.
    // This prevents the '.trim()' error and ensures graceful failure.
    if (!jsonString || jsonString.trim() === '') {
        throw new Error("AI returned no text content to parse.");
    }
    // The AI may wrap the JSON in ```json ... ```, so we strip it.
    const sanitizedString = jsonString.trim().replace(/^```json\s*|```\s*$/g, '');
    try {
        return JSON.parse(sanitizedString) as T;
    } catch (error) {
        console.error("Failed to parse sanitized JSON:", sanitizedString);
        throw new Error("AI returned a malformed JSON response.");
    }
};


/**
 * Calls the secure Vercel/serverless function API proxy.
 * @param modelOptions The request body to send to the Gemini API.
 * @returns The full response object from the Gemini API.
 */
const callApiProxy = async (modelOptions: any): Promise<any> => {
    const response = await fetch('/api/gemini', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(modelOptions),
    });

    if (!response.ok) {
        let errorDetails = `AI API call failed with status ${response.status}.`;
        try {
            // Vercel often sends JSON errors for function issues
            const errorData = await response.json();
            errorDetails = errorData.details || errorData.error || JSON.stringify(errorData);
        } catch (e) {
            // If the response isn't JSON (e.g., a gateway timeout), use the text body.
            const textError = await response.text();
            if (textError) errorDetails = textError;
        }
        console.error('API Proxy Error:', errorDetails);
        throw new Error(errorDetails);
    }
    
    // The proxy returns the entire GenerateContentResponse object from the SDK.
    return response.json();
};

/**
 * Standard safety settings to allow educational content generation.
 */
const safetySettings = [
  { category: HarmCategory.HARM_CATEGORY_HARASSMENT, threshold: HarmBlockThreshold.BLOCK_NONE },
  { category: HarmCategory.HARM_CATEGORY_HATE_SPEECH, threshold: HarmBlockThreshold.BLOCK_NONE },
  { category: HarmCategory.HARM_CATEGORY_SEXUALLY_EXPLICIT, threshold: HarmBlockThreshold.BLOCK_NONE },
  { category: HarmCategory.HARM_CATEGORY_DANGEROUS_CONTENT, threshold: HarmBlockThreshold.BLOCK_NONE },
];

/**
 * Standard system instructions for complex generation tasks to improve efficiency.
 */
const efficientGenerationSystemInstruction = "You are an expert educational content creator. Be efficient and generate the JSON output directly without any extra conversation or explanation.";


/**
 * Processes and formats errors from the AI service for user-facing display.
 * @param error The error object.
 * @param functionName The name of the service function where the error occurred.
 * @returns A user-friendly Error object.
 */
const handleGeminiError = (error: any, functionName: string): Error => {
    console.error(`Error in ${functionName}:`, error);
    
    let userMessage = "An AI feature failed. Please try again. If the problem persists, check your connection or the server logs.";

    if (error && typeof error.message === 'string') {
        const lowerMessage = error.message.toLowerCase();
        if (lowerMessage.includes('api key not valid') || lowerMessage.includes('api_key_invalid')) {
            userMessage = "AI Error: The API key configured on the server is invalid.";
        } else if (lowerMessage.includes('quota')) {
            userMessage = "AI Error: API quota exceeded. Please check your billing details.";
        } else if (lowerMessage.includes('timeout')) {
            userMessage = "AI Error: The request timed out. The task may be too complex. Please try simplifying it.";
        } else {
            userMessage = `AI Error: ${error.message}`;
        }
    }
    
    return new Error(userMessage);
};

// --- API FUNCTIONS ---

export const checkApiStatus = async (): Promise<{ status: 'success' | 'error'; message: string }> => {
    try {
        await callApiProxy({ model: 'gemini-2.5-flash', contents: 'test' });
        return { status: 'success', message: 'Connection successful. The secure AI proxy is working correctly.' };
    } catch (error) {
        const processedError = handleGeminiError(error, 'checkApiStatus');
        return { status: 'error', message: processedError.message };
    }
};

const performanceAnalysisSchema = { type: Type.ARRAY, items: { type: Type.OBJECT, properties: { studentName: { type: Type.STRING }, trendSummary: { type: Type.STRING }, recommendation: { type: Type.STRING } }, required: ["studentName", "trendSummary", "recommendation"] } };

export const analyzeStudentPerformance = async (students: Student[], grades: Grade[], anecdotes: Anecdote[]): Promise<AIAnalysisResult[]> => {
  const model = "gemini-2.5-pro";
  const studentData = students.map(student => ({
      name: `${student.firstName} ${student.lastName}`,
      grades: grades.filter(g => g.studentId === student.id).map(g => ({ subject: g.subject, quarter: g.quarter, percentage: ((g.score / g.maxScore) * 100).toFixed(2) })),
      anecdotes: anecdotes.filter(a => a.studentId === student.id).map(a => a.observation)
  }));

  const prompt = `As an expert teacher, analyze the provided data for a class of students. Identify students who are either excelling or at risk based on grade trends and anecdotal records. Focus on significant, consistent trends or notable outliers. For each identified student, provide their full name, a one-sentence trendSummary, and a personalized, actionable recommendation (intervention for at-risk, enrichment for excelling). Do not include students with stable or average performance.

    Here is the class data: ${JSON.stringify(studentData, null, 2)}

    Return the result as a JSON array. If no students show significant trends, return an empty array. Adhere strictly to the provided JSON schema.`;

  try {
    const response = await callApiProxy({
      model,
      contents: prompt,
      config: {
        responseMimeType: "application/json",
        responseSchema: performanceAnalysisSchema,
        safetySettings,
      },
      systemInstruction: "You are an expert teacher analyzing student data. Be concise and accurate."
    });

    const result = parseJsonFromAiResponse<AIAnalysisResult[] | AIAnalysisResult>(response.text);
    // Failsafe: If the AI returns a single object instead of an array, wrap it.
    if (!Array.isArray(result)) {
        return [result];
    }
    return result;

  } catch (error) {
    throw handleGeminiError(error, 'analyzeStudentPerformance');
  }
};

const gradeExtractionSchema = { type: Type.OBJECT, properties: { grades: { type: Type.ARRAY, items: { type: Type.OBJECT, properties: { studentName: { type: Type.STRING }, score: { type: Type.NUMBER }, maxScore: { type: Type.NUMBER } }, required: ["studentName", "score", "maxScore"] } } } };

export const extractGradesFromImage = async (base64Image: string, students: Student[]): Promise<ExtractedGrade[]> => {
    const model = "gemini-2.5-flash";
    const studentList = students.map(s => `${s.firstName} ${s.lastName}`).join(', ');
    const prompt = `Analyze the image of a grade sheet. Extract each student's name, their score, and the maximum possible score. Match names to this list: ${studentList}. Return data as JSON.`;

    try {
        const response = await callApiProxy({
            model,
            contents: { parts: [{ text: prompt }, { inlineData: { mimeType: 'image/jpeg', data: base64Image } }] },
            config: { responseMimeType: "application/json", responseSchema: gradeExtractionSchema, safetySettings },
        });

        const result = parseJsonFromAiResponse<{ grades?: ExtractedGrade[] }>(response.text);
        return result.grades || [];
    } catch (error) {
        throw handleGeminiError(error, 'extractGradesFromImage');
    }
};

const rephraseSchema = { type: Type.OBJECT, properties: { revisedText: { type: Type.STRING } }, required: ["revisedText"] };

export const rephraseAnecdote = async (text: string, mode: 'correct' | 'rephrase'): Promise<string> => {
    const model = "gemini-2.5-flash";
    const prompt = mode === 'correct' 
        ? `Correct grammar and spelling. Text: "${text}"`
        : `Rephrase for a formal, objective student record. Text: "${text}"`;

    try {
        const response = await callApiProxy({
            model, contents: prompt,
            config: { responseMimeType: "application/json", responseSchema: rephraseSchema, safetySettings },
        });
        const result = parseJsonFromAiResponse<{ revisedText: string }>(response.text);
        return result.revisedText;
    } catch (error) {
        throw handleGeminiError(error, 'rephraseAnecdote');
    }
};

const reportCardCommentSchema = { type: Type.OBJECT, properties: { strengths: { type: Type.STRING }, areasForImprovement: { type: Type.STRING }, closingStatement: { type: Type.STRING } }, required: ["strengths", "areasForImprovement", "closingStatement"] };

export const generateReportCardComment = async (student: Student, grades: Grade[], anecdotes: Anecdote[]): Promise<{strengths: string, areasForImprovement: string, closingStatement: string}> => {
    const model = "gemini-2.5-pro";
    const gradeSummary = grades.map(g => ({ subject: g.subject, type: g.type, percentage: ((g.score / g.maxScore) * 100).toFixed(0) }));
    const prompt = `As a caring teacher, write a report card comment for ${student.firstName} ${student.lastName}.
        Student Data: - Grades: ${JSON.stringify(gradeSummary)} - Anecdotes: ${JSON.stringify(anecdotes.map(a => a.observation))}
        Instructions: Write a positive paragraph on strengths, a constructive one on areas for improvement, and a brief closing statement. Format as JSON.`;
    
     try {
        const response = await callApiProxy({ model, contents: prompt, config: { responseMimeType: "application/json", responseSchema: reportCardCommentSchema, safetySettings } });
        return parseJsonFromAiResponse(response.text);
    } catch (error) {
        throw handleGeminiError(error, 'generateReportCardComment');
    }
};

const quoteSchema = { type: Type.OBJECT, properties: { quote: { type: Type.STRING }, author: { type: Type.STRING } }, required: ["quote", "author"] };

export const getInspirationalQuote = async (): Promise<{ quote: string; author: string }> => {
    const model = "gemini-2.5-flash";
    const prompt = "Generate a short, inspirational quote for a teacher about education or growth. Return as JSON.";
    try {
        const response = await callApiProxy({ model, contents: prompt, config: { responseMimeType: "application/json", responseSchema: quoteSchema, safetySettings } });
        const result = parseJsonFromAiResponse<{ quote: string; author: string }>(response.text);
        if (!result.quote || !result.author) throw new Error("AI returned an invalid quote structure.");
        return result;
    } catch (error) {
        return { quote: "The beautiful thing about learning is that no one can take it away from you.", author: "B.B. King" }; // Failsafe
    }
};

const certificateContentSchema = { type: Type.OBJECT, properties: { certificateText: { type: Type.STRING } }, required: ["certificateText"] };

export const generateCertificateContent = async (details: { awardTitle: string; tone: string; achievements?: string }): Promise<string> => {
    const model = "gemini-2.5-flash";
    const { awardTitle, tone, achievements } = details;
    const prompt = `Craft body content for a student certificate for "${awardTitle}" with a ${tone} tone. ${achievements ? `Mention: "${achievements}"` : ''} Use placeholders '^^[STUDENT_NAME]^^' and '##[AWARD_TYPE]##'. Return JSON.`;

    try {
        const response = await callApiProxy({ model, contents: prompt, config: { responseMimeType: "application/json", responseSchema: certificateContentSchema, safetySettings } });
        const result = parseJsonFromAiResponse<{ certificateText: string }>(response.text);
        return result.certificateText;
    } catch (error) {
        throw handleGeminiError(error, 'generateCertificateContent');
    }
};

export const processAttendanceCommand = async (command: string, students: Student[]): Promise<{ status: AttendanceStatus, studentIds: string[] } | null> => {
    const model = "gemini-2.5-flash";
    const updateAttendanceTool = {
        functionDeclarations: [{
            name: 'update_attendance',
            description: 'Updates the attendance status for one or more students.',
            parameters: {
                type: Type.OBJECT,
                properties: {
                    status: { type: Type.STRING, enum: ['present', 'absent', 'late'] },
                    studentNames: { type: Type.ARRAY, items: { type: Type.STRING }, description: "List of student first names, last names, or full names." },
                },
                required: ['status', 'studentNames'],
            },
        }],
    };

    const studentList = students.map(s => `${s.firstName} ${s.lastName}`).join(', ');
    const prompt = `Command: "${command}". Update attendance for students from this list: ${studentList}.`;

    try {
        const response = await callApiProxy({ model, contents: prompt, config: { tools: [updateAttendanceTool], safetySettings } });
        const fc = response.functionCalls?.[0];

        if (fc && fc.name === 'update_attendance' && fc.args.studentNames) {
            const status = fc.args.status as AttendanceStatus;
            const namesToFind = fc.args.studentNames.map((name: string) => name.toLowerCase());
            const studentIds = students
                .filter(s => namesToFind.some((name: string) => 
                    s.firstName.toLowerCase().includes(name) ||
                    s.lastName.toLowerCase().includes(name) ||
                    `${s.firstName} ${s.lastName}`.toLowerCase().includes(name)
                ))
                .map(s => s.id);
            return { status, studentIds };
        }
        return null;
    } catch (error) {
        throw handleGeminiError(error, 'processAttendanceCommand');
    }
};

// --- START OF MISSING FUNCTIONS ---

// --- DLP GENERATOR ---
const dlpProcedureSchema = { type: Type.OBJECT, properties: { title: { type: Type.STRING }, content: { type: Type.STRING, description: "Detailed teacher and student activities using markdown for formatting. MUST explicitly label main activities as '(LOTS)' or '(HOTS)'." }, ppst: { type: Type.STRING, description: "Relevant PPST indicator for the procedure." } }, required: ["title", "content", "ppst"] };
const dlpEvaluationQuestionSchema = { type: Type.OBJECT, properties: { question: { type: Type.STRING }, options: { type: Type.ARRAY, description: "An array of exactly 4 string options.", items: { type: Type.STRING } }, answer: { type: Type.STRING } }, required: ["question", "options", "answer"] };
const dlpContentSchema = { type: Type.OBJECT, properties: { contentStandard: { type: Type.STRING }, performanceStandard: { type: Type.STRING }, topic: { type: Type.STRING }, learningReferences: { type: Type.STRING }, learningMaterials: { type: Type.STRING }, procedures: { type: Type.ARRAY, items: dlpProcedureSchema }, evaluationQuestions: { type: Type.ARRAY, description: "An array of exactly 5 multiple-choice questions.", items: dlpEvaluationQuestionSchema }, remarksContent: { type: Type.STRING } }, required: ["contentStandard", "performanceStandard", "topic", "learningReferences", "learningMaterials", "procedures", "evaluationQuestions", "remarksContent"] };

export const generateDlpContent = async (details: {
    gradeLevel: string,
    learningCompetency: string,
    lessonObjective: string,
    previousLesson: string,
    selectedQuarter: string,
    subject: string,
    teacherPosition: string,
    language: 'English' | 'Filipino',
    dlpFormat: string,
}): Promise<DlpContent> => {
    const model = "gemini-2.5-pro";
    const prompt = `
        You are an expert instructional designer for the Philippine Department of Education. Your task is to generate a complete Daily Lesson Plan (DLP) strictly following DepEd Order No. 42, s. 2016.

        **User Inputs:**
        - Grade Level: ${details.gradeLevel}
        - Subject: ${details.subject}
        - Learning Competency: ${details.learningCompetency}
        - Lesson Objective: ${details.lessonObjective}
        - Previous Lesson Topic: ${details.previousLesson}
        - Language: ${details.language}
        - DLP Format: ${details.dlpFormat}
        - Teacher Position: ${details.teacherPosition} Teacher

        **Strict Generation Requirements:**

        1.  **Alignment:** ALL generated content (standards, topic, activities, evaluation) MUST be directly and strongly anchored to the provided **Learning Competency** and **Lesson Objective**.
        2.  **Standards & Topic:** Generate relevant Content and Performance Standards and a specific Topic based on the competency.
        3.  **Procedures/Activities:**
            *   Structure the procedures according to the specified format (${details.dlpFormat}).
            *   For each procedure step, provide detailed teacher and student activities in the \`content\` field. Use Markdown for formatting (e.g., bolding, lists).
            *   **Crucially, you MUST explicitly label the main cognitive activities as either '(LOTS)' for Lower-Order Thinking Skills or '(HOTS)' for Higher-Order Thinking Skills.** Ensure a logical progression from LOTS to HOTS.
            *   All activities must be designed to help learners achieve the stated **Lesson Objective**.
        4.  **Evaluation:**
            *   Create **exactly 5 multiple-choice questions**.
            *   Each question must have **exactly four (4) options**.
            *   Provide the correct letter or full text answer for each question.
            *   The evaluation MUST directly and accurately assess the achievement of the **Lesson Objective**.
        5.  **PPST Indicators:** For each procedure, assign relevant PPST indicators appropriate for a **${details.teacherPosition} Teacher**.
        6.  **Output Format:** Generate the output *directly* as a single JSON object. Do not include any extra text, conversation, or markdown formatting like \`\`\`json around the final JSON output.
    `;

    try {
        const response = await callApiProxy({
            model,
            contents: prompt,
            config: {
                responseMimeType: "application/json",
                responseSchema: dlpContentSchema,
                safetySettings
            },
            systemInstruction: "You are an expert DepEd instructional designer generating a structured DLP in JSON format."
        });
        return parseJsonFromAiResponse<DlpContent>(response.text);
    } catch (error) {
        throw handleGeminiError(error, 'generateDlpContent');
    }
};

// --- QUIZ GENERATOR ---
const rubricItemSchema = { type: Type.OBJECT, properties: { criteria: { type: Type.STRING }, points: { type: Type.NUMBER } }, required: ["criteria", "points"] };
const quizQuestionSchema = { type: Type.OBJECT, properties: { questionText: { type: Type.STRING }, options: { type: Type.ARRAY, items: { type: Type.STRING } }, correctAnswer: { type: Type.STRING } }, required: ["questionText", "correctAnswer"] };
const quizSectionSchema = { type: Type.OBJECT, properties: { instructions: { type: Type.STRING }, questions: { type: Type.ARRAY, items: quizQuestionSchema } }, required: ["instructions", "questions"] };
const quizContentSchema = {
    type: Type.OBJECT,
    properties: {
        quizTitle: { type: Type.STRING },
        tableOfSpecifications: { type: Type.ARRAY, items: { type: Type.OBJECT, properties: { objective: { type: Type.STRING }, cognitiveLevel: { type: Type.STRING }, itemNumbers: { type: Type.STRING } }, required: ["objective", "cognitiveLevel", "itemNumbers"] } },
        questionsByType: {
            type: Type.OBJECT,
            properties: {
                'Multiple Choice': quizSectionSchema,
                'True or False': quizSectionSchema,
                'Identification': quizSectionSchema,
            },
        },
        activities: {
            type: Type.ARRAY,
            items: {
                type: Type.OBJECT,
                properties: {
                    activityName: { type: Type.STRING },
                    activityInstructions: { type: Type.STRING },
                },
                required: ["activityName", "activityInstructions"],
            },
        },
    },
    required: ["quizTitle", "questionsByType", "activities"],
};


export const generateQuizContent = async (details: { topic: string; numQuestions: number; quizTypes: QuizType[]; subject: string; gradeLevel: string; }): Promise<GeneratedQuiz> => {
    const model = "gemini-2.5-pro";
    const prompt = `Generate a comprehensive quiz for a Grade ${details.gradeLevel} ${details.subject} class on the topic: "${details.topic}".
    
    The quiz must include:
    1. A suitable title for the quiz.
    2. A simple Table of Specifications (TOS) linking objectives to item numbers.
    3. The following quiz types: ${details.quizTypes.join(', ')}.
    4. Exactly ${details.numQuestions} questions for EACH specified quiz type.
    5. Clear instructions for each section.
    6. For Multiple Choice questions, provide 4 options.
    7. For all question types, provide the correct answer.
    8. Two creative, performance-based activities related to the topic. Do not generate rubrics for them yet.

    Return the result as a single JSON object.`;
    try {
        const response = await callApiProxy({
            model,
            contents: prompt,
            config: {
                responseMimeType: "application/json",
                responseSchema: quizContentSchema,
                safetySettings,
            },
            systemInstruction: efficientGenerationSystemInstruction,
        });
        return parseJsonFromAiResponse<GeneratedQuiz>(response.text);
    } catch (error) {
        throw handleGeminiError(error, 'generateQuizContent');
    }
};

export const generateRubricForActivity = async (details: { activityName: string, activityInstructions: string, totalPoints: number }): Promise<DlpRubricItem[]> => {
    const model = "gemini-2.5-flash";
    const prompt = `Create a simple scoring rubric for the following activity. The total score must add up to exactly ${details.totalPoints} points.
    
    Activity Name: ${details.activityName}
    Instructions: ${details.activityInstructions}
    
    Return as a JSON array where each item has "criteria" and "points".`;

    try {
        const response = await callApiProxy({
            model,
            contents: prompt,
            config: {
                responseMimeType: "application/json",
                responseSchema: { type: Type.ARRAY, items: rubricItemSchema },
                safetySettings,
            },
            systemInstruction: efficientGenerationSystemInstruction,
        });
        return parseJsonFromAiResponse<DlpRubricItem[]>(response.text);
    } catch (error) {
        throw handleGeminiError(error, 'generateRubricForActivity');
    }
};

// --- DLL (WEEKLY PLAN) GENERATOR ---
const dllDailyEntrySchema = { type: Type.OBJECT, properties: { monday: { type: Type.STRING }, tuesday: { type: Type.STRING }, wednesday: { type: Type.STRING }, thursday: { type: Type.STRING }, friday: { type: Type.STRING } }, required: ["monday", "tuesday", "wednesday", "thursday", "friday"] };
const dllProcedureSchema = { type: Type.OBJECT, properties: { procedure: { type: Type.STRING }, ...dllDailyEntrySchema.properties }, required: ["procedure", ...Object.keys(dllDailyEntrySchema.properties)] };
const dllContentSchema = { type: Type.OBJECT, properties: { contentStandard: { type: Type.STRING }, performanceStandard: { type: Type.STRING }, learningCompetencies: dllDailyEntrySchema, content: { type: Type.STRING }, learningResources: { type: Type.OBJECT, properties: { teacherGuidePages: dllDailyEntrySchema, learnerMaterialsPages: dllDailyEntrySchema, textbookPages: dllDailyEntrySchema, additionalMaterials: dllDailyEntrySchema, otherResources: dllDailyEntrySchema }, required: ["teacherGuidePages", "learnerMaterialsPages", "textbookPages", "additionalMaterials", "otherResources"] }, procedures: { type: Type.ARRAY, items: dllProcedureSchema }, remarks: { type: Type.STRING }, reflection: { type: Type.ARRAY, items: dllProcedureSchema } }, required: ["contentStandard", "performanceStandard", "learningCompetencies", "content", "learningResources", "procedures", "remarks", "reflection"] };

export const generateDllContent = async (details: { subject: string; gradeLevel: string; weeklyTopic: string; contentStandard: string; performanceStandard: string; dllFormat: string; language: 'English' | 'Filipino' }): Promise<DllContent> => {
    const model = "gemini-2.5-pro";
    const prompt = `Generate a complete Daily Lesson Log (DLL) for a Grade ${details.gradeLevel} ${details.subject} class for one week.
    - Topic for the week: ${details.weeklyTopic || `(Suggest a relevant topic for this grade level and subject)`}
    - Content Standard: ${details.contentStandard || `(Generate an appropriate standard)`}
    - Performance Standard: ${details.performanceStandard || `(Generate an appropriate standard)`}
    - Language: ${details.language}
    - DLL Format: ${details.dllFormat}

    Instructions:
    1.  Create daily learning competencies/objectives for Monday to Friday.
    2.  Fill in all sections (Content, Learning Resources, Procedures, Remarks, Reflection) with detailed, relevant, and coherent content for each day of the week.
    3.  The procedures must be well-structured and developmentally appropriate.
    4.  Return the output as a single JSON object.`;

    try {
        const response = await callApiProxy({
            model,
            contents: prompt,
            config: {
                responseMimeType: "application/json",
                responseSchema: dllContentSchema,
                safetySettings,
            },
            systemInstruction: "You are an expert DepEd teacher creating a detailed weekly lesson log in JSON format."
        });
        return parseJsonFromAiResponse<DllContent>(response.text);
    } catch (error) {
        throw handleGeminiError(error, 'generateDllContent');
    }
};

// --- LEARNING ACTIVITY SHEET (LAS) GENERATOR ---
const lasQuestionSchema = { type: Type.OBJECT, properties: { questionText: { type: Type.STRING }, type: { type: Type.STRING, enum: ['Identification', 'Essay', 'Problem-solving', 'Multiple Choice'] }, options: { type: Type.ARRAY, items: { type: Type.STRING } }, answer: { type: Type.STRING } }, required: ["questionText", "type"] };
const lasActivitySchema = { type: Type.OBJECT, properties: { title: { type: Type.STRING }, instructions: { type: Type.STRING }, questions: { type: Type.ARRAY, items: lasQuestionSchema }, rubric: { type: Type.ARRAY, items: rubricItemSchema } }, required: ["title", "instructions"] };
const lasContentSchema = { type: Type.OBJECT, properties: { activityTitle: { type: Type.STRING }, learningTarget: { type: Type.STRING }, references: { type: Type.STRING }, conceptNotes: { type: Type.ARRAY, items: { type: Type.OBJECT, properties: { title: { type: Type.STRING }, content: { type: Type.STRING } }, required: ["title", "content"] } }, activities: { type: Type.ARRAY, items: lasActivitySchema } }, required: ["activityTitle", "learningTarget", "references", "conceptNotes", "activities"] };

export const generateLearningActivitySheet = async (details: { subject: string, gradeLevel: string, learningCompetency: string, lessonObjective: string, activityType: string, language: 'English' | 'Filipino' }): Promise<LearningActivitySheet> => {
    const model = "gemini-2.5-pro";
    const prompt = `Create a comprehensive, DLP-style Learning Activity Sheet (LAS) in ${details.language}.
    
    Details:
    - Subject: Grade ${details.gradeLevel} ${details.subject}
    - Learning Competency: ${details.learningCompetency}
    - Learning Objective: ${details.lessonObjective}
    - Activity Focus: ${details.activityType}

    Instructions:
    1.  Create a main 'Activity Title' for the LAS.
    2.  Formulate a clear 'Learning Target' based on the objective.
    3.  Provide plausible 'References'.
    4.  Write comprehensive 'Concept Notes' with clear explanations and examples about the topic.
    5.  Design at least two distinct 'Activities' that align with the activity focus. Include questions (with answers for checkable types) and a scoring rubric for performance-based tasks.
    6.  Return as a single JSON object.`;
    try {
        const response = await callApiProxy({
            model,
            contents: prompt,
            config: {
                responseMimeType: "application/json",
                responseSchema: lasContentSchema,
                safetySettings,
            },
            systemInstruction: efficientGenerationSystemInstruction,
        });
        return parseJsonFromAiResponse<LearningActivitySheet>(response.text);
    } catch (error) {
        throw handleGeminiError(error, 'generateLearningActivitySheet');
    }
};

// --- EXAM GENERATOR ---
// This will be added in a future update as per the user's phased request.
// The types have been added to types.ts and the UI is in LessonPlanners.tsx.
// We just need to implement the service function here.
// For now, returning a placeholder to avoid breaking the app.
export const generateExamContent = async (objectives: ExamObjective[], details: { subject: string, gradeLevel: string }): Promise<GeneratedExam> => {
    // Placeholder - to be implemented fully later.
    console.warn("generateExamContent is not yet fully implemented.");
    return Promise.resolve({
        title: "Generated Exam (Placeholder)",
        tableOfSpecifications: [],
        questions: [],
        subject: details.subject,
        gradeLevel: details.gradeLevel,
    });
};

// --- END OF MISSING FUNCTIONS ---
