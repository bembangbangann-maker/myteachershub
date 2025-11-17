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

const quoteSchema = { type: Type.OBJECT, properties: { quote: { type: Type.STRING }, author: { type: Type.STRING } }, required: ["quote", "author"] };
export const getInspirationalQuote = async (): Promise<{ quote: string; author: string }> => {
    const model = "gemini-2.5-flash";
    const prompt = "Provide a short, inspirational quote suitable for a teacher. Include the author. Respond in JSON format.";
    try {
        const response = await callApiProxy({
            model, contents: prompt,
            config: { responseMimeType: "application/json", responseSchema: quoteSchema, safetySettings }
        });
        return parseJsonFromAiResponse<{ quote: string; author: string }>(response.text);
    } catch (error) {
        throw handleGeminiError(error, 'getInspirationalQuote');
    }
};

const attendanceCommandSchema = { type: Type.OBJECT, properties: { status: { type: Type.STRING, enum: ["present", "absent", "late"] }, studentIds: { type: Type.ARRAY, items: { type: Type.STRING } } }, required: ["status", "studentIds"] };
export const processAttendanceCommand = async (command: string, students: Student[]): Promise<{ status: AttendanceStatus; studentIds: string[] }> => {
    const model = "gemini-2.5-flash";
    const studentList = students.map(s => ({ id: s.id, name: `${s.firstName} ${s.lastName}` }));
    const prompt = `
        Analyze the command: "${command}"
        Determine the attendance status (present, absent, or late).
        Identify the students mentioned. Match them against this list: ${JSON.stringify(studentList)}.
        Return a JSON object with the status and a list of the corresponding student IDs.
        If the command says "everyone" or "all", include all student IDs.
    `;
    try {
        const response = await callApiProxy({
            model, contents: prompt,
            config: { responseMimeType: "application/json", responseSchema: attendanceCommandSchema, safetySettings }
        });
        return parseJsonFromAiResponse<{ status: AttendanceStatus; studentIds: string[] }>(response.text);
    } catch (error) {
        throw handleGeminiError(error, 'processAttendanceCommand');
    }
};

const reportCardCommentSchema = { type: Type.OBJECT, properties: { strengths: { type: Type.STRING }, areasForImprovement: { type: Type.STRING }, closingStatement: { type: Type.STRING } }, required: ["strengths", "areasForImprovement", "closingStatement"] };
export const generateReportCardComment = async (student: Student, grades: Grade[], anecdotes: Anecdote[]): Promise<{ strengths: string; areasForImprovement: string; closingStatement: string; }> => {
    const model = "gemini-2.5-pro";
    const prompt = `
        Generate a report card comment for ${student.firstName} ${student.lastName}.
        Data:
        Grades: ${JSON.stringify(grades)}
        Anecdotes: ${JSON.stringify(anecdotes.map(a => a.observation))}
        
        Structure the comment into three parts:
        1. "strengths": Positive observations.
        2. "areasForImprovement": Constructive feedback.
        3. "closingStatement": An encouraging final remark.

        Be professional, constructive, and personalized. Respond in JSON format.
    `;
    try {
        const response = await callApiProxy({
            model, contents: prompt,
            config: { responseMimeType: "application/json", responseSchema: reportCardCommentSchema, safetySettings }
        });
        return parseJsonFromAiResponse<{ strengths: string; areasForImprovement: string; closingStatement: string; }>(response.text);
    } catch (error) {
        throw handleGeminiError(error, 'generateReportCardComment');
    }
};

const certificateContentSchema = { type: Type.OBJECT, properties: { newContent: { type: Type.STRING } }, required: ["newContent"] };
export const generateCertificateContent = async (options: { awardTitle: string; tone: string; achievements: string; }): Promise<string> => {
    const model = "gemini-2.5-flash";
    const prompt = `
        Create a concise and powerful body text for an academic award certificate.
        - Award Title: "${options.awardTitle}"
        - Tone: ${options.tone}
        - Key Achievements to mention (optional): "${options.achievements}"

        Use these placeholders:
        - ^^[STUDENT_NAME]^^ for the student's name (make it very large).
        - ##[AWARD_TYPE]## for the specific award (make it large).
        - [GRADE_AND_SECTION], [GENERAL_AVERAGE], [SCHOOL_NAME], [DAY], [MONTH], [YEAR].
        
        Use markdown like **bold** and *italics*.
        Return a JSON object with a single key "newContent" containing the generated text.
    `;
    try {
        const response = await callApiProxy({
            model, contents: prompt,
            config: { responseMimeType: "application/json", responseSchema: certificateContentSchema, safetySettings }
        });
        const result = parseJsonFromAiResponse<{ newContent: string }>(response.text);
        return result.newContent;
    } catch (error) {
        throw handleGeminiError(error, 'generateCertificateContent');
    }
};

const dlpProcedureSchema = {
    type: Type.OBJECT,
    properties: {
        title: { type: Type.STRING, description: "A creative, descriptive title for the activity (e.g., 'Activity #1: Word Hunt'). Do not use generic titles like 'Teacher's Activity'." },
        content: { type: Type.STRING, description: "Detailed teacher and student activities. Use markdown for formatting: **bold** for emphasis, *italics* for special notes. Do not use asterisks (*), use bold or italics instead." },
        ppst: { type: Type.STRING, description: "Relevant PPST Indicator code, e.g., 1.1.2" }
    },
    required: ["title", "content", "ppst"]
};

const dlpEvaluationQuestionSchema = {
    type: Type.OBJECT,
    properties: {
        question: { type: Type.STRING },
        options: { type: Type.ARRAY, items: { type: Type.STRING } },
        answer: { type: Type.STRING }
    },
    required: ["question", "options", "answer"]
};

const dlpContentSchema = {
    type: Type.OBJECT,
    properties: {
        contentStandard: { type: Type.STRING },
        performanceStandard: { type: Type.STRING },
        topic: { type: Type.STRING },
        learningReferences: { type: Type.STRING, description: "List of references used." },
        learningMaterials: { type: Type.STRING, description: "List of materials needed." },
        procedures: { type: Type.ARRAY, items: dlpProcedureSchema },
        evaluationQuestions: { type: Type.ARRAY, items: dlpEvaluationQuestionSchema },
        remarksContent: { type: Type.STRING, description: "An empty string. The teacher fills this out manually." }
    },
    required: ["contentStandard", "performanceStandard", "topic", "learningReferences", "learningMaterials", "procedures", "evaluationQuestions", "remarksContent"]
};

export const generateDlpContent = async (options: {
    gradeLevel: string;
    subject: string;
    learningCompetency: string;
    lessonObjective: string;
    previousLesson: string;
    selectedQuarter: string;
    teacherPosition: string;
    language: 'English' | 'Filipino';
    dlpFormat: string;
}): Promise<DlpContent> => {
    const { gradeLevel, subject, learningCompetency, lessonObjective, previousLesson, selectedQuarter, teacherPosition, language, dlpFormat } = options;
    const model = "gemini-2.5-pro";

    const prompt = `
        You are an expert Filipino educator creating a Daily Lesson Plan (DLP).
        Your task is to generate a complete DLP in ${language} for a Grade ${gradeLevel} ${subject} class.
        
        Key Information:
        - Quarter: ${selectedQuarter}
        - Learning Competency: ${learningCompetency}
        - Specific Lesson Objective: ${lessonObjective}
        - Previous Lesson (for review context): ${previousLesson}
        - Teacher's Position (for PPST alignment): ${teacherPosition}
        - DLP Format: ${dlpFormat} (Use this format for the procedures. E.g., for 4As: Activity, Analysis, Abstraction, Application)

        Instructions:
        1.  Generate all sections of the DLP: Content Standard, Performance Standard, Topic, Learning References, Learning Materials, Procedures, and Evaluation Questions.
        2.  For the "Procedures" section:
            - Create a sequence of activities based on the selected DLP Format (${dlpFormat}).
            - For each procedure step (e.g., Motivation, Activity, Analysis), provide a creative and descriptive **title**. Do not use generic labels like "Teacher's Activity" or "Student's Activity". For example, use titles like "Activity #1: The Word Maze", "Group Discussion: Unpacking the Poem", or "Motivation: Picture Analysis".
            - The 'content' for each procedure should detail both the teacher's actions/instructions and the expected student activities. Use markdown for formatting: use **bold letters** for emphasis and *italics* for special notes. Do not use asterisks for lists or any other purpose.
            - Align each procedure with a relevant PPST indicator based on the teacher's position (${teacherPosition}).
        3.  Create 5 multiple-choice evaluation questions with 4 options each, and provide the correct answer.
        4.  For "remarksContent", provide an empty string "" as this section is for the teacher's handwritten notes after the lesson.
        5.  Strictly return the output as a JSON object adhering to the provided schema. Do not add any extra text or explanations.
    `;

    try {
        const response = await callApiProxy({
            model,
            contents: prompt,
            config: {
                responseMimeType: "application/json",
                responseSchema: dlpContentSchema,
                safetySettings,
            },
            systemInstruction: efficientGenerationSystemInstruction
        });

        return parseJsonFromAiResponse<DlpContent>(response.text);

    } catch (error) {
        throw handleGeminiError(error, 'generateDlpContent');
    }
};

const generatedQuizQuestionSchema = {
    type: Type.OBJECT,
    properties: {
        questionText: { type: Type.STRING },
        options: { type: Type.ARRAY, items: { type: Type.STRING }, description: "Only for Multiple Choice. Should have 4 options." },
        correctAnswer: { type: Type.STRING }
    },
    required: ["questionText", "correctAnswer"]
};
const generatedQuizSectionSchema = {
    type: Type.OBJECT,
    properties: {
        instructions: { type: Type.STRING },
        questions: { type: Type.ARRAY, items: generatedQuizQuestionSchema }
    },
    required: ["instructions", "questions"]
};
const dlpRubricItemSchema = {
    type: Type.OBJECT,
    properties: {
        criteria: { type: Type.STRING },
        points: { type: Type.NUMBER }
    },
    required: ["criteria", "points"]
};
const tosItemSchema = {
    type: Type.OBJECT,
    properties: {
        objective: { type: Type.STRING },
        cognitiveLevel: { type: Type.STRING },
        itemNumbers: { type: Type.STRING }
    },
    required: ["objective", "cognitiveLevel", "itemNumbers"]
};
const quizActivitySchema = {
    type: Type.OBJECT,
    properties: {
        activityName: { type: Type.STRING },
        activityInstructions: { type: Type.STRING },
        rubric: { type: Type.ARRAY, items: dlpRubricItemSchema }
    },
    required: ["activityName", "activityInstructions"]
};
const generatedQuizSchema = {
    type: Type.OBJECT,
    properties: {
        quizTitle: { type: Type.STRING },
        tableOfSpecifications: { type: Type.ARRAY, items: tosItemSchema },
        questionsByType: {
            type: Type.OBJECT,
            properties: {
                'Multiple Choice': { ...generatedQuizSectionSchema, description: "Multiple choice questions section" },
                'True or False': { ...generatedQuizSectionSchema, description: "True or False questions section" },
                'Identification': { ...generatedQuizSectionSchema, description: "Identification questions section" }
            }
        },
        activities: {
            type: Type.ARRAY,
            items: quizActivitySchema
        }
    },
    required: ["quizTitle", "questionsByType", "activities"]
};

export const generateQuizContent = async (options: { topic: string; numQuestions: number; quizTypes: QuizType[]; subject: string; gradeLevel: string; }): Promise<GeneratedQuiz> => {
    const { topic, numQuestions, quizTypes, subject, gradeLevel } = options;
    const model = "gemini-2.5-pro";
    const prompt = `
        Create a quiz for a Grade ${gradeLevel} ${subject} class on the topic: "${topic}".
        
        The quiz should include:
        1. A suitable title.
        2. A Table of Specifications (TOS) if applicable.
        3. ${numQuestions} questions for each of the following formats: ${quizTypes.join(', ')}.
        4. For Multiple Choice, provide 4 options.
        5. For Identification, the answer should be a single word or short phrase.
        6. For True or False, the answer is "True" or "False".
        7. One or two creative, engaging performance task activities related to the topic. Do not generate a rubric for these activities initially.
        8. Adhere strictly to the JSON schema.
    `;
    try {
        const response = await callApiProxy({
            model, contents: prompt,
            config: { responseMimeType: "application/json", responseSchema: generatedQuizSchema, safetySettings },
            systemInstruction: efficientGenerationSystemInstruction
        });
        return parseJsonFromAiResponse<GeneratedQuiz>(response.text);
    } catch (error) {
        throw handleGeminiError(error, 'generateQuizContent');
    }
};

const rubricSchema = { type: Type.ARRAY, items: dlpRubricItemSchema };
export const generateRubricForActivity = async (options: { activityName: string; activityInstructions: string; totalPoints: number; }): Promise<DlpRubricItem[]> => {
    const model = "gemini-2.5-flash";
    const prompt = `
        Create a simple rubric for the following activity. The total points for the rubric must sum up to exactly ${options.totalPoints}.
        - Activity Name: ${options.activityName}
        - Instructions: ${options.activityInstructions}
        
        Return a JSON array of criteria and points.
    `;
    try {
        const response = await callApiProxy({
            model, contents: prompt,
            config: { responseMimeType: "application/json", responseSchema: rubricSchema, safetySettings },
            systemInstruction: efficientGenerationSystemInstruction
        });
        return parseJsonFromAiResponse<DlpRubricItem[]>(response.text);
    } catch (error) {
        throw handleGeminiError(error, 'generateRubricForActivity');
    }
};

const dllDailyEntrySchema = { type: Type.OBJECT, properties: { monday: { type: Type.STRING }, tuesday: { type: Type.STRING }, wednesday: { type: Type.STRING }, thursday: { type: Type.STRING }, friday: { type: Type.STRING } }, required: ["monday", "tuesday", "wednesday", "thursday", "friday"] };
const dllProcedureSchema = { type: Type.OBJECT, properties: { procedure: { type: Type.STRING }, ...dllDailyEntrySchema.properties }, required: ["procedure", "monday", "tuesday", "wednesday", "thursday", "friday"] };
const dllContentSchema = { type: Type.OBJECT, properties: { contentStandard: { type: Type.STRING }, performanceStandard: { type: Type.STRING }, learningCompetencies: dllDailyEntrySchema, content: { type: Type.STRING }, learningResources: { type: Type.OBJECT, properties: { teacherGuidePages: dllDailyEntrySchema, learnerMaterialsPages: dllDailyEntrySchema, textbookPages: dllDailyEntrySchema, additionalMaterials: dllDailyEntrySchema, otherResources: dllDailyEntrySchema }, required: ["teacherGuidePages", "learnerMaterialsPages", "textbookPages", "additionalMaterials", "otherResources"] }, procedures: { type: Type.ARRAY, items: dllProcedureSchema }, remarks: { type: Type.STRING }, reflection: { type: Type.ARRAY, items: dllProcedureSchema } }, required: ["contentStandard", "performanceStandard", "learningCompetencies", "content", "learningResources", "procedures", "remarks", "reflection"] };
export const generateDllContent = async (options: { subject: string; gradeLevel: string; weeklyTopic?: string; contentStandard?: string; performanceStandard?: string; dllFormat: string; language: 'English' | 'Filipino' }): Promise<DllContent> => {
    const { subject, gradeLevel, weeklyTopic, contentStandard, performanceStandard, dllFormat, language } = options;
    const model = "gemini-2.5-pro";
    const prompt = `
        Create a complete Daily Lesson Log (DLL) in ${language} for a Grade ${gradeLevel} ${subject} class for one week.
        - Weekly Topic: ${weeklyTopic || '(Suggest a relevant topic based on the subject and grade)'}
        - Content Standard: ${contentStandard || '(Generate a relevant standard)'}
        - Performance Standard: ${performanceStandard || '(Generate a relevant standard)'}
        - DLL Format: ${dllFormat}

        Generate content for all sections: Objectives (Learning Competencies), Content, Learning Resources, and Procedures for Monday to Friday.
        Procedures should be detailed and follow the specified format.
        Important: For the 'remarks' field, provide an empty string "". For the 'reflection' field, provide an empty array []. These sections are for the teacher to fill out manually after the lesson.
        Return a JSON object adhering to the schema.
    `;
    try {
        const response = await callApiProxy({
            model, contents: prompt,
            config: { responseMimeType: "application/json", responseSchema: dllContentSchema, safetySettings },
            systemInstruction: efficientGenerationSystemInstruction
        });
        return parseJsonFromAiResponse<DllContent>(response.text);
    } catch (error) {
        throw handleGeminiError(error, 'generateDllContent');
    }
};

const lasQuestionSchema = { type: Type.OBJECT, properties: { questionText: { type: Type.STRING }, type: { type: Type.STRING, enum: ['Identification', 'Essay', 'Problem-solving', 'Multiple Choice'] }, options: { type: Type.ARRAY, items: { type: Type.STRING } }, answer: { type: Type.STRING } }, required: ["questionText", "type"] };
const lasActivitySchema = { type: Type.OBJECT, properties: { title: { type: Type.STRING }, instructions: { type: Type.STRING }, questions: { type: Type.ARRAY, items: lasQuestionSchema }, rubric: { type: Type.ARRAY, items: dlpRubricItemSchema } }, required: ["title", "instructions"] };
const lasContentSchema = { type: Type.OBJECT, properties: { activityTitle: { type: Type.STRING }, learningTarget: { type: Type.STRING }, references: { type: Type.STRING }, conceptNotes: { type: Type.ARRAY, items: { type: Type.OBJECT, properties: { title: { type: Type.STRING }, content: { type: Type.STRING } }, required: ["title", "content"] } }, activities: { type: Type.ARRAY, items: lasActivitySchema } }, required: ["activityTitle", "learningTarget", "references", "conceptNotes", "activities"] };
export const generateLearningActivitySheet = async (options: { subject: string; gradeLevel: string; learningCompetency: string; lessonObjective: string; activityType: string; language: 'English' | 'Filipino' }): Promise<LearningActivitySheet> => {
    const { subject, gradeLevel, learningCompetency, lessonObjective, activityType, language } = options;
    const model = "gemini-2.5-pro";
    const prompt = `
        Create a DLP-style Learning Activity Sheet (LAS) in ${language} for a Grade ${gradeLevel} ${subject} class.
        - Learning Competency: ${learningCompetency}
        - Lesson Objective: ${lessonObjective}
        - Activity Focus: ${activityType}

        Generate all parts: Title, Learning Target, References, Concept Notes, and Activities.
        - Concept Notes should be clear and concise.
        - Activities should be engaging and aligned with the objective and focus. Include questions or performance tasks.
        - For performance tasks, suggest a simple rubric.
        Return a JSON object adhering to the schema.
    `;
    try {
        const response = await callApiProxy({
            model, contents: prompt,
            config: { responseMimeType: "application/json", responseSchema: lasContentSchema, safetySettings },
            systemInstruction: efficientGenerationSystemInstruction
        });
        return parseJsonFromAiResponse<LearningActivitySheet>(response.text);
    } catch (error) {
        throw handleGeminiError(error, 'generateLearningActivitySheet');
    }
};

const generatedExamTosItemSchema = { type: Type.OBJECT, properties: { objective: { type: Type.STRING }, daysTaught: { type: Type.NUMBER }, percentage: { type: Type.STRING }, numItems: { type: Type.NUMBER }, itemPlacement: { type: Type.STRING }, remembering: { type: Type.STRING }, understanding: { type: Type.STRING }, applying: { type: Type.STRING }, analyzing: { type: Type.STRING }, evaluating: { type: Type.STRING }, creating: { type: Type.STRING } }, required: ["objective", "daysTaught", "percentage", "numItems", "itemPlacement", "remembering", "understanding", "applying", "analyzing", "evaluating", "creating"] };
const generatedExamSchema = { type: Type.OBJECT, properties: { title: { type: Type.STRING }, tableOfSpecifications: { type: Type.ARRAY, items: generatedExamTosItemSchema }, questions: { type: Type.ARRAY, items: generatedQuizQuestionSchema }, subject: { type: Type.STRING }, gradeLevel: { type: Type.STRING }, quarter: { type: Type.STRING } }, required: ["title", "tableOfSpecifications", "questions", "subject", "gradeLevel", "quarter"] };
export const generateExam = async (options: { objectives: { text: string; days: string }[]; subject: string; gradeLevel: string; quarter: string; }): Promise<GeneratedExam> => {
    const { objectives, subject, gradeLevel, quarter } = options;
    const model = "gemini-2.5-pro";
    const prompt = `
        You are an expert curriculum designer for the Department of Education. Your task is to generate a complete 50-item periodical examination.

        **Inputs:**
        - Subject: ${subject}
        - Grade Level: ${gradeLevel}
        - Quarter: ${quarter}
        - Learning Objectives and Days Taught: ${JSON.stringify(objectives)}

        **Instructions:**

        1.  **Create a Table of Specifications (TOS):**
            - The total number of items MUST be exactly 50.
            - Calculate the percentage for each objective based on the days taught: \`(Days for Objective / Total Days) * 100%\`.
            - Calculate the number of items for each objective: \`round(Percentage * 50)\`. Adjust rounding to ensure the total is exactly 50 items.
            - Distribute the number of items for each objective across the 6 levels of Bloom's Taxonomy: Remembering, Understanding, Applying, Analyzing, Evaluating, Creating. The distribution should be appropriate for the grade level and subject. Use empty strings "" for levels with 0 items.
            - Determine the item placement for each objective (e.g., "1-5", "6, 10, 12-14"). The placement must match the number of items.

        2.  **Generate Test Questions:**
            - Based STRICTLY on the TOS you just created, generate 50 multiple-choice questions.
            - Each question must have 4 options.
            - The questions must be numbered from 1 to 50 and must correspond to the 'Item Placement' in the TOS.
            - The cognitive level of each question must match the distribution in the TOS.

        3.  **Output:**
            - Generate a single, valid JSON object that adheres to the provided schema. Do not include any text or explanations outside the JSON structure.
    `;

    try {
        const response = await callApiProxy({
            model, contents: prompt,
            config: { responseMimeType: "application/json", responseSchema: generatedExamSchema, safetySettings },
            systemInstruction: efficientGenerationSystemInstruction
        });
        return parseJsonFromAiResponse<GeneratedExam>(response.text);
    } catch (error) {
        throw handleGeminiError(error, 'generateExam');
    }
};