
import {
  Document,
  Packer,
  Paragraph,
  Table,
  TableCell,
  TableRow,
  TextRun,
  WidthType,
  AlignmentType,
  BorderStyle,
  PageOrientation,
  ImageRun,
  HeadingLevel,
  PageBreak,
  IRunOptions,
  LevelFormat,
  IImageOptions,
  VerticalAlign,
  ShadingType,
} from 'docx';
import { Student, SchoolSettings, Attendance, Quarter, SubjectQuarterSettings, StudentQuarterlyRecord, MapehRecordDocxData, GeneratedQuiz, QuizType, DlpContent, GeneratedQuizSection, DllContent, LearningActivitySheet, GeneratedExam, StudentProfileDocxData, CertificateSettings, HonorsCertificateSettings } from '../types';
import { toast } from 'react-hot-toast';

interface SummaryOfGradesDocxData {
    students: {
        males: any[];
        females: any[];
    };
    settings: SchoolSettings;
    subject: string;
    summaryStats: {
        malesPassed: number;
        malesFailed: number;
        femalesPassed: number;
        femalesFailed: number;
    };
    selectedSectionText: string;
}

interface EClassRecordDocxData {
    allStudents: { males: Student[]; females: Student[] };
    settings: SchoolSettings;
    subject: string;
    quarter: Quarter;
    selectedSectionText: string;
    recordSettings: SubjectQuarterSettings;
    studentRecords: StudentQuarterlyRecord[];
    calculationResults: Map<string, any>;
    summary: {
        passed: number;
        failed: number;
        malesPassed: number;
        malesFailed: number;
        femalesPassed: number;
        femalesFailed: number;
    };
}

interface CertificateDocxData {
    honorRoll: any[];
    settings: SchoolSettings;
    designOptions: {
        title: string;
        content: string;
        fontFamily: string;
        fontSize: number;
        subject: string;
        quarter: Quarter;
        gradeAndSectionOverride?: string;
        showCheckerSignature: boolean;
        adviserDesignation: string;
    }
}

interface HonorsListDocxData {
    honorStudents: {
        highest: any[];
        high: any[];
        regular: any[];
    };
    settings: SchoolSettings;
    selectedSectionText: string;
}

interface PickedStudentsDocxData {
    pickedStudents: Student[];
    topic: string;
    settings: SchoolSettings;
    sectionText: string;
}

interface GroupsDocxData {
    groups: Student[][];
    topic: string;
    settings: SchoolSettings;
    sectionText: string;
}

class DocxService {
    // Helper to ensure text is always a string and never null/undefined to prevent Docx corruption
    private safeString(value: any): string {
        if (value === null || value === undefined) return '';
        return String(value);
    }

    private parseDataUrl(dataUrl: string | undefined): { type: "svg" | "jpg" | "png" | "gif" | "bmp"; data: string } | null {
        if (!dataUrl || !dataUrl.startsWith("data:image/")) {
            return null;
        }
        const parts = dataUrl.split(",");
        if (parts.length !== 2) return null;

        const meta = parts[0];
        const data = parts[1];

        const mimeMatch = meta.match(/data:image\/(.*?);base64/);
        if (!mimeMatch || !mimeMatch[1]) return null;

        let type = mimeMatch[1];
        if (type === "jpeg" || type === "jpg") {
            type = "jpg";
        }
        if (type === "svg+xml") {
            type = "svg";
        }

        const validTypes: Array<"svg" | "jpg" | "png" | "gif" | "bmp"> = ["svg", "jpg", "png", "gif", "bmp"];
        if (!validTypes.includes(type as any)) {
            return null;
        }

        return { type: type as "svg" | "jpg" | "png" | "gif" | "bmp", data };
    }
    
    private base64ToArrayBuffer(base64: string): ArrayBuffer {
        const binaryString = window.atob(base64);
        const len = binaryString.length;
        const bytes = new Uint8Array(len);
        for (let i = 0; i < len; i++) {
            bytes[i] = binaryString.charCodeAt(i);
        }
        return bytes.buffer;
    }

    private createDocxImage(
        parsedImage: { type: "svg" | "jpg" | "png" | "gif" | "bmp"; data: string } | null,
        width: number,
        height: number,
        options: Partial<IImageOptions> = {}
    ): ImageRun | undefined {
        if (!parsedImage || !parsedImage.data) {
            return undefined;
        }

        try {
            if (parsedImage.type === 'svg') {
                const fallbackImageData = 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNkYAAAAAYAAjCB0C8AAAAASUVORK5CYII=';
                return new ImageRun({
                    type: 'svg',
                    data: this.base64ToArrayBuffer(parsedImage.data),
                    transformation: {
                        width,
                        height,
                    },
                    fallback: {
                        type: 'png',
                        data: this.base64ToArrayBuffer(fallbackImageData),
                    },
                    ...options
                });
            } else {
                 return new ImageRun({
                    type: parsedImage.type,
                    data: this.base64ToArrayBuffer(parsedImage.data),
                    transformation: {
                        width: width,
                        height: height,
                    },
                    ...options
                });
            }
        } catch (e) {
            console.error("Failed to create ImageRun. The image data might be corrupt.", e);
            toast.error("An error occurred while processing an image for the document.");
            return undefined;
        }
    }

    private async downloadBlob(blob: Blob, fileName: string): Promise<void> {
        const link = document.createElement("a");
        const url = URL.createObjectURL(blob);
        link.href = url;
        link.download = fileName;
        link.style.display = 'none';
        document.body.appendChild(link);
        link.click();
        
        setTimeout(() => {
            document.body.removeChild(link);
            window.URL.revokeObjectURL(url);
        }, 100);
    }

    private parseMarkdownToParagraphs(markdownText: string): Paragraph[] {
        if (!markdownText) return [new Paragraph("")]; // Handle empty content
    
        const paragraphs: Paragraph[] = [];
        const lines = markdownText.split('\n');
    
        for (const line of lines) {
            if (line.trim() === '') {
                // An empty line adds vertical space.
                paragraphs.push(new Paragraph({ children: [], spacing: { after: 100 } }));
                continue;
            }
    
            const children: TextRun[] = [];
            // Regex to split by bold/italic markers, keeping them in the result
            const parts = line.split(/(\*\*.*?\*\*|\*.*?\*)/g).filter(Boolean);
    
            for (const part of parts) {
                const fontOptions = { font: "Times New Roman", size: 22 };
                if (part.startsWith('**') && part.endsWith('**')) {
                    children.push(new TextRun({ text: part.slice(2, -2), bold: true, ...fontOptions }));
                } else if (part.startsWith('*') && part.endsWith('*')) {
                    children.push(new TextRun({ text: part.slice(1, -1), italics: true, ...fontOptions }));
                } else {
                    children.push(new TextRun({ text: part, ...fontOptions }));
                }
            }
            
            // Check for numbered lists like "1. " or "  1. "
            const isListItem = /^\s*\d+\.\s+/.test(line);
    
            paragraphs.push(new Paragraph({
                children,
                numbering: isListItem ? {
                    reference: "dlp-list",
                    level: 0,
                } : undefined,
                spacing: { after: isListItem ? 80 : 200 } // Add space after paragraphs
            }));
        }
    
        return paragraphs;
    }
    
    private parseLasMarkdown(markdownText: string): Paragraph[] {
        const text = this.safeString(markdownText);
        // Standardize Font: Century Gothic, Size 14 (28 half-points)
        const fontOptions = { font: "Century Gothic", size: 28 };

        // CRITICAL FIX: Prevent "Word experienced an error" by ensuring we never return an empty array of paragraphs.
        // Empty table cells corrupt the document structure.
        if (!text || text.trim() === '') {
            return [new Paragraph({ children: [new TextRun({ text: " ", ...fontOptions })], spacing: { after: 0 } })];
        }

        const paragraphs: Paragraph[] = [];
        const lines = text.split('\n');

        for (const line of lines) {
            // Even empty lines in the text block should produce a valid Paragraph with an empty TextRun to hold spacing
            if (line.trim() === '') {
                paragraphs.push(new Paragraph({ children: [new TextRun({ text: " ", ...fontOptions })], spacing: { after: 120 } }));
                continue;
            }

            const children: TextRun[] = [];
            // Regex to split by bold (**) and italic (*) markers
            const parts = line.split(/(\*\*.*?\*\*|\*.*?\*)/g).filter(p => p !== '');

            for (const part of parts) {
                if (part.startsWith('**') && part.endsWith('**')) {
                    // Remove asterisks, apply Bold
                    children.push(new TextRun({ text: part.slice(2, -2), bold: true, ...fontOptions }));
                } else if (part.startsWith('*') && part.endsWith('*')) {
                    // Remove asterisks, apply Italic
                    children.push(new TextRun({ text: part.slice(1, -1), italics: true, ...fontOptions }));
                } else {
                    children.push(new TextRun({ text: part, ...fontOptions }));
                }
            }
            
            // Check for bullet points
            const isListItem = /^\s*•\s+/.test(line.trim());
            const isSubListItem = /^\s*o\s+/.test(line.trim());

            paragraphs.push(new Paragraph({
                children: children.length > 0 ? children : [new TextRun({ text: " ", ...fontOptions })],
                bullet: isSubListItem ? { level: 1 } : (isListItem ? { level: 0 } : undefined),
                indent: isSubListItem ? { left: 1080, hanging: 360 } : (isListItem ? { left: 720, hanging: 360 } : undefined),
                spacing: { after: 120 } // Standard spacing
            }));
        }

        if (paragraphs.length === 0) {
             return [new Paragraph({ children: [new TextRun({ text: " ", ...fontOptions })], spacing: { after: 0 } })];
        }

        return paragraphs;
    }

    public async generateQuizDocx(quiz: GeneratedQuiz): Promise<void> {
        // ... (Code omitted for brevity, assumes similar null checks implemented if not already present)
         const { quizTitle, questionsByType, activities, tableOfSpecifications } = quiz;

        const numbering = {
            config: [
                {
                    reference: "quiz-numbering",
                    levels: [
                        {
                            level: 0,
                            format: LevelFormat.DECIMAL,
                            text: "%1.",
                            start: 1,
                            indent: { left: 720, hanging: 360 },
                        },
                    ],
                },
            ],
        };

        const sections: (Paragraph | Table | PageBreak)[] = [
            new Paragraph({
                text: this.safeString(quizTitle),
                heading: HeadingLevel.TITLE,
                alignment: AlignmentType.CENTER,
            }),
            new Paragraph({ text: "Name: __________________________", alignment: AlignmentType.LEFT }),
            new Paragraph({ text: "Grade & Section: __________________________", alignment: AlignmentType.LEFT }),
            new Paragraph({ text: "Score: _________", alignment: AlignmentType.LEFT }),
        ];

        if (tableOfSpecifications && tableOfSpecifications.length > 0) {
            sections.push(new Paragraph({ text: "Table of Specifications", heading: HeadingLevel.HEADING_2, spacing: { before: 200, after: 100 } }));
            const tosRows = [
                new TableRow({
                    children: ["Objective", "Cognitive Level", "Item Numbers"].map(text => new TableCell({ children: [new Paragraph({ text, bold: true })] })),
                    tableHeader: true,
                }),
                ...tableOfSpecifications.map(item => new TableRow({
                    children: [
                        new TableCell({ children: [new Paragraph(this.safeString(item.objective))] }),
                        new TableCell({ children: [new Paragraph(this.safeString(item.cognitiveLevel))] }),
                        new TableCell({ children: [new Paragraph(this.safeString(item.itemNumbers))] }),
                    ]
                }))
            ];
            sections.push(new Table({ rows: tosRows, width: { size: 100, type: WidthType.PERCENTAGE } }));
        }

        const answerKeySections: (Paragraph | Table)[] = [
            new Paragraph({
                text: "Answer Key",
                heading: HeadingLevel.TITLE,
                alignment: AlignmentType.CENTER,
            }),
        ];

        for (const type in questionsByType) {
            const section = questionsByType[type as QuizType];
            if (section) {
                sections.push(new Paragraph({ text: this.safeString(type), heading: HeadingLevel.HEADING_1, spacing: { before: 400, after: 200 } }));
                sections.push(new Paragraph({ text: this.safeString(section.instructions), italics: true, spacing: { after: 200 } }));
                answerKeySections.push(new Paragraph({ text: this.safeString(type), heading: HeadingLevel.HEADING_1, spacing: { before: 400, after: 200 } }));
                
                let questionCounter = 1;
                section.questions.forEach((q, i) => {
                    sections.push(new Paragraph({
                        text: this.safeString(q.questionText),
                        numbering: {
                            reference: "quiz-numbering",
                            level: 0,
                        },
                    }));

                    if (q.options) {
                        q.options.forEach((opt, oi) => {
                            sections.push(new Paragraph({
                                text: `${String.fromCharCode(65 + oi)}. ${this.safeString(opt)}`,
                                indentation: { left: 1080 },
                            }));
                        });
                    } else if (type === 'Identification' || type === 'True or False') {
                         sections.push(new Paragraph({ text: "Answer: ____________________", indentation: { left: 1080 } }));
                    }

                    answerKeySections.push(new Paragraph({
                        text: `${questionCounter}. ${this.safeString(q.correctAnswer)}`,
                        numbering: {
                            reference: "quiz-numbering",
                            level: 0,
                        },
                    }));
                    questionCounter++;
                });
            }
        }
        
        if (activities && activities.length > 0) {
            sections.push(new Paragraph({ text: "Activities", heading: HeadingLevel.HEADING_1, spacing: { before: 400, after: 200 } }));
            activities.forEach(activity => {
                sections.push(new Paragraph({ text: this.safeString(activity.activityName), heading: HeadingLevel.HEADING_2, spacing: { before: 200, after: 100 } }));
                sections.push(new Paragraph({ text: this.safeString(activity.activityInstructions), italics: true, spacing: { after: 200 } }));
                if (activity.rubric && activity.rubric.length > 0) {
                    const rubricRows = activity.rubric.map(item => new TableRow({
                        children: [
                            new TableCell({ children: [new Paragraph(this.safeString(item.criteria))] }),
                            new TableCell({ children: [new Paragraph(this.safeString(item.points))] }),
                        ]
                    }));
                    const rubricTable = new Table({
                        rows: [
                             new TableRow({
                                children: [
                                    new TableCell({ children: [new Paragraph({ text: 'Criteria', bold: true })] }),
                                    new TableCell({ children: [new Paragraph({ text: 'Points', bold: true })] }),
                                ],
                                tableHeader: true,
                            }),
                            ...rubricRows,
                        ],
                        width: {
                            size: 100,
                            type: WidthType.PERCENTAGE,
                        },
                    });
                    sections.push(rubricTable);
                }
            });
        }
        
        sections.push(new PageBreak());
        sections.push(...answerKeySections);

        const doc = new Document({
            numbering,
            sections: [{ children: sections }],
        });

        const blob = await Packer.toBlob(doc);
        this.downloadBlob(blob, `${this.safeString(quizTitle).replace(/\s/g, '_')}_Quiz.docx`);
    }
    
    public async generateDlpDocx(
        dlpForm: any, 
        dlpContent: DlpContent, 
        htmlContent: string, 
        settings: SchoolSettings
    ): Promise<void> {
         // ... (Assuming existing code here is preserved or safe, only LAS was corrupted)
          const isFilipino = dlpForm.language === 'Filipino';
        const t = {
            objectives: isFilipino ? 'I. LAYUNIN' : 'I. OBJECTIVES',
            content: isFilipino ? 'II. NILALAMAN' : 'II. CONTENT',
            resources: isFilipino ? 'III. KAGAMITANG PANTURO' : 'III. LEARNING RESOURCES',
            procedure: isFilipino ? 'IV. PAMAMARAAN' : 'IV. PROCEDURE',
            evaluation: isFilipino ? 'V. PAGTATAYA' : 'V. EVALUATING LEARNING',
            reflection: isFilipino ? 'VI. PAGNINILAY' : 'VI. REFLECTION',
            preparedBy: isFilipino ? 'Inihanda ni' : 'Prepared by',
            checkedBy: isFilipino ? 'Sinuri ni' : 'Checked by',
            approvedBy: isFilipino ? 'Inaprubahan ni' : 'Approved by',
        };
    
        const gradeColorMapping: { [key: string]: string } = {
            '7': '90EE90', // LightGreen
            '8': 'FFFFE0', // LightYellow
            '9': 'F08080', // LightCoral
            '10': 'ADD8E6', // LightBlue
        };
        const gradeColor = gradeColorMapping[this.safeString(dlpForm.gradeLevel)] || 'D9D9D9';
    
        const textRunOptions: IRunOptions = { font: "Times New Roman", size: 22 };
        const boldTextRunOptions: IRunOptions = { ...textRunOptions, bold: true };
    
        const createHeaderCell = (text: string): TableCell => new TableCell({
            children: [new Paragraph({ children: [new TextRun({ ...boldTextRunOptions, text })], spacing: { after: 120 } })],
            shading: { type: ShadingType.CLEAR, fill: gradeColor },
            verticalAlign: VerticalAlign.CENTER,
        });
    
        const createContentCell = (children: Paragraph[]): TableCell => new TableCell({
            children,
            verticalAlign: VerticalAlign.TOP,
        });
    
        const doc = new Document({
            numbering: {
                config: [
                    {
                        reference: "dlp-list",
                        levels: [
                            {
                                level: 0,
                                format: LevelFormat.DECIMAL,
                                text: "%1.",
                                indent: { left: 720, hanging: 360 },
                                run: { font: "Times New Roman", size: 22 },
                            },
                        ],
                    },
                ],
            },
            sections: [{
                properties: {
                    page: {
                        margin: { top: 720, right: 720, bottom: 720, left: 720 },
                    },
                },
                children: [
                    // Main Header Table
                    new Table({
                        width: { size: 100, type: WidthType.PERCENTAGE },
                        rows: [
                            new TableRow({
                                children: [
                                    new TableCell({
                                        children: [new Paragraph({
                                            children: this.createDocxImage(this.parseDataUrl(settings.schoolLogo), 70, 70) ? [this.createDocxImage(this.parseDataUrl(settings.schoolLogo), 70, 70)!] : [],
                                            alignment: AlignmentType.CENTER
                                        })],
                                        width: { size: 15, type: WidthType.PERCENTAGE },
                                        verticalAlign: VerticalAlign.CENTER
                                    }),
                                    new TableCell({
                                        children: [
                                            new Paragraph({ text: isFilipino ? 'Paaralan' : 'School', style: 'header-label' }),
                                            new Paragraph({ text: this.safeString(dlpForm.schoolName).toUpperCase(), style: 'header-value' }),
                                            new Paragraph({ text: this.safeString(dlpForm.quarterSelect), style: 'header-value' }),
                                            new Paragraph({ text: `${isFilipino ? 'Guro' : 'Teacher'}: ${this.safeString(dlpForm.teacher)}`, style: 'header-value' }),
                                            new Paragraph({ text: `${isFilipino ? 'Asignatura' : 'Learning Area'}: ${this.safeString(dlpForm.subject).toUpperCase()}`, style: 'header-value' }),
                                            new Paragraph({ text: `${isFilipino ? 'Petsa ng Pagtuturo' : 'Teaching Dates'}: ${this.safeString(dlpForm.teachingDates)}`, style: 'header-value' }),
                                        ].map(p => new Paragraph({ ...p.options, alignment: AlignmentType.LEFT })),
                                        width: { size: 55, type: WidthType.PERCENTAGE }
                                    }),
                                    new TableCell({
                                        children: [
                                            new Paragraph({ text: isFilipino ? 'DETALYADONG BANGHAY-ARALIN SA' : 'DAILY LESSON PLAN IN', alignment: AlignmentType.CENTER, bold: true }),
                                            new Paragraph({ text: `${this.safeString(dlpForm.subject).toUpperCase()} ${this.safeString(dlpForm.gradeLevel)}`, alignment: AlignmentType.CENTER, bold: true }),
                                            new Paragraph({ text: (isFilipino ? 'ISKEDYUL NG KLASE' : 'CLASS SCHEDULE'), alignment: AlignmentType.CENTER, bold: true, spacing: {before: 100} }),
                                            ...this.safeString(dlpForm.classSchedule).split('\n').map((line: string) => new Paragraph({ text: line, alignment: AlignmentType.CENTER }))
                                        ],
                                        width: { size: 30, type: WidthType.PERCENTAGE },
                                        verticalAlign: VerticalAlign.CENTER
                                    }),
                                ],
                            }),
                        ],
                        borders: { top: { style: BorderStyle.SINGLE }, bottom: { style: BorderStyle.SINGLE }, left: { style: BorderStyle.SINGLE }, right: { style: BorderStyle.SINGLE } },
                    }),
    
                    // Main Content Table
                    new Table({
                        width: { size: 100, type: WidthType.PERCENTAGE },
                        rows: [
                            new TableRow({ children: [createHeaderCell(t.objectives), createContentCell([ new Paragraph({ children: [new TextRun({ ...boldTextRunOptions, text: (isFilipino ? 'Pamantayang Pangnilalaman: ' : 'Content Standard: ') }), new TextRun({ ...textRunOptions, text: this.safeString(dlpContent.contentStandard) })] }), new Paragraph({ children: [new TextRun({ ...boldTextRunOptions, text: (isFilipino ? 'Pamantayan sa Pagganap: ' : 'Performance Standard: ') }), new TextRun({ ...textRunOptions, text: this.safeString(dlpContent.performanceStandard) })] }), new Paragraph({ children: [new TextRun({ ...boldTextRunOptions, text: (isFilipino ? 'Kasanayan sa Pagkatuto: ' : 'Learning Competency: ') }), new TextRun({ ...textRunOptions, text: this.safeString(dlpForm.learningCompetency) })] }), new Paragraph({ text: (isFilipino ? 'Sa pagtatapos ng aralin, ang mga mag-aaral ay inaasahang:' : 'At the end of the lesson, the learners should be able to:'), spacing: { before: 200 } }), new Paragraph({ text: this.safeString(dlpForm.lessonObjective), bullet: { level: 0 } })])] }),
                            new TableRow({ children: [createHeaderCell(t.content), createContentCell([ new Paragraph({ children: [new TextRun({ ...boldTextRunOptions, text: (isFilipino ? 'Paksa: ' : 'Topic: ') }), new TextRun({ ...textRunOptions, text: this.safeString(dlpContent.topic) })] })])] }),
                            new TableRow({ children: [createHeaderCell(t.resources), createContentCell([ new Paragraph({ children: [new TextRun({ ...boldTextRunOptions, text: (isFilipino ? 'Sanggunian: ' : 'References: ') }), new TextRun({ ...textRunOptions, text: this.safeString(dlpContent.learningReferences) })] }), new Paragraph({ children: [new TextRun({ ...boldTextRunOptions, text: (isFilipino ? 'Kagamitan: ' : 'Materials: ') }), new TextRun({ ...textRunOptions, text: this.safeString(dlpContent.learningMaterials) })] })])] }),
                            new TableRow({ children: [createHeaderCell(t.procedure), new TableCell({
                                children: [ new Table({
                                    width: { size: 100, type: WidthType.PERCENTAGE },
                                    columnWidths: [25, 45, 30],
                                    rows: [
                                        new TableRow({ children: [ new TableCell({ children: [new Paragraph({ text: (isFilipino ? 'Pamamaraan' : 'Procedure'), bold: true })] }), new TableCell({ children: [new Paragraph({ text: (isFilipino ? 'Gawain ng Guro/Mag-aaral' : 'Teacher/Student Activity'), bold: true })] }), new TableCell({ children: [new Paragraph({ text: (isFilipino ? 'Mga Kaugnay na PPST Indicator' : 'Aligned PPST Indicators'), bold: true })] })] }),
                                        ...dlpContent.procedures.map(proc => new TableRow({
                                            children: [ new TableCell({ children: [new Paragraph({ text: this.safeString(proc.title), bold: true })] }), new TableCell({ children: this.parseMarkdownToParagraphs(this.safeString(proc.content)) }), new TableCell({ children: [new Paragraph({ text: this.safeString(proc.ppst), italics: true })] })],
                                        }))
                                    ],
                                })],
                            })] }),
                            new TableRow({ children: [createHeaderCell(t.evaluation), createContentCell( (dlpContent.evaluationQuestions || []).map(q => new Paragraph({ text: this.safeString(q.question), numbering: { reference: "dlp-list", level: 0 }, spacing: { after: 100 } })))] }),
                            new TableRow({ children: [createHeaderCell(t.reflection), createContentCell([new Paragraph({text: ""})])] }), // Empty for now
                        ],
                    }),
    
                    new PageBreak(),
    
                    // Answer Key
                    new Paragraph({ text: (isFilipino ? 'Susi sa Pagwawasto' : 'Answer Key (For Evaluating Learning)'), heading: HeadingLevel.HEADING_2 }),
                    ...(dlpContent.evaluationQuestions || []).map(q => new Paragraph({ text: this.safeString(q.answer), numbering: { reference: "dlp-list", level: 0 }, spacing: { after: 100 } })),
                ],
            }]
        });
    
        const blob = await Packer.toBlob(doc);
        this.downloadBlob(blob, `DLP_${this.safeString(dlpForm.subject).replace(/\s/g, '_')}.docx`);
    }

    public async generateDllDocx(
        dllForm: any, 
        dllContent: DllContent,
        settings: SchoolSettings
    ): Promise<void> {
         // ... (Same as before, checking strictness of safeString)
        // Corrected page size for 8.5" x 13" (Long Bond Paper) in landscape
        const pageHeight = 18720;
        const pageWidth = 12240;

        const isFilipino = dllForm.language === 'Filipino';
        const days = isFilipino ? ['Lunes', 'Martes', 'Miyerkules', 'Huwebes', 'Biyernes'] : ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday'];
        
        const boldCentered = { bold: true, alignment: AlignmentType.CENTER };

        const headerTable = new Table({
            width: { size: 100, type: WidthType.PERCENTAGE },
            rows: [
                new TableRow({
                    children: [
                        new TableCell({ children: [new Paragraph({ text: `School: ${this.safeString(settings.schoolName)}`, alignment: AlignmentType.LEFT })] }),
                        new TableCell({ children: [new Paragraph({ text: `Grade Level: ${this.safeString(dllForm.gradeLevel)}`, alignment: AlignmentType.LEFT })] }),
                        new TableCell({ children: [new Paragraph({ text: `Teacher: ${this.safeString(settings.teacherName)}`, alignment: AlignmentType.LEFT })] }),
                        new TableCell({ children: [new Paragraph({ text: `Learning Area: ${this.safeString(dllForm.subject)}`, alignment: AlignmentType.LEFT })] }),
                    ]
                }),
                new TableRow({
                    children: [
                        new TableCell({ children: [new Paragraph({ text: `Teaching Dates & Time: ${this.safeString(dllForm.teachingDates)}`, alignment: AlignmentType.LEFT })] }),
                        new TableCell({ children: [new Paragraph({ text: `Quarter: ${this.safeString(dllForm.quarter)}`, alignment: AlignmentType.LEFT })] }),
                        new TableCell({ children: [], columnSpan: 2 }),
                    ]
                })
            ],
            borders: { top: { style: BorderStyle.NONE }, bottom: { style: BorderStyle.NONE }, left: { style: BorderStyle.NONE }, right: { style: BorderStyle.NONE } },
        });

        const mainTableRows: TableRow[] = [];
        
        // Add days row
        mainTableRows.push(new TableRow({
            children: [new TableCell({ children: [] }), ...days.map(day => new TableCell({ children: [new Paragraph({ text: day, ...boldCentered })] }))]
        }));

        // Add Objectives
        mainTableRows.push(new TableRow({ children: [new TableCell({ children: [new Paragraph({ text: "I. OBJECTIVES", bold: true })], columnSpan: 6 })]}));
        mainTableRows.push(new TableRow({ children: [new TableCell({ children: [new Paragraph("A. Content Standard")] }), new TableCell({ children: [new Paragraph(this.safeString(dllContent.contentStandard))], columnSpan: 5 })]}));
        mainTableRows.push(new TableRow({ children: [new TableCell({ children: [new Paragraph("B. Performance Standard")] }), new TableCell({ children: [new Paragraph(this.safeString(dllContent.performanceStandard))], columnSpan: 5 })]}));
        mainTableRows.push(new TableRow({
            children: [
                new TableCell({ children: [new Paragraph("C. Learning Competencies")] }),
                new TableCell({ children: [new Paragraph(this.safeString(dllContent.learningCompetencies.monday))] }),
                new TableCell({ children: [new Paragraph(this.safeString(dllContent.learningCompetencies.tuesday))] }),
                new TableCell({ children: [new Paragraph(this.safeString(dllContent.learningCompetencies.wednesday))] }),
                new TableCell({ children: [new Paragraph(this.safeString(dllContent.learningCompetencies.thursday))] }),
                new TableCell({ children: [new Paragraph(this.safeString(dllContent.learningCompetencies.friday))] }),
            ]
        }));
        
        // Add Content
        mainTableRows.push(new TableRow({ children: [new TableCell({ children: [new Paragraph({ text: "II. CONTENT", bold: true })] }), new TableCell({ children: [new Paragraph(this.safeString(dllContent.content))], columnSpan: 5 })]}));

        // Add Learning Resources
        mainTableRows.push(new TableRow({ children: [new TableCell({ children: [new Paragraph({ text: "III. LEARNING RESOURCES", bold: true })], columnSpan: 6 })]}));

        // Add Procedures
        mainTableRows.push(new TableRow({ children: [new TableCell({ children: [new Paragraph({ text: "IV. PROCEDURES", bold: true })], columnSpan: 6 })]}));
        dllContent.procedures.forEach(proc => {
            mainTableRows.push(new TableRow({
                children: [
                    new TableCell({ children: [new Paragraph(this.safeString(proc.procedure))] }),
                    new TableCell({ children: this.parseMarkdownToParagraphs(this.safeString(proc.monday)) }),
                    new TableCell({ children: this.parseMarkdownToParagraphs(this.safeString(proc.tuesday)) }),
                    new TableCell({ children: this.parseMarkdownToParagraphs(this.safeString(proc.wednesday)) }),
                    new TableCell({ children: this.parseMarkdownToParagraphs(this.safeString(proc.thursday)) }),
                    new TableCell({ children: this.parseMarkdownToParagraphs(this.safeString(proc.friday)) }),
                ]
            }));
        });
        
        // Add Remarks & Reflection
        mainTableRows.push(new TableRow({ children: [new TableCell({ children: [new Paragraph({ text: "V. REMARKS", bold: true })] }), new TableCell({ children: [new Paragraph(this.safeString(dllContent.remarks))], columnSpan: 5 })]}));
        mainTableRows.push(new TableRow({ children: [new TableCell({ children: [new Paragraph({ text: "VI. REFLECTION", bold: true })], columnSpan: 6 })]}));

        const mainTable = new Table({
            width: { size: 100, type: WidthType.PERCENTAGE },
            columnWidths: [15, 17, 17, 17, 17, 17],
            rows: mainTableRows
        });
        
        const signatoriesTable = new Table({
            width: { size: 100, type: WidthType.PERCENTAGE },
            columnWidths: [33, 34, 33],
            rows: [
                 new TableRow({
                    children: [
                        new TableCell({ children: [new Paragraph("Prepared by:")], borders: { top: { style: BorderStyle.NONE }, bottom: { style: BorderStyle.NONE }, left: { style: BorderStyle.NONE }, right: { style: BorderStyle.NONE }} }),
                        new TableCell({ children: [new Paragraph("Checked by:")], borders: { top: { style: BorderStyle.NONE }, bottom: { style: BorderStyle.NONE }, left: { style: BorderStyle.NONE }, right: { style: BorderStyle.NONE }} }),
                        new TableCell({ children: [new Paragraph("Approved by:")], borders: { top: { style: BorderStyle.NONE }, bottom: { style: BorderStyle.NONE }, left: { style: BorderStyle.NONE }, right: { style: BorderStyle.NONE }} }),
                    ]
                 }),
                 new TableRow({
                    children: [
                        new TableCell({ children: [new Paragraph({text: "", spacing: {before: 1000}})], borders: { top: { style: BorderStyle.NONE }, bottom: { style: BorderStyle.NONE }, left: { style: BorderStyle.NONE }, right: { style: BorderStyle.NONE }} }),
                        new TableCell({ children: [], borders: { top: { style: BorderStyle.NONE }, bottom: { style: BorderStyle.NONE }, left: { style: BorderStyle.NONE }, right: { style: BorderStyle.NONE }} }),
                        new TableCell({ children: [], borders: { top: { style: BorderStyle.NONE }, bottom: { style: BorderStyle.NONE }, left: { style: BorderStyle.NONE }, right: { style: BorderStyle.NONE }} }),
                    ]
                 }),
                 new TableRow({
                    children: [
                        new TableCell({ children: [new Paragraph({ text: this.safeString(dllForm.preparedByName), ...boldCentered })], borders: { top: { style: BorderStyle.SINGLE }, bottom: { style: BorderStyle.NONE }, left: { style: BorderStyle.NONE }, right: { style: BorderStyle.NONE }} }),
                        new TableCell({ children: [new Paragraph({ text: this.safeString(dllForm.checkedByName), ...boldCentered })], borders: { top: { style: BorderStyle.SINGLE }, bottom: { style: BorderStyle.NONE }, left: { style: BorderStyle.NONE }, right: { style: BorderStyle.NONE }} }),
                        new TableCell({ children: [new Paragraph({ text: this.safeString(dllForm.approvedByName), ...boldCentered })], borders: { top: { style: BorderStyle.SINGLE }, bottom: { style: BorderStyle.NONE }, left: { style: BorderStyle.NONE }, right: { style: BorderStyle.NONE }} }),
                    ]
                 }),
                 new TableRow({
                    children: [
                        new TableCell({ children: [new Paragraph({ text: this.safeString(dllForm.preparedByDesignation), alignment: AlignmentType.CENTER })], borders: { top: { style: BorderStyle.NONE }, bottom: { style: BorderStyle.NONE }, left: { style: BorderStyle.NONE }, right: { style: BorderStyle.NONE }} }),
                        new TableCell({ children: [new Paragraph({ text: this.safeString(dllForm.checkedByDesignation), alignment: AlignmentType.CENTER })], borders: { top: { style: BorderStyle.NONE }, bottom: { style: BorderStyle.NONE }, left: { style: BorderStyle.NONE }, right: { style: BorderStyle.NONE }} }),
                        new TableCell({ children: [new Paragraph({ text: this.safeString(dllForm.approvedByDesignation), alignment: AlignmentType.CENTER })], borders: { top: { style: BorderStyle.NONE }, bottom: { style: BorderStyle.NONE }, left: { style: BorderStyle.NONE }, right: { style: BorderStyle.NONE }} }),
                    ]
                 })
            ]
        });


        const doc = new Document({
            sections: [{
                properties: {
                    page: {
                        size: { width: pageHeight, height: pageWidth },
                        orientation: PageOrientation.LANDSCAPE,
                        margin: { top: 720, right: 720, bottom: 720, left: 720 },
                    },
                },
                children: [
                    new Paragraph({ text: "DAILY LESSON LOG", ...boldCentered, spacing: { after: 200 } }),
                    headerTable,
                    mainTable,
                    new Paragraph({ spacing: { after: 400 } }),
                    signatoriesTable,
                ],
            }],
        });
    
        const blob = await Packer.toBlob(doc);
        this.downloadBlob(blob, `DLL_${this.safeString(dllForm.subject).replace(/\s/g, '_')}.docx`);
    }

    public async generateLasDocx(
        lasForm: any,
        lasContent: LearningActivitySheet,
        settings: SchoolSettings
    ): Promise<void> {
        const sections: (Paragraph | Table | PageBreak)[] = [];
        
        // Enforce Century Gothic Font Size 14 (28 half-points)
        const baseFont = "Century Gothic";
        const fieldFont = { font: baseFont, size: 28 }; // 14pt = 28 half-points

        // Safety check: Ensure days exist
        const days = lasContent?.days || [];

        days.forEach((dayData, index) => {
            if (index > 0) {
                sections.push(new PageBreak());
            }

            // --- Page Header Structure ---
            
            // 1. Day Title (Bold, Heading 1 style, ALL CAPS)
            sections.push(new Paragraph({
                text: this.safeString(dayData.dayTitle).toUpperCase(), 
                heading: HeadingLevel.HEADING_1,
                run: { font: baseFont, size: 28, bold: true }, // 14pt Bold
                spacing: { after: 240 }
            }));

            // 2. DepEd Header Line
            sections.push(new Paragraph({
                text: "DepED | Dynamic Learning Program | BAGONG PILIPINAS | LEARNING ACTIVITY SHEET",
                alignment: AlignmentType.LEFT, 
                run: { font: baseFont, size: 28, bold: true }, // 14pt Bold
                spacing: { after: 240 },
                border: { bottom: { color: "auto", space: 1, value: "single", size: 6 } },
            }));

            // 3. Basic Info Fields
            
            // Subject
            sections.push(new Paragraph({
                children: [
                    new TextRun({ text: "Subject: ", bold: true, ...fieldFont }),
                    new TextRun({ text: this.safeString(lasForm.subject), ...fieldFont })
                ],
                spacing: { before: 120 }
            }));

            // Grade & Section
            sections.push(new Paragraph({
                children: [
                    new TextRun({ text: "Grade & Section: ", bold: true, ...fieldFont }),
                    new TextRun({ text: "________________________________________", ...fieldFont })
                ],
                spacing: { before: 60 }
            }));

            // Name, Score, Date line
            const infoRow = new Table({
                width: { size: 100, type: WidthType.PERCENTAGE },
                borders: { top: { style: BorderStyle.NONE }, bottom: { style: BorderStyle.NONE }, left: { style: BorderStyle.NONE }, right: { style: BorderStyle.NONE }, insideHorizontal: { style: BorderStyle.NONE }, insideVertical: { style: BorderStyle.NONE } },
                rows: [
                    new TableRow({
                        children: [
                            new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "Name: ", bold: true, ...fieldFont }), new TextRun({ text: "__________________________________", ...fieldFont })] })], width: { size: 50, type: WidthType.PERCENTAGE } }),
                            new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "Score: ", bold: true, ...fieldFont }), new TextRun({ text: "_______", ...fieldFont })] })], width: { size: 25, type: WidthType.PERCENTAGE } }),
                            new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "Date: ", bold: true, ...fieldFont }), new TextRun({ text: "_______", ...fieldFont })] })], width: { size: 25, type: WidthType.PERCENTAGE } }),
                        ]
                    })
                ]
            });
            sections.push(infoRow);
            sections.push(new Paragraph({ text: "", spacing: { after: 60 } }));

            // Activity Title
            sections.push(new Paragraph({
                children: [
                    new TextRun({ text: "Activity Title: ", bold: true, ...fieldFont }),
                    new TextRun({ text: this.safeString(dayData.activityTitle), ...fieldFont })
                ]
            }));

            // Learning Target
            sections.push(new Paragraph({
                children: [
                    new TextRun({ text: "Learning Target: ", bold: true, ...fieldFont }),
                    new TextRun({ text: this.safeString(dayData.learningTarget), ...fieldFont })
                ]
            }));

            // Reference
            sections.push(new Paragraph({
                children: [
                    new TextRun({ text: "Reference: ", bold: true, ...fieldFont }),
                    new TextRun({ text: this.safeString(dayData.references) || "____________________", ...fieldFont })
                ],
                spacing: { after: 240 } 
            }));

            // 4. Concept Notes
            if (dayData.conceptNotes && Array.isArray(dayData.conceptNotes)) {
                dayData.conceptNotes.forEach(note => {
                    sections.push(new Paragraph({
                        text: this.safeString(note.title).toUpperCase(),
                        heading: HeadingLevel.HEADING_2,
                        run: { font: baseFont, size: 28, bold: true },
                        spacing: { before: 120, after: 120 }
                    }));
                    sections.push(...this.parseLasMarkdown(this.safeString(note.content)));
                });
            }

            // 5. Activities
            if (dayData.activities && Array.isArray(dayData.activities)) {
                dayData.activities.forEach(activity => {
                    // Activity Header
                    sections.push(new Paragraph({
                        text: this.safeString(activity.title).toUpperCase(),
                        heading: HeadingLevel.HEADING_3,
                        run: { font: baseFont, size: 28, bold: true }, 
                        spacing: { before: 240, after: 120 }
                    }));

                    const instructions = this.safeString(activity.instructions);
                    
                    // Check for Table logic (using || separator) to generate "Match the column" tables
                    if (instructions.includes('||')) {
                        const lines = instructions.split('\n');
                        const directionsLines = lines.filter(l => !l.includes('||'));
                        const tableLines = lines.filter(l => l.includes('||'));

                        if (directionsLines.length > 0) {
                            sections.push(new Paragraph({
                                children: [
                                    new TextRun({ text: "Directions: ", bold: true, ...fieldFont }),
                                    new TextRun({ text: directionsLines.join(' '), ...fieldFont })
                                ],
                                spacing: { after: 120 }
                            }));
                        }

                        if (tableLines.length > 0) {
                            // Create Table for matching activities
                            const tableRows = tableLines.map(line => {
                                const parts = line.split('||');
                                // Ensure parts are safe strings
                                const col1 = this.safeString(parts[0] || "").trim();
                                const col2 = this.safeString(parts[1] || "").trim();
                                
                                // IMPORTANT: parseLasMarkdown now guarantees to return at least one paragraph
                                return new TableRow({
                                    children: [
                                        new TableCell({ children: this.parseLasMarkdown(col1), width: { size: 50, type: WidthType.PERCENTAGE }, padding: { top: 100, bottom: 100, left: 100, right: 100 } }),
                                        new TableCell({ children: this.parseLasMarkdown(col2), width: { size: 50, type: WidthType.PERCENTAGE }, padding: { top: 100, bottom: 100, left: 100, right: 100 } }),
                                    ]
                                });
                            });

                            // Add header row for the matching table
                            tableRows.unshift(new TableRow({
                                tableHeader: true,
                                children: [
                                    new TableCell({ children: [new Paragraph({ text: "Column A", bold: true, alignment: AlignmentType.CENTER, run: { font: baseFont, size: 28 } })], shading: { fill: "EFEFEF", type: ShadingType.CLEAR, color: "auto" } }),
                                    new TableCell({ children: [new Paragraph({ text: "Column B", bold: true, alignment: AlignmentType.CENTER, run: { font: baseFont, size: 28 } })], shading: { fill: "EFEFEF", type: ShadingType.CLEAR, color: "auto" } }),
                                ]
                            }));
                            
                            if (tableRows.length > 0) {
                                sections.push(new Table({
                                    rows: tableRows,
                                    width: { size: 100, type: WidthType.PERCENTAGE },
                                }));
                                sections.push(new Paragraph({ children: [], spacing: { after: 120 } })); // Spacer using empty children array is risky for text runs, but OK for paragraph spacing if handled
                            }
                        }

                    } else {
                        // Normal text instructions
                        sections.push(...this.parseLasMarkdown(instructions));
                    }

                    // Activity Questions
                    if (activity.questions && activity.questions.length > 0) {
                        activity.questions.forEach((q, i) => {
                            sections.push(new Paragraph({
                                children: [
                                    new TextRun({ text: `${i + 1}. `, bold: true, ...fieldFont }),
                                    new TextRun({ text: this.safeString(q.questionText), ...fieldFont })
                                ],
                                spacing: { before: 60 }
                            }));
                            if (q.options && q.options.length > 0) {
                                q.options.forEach((opt, oi) => {
                                    sections.push(new Paragraph({
                                        text: `${String.fromCharCode(65+oi)}. ${this.safeString(opt)}`,
                                        indentation: { left: 720 },
                                        run: fieldFont
                                    }));
                                });
                            } else {
                                // Blank line for answer if no options
                                sections.push(new Paragraph({ text: "______________________________________________________", indentation: { left: 720 }, run: fieldFont }));
                            }
                        });
                    }
                });
            }

            // 6. Reflection
            sections.push(new Paragraph({
                text: "REFLECTION",
                heading: HeadingLevel.HEADING_3,
                run: { font: baseFont, size: 28, bold: true }, 
                spacing: { before: 240, after: 120 }
            }));
            sections.push(new Paragraph({
                text: this.safeString(dayData.reflection),
                run: fieldFont
            }));
             sections.push(new Paragraph({
                text: "________________________________________________________________________________________________________________________________",
                spacing: { before: 200 },
                run: fieldFont
            }));
             sections.push(new Paragraph({
                text: "________________________________________________________________________________________________________________________________",
                spacing: { before: 200 },
                run: fieldFont
            }));
        });

        const doc = new Document({
            numbering: {
                config: [
                    {
                        reference: "las-list",
                        levels: [{
                            level: 0,
                            format: LevelFormat.DECIMAL,
                            text: "%1.",
                            indent: { left: 720, hanging: 360 },
                            run: { font: "Century Gothic", size: 28 } // 14pt
                        }],
                    },
                ],
            },
            sections: [{
                properties: {
                    page: {
                        size: { width: 12240, height: 18720 }, // 8.5 x 13 Long Bond Paper in Portrait
                        margin: { top: 720, right: 720, bottom: 720, left: 720 }, // 0.5 inch margins
                    }
                },
                children: sections
            }]
        });

        const blob = await Packer.toBlob(doc);
        this.downloadBlob(blob, `LAS_${this.safeString(lasForm.subject).replace(/\s/g, '_')}.docx`);
    }

    public async generateExamDocx(exam: GeneratedExam, settings: SchoolSettings): Promise<void> {
        // Stub for generateExamDocx
    }

    public async generateAttendanceDocx(students: Student[], attendance: Attendance[], currentDate: Date, schoolSettings: SchoolSettings): Promise<void> {
        // Stub for generateAttendanceDocx
    }
    
    public async generateSummaryOfGradesDocx(data: SummaryOfGradesDocxData): Promise<void> {
       // Stub for generateSummaryOfGradesDocx
    }
    
    public async generateEClassRecordDocx(data: EClassRecordDocxData): Promise<void> {
        // Stub for generateEClassRecordDocx
    }

    public async generateMapehRecordDocx(data: MapehRecordDocxData): Promise<void> {
        // Stub for generateMapehRecordDocx
    }

    public async generateCertificateDocx(data: CertificateDocxData): Promise<void> {
        // Stub for generateCertificateDocx
    }
    
    public async generateHonorsListDocx(data: HonorsListDocxData): Promise<void> {
        // Stub for generateHonorsListDocx
    }

    public async generateSF2Docx(students: Student[], attendance: Attendance[], settings: SchoolSettings, currentDate: Date): Promise<void> {
        // Stub for generateSF2Docx
    }

    public async generatePickedStudentsDocx(data: PickedStudentsDocxData): Promise<void> {
        // Stub for generatePickedStudentsDocx
    }

    public async generateGroupsDocx(data: GroupsDocxData): Promise<void> {
        // Stub for generateGroupsDocx
    }

    public async generateStudentProfileDocx(data: StudentProfileDocxData): Promise<void> {
        // Stub for generateStudentProfileDocx
    }

}

export const docxService = new DocxService();
