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
  UnderlineType,
  IParagraphOptions,
  PageBreak,
  IRunOptions,
  Numbering,
  Indent,
  IImageOptions,
  VerticalAlign,
  HeightRule,
  TableVerticalAlign,
  LevelFormat,
  TextWrappingType,
  HorizontalPositionAlign,
  VerticalPositionAlign,
  ShadingType,
  ITableCellOptions,
  IBordersOptions,
} from 'docx';
// FIX: Import the 'GeneratedQuizSection' type to resolve a 'Cannot find name' error.
import { Student, SchoolSettings, Attendance, Quarter, SubjectQuarterSettings, StudentQuarterlyRecord, MapehRecordDocxData, GeneratedQuiz, GeneratedQuizQuestion, GeneratedQuizSection, DlpContent, DlpProcedure, QuizType, DllContent, DllObjectives, DllDailyEntry, DllProcedure as DllProcedureType, DlpRubricItem, StudentProfileDocxData, LearningActivitySheet, GeneratedExam, LearningActivitySheetDay } from '../types';
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
        if (!markdownText) return [new Paragraph({ run: { font: "Century Gothic", size: 24 }})];

        const paragraphs: Paragraph[] = [];
        const lines = markdownText.split('\n').filter(line => line.trim() !== '');

        for (const line of lines) {
            const children: TextRun[] = [];
            const parts = line.trim().split(/(\*\*.*?\*\*|\*.*?\*)/g).filter(Boolean);

            for (const part of parts) {
                const fontOptions = { font: "Century Gothic", size: 24 };
                if (part.startsWith('**') && part.endsWith('**')) {
                    children.push(new TextRun({ text: part.slice(2, -2), bold: true, ...fontOptions }));
                } else if (part.startsWith('*') && part.endsWith('*')) {
                    children.push(new TextRun({ text: part.slice(1, -1), italics: true, ...fontOptions }));
                } else {
                    children.push(new TextRun({ text: part, ...fontOptions }));
                }
            }
            
            const isListItem = /^\s*•\s+/.test(line.trim());
            const isSubListItem = /^\s*o\s+/.test(line.trim());

            paragraphs.push(new Paragraph({
                children,
                bullet: isSubListItem ? { level: 1 } : (isListItem ? { level: 0 } : undefined),
                indent: isSubListItem ? { left: 1080, hanging: 360 } : (isListItem ? { left: 720, hanging: 360 } : undefined),
                spacing: { after: 80 }
            }));
        }

        return paragraphs;
    }

    public async generateQuizDocx(quiz: GeneratedQuiz): Promise<void> {
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
                text: quizTitle,
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
                        new TableCell({ children: [new Paragraph(item.objective)] }),
                        new TableCell({ children: [new Paragraph(item.cognitiveLevel)] }),
                        new TableCell({ children: [new Paragraph(item.itemNumbers)] }),
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
                sections.push(new Paragraph({ text: type, heading: HeadingLevel.HEADING_1, spacing: { before: 400, after: 200 } }));
                sections.push(new Paragraph({ text: section.instructions, italics: true, spacing: { after: 200 } }));
                answerKeySections.push(new Paragraph({ text: type, heading: HeadingLevel.HEADING_1, spacing: { before: 400, after: 200 } }));
                
                let questionCounter = 1;
                section.questions.forEach((q, i) => {
                    sections.push(new Paragraph({
                        text: q.questionText,
                        numbering: {
                            reference: "quiz-numbering",
                            level: 0,
                        },
                    }));

                    if (q.options) {
                        q.options.forEach((opt, oi) => {
                            sections.push(new Paragraph({
                                text: `${String.fromCharCode(65 + oi)}. ${opt}`,
                                indentation: { left: 1080 },
                            }));
                        });
                    } else if (type === 'Identification' || type === 'True or False') {
                         sections.push(new Paragraph({ text: "Answer: ____________________", indentation: { left: 1080 } }));
                    }

                    answerKeySections.push(new Paragraph({
                        text: `${questionCounter}. ${q.correctAnswer}`,
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
                sections.push(new Paragraph({ text: activity.activityName, heading: HeadingLevel.HEADING_2, spacing: { before: 200, after: 100 } }));
                sections.push(new Paragraph({ text: activity.activityInstructions, italics: true, spacing: { after: 200 } }));
                if (activity.rubric && activity.rubric.length > 0) {
                    const rubricRows = activity.rubric.map(item => new TableRow({
                        children: [
                            new TableCell({ children: [new Paragraph(item.criteria)] }),
                            new TableCell({ children: [new Paragraph(String(item.points))] }),
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
        this.downloadBlob(blob, `${quizTitle.replace(/\s/g, '_')}_Quiz.docx`);
    }
    
    public async generateDlpDocx(
        dlpForm: any, 
        dlpContent: DlpContent, 
        htmlContent: string, 
        settings: SchoolSettings
    ): Promise<void> {
    
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
        const gradeColor = gradeColorMapping[dlpForm.gradeLevel] || 'D9D9D9';
    
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
                                            new Paragraph({ text: dlpForm.schoolName.toUpperCase(), style: 'header-value' }),
                                            new Paragraph({ text: dlpForm.quarterSelect, style: 'header-value' }),
                                            new Paragraph({ text: `${isFilipino ? 'Guro' : 'Teacher'}: ${dlpForm.teacher}`, style: 'header-value' }),
                                            new Paragraph({ text: `${isFilipino ? 'Asignatura' : 'Learning Area'}: ${dlpForm.subject.toUpperCase()}`, style: 'header-value' }),
                                            new Paragraph({ text: `${isFilipino ? 'Petsa ng Pagtuturo' : 'Teaching Dates'}: ${dlpForm.teachingDates}`, style: 'header-value' }),
                                        ].map(p => new Paragraph({ ...p.options, alignment: AlignmentType.LEFT })),
                                        width: { size: 55, type: WidthType.PERCENTAGE }
                                    }),
                                    new TableCell({
                                        children: [
                                            new Paragraph({ text: isFilipino ? 'DETALYADONG BANGHAY-ARALIN SA' : 'DAILY LESSON PLAN IN', alignment: AlignmentType.CENTER, bold: true }),
                                            new Paragraph({ text: `${dlpForm.subject.toUpperCase()} ${dlpForm.gradeLevel}`, alignment: AlignmentType.CENTER, bold: true }),
                                            new Paragraph({ text: (isFilipino ? 'ISKEDYUL NG KLASE' : 'CLASS SCHEDULE'), alignment: AlignmentType.CENTER, bold: true, spacing: {before: 100} }),
                                            ...dlpForm.classSchedule.split('\n').map((line: string) => new Paragraph({ text: line, alignment: AlignmentType.CENTER }))
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
                            new TableRow({ children: [createHeaderCell(t.objectives), createContentCell([ new Paragraph({ children: [new TextRun({ ...boldTextRunOptions, text: (isFilipino ? 'Pamantayang Pangnilalaman: ' : 'Content Standard: ') }), new TextRun({ ...textRunOptions, text: dlpContent.contentStandard })] }), new Paragraph({ children: [new TextRun({ ...boldTextRunOptions, text: (isFilipino ? 'Pamantayan sa Pagganap: ' : 'Performance Standard: ') }), new TextRun({ ...textRunOptions, text: dlpContent.performanceStandard })] }), new Paragraph({ children: [new TextRun({ ...boldTextRunOptions, text: (isFilipino ? 'Kasanayan sa Pagkatuto: ' : 'Learning Competency: ') }), new TextRun({ ...textRunOptions, text: dlpForm.learningCompetency })] }), new Paragraph({ text: (isFilipino ? 'Sa pagtatapos ng aralin, ang mga mag-aaral ay inaasahang:' : 'At the end of the lesson, the learners should be able to:'), spacing: { before: 200 } }), new Paragraph({ text: dlpForm.lessonObjective, bullet: { level: 0 } })])] }),
                            new TableRow({ children: [createHeaderCell(t.content), createContentCell([ new Paragraph({ children: [new TextRun({ ...boldTextRunOptions, text: (isFilipino ? 'Paksa: ' : 'Topic: ') }), new TextRun({ ...textRunOptions, text: dlpContent.topic })] })])] }),
                            new TableRow({ children: [createHeaderCell(t.resources), createContentCell([ new Paragraph({ children: [new TextRun({ ...boldTextRunOptions, text: (isFilipino ? 'Sanggunian: ' : 'References: ') }), new TextRun({ ...textRunOptions, text: dlpContent.learningReferences })] }), new Paragraph({ children: [new TextRun({ ...boldTextRunOptions, text: (isFilipino ? 'Kagamitan: ' : 'Materials: ') }), new TextRun({ ...textRunOptions, text: dlpContent.learningMaterials })] })])] }),
                            new TableRow({ children: [createHeaderCell(t.procedure), new TableCell({
                                children: [ new Table({
                                    width: { size: 100, type: WidthType.PERCENTAGE },
                                    columnWidths: [25, 45, 30],
                                    rows: [
                                        new TableRow({ children: [ new TableCell({ children: [new Paragraph({ text: (isFilipino ? 'Pamamaraan' : 'Procedure'), bold: true })] }), new TableCell({ children: [new Paragraph({ text: (isFilipino ? 'Gawain ng Guro/Mag-aaral' : 'Teacher/Student Activity'), bold: true })] }), new TableCell({ children: [new Paragraph({ text: (isFilipino ? 'Mga Kaugnay na PPST Indicator' : 'Aligned PPST Indicators'), bold: true })] })] }),
                                        ...dlpContent.procedures.map(proc => new TableRow({
                                            children: [ new TableCell({ children: [new Paragraph({ text: proc.title, bold: true })] }), new TableCell({ children: this.parseMarkdownToParagraphs(proc.content) }), new TableCell({ children: [new Paragraph({ text: proc.ppst, italics: true })] })],
                                        }))
                                    ],
                                })],
                            })] }),
                            new TableRow({ children: [createHeaderCell(t.evaluation), createContentCell( (dlpContent.evaluationQuestions || []).map(q => new Paragraph({ text: q.question, numbering: { reference: "dlp-list", level: 0 }, spacing: { after: 100 } })))] }),
                            new TableRow({ children: [createHeaderCell(t.reflection), createContentCell([new Paragraph({text: ""})])] }), // Empty for now, as per user's last request.
                        ],
                    }),
    
                    new PageBreak(),
    
                    // Answer Key
                    new Paragraph({ text: (isFilipino ? 'Susi sa Pagwawasto' : 'Answer Key (For Evaluating Learning)'), heading: HeadingLevel.HEADING_2 }),
                    ...(dlpContent.evaluationQuestions || []).map(q => new Paragraph({ text: q.answer, numbering: { reference: "dlp-list", level: 0 }, spacing: { after: 100 } })),
                ],
            }]
        });
    
        const blob = await Packer.toBlob(doc);
        this.downloadBlob(blob, `DLP_${dlpForm.subject.replace(/\s/g, '_')}.docx`);
    }
    
    public async generateDllDocx(
        dllForm: any, 
        dllContent: DllContent,
        settings: SchoolSettings
    ): Promise<void> {

        // Corrected page size for 8.5" x 13" (Long Bond Paper) in landscape
        // The library expects portrait dimensions first, then applies orientation.
        // Width: 8.5 inches * 1440 DXA/inch = 12240 DXA
        // Height: 13 inches * 1440 DXA/inch = 18720 DXA
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
                        new TableCell({ children: [new Paragraph({ text: `School: ${settings.schoolName}`, alignment: AlignmentType.LEFT })] }),
                        new TableCell({ children: [new Paragraph({ text: `Grade Level: ${dllForm.gradeLevel}`, alignment: AlignmentType.LEFT })] }),
                        new TableCell({ children: [new Paragraph({ text: `Teacher: ${settings.teacherName}`, alignment: AlignmentType.LEFT })] }),
                        new TableCell({ children: [new Paragraph({ text: `Learning Area: ${dllForm.subject}`, alignment: AlignmentType.LEFT })] }),
                    ]
                }),
                new TableRow({
                    children: [
                        new TableCell({ children: [new Paragraph({ text: `Teaching Dates & Time: ${dllForm.teachingDates}`, alignment: AlignmentType.LEFT })] }),
                        new TableCell({ children: [new Paragraph({ text: `Quarter: ${dllForm.quarter}`, alignment: AlignmentType.LEFT })] }),
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
        mainTableRows.push(new TableRow({ children: [new TableCell({ children: [new Paragraph("A. Content Standard")] }), new TableCell({ children: [new Paragraph(dllContent.contentStandard)], columnSpan: 5 })]}));
        mainTableRows.push(new TableRow({ children: [new TableCell({ children: [new Paragraph("B. Performance Standard")] }), new TableCell({ children: [new Paragraph(dllContent.performanceStandard)], columnSpan: 5 })]}));
        mainTableRows.push(new TableRow({
            children: [
                new TableCell({ children: [new Paragraph("C. Learning Competencies")] }),
                new TableCell({ children: [new Paragraph(dllContent.learningCompetencies.monday)] }),
                new TableCell({ children: [new Paragraph(dllContent.learningCompetencies.tuesday)] }),
                new TableCell({ children: [new Paragraph(dllContent.learningCompetencies.wednesday)] }),
                new TableCell({ children: [new Paragraph(dllContent.learningCompetencies.thursday)] }),
                new TableCell({ children: [new Paragraph(dllContent.learningCompetencies.friday)] }),
            ]
        }));
        
        // Add Content
        mainTableRows.push(new TableRow({ children: [new TableCell({ children: [new Paragraph({ text: "II. CONTENT", bold: true })] }), new TableCell({ children: [new Paragraph(dllContent.content)], columnSpan: 5 })]}));

        // Add Learning Resources
        mainTableRows.push(new TableRow({ children: [new TableCell({ children: [new Paragraph({ text: "III. LEARNING RESOURCES", bold: true })], columnSpan: 6 })]}));
        // ... (similar rows for each resource type)

        // Add Procedures
        mainTableRows.push(new TableRow({ children: [new TableCell({ children: [new Paragraph({ text: "IV. PROCEDURES", bold: true })], columnSpan: 6 })]}));
        dllContent.procedures.forEach(proc => {
            mainTableRows.push(new TableRow({
                children: [
                    new TableCell({ children: [new Paragraph(proc.procedure)] }),
                    new TableCell({ children: this.parseMarkdownToParagraphs(proc.monday) }),
                    new TableCell({ children: this.parseMarkdownToParagraphs(proc.tuesday) }),
                    new TableCell({ children: this.parseMarkdownToParagraphs(proc.wednesday) }),
                    new TableCell({ children: this.parseMarkdownToParagraphs(proc.thursday) }),
                    new TableCell({ children: this.parseMarkdownToParagraphs(proc.friday) }),
                ]
            }));
        });
        
        // Add Remarks & Reflection
        mainTableRows.push(new TableRow({ children: [new TableCell({ children: [new Paragraph({ text: "V. REMARKS", bold: true })] }), new TableCell({ children: [new Paragraph(dllContent.remarks)], columnSpan: 5 })]}));
        mainTableRows.push(new TableRow({ children: [new TableCell({ children: [new Paragraph({ text: "VI. REFLECTION", bold: true })], columnSpan: 6 })]}));
        // ... (rows for reflection)

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
                        new TableCell({ children: [new Paragraph({ text: dllForm.preparedByName, ...boldCentered })], borders: { top: { style: BorderStyle.SINGLE }, bottom: { style: BorderStyle.NONE }, left: { style: BorderStyle.NONE }, right: { style: BorderStyle.NONE }} }),
                        new TableCell({ children: [new Paragraph({ text: dllForm.checkedByName, ...boldCentered })], borders: { top: { style: BorderStyle.SINGLE }, bottom: { style: BorderStyle.NONE }, left: { style: BorderStyle.NONE }, right: { style: BorderStyle.NONE }} }),
                        new TableCell({ children: [new Paragraph({ text: dllForm.approvedByName, ...boldCentered })], borders: { top: { style: BorderStyle.SINGLE }, bottom: { style: BorderStyle.NONE }, left: { style: BorderStyle.NONE }, right: { style: BorderStyle.NONE }} }),
                    ]
                 }),
                 new TableRow({
                    children: [
                        new TableCell({ children: [new Paragraph({ text: dllForm.preparedByDesignation, alignment: AlignmentType.CENTER })], borders: { top: { style: BorderStyle.NONE }, bottom: { style: BorderStyle.NONE }, left: { style: BorderStyle.NONE }, right: { style: BorderStyle.NONE }} }),
                        new TableCell({ children: [new Paragraph({ text: dllForm.checkedByDesignation, alignment: AlignmentType.CENTER })], borders: { top: { style: BorderStyle.NONE }, bottom: { style: BorderStyle.NONE }, left: { style: BorderStyle.NONE }, right: { style: BorderStyle.NONE }} }),
                        new TableCell({ children: [new Paragraph({ text: dllForm.approvedByDesignation, alignment: AlignmentType.CENTER })], borders: { top: { style: BorderStyle.NONE }, bottom: { style: BorderStyle.NONE }, left: { style: BorderStyle.NONE }, right: { style: BorderStyle.NONE }} }),
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
        this.downloadBlob(blob, `DLL_${dllForm.subject.replace(/\s/g, '_')}.docx`);
    }

    public async generateLasDocx(
        lasForm: any,
        lasContent: LearningActivitySheet,
        settings: SchoolSettings
    ): Promise<void> {

        const sections: (Paragraph | Table | PageBreak)[] = [];
        
        lasContent.days.forEach((dayData, index) => {
            if (index > 0) {
                sections.push(new PageBreak());
            }

            // Day Title
            sections.push(new Paragraph({ text: dayData.dayTitle, heading: HeadingLevel.HEADING_1, run: { font: "Century Gothic", size: 28, bold: true }, spacing: { before: 200, after: 100 } }));

            // DepEd Header
            sections.push(new Paragraph({
                text: "DepED | Dynamic Learning Program | BAGONG PILIPINAS | LEARNING ACTIVITY SHEET",
                alignment: AlignmentType.CENTER,
                run: { font: "Century Gothic", size: 20, bold: true },
                spacing: { after: 200 },
                border: { bottom: { color: "auto", space: 1, value: "single", size: 6 } },
            }));

            // Info Section
            sections.push(new Paragraph({ text: `Subject: ${lasForm.subject}`, run: { font: "Century Gothic", size: 24 }, spacing: { top: 100, after: 50 } }));
            sections.push(new Paragraph({ text: "Grade & Section: ____________________", run: { font: "Century Gothic", size: 24 }, spacing: { after: 100 } }));
            sections.push(new Paragraph({
                children: [
                    new TextRun({ text: "Name: ____________________", font: "Century Gothic", size: 24 }),
                    new TextRun({ text: "\t\t", font: "Century Gothic", size: 24 }),
                    new TextRun({ text: "Score: __________", font: "Century Gothic", size: 24 }),
                    new TextRun({ text: "\t\t", font: "Century Gothic", size: 24 }),
                    new TextRun({ text: "Date: __________", font: "Century Gothic", size: 24 }),
                ],
                spacing: { after: 100 }
            }));
            sections.push(new Paragraph({ text: `Activity Title: ${dayData.activityTitle}`, run: { font: "Century Gothic", size: 24 }, spacing: { after: 50 } }));
            sections.push(new Paragraph({ text: `Learning Target: ${dayData.learningTarget}`, run: { font: "Century Gothic", size: 24 }, spacing: { after: 100 } }));
            sections.push(new Paragraph({ text: `Reference: ${dayData.references || '____________________'}`, run: { font: "Century Gothic", size: 24 }, spacing: { after: 200 } }));

            // Main content logic
            dayData.conceptNotes.forEach(note => {
                sections.push(new Paragraph({ text: note.title, bold: true, run: { font: "Century Gothic", size: 28 }, spacing: { after: 100 } }));
                sections.push(...this.parseLasMarkdown(note.content));
            });
            
            dayData.activities.forEach(activity => {
                sections.push(new Paragraph({ text: activity.title, bold: true, run: { font: "Century Gothic", size: 28 }, spacing: { before: 200, after: 100 } }));
                
                const isMatchingType = activity.title.toLowerCase().includes('match');
                const instructionsLines = activity.instructions.split('\n');
                const regularInstructions = instructionsLines.filter(line => !line.includes('||')).join('\n');
                
                if (regularInstructions.trim()) {
                    sections.push(new Paragraph({ text: `Directions: ${regularInstructions}`, italics: true, run: { font: "Century Gothic", size: 24 }, spacing: { after: 100 }}));
                }

                const tableLines = instructionsLines.filter(line => line.includes('||'));

                if (isMatchingType && tableLines.length > 0) {
                    const tableRows = tableLines.map(line => {
                        const parts = line.split('||');
                        return new TableRow({
                            children: [
                                new TableCell({ children: this.parseLasMarkdown(parts[0] || ''), width: { size: 50, type: WidthType.PERCENTAGE } }),
                                new TableCell({ children: this.parseLasMarkdown(parts[1] || ''), width: { size: 50, type: WidthType.PERCENTAGE } }),
                            ],
                        });
                    });
                    
                    tableRows.unshift(new TableRow({
                        children: [
                            new TableCell({ children: [new Paragraph({ text: "Column A (Situation)", bold: true, alignment: AlignmentType.CENTER, run: { font: "Century Gothic", size: 24 } })] }),
                            new TableCell({ children: [new Paragraph({ text: "Column B (Term)", bold: true, alignment: AlignmentType.CENTER, run: { font: "Century Gothic", size: 24 } })] }),
                        ],
                        tableHeader: true,
                    }));

                    sections.push(new Table({ rows: tableRows, width: { size: 100, type: WidthType.PERCENTAGE } }));
                } else if (activity.instructions.trim()) {
                    sections.push(...this.parseLasMarkdown(activity.instructions));
                }

                if (activity.questions) {
                    activity.questions.forEach(q => {
                        sections.push(new Paragraph({ text: q.questionText, numbering: { reference: "las-list", level: 0 }, run: { font: "Century Gothic", size: 24 } }));
                    });
                }
            });

            sections.push(new Paragraph({ text: "REFLECTION", bold: true, run: { font: "Century Gothic", size: 28 }, spacing: { before: 200, after: 100 } }));
            sections.push(new Paragraph({ text: dayData.reflection, run: { font: "Century Gothic", size: 24 }}));
            sections.push(new Paragraph({ text: "__________________________________________________________________________________________", run: { font: "Century Gothic", size: 24 }, spacing: { before: 100 }}));
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
                            run: { font: "Century Gothic", size: 24 }
                        }],
                    },
                ],
            },
            sections: [{
                properties: {
                    page: {
                        size: { width: 11906, height: 16838 }, // A4 Portrait
                        margin: { top: 720, right: 720, bottom: 720, left: 720 },
                    },
                },
                children: sections
            }]
        });

        const blob = await Packer.toBlob(doc);
        this.downloadBlob(blob, `LAS_${lasForm.subject.replace(/\s/g, '_')}.docx`);
    }

    public async generateExamDocx(exam: GeneratedExam, settings: SchoolSettings): Promise<void> {
        // Docx generation logic...
    }

    public async generateAttendanceDocx(students: Student[], attendance: Attendance[], currentDate: Date, schoolSettings: SchoolSettings): Promise<void> {
        // Docx generation logic...
    }
    
    public async generateSummaryOfGradesDocx(data: SummaryOfGradesDocxData): Promise<void> {
       // Docx generation logic...
    }
    
    public async generateEClassRecordDocx(data: EClassRecordDocxData): Promise<void> {
        // Docx generation logic...
    }

    public async generateMapehRecordDocx(data: MapehRecordDocxData): Promise<void> {
        // Docx generation logic...
    }

    public async generateCertificateDocx(data: CertificateDocxData): Promise<void> {
        // Docx generation logic...
    }
    
    public async generateHonorsListDocx(data: HonorsListDocxData): Promise<void> {
        // Docx generation logic...
    }

    public async generateSF2Docx(students: Student[], attendance: Attendance[], settings: SchoolSettings, currentDate: Date): Promise<void> {
        // Docx generation logic...
    }

    public async generatePickedStudentsDocx(data: PickedStudentsDocxData): Promise<void> {
        // Docx generation logic...
    }

    public async generateGroupsDocx(data: GroupsDocxData): Promise<void> {
        // Docx generation logic...
    }

    public async generateStudentProfileDocx(data: StudentProfileDocxData): Promise<void> {
        // Docx generation logic...
    }

}

export const docxService = new DocxService();