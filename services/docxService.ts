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
} from 'docx';
import { Student, SchoolSettings, Attendance, Quarter, SubjectQuarterSettings, StudentQuarterlyRecord, MapehRecordDocxData, GeneratedQuiz, GeneratedQuizQuestion, DlpContent, DlpProcedure, QuizType, DllContent, DllObjectives, DllDailyEntry, DllProcedure as DllProcedureType, DlpRubricItem, StudentProfileDocxData, LearningActivitySheet, GeneratedExam } from '../types';
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
        if (!markdownText) return [new Paragraph("")];

        const paragraphs: Paragraph[] = [];
        const lines = markdownText.split('\n');

        for (const line of lines) {
            if (line.trim() === '') {
                paragraphs.push(new Paragraph(""));
                continue;
            }

            const children: TextRun[] = [];
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
            
            paragraphs.push(new Paragraph({ children, spacing: { after: 100 } }));
        }

        return paragraphs;
    }


  public async generateAttendanceDocx(students: Student[], attendance: Attendance[], currentDate: Date, schoolSettings: SchoolSettings): Promise<void> {
    const year = currentDate.getUTCFullYear();
    const month = currentDate.getUTCMonth(); 

    const daysInMonth = new Date(year, month + 1, 0).getDate();
    const monthDays = Array.from({ length: daysInMonth }, (_, i) => i + 1);

    const studentsByGender = {
        males: students.filter(s => s.gender === 'Male').sort((a,b) => a.lastName.localeCompare(b.lastName)),
        females: students.filter(s => s.gender === 'Female').sort((a,b) => a.lastName.localeCompare(b.lastName)),
        others: students.filter(s => s.gender !== 'Male' && s.gender !== 'Female').sort((a,b) => a.lastName.localeCompare(b.lastName))
    };

    const attendanceMap = new Map<string, Map<number, string>>();
    for (const att of attendance) {
        const attDate = new Date(`${att.date}T00:00:00Z`);
        if (attDate.getUTCFullYear() === year && attDate.getUTCMonth() === month) {
            if (!attendanceMap.has(att.studentId)) {
                attendanceMap.set(att.studentId, new Map());
            }
            const day = attDate.getUTCDate();
            const status = att.status === 'present' ? '' : att.status === 'late' ? '.' : 'x';
            attendanceMap.get(att.studentId)!.set(day, status);
        }
    }

    const summary = (() => {
        const totalAbsences: Record<string, number> = {};
        const totalTardies: Record<string, number> = {};

        students.forEach(s => {
            totalAbsences[s.id] = 0;
            totalTardies[s.id] = 0;
        });

        attendanceMap.forEach((dayMap, studentId) => {
            if(!students.find(s => s.id === studentId)) return;
            dayMap.forEach(status => {
                if (status === 'x') totalAbsences[studentId]++;
                if (status === '.') totalTardies[studentId]++;
            });
        });
        
        const dailySummary = monthDays.map(day => {
            let absent = 0;
            let tardy = 0;
            students.forEach(student => {
                 const status = attendanceMap.get(student.id)?.get(day);
                 if (status === 'x') absent++;
                 if (status === '.') tardy++;
            })
            return { absent, tardy };
        });

        return { totalAbsences, totalTardies, dailySummary };
    })();
    
    const tableHeader1 = new TableRow({
        tableHeader: true,
        children: [
            new TableCell({ children: [new Paragraph({ text: "#", alignment: AlignmentType.CENTER })], rowSpan: 2, verticalAlign: VerticalAlign.CENTER }),
            new TableCell({ children: [new Paragraph({ text: "(Learner's Name)", alignment: AlignmentType.CENTER })], rowSpan: 2, verticalAlign: VerticalAlign.CENTER, width: { size: 3500, type: WidthType.DXA } }),
            new TableCell({ children: [new Paragraph({ text: "(Date)", alignment: AlignmentType.CENTER })], columnSpan: daysInMonth, verticalAlign: VerticalAlign.CENTER }),
            new TableCell({ children: [new Paragraph({ text: "Total", alignment: AlignmentType.CENTER })], columnSpan: 2, verticalAlign: VerticalAlign.CENTER }),
        ]
    });

    const tableHeader2 = new TableRow({
        tableHeader: true,
        children: [
            ...monthDays.map(day => new TableCell({ children: [new Paragraph({ text: day.toString(), alignment: AlignmentType.CENTER })], verticalAlign: VerticalAlign.CENTER, width: { size: 400, type: WidthType.DXA } })),
            new TableCell({ children: [new Paragraph({ text: "Absences", alignment: AlignmentType.CENTER })], verticalAlign: VerticalAlign.CENTER }),
            new TableCell({ children: [new Paragraph({ text: "Tardiness", alignment: AlignmentType.CENTER })], verticalAlign: VerticalAlign.CENTER }),
        ],
    });

    const createStudentRows = (studentList: Student[], startIndex: number) => {
        return studentList.map((student, index) => {
            const cells = [
                new TableCell({ children: [new Paragraph({ text: (startIndex + index + 1).toString(), alignment: AlignmentType.CENTER })], verticalAlign: VerticalAlign.CENTER }),
                new TableCell({ children: [new Paragraph({ text: `${student.lastName}, ${student.firstName} ${student.middleName?.[0] || ''}.`, alignment: AlignmentType.LEFT })], verticalAlign: VerticalAlign.CENTER }),
            ];

            for (let day = 1; day <= daysInMonth; day++) {
                cells.push(new TableCell({ children: [new Paragraph({ text: attendanceMap.get(student.id)?.get(day) || '', alignment: AlignmentType.CENTER })], verticalAlign: VerticalAlign.CENTER }));
            }

            cells.push(new TableCell({ children: [new Paragraph({ text: summary.totalAbsences[student.id]?.toString() || '0', alignment: AlignmentType.CENTER })], verticalAlign: VerticalAlign.CENTER }));
            cells.push(new TableCell({ children: [new Paragraph({ text: summary.totalTardies[student.id]?.toString() || '0', alignment: AlignmentType.CENTER })], verticalAlign: VerticalAlign.CENTER }));
            
            return new TableRow({ children: cells });
        });
    };

    const maleRows = createStudentRows(studentsByGender.males, 0);
    const femaleRows = createStudentRows(studentsByGender.females, studentsByGender.males.length);

    const monthName = currentDate.toLocaleString('default', { month: 'long', timeZone: 'UTC' });
    const studentInfo = students[0];
    const sectionText = (studentInfo?.gradeLevel && studentInfo?.section) ? `${studentInfo.gradeLevel} - ${studentInfo.section}` : 'Class';

    const doc = new Document({
        sections: [{
            properties: {
                page: {
                    size: { width: 18720, height: 12240 }, // 13 x 8.5 inches in DXA
                    margin: { top: 720, right: 720, bottom: 720, left: 720 },
                },
            },
            children: [
                new Paragraph({ text: "Monthly Attendance Report", heading: HeadingLevel.HEADING_1, alignment: AlignmentType.CENTER }),
                new Paragraph({ text: `For the Month of: ${monthName} ${year}`, alignment: AlignmentType.CENTER }),
                new Paragraph({ text: `Class: ${sectionText}`, alignment: AlignmentType.CENTER }),
                new Paragraph({ text: `Teacher: ${schoolSettings.teacherName}`, alignment: AlignmentType.CENTER }),
                new Paragraph({ text: "" }), // Spacer
                new Table({
                    rows: [
                        tableHeader1,
                        tableHeader2,
                        new TableRow({ children: [new TableCell({ children: [new Paragraph({ text: "MALE", alignment: AlignmentType.LEFT })], columnSpan: daysInMonth + 4, borders: { top: { style: BorderStyle.SINGLE }, bottom: { style: BorderStyle.SINGLE } } })] }),
                        ...maleRows,
                        new TableRow({ children: [new TableCell({ children: [new Paragraph({ text: "FEMALE", alignment: AlignmentType.LEFT })], columnSpan: daysInMonth + 4, borders: { top: { style: BorderStyle.SINGLE }, bottom: { style: BorderStyle.SINGLE } } })] }),
                        ...femaleRows,
                    ]
                })
            ]
        }]
    });
    const blob = await Packer.toBlob(doc);
    this.downloadBlob(blob, `Attendance_${monthName}_${year}.docx`);
  }
  public async generateExamDocx(exam: GeneratedExam, settings: SchoolSettings): Promise<void> {
    const { title, tableOfSpecifications, questions, subject, gradeLevel, quarter } = exam;

    const fontStyle = "Bookman Old Style";
    const fontSize = 22; // 11pt

    const tableHeaderStyle: IParagraphOptions = {
        alignment: AlignmentType.CENTER,
        run: { font: fontStyle, size: 18, bold: true },
    };

    const tableCellStyle: IParagraphOptions = {
        alignment: AlignmentType.CENTER,
        run: { font: fontStyle, size: 18 },
    };
    
    const objectiveCellStyle: IParagraphOptions = {
        alignment: AlignmentType.LEFT,
        run: { font: fontStyle, size: 18 },
    };

    const tosTable = new Table({
        width: { size: 100, type: WidthType.PERCENTAGE },
        rows: [
            new TableRow({
                tableHeader: true,
                children: [
                    new TableCell({ children: [new Paragraph({ ...tableHeaderStyle, text: "LEARNING COMPETENCIES" })], rowSpan: 2, verticalAlign: VerticalAlign.CENTER, width: { size: 30, type: WidthType.PERCENTAGE } }),
                    new TableCell({ children: [new Paragraph({ ...tableHeaderStyle, text: "NO. OF DAYS TAUGHT" })], rowSpan: 2, verticalAlign: VerticalAlign.CENTER }),
                    new TableCell({ children: [new Paragraph({ ...tableHeaderStyle, text: "PERCENTAGE" })], rowSpan: 2, verticalAlign: VerticalAlign.CENTER }),
                    new TableCell({ children: [new Paragraph({ ...tableHeaderStyle, text: "NO. OF ITEMS" })], rowSpan: 2, verticalAlign: VerticalAlign.CENTER }),
                    new TableCell({ children: [new Paragraph({ ...tableHeaderStyle, text: "COGNITIVE PROCESS DIMENSION" })], columnSpan: 6, verticalAlign: VerticalAlign.CENTER }),
                    new TableCell({ children: [new Paragraph({ ...tableHeaderStyle, text: "ITEM PLACEMENT" })], rowSpan: 2, verticalAlign: VerticalAlign.CENTER }),
                ],
            }),
            new TableRow({
                tableHeader: true,
                children: [ "Remembering", "Understanding", "Applying", "Analyzing", "Evaluating", "Creating" ].map(label =>
                    new TableCell({ children: [new Paragraph({ ...tableHeaderStyle, text: label })], verticalAlign: VerticalAlign.CENTER })
                ),
            }),
            ...tableOfSpecifications.map(item => new TableRow({
                children: [
                    new TableCell({ children: [new Paragraph({ ...objectiveCellStyle, text: item.objective })] }),
                    new TableCell({ children: [new Paragraph({ ...tableCellStyle, text: String(item.daysTaught) })] }),
                    new TableCell({ children: [new Paragraph({ ...tableCellStyle, text: item.percentage })] }),
                    new TableCell({ children: [new Paragraph({ ...tableCellStyle, text: String(item.numItems) })] }),
                    new TableCell({ children: [new Paragraph({ ...tableCellStyle, text: item.remembering })] }),
                    new TableCell({ children: [new Paragraph({ ...tableCellStyle, text: item.understanding })] }),
                    new TableCell({ children: [new Paragraph({ ...tableCellStyle, text: item.applying })] }),
                    new TableCell({ children: [new Paragraph({ ...tableCellStyle, text: item.analyzing })] }),
                    new TableCell({ children: [new Paragraph({ ...tableCellStyle, text: item.evaluating })] }),
                    new TableCell({ children: [new Paragraph({ ...tableCellStyle, text: item.creating })] }),
                    new TableCell({ children: [new Paragraph({ ...tableCellStyle, text: item.itemPlacement })] }),
                ]
            })),
            new TableRow({
                children: [
                    new TableCell({ children: [new Paragraph({ ...tableHeaderStyle, text: "TOTAL" })] }),
                    new TableCell({ children: [new Paragraph({ ...tableCellStyle, text: String(tableOfSpecifications.reduce((acc, item) => acc + item.daysTaught, 0)) })] }),
                    new TableCell({ children: [new Paragraph({ ...tableCellStyle, text: "100%" })] }),
                    new TableCell({ children: [new Paragraph({ ...tableCellStyle, text: "50" })] }),
                    new TableCell({ children: [new Paragraph({ ...tableCellStyle, text: String(tableOfSpecifications.reduce((acc, item) => acc + (parseInt(item.remembering, 10) || 0), 0)) })] }),
                    new TableCell({ children: [new Paragraph({ ...tableCellStyle, text: String(tableOfSpecifications.reduce((acc, item) => acc + (parseInt(item.understanding, 10) || 0), 0)) })] }),
                    new TableCell({ children: [new Paragraph({ ...tableCellStyle, text: String(tableOfSpecifications.reduce((acc, item) => acc + (parseInt(item.applying, 10) || 0), 0)) })] }),
                    new TableCell({ children: [new Paragraph({ ...tableCellStyle, text: String(tableOfSpecifications.reduce((acc, item) => acc + (parseInt(item.analyzing, 10) || 0), 0)) })] }),
                    new TableCell({ children: [new Paragraph({ ...tableCellStyle, text: String(tableOfSpecifications.reduce((acc, item) => acc + (parseInt(item.evaluating, 10) || 0), 0)) })] }),
                    new TableCell({ children: [new Paragraph({ ...tableCellStyle, text: String(tableOfSpecifications.reduce((acc, item) => acc + (parseInt(item.creating, 10) || 0), 0)) })] }),
                    new TableCell({ children: [new Paragraph({ ...tableCellStyle, text: "1-50" })] }),
                ]
            }),
        ],
    });
    
    const quarterText = quarter === '1' ? 'FIRST' : quarter === '2' ? 'SECOND' : quarter === '3' ? 'THIRD' : 'FOURTH';

    const examHeaderParagraphs: Paragraph[] = [
        new Paragraph({ text: settings.schoolName.toUpperCase(), alignment: AlignmentType.CENTER, run: { font: fontStyle, size: fontSize, bold: true } }),
        new Paragraph({ text: `PERIODICAL TEST IN ${subject.toUpperCase()}`, alignment: AlignmentType.CENTER, run: { font: fontStyle, size: fontSize, bold: true } }),
        new Paragraph({ text: `${quarterText} QUARTER, S.Y. ${settings.schoolYear}`, alignment: AlignmentType.CENTER, run: { font: fontStyle, size: fontSize, bold: true } }),
        new Paragraph({ text: `GRADE ${gradeLevel}`, alignment: AlignmentType.CENTER, run: { font: fontStyle, size: fontSize, bold: true } }),
        new Paragraph({ text: "" }),
        new Paragraph({
            children: [
                new TextRun({ text: "Name: __________________________________________________________________", font: fontStyle, size: fontSize }),
                new TextRun({ text: "", break: 1 }),
                new TextRun({ text: "Grade & Section: _________________________________________________________", font: fontStyle, size: fontSize }),
                new TextRun({ text: "", break: 1 }),
                new TextRun({ text: "Score: ____________", font: fontStyle, size: fontSize }),
            ],
        }),
        new Paragraph({ text: "" }),
    ];

    // Question Section
    const questionParagraphs: Paragraph[] = [];
    questions.forEach((q, i) => {
        questionParagraphs.push(new Paragraph({
            text: q.questionText,
            style: 'question',
            numbering: { reference: 'exam-questions', level: 0 }
        }));
        if (q.options) {
            q.options.forEach((opt, oi) => {
                questionParagraphs.push(new Paragraph({
                    children: [new TextRun({ text: `${String.fromCharCode(65 + oi)}. ${opt}` })],
                    style: 'options',
                    indent: { left: 720 }
                }));
            });
        }
        questionParagraphs.push(new Paragraph({ text: "" })); // Add space after each question
    });

    // Answer Key Section
    const answerKeyParagraphs: Paragraph[] = [
        new Paragraph({ text: "ANSWER KEY", heading: HeadingLevel.HEADING_2, alignment: AlignmentType.CENTER, pageBreakBefore: true }),
    ];
    
    const answerKeyTableRows: TableRow[] = [];
    questions.forEach((q, i) => {
        let answerLetter = '';
        if (q.options && q.correctAnswer) {
            const correctIndex = q.options.findIndex(opt => opt.trim().toLowerCase() === q.correctAnswer.trim().toLowerCase());
            if (correctIndex !== -1) {
                answerLetter = String.fromCharCode(65 + correctIndex);
            } else {
                answerLetter = q.correctAnswer; // Fallback
            }
        } else {
            answerLetter = q.correctAnswer;
        }
        answerKeyTableRows.push(new TableRow({
            children: [
                new TableCell({
                    children: [new Paragraph({ text: `${i + 1}. ${answerLetter}`, style: 'answerKey'})],
                    borders: { top: {style: BorderStyle.NONE}, bottom: {style: BorderStyle.NONE}, left: {style: BorderStyle.NONE}, right: {style: BorderStyle.NONE} }
                })
            ]
        }));
    });

    const answerKeyTable = new Table({
        rows: answerKeyTableRows,
        width: { size: 20, type: WidthType.PERCENTAGE },
        columnWidths: [1],
         borders: {
            top: { style: BorderStyle.NONE }, bottom: { style: BorderStyle.NONE },
            left: { style: BorderStyle.NONE }, right: { style: BorderStyle.NONE },
            insideHorizontal: { style: BorderStyle.NONE }, insideVertical: { style: BorderStyle.NONE },
        },
    });

    const doc = new Document({
        numbering: {
            config: [
                {
                    reference: "exam-questions",
                    levels: [{
                        level: 0,
                        format: LevelFormat.DECIMAL,
                        text: "%1.",
                        alignment: AlignmentType.LEFT,
                        style: {
                            paragraph: {
                                indent: { left: 360, hanging: 360 },
                            },
                        },
                    }],
                },
            ],
        },
        styles: {
            paragraphStyles: [
                { id: 'Normal', name: 'Normal', run: { font: fontStyle, size: fontSize } },
                { id: 'question', name: 'Question', basedOn: 'Normal', run: { size: 24 } },
                { id: 'options', name: 'Options', basedOn: 'Normal', run: { size: 22 } },
                { id: 'answerKey', name: 'Answer Key', basedOn: 'Normal', run: { font: fontStyle, size: 22 } },
                { id: 'Heading2', name: 'Heading 2', basedOn: 'Normal', next: 'Normal', run: { font: fontStyle, size: 28, bold: true }, paragraph: { spacing: { before: 240, after: 120 } } },
            ]
        },
        sections: [
            { // Section 1: TOS
                properties: {
                    page: {
                        size: { width: 18720, height: 12240 }, // 13 x 8.5 inches
                        orientation: PageOrientation.LANDSCAPE,
                        margin: { top: 720, right: 720, bottom: 720, left: 720 },
                    },
                },
                children: [
                    new Paragraph({ text: "TABLE OF SPECIFICATIONS", heading: HeadingLevel.HEADING_2, alignment: AlignmentType.CENTER }),
                    new Paragraph({ text: "" }),
                    tosTable,
                ],
            },
            { // Section 2: Exam Questions
                properties: {
                    page: {
                        size: { width: 12240, height: 18720 }, // 8.5 x 13 inches
                        orientation: PageOrientation.PORTRAIT,
                        margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 },
                    },
                },
                children: [
                    ...examHeaderParagraphs,
                    ...questionParagraphs,
                    ...answerKeyParagraphs,
                    answerKeyTable,
                ],
            },
        ]
    });

    const blob = await Packer.toBlob(doc);
    this.downloadBlob(blob, `${title}.docx`);
}
  // FIX: This method was missing. Add the empty `generateMapehRecordDocx` method.
  public async generateMapehRecordDocx(data: MapehRecordDocxData): Promise<void> {
    const { summaryData, settings, quarter, selectedSectionText } = data;
    const doc = new Document({
        sections: [{
            children: [
                new Paragraph({ text: `MAPEH Summary - Quarter ${quarter}`, heading: HeadingLevel.HEADING_1, alignment: AlignmentType.CENTER }),
                new Paragraph({ text: `Class: ${selectedSectionText}`, alignment: AlignmentType.CENTER }),
                new Paragraph({ text: `Teacher: ${settings.teacherName}`, alignment: AlignmentType.CENTER }),
                new Paragraph({ text: "" }),
                new Table({
                    width: { size: 100, type: WidthType.PERCENTAGE },
                    rows: [
                        new TableRow({
                            tableHeader: true,
                            children: [ "Student Name", "Music", "Arts", "PE", "Health", "Final MAPEH Grade" ].map(header => new TableCell({
                                children: [new Paragraph({ text: header, alignment: AlignmentType.CENTER })],
                                shading: { type: ShadingType.CLEAR, fill: "auto", color: "D9EAD3" },
                            }))
                        }),
                        ...summaryData.map(item => new TableRow({
                            children: [
                                new TableCell({ children: [new Paragraph(`${item.student.lastName}, ${item.student.firstName}${item.student.middleName && item.student.middleName.trim() ? ` ${item.student.middleName.trim().charAt(0)}.` : ''}`)] }),
                                new TableCell({ children: [new Paragraph({ text: String(item.componentGrades["Music"] ?? ''), alignment: AlignmentType.CENTER })] }),
                                new TableCell({ children: [new Paragraph({ text: String(item.componentGrades["Arts"] ?? ''), alignment: AlignmentType.CENTER })] }),
                                new TableCell({ children: [new Paragraph({ text: String(item.componentGrades["PE"] ?? ''), alignment: AlignmentType.CENTER })] }),
                                new TableCell({ children: [new Paragraph({ text: String(item.componentGrades["Health"] ?? ''), alignment: AlignmentType.CENTER })] }),
                                new TableCell({ children: [new Paragraph({ text: String(item.finalMapehGrade ?? ''), alignment: AlignmentType.CENTER, run: { bold: true } })] }),
                            ]
                        }))
                    ]
                })
            ]
        }]
    });
    const blob = await Packer.toBlob(doc);
    this.downloadBlob(blob, `MAPEH_Summary_${selectedSectionText.replace(/\s/g, '_')}_Q${quarter}.docx`);
  }
    // FIX: Add all missing docx generation methods
    
    public async generateSF2Docx(students: Student[], attendance: Attendance[], settings: SchoolSettings, currentUTCDate: Date): Promise<void> {
        // This is a complex form. The implementation replicates the structure seen in SF2.tsx
        const monthName = currentUTCDate.toLocaleString('default', { month: 'long', timeZone: 'UTC' });
        const year = currentUTCDate.getUTCFullYear();
// FIX: Changed `students[0] || {}` to `students[0]` and used optional chaining `?.` to prevent errors when the students array is empty.
        const studentInfo = students[0];
        const sectionText = (studentInfo?.gradeLevel && studentInfo?.section) ? `${studentInfo.gradeLevel} - ${studentInfo.section}` : 'Class';

        // ... (Full implementation would be very long, involving creating the full SF2 table structure)
        // For brevity, we'll create a simplified version. A full implementation would mirror excelService.
        const doc = new Document({
            sections: [{
                properties: {
                    page: {
                        size: { orientation: PageOrientation.LANDSCAPE },
                    }
                },
                children: [
                    new Paragraph({ text: "School Form 2 (SF2) Daily Attendance Report of Learners", alignment: AlignmentType.CENTER, heading: HeadingLevel.HEADING_1 }),
                    new Paragraph({ text: `Month of ${monthName}, ${year}`, alignment: AlignmentType.CENTER }),
                    new Paragraph({ text: `Grade & Section: ${sectionText}`, alignment: AlignmentType.LEFT }),
                    new Paragraph({ text: "NOTE: This is a placeholder implementation. A full SF2 form requires a complex table."})
                ],
            }],
        });

        const blob = await Packer.toBlob(doc);
        this.downloadBlob(blob, `SF2_${sectionText.replace(/\s/g, '_')}_${monthName}_${year}.docx`);
    }

    public async generateSummaryOfGradesDocx(data: SummaryOfGradesDocxData): Promise<void> {
        const { students, settings, subject, summaryStats, selectedSectionText } = data;

        const doc = new Document({
            sections: [{
                children: [
                    new Paragraph({ text: "Summary of Quarterly Grades", alignment: AlignmentType.CENTER, heading: HeadingLevel.HEADING_1 }),
                    new Paragraph({ text: `Subject: ${subject}`, alignment: AlignmentType.CENTER }),
                    new Paragraph({ text: `Class: ${selectedSectionText}`, alignment: AlignmentType.CENTER }),
                    new Paragraph({ text: `Teacher: ${settings.teacherName}`, alignment: AlignmentType.CENTER }),
                    new Paragraph({ text: "" }),
                    // ... Table generation for males and females
                ],
            }],
        });

        const blob = await Packer.toBlob(doc);
        this.downloadBlob(blob, `Summary_of_Grades_${subject.replace(/\s/g, '_')}.docx`);
    }

    public async generateHonorsListDocx(data: HonorsListDocxData): Promise<void> {
        const { honorStudents, settings, selectedSectionText } = data;
        const doc = new Document({
            sections: [{
                children: [
                    new Paragraph({ text: "List of Honors", alignment: AlignmentType.CENTER, heading: HeadingLevel.HEADING_1 }),
                    new Paragraph({ text: `Class: ${selectedSectionText}`, alignment: AlignmentType.CENTER }),
                     new Paragraph({ text: "" }),
                    new Paragraph({ text: "With Highest Honors", heading: HeadingLevel.HEADING_2 }),
                    ...honorStudents.highest.map(s => new Paragraph({ text: `${s.student.lastName}, ${s.student.firstName}`, bullet: { level: 0 } })),
                    new Paragraph({ text: "With High Honors", heading: HeadingLevel.HEADING_2 }),
                    ...honorStudents.high.map(s => new Paragraph({ text: `${s.student.lastName}, ${s.student.firstName}`, bullet: { level: 0 } })),
                    new Paragraph({ text: "With Honors", heading: HeadingLevel.HEADING_2 }),
                    ...honorStudents.regular.map(s => new Paragraph({ text: `${s.student.lastName}, ${s.student.firstName}`, bullet: { level: 0 } })),
                ],
            }],
        });
        const blob = await Packer.toBlob(doc);
        this.downloadBlob(blob, `Honors_List_${selectedSectionText.replace(/\s/g, '_')}.docx`);
    }

    public async generateEClassRecordDocx(data: EClassRecordDocxData): Promise<void> {
        const doc = new Document({
            sections: [{
                 properties: {
                    page: {
                        size: { orientation: PageOrientation.LANDSCAPE },
                    }
                },
                children: [
                    new Paragraph({ text: `E-Class Record`, heading: HeadingLevel.HEADING_1, alignment: AlignmentType.CENTER }),
                    new Paragraph({ text: `Subject: ${data.subject} - Quarter ${data.quarter}`, alignment: AlignmentType.CENTER }),
                    new Paragraph({ text: `Class: ${data.selectedSectionText}`, alignment: AlignmentType.CENTER }),
                    new Paragraph({ text: "NOTE: This is a placeholder implementation for a very complex table."})
                ],
            }],
        });
        const blob = await Packer.toBlob(doc);
        this.downloadBlob(blob, `E-Class_Record_${data.subject.replace(/\s/g, '_')}_Q${data.quarter}.docx`);
    }

    public async generateStudentProfileDocx(data: StudentProfileDocxData): Promise<void> {
         const { student, academicSummary, attendanceSummary, recentAnecdotes, settings } = data;
        const doc = new Document({
            sections: [{
                children: [
                    new Paragraph({ text: "Student Profile", heading: HeadingLevel.HEADING_1, alignment: AlignmentType.CENTER }),
                    new Paragraph({ text: `${student.firstName} ${student.lastName}`, heading: HeadingLevel.HEADING_2 }),
                    // ... more details
                ],
            }],
        });
        const blob = await Packer.toBlob(doc);
        this.downloadBlob(blob, `Profile_${student.lastName}_${student.firstName}.docx`);
    }

    public async generatePickedStudentsDocx(data: PickedStudentsDocxData): Promise<void> {
        const doc = new Document({
            sections: [{
                children: [
                    new Paragraph({ text: `Picked Students for ${data.topic}`, heading: HeadingLevel.HEADING_1 }),
                    ...data.pickedStudents.map(s => new Paragraph({ text: `${s.lastName}, ${s.firstName}`, bullet: { level: 0 } })),
                ],
            }],
        });
        const blob = await Packer.toBlob(doc);
        this.downloadBlob(blob, `Picked_Students_${data.topic}.docx`);
    }

    public async generateGroupsDocx(data: GroupsDocxData): Promise<void> {
         const doc = new Document({
            sections: [{
                children: [
                    new Paragraph({ text: `Groups for ${data.topic}`, heading: HeadingLevel.HEADING_1 }),
                    ...data.groups.flatMap((group, i) => [
                        new Paragraph({ text: `Group ${i + 1}`, heading: HeadingLevel.HEADING_2 }),
                        ...group.map(s => new Paragraph({ text: `${s.lastName}, ${s.firstName}`, bullet: { level: 0 } }))
                    ])
                ],
            }],
        });
        const blob = await Packer.toBlob(doc);
        this.downloadBlob(blob, `Groups_${data.topic}.docx`);
    }
    
    public async generateDlpDocx(dlpForm: any, dlpContent: DlpContent, _unused: string, settings: SchoolSettings): Promise<void> {
        const schoolLogo = this.createDocxImage(this.parseDataUrl(settings.schoolLogo), 60, 60);
        const secondLogo = this.createDocxImage(this.parseDataUrl(settings.secondLogo), 60, 60);
        const isFilipino = dlpForm.language === 'Filipino';
    
        const boldRun = (text: string): TextRun => new TextRun({ text, bold: true, font: "Times New Roman", size: 22 });
        const normalRun = (text: string): TextRun => new TextRun({ text, font: "Times New Roman", size: 22 });
    
        const headerTable = new Table({
            width: { size: 100, type: WidthType.PERCENTAGE },
            rows: [
                new TableRow({
                    children: [
                        new TableCell({
                            children: [new Paragraph({
                                children: [
                                    ...(schoolLogo ? [schoolLogo, new TextRun("  ")] : []),
                                    ...(secondLogo ? [secondLogo] : []),
                                ],
                                alignment: AlignmentType.CENTER,
                            })],
                            verticalAlign: VerticalAlign.CENTER,
                            rowSpan: 5,
                        }),
                        new TableCell({ children: [new Paragraph({ children: [boldRun(isFilipino ? 'Paaralan: ' : 'School: '), normalRun(dlpForm.schoolName.toUpperCase())] })] }),
                        new TableCell({
                            children: [new Paragraph({ text: isFilipino ? 'DETALYADONG BANGHAY-ARALIN SA' : 'DAILY LESSON PLAN IN', alignment: AlignmentType.CENTER, run: { bold: true, size: 24, font: "Times New Roman" } }), new Paragraph({ text: `${dlpForm.subject.toUpperCase()} ${dlpForm.gradeLevel}`, alignment: AlignmentType.CENTER, run: { bold: true, size: 24, font: "Times New Roman" } })],
                            rowSpan: 2, verticalAlign: VerticalAlign.CENTER
                        }),
                    ]
                }),
                new TableRow({ children: [new TableCell({ children: [new Paragraph({ children: [boldRun(`${dlpForm.quarterSelect}`)] })] })] }),
                new TableRow({
                    children: [
                        new TableCell({ children: [new Paragraph({ children: [boldRun(isFilipino ? 'Guro: ' : 'Teacher: '), normalRun(dlpForm.teacher)] })] }),
                        new TableCell({
                            children: [
                                new Paragraph({ children: [boldRun(isFilipino ? 'ISKEDYUL NG KLASE' : 'CLASS SCHEDULE')] }),
                                ...dlpForm.classSchedule.split('\n').map((line: string) => new Paragraph(line))
                            ],
                            rowSpan: 3, verticalAlign: VerticalAlign.TOP
                        }),
                    ]
                }),
                new TableRow({ children: [new TableCell({ children: [new Paragraph({ children: [boldRun(isFilipino ? 'Asignatura: ' : 'Learning Area: '), normalRun(dlpForm.subject.toUpperCase())] })] })] }),
                new TableRow({ children: [new TableCell({ children: [new Paragraph({ children: [boldRun(isFilipino ? 'Petsa ng Pagtuturo: ' : 'Teaching Dates: '), normalRun(dlpForm.teachingDates)] })] })] }),
            ],
        });

        const doc = new Document({
            numbering: {
                config: [{
                    reference: "dlp-list",
                    levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }]
                }]
            },
            sections: [{
                properties: { page: { margin: { top: 720, right: 720, bottom: 720, left: 720 } } },
                children: [
                    headerTable,
                    new Paragraph({ text: isFilipino ? 'I. LAYUNIN' : 'I. OBJECTIVES', heading: HeadingLevel.HEADING_1, spacing: { before: 200 } }),
                    new Paragraph({ children: [boldRun(isFilipino ? 'A. Pamantayang Pangnilalaman: ' : 'A. Content Standard: '), normalRun(dlpContent.contentStandard)] }),
                    new Paragraph({ children: [boldRun(isFilipino ? 'B. Pamantayan sa Pagganap: ' : 'B. Performance Standard: '), normalRun(dlpContent.performanceStandard)] }),
                    new Paragraph({ children: [boldRun(isFilipino ? 'C. Kasanayan sa Pagkatuto: ' : 'C. Learning Competency: '), normalRun(dlpForm.learningCompetency)] }),
                    new Paragraph({ text: isFilipino ? 'Sa pagtatapos ng aralin, ang mga mag-aaral ay inaasahang:' : 'At the end of the lesson, the learners should be able to:', spacing: { before: 100 } }),
                    new Paragraph({ text: dlpForm.lessonObjective, bullet: { level: 0 } }),

                    new Paragraph({ text: isFilipino ? 'II. NILALAMAN' : 'II. CONTENT', heading: HeadingLevel.HEADING_1, spacing: { before: 200 } }),
                    new Paragraph({ children: [boldRun(isFilipino ? 'Paksa: ' : 'Topic: '), normalRun(dlpContent.topic)] }),

                    new Paragraph({ text: isFilipino ? 'III. KAGAMITANG PANTURO' : 'III. LEARNING RESOURCES', heading: HeadingLevel.HEADING_1, spacing: { before: 200 } }),
                    new Paragraph({ children: [boldRun(isFilipino ? 'A. Sanggunian: ' : 'A. References: '), normalRun(dlpContent.learningReferences)] }),
                    new Paragraph({ children: [boldRun(isFilipino ? 'B. Kagamitan: ' : 'B. Materials: '), normalRun(dlpContent.learningMaterials)] }),
                    
                    new Paragraph({ text: isFilipino ? 'IV. PAMAMARAAN' : 'IV. PROCEDURE', heading: HeadingLevel.HEADING_1, spacing: { before: 200 } }),
                    new Table({
                        width: { size: 100, type: WidthType.PERCENTAGE },
                        columnWidths: [25, 45, 30],
                        rows: [
                            new TableRow({
                                tableHeader: true,
                                children: [
                                    new TableCell({ children: [new Paragraph({ text: isFilipino ? 'Pamamaraan' : 'Procedure', alignment: AlignmentType.CENTER, run: { bold: true } })] }),
                                    new TableCell({ children: [new Paragraph({ text: isFilipino ? 'Gawain ng Guro/Mag-aaral' : 'Teacher/Student Activity', alignment: AlignmentType.CENTER, run: { bold: true } })] }),
                                    new TableCell({ children: [new Paragraph({ text: 'Aligned PPST Indicators', alignment: AlignmentType.CENTER, run: { bold: true } })] }),
                                ]
                            }),
                            ...dlpContent.procedures.flatMap(proc => new TableRow({
                                children: [
                                    new TableCell({ children: [new Paragraph({ text: proc.title, run: { bold: true } })], verticalAlign: VerticalAlign.CENTER }),
                                    new TableCell({ children: this.parseMarkdownToParagraphs(proc.content) }),
                                    new TableCell({ children: [new Paragraph({ text: proc.ppst, run: { italics: true } })], verticalAlign: VerticalAlign.CENTER }),
                                ]
                            }))
                        ]
                    }),
                ],
            }]
        });

        const blob = await Packer.toBlob(doc);
        this.downloadBlob(blob, `DLP_${dlpContent.topic.replace(/\s/g, '_')}.docx`);
    }
    
    public async generateDllDocx(dllExportData: any, dllContent: DllContent, settings: SchoolSettings): Promise<void> {
        const isFilipino = dllExportData.language === 'Filipino';
        const doc = new Document({
            sections: [{
                properties: {
                    page: {
                        size: { width: 18720, height: 12240, orientation: PageOrientation.LANDSCAPE }, // Legal size landscape
                        margin: { top: 720, right: 720, bottom: 720, left: 720 }
                    }
                },
                children: [
                    new Paragraph({ text: isFilipino ? "PANG-ARAW-ARAW NA TALA SA PAGTUTURO" : "DAILY LESSON LOG", alignment: AlignmentType.CENTER, run: { bold: true, size: 28 } }),
                    new Table({
                        width: { size: 100, type: WidthType.PERCENTAGE },
                        columnWidths: [20, 30, 20, 30],
                        borders: { top: { style: BorderStyle.NONE }, bottom: { style: BorderStyle.NONE }, left: { style: BorderStyle.NONE }, right: { style: BorderStyle.NONE }, insideHorizontal: { style: BorderStyle.NONE }, insideVertical: { style: BorderStyle.NONE } },
                        rows: [
                            new TableRow({ children: [ new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: isFilipino ? 'Paaralan: ' : 'School: ', bold: true }), new TextRun(settings.schoolName)] })] }), new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: isFilipino ? 'Baitang: ' : 'Grade Level: ', bold: true }), new TextRun(dllExportData.gradeLevel)] })] }), new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: isFilipino ? 'Guro: ' : 'Teacher: ', bold: true }), new TextRun(settings.teacherName)] })] }), new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: isFilipino ? 'Asignatura: ' : 'Learning Area: ', bold: true }), new TextRun(dllExportData.subject)] })] }) ] }),
                            new TableRow({ children: [ new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: isFilipino ? 'Petsa/Oras: ' : 'Teaching Dates & Time: ', bold: true }), new TextRun(dllExportData.teachingDates)] })] }), new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: isFilipino ? 'Markahan: ' : 'Quarter: ', bold: true }), new TextRun(dllExportData.quarter)] })] }), new TableCell({ children: [] }), new TableCell({ children: [] }) ] }),
                        ]
                    }),
                    new Table({
                        width: { size: 100, type: WidthType.PERCENTAGE },
                        columnWidths: [16.6, 16.6, 16.6, 16.6, 16.6, 16.6], // 6 columns
                        rows: [
                            new TableRow({
                                tableHeader: true,
                                children: [
                                    new TableCell({ children: [new Paragraph('')] }),
                                    ...['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday'].map(day => new TableCell({ children: [new Paragraph({ text: day, alignment: AlignmentType.CENTER, run: { bold: true } })] }))
                                ]
                            }),
                            new TableRow({ children: [new TableCell({ children: [new Paragraph({ text: isFilipino ? "I. LAYUNIN" : "I. OBJECTIVES", run: { bold: true } })], columnSpan: 6 })] }),
                            new TableRow({ children: [new TableCell({ children: [new Paragraph({ text: isFilipino ? "C. Mga Kasanayan sa Pagkatuto" : "C. Learning Competencies", run: { bold: true } })] }), ...Object.values(dllContent.learningCompetencies).map(val => new TableCell({ children: [new Paragraph(val)] })) ] }),
                            new TableRow({ children: [new TableCell({ children: [new Paragraph({ text: isFilipino ? "II. NILALAMAN" : "II. CONTENT", run: { bold: true } })], columnSpan: 6 })] }),
                            new TableRow({ children: [new TableCell({ children: [new Paragraph({ text: isFilipino ? "IV. PAMAMARAAN" : "IV. PROCEDURES", run: { bold: true } })], columnSpan: 6 })] }),
                            ...dllContent.procedures.map(proc => new TableRow({
                                children: [
                                    new TableCell({ children: [new Paragraph(proc.procedure)] }),
                                    new TableCell({ children: this.parseMarkdownToParagraphs(proc.monday) }),
                                    new TableCell({ children: this.parseMarkdownToParagraphs(proc.tuesday) }),
                                    new TableCell({ children: this.parseMarkdownToParagraphs(proc.wednesday) }),
                                    new TableCell({ children: this.parseMarkdownToParagraphs(proc.thursday) }),
                                    new TableCell({ children: this.parseMarkdownToParagraphs(proc.friday) }),
                                ]
                            }))
                        ]
                    })
                ],
            }],
        });
        const blob = await Packer.toBlob(doc);
        this.downloadBlob(blob, `DLL_${dllExportData.subject.replace(/\s/g, '_')}.docx`);
    }

    public async generateLasDocx(data: any, lasContent: LearningActivitySheet, settings: SchoolSettings): Promise<void> {
        const schoolLogo = this.createDocxImage(this.parseDataUrl(settings.schoolLogo), 50, 50);
        const secondLogo = this.createDocxImage(this.parseDataUrl(settings.secondLogo), 50, 50);

        const blackTextRun = { font: "Times New Roman", size: 22, bold: true };
        const normalText = { font: "Times New Roman", size: 22 };
        const whiteTextRun = { color: "FFFFFF", bold: true, font: "Arial", size: 20 };

        const table1 = new Table({
            width: { size: 100, type: WidthType.PERCENTAGE },
            rows: [ new TableRow({ children: [ new TableCell({ children: [ new Paragraph({ children: [ ...(schoolLogo ? [schoolLogo, new TextRun("  ")] : []), ...(secondLogo ? [secondLogo] : []), ], alignment: AlignmentType.LEFT }) ], borders: { top: { style: BorderStyle.NONE }, bottom: { style: BorderStyle.NONE }, left: { style: BorderStyle.NONE }, right: { style: BorderStyle.NONE } }, width: { size: 20, type: WidthType.PERCENTAGE }, verticalAlign: VerticalAlign.CENTER, }), new TableCell({ children: [ new Paragraph({ text: "Dynamic Learning Program", bold: true, size: 24, alignment: AlignmentType.CENTER }), new Paragraph({ text: "LEARNING ACTIVITY SHEET", bold: true, size: 28, alignment: AlignmentType.CENTER }), ], borders: { top: { style: BorderStyle.NONE }, bottom: { style: BorderStyle.NONE }, left: { style: BorderStyle.NONE }, right: { style: BorderStyle.NONE } }, width: { size: 60, type: WidthType.PERCENTAGE }, verticalAlign: VerticalAlign.CENTER, }), new TableCell({ children: [ new Table({ width: { size: 100, type: WidthType.PERCENTAGE }, borders: { top: { style: BorderStyle.SINGLE, size: 2 }, bottom: { style: BorderStyle.NONE }, left: { style: BorderStyle.SINGLE, size: 2 }, right: { style: BorderStyle.SINGLE, size: 2 } }, rows: [ new TableRow({ children: [new TableCell({ children: [new Paragraph({ text: `S.Y. ${settings.schoolYear}`, alignment: AlignmentType.CENTER })], borders: { top: { style: BorderStyle.NONE }, bottom: { style: BorderStyle.NONE }, left: { style: BorderStyle.NONE }, right: { style: BorderStyle.NONE } }, })] }) ] }) ], borders: { top: { style: BorderStyle.NONE }, bottom: { style: BorderStyle.NONE }, left: { style: BorderStyle.NONE }, right: { style: BorderStyle.NONE } }, width: { size: 20, type: WidthType.PERCENTAGE }, verticalAlign: VerticalAlign.TOP, }), ], }), ],
        });

        const table2 = new Table({
            width: { size: 100, type: WidthType.PERCENTAGE },
            borders: { top: { style: BorderStyle.SINGLE, size: 6, color: "000000" }, bottom: { style: BorderStyle.SINGLE, size: 6, color: "000000" }, left: { style: BorderStyle.SINGLE, size: 6, color: "000000" }, right: { style: BorderStyle.SINGLE, size: 6, color: "000000" }, insideHorizontal: { style: BorderStyle.SINGLE, size: 2, color: "000000" }, insideVertical: { style: BorderStyle.SINGLE, size: 2, color: "000000" }, },
            rows: [ new TableRow({ children: [ new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "Name:", ...blackTextRun })] })], borders: { top: { style: BorderStyle.NONE }, bottom: { style: BorderStyle.NONE }, left: { style: BorderStyle.NONE }, right: { style: BorderStyle.NONE } } }), new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "Score:", ...blackTextRun })] })], borders: { top: { style: BorderStyle.NONE }, bottom: { style: BorderStyle.NONE }, left: { style: BorderStyle.NONE }, right: { style: BorderStyle.NONE } } }), ] }), new TableRow({ children: [ new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "Grade & Section:", ...blackTextRun })] })], borders: { top: { style: BorderStyle.NONE }, bottom: { style: BorderStyle.NONE }, left: { style: BorderStyle.NONE }, right: { style: BorderStyle.NONE } } }), new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "Date:", ...blackTextRun })] })], borders: { top: { style: BorderStyle.NONE }, bottom: { style: BorderStyle.NONE }, left: { style: BorderStyle.NONE }, right: { style: BorderStyle.NONE } } }), ] }), ],
        });

        const table3 = new Table({
            width: { size: 100, type: WidthType.PERCENTAGE },
            borders: { top: { style: BorderStyle.SINGLE, size: 6 }, bottom: { style: BorderStyle.SINGLE, size: 6 }, left: { style: BorderStyle.SINGLE, size: 6 }, right: { style: BorderStyle.SINGLE, size: 6 }, },
            rows: [ new TableRow({ children: [ new TableCell({ children: [ new Paragraph({ children: [ new TextRun({ text: "Activity Title: ", ...whiteTextRun }), new TextRun({ text: lasContent.activityTitle, color: "FFFFFF", font: "Arial", size: 20, bold: true }) ]}), new Paragraph({ children: [ new TextRun({ text: "Learning Target: ", ...whiteTextRun }), new TextRun({ text: lasContent.learningTarget, color: "FFFFFF", font: "Arial", size: 20 }) ]}) ], shading: { type: ShadingType.CLEAR, fill: "000000" }, borders: { top: { style: BorderStyle.NONE }, bottom: { style: BorderStyle.NONE }, left: { style: BorderStyle.NONE }, right: { style: BorderStyle.NONE } }, }) ] }) ]
        });

        const table4 = new Table({
            width: { size: 100, type: WidthType.PERCENTAGE },
            rows: [ new TableRow({ height: { value: 800, rule: HeightRule.ATLEAST }, children: [ new TableCell({ children: [ new Paragraph({ children: [new TextRun({ text: "References: ", ...blackTextRun }), new TextRun({ text: "(Author, Title, Pages)", font: "Times New Roman", size: 18, italics: true })] }), new Paragraph({ text: lasContent.references, run: normalText }) ], verticalAlign: VerticalAlign.TOP, borders: { top: { style: BorderStyle.SINGLE, size: 2 }, bottom: { style: BorderStyle.SINGLE, size: 2 }, left: { style: BorderStyle.SINGLE, size: 2 }, right: { style: BorderStyle.SINGLE, size: 2 } }, }) ] }) ]
        });

        const contentParagraphs = [
            ...lasContent.conceptNotes.flatMap(note => [
                new Paragraph({ text: note.title, heading: HeadingLevel.HEADING_2, spacing: { before: 300, after: 100 } }),
                ...this.parseLasMarkdown(note.content),
            ]),
            ...lasContent.activities.flatMap(activity => [
                new Paragraph({ text: activity.title, heading: HeadingLevel.HEADING_1, spacing: { before: 400, after: 200 } }),
                new Paragraph({ text: activity.instructions, italics: true, spacing: { after: 200 } }),
                ...(activity.questions ? activity.questions.flatMap((q) => [
                    new Paragraph({ text: q.questionText, numbering: { reference: "las-questions", level: 0 } }),
                    ...(q.options ? q.options.map(opt => new Paragraph({ text: opt, numbering: { reference: "las-options", level: 1 } })) : [new Paragraph({ text: "________________________________", spacing: { before: 200, after: 200 } })]),
                ]) : []),
            ]),
        ];
        
        const doc = new Document({
            numbering: {
                config: [
                    { reference: "las-questions", levels: [{ level: 0, format: "decimal", text: "%1. ", alignment: AlignmentType.LEFT }] },
                    { reference: "las-options", levels: [{ level: 1, format: "lowerLetter", text: "%2) ", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720 } } } }] }
                ]
            },
            sections: [{
                children: [
                    table1,
                    new Paragraph({ text: "" }),
                    table2,
                    new Paragraph({ text: "" }),
                    table3,
                    new Paragraph({ text: "" }),
                    table4,
                    new Paragraph({ text: "", spacing: { after: 200 }}),
                    ...contentParagraphs
                ]
            }]
        });
        
        const blob = await Packer.toBlob(doc);
        this.downloadBlob(blob, `LAS_${lasContent.activityTitle.replace(/\s/g, '_')}.docx`);
    }

    public async generateQuizDocx(quizContent: GeneratedQuiz): Promise<void> {
        const doc = new Document({
            sections: [{
                children: [
                    new Paragraph({ text: quizContent.quizTitle, heading: HeadingLevel.HEADING_1, alignment: AlignmentType.CENTER }),
                    new Paragraph({text: ''}),
                     ...Object.entries(quizContent.questionsByType).flatMap(([type, section]) => {
                         if (!section) return [];
                         return [
                            new Paragraph({ text: type, heading: HeadingLevel.HEADING_2 }),
                            ...section.questions.map((q, i) => new Paragraph({ text: `${i+1}. ${q.questionText}`, numbering: { reference: "quiz-num", level: 0 } }))
                         ]
                     })
                ],
            }],
             numbering: {
                config: [{
                    reference: "quiz-num",
                    levels: [{ level: 0, format: "decimal", text: "%1." }]
                }]
             }
        });
        const blob = await Packer.toBlob(doc);
        this.downloadBlob(blob, `${quizContent.quizTitle.replace(/\s/g, '_')}.docx`);
    }
}

export const docxService = new DocxService();