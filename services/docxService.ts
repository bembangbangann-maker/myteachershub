
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
  Header,
  convertInchesToTwip,
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
                } as any);
            } else {
                 return new ImageRun({
                    type: parsedImage.type as "jpg" | "png" | "gif" | "bmp",
                    data: this.base64ToArrayBuffer(parsedImage.data),
                    transformation: {
                        width: width,
                        height: height,
                    },
                    ...options
                } as any);
            }
        } catch (e) {
            console.error("Failed to create ImageRun. The image data might be corrupt.", e);
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
        const baseFont = "Century Gothic";
        const fontSize = 24; // 12pt
        const fontOptions = { font: baseFont, size: fontSize };

        if (!text || text.trim() === '') {
            return [new Paragraph({ children: [new TextRun({ text: " ", ...fontOptions })], spacing: { after: 0 } })];
        }

        const paragraphs: Paragraph[] = [];
        const lines = text.split('\n');

        for (const line of lines) {
            if (line.trim() === '') {
                paragraphs.push(new Paragraph({ children: [new TextRun({ text: " ", ...fontOptions })], spacing: { after: 120 } }));
                continue;
            }

            const children: TextRun[] = [];
            const parts = line.split(/(\*\*.*?\*\*|\*.*?\*)/g).filter(p => p !== '');

            for (const part of parts) {
                if (part.startsWith('**') && part.endsWith('**')) {
                    children.push(new TextRun({ text: part.slice(2, -2), bold: true, ...fontOptions }));
                } else if (part.startsWith('*') && part.endsWith('*')) {
                    children.push(new TextRun({ text: part.slice(1, -1), italics: true, ...fontOptions }));
                } else {
                    children.push(new TextRun({ text: part, ...fontOptions }));
                }
            }
            
            const isListItem = /^\s*•\s+/.test(line.trim()) || /^\d+\./.test(line.trim());
            const isSubListItem = /^\s*o\s+/.test(line.trim());

            paragraphs.push(new Paragraph({
                children: children.length > 0 ? children : [new TextRun({ text: " ", ...fontOptions })],
                bullet: isSubListItem ? { level: 1 } : (isListItem ? { level: 0 } : undefined),
                indent: isSubListItem ? { left: 1080, hanging: 360 } : (isListItem ? { left: 720, hanging: 360 } : undefined),
                spacing: { after: 120 } 
            }));
        }

        if (paragraphs.length === 0) {
             return [new Paragraph({ children: [new TextRun({ text: " ", ...fontOptions })], spacing: { after: 0 } })];
        }

        return paragraphs;
    }

    public async generateLasDocx(
        lasForm: any,
        lasContent: LearningActivitySheet,
        settings: SchoolSettings
    ): Promise<void> {
        const sections: (Paragraph | Table | Paragraph)[] = [];
        const baseFont = "Century Gothic";
        const headerFontSize = 22; // 11pt
        const contentFontSize = 24; // 12pt
        
        const headerFont = { font: baseFont, size: headerFontSize, bold: true };
        const fieldFont = { font: baseFont, size: headerFontSize };
        const contentFont = { font: baseFont, size: contentFontSize };
        const titleFont = { font: baseFont, size: 32, bold: true }; // 16pt

        const days = lasContent?.days || [];
        
        if (days.length === 0) {
             sections.push(new Paragraph({ text: "No content generated for LAS." }));
        }

        days.forEach((dayData, index) => {
            if (index > 0) {
                sections.push(new Paragraph({ children: [new PageBreak()] }));
            }

            // --- HEADER ---
            
            // Logos
            const secondLogoRun = this.createDocxImage(this.parseDataUrl(settings.secondLogo), 50, 50) || new TextRun("");
            const schoolLogoRun = this.createDocxImage(this.parseDataUrl(settings.schoolLogo), 50, 50) || new TextRun("");

            // Top Table (Logos | Title | Info)
            const topHeaderTable = new Table({
                width: { size: 100, type: WidthType.PERCENTAGE },
                rows: [
                    new TableRow({
                        children: [
                            // Col 1: Logos
                            new TableCell({
                                width: { size: 15, type: WidthType.PERCENTAGE },
                                children: [
                                    new Paragraph({
                                        children: [secondLogoRun, new TextRun("  "), schoolLogoRun],
                                        alignment: AlignmentType.CENTER,
                                    })
                                ],
                                borders: { 
                                    top: { style: BorderStyle.SINGLE, size: 4 }, 
                                    bottom: { style: BorderStyle.SINGLE, size: 4 }, 
                                    left: { style: BorderStyle.SINGLE, size: 4 }, 
                                    right: { style: BorderStyle.NONE } 
                                },
                                verticalAlign: VerticalAlign.CENTER,
                            }),
                            // Col 2: Dynamic Learning Program
                            new TableCell({
                                width: { size: 50, type: WidthType.PERCENTAGE },
                                children: [
                                    new Paragraph({
                                        children: [new TextRun({ text: "Dynamic Learning Program", font: baseFont, size: 28, bold: true })],
                                        alignment: AlignmentType.CENTER,
                                    })
                                ],
                                borders: { 
                                    top: { style: BorderStyle.SINGLE, size: 4 }, 
                                    bottom: { style: BorderStyle.SINGLE, size: 4 }, 
                                    left: { style: BorderStyle.NONE }, 
                                    right: { style: BorderStyle.NONE } 
                                },
                                verticalAlign: VerticalAlign.CENTER,
                            }),
                            // Col 3: Info Box (Nested Table)
                            new TableCell({
                                width: { size: 35, type: WidthType.PERCENTAGE },
                                children: [
                                    new Table({
                                        width: { size: 100, type: WidthType.PERCENTAGE },
                                        rows: [
                                            new TableRow({
                                                children: [
                                                    new TableCell({
                                                        children: [new Paragraph({ children: [new TextRun({ text: `S.Y. ${this.safeString(settings.schoolYear)}`, ...headerFont })], alignment: AlignmentType.CENTER })],
                                                        borders: { bottom: { style: BorderStyle.SINGLE, size: 4 }, top: { style: BorderStyle.NONE }, left: { style: BorderStyle.SINGLE, size: 4 }, right: { style: BorderStyle.NONE } }
                                                    })
                                                ]
                                            }),
                                            new TableRow({
                                                children: [
                                                    new TableCell({
                                                        children: [new Paragraph({ children: [new TextRun({ text: "Subject: ", ...headerFont }), new TextRun({ text: this.safeString(lasForm.subject), ...fieldFont })], alignment: AlignmentType.LEFT })],
                                                        borders: { bottom: { style: BorderStyle.SINGLE, size: 4 }, top: { style: BorderStyle.NONE }, left: { style: BorderStyle.SINGLE, size: 4 }, right: { style: BorderStyle.NONE } }
                                                    })
                                                ]
                                            }),
                                            new TableRow({
                                                children: [
                                                    new TableCell({
                                                        children: [new Paragraph({ children: [new TextRun({ text: "Q1 - LAS - _________", ...headerFont })], alignment: AlignmentType.CENTER })],
                                                        borders: { bottom: { style: BorderStyle.NONE }, top: { style: BorderStyle.NONE }, left: { style: BorderStyle.SINGLE, size: 4 }, right: { style: BorderStyle.NONE } }
                                                    })
                                                ]
                                            }),
                                        ]
                                    })
                                ],
                                borders: { 
                                    top: { style: BorderStyle.SINGLE, size: 4 }, 
                                    bottom: { style: BorderStyle.SINGLE, size: 4 }, 
                                    left: { style: BorderStyle.NONE }, 
                                    right: { style: BorderStyle.SINGLE, size: 4 } 
                                },
                                margins: { top: 0, bottom: 0, left: 0, right: 0 }
                            })
                        ]
                    })
                ]
            });
            sections.push(topHeaderTable);

            // Title
            sections.push(new Paragraph({
                children: [new TextRun({ text: "LEARNING ACTIVITY SHEET", ...titleFont })],
                alignment: AlignmentType.CENTER,
                spacing: { before: 120, after: 120 }
            }));

            // Student Info Table
            const studentInfoTable = new Table({
                width: { size: 100, type: WidthType.PERCENTAGE },
                rows: [
                    new TableRow({
                        children: [
                            new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "Name:", ...headerFont })] })], width: { size: 10, type: WidthType.PERCENTAGE } }),
                            new TableCell({ children: [new Paragraph({ text: "", ...fieldFont })], width: { size: 60, type: WidthType.PERCENTAGE } }),
                            new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "Score:", ...headerFont })] })], width: { size: 10, type: WidthType.PERCENTAGE } }),
                            new TableCell({ children: [new Paragraph({ text: "", ...fieldFont })], width: { size: 20, type: WidthType.PERCENTAGE } }),
                        ]
                    }),
                    new TableRow({
                        children: [
                            new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "Grade & Section:", ...headerFont })] })], width: { size: 20, type: WidthType.PERCENTAGE } }),
                            new TableCell({ children: [new Paragraph({ text: this.safeString(lasForm.gradeLevel), ...fieldFont })], width: { size: 50, type: WidthType.PERCENTAGE } }),
                            new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "Date:", ...headerFont })] })], width: { size: 10, type: WidthType.PERCENTAGE } }),
                            new TableCell({ children: [new Paragraph({ text: "", ...fieldFont })], width: { size: 20, type: WidthType.PERCENTAGE } }),
                        ]
                    })
                ]
            });
            sections.push(studentInfoTable);

            // Activity Type Table
            const check = (label: string) => {
                const isChecked = this.safeString(lasForm.activityType).toLowerCase().includes(label.toLowerCase().split(':')[0].split(' ')[0]);
                return isChecked ? `\u2611 ${label}` : `\u2610 ${label}`;
            };

            const activityTypeTable = new Table({
                width: { size: 100, type: WidthType.PERCENTAGE },
                rows: [
                    new TableRow({
                        children: [
                            new TableCell({ 
                                children: [new Paragraph({ children: [new TextRun({ text: "Type of Activity: (Check or choose from below.)", ...headerFont, italics: true })] })], 
                                columnSpan: 4,
                                borders: { bottom: { style: BorderStyle.NONE, size: 0 } }
                            })
                        ]
                    }),
                    new TableRow({
                        children: [
                            new TableCell({ children: [new Paragraph({ text: check("Concept Notes"), ...fieldFont })], width: { size: 25, type: WidthType.PERCENTAGE }, borders: { top: { style: BorderStyle.NONE, size: 0 }, bottom: { style: BorderStyle.NONE, size: 0 }, right: { style: BorderStyle.NONE, size: 0 } } }),
                            new TableCell({ children: [new Paragraph({ text: check("Performance Task"), ...fieldFont })], width: { size: 25, type: WidthType.PERCENTAGE }, borders: { top: { style: BorderStyle.NONE, size: 0 }, bottom: { style: BorderStyle.NONE, size: 0 }, left: { style: BorderStyle.NONE, size: 0 }, right: { style: BorderStyle.NONE, size: 0 } } }),
                            new TableCell({ children: [new Paragraph({ text: check("Formal Theme"), ...fieldFont })], width: { size: 25, type: WidthType.PERCENTAGE }, borders: { top: { style: BorderStyle.NONE, size: 0 }, bottom: { style: BorderStyle.NONE, size: 0 }, left: { style: BorderStyle.NONE, size: 0 }, right: { style: BorderStyle.NONE, size: 0 } } }),
                            new TableCell({ children: [new Paragraph({ text: check("Others:"), ...fieldFont })], width: { size: 25, type: WidthType.PERCENTAGE }, borders: { top: { style: BorderStyle.NONE, size: 0 }, bottom: { style: BorderStyle.NONE, size: 0 }, left: { style: BorderStyle.NONE, size: 0 } } }),
                        ]
                    }),
                    new TableRow({
                        children: [
                            new TableCell({ children: [new Paragraph({ text: check("Skills: Exercise / Drill"), ...fieldFont })], width: { size: 25, type: WidthType.PERCENTAGE }, borders: { top: { style: BorderStyle.NONE, size: 0 }, right: { style: BorderStyle.NONE, size: 0 } } }),
                            new TableCell({ children: [new Paragraph({ text: check("Illustration"), ...fieldFont })], width: { size: 25, type: WidthType.PERCENTAGE }, borders: { top: { style: BorderStyle.NONE, size: 0 }, left: { style: BorderStyle.NONE, size: 0 }, right: { style: BorderStyle.NONE, size: 0 } } }),
                            new TableCell({ children: [new Paragraph({ text: check("Informal Theme"), ...fieldFont })], width: { size: 25, type: WidthType.PERCENTAGE }, borders: { top: { style: BorderStyle.NONE, size: 0 }, left: { style: BorderStyle.NONE, size: 0 }, right: { style: BorderStyle.NONE, size: 0 } } }),
                            new TableCell({ children: [new Paragraph({ text: "____________________", ...fieldFont })], width: { size: 25, type: WidthType.PERCENTAGE }, borders: { top: { style: BorderStyle.NONE, size: 0 }, left: { style: BorderStyle.NONE, size: 0 } } }),
                        ]
                    })
                ]
            });
            sections.push(activityTypeTable);

            // Details Table (Title, Target, Ref)
            const detailsTable = new Table({
                width: { size: 100, type: WidthType.PERCENTAGE },
                rows: [
                    new TableRow({ children: [new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "Activity Title: ", ...headerFont })] })], width: { size: 25, type: WidthType.PERCENTAGE } }), new TableCell({ children: [new Paragraph({ text: this.safeString(dayData.activityTitle), ...fieldFont })], width: { size: 75, type: WidthType.PERCENTAGE } })] }),
                    new TableRow({ children: [new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "Learning Target: ", ...headerFont })] })], width: { size: 25, type: WidthType.PERCENTAGE } }), new TableCell({ children: [new Paragraph({ text: this.safeString(dayData.learningTarget), ...fieldFont })], width: { size: 75, type: WidthType.PERCENTAGE } })] }),
                    new TableRow({ children: [new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "References: ", ...headerFont }), new TextRun({ text: "(Author, Title, Pages)", size: 16, italics: true })] })], width: { size: 25, type: WidthType.PERCENTAGE } }), new TableCell({ children: [new Paragraph({ text: this.safeString(dayData.references), ...fieldFont })], width: { size: 75, type: WidthType.PERCENTAGE } })] }),
                ]
            });
            sections.push(detailsTable);
            
            // Content Spacer
            sections.push(new Paragraph({ text: "", spacing: { after: 240 } }));

            // --- CONTENT BODY ---
            if (dayData.conceptNotes && Array.isArray(dayData.conceptNotes)) {
                dayData.conceptNotes.forEach(note => {
                    sections.push(new Paragraph({
                        children: [new TextRun({ text: this.safeString(note.title).toUpperCase(), ...headerFont })],
                        spacing: { before: 120, after: 120 }
                    }));
                    sections.push(...this.parseLasMarkdown(this.safeString(note.content)));
                });
            }

             if (dayData.activities && Array.isArray(dayData.activities)) {
                dayData.activities.forEach(activity => {
                    sections.push(new Paragraph({
                        children: [new TextRun({ text: this.safeString(activity.title).toUpperCase(), ...headerFont })],
                        spacing: { before: 240, after: 120 }
                    }));

                    const instructions = this.safeString(activity.instructions);
                    
                    // Handle "Matching" type activities if they use '||' separator
                    if (instructions.includes('||')) {
                        const lines = instructions.split('\n');
                        // separate intro lines from table lines
                        const directionsLines = lines.filter(l => !l.includes('||'));
                        const tableLines = lines.filter(l => l.includes('||'));

                        if (directionsLines.length > 0) {
                            sections.push(new Paragraph({
                                children: [
                                    new TextRun({ text: "Directions: ", ...headerFont }),
                                    new TextRun({ text: directionsLines.join(' '), ...contentFont })
                                ],
                                spacing: { after: 120 }
                            }));
                        }

                        if (tableLines.length > 0) {
                             const tableRows = tableLines.map(line => {
                                const parts = line.split('||');
                                const col1 = this.safeString(parts[0] || "").trim();
                                const col2 = this.safeString(parts[1] || "").trim();
                                return new TableRow({
                                    children: [
                                        new TableCell({ children: this.parseLasMarkdown(col1), width: { size: 50, type: WidthType.PERCENTAGE }, margins: { top: 100, bottom: 100, left: 100, right: 100 } }),
                                        new TableCell({ children: this.parseLasMarkdown(col2), width: { size: 50, type: WidthType.PERCENTAGE }, margins: { top: 100, bottom: 100, left: 100, right: 100 } }),
                                    ]
                                });
                            });

                            // Add Header Row for Table
                            tableRows.unshift(new TableRow({
                                tableHeader: true,
                                children: [
                                    new TableCell({ 
                                        children: [new Paragraph({ children: [new TextRun({ text: "Column A", ...headerFont })] })], 
                                        width: { size: 50, type: WidthType.PERCENTAGE },
                                        shading: { fill: "FFFFFF" }
                                    }),
                                    new TableCell({ 
                                        children: [new Paragraph({ children: [new TextRun({ text: "Column B", ...headerFont })] })], 
                                        width: { size: 50, type: WidthType.PERCENTAGE },
                                        shading: { fill: "FFFFFF" }
                                    }),
                                ]
                            }));
                            
                            sections.push(new Table({
                                rows: tableRows,
                                width: { size: 100, type: WidthType.PERCENTAGE },
                            }));
                            sections.push(new Paragraph({ text: "", spacing: { after: 120 } }));
                        }
                    } else {
                         sections.push(...this.parseLasMarkdown(instructions));
                    }

                     if (activity.questions && activity.questions.length > 0) {
                        activity.questions.forEach((q, i) => {
                            sections.push(new Paragraph({
                                children: [
                                    new TextRun({ text: `${i + 1}. `, ...headerFont }),
                                    new TextRun({ text: this.safeString(q.questionText), ...contentFont })
                                ],
                                spacing: { before: 60 }
                            }));
                            if (q.options && q.options.length > 0) {
                                q.options.forEach((opt, oi) => {
                                    sections.push(new Paragraph({
                                        text: `${String.fromCharCode(65+oi)}. ${this.safeString(opt)}`,
                                        indent: { left: 720 },
                                        style: "contentFont"
                                    }));
                                });
                            } else {
                                sections.push(new Paragraph({ text: "________________________________________", indent: { left: 720 }, style: "contentFont" }));
                            }
                        });
                    }
                });
             }

             // Reflection
             sections.push(new Paragraph({
                children: [new TextRun({ text: "REFLECTION", ...headerFont })],
                spacing: { before: 240, after: 120 }
            }));
            sections.push(new Paragraph({
                children: [new TextRun({ text: this.safeString(dayData.reflection), ...contentFont })],
            }));
            sections.push(new Paragraph({
                children: [new TextRun({ text: "________________________________________________________________________________________________________________________________", ...contentFont })],
                spacing: { before: 200 },
            }));
             sections.push(new Paragraph({
                children: [new TextRun({ text: "________________________________________________________________________________________________________________________________", ...contentFont })],
                spacing: { before: 200 },
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
                            style: {
                                paragraph: {
                                    indent: { left: 720, hanging: 360 },
                                },
                                run: { font: "Century Gothic", size: 28 }
                            }
                        }],
                    },
                ],
            },
            styles: {
                paragraphStyles: [
                    {
                        id: "contentFont",
                        name: "Content Font",
                        run: { font: "Century Gothic", size: 28 },
                    }
                ]
            },
            sections: [{
                properties: {
                    page: {
                        size: { width: 12240, height: 18720 }, // 8.5 x 13 Long Bond
                        margin: { top: 720, right: 720, bottom: 720, left: 720 }, // 0.5 inch
                    }
                },
                children: sections
            }]
        });

        const blob = await Packer.toBlob(doc);
        this.downloadBlob(blob, `LAS_${this.safeString(lasForm.subject).replace(/\s/g, '_')}.docx`);
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
                            style: {
                                paragraph: {
                                    indent: { left: 720, hanging: 360 },
                                }
                            },
                        },
                    ],
                },
            ],
        };

        const sections: (Paragraph | Table | Paragraph)[] = [
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
                    children: ["Objective", "Cognitive Level", "Item Numbers"].map(text => new TableCell({ children: [new Paragraph({ children: [new TextRun({ text, bold: true })] })] })),
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
                sections.push(new Paragraph({ children: [new TextRun({ text: this.safeString(section.instructions), italics: true })], spacing: { after: 200 } }));
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
                                indent: { left: 1080 },
                            }));
                        });
                    } else if (type === 'Identification' || type === 'True or False') {
                         sections.push(new Paragraph({ text: "Answer: ____________________", indent: { left: 1080 } }));
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
                sections.push(new Paragraph({ children: [new TextRun({ text: this.safeString(activity.activityInstructions), italics: true })], spacing: { after: 200 } }));
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
                                    new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: 'Criteria', bold: true })] })] }),
                                    new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: 'Points', bold: true })] })] }),
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
        
        sections.push(new Paragraph({ children: [new PageBreak()] }));
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
        
        const headerPars = [
            { label: isFilipino ? 'Paaralan' : 'School', value: this.safeString(dlpForm.schoolName).toUpperCase() },
            { value: this.safeString(dlpForm.quarterSelect) },
            { label: isFilipino ? 'Guro' : 'Teacher', value: this.safeString(dlpForm.teacher) },
            { label: isFilipino ? 'Asignatura' : 'Learning Area', value: this.safeString(dlpForm.subject).toUpperCase() },
            { label: isFilipino ? 'Petsa ng Pagtuturo' : 'Teaching Dates', value: this.safeString(dlpForm.teachingDates) },
        ];
    
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
                                style: {
                                    paragraph: {
                                        indent: { left: 720, hanging: 360 },
                                    },
                                    run: { font: "Times New Roman", size: 22 },
                                },
                            },
                        ],
                    },
                ],
            },
            styles: {
                paragraphStyles: [
                    { id: 'header-label', name: 'Header Label', run: { font: 'Arial', size: 16 }, paragraph: { spacing: { after: 0 } } },
                    { id: 'header-value', name: 'Header Value', run: { font: 'Arial', size: 16, bold: true }, paragraph: { spacing: { after: 40 } } }
                ]
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
                                        children: headerPars.map(p => {
                                            if (p.label) {
                                                return new Paragraph({ children: [new TextRun({ text: `${p.label}: `, size: 16 }), new TextRun({ text: p.value, bold: true, size: 16 })], alignment: AlignmentType.LEFT });
                                            } else {
                                                return new Paragraph({ children: [new TextRun({ text: p.value, bold: true, size: 16 })], alignment: AlignmentType.LEFT });
                                            }
                                        }),
                                        width: { size: 55, type: WidthType.PERCENTAGE }
                                    }),
                                    new TableCell({
                                        children: [
                                            new Paragraph({ children: [new TextRun({ text: isFilipino ? 'DETALYADONG BANGHAY-ARALIN SA' : 'DAILY LESSON PLAN IN', bold: true })], alignment: AlignmentType.CENTER }),
                                            new Paragraph({ children: [new TextRun({ text: `${this.safeString(dlpForm.subject).toUpperCase()} ${this.safeString(dlpForm.gradeLevel)}`, bold: true })], alignment: AlignmentType.CENTER }),
                                            new Paragraph({ children: [new TextRun({ text: (isFilipino ? 'ISKEDYUL NG KLASE' : 'CLASS SCHEDULE'), bold: true })], alignment: AlignmentType.CENTER, spacing: {before: 100} }),
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
                                    rows: [
                                        new TableRow({ children: [ new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: (isFilipino ? 'Pamamaraan' : 'Procedure'), bold: true })] })], width: { size: 25, type: WidthType.PERCENTAGE } }), new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: (isFilipino ? 'Gawain ng Guro/Mag-aaral' : 'Teacher/Student Activity'), bold: true })] })], width: { size: 45, type: WidthType.PERCENTAGE } }), new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: (isFilipino ? 'Mga Kaugnay na PPST Indicator' : 'Aligned PPST Indicators'), bold: true })] })], width: { size: 30, type: WidthType.PERCENTAGE } })] }),
                                        ...dlpContent.procedures.map(proc => new TableRow({
                                            children: [ new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: this.safeString(proc.title), bold: true })] })], width: { size: 25, type: WidthType.PERCENTAGE } }), new TableCell({ children: this.parseMarkdownToParagraphs(this.safeString(proc.content)), width: { size: 45, type: WidthType.PERCENTAGE } }), new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: this.safeString(proc.ppst), italics: true })] })], width: { size: 30, type: WidthType.PERCENTAGE } })],
                                        }))
                                    ],
                                })],
                            })] }),
                            new TableRow({ children: [createHeaderCell(t.evaluation), createContentCell( (dlpContent.evaluationQuestions || []).map(q => new Paragraph({ text: this.safeString(q.question), numbering: { reference: "dlp-list", level: 0 }, spacing: { after: 100 } })))] }),
                            new TableRow({ children: [createHeaderCell(t.reflection), createContentCell([new Paragraph({text: ""})])] }), // Empty for now
                        ],
                    }),
    
                    new Paragraph({ children: [new PageBreak()] }),
    
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
         const pageHeight = 18720;
        const pageWidth = 12240;

        const isFilipino = dllForm.language === 'Filipino';
        const days = isFilipino ? ['Lunes', 'Martes', 'Miyerkules', 'Huwebes', 'Biyernes'] : ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday'];
        
        const boldCentered = { children: [], alignment: AlignmentType.CENTER };

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
            children: [new TableCell({ children: [] }), ...days.map(day => new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: day, bold: true })], alignment: AlignmentType.CENTER })] }))]
        }));

        // Add Objectives
        mainTableRows.push(new TableRow({ children: [new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "I. OBJECTIVES", bold: true })] })], columnSpan: 6 })]}));
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
        mainTableRows.push(new TableRow({ children: [new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "II. CONTENT", bold: true })] })] }), new TableCell({ children: [new Paragraph(this.safeString(dllContent.content))], columnSpan: 5 })]}));

        // Add Learning Resources
        mainTableRows.push(new TableRow({ children: [new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "III. LEARNING RESOURCES", bold: true })] })], columnSpan: 6 })]}));

        // Add Procedures
        mainTableRows.push(new TableRow({ children: [new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "IV. PROCEDURES", bold: true })] })], columnSpan: 6 })]}));
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
        mainTableRows.push(new TableRow({ children: [new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "V. REMARKS", bold: true })] })] }), new TableCell({ children: [new Paragraph(this.safeString(dllContent.remarks))], columnSpan: 5 })]}));
        mainTableRows.push(new TableRow({ children: [new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "VI. REFLECTION", bold: true })] })], columnSpan: 6 })]}));

        const mainTable = new Table({
            width: { size: 100, type: WidthType.PERCENTAGE },
            rows: mainTableRows
        });
        
        const signatoriesTable = new Table({
            width: { size: 100, type: WidthType.PERCENTAGE },
            rows: [
                 new TableRow({
                    children: [
                        new TableCell({ children: [new Paragraph("Prepared by:")], width: {size: 33, type: WidthType.PERCENTAGE}, borders: { top: { style: BorderStyle.NONE }, bottom: { style: BorderStyle.NONE }, left: { style: BorderStyle.NONE }, right: { style: BorderStyle.NONE }} }),
                        new TableCell({ children: [new Paragraph("Checked by:")], width: {size: 34, type: WidthType.PERCENTAGE}, borders: { top: { style: BorderStyle.NONE }, bottom: { style: BorderStyle.NONE }, left: { style: BorderStyle.NONE }, right: { style: BorderStyle.NONE }} }),
                        new TableCell({ children: [new Paragraph("Approved by:")], width: {size: 33, type: WidthType.PERCENTAGE}, borders: { top: { style: BorderStyle.NONE }, bottom: { style: BorderStyle.NONE }, left: { style: BorderStyle.NONE }, right: { style: BorderStyle.NONE }} }),
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
                        new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: this.safeString(dllForm.preparedByName), bold: true })], alignment: AlignmentType.CENTER })], borders: { top: { style: BorderStyle.SINGLE }, bottom: { style: BorderStyle.NONE }, left: { style: BorderStyle.NONE }, right: { style: BorderStyle.NONE }} }),
                        new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: this.safeString(dllForm.checkedByName), bold: true })], alignment: AlignmentType.CENTER })], borders: { top: { style: BorderStyle.SINGLE }, bottom: { style: BorderStyle.NONE }, left: { style: BorderStyle.NONE }, right: { style: BorderStyle.NONE }} }),
                        new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: this.safeString(dllForm.approvedByName), bold: true })], alignment: AlignmentType.CENTER })], borders: { top: { style: BorderStyle.SINGLE }, bottom: { style: BorderStyle.NONE }, left: { style: BorderStyle.NONE }, right: { style: BorderStyle.NONE }} }),
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
                        size: { width: pageHeight, height: pageWidth, orientation: PageOrientation.LANDSCAPE },
                        margin: { top: 720, right: 720, bottom: 720, left: 720 },
                    },
                },
                children: [
                    new Paragraph({ children: [new TextRun({ text: "DAILY LESSON LOG", bold: true })], alignment: AlignmentType.CENTER, spacing: { after: 200 } }),
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

    public async generateExamDocx(
        exam: GeneratedExam,
        settings: SchoolSettings
    ): Promise<void> {
         const sections: (Paragraph | Table | Paragraph)[] = [];

        // Header
        const secondLogoRun = this.createDocxImage(this.parseDataUrl(settings.secondLogo), 60, 60) || new TextRun("");
        const schoolLogoRun = this.createDocxImage(this.parseDataUrl(settings.schoolLogo), 60, 60) || new TextRun("");

        const headerTable = new Table({
            width: { size: 100, type: WidthType.PERCENTAGE },
            rows: [
                new TableRow({
                    children: [
                        new TableCell({ children: [new Paragraph({ children: [secondLogoRun, new TextRun("  "), schoolLogoRun], alignment: AlignmentType.LEFT })], borders: { top: { style: BorderStyle.NONE }, bottom: { style: BorderStyle.NONE }, left: { style: BorderStyle.NONE }, right: { style: BorderStyle.NONE } }, width: { size: 20, type: WidthType.PERCENTAGE } }),
                        new TableCell({ children: [new Paragraph({ text: `Republic of the Philippines`, alignment: AlignmentType.CENTER }), new Paragraph({ children: [new TextRun({ text: `Department of Education`, bold: true })], alignment: AlignmentType.CENTER }), new Paragraph({ children: [new TextRun({ text: settings.schoolName.toUpperCase(), bold: true })], alignment: AlignmentType.CENTER })], borders: { top: { style: BorderStyle.NONE }, bottom: { style: BorderStyle.NONE }, left: { style: BorderStyle.NONE }, right: { style: BorderStyle.NONE } }, width: { size: 60, type: WidthType.PERCENTAGE } }),
                         new TableCell({ children: [], borders: { top: { style: BorderStyle.NONE }, bottom: { style: BorderStyle.NONE }, left: { style: BorderStyle.NONE }, right: { style: BorderStyle.NONE } }, width: { size: 20, type: WidthType.PERCENTAGE } }),
                    ]
                })
            ]
        });
        sections.push(headerTable);
        sections.push(new Paragraph({ text: exam.title.toUpperCase(), heading: HeadingLevel.TITLE, alignment: AlignmentType.CENTER, spacing: { before: 200, after: 200 } }));
        sections.push(new Paragraph({ text: `Subject: ${exam.subject} | Grade: ${exam.gradeLevel} | Quarter: ${exam.quarter}`, alignment: AlignmentType.CENTER, spacing: { after: 200 } }));
        sections.push(new Paragraph({ text: `Name: __________________________________  Score: ______`, spacing: { after: 400 } }));

        // Questions
        sections.push(new Paragraph({ children: [new TextRun({ text: "TEST I. MULTIPLE CHOICE", bold: true })] }));
        sections.push(new Paragraph({ children: [new TextRun({ text: "Directions: Read each question carefully and circle the letter of the correct answer.", italics: true })], spacing: { after: 200 } }));

        exam.questions.forEach((q, i) => {
             sections.push(new Paragraph({
                children: [
                    new TextRun({ text: `${i + 1}. ${q.questionText}`, bold: true })
                ],
                spacing: { before: 100 }
            }));
            
            if (q.options) {
                q.options.forEach((opt, oi) => {
                    sections.push(new Paragraph({
                        text: `${String.fromCharCode(65+oi)}. ${opt}`,
                        indent: { left: 720 }
                    }));
                });
            }
        });

        // TOS (New Page)
        sections.push(new Paragraph({ children: [new PageBreak()] }));
        sections.push(new Paragraph({ text: "TABLE OF SPECIFICATIONS", heading: HeadingLevel.HEADING_2, alignment: AlignmentType.CENTER }));
        
        const tosHeaderRow = new TableRow({
            children: ["Competencies", "Days", "%", "No. Items", "Placement", "R", "U", "Ap", "An", "E", "C"].map(t => new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: t, bold: true, size: 16 })] })], shading: { fill: "E7E6E6" } })),
            tableHeader: true
        });

        const tosRows = exam.tableOfSpecifications.map(row => new TableRow({
            children: [
                new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: row.objective, size: 16 })] })] }),
                new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: String(row.daysTaught), size: 16 })] })] }),
                new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: row.percentage, size: 16 })] })] }),
                new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: String(row.numItems), size: 16 })] })] }),
                new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: row.itemPlacement, size: 16 })] })] }),
                new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: row.remembering, size: 16 })] })] }),
                new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: row.understanding, size: 16 })] })] }),
                new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: row.applying, size: 16 })] })] }),
                new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: row.analyzing, size: 16 })] })] }),
                new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: row.evaluating, size: 16 })] })] }),
                new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: row.creating, size: 16 })] })] }),
            ]
        }));

        sections.push(new Table({
            rows: [tosHeaderRow, ...tosRows],
            width: { size: 100, type: WidthType.PERCENTAGE }
        }));
        
        // Key (New Page)
        sections.push(new Paragraph({ children: [new PageBreak()] }));
        sections.push(new Paragraph({ text: "ANSWER KEY", heading: HeadingLevel.HEADING_2, alignment: AlignmentType.CENTER }));
        exam.questions.forEach((q, i) => {
            sections.push(new Paragraph({ text: `${i+1}. ${q.correctAnswer}` }));
        });

        const doc = new Document({
            sections: [{ children: sections }]
        });
        
        const blob = await Packer.toBlob(doc);
        this.downloadBlob(blob, `Exam_${exam.subject.replace(/\s/g, '_')}_Q${exam.quarter}.docx`);
    }
    
    public async generatePickedStudentsDocx(data: PickedStudentsDocxData): Promise<void> {
         const doc = new Document({
            sections: [{
                children: [
                    new Paragraph({ text: `Recitation List: ${data.topic}`, heading: HeadingLevel.TITLE }),
                    new Paragraph({ text: data.sectionText }),
                    new Paragraph({ text: `Date: ${new Date().toLocaleDateString()}` }),
                    ...data.pickedStudents.map((s, i) => new Paragraph({ text: `${i + 1}. ${s.firstName} ${s.lastName}` }))
                ]
            }]
        });
        const blob = await Packer.toBlob(doc);
        this.downloadBlob(blob, `Recitation_List_${data.topic}.docx`);
    }
    
    public async generateGroupsDocx(data: GroupsDocxData): Promise<void> {
        const sections: (Paragraph)[] = [
            new Paragraph({ text: `Group Assignments: ${data.topic}`, heading: HeadingLevel.TITLE }),
            new Paragraph({ text: data.sectionText }),
            new Paragraph({ text: `Date: ${new Date().toLocaleDateString()}` }),
        ];
        
        data.groups.forEach((group, i) => {
            sections.push(new Paragraph({ text: `Group ${i + 1}`, heading: HeadingLevel.HEADING_2, spacing: { before: 200 } }));
            group.forEach(s => sections.push(new Paragraph({ text: `${s.firstName} ${s.lastName}`, bullet: { level: 0 } })));
        });

        const doc = new Document({ sections: [{ children: sections }] });
        const blob = await Packer.toBlob(doc);
        this.downloadBlob(blob, `Groups_${data.topic}.docx`);
    }
    
    public async generateHonorsListDocx(data: HonorsListDocxData): Promise<void> {
         const sections: (Paragraph | Table)[] = [
            new Paragraph({ text: "HONOR ROLL", heading: HeadingLevel.TITLE, alignment: AlignmentType.CENTER }),
            new Paragraph({ text: `School Year: ${data.settings.schoolYear}`, alignment: AlignmentType.CENTER }),
            new Paragraph({ text: `Class: ${data.selectedSectionText}`, alignment: AlignmentType.CENTER, spacing: { after: 200 } }),
        ];

        const addCategory = (title: string, students: any[]) => {
            if (students.length > 0) {
                sections.push(new Paragraph({ text: title, heading: HeadingLevel.HEADING_2, spacing: { before: 200, after: 100 } }));
                const rows = students.map((s, i) => new TableRow({
                    children: [
                        new TableCell({ children: [new Paragraph(String(i + 1))] }),
                        new TableCell({ children: [new Paragraph(`${s.student.lastName}, ${s.student.firstName}`)] }),
                        new TableCell({ children: [new Paragraph(s.generalAvg.toFixed(0))] }),
                    ]
                }));
                 sections.push(new Table({
                    rows: [
                        new TableRow({ children: ["Rank", "Name", "Average"].map(t => new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: t, bold: true })] })] })), tableHeader: true }),
                        ...rows
                    ],
                    width: { size: 100, type: WidthType.PERCENTAGE }
                }));
            }
        };

        addCategory("WITH HIGHEST HONORS (98-100)", data.honorStudents.highest);
        addCategory("WITH HIGH HONORS (95-97)", data.honorStudents.high);
        addCategory("WITH HONORS (90-94)", data.honorStudents.regular);

        const doc = new Document({ sections: [{ children: sections }] });
        const blob = await Packer.toBlob(doc);
        this.downloadBlob(blob, `Honor_Roll_${data.selectedSectionText}.docx`);
    }
    
    public async generateStudentProfileDocx(data: StudentProfileDocxData): Promise<void> {
         const sections: (Paragraph | Table)[] = []; // Removed ImageRun
         // Header
         sections.push(new Paragraph({ text: data.settings.schoolName.toUpperCase(), heading: HeadingLevel.TITLE, alignment: AlignmentType.CENTER }));
         sections.push(new Paragraph({ text: "STUDENT PROFILE", heading: HeadingLevel.HEADING_2, alignment: AlignmentType.CENTER }));
         
         sections.push(new Paragraph({ children: [new TextRun({ text: `Name: ${data.student.firstName} ${data.student.lastName}`, bold: true })], spacing: { before: 200 } }));
         sections.push(new Paragraph({ text: `LRN: ${data.student.lrn || 'N/A'}   |   Grade/Section: ${data.student.gradeLevel} - ${data.student.section}` }));

         // Grades
         sections.push(new Paragraph({ text: "Academic Record", heading: HeadingLevel.HEADING_3, spacing: { before: 200 } }));
         const gradeRows = data.academicSummary.map(sub => new TableRow({
             children: [
                 new TableCell({ children: [new Paragraph(sub.subject)] }),
                 ...sub.grades.map(g => new TableCell({ children: [new Paragraph(g !== null ? String(g) : '-')],  })),
             ]
         }));
         sections.push(new Table({
             rows: [
                 new TableRow({ children: ["Subject", "Q1", "Q2", "Q3", "Q4"].map(t => new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: t, bold: true })] })] })), tableHeader: true }),
                 ...gradeRows
             ],
             width: { size: 100, type: WidthType.PERCENTAGE }
         }));
         
         // Anecdotes
         if (data.recentAnecdotes.length > 0) {
             sections.push(new Paragraph({ text: "Recent Observations", heading: HeadingLevel.HEADING_3, spacing: { before: 200 } }));
             data.recentAnecdotes.forEach(a => {
                 sections.push(new Paragraph({ children: [new TextRun({ text: new Date(a.date).toLocaleDateString(), bold: true })] }));
                 sections.push(new Paragraph({ children: [new TextRun({ text: a.observation, italics: true })], spacing: { after: 100 } }));
             });
         }

         const doc = new Document({ sections: [{ children: sections }] });
         const blob = await Packer.toBlob(doc);
         this.downloadBlob(blob, `Profile_${data.student.lastName}.docx`);
    }

    public async generateAttendanceDocx(
        students: Student[], 
        attendance: Attendance[], 
        currentDate: Date, 
        settings: SchoolSettings
    ): Promise<void> {
        // Simplified implementation similar to generateAttendanceXlsx but for Docx
         const doc = new Document({
            sections: [{
                children: [new Paragraph("Attendance Report (See Excel export for detailed view)")]
            }]
        });
        const blob = await Packer.toBlob(doc);
        this.downloadBlob(blob, `Attendance_Report.docx`);
    }

    public async generateSF2Docx(
        students: Student[],
        attendance: Attendance[],
        settings: SchoolSettings,
        currentDate: Date
    ): Promise<void> {
         const doc = new Document({
            sections: [{
                properties: { page: { size: { width: 12240, height: 15840 } } }, // Legal sizeish
                children: [new Paragraph("School Form 2 (SF2) - Automated Generation Pending Update")]
            }]
        });
        const blob = await Packer.toBlob(doc);
        this.downloadBlob(blob, `SF2_${currentDate.getMonth()}.docx`);
    }

    public async generateMapehRecordDocx(data: MapehRecordDocxData): Promise<void> {
          const doc = new Document({
            sections: [{
                children: [new Paragraph("MAPEH Record")]
            }]
        });
        const blob = await Packer.toBlob(doc);
        this.downloadBlob(blob, `MAPEH_Record.docx`);
    }
    
    public async generateSummaryOfGradesDocx(data: SummaryOfGradesDocxData): Promise<void> {
         const doc = new Document({
            sections: [{
                children: [new Paragraph("Summary of Grades")]
            }]
        });
        const blob = await Packer.toBlob(doc);
        this.downloadBlob(blob, `Summary_of_Grades.docx`);
    }
    
    public async generateEClassRecordDocx(data: EClassRecordDocxData): Promise<void> {
        const { allStudents, settings, subject, quarter, selectedSectionText, recordSettings, studentRecords, calculationResults } = data;

        const rows: TableRow[] = [];

        // Header Rows
        const headerRow = new TableRow({
            children: [
                new TableCell({ children: [new Paragraph("Learner's Names")], rowSpan: 2 }),
                new TableCell({ children: [new Paragraph(`Written Works (${recordSettings.wwPercentage * 100}%)`)], columnSpan: 3 }), // Simplified
                new TableCell({ children: [new Paragraph(`Performance Tasks (${recordSettings.ptPercentage * 100}%)`)], columnSpan: 3 }), // Simplified
                new TableCell({ children: [new Paragraph(`Quarterly Assessment (${recordSettings.qaPercentage * 100}%)`)], columnSpan: 3 }), // Simplified
                new TableCell({ children: [new Paragraph("Initial Grade")], rowSpan: 2 }),
                new TableCell({ children: [new Paragraph("Quarterly Grade")], rowSpan: 2 }),
            ]
        });
        rows.push(headerRow);
        
        const subHeaderRow = new TableRow({
            children: [
                new TableCell({ children: [new Paragraph("Total")] }), new TableCell({ children: [new Paragraph("PS")] }), new TableCell({ children: [new Paragraph("WS")] }),
                new TableCell({ children: [new Paragraph("Total")] }), new TableCell({ children: [new Paragraph("PS")] }), new TableCell({ children: [new Paragraph("WS")] }),
                new TableCell({ children: [new Paragraph("Score")] }), new TableCell({ children: [new Paragraph("PS")] }), new TableCell({ children: [new Paragraph("WS")] }),
            ]
        });
        rows.push(subHeaderRow);

        const addStudentRows = (studentList: Student[]) => {
            studentList.forEach(student => {
                const calcs = calculationResults.get(student.id) || {};
                rows.push(new TableRow({
                    children: [
                        new TableCell({ children: [new Paragraph(`${student.lastName}, ${student.firstName}`)] }),
                        new TableCell({ children: [new Paragraph(String(calcs.wwTotal || 0))] }), new TableCell({ children: [new Paragraph(calcs.wwPs?.toFixed(2) || "")] }), new TableCell({ children: [new Paragraph(calcs.wwWs?.toFixed(2) || "")] }),
                        new TableCell({ children: [new Paragraph(String(calcs.ptTotal || 0))] }), new TableCell({ children: [new Paragraph(calcs.ptPs?.toFixed(2) || "")] }), new TableCell({ children: [new Paragraph(calcs.ptWs?.toFixed(2) || "")] }),
                        new TableCell({ children: [new Paragraph(String(data.studentRecords.find(r => r.studentId === student.id)?.quarterlyAssessment || 0))] }), new TableCell({ children: [new Paragraph(calcs.qaPs?.toFixed(2) || "")] }), new TableCell({ children: [new Paragraph(calcs.qaWs?.toFixed(2) || "")] }),
                        new TableCell({ children: [new Paragraph(calcs.initialGrade?.toFixed(2) || "")] }),
                        new TableCell({ children: [new Paragraph(String(calcs.quarterlyGrade || ""))] }),
                    ]
                }));
            });
        };

        if (allStudents.males.length > 0) {
            rows.push(new TableRow({ children: [new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "MALE", bold: true })] })], columnSpan: 12 })] }));
            addStudentRows(allStudents.males);
        }
        if (allStudents.females.length > 0) {
             rows.push(new TableRow({ children: [new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "FEMALE", bold: true })] })], columnSpan: 12 })] }));
            addStudentRows(allStudents.females);
        }

        const table = new Table({
            width: { size: 100, type: WidthType.PERCENTAGE },
            rows: rows
        });

        const doc = new Document({
            sections: [{
                properties: {
                    page: {
                         size: { width: 12240, height: 15840, orientation: PageOrientation.LANDSCAPE } // Legal Landscape approx
                    }
                },
                children: [
                    new Paragraph({ text: `E-Class Record: ${subject} - Q${quarter}`, heading: HeadingLevel.TITLE }),
                    new Paragraph({ text: `Class: ${selectedSectionText}` }),
                    new Paragraph({ text: `Teacher: ${settings.teacherName}`, spacing: { after: 200 } }),
                    table
                ]
            }]
        });

        const blob = await Packer.toBlob(doc);
        this.downloadBlob(blob, `EClassRecord_${subject}_Q${quarter}.docx`);
    }
}

export const docxService = new DocxService();
