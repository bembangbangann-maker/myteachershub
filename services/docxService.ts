import {
    Document,
    Packer,
    Paragraph,
    TextRun,
    Table,
    TableRow,
    TableCell,
    WidthType,
    AlignmentType,
    BorderStyle,
    VerticalAlign,
    LevelFormat,
    PageBreak,
    ImageRun,
    Header,
    Footer,
    convertInchesToTwip,
    UnderlineType
} from "docx";
import { saveAs } from "file-saver";
import {
    Student,
    SchoolSettings,
    Attendance,
    Quarter,
    SubjectQuarterSettings,
    StudentQuarterlyRecord,
    MapehRecordDocxData,
    SummaryOfGradesData,
    LearningActivitySheet,
    DlpContent,
    DllContent,
    GeneratedQuiz,
    GeneratedExam
} from "../types";

export class DocxService {
    // Helper to safely return string
    private safeString(str: string | undefined | null): string {
        return str || "";
    }

    // Helper to save blob
    private downloadBlob(blob: Blob, fileName: string) {
        saveAs(blob, fileName);
    }

    // Helper to parse data URL for images
    private parseDataUrl(dataUrl: string): string | Uint8Array {
        if (!dataUrl) return "";
        const base64Image = dataUrl.split(',')[1];
        if (!base64Image) return "";
        return Uint8Array.from(atob(base64Image), c => c.charCodeAt(0));
    }

    // Helper to create image run
    private createDocxImage(data: string | Uint8Array, width: number, height: number): ImageRun | undefined {
        if (!data) return undefined;
        try {
            return new ImageRun({
                data: data,
                transformation: {
                    width: width,
                    height: height,
                },
            });
        } catch (e) {
            console.error("Error creating image run", e);
            return undefined;
        }
    }

    // Helper to parse markdown-like syntax for LAS
    private parseLasMarkdown(text: string): Paragraph[] {
        const paragraphs: Paragraph[] = [];
        const lines = text.split('\n');
        
        lines.forEach(line => {
            // Basic markdown parsing for bold (**text**)
            const parts = line.split(/(\*\*.*?\*\*)/g);
            const children: TextRun[] = parts.map(part => {
                if (part.startsWith('**') && part.endsWith('**')) {
                    return new TextRun({
                        text: part.slice(2, -2),
                        bold: true,
                        font: "Century Gothic",
                        size: 24 // 12pt
                    });
                }
                return new TextRun({
                    text: part,
                    font: "Century Gothic",
                    size: 24 // 12pt
                });
            });
            
            paragraphs.push(new Paragraph({ children }));
        });
        
        return paragraphs;
    }

    // ... Methods called by components ...

    public async generateAttendanceDocx(
        students: Student[],
        attendance: Attendance[],
        currentDate: Date,
        settings: SchoolSettings
    ): Promise<void> {
        // Placeholder implementation - create a simple document
        const doc = new Document({
            sections: [{
                children: [
                    new Paragraph({ text: "Attendance Report", heading: "Heading1" }),
                    new Paragraph({ text: `Date: ${currentDate.toLocaleDateString()}` }),
                    new Paragraph({ text: `Class: ${students[0]?.gradeLevel || ''} - ${students[0]?.section || ''}` }),
                    new Paragraph({ text: "Attendance data table generated here." })
                ]
            }]
        });
        const blob = await Packer.toBlob(doc);
        this.downloadBlob(blob, "Attendance_Report.docx");
    }

    public async generateSF2Docx(
        students: Student[],
        attendance: Attendance[],
        settings: SchoolSettings,
        currentDate: Date
    ): Promise<void> {
        // Placeholder
        const doc = new Document({
            sections: [{
                children: [
                    new Paragraph({ text: "School Form 2 (SF2)", heading: "Heading1" }),
                    new Paragraph({ text: "Daily Attendance Report of Learners" }),
                ]
            }]
        });
        const blob = await Packer.toBlob(doc);
        this.downloadBlob(blob, "SF2.docx");
    }

    public async generateEClassRecordDocx(data: any): Promise<void> {
        const doc = new Document({
            sections: [{
                children: [
                    new Paragraph({ text: "E-Class Record", heading: "Heading1" }),
                    new Paragraph({ text: `Subject: ${data.subject}` }),
                ]
            }]
        });
        const blob = await Packer.toBlob(doc);
        this.downloadBlob(blob, "EClassRecord.docx");
    }

    public async generateMapehRecordDocx(data: MapehRecordDocxData): Promise<void> {
        const doc = new Document({
            sections: [{
                children: [
                    new Paragraph({ text: "MAPEH Record", heading: "Heading1" }),
                ]
            }]
        });
        const blob = await Packer.toBlob(doc);
        this.downloadBlob(blob, "MapehRecord.docx");
    }

    public async generateSummaryOfGradesDocx(data: any): Promise<void> {
        const doc = new Document({
            sections: [{
                children: [
                    new Paragraph({ text: "Summary of Grades", heading: "Heading1" }),
                ]
            }]
        });
        const blob = await Packer.toBlob(doc);
        this.downloadBlob(blob, "SummaryOfGrades.docx");
    }

    public async generateHonorsListDocx(data: any): Promise<void> {
        const doc = new Document({
            sections: [{
                children: [
                    new Paragraph({ text: "Honors List", heading: "Heading1" }),
                ]
            }]
        });
        const blob = await Packer.toBlob(doc);
        this.downloadBlob(blob, "HonorsList.docx");
    }

    public async generateStudentProfileDocx(data: any): Promise<void> {
        const doc = new Document({
            sections: [{
                children: [
                    new Paragraph({ text: "Student Profile", heading: "Heading1" }),
                    new Paragraph({ text: `Name: ${data.student.firstName} ${data.student.lastName}` }),
                ]
            }]
        });
        const blob = await Packer.toBlob(doc);
        this.downloadBlob(blob, "StudentProfile.docx");
    }

    public async generatePickedStudentsDocx(data: any): Promise<void> {
        const doc = new Document({
            sections: [{
                children: [
                    new Paragraph({ text: "Picked Students", heading: "Heading1" }),
                ]
            }]
        });
        const blob = await Packer.toBlob(doc);
        this.downloadBlob(blob, "PickedStudents.docx");
    }

    public async generateGroupsDocx(data: any): Promise<void> {
        const doc = new Document({
            sections: [{
                children: [
                    new Paragraph({ text: "Generated Groups", heading: "Heading1" }),
                ]
            }]
        });
        const blob = await Packer.toBlob(doc);
        this.downloadBlob(blob, "Groups.docx");
    }

    public async generateDlpDocx(form: any, content: DlpContent, extra: any, settings: SchoolSettings): Promise<void> {
        const doc = new Document({
            sections: [{
                children: [
                    new Paragraph({ text: "Daily Lesson Plan", heading: "Heading1" }),
                    new Paragraph({ text: `Topic: ${content.topic}` }),
                ]
            }]
        });
        const blob = await Packer.toBlob(doc);
        this.downloadBlob(blob, "DLP.docx");
    }

    public async generateDllDocx(form: any, content: DllContent, settings: SchoolSettings): Promise<void> {
        const doc = new Document({
            sections: [{
                children: [
                    new Paragraph({ text: "Daily Lesson Log", heading: "Heading1" }),
                ]
            }]
        });
        const blob = await Packer.toBlob(doc);
        this.downloadBlob(blob, "DLL.docx");
    }

    public async generateQuizDocx(content: GeneratedQuiz): Promise<void> {
        const doc = new Document({
            sections: [{
                children: [
                    new Paragraph({ text: content.quizTitle, heading: "Heading1" }),
                ]
            }]
        });
        const blob = await Packer.toBlob(doc);
        this.downloadBlob(blob, "Quiz.docx");
    }

    public async generateExamDocx(content: GeneratedExam, settings: SchoolSettings): Promise<void> {
        const doc = new Document({
            sections: [{
                children: [
                    new Paragraph({ text: content.title, heading: "Heading1" }),
                ]
            }]
        });
        const blob = await Packer.toBlob(doc);
        this.downloadBlob(blob, "Exam.docx");
    }

    public async generateLasDocx(
        lasForm: any,
        lasContent: LearningActivitySheet,
        settings: SchoolSettings
    ): Promise<void> {
        const sections: (Paragraph | Table)[] = []; 
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
                // Simple check logic - if content seems like performance task, check it
                const activityType = this.safeString(lasForm.activityType).toLowerCase();
                const titleLower = this.safeString(dayData.dayTitle).toLowerCase();
                
                if (label.includes("Concept Notes") && !titleLower.includes("performance task")) return `\u2611 ${label}`;
                if (label.includes("Performance Task") && titleLower.includes("performance task")) return `\u2611 ${label}`;
                
                return `\u2610 ${label}`;
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

            // Details Table (Title, Target, Ref) + COMPETENCY
            const detailsTable = new Table({
                width: { size: 100, type: WidthType.PERCENTAGE },
                rows: [
                    new TableRow({ children: [new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "Activity Title: ", ...headerFont })] })], width: { size: 25, type: WidthType.PERCENTAGE } }), new TableCell({ children: [new Paragraph({ text: this.safeString(dayData.activityTitle), ...fieldFont })], width: { size: 75, type: WidthType.PERCENTAGE } })] }),
                    new TableRow({ children: [new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "Learning Target: ", ...headerFont })] })], width: { size: 25, type: WidthType.PERCENTAGE } }), new TableCell({ children: [new Paragraph({ text: this.safeString(dayData.learningTarget), ...fieldFont })], width: { size: 75, type: WidthType.PERCENTAGE } })] }),
                    // Added Learning Competency Row as requested
                    new TableRow({ children: [new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "Learning Competency: ", ...headerFont })] })], width: { size: 25, type: WidthType.PERCENTAGE } }), new TableCell({ children: [new Paragraph({ text: this.safeString(lasForm.learningCompetency), ...fieldFont })], width: { size: 75, type: WidthType.PERCENTAGE } })] }),
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
                                        children: [new TextRun({ text: `${String.fromCharCode(65+oi)}. ${this.safeString(opt)}`, ...contentFont })],
                                        indent: { left: 720 }
                                    }));
                                });
                            } else {
                                sections.push(new Paragraph({ children: [new TextRun({ text: "________________________________________", ...contentFont })], indent: { left: 720 } }));
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
}

export const docxService = new DocxService();