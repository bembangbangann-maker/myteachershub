import React, { useState, useEffect, useMemo, useCallback } from 'react';
import { toast } from 'react-hot-toast';
import { useAppContext } from '../contexts/AppContext';
import { generateDlpContent, generateQuizContent, generateRubricForActivity, generateDllContent, generateLearningActivitySheet, generateExam } from '../services/geminiService';
import { DlpContent, GeneratedQuiz, QuizType, DlpRubricItem, GeneratedQuizSection, DllContent, DlpProcedure, LearningActivitySheet, SchoolSettings, ExamObjective, GeneratedExam, LearningActivitySheetDay } from '../types';
import Header from './Header';
import { SparklesIcon, DownloadIcon, ClipboardCheckIcon, PlusIcon, TrashIcon, RefreshCwIcon } from './icons';
import { docxService } from '../services/docxService';

const TabButton: React.FC<{ label: string, icon: React.ReactNode, isActive: boolean, onClick: () => void }> = ({ label, icon, isActive, onClick }) => (
    <button onClick={onClick} className={`flex items-center gap-2 px-4 py-3 text-sm font-semibold transition-colors border-b-2 ${isActive ? 'border-primary text-primary' : 'border-transparent text-base-content/70 hover:text-base-content'}`}>
        {icon} {label}
    </button>
);

const InputField: React.FC<{ id: string, label: string, value: string, onChange: any, type?: string, required?: boolean, placeholder?: string }> = ({ id, label, value, onChange, type = 'text', required = false, placeholder='' }) => (
    <div>
        <label htmlFor={id} className="block text-sm font-medium text-base-content mb-1">{label}{required && <span className="text-error">*</span>}</label>
        <input type={type} id={id} value={value} onChange={onChange} required={required} placeholder={placeholder} className="w-full bg-base-100 border border-base-300 rounded-md p-2 h-10 text-base-content" />
    </div>
);

const TextAreaField: React.FC<{ id: string, label: string, value: string, onChange: any, rows?: number, required?: boolean, placeholder?: string }> = ({ id, label, value, onChange, rows = 3, required = false, placeholder='' }) => (
    <div>
        <label htmlFor={id} className="block text-sm font-medium text-base-content mb-1">{label}{required && <span className="text-error">*</span>}</label>
        <textarea id={id} value={value} onChange={onChange} rows={rows} required={required} placeholder={placeholder} className="w-full bg-base-100 border border-base-300 rounded-md p-2 text-base-content" />
    </div>
);

const gradeLevels = ['Kindergarten', '1', '2', '3', '4', '5', '6', '7', '8', '9', '10', '11', '12'];
const quizTypesList: QuizType[] = ['Multiple Choice', 'True or False', 'Identification'];
const activityTypes = [
    "Concept Notes", "Skills: Exercise / Drill", "Performance Task", "Illustration / Drawing", "Formal Theme",
    "Informal Theme", "Guided Practice", "Independent Practice", "Group Activity", "Problem Solving",
    "Creative Output", "Inquiry-Based Learning", "Experiment / Investigation", "Others"
];

type ActiveTab = 'dlp' | 'dll' | 'quiz' | 'las' | 'exam';

const LessonPlanners: React.FC = () => {
    const { settings } = useAppContext();
    const [activeTab, setActiveTab] = useState<ActiveTab>('dlp');
    const [isLoading, setIsLoading] = useState(false);
    
    // DLP State
    const [teacherPosition, setTeacherPosition] = useState<'Beginning' | 'Proficient' | 'Highly Proficient' | 'Distinguished'>('Beginning');
    const [dlpForm, setDlpForm] = useState({
        teacher: settings.teacherName || '',
        schoolName: settings.schoolName || '',
        subject: 'English',
        teachingDates: '',
        classSchedule: '',
        gradeLevel: '9',
        quarterSelect: '1ST QUARTER',
        learningCompetency: '',
        lessonObjective: '',
        previousLesson: '',
        preparedByName: settings.teacherName.toUpperCase() || '',
        preparedByDesignation: 'Secondary School Teacher I, Grade 9\nENGLISH Teacher',
        checkedByName: (settings.checkedBy || '').toUpperCase(),
        checkedByDesignation: settings.checkerDesignation || 'Learning Area Coordinator',
        approvedByName: (settings.principalName || '').toUpperCase(),
        approvedByDesignation: settings.principalDesignation || 'School Principal II',
        language: 'English',
        dlpFormat: 'Standard DepEd',
    });
    const [dlpContent, setDlpContent] = useState<DlpContent | null>(null);

    // Quiz State
    const [quizForm, setQuizForm] = useState({
        quizTopic: '',
        numQuestions: 10,
        quizTypes: ['Multiple Choice'] as QuizType[],
        subject: 'English',
        gradeLevel: '9',
    });
    const [quizContent, setQuizContent] = useState<GeneratedQuiz | null>(null);
    const [activityPoints, setActivityPoints] = useState<{ [index: number]: string }>({});
    const [generatingRubricIndex, setGeneratingRubricIndex] = useState<number | null>(null);

    // DLL State
    const [dllFormat, setDllFormat] = useState('Standard');
    const [dllForm, setDllForm] = useState({
        subject: 'English',
        gradeLevel: '10',
        weeklyTopic: '',
        contentStandard: '',
        performanceStandard: '',
        teachingDates: '',
        quarter: '3',
        preparedByName: settings.teacherName.toUpperCase() || '',
        preparedByDesignation: 'Teacher',
        checkedByName: (settings.checkedBy || '').toUpperCase(),
        checkedByDesignation: settings.checkerDesignation || 'Learning Area Coordinator',
        approvedByName: (settings.principalName || '').toUpperCase(),
        approvedByDesignation: settings.principalDesignation || 'School Principal II',
        language: 'English',
    });
    const [dllContent, setDllContent] = useState<DllContent | null>(null);

    // LAS State
    const [lasForm, setLasForm] = useState({
        subject: 'Filipino',
        gradeLevel: '7',
        learningCompetency: '',
        lessonObjective: '',
        activityType: 'Guided Practice',
        language: 'Filipino',
    });
    const [lasContent, setLasContent] = useState<LearningActivitySheet | null>(null);

    // Exam State
    const [examObjectives, setExamObjectives] = useState<ExamObjective[]>([{ id: `obj-${Date.now()}`, text: '', days: '' }]);
    const [examSubject, setExamSubject] = useState('Science');
    const [examGradeLevel, setExamGradeLevel] = useState('10');
    const [examQuarter, setExamQuarter] = useState<string>('1');
    const [examContent, setExamContent] = useState<GeneratedExam | null>(null);


    // Persist form state to localStorage
    useEffect(() => {
        try {
            const savedState = localStorage.getItem('lessonPlannersState_v2');
            if (savedState) {
                const state = JSON.parse(savedState);
                if (state.dlpForm) setDlpForm(prev => ({...prev, ...state.dlpForm}));
                if (state.dllForm) setDllForm(prev => ({...prev, ...state.dllForm}));
                if (state.quizForm) setQuizForm(prev => ({...prev, ...state.quizForm}));
                if (state.lasForm) setLasForm(prev => ({...prev, ...state.lasForm}));
                if (state.activeTab) setActiveTab(state.activeTab);
                if (state.dlpContent) setDlpContent(state.dlpContent);
                if (state.dllContent) setDllContent(state.dllContent);
                if (state.quizContent) setQuizContent(state.quizContent);
                if (state.lasContent) setLasContent(state.lasContent);
                if (state.teacherPosition) setTeacherPosition(state.teacherPosition);
                if (state.dllFormat) setDllFormat(state.dllFormat);
                if (state.examObjectives) setExamObjectives(state.examObjectives);
                if (state.examSubject) setExamSubject(state.examSubject);
                if (state.examGradeLevel) setExamGradeLevel(state.examGradeLevel);
                if (state.examQuarter) setExamQuarter(state.examQuarter);
                if (state.examContent) setExamContent(state.examContent);
            }
        } catch (e) { console.error("Could not parse saved lesson planner state.", e); }
    }, []);

    useEffect(() => {
        const stateToSave = { dlpForm, dllForm, quizForm, lasForm, activeTab, dlpContent, dllContent, quizContent, lasContent, teacherPosition, dllFormat, examObjectives, examSubject, examGradeLevel, examQuarter, examContent };
        localStorage.setItem('lessonPlannersState_v2', JSON.stringify(stateToSave));
    }, [dlpForm, dllForm, quizForm, lasForm, activeTab, dlpContent, dllContent, quizContent, lasContent, teacherPosition, dllFormat, examObjectives, examSubject, examGradeLevel, examQuarter, examContent]);


    useEffect(() => {
        setDlpForm(prev => ({
            ...prev,
            teacher: settings.teacherName,
            schoolName: settings.schoolName,
            preparedByName: settings.teacherName.toUpperCase(),
            checkedByName: (settings.checkedBy || '').toUpperCase(),
            checkedByDesignation: settings.checkerDesignation || 'Learning Area Coordinator',
            approvedByName: (settings.principalName || '').toUpperCase(),
            approvedByDesignation: settings.principalDesignation || 'School Principal II',
        }));
        setDllForm(prev => ({
            ...prev,
            preparedByName: settings.teacherName.toUpperCase(),
            checkedByName: (settings.checkedBy || '').toUpperCase(),
            checkedByDesignation: settings.checkerDesignation || 'Learning Area Coordinator',
            approvedByName: (settings.principalName || '').toUpperCase(),
            approvedByDesignation: settings.principalDesignation || 'School Principal II',
        }));
    }, [settings]);

    const handleDlpFormChange = (e: React.ChangeEvent<HTMLInputElement | HTMLTextAreaElement | HTMLSelectElement>) => {
        const { id, value } = e.target;
        setDlpForm(prev => ({ ...prev, [id]: value }));
        if (id === 'teacher') {
            setDlpForm(prev => ({ ...prev, preparedByName: value.toUpperCase() }));
        }
    };
    
    const handleDllFormChange = (e: React.ChangeEvent<HTMLInputElement | HTMLTextAreaElement | HTMLSelectElement>) => {
        const { id, value } = e.target;
        setDllForm(prev => ({...prev, [id]: value}));
    };

    const handleLasFormChange = (e: React.ChangeEvent<HTMLInputElement | HTMLTextAreaElement | HTMLSelectElement>) => {
        const { id, value } = e.target;
        setLasForm(prev => ({ ...prev, [id]: value }));
    };

    const handleQuizFormChange = (e: React.ChangeEvent<HTMLInputElement | HTMLTextAreaElement | HTMLSelectElement>) => {
        const { id, value } = e.target;
        setQuizForm(prev => ({
            ...prev,
            [id]: id === 'numQuestions' ? Number(value) : value,
        }));
    };

    const handleQuizTypeChange = (quizType: QuizType) => {
        setQuizForm(prev => {
            const newQuizTypes = prev.quizTypes.includes(quizType)
                ? prev.quizTypes.filter(t => t !== quizType)
                : [...prev.quizTypes, quizType];
            return { ...prev, quizTypes: newQuizTypes };
        });
    };

    const handleActivityPointsChange = (index: number, value: string) => {
        setActivityPoints(prev => ({ ...prev, [index]: value }));
    };

    const handleGenerateRubric = async (activityIndex: number) => {
        if (!quizContent?.activities[activityIndex]) return;

        const points = parseInt(activityPoints[activityIndex] || '0', 10);
        if (isNaN(points) || points <= 0) {
            toast.error("Please enter a valid number of points.");
            return;
        }

        setGeneratingRubricIndex(activityIndex);
        const toastId = toast.loading(`Generating rubric for activity ${activityIndex + 1}...`);
        try {
            const activity = quizContent.activities[activityIndex];
            const newRubric = await generateRubricForActivity({
                activityName: activity.activityName,
                activityInstructions: activity.activityInstructions,
                totalPoints: points,
            });
            
            setQuizContent(prev => {
                if (!prev) return null;
                const newActivities = [...prev.activities];
                newActivities[activityIndex] = { ...newActivities[activityIndex], rubric: newRubric };
                return { ...prev, activities: newActivities };
            });

            toast.success("Rubric generated successfully!", { id: toastId });

        } catch(error) {
            let message = "An unknown error occurred.";
            if (error instanceof Error) message = error.message;
            toast.error(message, { id: toastId });
        } finally {
            setGeneratingRubricIndex(null);
        }
    };

    const handleAddExamObjective = () => {
        setExamObjectives(prev => [...prev, { id: `obj-${Date.now()}`, text: '', days: '' }]);
    };

    const handleRemoveExamObjective = (id: string) => {
        if (examObjectives.length > 1) {
            setExamObjectives(prev => prev.filter(obj => obj.id !== id));
        } else {
            toast.error("You must have at least one objective.");
        }
    };

    const handleExamObjectiveChange = (id: string, field: 'text' | 'days', value: string) => {
        setExamObjectives(prev => prev.map(obj => obj.id === id ? { ...obj, [field]: value } : obj));
    };

    const handleGenerateExam = async () => {
        const objectivesWithDays = examObjectives
            .map(obj => ({ text: obj.text.trim(), days: obj.days.trim() }))
            .filter(obj => obj.text && obj.days && !isNaN(parseInt(obj.days, 10)) && parseInt(obj.days, 10) > 0);
        
        if (objectivesWithDays.length === 0) {
            toast.error("Please provide at least one valid learning objective with the number of days taught.");
            return;
        }

        setIsLoading(true);
        setExamContent(null);
        const toastId = toast.loading('Generating 50-Item Examination...');

        try {
            const content = await generateExam({
                objectives: objectivesWithDays,
                subject: examSubject,
                gradeLevel: examGradeLevel,
                quarter: examQuarter,
            });
            setExamContent(content);
            toast.success('Examination generated successfully!', { id: toastId });
        } catch (error) {
            let message = "An unknown error occurred during exam generation.";
            if (error instanceof Error) message = error.message;
            toast.error(message, { id: toastId });
        } finally {
            setIsLoading(false);
        }
    };

    const generateDLP = async () => {
        const requiredFields: (keyof typeof dlpForm)[] = ['teacher', 'schoolName', 'subject', 'teachingDates', 'classSchedule', 'gradeLevel', 'learningCompetency', 'lessonObjective', 'previousLesson'];
        if (requiredFields.some(field => !dlpForm[field as keyof typeof dlpForm].trim())) {
            toast.error('Please fill in all required DLP fields.');
            return;
        }
        setIsLoading(true);
        setDlpContent(null);
        const toastId = toast.loading('Generating Daily Lesson Plan...', {
            style: { background: 'var(--info)', color: 'white' },
            iconTheme: { primary: 'white', secondary: 'var(--info)' },
        });
        try {
            const content = await generateDlpContent({
                gradeLevel: dlpForm.gradeLevel,
                learningCompetency: dlpForm.learningCompetency,
                lessonObjective: dlpForm.lessonObjective,
                previousLesson: dlpForm.previousLesson,
                selectedQuarter: dlpForm.quarterSelect,
                subject: dlpForm.subject,
                teacherPosition,
                language: dlpForm.language as 'English' | 'Filipino',
                dlpFormat: dlpForm.dlpFormat,
            });
            setDlpContent(content);
            toast.success('DLP generated successfully!', { id: toastId });
        } catch (error) {
            let message = "An unknown error occurred.";
            if (error instanceof Error) message = error.message;
            toast.error(message, { id: toastId });
        } finally {
            setIsLoading(false);
        }
    };
    
    const generateDLL = async () => {
        if (!dllForm.subject || !dllForm.gradeLevel) {
            toast.error('Please provide a Subject and Grade Level.');
            return;
        }
        setIsLoading(true);
        setDllContent(null);
        const toastId = toast.loading('Generating Weekly Plan...');
        try {
            const content = await generateDllContent({
                ...dllForm,
                language: dllForm.language as 'English' | 'Filipino',
                dllFormat: dllFormat,
            });
            setDllContent(content);
            toast.success('Weekly Plan generated successfully!', { id: toastId });
        } catch (error) {
            let message = "An unknown error occurred.";
            if (error instanceof Error) message = error.message;
            toast.error(message, { id: toastId });
        } finally {
            setIsLoading(false);
        }
    };

    const generateLAS = async () => {
        if (!lasForm.subject.trim() || !lasForm.learningCompetency.trim() || !lasForm.lessonObjective.trim()) {
            toast.error("Please fill in the Subject, Learning Competency, and Lesson Objective.");
            return;
        }
        setIsLoading(true);
        setLasContent(null);
        const toastId = toast.loading('Generating Learning Activity Sheet...');
        try {
            const content = await generateLearningActivitySheet({
                ...lasForm,
                language: lasForm.language as 'English' | 'Filipino',
            });
            setLasContent(content);
            toast.success('Learning Sheet generated successfully!', { id: toastId });
        } catch (error) {
            let message = "An unknown error occurred.";
            if (error instanceof Error) message = error.message;
            toast.error(message, { id: toastId });
        } finally {
            setIsLoading(false);
        }
    };

    const generateQuiz = async () => {
        if (!quizForm.quizTopic.trim() || quizForm.quizTypes.length === 0) {
            toast.error('Please provide a topic and select at least one quiz format.');
            return;
        }
        setIsLoading(true);
        setQuizContent(null);
        const toastId = toast.loading('Generating Quiz & Activities...', {
             style: { background: 'var(--info)', color: 'white' },
            iconTheme: { primary: 'white', secondary: 'var(--info)' },
        });
        try {
            const content = await generateQuizContent({
                topic: quizForm.quizTopic,
                numQuestions: quizForm.numQuestions,
                quizTypes: quizForm.quizTypes,
                subject: quizForm.subject,
                gradeLevel: quizForm.gradeLevel,
            });
            setQuizContent(content);
            toast.success('Quiz generated successfully!', { id: toastId });
        } catch (error) {
            let message = "An unknown error occurred.";
            if (error instanceof Error) message = error.message;
            toast.error(message, { id: toastId });
        } finally {
            setIsLoading(false);
        }
    };
    
    const handleDownloadDlpDocx = async () => {
        if (!dlpContent) {
            toast.error("No DLP content to download.");
            return;
        }
        setIsLoading(true);
        const toastId = toast.loading('Generating Word document...');
        try {
            await docxService.generateDlpDocx(dlpForm, dlpContent, "", settings);
            toast.success('DLP downloaded successfully!', { id: toastId });
        } catch (error) {
             let message = "An unknown error occurred.";
            if (error instanceof Error) message = error.message;
            toast.error(message, { id: toastId });
        } finally {
            setIsLoading(false);
        }
    };

    const handleDownloadDllDocx = async () => {
        if (!dllContent) {
            toast.error("No Weekly Plan content to download.");
            return;
        }
        setIsLoading(true);
        const toastId = toast.loading('Generating Word document...');
        try {
            const dllExportData = {
                ...dllForm, // Pass all form fields including signatories
                teacher: settings.teacherName,
                schoolName: settings.schoolName,
            };
            await docxService.generateDllDocx(dllExportData, dllContent, settings);
            toast.success('Weekly Plan downloaded successfully!', { id: toastId });
        } catch (error) {
            let message = "An unknown error occurred.";
            if (error instanceof Error) message = error.message;
            toast.error(message, { id: toastId });
        } finally {
            setIsLoading(false);
        }
    };

     const handleDownloadLasDocx = async () => {
        if (!lasContent) {
            toast.error("No Learning Sheet content to download.");
            return;
        }
        setIsLoading(true);
        const toastId = toast.loading('Generating Word document...');
        try {
            await docxService.generateLasDocx({
                schoolYear: settings.schoolYear,
                ...lasForm
            }, lasContent, settings);
            toast.success('Learning Sheet downloaded successfully!', { id: toastId });
        } catch (error) {
            let message = "An unknown error occurred.";
            if (error instanceof Error) message = error.message;
            toast.error(message, { id: toastId });
        } finally {
            setIsLoading(false);
        }
    };

    const handleDownloadQuizDocx = async () => {
        if (!quizContent) {
            toast.error("No quiz content to download.");
            return;
        }
        setIsLoading(true);
        const toastId = toast.loading('Generating Word document...');
        try {
            await docxService.generateQuizDocx(quizContent);
            toast.success('Quiz downloaded successfully!', { id: toastId });
        } catch (error) {
            let message = "An unknown error occurred.";
            if (error instanceof Error) message = error.message;
            toast.error(message, { id: toastId });
        } finally {
            setIsLoading(false);
        }
    };

    const handleDownloadExamDocx = async () => {
        if (!examContent) {
            toast.error("No exam content to download.");
            return;
        }
        setIsLoading(true);
        const toastId = toast.loading('Generating Examination Word document...');
        try {
            await docxService.generateExamDocx(examContent, settings);
            toast.success('Examination downloaded successfully!', { id: toastId });
        } catch (error) {
            let message = "An unknown error occurred during DOCX generation.";
            if (error instanceof Error) message = error.message;
            toast.error(message, { id: toastId });
        } finally {
            setIsLoading(false);
        }
    };

    const dlpOutputHtml = useMemo(() => {
        if (!dlpContent) return { mainContent: '', answerKeyHtml: '', reflectionTableHtml: ''};

        const isFilipino = dlpForm.language === 'Filipino';
        const t = {
            objectives: isFilipino ? 'I. LAYUNIN' : 'I. OBJECTIVES',
            contentStandard: isFilipino ? 'Pamantayang Pangnilalaman:' : 'Content Standard:',
            performanceStandard: isFilipino ? 'Pamantayan sa Pagganap:' : 'Performance Standard:',
            learningCompetency: isFilipino ? 'Kasanayan sa Pagkatuto:' : 'Learning Competency:',
            atTheEnd: isFilipino ? 'Sa pagtatapos ng aralin, ang mga mag-aaral ay inaasahang:' : 'At the end of the lesson, the learners should be able to:',
            content: isFilipino ? 'II. NILALAMAN' : 'II. CONTENT',
            topic: isFilipino ? 'Paksa:' : 'Topic:',
            resources: isFilipino ? 'III. KAGAMITANG PANTURO' : 'III. LEARNING RESOURCES',
            references: isFilipino ? 'Sanggunian:' : 'References:',
            materials: isFilipino ? 'Kagamitan:' : 'Materials:',
            procedure: isFilipino ? 'IV. PAMAMARAAN' : 'IV. PROCEDURE',
            remarks: isFilipino ? 'V. MGA TALA' : 'V. REMARKS',
            reflection: isFilipino ? 'VI. PAGNINILAY' : 'VI. REFLECTION',
        };

        const tableStyle = 'width: 100%; border-collapse: collapse; table-layout: fixed;';
        const cellStyle = 'padding: 8px; border: 1px solid var(--base-300); vertical-align: top; text-align: left;';
        const headerCellStyle = `${cellStyle} font-weight: bold; width: 25%;`;
        const contentCellStyle = `${cellStyle} width: 45%;`;
        const ppstCellStyle = `${cellStyle} width: 30%; font-style: italic; font-size: 0.9em; color: var(--primary);`;
        
        const scheduleHtml = dlpForm.classSchedule.split('\n').map(line => `<span>${line}</span>`).join('<br>');
        
        const mainContent = `
            <div class="font-serif text-sm">
                <table style="${tableStyle}">
                     <tr><td style="${cellStyle}; width: 15%; vertical-align: middle; text-align: center;" rowspan="5">${settings.schoolLogo ? `<img src="${settings.schoolLogo}" alt="logo" style="width: 60px; height: 60px; margin: auto;"/>` : ''}</td><td style="${cellStyle}; width: 55%;"><strong>${isFilipino ? 'Paaralan' : 'School'}:</strong> ${dlpForm.schoolName.toUpperCase()}</td><td style="${cellStyle}; width: 30%; text-align: center; vertical-align: middle;" rowspan="2"><strong>${isFilipino ? 'DETALYADONG BANGHAY-ARALIN SA' : 'DAILY LESSON PLAN IN'}<br/>${dlpForm.subject.toUpperCase()} ${dlpForm.gradeLevel}</strong></td></tr>
                    <tr><td style="${cellStyle}"><strong>${dlpForm.quarterSelect}</strong></td></tr>
                    <tr><td style="${cellStyle}"><strong>${isFilipino ? 'Guro' : 'Teacher'}:</strong> ${dlpForm.teacher}</td><td style="${cellStyle}" rowspan="3"><strong>${isFilipino ? 'ISKEDYUL NG KLASE' : 'CLASS SCHEDULE'}</strong><br/>${scheduleHtml}</td></tr>
                    <tr><td style="${cellStyle}"><strong>${isFilipino ? 'Asignatura' : 'Learning Area'}:</strong> ${dlpForm.subject.toUpperCase()}</td></tr>
                    <tr><td style="${cellStyle}"><strong>${isFilipino ? 'Petsa ng Pagtuturo' : 'Teaching Dates'}:</strong> ${dlpForm.teachingDates}</td></tr>
                </table>
                <h3 class="text-lg font-bold mt-4 mb-2 bg-base-300/30 p-1">${t.objectives}</h3>
                <table style="${tableStyle}">
                    <tr><td style="${headerCellStyle}">${t.contentStandard}</td><td style="${cellStyle}" colspan="2">${dlpContent.contentStandard}</td></tr>
                    <tr><td style="${headerCellStyle}">${t.performanceStandard}</td><td style="${cellStyle}" colspan="2">${dlpContent.performanceStandard}</td></tr>
                    <tr><td style="${headerCellStyle}">${t.learningCompetency}</td><td style="${cellStyle}" colspan="2">${dlpForm.learningCompetency}</td></tr>
                    <tr><td style="${cellStyle}" colspan="3">${t.atTheEnd}</td></tr>
                    <tr><td style="${cellStyle}" colspan="3"><ul class="list-disc ml-8"><li>${dlpForm.lessonObjective}</li></ul></td></tr>
                </table>
                <h3 class="text-lg font-bold mt-4 mb-2 bg-base-300/30 p-1">${t.content}</h3>
                <table style="${tableStyle}"><tr><td style="${headerCellStyle}">${t.topic}</td><td style="${cellStyle}" colspan="2">${dlpContent.topic}</td></tr></table>
                <h3 class="text-lg font-bold mt-4 mb-2 bg-base-300/30 p-1">${t.resources}</h3>
                <table style="${tableStyle}">
                    <tr><td style="${headerCellStyle}">${t.references}</td><td style="${cellStyle}" colspan="2">${dlpContent.learningReferences}</td></tr>
                    <tr><td style="${headerCellStyle}">${t.materials}</td><td style="${cellStyle}" colspan="2">${dlpContent.learningMaterials}</td></tr>
                </table>
                <h3 class="text-lg font-bold mt-4 mb-2 bg-base-300/30 p-1">${t.procedure}</h3>
                <table style="${tableStyle}">
                    <thead><tr><th style="${headerCellStyle}">${isFilipino ? 'Pamamaraan' : 'Procedure'}</th><th style="${contentCellStyle}">${isFilipino ? 'Gawain ng Guro/Mag-aaral' : 'Teacher/Student Activity'}</th><th style="${ppstCellStyle}">${isFilipino ? 'Mga Kaugnay na PPST Indicator' : 'Aligned PPST Indicators'}</th></tr></thead>
                    <tbody>
                        ${dlpContent.procedures.map(proc => `
                            <tr>
                                <td style="${headerCellStyle}">${proc.title}</td>
                                <td style="${contentCellStyle}">${proc.content.replace(/\n/g, '<br/>')}</td>
                                <td style="${ppstCellStyle}">${proc.ppst}</td>
                            </tr>
                        `).join('')}
                    </tbody>
                </table>
        `;
        const sectionsForReflection = (dlpForm.classSchedule || '').split('\n').map(line => {
            const parts = line.match(/([Gg]?\d+\s*-\s*[\w\s]+|[\w\s]+)/);
            return parts ? parts[0].trim().replace(/,/g, '') : line.trim();
        }).filter(Boolean);

        const reflectionTableHtml = `
            <h3 class="text-lg font-bold mt-4 mb-2 bg-base-300/30 p-1">${t.remarks}</h3>
            <div style="border: 1px solid var(--base-300); padding: 8px; min-height: 80px;">
                <p style="border-bottom: 1px solid var(--base-300); height: 24px;">${dlpContent.remarksContent || ''}</p>
            </div>
            <h3 class="text-lg font-bold mt-4 mb-2 bg-base-300/30 p-1">${t.reflection}</h3>
            <table style="${tableStyle.replace('table-layout: fixed;', '')}">
                <tbody>
                    <tr><td style="padding: 8px; border: 1px solid var(--base-300); vertical-align: top; text-align: left; font-weight: bold; width: 40%;">${isFilipino ? 'A. Bilang ng mag-aaral na nakakuha ng 80% sa pagtataya' : 'A. No. of learners who earned 80% in the evaluation'}</td><td style="padding: 8px; border: 1px solid var(--base-300); vertical-align: top; text-align: left; width: 60%;">${sectionsForReflection.length > 0 ? sectionsForReflection.map(sec => `<p>___ out of ___ learners earned 80% and above - ${sec}</p>`).join('') : `<p>___ out of ___ learners earned 80% and above</p>`}</td></tr>
                    <tr><td style="padding: 8px; border: 1px solid var(--base-300); vertical-align: top; text-align: left; font-weight: bold; width: 40%;">${isFilipino ? 'B. Bilang ng mag-aaral na nangangailangan ng remediation na nakakuha ng mababa sa 80%' : 'B. No. of learners who require additional activities for remediation who score below 80%'}</td><td style="padding: 8px; border: 1px solid var(--base-300); vertical-align: top; text-align: left; width: 60%;">${sectionsForReflection.length > 0 ? sectionsForReflection.map(sec => `<p>___ out of ___ learners require additional activities - ${sec}</p>`).join('') : `<p>___ out of ___ learners require additional activities</p>`}</td></tr>
                    <tr><td style="padding: 8px; border: 1px solid var(--base-300); vertical-align: top; text-align: left; font-weight: bold; width: 40%;">${isFilipino ? 'C. Nakatulong ba ang remedial? Bilang ng mag-aaral na nakaunawa sa aralin.' : 'C. Did the remedial lessons work? No. of learners who have caught up with the lessons.'}</td><td style="padding: 8px; border: 1px solid var(--base-300); vertical-align: top; text-align: left; width: 60%;"><p><span>☐</span> ${isFilipino ? 'Oo' : 'YES'} <span>☐</span> ${isFilipino ? 'Hindi' : 'NO'}</p><p><span>☐</span> ___ ${isFilipino ? 'na mag-aaral ang nakaunawa sa aralin' : 'learners caught up with the lesson'}</p></td></tr>
                    <tr><td style="padding: 8px; border: 1px solid var(--base-300); vertical-align: top; text-align: left; font-weight: bold; width: 40%;">${isFilipino ? 'D. Bilang ng mga mag-aaral na magpapatuloy sa remediation.' : 'D. No. of learners who continue to require remediation'}</td><td style="padding: 8px; border: 1px solid var(--base-300); vertical-align: top; text-align: left; width: 60%;"><p><span>☐</span> ___ ${isFilipino ? 'na mag-aaral ang magpapatuloy sa remediation' : 'learners continue to require remediation'}</p></td></tr>
                    <tr><td style="padding: 8px; border: 1px solid var(--base-300); vertical-align: top; text-align: left; font-weight: bold; width: 40%;">${isFilipino ? 'E. Alin sa mga istratehiyang pagtuturo nakatulong ng lubos? Paano ito nakatulong?' : 'E. Which of my teaching strategies work well? Why did this work?'}</td><td style="padding: 8px; border: 1px solid var(--base-300); vertical-align: top; text-align: left; width: 60%;"><p><span>☐</span> experiment</p><p><span>☐</span> collaborative learning</p><p><span>☐</span> differentiated instruction</p><p><span>☐</span> lecture</p><p><span>☐</span> think-pair-share</p><p><span>☐</span> role play</p><p><span>☐</span> discovery</p><p><span>☐</span> others</p></td></tr>
                    <tr><td style="padding: 8px; border: 1px solid var(--base-300); vertical-align: top; text-align: left; font-weight: bold; width: 40%;">${isFilipino ? 'F. Anong suliranin ang aking naranasan na solusyunan sa tulong ng aking punungguro at superbisor?' : 'F. What difficulties did I encounter which my principal or supervisor can help me solve?'}</td><td style="padding: 8px; border: 1px solid var(--base-300); vertical-align: top; text-align: left; width: 60%;"></td></tr>
                    <tr><td style="padding: 8px; border: 1px solid var(--base-300); vertical-align: top; text-align: left; font-weight: bold; width: 40%;">${isFilipino ? 'G. Anong kagamitang panturo ang aking nadibuho na nais kong ibahagi sa mga kapwa ko guro?' : 'G. What innovation or localized materials did I use/discover which I wish to share with other teachers?'}</td><td style="padding: 8px; border: 1px solid var(--base-300); vertical-align: top; text-align: left; width: 60%;"></td></tr>
                </tbody>
            </table>
        `;

        const answerKeyHtml = `
            <h3 class="text-lg font-bold mt-4 mb-2 bg-base-300/30 p-1">${isFilipino ? 'Susi sa Pagwawasto' : 'Answer Key'} (For Evaluating Learning)</h3>
            <ol class="list-decimal list-inside">
                ${dlpContent.evaluationQuestions.map(q => `<li>${q.answer}</li>`).join('')}
            </ol>
        `;

        return { mainContent, answerKeyHtml, reflectionTableHtml };
    }, [dlpContent, dlpForm, settings]);

    return (
        <div className="min-h-screen">
            <Header title="AI Lesson & Assessment Generators" />
            <div className="p-4 md:p-8">
                <div className="flex border-b border-base-300 mb-6 flex-wrap">
                    <TabButton label="Daily Lesson Plan (DLP)" icon={<ClipboardCheckIcon className="w-5 h-5"/>} isActive={activeTab === 'dlp'} onClick={() => setActiveTab('dlp')} />
                    <TabButton label="Daily Lesson Log (DLL)" icon={<ClipboardCheckIcon className="w-5 h-5"/>} isActive={activeTab === 'dll'} onClick={() => setActiveTab('dll')} />
                    <TabButton label="Learning Activity Sheet (LAS)" icon={<ClipboardCheckIcon className="w-5 h-5"/>} isActive={activeTab === 'las'} onClick={() => setActiveTab('las')} />
                    <TabButton label="Quiz Generator" icon={<ClipboardCheckIcon className="w-5 h-5"/>} isActive={activeTab === 'quiz'} onClick={() => setActiveTab('quiz')} />
                    <TabButton label="Exam Generator" icon={<ClipboardCheckIcon className="w-5 h-5"/>} isActive={activeTab === 'exam'} onClick={() => setActiveTab('exam')} />
                </div>

                <div className="grid grid-cols-1 lg:grid-cols-2 gap-8">
                    {/* Controls Column */}
                    <div className="bg-base-200 p-6 rounded-xl shadow-lg self-start">
                        {/* DLP FORM */}
                        {activeTab === 'dlp' && (
                            <form onSubmit={(e) => { e.preventDefault(); generateDLP(); }} className="space-y-4">
                                <h2 className="text-xl font-bold text-base-content mb-2">DLP Details</h2>
                                <div className="grid grid-cols-2 gap-4">
                                    <InputField id="subject" label="Subject" value={dlpForm.subject} onChange={handleDlpFormChange} required />
                                    <InputField id="gradeLevel" label="Grade Level" value={dlpForm.gradeLevel} onChange={handleDlpFormChange} required />
                                </div>
                                <TextAreaField id="learningCompetency" label="Learning Competency" value={dlpForm.learningCompetency} onChange={handleDlpFormChange} required placeholder="e.g., EN9G-IIa-19: Use adverbs in narration" />
                                <TextAreaField id="lessonObjective" label="Specific Lesson Objective" value={dlpForm.lessonObjective} onChange={handleDlpFormChange} required placeholder="e.g., Identify and use adverbs of manner in sentences." />
                                <InputField id="previousLesson" label="Previous Lesson" value={dlpForm.previousLesson} onChange={handleDlpFormChange} required placeholder="e.g., Types of Adjectives" />
                                 <div className="grid grid-cols-2 gap-4">
                                    <div><label htmlFor="language" className="block text-sm font-medium text-base-content mb-1">Language</label><select id="language" value={dlpForm.language} onChange={handleDlpFormChange} className="w-full bg-base-100 border border-base-300 rounded-md p-2 h-10"><option>English</option><option>Filipino</option></select></div>
                                    <div><label htmlFor="dlpFormat" className="block text-sm font-medium text-base-content mb-1">DLP Format</label><select id="dlpFormat" value={dlpForm.dlpFormat} onChange={handleDlpFormChange} className="w-full bg-base-100 border border-base-300 rounded-md p-2 h-10"><option>Standard DepEd</option><option>4As</option><option>5Es</option><option>Explicit Instruction</option></select></div>
                                </div>
                                <div className="grid grid-cols-1 md:grid-cols-2 gap-4 border-t border-base-300 pt-4">
                                     <InputField id="teacher" label="Teacher's Name" value={dlpForm.teacher} onChange={handleDlpFormChange} required />
                                     <InputField id="schoolName" label="School Name" value={dlpForm.schoolName} onChange={handleDlpFormChange} required />
                                     <InputField id="teachingDates" label="Teaching Dates & Time" value={dlpForm.teachingDates} onChange={handleDlpFormChange} placeholder="e.g., Oct. 2-6, 2023 (9:00 AM)" />
                                     <div><label htmlFor="quarterSelect" className="block text-sm font-medium text-base-content mb-1">Quarter</label><select id="quarterSelect" value={dlpForm.quarterSelect} onChange={handleDlpFormChange} className="w-full bg-base-100 border border-base-300 rounded-md p-2 h-10"><option>1ST QUARTER</option><option>2ND QUARTER</option><option>3RD QUARTER</option><option>4TH QUARTER</option></select></div>
                                     <TextAreaField id="classSchedule" label="Class Schedule" value={dlpForm.classSchedule} onChange={handleDlpFormChange} rows={2} placeholder="e.g., 9:00-10:00 AM - Grade 9 Ruby" />
                                     <div><label htmlFor="teacherPosition" className="block text-sm font-medium text-base-content mb-1">Teacher Position</label><select id="teacherPosition" value={teacherPosition} onChange={(e) => setTeacherPosition(e.target.value as any)} className="w-full bg-base-100 border border-base-300 rounded-md p-2 h-10"><option value="Beginning">Beginning (Teacher I-III)</option><option value="Proficient">Proficient (Master Teacher I-II)</option><option value="Highly Proficient">Highly Proficient (Master Teacher III-IV)</option><option value="Distinguished">Distinguished</option></select></div>
                                </div>
                                <button type="submit" disabled={isLoading} className="w-full flex items-center justify-center bg-primary hover:bg-primary-focus text-white font-bold py-3 px-4 rounded-lg text-lg mt-4 disabled:opacity-50">
                                    <SparklesIcon className={`w-6 h-6 mr-3 ${isLoading ? 'animate-spin' : ''}`} /> {isLoading ? 'Generating DLP...' : 'Generate Full DLP'}
                                </button>
                            </form>
                        )}

                        {/* DLL FORM */}
                        {activeTab === 'dll' && (
                             <form onSubmit={(e) => { e.preventDefault(); generateDLL(); }} className="space-y-4">
                                <h2 className="text-xl font-bold text-base-content mb-2">Weekly Plan Details</h2>
                                <div className="grid grid-cols-2 gap-4">
                                    <InputField id="subject" label="Subject" value={dllForm.subject} onChange={handleDllFormChange} required />
                                    <InputField id="gradeLevel" label="Grade Level" value={dllForm.gradeLevel} onChange={handleDllFormChange} required />
                                </div>
                                <InputField id="weeklyTopic" label="Weekly Topic" value={dllForm.weeklyTopic} onChange={handleDllFormChange} required placeholder="e.g., Linear Equations" />
                                <TextAreaField id="contentStandard" label="Content Standard" value={dllForm.contentStandard} onChange={handleDllFormChange} placeholder="(Optional) AI can generate this" />
                                <TextAreaField id="performanceStandard" label="Performance Standard" value={dllForm.performanceStandard} onChange={handleDllFormChange} placeholder="(Optional) AI can generate this" />
                                <div className="grid grid-cols-2 gap-4">
                                     <InputField id="quarter" label="Quarter" value={dllForm.quarter} onChange={handleDllFormChange} placeholder="e.g., 1" />
                                     <InputField id="teachingDates" label="Teaching Dates" value={dllForm.teachingDates} onChange={handleDllFormChange} placeholder="e.g., Sept 4-8, 2023" />
                                </div>
                                <div className="grid grid-cols-2 gap-4">
                                    <div><label htmlFor="dllLanguage" className="block text-sm font-medium text-base-content mb-1">Language</label><select id="language" value={dllForm.language} onChange={handleDllFormChange} className="w-full bg-base-100 border border-base-300 rounded-md p-2 h-10"><option>English</option><option>Filipino</option></select></div>
                                    <div><label htmlFor="dllFormat" className="block text-sm font-medium text-base-content mb-1">DLL Format</label><select id="dllFormat" value={dllFormat} onChange={(e) => setDllFormat(e.target.value)} className="w-full bg-base-100 border border-base-300 rounded-md p-2 h-10"><option>Standard</option><option>4As</option><option>5Es</option></select></div>
                                </div>
                                <button type="submit" disabled={isLoading} className="w-full flex items-center justify-center bg-primary hover:bg-primary-focus text-white font-bold py-3 px-4 rounded-lg text-lg mt-4 disabled:opacity-50">
                                    <SparklesIcon className={`w-6 h-6 mr-3 ${isLoading ? 'animate-spin' : ''}`} /> {isLoading ? 'Generating Plan...' : 'Generate Weekly Plan'}
                                </button>
                             </form>
                        )}

                        {/* QUIZ FORM */}
                        {activeTab === 'quiz' && (
                            <form onSubmit={(e) => { e.preventDefault(); generateQuiz(); }} className="space-y-4">
                                <h2 className="text-xl font-bold text-base-content mb-2">Quiz Details</h2>
                                <TextAreaField id="quizTopic" label="Quiz Topic / Coverage" value={quizForm.quizTopic} onChange={handleQuizFormChange} required placeholder="e.g., Photosynthesis and Cellular Respiration" />
                                <div className="grid grid-cols-2 gap-4">
                                    <InputField id="subject" label="Subject" value={quizForm.subject} onChange={handleQuizFormChange} required />
                                    <InputField id="gradeLevel" label="Grade Level" value={quizForm.gradeLevel} onChange={handleQuizFormChange} required />
                                </div>
                                <InputField id="numQuestions" label="Number of Questions (per type)" type="number" value={quizForm.numQuestions.toString()} onChange={handleQuizFormChange} required />
                                <div>
                                    <label className="block text-sm font-medium text-base-content mb-2">Quiz Types</label>
                                    <div className="flex flex-wrap gap-3">
                                        {quizTypesList.map(type => (
                                            <label key={type} className="flex items-center gap-2 cursor-pointer bg-base-100 p-2 rounded-md border border-base-300 hover:bg-base-300">
                                                <input type="checkbox" checked={quizForm.quizTypes.includes(type)} onChange={() => handleQuizTypeChange(type)} className="checkbox checkbox-primary checkbox-sm" />
                                                <span className="text-sm">{type}</span>
                                            </label>
                                        ))}
                                    </div>
                                </div>
                                <button type="submit" disabled={isLoading} className="w-full flex items-center justify-center bg-primary hover:bg-primary-focus text-white font-bold py-3 px-4 rounded-lg text-lg mt-4 disabled:opacity-50">
                                    <SparklesIcon className={`w-6 h-6 mr-3 ${isLoading ? 'animate-spin' : ''}`} /> {isLoading ? 'Generating Quiz...' : 'Generate Quiz'}
                                </button>
                            </form>
                        )}

                        {/* LAS FORM */}
                         {activeTab === 'las' && (
                            <form onSubmit={(e) => { e.preventDefault(); generateLAS(); }} className="space-y-4">
                                <h2 className="text-xl font-bold text-base-content mb-2">Activity Sheet Details</h2>
                                <div className="grid grid-cols-2 gap-4">
                                    <InputField id="subject" label="Subject" value={lasForm.subject} onChange={handleLasFormChange} required />
                                    <InputField id="gradeLevel" label="Grade Level" value={lasForm.gradeLevel} onChange={handleLasFormChange} required />
                                </div>
                                <TextAreaField id="learningCompetency" label="Learning Competency" value={lasForm.learningCompetency} onChange={handleLasFormChange} required placeholder="e.g., Nasusuri ang mga elementong linggwistiko" />
                                <TextAreaField id="lessonObjective" label="Lesson Objective" value={lasForm.lessonObjective} onChange={handleLasFormChange} required placeholder="e.g., Makasusulat ng maikling tula..." />
                                <div>
                                    <label htmlFor="activityType" className="block text-sm font-medium text-base-content mb-1">Type of Activity</label>
                                    <select id="activityType" value={lasForm.activityType} onChange={handleLasFormChange} className="w-full bg-base-100 border border-base-300 rounded-md p-2 h-10">
                                        {activityTypes.map(type => <option key={type} value={type}>{type}</option>)}
                                    </select>
                                </div>
                                 <div>
                                    <label htmlFor="lasLanguage" className="block text-sm font-medium text-base-content mb-1">Language</label>
                                    <select id="language" value={lasForm.language} onChange={handleLasFormChange} className="w-full bg-base-100 border border-base-300 rounded-md p-2 h-10">
                                        <option>English</option><option>Filipino</option>
                                    </select>
                                </div>
                                <button type="submit" disabled={isLoading} className="w-full flex items-center justify-center bg-primary hover:bg-primary-focus text-white font-bold py-3 px-4 rounded-lg text-lg mt-4 disabled:opacity-50">
                                    <SparklesIcon className={`w-6 h-6 mr-3 ${isLoading ? 'animate-spin' : ''}`} /> {isLoading ? 'Generating Sheet...' : 'Generate Activity Sheet'}
                                </button>
                            </form>
                         )}

                         {/* EXAM FORM */}
                         {activeTab === 'exam' && (
                             <div className="space-y-4">
                                 <h2 className="text-xl font-bold text-base-content mb-2">Periodical Exam Details</h2>
                                 <div className="grid grid-cols-3 gap-4">
                                     <div className="col-span-1"><InputField id="examSubject" label="Subject" value={examSubject} onChange={(e: any) => setExamSubject(e.target.value)} required /></div>
                                     <div className="col-span-1"><InputField id="examGradeLevel" label="Grade Level" value={examGradeLevel} onChange={(e: any) => setExamGradeLevel(e.target.value)} required /></div>
                                     <div className="col-span-1"><InputField id="examQuarter" label="Quarter" value={examQuarter} onChange={(e: any) => setExamQuarter(e.target.value)} required /></div>
                                 </div>
                                 
                                 <div>
                                     <label className="block text-sm font-medium text-base-content mb-2">Learning Objectives & Days Taught</label>
                                     <div className="space-y-2 max-h-64 overflow-y-auto p-1">
                                         {examObjectives.map((obj, index) => (
                                             <div key={obj.id} className="flex items-center gap-2">
                                                 <span className="text-sm font-bold w-6">{index + 1}.</span>
                                                 <input type="text" value={obj.text} onChange={e => handleExamObjectiveChange(obj.id, 'text', e.target.value)} placeholder="Learning Objective" className="flex-grow bg-base-100 border border-base-300 rounded-md p-2 h-10 text-sm" />
                                                 <input type="number" value={obj.days} onChange={e => handleExamObjectiveChange(obj.id, 'days', e.target.value)} placeholder="Days" className="w-20 bg-base-100 border border-base-300 rounded-md p-2 h-10 text-sm text-center" />
                                                 <button onClick={() => handleRemoveExamObjective(obj.id)} className="p-2 text-error hover:bg-base-300 rounded-md"><TrashIcon className="w-4 h-4" /></button>
                                             </div>
                                         ))}
                                     </div>
                                     <button onClick={handleAddExamObjective} className="mt-2 flex items-center text-primary hover:underline text-sm font-semibold"><PlusIcon className="w-4 h-4 mr-1"/> Add Objective</button>
                                 </div>

                                 <button onClick={handleGenerateExam} disabled={isLoading} className="w-full flex items-center justify-center bg-primary hover:bg-primary-focus text-white font-bold py-3 px-4 rounded-lg text-lg mt-4 disabled:opacity-50">
                                    <SparklesIcon className={`w-6 h-6 mr-3 ${isLoading ? 'animate-spin' : ''}`} /> {isLoading ? 'Generating Exam...' : 'Generate Examination'}
                                </button>
                             </div>
                         )}
                    </div>

                    {/* Preview Column */}
                    <div className="bg-base-200 p-6 rounded-xl shadow-lg">
                        <div className="flex justify-between items-center mb-4">
                            <h2 className="text-xl font-bold text-base-content">Preview</h2>
                            { (activeTab === 'dlp' && dlpContent) && <button onClick={handleDownloadDlpDocx} disabled={isLoading} className="flex items-center gap-2 bg-secondary hover:bg-secondary-focus text-white font-bold py-2 px-3 rounded-lg text-sm"><DownloadIcon className="w-4 h-4"/> Download Word File</button> }
                            { (activeTab === 'dll' && dllContent) && <button onClick={handleDownloadDllDocx} disabled={isLoading} className="flex items-center gap-2 bg-secondary hover:bg-secondary-focus text-white font-bold py-2 px-3 rounded-lg text-sm"><DownloadIcon className="w-4 h-4"/> Download Word File</button> }
                            { (activeTab === 'las' && lasContent) && <button onClick={handleDownloadLasDocx} disabled={isLoading} className="flex items-center gap-2 bg-secondary hover:bg-secondary-focus text-white font-bold py-2 px-3 rounded-lg text-sm"><DownloadIcon className="w-4 h-4"/> Download Word File</button> }
                            { (activeTab === 'quiz' && quizContent) && <button onClick={handleDownloadQuizDocx} disabled={isLoading} className="flex items-center gap-2 bg-secondary hover:bg-secondary-focus text-white font-bold py-2 px-3 rounded-lg text-sm"><DownloadIcon className="w-4 h-4"/> Download Word File</button> }
                            { (activeTab === 'exam' && examContent) && <button onClick={handleDownloadExamDocx} disabled={isLoading} className="flex items-center gap-2 bg-secondary hover:bg-secondary-focus text-white font-bold py-2 px-3 rounded-lg text-sm"><DownloadIcon className="w-4 h-4"/> Download Word File</button> }
                        </div>
                        <div className="bg-base-100 p-4 rounded-md min-h-[50vh] max-h-[80vh] overflow-y-auto prose prose-sm max-w-none prose-headings:text-primary prose-strong:text-base-content">
                            {isLoading && (
                                <div className="flex flex-col items-center justify-center h-full text-center">
                                    <SparklesIcon className="w-16 h-16 text-primary animate-pulse" />
                                    <p className="mt-4 font-semibold text-base-content">Generating content, please wait...</p>
                                    <p className="text-sm text-base-content/70">This may take up to a minute for complex documents.</p>
                                </div>
                            )}
                            {!isLoading && activeTab === 'dlp' && dlpContent && (
                                <div dangerouslySetInnerHTML={{ __html: dlpOutputHtml.mainContent + dlpOutputHtml.answerKeyHtml + dlpOutputHtml.reflectionTableHtml }} />
                            )}
                            {!isLoading && activeTab === 'dll' && dllContent && (
                                <div className="font-serif text-sm">
                                    <h3 className="text-center font-bold text-lg mb-4">DAILY LESSON LOG</h3>
                                    <div className="grid grid-cols-2 gap-x-4 mb-4 text-xs border-b border-base-300 pb-4">
                                        <div><strong>Teacher:</strong> {settings.teacherName}</div>
                                        <div><strong>Grade Level:</strong> {dllForm.gradeLevel}</div>
                                        <div><strong>Subject:</strong> {dllForm.subject}</div>
                                        <div><strong>Quarter:</strong> {dllForm.quarter}</div>
                                        <div className="col-span-2"><strong>Teaching Dates:</strong> {dllForm.teachingDates}</div>
                                    </div>
                                    <div className="space-y-4">
                                        <div>
                                            <h4 className="font-bold bg-base-300/30 p-1">I. OBJECTIVES</h4>
                                            <p><strong>A. Content Standard:</strong> {dllContent.contentStandard}</p>
                                            <p><strong>B. Performance Standard:</strong> {dllContent.performanceStandard}</p>
                                        </div>
                                        <div>
                                            <h4 className="font-bold bg-base-300/30 p-1">IV. PROCEDURES</h4>
                                            <table className="w-full border-collapse border border-base-300 mt-2 text-xs">
                                                <thead>
                                                    <tr className="bg-base-200"><th className="border p-1">Procedure</th><th className="border p-1">Monday</th><th className="border p-1">Tuesday</th><th className="border p-1">Wednesday</th><th className="border p-1">Thursday</th><th className="border p-1">Friday</th></tr>
                                                </thead>
                                                <tbody>
                                                    {dllContent.procedures.map((proc, i) => (
                                                        <tr key={i}>
                                                            <td className="border p-1 font-semibold">{proc.procedure}</td>
                                                            <td className="border p-1">{proc.monday}</td>
                                                            <td className="border p-1">{proc.tuesday}</td>
                                                            <td className="border p-1">{proc.wednesday}</td>
                                                            <td className="border p-1">{proc.thursday}</td>
                                                            <td className="border p-1">{proc.friday}</td>
                                                        </tr>
                                                    ))}
                                                </tbody>
                                            </table>
                                        </div>
                                    </div>
                                </div>
                            )}
                            {!isLoading && activeTab === 'las' && lasContent && (
                                <div className="font-serif text-sm space-y-6">
                                    {lasContent.days.map((day, i) => (
                                        <div key={i} className="border-b-4 border-base-300 pb-6">
                                            <h3 className="text-xl font-bold text-center mb-1">{day.dayTitle}</h3>
                                            <div className="text-center text-xs mb-4 border-b border-black pb-2">
                                                <p className="font-bold">DepED | Dynamic Learning Program | BAGONG PILIPINAS | LEARNING ACTIVITY SHEET</p>
                                            </div>
                                            <div className="grid grid-cols-2 gap-x-4 text-xs mb-4">
                                                <p><strong>Subject:</strong> {lasForm.subject}</p>
                                                <p><strong>Grade & Section:</strong> _________</p>
                                                <p><strong>Activity Title:</strong> {day.activityTitle}</p>
                                                <p><strong>Learning Target:</strong> {day.learningTarget}</p>
                                            </div>
                                            
                                            {day.conceptNotes.map((note, idx) => (
                                                <div key={idx} className="mb-4">
                                                    <h4 className="font-bold uppercase mb-1">{note.title}</h4>
                                                    <div className="text-justify whitespace-pre-wrap">{note.content}</div>
                                                </div>
                                            ))}

                                            {day.activities.map((act, idx) => (
                                                <div key={idx} className="mb-4">
                                                    <h4 className="font-bold uppercase mb-1">{act.title}</h4>
                                                    <div className="whitespace-pre-wrap mb-2">{act.instructions}</div>
                                                    {act.questions && (
                                                        <ol className="list-decimal list-inside ml-2">
                                                            {act.questions.map((q, qId) => <li key={qId}>{q.questionText}</li>)}
                                                        </ol>
                                                    )}
                                                </div>
                                            ))}
                                            
                                            <div className="mt-4">
                                                <h4 className="font-bold uppercase">REFLECTION</h4>
                                                <p>{day.reflection}</p>
                                                <div className="border-b border-black mt-6"></div>
                                            </div>
                                        </div>
                                    ))}
                                </div>
                            )}
                            {!isLoading && activeTab === 'quiz' && quizContent && (
                                <div className="font-serif text-sm">
                                    <h3 className="text-center font-bold text-lg">{quizContent.quizTitle}</h3>
                                    {Object.entries(quizContent.questionsByType).map(([type, sec]) => {
                                        const section = sec as GeneratedQuizSection;
                                        return (
                                        <div key={type} className="mt-4">
                                            <h4 className="font-bold uppercase">{type}</h4>
                                            <p className="italic text-xs mb-2">{section.instructions}</p>
                                            <ol className="list-decimal list-inside">
                                                {section.questions.map((q, i) => (
                                                    <li key={i} className="mb-1">
                                                        {q.questionText}
                                                        {q.options && (
                                                            <ul className="list-[upper-alpha] list-inside ml-4 grid grid-cols-2">
                                                                {q.options.map((opt, oi) => <li key={oi}>{opt}</li>)}
                                                            </ul>
                                                        )}
                                                    </li>
                                                ))}
                                            </ol>
                                        </div>
                                    )})}
                                    {quizContent.activities && quizContent.activities.length > 0 && (
                                        <div className="mt-6 border-t pt-4">
                                            <h4 className="font-bold uppercase mb-2">Performance Tasks / Activities</h4>
                                            {quizContent.activities.map((act, i) => (
                                                <div key={i} className="mb-4 bg-base-200/50 p-3 rounded-md">
                                                    <p className="font-bold">{act.activityName}</p>
                                                    <p className="text-xs italic mb-2">{act.activityInstructions}</p>
                                                    
                                                    {act.rubric ? (
                                                        <div className="mt-2">
                                                            <p className="font-bold text-xs">Rubric:</p>
                                                            <table className="w-full text-xs border-collapse border border-base-300 mt-1">
                                                                <thead><tr className="bg-base-300"><th className="border p-1 text-left">Criteria</th><th className="border p-1 w-16">Points</th></tr></thead>
                                                                <tbody>
                                                                    {act.rubric.map((r, ri) => <tr key={ri}><td className="border p-1">{r.criteria}</td><td className="border p-1 text-center">{r.points}</td></tr>)}
                                                                </tbody>
                                                            </table>
                                                        </div>
                                                    ) : (
                                                        <div className="flex items-center gap-2 mt-2">
                                                            <input type="number" placeholder="Total Points" value={activityPoints[i] || ''} onChange={e => handleActivityPointsChange(i, e.target.value)} className="w-24 p-1 text-xs border rounded bg-base-100" />
                                                            <button onClick={() => handleGenerateRubric(i)} disabled={generatingRubricIndex === i} className="text-xs bg-primary text-white px-2 py-1 rounded flex items-center gap-1">
                                                                {generatingRubricIndex === i ? <RefreshCwIcon className="w-3 h-3 animate-spin"/> : <SparklesIcon className="w-3 h-3"/>} Generate Rubric
                                                            </button>
                                                        </div>
                                                    )}
                                                </div>
                                            ))}
                                        </div>
                                    )}
                                    <div className="mt-8 border-t-2 border-dashed pt-4">
                                        <h4 className="font-bold text-center mb-2">Answer Key</h4>
                                        <div className="grid grid-cols-2 gap-4 text-xs">
                                            {Object.entries(quizContent.questionsByType).map(([type, sec]) => {
                                                const section = sec as GeneratedQuizSection;
                                                return (
                                                <div key={type}>
                                                    <p className="font-bold underline">{type}</p>
                                                    <ol className="list-decimal list-inside">
                                                        {section.questions.map((q, i) => <li key={i}>{q.correctAnswer}</li>)}
                                                    </ol>
                                                </div>
                                            )})}
                                        </div>
                                    </div>
                                </div>
                            )}
                             {!isLoading && activeTab === 'exam' && examContent && (
                                <div className="font-serif text-sm">
                                    <h3 className="text-center font-bold text-lg uppercase">{examContent.title}</h3>
                                    <div className="mt-4 mb-6">
                                        <h4 className="font-bold text-center mb-2">TABLE OF SPECIFICATIONS</h4>
                                        <table className="w-full border-collapse border border-black text-[10px] text-center">
                                            <thead>
                                                <tr className="bg-gray-100">
                                                    <th className="border border-black p-1 w-1/3">Learning Competencies</th>
                                                    <th className="border border-black p-1">Days Taught</th>
                                                    <th className="border border-black p-1">%</th>
                                                    <th className="border border-black p-1">No. of Items</th>
                                                    <th className="border border-black p-1">Item Placement</th>
                                                    <th className="border border-black p-1">R</th>
                                                    <th className="border border-black p-1">U</th>
                                                    <th className="border border-black p-1">Ap</th>
                                                    <th className="border border-black p-1">An</th>
                                                    <th className="border border-black p-1">E</th>
                                                    <th className="border border-black p-1">C</th>
                                                </tr>
                                            </thead>
                                            <tbody>
                                                {examContent.tableOfSpecifications.map((row, i) => (
                                                    <tr key={i}>
                                                        <td className="border border-black p-1 text-left">{row.objective}</td>
                                                        <td className="border border-black p-1">{row.daysTaught}</td>
                                                        <td className="border border-black p-1">{row.percentage}</td>
                                                        <td className="border border-black p-1">{row.numItems}</td>
                                                        <td className="border border-black p-1">{row.itemPlacement}</td>
                                                        <td className="border border-black p-1">{row.remembering}</td>
                                                        <td className="border border-black p-1">{row.understanding}</td>
                                                        <td className="border border-black p-1">{row.applying}</td>
                                                        <td className="border border-black p-1">{row.analyzing}</td>
                                                        <td className="border border-black p-1">{row.evaluating}</td>
                                                        <td className="border border-black p-1">{row.creating}</td>
                                                    </tr>
                                                ))}
                                                <tr className="font-bold bg-gray-100">
                                                    <td className="border border-black p-1 text-right">TOTAL</td>
                                                    <td className="border border-black p-1">{examContent.tableOfSpecifications.reduce((sum, row) => sum + row.daysTaught, 0)}</td>
                                                    <td className="border border-black p-1">100%</td>
                                                    <td className="border border-black p-1">50</td>
                                                    <td className="border border-black p-1" colSpan={7}></td>
                                                </tr>
                                            </tbody>
                                        </table>
                                    </div>

                                    <div className="mt-6">
                                        <h4 className="font-bold mb-2">TEST QUESTIONS</h4>
                                        <ol className="list-decimal list-inside space-y-2">
                                            {examContent.questions.map((q, i) => (
                                                <li key={i} className="break-inside-avoid">
                                                    <span className="font-semibold">{q.questionText}</span>
                                                    <div className="ml-4 grid grid-cols-2 gap-1 mt-1 text-xs">
                                                        {q.options?.map((opt, oi) => (
                                                            <div key={oi}><span className="font-bold">{String.fromCharCode(65+oi)}.</span> {opt}</div>
                                                        ))}
                                                    </div>
                                                </li>
                                            ))}
                                        </ol>
                                    </div>
                                    
                                     <div className="mt-8 border-t-2 border-dashed pt-4 break-inside-avoid">
                                        <h4 className="font-bold text-center mb-2">Answer Key</h4>
                                        <div className="grid grid-cols-5 gap-2 text-xs">
                                            {examContent.questions.map((q, i) => (
                                                 <div key={i}><strong>{i+1}.</strong> {q.correctAnswer}</div>
                                            ))}
                                        </div>
                                    </div>
                                </div>
                             )}

                             {!isLoading && !dlpContent && !dllContent && !lasContent && !quizContent && !examContent && (
                                <div className="flex flex-col items-center justify-center h-full text-center">
                                    <ClipboardCheckIcon className="w-24 h-24 text-base-300" />
                                    <p className="mt-4 text-base-content/70">Fill in the details on the left and click "Generate" to see your document here.</p>
                                </div>
                             )}
                        </div>
                    </div>
                </div>
            </div>
        </div>
    );
};

export default LessonPlanners;