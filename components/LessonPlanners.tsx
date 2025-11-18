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
const jhsGradeLevels = ['7', '8', '9', '10'];

const subjectAreas = {
  "Elementary (K-6)": ["Kindergarten (Domains)", "Mother Tongue", "Filipino", "English", "Mathematics", "Science", "Araling Panlipunan (AP)", "Edukasyon sa Pagpapakatao (EsP)", "Music", "Arts", "Physical Education (PE)", "Health", "Edukasyong Pantahanan at Pangkabuhayan (EPP)", "Technology and Livelihood Education (TLE)"],
  "Junior High School (Grades 7-10)": ["Filipino", "English", "Mathematics", "Science", "Araling Panlipunan (AP)", "Edukasyon sa Pagpapakatao (EsP)", "Technology and Livelihood Education (TLE)", "Music", "Arts", "Physical Education (PE)", "Health"],
  "Senior High School - Core (Grades 11-12)": ["21st Century Literature from the Philippines and the World", "Contemporary Philippine Arts from the Regions", "Earth and Life Science", "General Mathematics", "Introduction to the Philosophy of the Human Person", "Komunikasyon at Pananaliksik sa Wika at Kulturang Pilipino", "Media and Information Literacy", "Oral Communication in Context", "Pagbasa at Pagsusuri ng Iba't Ibang Teksto Tungo sa Pananaliksik", "Personal Development", "Physical Education and Health", "Physical Science", "Reading and Writing Skills", "Statistics and Probability", "Understanding Culture, Society and Politics"],
  "Senior High School - Applied (Grades 11-12)": ["Empowerment Technologies", "English for Academic and Professional Purposes", "Entrepreneurship", "Filipino sa Piling Larang", "Practical Research 1", "Practical Research 2"]
};

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
    // FIX: Add state for exam quarter to resolve missing property error.
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
                // Exam state
                if (state.examObjectives) setExamObjectives(state.examObjectives);
                if (state.examSubject) setExamSubject(state.examSubject);
                if (state.examGradeLevel) setExamGradeLevel(state.examGradeLevel);
                // FIX: Load saved exam quarter from local storage.
                if (state.examQuarter) setExamQuarter(state.examQuarter);
                if (state.examContent) setExamContent(state.examContent);
            }
        } catch (e) { console.error("Could not parse saved lesson planner state.", e); }
    }, []);

    useEffect(() => {
        // FIX: Include examQuarter in the state saved to local storage.
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
                // FIX: Pass the examQuarter state to the generateExam function.
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

    // This is the reconstructed UI.
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
                                <button type="submit" disabled={isLoading} className="w-full flex items-center justify-center bg-primary hover:bg-primary-focus text-white font-bold py-3 px-4 rounded-lg text-lg mt-4 disabled:opacity-50">
                                    <SparklesIcon className={`w-6 h-6 mr-3 ${isLoading ? 'animate-spin' : ''}`} /> {isLoading ? 'Generating DLP...' : 'Generate Full DLP'}
                                </button>
                            </form>
                        )}
                        {/* Other forms would go here */}
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
                            {/* Other previews would go here */}
                             {!isLoading && !dlpContent && (
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
