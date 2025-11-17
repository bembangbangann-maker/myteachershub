import React, { useState, useEffect, useMemo, useCallback } from 'react';
import { toast } from 'react-hot-toast';
import { useAppContext } from '../contexts/AppContext';
import { generateDlpContent, generateQuizContent, generateRubricForActivity, generateDllContent, generateLearningActivitySheet, generateExam } from '../services/geminiService';
import { DlpContent, GeneratedQuiz, QuizType, DlpRubricItem, GeneratedQuizSection, DllContent, DlpProcedure, LearningActivitySheet, SchoolSettings, ExamObjective, GeneratedExam } from '../types';
import Header from './Header';
import { SparklesIcon, DownloadIcon, ClipboardCheckIcon, PlusIcon, TrashIcon } from './icons';
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
                    <tr><td style="padding: 8px; border: 1px solid var(--base-300); vertical-align: top; text-align: left; font-weight: bold; width: 40%;">${isFilipino ? 'E. Alin sa mga istratehiyang pagtuturo nakatulong ng lubos? Paano ito nakatulong?' : 'E. Which of my teaching strategies work well? Why did this work?'}</td><td style="padding: 8px; border: 1px solid var(--base-300); vertical-align: top; text-align: left; width: 60%;"><p><span>☐</span> experiment</p><p><span>☐</span> collaborative learning</p><p><span>☐</span> differentiated instruction</p><p><span>☐</span> lecture</p><p><span>☐</span> think-pair-share</p><p><span>☐</span> role play</p><p><span>☐</span> discovery</p><p><span>☐</span> board work</p><p>${isFilipino ? 'Bakit' : 'Why'}? ____________________</p></td></tr>
                    <tr><td style="padding: 8px; border: 1px solid var(--base-300); vertical-align: top; text-align: left; font-weight: bold; width: 40%;">${isFilipino ? 'F. Anong suliranin ang aking naranasan na solusyunan sa tulong ang aking punungguro at superbisor?' : 'F. What difficulties did I encounter which my principal or supervisor can help me solve?'}</td><td style="padding: 8px; border: 1px solid var(--base-300); vertical-align: top; text-align: left; width: 60%;"><p><span>☐</span> bullying among students</p><p><span>☐</span> student's behavior/attitude</p><p><span>☐</span> unavailable technology/equipment (AVR/LCD)</p><p><span>☐</span> internet lab</p><p>${isFilipino ? 'Bakit' : 'Why'}? ____________________</p></td></tr>
                    <tr><td style="padding: 8px; border: 1px solid var(--base-300); vertical-align: top; text-align: left; font-weight: bold; width: 40%;">${isFilipino ? 'G. Anong kagamitang panturo ang aking nadibuho na nais kong ibahagi sa mga kapwa ko guro.' : 'G. What innovation or localized materials did I use / discover which I wish to share with other teachers.'}</td><td style="padding: 8px; border: 1px solid var(--base-300); vertical-align: top; text-align: left; width: 60%;"><p><span>☐</span> localized videos</p><p><span>☐</span> colorful worksheets</p><p><span>☐</span> local jingle composition</p><p>${isFilipino ? 'Bakit' : 'Why'}? ____________________</p></td></tr>
                </tbody>
            </table>
        `;
        const answerKeyHtml = `
            <div class="page-break" style="page-break-before: always;"></div>
            <h3 class="text-lg font-bold mt-4 mb-2">${isFilipino ? 'Susi sa Pagwawasto' : 'Answer Key (For Evaluating Learning)'}</h3>
            <ol class="list-decimal ml-6">
                ${(dlpContent.evaluationQuestions || []).map(q => `<li>${q.answer}</li>`).join('')}
            </ol>
        `;

        return { mainContent: mainContent + '</div>', answerKeyHtml, reflectionTableHtml };

    }, [dlpContent, dlpForm, settings.schoolLogo]);
    
    const { dllHeaders, dllDays } = useMemo(() => {
        const isFilipino = dllForm.language === 'Filipino';
        return {
            dllHeaders: {
                section: isFilipino ? 'Seksyon' : 'Section',
                contentStandard: isFilipino ? 'A. Pamantayang Pangnilalaman' : 'A. Content Standard',
                performanceStandard: isFilipino ? 'B. Pamantayan sa Pagganap' : 'B. Performance Standard',
                learningCompetencies: isFilipino ? 'C. Mga Kasanayan sa Pagkatuto' : 'C. Learning Competencies',
                content: isFilipino ? 'II. NILALAMAN' : 'II. CONTENT',
                resources: isFilipino ? 'III. KAGAMITANG PANTURO' : 'III. LEARNING RESOURCES',
                procedures: isFilipino ? 'IV. PAMAMARAAN' : 'IV. PROCEDURES',
                remarks: isFilipino ? 'V. MGA TALA' : 'V. REMARKS',
                reflection: isFilipino ? 'VI. PAGNINILAY' : 'VI. REFLECTION',
            },
            dllDays: isFilipino ? ['Lunes', 'Martes', 'Miyerkules', 'Huwebes', 'Biyernes'] : ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday'],
        };
    }, [dllForm.language]);


    return (
        <div className="min-h-screen">
            <Header title="Lesson Planners" />
            <div className="p-4 md:p-8">
                 <div className="flex border-b border-base-300 mb-6">
                    <TabButton label="DLP Generator" icon={<SparklesIcon className="w-4 h-4" />} isActive={activeTab === 'dlp'} onClick={() => setActiveTab('dlp')} />
                    <TabButton label="Weekly Plan Generator" icon={<SparklesIcon className="w-4 h-4" />} isActive={activeTab === 'dll'} onClick={() => setActiveTab('dll')} />
                    <TabButton label="Quiz Generator" icon={<SparklesIcon className="w-4 h-4" />} isActive={activeTab === 'quiz'} onClick={() => setActiveTab('quiz')} />
                    <TabButton label="Learning Sheets" icon={<ClipboardCheckIcon className="w-4 h-4" />} isActive={activeTab === 'las'} onClick={() => setActiveTab('las')} />
                    <TabButton label="Exam Generator" icon={<ClipboardCheckIcon className="w-4 h-4" />} isActive={activeTab === 'exam'} onClick={() => setActiveTab('exam')} />
                </div>
                <div className="grid grid-cols-1 lg:grid-cols-3 gap-8">
                    {/* Left Column: Forms */}
                    <div className="lg:col-span-1 bg-base-200 p-6 rounded-xl shadow-lg self-start">
                        {activeTab === 'dlp' ? (
                             <DlpFormUI dlpForm={dlpForm} handleDlpFormChange={handleDlpFormChange} teacherPosition={teacherPosition} setTeacherPosition={setTeacherPosition} generateDLP={generateDLP} isLoading={isLoading} />
                        ) : activeTab === 'dll' ? (
                            <DllFormUI dllForm={dllForm} handleDllFormChange={handleDllFormChange} dllFormat={dllFormat} setDllFormat={setDllFormat} generateDLL={generateDLL} isLoading={isLoading} />
                        ) : activeTab === 'las' ? (
                            <LasFormUI lasForm={lasForm} handleLasFormChange={handleLasFormChange} generateLAS={generateLAS} isLoading={isLoading} />
                        ) : activeTab === 'exam' ? (
                            <ExamGeneratorForm
                                examObjectives={examObjectives}
                                onAddObjective={handleAddExamObjective}
                                onRemoveObjective={handleRemoveExamObjective}
                                onObjectiveChange={handleExamObjectiveChange}
                                subject={examSubject}
                                setSubject={setExamSubject}
                                gradeLevel={examGradeLevel}
                                setGradeLevel={setExamGradeLevel}
                                quarter={examQuarter}
                                setQuarter={setExamQuarter}
                                onGenerate={handleGenerateExam}
                                isLoading={isLoading}
                            />
                        ) : (
                            <QuizFormUI quizForm={quizForm} handleQuizFormChange={handleQuizFormChange} handleQuizTypeChange={handleQuizTypeChange} generateQuiz={generateQuiz} isLoading={isLoading} />
                        )}
                    </div>
                    {/* Right Column: Previews */}
                     <div className="lg:col-span-2 bg-base-200 rounded-xl shadow-lg flex flex-col">
                        {isLoading && (<div className="flex-grow flex items-center justify-center text-center p-16"><div className="flex flex-col items-center"><SparklesIcon className="w-16 h-16 mx-auto text-primary animate-pulse mb-4" /><h3 className="text-2xl font-bold">AI is Generating...</h3><p className="mt-2">This may take a moment.</p></div></div>)}
                        {!isLoading && activeTab === 'dlp' && !dlpContent && (<div className="flex-grow flex items-center justify-center text-center p-16"><div><h3 className="text-2xl font-bold">DLP Preview</h3><p className="mt-2">Your generated Daily Lesson Plan will appear here.</p></div></div>)}
                        {!isLoading && activeTab === 'dll' && !dllContent && (<div className="flex-grow flex items-center justify-center text-center p-16"><div><h3 className="text-2xl font-bold">Weekly Plan Preview</h3><p className="mt-2">Your generated Weekly Plan will appear here.</p></div></div>)}
                        {!isLoading && activeTab === 'quiz' && !quizContent && (<div className="flex-grow flex items-center justify-center text-center p-16"><div><h3 className="text-2xl font-bold">Quiz Preview</h3><p className="mt-2">Your generated quiz will appear here.</p></div></div>)}
                        {!isLoading && activeTab === 'las' && !lasContent && (<div className="flex-grow flex items-center justify-center text-center p-16"><div><h3 className="text-2xl font-bold">Learning Sheet Preview</h3><p className="mt-2">Your generated activity sheet will appear here.</p></div></div>)}
                        {!isLoading && activeTab === 'exam' && !examContent && (<div className="flex-grow flex items-center justify-center text-center p-16"><div><h3 className="text-2xl font-bold">Exam Preview</h3><p className="mt-2">Your generated 50-item examination will appear here.</p></div></div>)}
                        
                        {!isLoading && dlpContent && activeTab === 'dlp' && (
                            <DlpPreview dlpOutputHtml={dlpOutputHtml} handleDownloadDlpDocx={handleDownloadDlpDocx} isLoading={isLoading} />
                        )}

                        {!isLoading && dllContent && activeTab === 'dll' && (
                             <DllPreview dllContent={dllContent} dllHeaders={dllHeaders} dllDays={dllDays} handleDownloadDllDocx={handleDownloadDllDocx} isLoading={isLoading} />
                        )}
                        
                        {!isLoading && lasContent && activeTab === 'las' && (
                           <LasPreview lasContent={lasContent} settings={settings} lasForm={lasForm} onDownload={handleDownloadLasDocx} />
                        )}

                        {!isLoading && quizContent && activeTab === 'quiz' && (
                           <QuizPreview quizContent={quizContent} activityPoints={activityPoints} handleActivityPointsChange={handleActivityPointsChange} handleGenerateRubric={handleGenerateRubric} generatingRubricIndex={generatingRubricIndex} handleDownloadQuizDocx={handleDownloadQuizDocx} isLoading={isLoading} />
                        )}

                        {!isLoading && examContent && activeTab === 'exam' && (
                            <ExamPreview examContent={examContent} onDownload={handleDownloadExamDocx} isLoading={isLoading} />
                        )}
                    </div>
                </div>
            </div>
        </div>
    );
};

// --- START SUB-COMPONENTS ---
// To keep the main component cleaner, UI sections are broken down.

const DlpFormUI = ({ dlpForm, handleDlpFormChange, teacherPosition, setTeacherPosition, generateDLP, isLoading }: any) => (
    <div className="space-y-4">
        <h3 className="text-xl font-bold text-base-content mb-4 flex items-center"><SparklesIcon className="w-6 h-6 mr-2 text-primary" />DLP Generator</h3>
        {/* Form fields here, extracted for clarity */}
        <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
            <InputField id="teacher" label="Teacher" value={dlpForm.teacher} onChange={handleDlpFormChange} required />
            <InputField id="schoolName" label="School Name" value={dlpForm.schoolName} onChange={handleDlpFormChange} required />
             <div>
                <label htmlFor="gradeLevel" className="block text-sm font-medium text-base-content mb-1">Grade Level<span className="text-error">*</span></label>
                <select id="gradeLevel" value={dlpForm.gradeLevel} onChange={handleDlpFormChange} className="w-full bg-base-100 border border-base-300 rounded-md p-2 h-10">
                    {gradeLevels.map(grade => ( <option key={grade} value={grade}>{grade === 'Kindergarten' ? 'Kindergarten' : `Grade ${grade}`}</option> ))}
                </select>
            </div>
             <div><label htmlFor="quarterSelect" className="block text-sm font-medium text-base-content mb-1">Quarter<span className="text-error">*</span></label><select id="quarterSelect" value={dlpForm.quarterSelect} onChange={handleDlpFormChange} className="w-full bg-base-100 border border-base-300 rounded-md p-2 h-10"><option>1ST QUARTER</option><option>2ND QUARTER</option><option>3RD QUARTER</option><option>4TH QUARTER</option></select></div>
        </div>
        <TextAreaField id="classSchedule" label="Class Schedule" value={dlpForm.classSchedule} onChange={handleDlpFormChange} rows={2} required placeholder="e.g., 12:40 - 1:20 PM, G9-Gentleness"/>
        <TextAreaField id="learningCompetency" label="Learning Competency" value={dlpForm.learningCompetency} onChange={handleDlpFormChange} required placeholder="Paste the learning competency here..." />
        <TextAreaField id="lessonObjective" label="Lesson Objective" value={dlpForm.lessonObjective} onChange={handleDlpFormChange} required placeholder="e.g., construct if clauses using the structure of Second Conditionals" />
        <div className="pt-4"><button onClick={generateDLP} disabled={isLoading} className="w-full flex items-center justify-center bg-primary hover:bg-primary-focus text-white font-bold py-3 px-4 rounded-lg"><SparklesIcon className="w-5 h-5 mr-2" />{isLoading ? 'Generating...' : 'Generate Full DLP'}</button></div>
    </div>
);

const DllFormUI = ({ dllForm, handleDllFormChange, dllFormat, setDllFormat, generateDLL, isLoading }: any) => (
    <div className="space-y-4">
        <h3 className="text-xl font-bold text-base-content mb-4 flex items-center"><SparklesIcon className="w-6 h-6 mr-2 text-primary" />Weekly Plan Generator</h3>
        <div className="grid grid-cols-2 gap-4">
            <div>
                <label htmlFor="subject" className="block text-sm font-medium text-base-content mb-1">Subject<span className="text-error">*</span></label>
                <select id="subject" value={dllForm.subject} onChange={handleDllFormChange} className="w-full bg-base-100 border border-base-300 rounded-md p-2 h-10">
                    {Object.entries(subjectAreas).map(([group, subjects]) => (
                        <optgroup label={group} key={group}>
                            {subjects.map(subj => <option key={subj} value={subj}>{subj}</option>)}
                        </optgroup>
                    ))}
                </select>
            </div>
            <div>
                <label htmlFor="gradeLevel" className="block text-sm font-medium text-base-content mb-1">Grade Level<span className="text-error">*</span></label>
                <select id="gradeLevel" value={dllForm.gradeLevel} onChange={handleDllFormChange} className="w-full bg-base-100 border border-base-300 rounded-md p-2 h-10">
                    {gradeLevels.map(grade => ( <option key={grade} value={grade}>{grade === 'Kindergarten' ? 'Kindergarten' : `Grade ${grade}`}</option> ))}
                </select>
            </div>
        </div>
        <InputField id="teachingDates" label="Teaching Dates" value={dllForm.teachingDates} onChange={handleDllFormChange} placeholder="e.g., November 6-10, 2023" />
        <TextAreaField id="weeklyTopic" label="Weekly Topic" value={dllForm.weeklyTopic} onChange={handleDllFormChange} rows={2} placeholder="Main topic for the week" />
        <TextAreaField id="contentStandard" label="Content Standard" value={dllForm.contentStandard} onChange={handleDllFormChange} rows={3} placeholder="Paste the content standard here..." />
        <TextAreaField id="performanceStandard" label="Performance Standard" value={dllForm.performanceStandard} onChange={handleDllFormChange} rows={3} placeholder="Paste the performance standard here..." />

        <div className="grid grid-cols-2 gap-4">
             <div>
                <label htmlFor="dllFormat" className="block text-sm font-medium text-base-content mb-1">DLL Format</label>
                <select id="dllFormat" value={dllFormat} onChange={(e: any) => setDllFormat(e.target.value)} className="w-full bg-base-100 border border-base-300 rounded-md p-2 h-10">
                    <option>Standard</option>
                    <option>4As (Activity, Analysis, Abstraction, Application)</option>
                    <option>5Es (Engage, Explore, Explain, Elaborate, Evaluate)</option>
                </select>
            </div>
            <div>
                <label htmlFor="language" className="block text-sm font-medium text-base-content mb-1">Language</label>
                <select id="language" value={dllForm.language} onChange={handleDllFormChange} className="w-full bg-base-100 border border-base-300 rounded-md p-2 h-10">
                    <option>English</option>
                    <option>Filipino</option>
                </select>
            </div>
        </div>
        <div className="pt-4"><button onClick={generateDLL} disabled={isLoading} className="w-full flex items-center justify-center bg-primary hover:bg-primary-focus text-white font-bold py-3 px-4 rounded-lg"><SparklesIcon className="w-5 h-5 mr-2" />{isLoading ? 'Generating...' : 'Generate Weekly Plan'}</button></div>
    </div>
);

const LasFormUI = ({ lasForm, handleLasFormChange, generateLAS, isLoading }: any) => (
    <div className="space-y-4">
        <h3 className="text-xl font-bold text-base-content mb-4 flex items-center"><ClipboardCheckIcon className="w-6 h-6 mr-2 text-primary" />DLP-Style Learning Activity Sheet</h3>
        <div className="grid grid-cols-2 gap-4">
            <div>
                <label htmlFor="subject" className="block text-sm font-medium text-base-content mb-1">Subject<span className="text-error">*</span></label>
                <select id="subject" value={lasForm.subject} onChange={handleLasFormChange} className="w-full bg-base-100 border border-base-300 rounded-md p-2 h-10">
                    {Object.entries(subjectAreas).map(([group, subjects]) => (
                        <optgroup label={group} key={group}>
                            {subjects.map(subj => <option key={subj} value={subj}>{subj}</option>)}
                        </optgroup>
                    ))}
                </select>
            </div>
            <div>
                <label htmlFor="gradeLevel" className="block text-sm font-medium text-base-content mb-1">Grade Level<span className="text-error">*</span></label>
                <select id="gradeLevel" value={lasForm.gradeLevel} onChange={handleLasFormChange} className="w-full bg-base-100 border border-base-300 rounded-md p-2 h-10">
                    {gradeLevels.map(grade => ( <option key={grade} value={grade}>{grade === 'Kindergarten' ? 'Kindergarten' : `Grade ${grade}`}</option> ))}
                </select>
            </div>
        </div>
        
        <TextAreaField id="learningCompetency" label="Learning Competency" value={lasForm.learningCompetency} onChange={handleLasFormChange} required placeholder="Paste the learning competency here..." />
        
        <TextAreaField id="lessonObjective" label="Lesson Objective" value={lasForm.lessonObjective} onChange={handleLasFormChange} required placeholder="e.g., Identify the parts of a plant" />
        
        <div className="grid grid-cols-2 gap-4">
            <div>
                <label htmlFor="activityType" className="block text-sm font-medium text-base-content mb-1">Activity Focus/Type</label>
                <select id="activityType" value={lasForm.activityType} onChange={handleLasFormChange} className="w-full bg-base-100 border border-base-300 rounded-md p-2 h-10">
                    {activityTypes.map(type => <option key={type} value={type}>{type}</option>)}
                </select>
            </div>
            <div>
                <label htmlFor="language" className="block text-sm font-medium text-base-content mb-1">Language</label>
                <select id="language" value={lasForm.language} onChange={handleLasFormChange} className="w-full bg-base-100 border border-base-300 rounded-md p-2 h-10">
                    <option>English</option>
                    <option>Filipino</option>
                </select>
            </div>
        </div>
        <div className="pt-4"><button onClick={generateLAS} disabled={isLoading} className="w-full flex items-center justify-center bg-primary hover:bg-primary-focus text-white font-bold py-3 px-4 rounded-lg"><SparklesIcon className="w-5 h-5 mr-2" />{isLoading ? 'Generating...' : 'Generate Learning Sheet'}</button></div>
    </div>
);

const QuizFormUI = ({ quizForm, handleQuizFormChange, handleQuizTypeChange, generateQuiz, isLoading }: any) => (
     <form onSubmit={(e) => { e.preventDefault(); generateQuiz(); }} className="space-y-4">
        <h3 className="text-xl font-bold text-base-content mb-4 flex items-center"><SparklesIcon className="w-6 h-6 mr-2 text-primary" />Quiz Generator</h3>
        <InputField id="quizTopic" label="Quiz Topic" value={quizForm.quizTopic} onChange={handleQuizFormChange} required placeholder="e.g., Parts of a Cell" />

        <div className="grid grid-cols-2 gap-4">
            <div>
                <label htmlFor="subject" className="block text-sm font-medium text-base-content mb-1">Subject<span className="text-error">*</span></label>
                <select id="subject" name="subject" value={quizForm.subject} onChange={handleQuizFormChange} className="w-full bg-base-100 border border-base-300 rounded-md p-2 h-10">
                    {Object.entries(subjectAreas).map(([group, subjects]) => (
                        <optgroup label={group} key={group}>
                            {subjects.map(subj => <option key={subj} value={subj}>{subj}</option>)}
                        </optgroup>
                    ))}
                </select>
            </div>
            <div>
                <label htmlFor="gradeLevel" className="block text-sm font-medium text-base-content mb-1">Grade Level<span className="text-error">*</span></label>
                <select id="gradeLevel" name="gradeLevel" value={quizForm.gradeLevel} onChange={handleQuizFormChange} className="w-full bg-base-100 border border-base-300 rounded-md p-2 h-10">
                    {gradeLevels.map(grade => ( <option key={grade} value={grade}>{grade === 'Kindergarten' ? 'Kindergarten' : `Grade ${grade}`}</option>))}
                </select>
            </div>
        </div>
        
        <div>
            <label htmlFor="numQuestions" className="block text-sm font-medium text-base-content mb-1">Number of Questions per Type</label>
            <input type="number" id="numQuestions" value={quizForm.numQuestions} onChange={handleQuizFormChange} className="w-full bg-base-100 border border-base-300 rounded-md p-2 h-10" min="1" max="20" />
        </div>

        <div>
            <label className="block text-sm font-medium text-base-content mb-1">Quiz Formats</label>
            <div className="grid grid-cols-2 gap-2">
                {(['Multiple Choice', 'True or False', 'Identification'] as QuizType[]).map(type => (
                    <label key={type} className="flex items-center gap-2 bg-base-100 p-2 rounded-md">
                        <input
                            type="checkbox"
                            checked={quizForm.quizTypes.includes(type)}
                            onChange={() => handleQuizTypeChange(type)}
                            className="checkbox checkbox-primary"
                        />
                        <span className="text-sm">{type}</span>
                    </label>
                ))}
            </div>
        </div>
        <div className="pt-4"><button type="submit" disabled={isLoading} className="w-full flex items-center justify-center bg-primary hover:bg-primary-focus text-white font-bold py-3 px-4 rounded-lg"><SparklesIcon className="w-5 h-5 mr-2" />{isLoading ? 'Generating...' : 'Generate Quiz'}</button></div>
    </form>
);

const ExamGeneratorForm = ({ examObjectives, onAddObjective, onRemoveObjective, onObjectiveChange, subject, setSubject, gradeLevel, setGradeLevel, quarter, setQuarter, onGenerate, isLoading }: any) => (
    <div className="space-y-4">
        <h3 className="text-xl font-bold text-base-content mb-4 flex items-center"><ClipboardCheckIcon className="w-6 h-6 mr-2 text-primary" />50-Item Exam Generator</h3>
        <div className="grid grid-cols-1 md:grid-cols-4 gap-4">
            <div className="md:col-span-2">
                <InputField id="examSubject" label="Subject" value={subject} onChange={(e: any) => setSubject(e.target.value)} required />
            </div>
            <div>
                <label htmlFor="examGradeLevel" className="block text-sm font-medium text-base-content mb-1">Grade Level<span className="text-error">*</span></label>
                <select id="examGradeLevel" value={gradeLevel} onChange={(e) => setGradeLevel(e.target.value)} className="w-full bg-base-100 border border-base-300 rounded-md p-2 h-10">
                    {gradeLevels.map(grade => ( <option key={grade} value={grade}>{grade === 'Kindergarten' ? 'Kindergarten' : `Grade ${grade}`}</option>))}
                </select>
            </div>
            <div>
                <label htmlFor="examQuarter" className="block text-sm font-medium text-base-content mb-1">Quarter<span className="text-error">*</span></label>
                <select id="examQuarter" value={quarter} onChange={(e) => setQuarter(e.target.value)} className="w-full bg-base-100 border border-base-300 rounded-md p-2 h-10">
                    <option value="1">1st</option>
                    <option value="2">2nd</option>
                    <option value="3">3rd</option>
                    <option value="4">4th</option>
                </select>
            </div>
        </div>
        <div className="space-y-3">
            <label className="block text-sm font-medium text-base-content">Learning Objectives</label>
            {examObjectives.map((obj: ExamObjective, index: number) => (
                <div key={obj.id} className="grid grid-cols-[1fr,80px,auto] gap-2 items-end">
                    <TextAreaField id={`obj-text-${obj.id}`} label={`Objective ${index + 1}`} value={obj.text} onChange={(e: any) => onObjectiveChange(obj.id, 'text', e.target.value)} rows={2} placeholder="Enter a learning objective" required />
                    <InputField id={`obj-days-${obj.id}`} label="Days Taught" value={obj.days} onChange={(e: any) => onObjectiveChange(obj.id, 'days', e.target.value)} type="number" required />
                    <button type="button" onClick={() => onRemoveObjective(obj.id)} className="h-10 px-3 bg-error hover:bg-red-700 text-white font-bold rounded-lg"><TrashIcon className="w-5 h-5"/></button>
                </div>
            ))}
        </div>
        <button type="button" onClick={onAddObjective} className="flex items-center gap-2 text-sm text-primary hover:underline"><PlusIcon className="w-4 h-4" /> Add Another Objective</button>
        <div className="pt-4"><button onClick={onGenerate} disabled={isLoading} className="w-full flex items-center justify-center bg-primary hover:bg-primary-focus text-white font-bold py-3 px-4 rounded-lg"><SparklesIcon className="w-5 h-5 mr-2" />{isLoading ? 'Generating...' : 'Generate Exam'}</button></div>
    </div>
);

const DlpPreview = ({ dlpOutputHtml, handleDownloadDlpDocx, isLoading }: any) => (
    <>
        <div className="p-4 border-b border-base-300 flex justify-between items-center flex-shrink-0"><h3 className="text-xl font-bold">Generated DLP</h3><button onClick={handleDownloadDlpDocx} disabled={isLoading} className="flex items-center bg-secondary hover:bg-secondary-focus text-white font-bold py-2 px-4 rounded-lg"><DownloadIcon className="w-5 h-5 mr-2"/>Download Word File</button></div>
        <div className="p-6 overflow-y-auto flex-grow min-h-0 dlp-preview">
            <div className="overflow-x-auto" dangerouslySetInnerHTML={{ __html: dlpOutputHtml.mainContent }}></div>
            <div className="overflow-x-auto" dangerouslySetInnerHTML={{ __html: dlpOutputHtml.reflectionTableHtml }}></div>
            <div className="overflow-x-auto" dangerouslySetInnerHTML={{ __html: dlpOutputHtml.answerKeyHtml }}></div>
        </div>
    </>
);

const DllPreview = ({ dllContent, dllHeaders, dllDays, handleDownloadDllDocx, isLoading }: any) => (
     <>
        <div className="p-4 border-b border-base-300 flex justify-between items-center flex-shrink-0">
            <h3 className="text-xl font-bold">Generated Weekly Plan Preview</h3>
            <button onClick={handleDownloadDllDocx} disabled={isLoading} className="flex items-center bg-secondary hover:bg-secondary-focus text-white font-bold py-2 px-4 rounded-lg">
                <DownloadIcon className="w-5 h-5 mr-2"/>Download Word File
            </button>
        </div>
        <div className="p-6 overflow-y-auto flex-grow min-h-0 text-sm">
            <div className="space-y-4">
                <div><strong>Content Standard:</strong> {dllContent.contentStandard}</div>
                <div><strong>Performance Standard:</strong> {dllContent.performanceStandard}</div>
                <div><strong>Content:</strong> {dllContent.content}</div>
            </div>
            <div className="overflow-x-auto mt-4">
                <table className="w-full border-collapse">
                    <thead>
                        <tr className="bg-base-300/50">
                            <th className="p-2 border border-base-300 w-1/6"></th>
                            {dllDays.map((day: string) => <th key={day} className="p-2 border border-base-300">{day}</th>)}
                        </tr>
                    </thead>
                    <tbody>
                        <tr>
                            <td className="p-2 border border-base-300 font-bold">{dllHeaders.learningCompetencies}</td>
                            <td className="p-2 border border-base-300">{dllContent.learningCompetencies.monday}</td>
                            <td className="p-2 border border-base-300">{dllContent.learningCompetencies.tuesday}</td>
                            <td className="p-2 border border-base-300">{dllContent.learningCompetencies.wednesday}</td>
                            <td className="p-2 border border-base-300">{dllContent.learningCompetencies.thursday}</td>
                            <td className="p-2 border border-base-300">{dllContent.learningCompetencies.friday}</td>
                        </tr>
                        <tr className="bg-base-300/30"><td colSpan={6} className="p-2 border border-base-300 font-bold">{dllHeaders.procedures}</td></tr>
                        {dllContent.procedures.map((proc: any, index: number) => (
                            <tr key={index}>
                                <td className="p-2 border border-base-300 font-semibold">{proc.procedure}</td>
                                {(['monday', 'tuesday', 'wednesday', 'thursday', 'friday'] as const).map(day => (
                                    <td key={day} className="p-2 border border-base-300 whitespace-pre-wrap">{proc[day]}</td>
                                ))}
                            </tr>
                        ))}
                    </tbody>
                </table>
            </div>
        </div>
    </>
);

const LasPreview = ({ lasContent, settings, lasForm, onDownload }: any) => (
    <>
        <div className="p-4 border-b border-base-300 flex justify-between items-center flex-shrink-0">
            <h3 className="text-xl font-bold">{lasContent.activityTitle}</h3>
            <button onClick={onDownload} className="flex items-center bg-secondary hover:bg-secondary-focus text-white font-bold py-2 px-4 rounded-lg"><DownloadIcon className="w-5 h-5 mr-2"/>Download Word File</button>
        </div>
        <div className="p-4 md:p-6 overflow-y-auto flex-grow min-h-0 bg-base-100">
            <div className="bg-white text-black p-4 font-serif border-2 border-black max-w-4xl mx-auto">
                <header className="flex justify-between items-start mb-2 border-b-2 border-black pb-2">
                    <div className="flex items-center gap-2">
                        {settings.schoolLogo && <img src={settings.schoolLogo} alt="School Logo" className="h-16 w-16 object-contain" />}
                        {settings.secondLogo && <img src={settings.secondLogo} alt="Second Logo" className="h-16 w-16 object-contain" />}
                    </div>
                    <div className="text-center">
                        <p className="font-bold text-lg">Dynamic Learning Program</p>
                        <p className="font-bold text-base">LEARNING ACTIVITY SHEET</p>
                    </div>
                    <div className="text-sm">
                        <div className="border-2 border-black p-1">S.Y. {settings.schoolYear}</div>
                    </div>
                </header>
                <table className="w-full border-collapse border-2 border-black text-sm mb-2">
                    <tbody>
                        <tr>
                            <td className="border border-black p-1 w-2/3"><strong>Name:</strong></td>
                            <td className="border border-black p-1 w-1/3"><strong>Score:</strong></td>
                        </tr>
                         <tr>
                            <td className="border border-black p-1"><strong>Grade & Section:</strong></td>
                            <td className="border border-black p-1"><strong>Date:</strong></td>
                        </tr>
                    </tbody>
                </table>
                <div className="border-2 border-black">
                    <p className="bg-black text-white font-bold p-1">Activity Title: <span className="font-normal">{lasContent.activityTitle}</span></p>
                    <p className="bg-black text-white font-bold p-1">Learning Target: <span className="font-normal">{lasContent.learningTarget}</span></p>
                    <div className="p-2">
                        <p className="font-bold">References: <span className="italic text-xs">(Author, Title, Pages)</span></p>
                        <p className="pl-4">{lasContent.references}</p>
                    </div>
                </div>
                 <div className="mt-4 space-y-4">
                    {lasContent.conceptNotes.map((note: any, index: number) => (
                        <div key={`note-${index}`}>
                            <h4 className="font-bold underline">{note.title}</h4>
                            <p className="whitespace-pre-wrap">{note.content}</p>
                        </div>
                    ))}
                     {lasContent.activities.map((activity: any, index: number) => (
                        <div key={`activity-${index}`}>
                            <h4 className="font-bold text-lg underline">{activity.title}</h4>
                            <p className="italic mb-2">{activity.instructions}</p>
                            {activity.questions && (
                                <ol className="list-decimal list-inside space-y-2">
                                    {activity.questions.map((q: any, qIndex: number) => <li key={qIndex}>{q.questionText}</li>)}
                                </ol>
                            )}
                        </div>
                    ))}
                </div>
            </div>
        </div>
    </>
);

const QuizPreview = ({ quizContent, activityPoints, handleActivityPointsChange, handleGenerateRubric, generatingRubricIndex, handleDownloadQuizDocx, isLoading }: any) => (
    <>
        <div className="p-4 border-b border-base-300 flex justify-between items-center flex-shrink-0">
            <h3 className="text-xl font-bold">{quizContent.quizTitle}</h3>
            <button onClick={handleDownloadQuizDocx} disabled={isLoading} className="flex items-center bg-secondary hover:bg-secondary-focus text-white font-bold py-2 px-4 rounded-lg">
                <DownloadIcon className="w-5 h-5 mr-2"/>Download Word File
            </button>
        </div>
        <div className="p-6 overflow-y-auto flex-grow min-h-0">
            <div className="space-y-8">
                {Object.entries(quizContent.questionsByType).map(([type, section]) => {
                    if (!section) return null;
                    const typedSection = section as GeneratedQuizSection;
                    return (
                        <div key={type}>
                            <h4 className="text-lg font-bold text-primary border-b-2 border-primary mb-2 pb-1">{type}</h4>
                            <p className="text-sm italic mb-4">{typedSection.instructions}</p>
                            <ol className="list-decimal list-inside space-y-4">
                                {typedSection.questions.map((q, i) => (
                                    <li key={i}>
                                        <p>{q.questionText}</p>
                                        {q.options && (
                                            <ul className="list-none pl-6 mt-1 grid grid-cols-1 md:grid-cols-2 gap-x-4">
                                                {q.options.map((opt, oi) => <li key={oi}>{String.fromCharCode(65 + oi)}. {opt}</li>)}
                                            </ul>
                                        )}
                                    </li>
                                ))}
                            </ol>
                        </div>
                    );
                })}
                {quizContent.activities && quizContent.activities.length > 0 && (
                    <div>
                        <h4 className="text-lg font-bold text-primary border-b-2 border-primary mb-2 pb-1">Activities</h4>
                        {quizContent.activities.map((activity: any, index: number) => (
                            <div key={index} className="bg-base-100 p-4 rounded-lg mb-4">
                                <h5 className="font-bold">{activity.activityName}</h5>
                                <p className="text-sm italic my-2">{activity.activityInstructions}</p>
                                {activity.rubric ? (
                                    <div>
                                        <h6 className="text-xs font-bold uppercase mt-2 mb-1">Rubric</h6>
                                        <table className="w-full text-xs"><tbody>{activity.rubric.map((r: DlpRubricItem) => <tr key={r.criteria} className="border-b border-base-300"><td className="py-1">{r.criteria}</td><td className="py-1 text-right font-bold">{r.points} pts</td></tr>)}</tbody></table>
                                    </div>
                                ) : (
                                    <div className="flex items-center gap-2 mt-2">
                                        <input type="number" placeholder="Total Pts" value={activityPoints[index] || ''} onChange={e => handleActivityPointsChange(index, e.target.value)} className="w-24 bg-base-300 p-2 rounded-md h-8 text-sm" />
                                        <button onClick={() => handleGenerateRubric(index)} disabled={generatingRubricIndex === index} className="flex items-center bg-primary hover:bg-primary-focus text-white font-bold py-1 px-3 rounded-lg text-sm"><SparklesIcon className="w-4 h-4 mr-1"/>{generatingRubricIndex === index ? 'Generating...' : 'Generate Rubric'}</button>
                                    </div>
                                )}
                            </div>
                        ))}
                    </div>
                )}
            </div>
        </div>
    </>
);

const ExamPreview = ({ examContent, onDownload, isLoading }: { examContent: GeneratedExam, onDownload: () => void, isLoading: boolean }) => (
    <>
        <div className="p-4 border-b border-base-300 flex justify-between items-center flex-shrink-0">
            <h3 className="text-xl font-bold">{examContent.title}</h3>
            <button onClick={onDownload} disabled={isLoading} className="flex items-center bg-secondary hover:bg-secondary-focus text-white font-bold py-2 px-4 rounded-lg">
                <DownloadIcon className="w-5 h-5 mr-2"/>Download Word File
            </button>
        </div>
        <div className="p-6 overflow-y-auto flex-grow min-h-0">
            <div className="space-y-8">
                <div>
                    <h4 className="text-lg font-bold text-primary mb-2">Table of Specifications</h4>
                    <div className="overflow-x-auto">
                        <table className="w-full border-collapse text-xs">
                            <thead className="bg-base-300/50 text-center">
                                <tr>
                                    <th className="p-2 border border-base-300" rowSpan={2}>Learning Objective</th>
                                    <th className="p-2 border border-base-300" rowSpan={2}>Days</th>
                                    <th className="p-2 border border-base-300" rowSpan={2}>%</th>
                                    <th className="p-2 border border-base-300" rowSpan={2}>Items</th>
                                    <th className="p-2 border border-base-300" colSpan={6}>Cognitive Level</th>
                                    <th className="p-2 border border-base-300" rowSpan={2}>Placement</th>
                                </tr>
                                <tr>
                                    <th className="p-1 border border-base-300 font-normal">Rem</th>
                                    <th className="p-1 border border-base-300 font-normal">Und</th>
                                    <th className="p-1 border border-base-300 font-normal">App</th>
                                    <th className="p-1 border border-base-300 font-normal">Ana</th>
                                    <th className="p-1 border border-base-300 font-normal">Eva</th>
                                    <th className="p-1 border border-base-300 font-normal">Cre</th>
                                </tr>
                            </thead>
                            <tbody>
                                {examContent.tableOfSpecifications.map((item, index) => (
                                    <tr key={index} className="border-b border-base-300">
                                        <td className="p-2 border border-base-300">{item.objective}</td>
                                        <td className="p-2 border border-base-300 text-center">{item.daysTaught}</td>
                                        <td className="p-2 border border-base-300 text-center">{item.percentage}</td>
                                        <td className="p-2 border border-base-300 text-center">{item.numItems}</td>
                                        <td className="p-2 border border-base-300 text-center">{item.remembering}</td>
                                        <td className="p-2 border border-base-300 text-center">{item.understanding}</td>
                                        <td className="p-2 border border-base-300 text-center">{item.applying}</td>
                                        <td className="p-2 border border-base-300 text-center">{item.analyzing}</td>
                                        <td className="p-2 border border-base-300 text-center">{item.evaluating}</td>
                                        <td className="p-2 border border-base-300 text-center">{item.creating}</td>
                                        <td className="p-2 border border-base-300 text-center">{item.itemPlacement}</td>
                                    </tr>
                                ))}
                            </tbody>
                        </table>
                    </div>
                </div>
                <div>
                    <h4 className="text-lg font-bold text-primary mb-2">Test Questions</h4>
                    <ol className="list-decimal list-inside space-y-4 text-sm">
                        {examContent.questions.map((q, i) => (
                            <li key={i}>
                                <p>{q.questionText}</p>
                                {q.options && 
                                    <ul className="list-none pl-6 mt-1 grid grid-cols-1 md:grid-cols-2 gap-x-4">
                                        {q.options.map((opt, oi) => <li key={oi}>{String.fromCharCode(65 + oi)}. {opt}</li>)}
                                    </ul>
                                }
                            </li>
                        ))}
                    </ol>
                </div>
            </div>
        </div>
    </>
);

export default LessonPlanners;