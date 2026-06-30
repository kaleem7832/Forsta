import React, { useState, useRef } from "react";
import {
  Upload,
  FileText,
  Download,
  CheckCircle,
  AlertCircle,
  HelpCircle,
  RefreshCw,
  Plus,
  Trash2,
  Edit3,
  Check,
  Eye,
  Settings,
  Grid as GridIcon,
  ChevronRight,
  Sparkles,
  Info,
  List,
  EyeOff,
  Copy
} from "lucide-react";
import { motion, AnimatePresence } from "motion/react";

interface Question {
  id: string;
  type: string;
  questionText: string;
  title: string;
  instruction: string;
  options: string[];
  gridHeaders?: string[];
  gridRows?: string[];
  tags: string[];
  isIgnored: boolean;
  originalText?: string;
  isScreener?: boolean;
}

const QUESTION_TYPES = [
  { value: "single", label: "Single Choice", instruction: "Please select one response." },
  { value: "multi", label: "Multiple Choice", instruction: "Please select all that apply." },
  { value: "grid", label: "Grid / Matrix", instruction: "Please select one response per row." },
  { value: "ranking", label: "Ranking", instruction: "Please rank all that apply." },
  { value: "numeric", label: "Numeric Input", instruction: "Please provide the answer." },
  { value: "numeric list", label: "Numeric List", instruction: "Please provide the answers." },
  { value: "text", label: "Open Text", instruction: "Please provide the answer." },
  { value: "text list", label: "Open Text List", instruction: "Please provide the answers." },
  { value: "date", label: "Date Input", instruction: "Please provide the date." },
  { value: "info", label: "Informational Text", instruction: "" }
];

export default function App() {
  const [file, setFile] = useState<File | null>(null);
  const [isDragging, setIsDragging] = useState(false);
  const [loading, setLoading] = useState(false);
  const [loadingStep, setLoadingStep] = useState(0);
  const [error, setError] = useState<string | null>(null);
  const [originalName, setOriginalName] = useState<string>("");
  const [questions, setQuestions] = useState<Question[]>([]);
  const [activeTab, setActiveTab] = useState<"builder" | "export">("builder");
  const [expandedQuestionId, setExpandedQuestionId] = useState<string | null>(null);
  const [copiedIndex, setCopiedIndex] = useState<boolean>(false);
  
  // Rules / Cheatsheet drawer state
  const [showCheatsheet, setShowCheatsheet] = useState(false);

  const fileInputRef = useRef<HTMLInputElement>(null);

  const loadingSteps = [
    "Reading file & extracting paragraphs from document...",
    "Sending content to Gemini for structure analysis...",
    "Identifying S-Screeners and Q-Main questions...",
    "Applying special option codes (97, 98, 99) automatically...",
    "Structuring Forsta tagging matrix & instructions...",
    "Finalizing raw preview schema..."
  ];

  const handleDragOver = (e: React.DragEvent) => {
    e.preventDefault();
    setIsDragging(true);
  };

  const handleDragLeave = () => {
    setIsDragging(false);
  };

  const handleDrop = (e: React.DragEvent) => {
    e.preventDefault();
    setIsDragging(false);
    setError(null);
    if (e.dataTransfer.files && e.dataTransfer.files[0]) {
      const droppedFile = e.dataTransfer.files[0];
      if (droppedFile.name.endsWith(".docx")) {
        setFile(droppedFile);
      } else {
        setError("Only Word Documents (.docx) are supported.");
      }
    }
  };

  const handleFileSelect = (e: React.ChangeEvent<HTMLInputElement>) => {
    setError(null);
    if (e.target.files && e.target.files[0]) {
      const selectedFile = e.target.files[0];
      if (selectedFile.name.endsWith(".docx")) {
        setFile(selectedFile);
      } else {
        setError("Only Word Documents (.docx) are supported.");
      }
    }
  };

  const parseDocument = async () => {
    if (!file) return;
    setLoading(true);
    setLoadingStep(0);
    setError(null);

    // Simulate stepping through loader for smooth visual feedback
    const interval = setInterval(() => {
      setLoadingStep((prev) => {
        if (prev < loadingSteps.length - 1) return prev + 1;
        return prev;
      });
    }, 2800);

    try {
      const formData = new FormData();
      formData.append("file", file);

      const response = await fetch("/api/parse-questionnaire", {
        method: "POST",
        body: formData,
      });

      clearInterval(interval);

      const contentType = response.headers.get("content-type");
      if (!contentType || !contentType.includes("application/json")) {
        const textResponse = await response.text();
        console.error("Non-JSON Server response:", textResponse);
        
        if (response.status === 413) {
          throw new Error("The uploaded file is too large. Please select a smaller Word Document.");
        }
        
        throw new Error("The server returned an invalid response. This usually happens if the backend is restarting or your Gemini API key is missing. Please verify your API keys in the Settings menu and try again.");
      }

      const data = await response.json();

      if (!response.ok || !data.success) {
        throw new Error(data.error || "Failed to process the document.");
      }

      setQuestions(data.questions);
      setOriginalName(data.originalName);
      if (data.questions.length > 0) {
        setExpandedQuestionId(data.questions[0].id);
      }
    } catch (err: any) {
      clearInterval(interval);
      console.error(err);
      setError(err.message || "An error occurred while connecting to the server.");
    } finally {
      setLoading(false);
    }
  };

  const downloadDocx = async () => {
    try {
      const response = await fetch("/api/generate-docx", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          originalName: originalName || file?.name || "survey_questionnaire.docx",
          questions,
        }),
      });

      if (!response.ok) {
        throw new Error("Failed to generate and download document.");
      }

      const filenameHeader = response.headers.get("X-Filename") || "tagged_questionnaire.docx";
      const blob = await response.blob();
      const url = window.URL.createObjectURL(blob);
      const link = document.createElement("a");
      link.href = url;
      link.setAttribute("download", filenameHeader);
      document.body.appendChild(link);
      link.click();
      link.parentNode?.removeChild(link);
    } catch (err: any) {
      alert("Error: " + err.message);
    }
  };

  // Helper to construct the plain text block representation
  const getForstaMarkup = (q: Question) => {
    if (q.isIgnored) {
      return `[[begin ignore]]\n${q.id}. ${q.questionText} (Ignored Question)\n[[end ignore]]`;
    }

    let result = "";
    const hasRandomize = q.tags.includes("randomize");
    const typeLabel = hasRandomize ? `${q.type}, randomize` : q.type;

    // Head
    result += `${q.id}. ${q.questionText} [[${typeLabel}]]\n`;
    
    // Title Line
    const extraTags = q.tags.filter(t => t !== q.type && t !== "randomize" && t !== "required" && t !== "not required");
    const tagsStr = extraTags.length > 0 ? " " + extraTags.map(t => `[[${t}]]`).join(" ") : "";
    result += `${q.id} [[title]]${tagsStr}\n`;

    // Instruction Line
    if (q.instruction) {
      result += `${q.instruction} [[instruction]]\n`;
    }

    if (q.type === "grid") {
      const h = q.gridHeaders || [];
      const r = q.gridRows || [];
      result += `||${h.join("|")}|\n`;
      result += `|${h.map(() => "-").join("|")}|\n`;
      r.forEach(row => {
        result += `|${row}|${h.map(() => "").join("|")}|\n`;
      });
    } else {
      q.options.forEach(opt => {
        let cleanOpt = opt.trim();
        if (cleanOpt.startsWith("* ")) {
          cleanOpt = cleanOpt.substring(2).trim();
        } else if (cleanOpt.startsWith("*")) {
          cleanOpt = cleanOpt.substring(1).trim();
        } else if (cleanOpt.startsWith("- ")) {
          cleanOpt = cleanOpt.substring(2).trim();
        } else if (cleanOpt.startsWith("-")) {
          cleanOpt = cleanOpt.substring(1).trim();
        }
        result += `- ${cleanOpt}\n`;
      });
    }

    return result;
  };

  const getCombinedTextMarkup = () => {
    return questions
      .filter((q) => !q.isIgnored)
      .map((q) => getForstaMarkup(q))
      .join("\n\n");
  };

  const copyToClipboard = () => {
    const text = getCombinedTextMarkup();
    navigator.clipboard.writeText(text);
    setCopiedIndex(true);
    setTimeout(() => setCopiedIndex(false), 2000);
  };

  // Edit Question Fields in State
  const updateQuestionField = (id: string, field: keyof Question, value: any) => {
    setQuestions(prev =>
      prev.map(q => {
        if (q.id === id) {
          const updated = { ...q, [field]: value };
          // If type changes, automatically set default description helper
          if (field === "type") {
            const mappedType = QUESTION_TYPES.find(qt => qt.value === value);
            if (mappedType) {
              updated.instruction = mappedType.instruction;
            }
          }
          return updated;
        }
        return q;
      })
    );
  };

  const handleOptionChange = (qId: string, optIndex: number, newVal: string) => {
    setQuestions(prev =>
      prev.map(q => {
        if (q.id === qId) {
          const updatedOptions = [...q.options];
          updatedOptions[optIndex] = newVal;
          return { ...q, options: updatedOptions };
        }
        return q;
      })
    );
  };

  const deleteOption = (qId: string, optIndex: number) => {
    setQuestions(prev =>
      prev.map(q => {
        if (q.id === qId) {
          const updatedOptions = q.options.filter((_, i) => i !== optIndex);
          return { ...q, options: updatedOptions };
        }
        return q;
      })
    );
  };

  const addOption = (qId: string) => {
    setQuestions(prev =>
      prev.map(q => {
        if (q.id === qId) {
          return { ...q, options: [...q.options, "New Option Choice"] };
        }
        return q;
      })
    );
  };

  // Adding new Custom Blank Question
  const addNewQuestion = () => {
    const nextNum = questions.length + 1;
    const newQ: Question = {
      id: `Q${nextNum}`,
      type: "single",
      questionText: `What is your opinion on survey question ${nextNum}?`,
      title: `Q${nextNum}`,
      instruction: "Please select one response.",
      options: ["Option 1", "Option 2", "Other (please specify) [[97, other]]"],
      tags: ["single", "required"],
      isIgnored: false,
    };
    setQuestions([...questions, newQ]);
    setExpandedQuestionId(newQ.id);
  };

  const resetUpload = () => {
    setFile(null);
    setQuestions([]);
    setError(null);
  };

  return (
    <div className="min-h-screen bg-slate-50 font-sans text-slate-800 antialiased flex flex-col">
      
      {/* Decorative top strip */}
      <div className="h-1.5 w-full bg-gradient-to-r from-cyan-500 via-indigo-600 to-emerald-500" />

      {/* Main Elegant Navigation */}
      <header className="bg-white border-b border-slate-200 py-4 px-6 md:px-12 sticky top-0 z-45 shadow-sm flex items-center justify-between">
        <div className="flex items-center gap-3">
          <div className="p-2.5 bg-indigo-50 text-indigo-600 rounded-xl border border-indigo-100 shadow-sm">
            <Sparkles className="w-6 h-6 animate-pulse" />
          </div>
          <div>
            <span className="text-xs uppercase font-bold tracking-widest text-indigo-500">Forsta Assistant</span>
            <h1 className="font-sans font-bold text-lg text-slate-900 tracking-tight flex items-center gap-2">
              Survey Questionnaire Tagging Studio
            </h1>
          </div>
        </div>

        <div className="flex items-center gap-2.5">
          <button
            onClick={() => setShowCheatsheet(true)}
            className="flex items-center gap-2 px-3.5 py-2 text-sm font-medium text-slate-600 hover:text-slate-900 bg-slate-100 hover:bg-slate-200 rounded-lg transition-all cursor-pointer"
          >
            <HelpCircle className="w-4 h-4 text-indigo-600" />
            <span className="hidden sm:inline">Tag Rules Cheatsheet</span>
          </button>
          
          {questions.length > 0 && (
            <button
              onClick={resetUpload}
              className="flex items-center gap-1.5 px-3.5 py-2 text-sm font-medium text-slate-600 hover:text-indigo-600 hover:bg-indigo-50 rounded-lg transition-all cursor-pointer"
            >
              <RefreshCw className="w-4 h-4" />
              <span>Reset</span>
            </button>
          )}
        </div>
      </header>

      {/* Primary Workspace */}
      <main className="flex-1 max-w-7xl w-full mx-auto p-4 md:p-8 flex flex-col gap-6">

        {/* Informative Help Alert banner if no document uploaded */}
        {!file && (
          <div className="bg-gradient-to-r from-slate-900 to-indigo-950 text-white rounded-2xl p-6 shadow-xl border border-slate-800 overflow-hidden relative">
            <div className="absolute top-0 right-0 p-8 opacity-10 pointer-events-none">
              <Sparkles className="w-40 h-40" />
            </div>
            <div className="max-w-2xl">
              <span className="inline-block px-3 py-1 bg-indigo-500/30 border border-indigo-400/20 text-indigo-300 rounded-full text-xs font-semibold mb-3">
                Powered by Gemini-3.5-Flash
              </span>
              <h2 className="text-xl md:text-2xl font-bold mb-2 tracking-tight line-height-relaxed">
                Turn standard Word Questionnaires into formatted Forsta Tagged Documents automatically.
              </h2>
              <p className="text-sm text-slate-300 leading-relaxed">
                Upload your word document questionnaire. Our specialized AI parsing rules automatically separate Screener and Main questions, identify question types, formulate proper instructions, assign unique question titles, inject custom option codes (like 97, 98, NA codes), and apply randomize parameters.
              </p>
            </div>
          </div>
        )}

        {/* Step 1: File Selection & Analysis */}
        {!questions.length && !loading && (
          <div className="grid grid-cols-1 md:grid-cols-3 gap-6">
            
            {/* Drag Drop Area */}
            <div className="md:col-span-2">
              <div
                onDragOver={handleDragOver}
                onDragLeave={handleDragLeave}
                onDrop={handleDrop}
                onClick={() => fileInputRef.current?.click()}
                className={`border-2 border-dashed rounded-2xl p-12 text-center cursor-pointer transition-all duration-300 min-h-[350px] flex flex-col items-center justify-center bg-white shadow-sm ${
                  isDragging
                    ? "border-indigo-600 bg-indigo-50/50 scale-[1.01]"
                    : "border-slate-300 hover:border-indigo-400 hover:shadow-md"
                }`}
              >
                <input
                  type="file"
                  ref={fileInputRef}
                  onChange={handleFileSelect}
                  accept=".docx"
                  className="hidden"
                />
                <div className={`p-5 rounded-2xl mb-4 transition-all duration-300 ${isDragging ? "bg-indigo-600 text-white" : "bg-slate-100 text-slate-500"}`}>
                  <Upload className="w-10 h-10" />
                </div>
                
                {file ? (
                  <div className="space-y-2">
                    <p className="text-xs uppercase tracking-wider font-bold text-indigo-500">Ready to Analyze</p>
                    <h3 className="font-bold text-slate-900 text-lg flex items-center justify-center gap-1.5">
                      <FileText className="w-5 h-5 text-indigo-600" />
                      {file.name}
                    </h3>
                    <p className="text-xs text-slate-400 font-mono font-bold">Size: {(file.size / 1024).toFixed(1)} KB</p>
                  </div>
                ) : (
                  <div>
                    <h3 className="font-bold text-slate-900 text-lg mb-1">
                      Drag & drop your Word Document here
                    </h3>
                    <p className="text-sm text-slate-500 max-w-sm mx-auto mb-4 leading-relaxed">
                      Please upload only .docx files. The AI parser will automatically scan and format standard survey questions.
                    </p>
                    <span className="inline-flex py-1.5 px-4 bg-indigo-50 text-indigo-600 rounded-lg text-sm font-semibold hover:bg-indigo-100 transition-colors">
                      Browse Files
                    </span>
                  </div>
                )}
              </div>

              {/* Error Box */}
              {error && (
                <div className="mt-4 p-4 bg-rose-50 border border-rose-150 rounded-xl text-rose-800 flex items-start gap-3">
                  <AlertCircle className="w-5 h-5 text-rose-500 shrink-0 mt-0.5" />
                  <div>
                    <h4 className="font-bold text-sm">Processing Error</h4>
                    <p className="text-xs text-rose-700 leading-relaxed mt-0.5">{error}</p>
                  </div>
                </div>
              )}

              {/* Action Button */}
              {file && (
                <div className="mt-6 flex justify-end">
                  <button
                    onClick={parseDocument}
                    className="flex items-center gap-2 px-6 py-3 bg-indigo-600 hover:bg-indigo-700 text-white font-bold rounded-xl shadow-lg hover:shadow-indigo-500/20 hover:-translate-y-0.5 transition-all w-full md:w-auto justify-center cursor-pointer"
                  >
                    <Sparkles className="w-5 h-5" />
                    <span>Run AI Tagging Engine</span>
                  </button>
                </div>
              )}
            </div>

            {/* Quick Rules & Guidelines Checklist */}
            <div className="bg-white rounded-2xl p-6 border border-slate-200 shadow-sm flex flex-col justify-between">
              <div>
                <h3 className="font-sans font-bold text-slate-900 mb-4 flex items-center gap-1.5 border-b border-slate-100 pb-3 text-base">
                  <Info className="w-5 h-5 text-indigo-600" />
                  Predefined AI Parser Rules
                </h3>
                
                <ul className="space-y-3.5 text-sm text-slate-600 leading-tight">
                  <li className="flex items-start gap-2.5">
                    <CheckCircle className="w-4 h-4 text-emerald-500 shrink-0 mt-0.5" />
                    <span>Starts from <strong>screener questions</strong> and groups appropriately.</span>
                  </li>
                  <li className="flex items-start gap-2.5">
                    <CheckCircle className="w-4 h-4 text-emerald-500 shrink-0 mt-0.5" />
                    <span>Prepends <strong>"S"</strong> for Screeners and <strong>"Q"</strong> for Main Questions.</span>
                  </li>
                  <li className="flex items-start gap-2.5">
                    <CheckCircle className="w-4 h-4 text-emerald-500 shrink-0 mt-0.5" />
                    <span>Converts decimal numbering like 1.2 to <strong>Q1x2 / S1x2</strong>.</span>
                  </li>
                  <li className="flex items-start gap-2.5">
                    <CheckCircle className="w-4 h-4 text-emerald-500 shrink-0 mt-0.5" />
                    <span>Appends Forsta instructions based on mapped question types.</span>
                  </li>
                  <li className="flex items-start gap-2.5">
                    <CheckCircle className="w-4 h-4 text-emerald-500 shrink-0 mt-0.5" />
                    <span>Identifies and inserts <strong>97, 98, and 99 special option codes</strong> automatically.</span>
                  </li>
                  <li className="flex items-start gap-2.5">
                    <CheckCircle className="w-4 h-4 text-emerald-500 shrink-0 mt-0.5" />
                    <span>Ignores questions that contain <strong>"answer per column"</strong> criteria.</span>
                  </li>
                </ul>
              </div>

              <div className="mt-6 pt-4 border-t border-slate-100">
                <p className="text-xs text-slate-400">
                  All formatting is built to comply precisely with Forsta's official smart survey importing parser specs.
                </p>
              </div>
            </div>
          </div>
        )}

        {/* Loading state spinner */}
        {loading && (
          <div className="bg-white border border-slate-200 rounded-3xl p-16 shadow-xl flex flex-col items-center justify-center max-w-xl mx-auto text-center my-12 relative overflow-hidden">
            <div className="absolute top-0 right-0 w-32 h-32 bg-indigo-50/50 rounded-full blur-2xl -z-10" />
            <div className="relative mb-6">
              <div className="w-16 h-16 border-4 border-slate-100 border-t-indigo-600 rounded-full animate-spin" />
              <div className="absolute top-1/2 left-1/2 -translate-x-1/2 -translate-y-1/2 text-indigo-650">
                <Sparkles className="w-6 h-6 animate-pulse" />
              </div>
            </div>
            
            <h3 className="font-bold text-slate-900 text-xl tracking-tight mb-2">Analyzing Questionnaire</h3>
            <p className="text-sm text-slate-500 max-w-xs mb-6 text-center leading-normal">
              Our AI is applying Forsta parsing regulations onto your survey structure.
            </p>
            
            <div className="w-full bg-slate-100 h-1.5 rounded-full overflow-hidden mb-3">
              <motion.div
                className="bg-indigo-600 h-full rounded-full"
                animate={{ width: `${((loadingStep + 1) / loadingSteps.length) * 100}%` }}
                transition={{ duration: 0.5 }}
              />
            </div>
            <p className="text-xs text-indigo-650 font-mono font-semibold h-4 animate-pulse">
              {loadingSteps[loadingStep]}
            </p>
          </div>
        )}

        {/* Step 2: Main Workspace Panel once loaded */}
        {questions.length > 0 && (
          <div className="flex flex-col gap-6">

            {/* Quick Status / Control Summary */}
            <div className="bg-white rounded-2xl p-4 md:p-6 border border-slate-200 shadow-sm flex flex-col md:flex-row items-center justify-between gap-4">
              <div className="flex items-center gap-3">
                <div className="p-2 bg-emerald-50 border border-emerald-150 rounded-xl text-emerald-600 shadow-sm">
                  <CheckCircle className="w-6 h-6" />
                </div>
                <div>
                  <p className="text-xs text-slate-400 font-semibold tracking-wider uppercase mb-0.5">Analysis Complete</p>
                  <h2 className="font-bold text-slate-800 text-base md:text-lg">
                    {questions.length} questions mapped & tagged
                  </h2>
                </div>
              </div>

              {/* Action Buttons */}
              <div className="flex items-center gap-3 w-full md:w-auto">
                <button
                  onClick={addNewQuestion}
                  className="flex items-center gap-1.5 px-4 py-2.5 bg-slate-100 hover:bg-slate-200 text-slate-700 hover:text-slate-900 font-semibold rounded-xl text-sm transition-all cursor-pointer"
                >
                  <Plus className="w-4 h-4" />
                  <span>Add Question</span>
                </button>
                <button
                  onClick={downloadDocx}
                  className="flex items-center gap-2 px-5 py-2.5 bg-emerald-600 hover:bg-emerald-700 text-white font-bold rounded-xl text-sm shadow-md hover:shadow-emerald-500/10 hover:-translate-y-0.5 transition-all cursor-pointer"
                >
                  <Download className="w-4 h-4" />
                  <span>Download Tagged .docx</span>
                </button>
              </div>
            </div>

            {/* Layout tabs to toggle between Visual Builder Grid / Pure Text Export Markup */}
            <div className="flex gap-4 border-b border-slate-200">
              <button
                onClick={() => setActiveTab("builder")}
                className={`py-3 px-1.5 font-bold text-sm tracking-tight border-b-2 flex items-center gap-2 transition-all cursor-pointer ${
                  activeTab === "builder"
                    ? "border-indigo-600 text-indigo-600"
                    : "border-transparent text-slate-500 hover:text-slate-800"
                }`}
              >
                <Settings className="w-4 h-4" />
                <span>Verify & Edit Question Segments</span>
              </button>
              <button
                onClick={() => setActiveTab("export")}
                className={`py-3 px-1.5 font-bold text-sm tracking-tight border-b-2 flex items-center gap-2 transition-all cursor-pointer ${
                  activeTab === "export"
                    ? "border-indigo-600 text-indigo-600"
                    : "border-transparent text-slate-500 hover:text-slate-800"
                }`}
              >
                <Eye className="w-4 h-4" />
                <span>Raw Program Syntax Output</span>
              </button>
            </div>

            {/* Content Tabs */}
            {activeTab === "builder" ? (
              <div className="grid grid-cols-1 lg:grid-cols-12 gap-6">
                
                {/* Visual Questions List */}
                <div className="lg:col-span-8 flex flex-col gap-4">
                  {questions.map((q, qIndex) => {
                    const isExpanded = expandedQuestionId === q.id;
                    const questionTypeInfo = QUESTION_TYPES.find(qt => qt.value === q.type);

                    return (
                      <div
                        key={q.id}
                        className={`bg-white rounded-2xl border transition-all duration-200 overflow-hidden ${
                          q.isIgnored
                            ? "opacity-55 border-slate-200 bg-slate-50/70"
                            : isExpanded
                            ? "border-indigo-500 ring-2 ring-indigo-50 shadow-md"
                            : "border-slate-200 hover:border-slate-300 shadow-sm"
                        }`}
                      >
                        {/* Header Area */}
                        <div
                          onClick={() => setExpandedQuestionId(isExpanded ? null : q.id)}
                          className="p-4 md:p-5 flex items-center justify-between cursor-pointer select-none"
                        >
                          <div className="flex items-center gap-3 min-w-0 flex-1">
                            {/* Question ID tag */}
                            <span className="flex-shrink-0 inline-flex items-center justify-center font-mono font-bold text-xs bg-slate-100 px-2.5 py-1 rounded-md text-slate-700 capitalize border border-slate-200">
                              {q.id}
                            </span>
                            
                            {/* Question text snippet */}
                            <h4 className={`font-semibold text-slate-800 truncate text-sm md:text-base ${q.isIgnored ? "line-through text-slate-400" : ""}`}>
                              {q.questionText || "(No question text defined)"}
                            </h4>

                            {q.isIgnored && (
                              <span className="inline-flex text-xs px-2 py-0.5 rounded-full bg-rose-50 border border-rose-100 text-rose-600 font-semibold align-middle">
                                Ignored
                              </span>
                            )}
                          </div>

                          <div className="flex items-center gap-2 flex-shrink-0 ml-3">
                            {/* Question type tag pill */}
                            {!q.isIgnored && (
                              <span className="hidden sm:inline-flex items-center text-xs px-2.5 py-1 rounded-full bg-indigo-50 border border-indigo-100 text-indigo-700 font-medium">
                                {questionTypeInfo?.label || q.type}
                              </span>
                            )}
                            <ChevronRight className={`w-5 h-5 text-slate-400 transition-transform ${isExpanded ? "rotate-90" : ""}`} />
                          </div>
                        </div>

                        {/* Expandable Details Form */}
                        <AnimatePresence>
                          {isExpanded && (
                            <motion.div
                              initial={{ height: 0, opacity: 0 }}
                              animate={{ height: "auto", opacity: 1 }}
                              exit={{ height: 0, opacity: 0 }}
                              transition={{ duration: 0.2 }}
                              className="border-t border-slate-100 bg-white"
                            >
                              <div className="p-5 flex flex-col gap-5">
                                
                                {/* Question Definition Fields */}
                                <div className="grid grid-cols-1 md:grid-cols-12 gap-4">
                                  
                                  {/* Code / ID */}
                                  <div className="md:col-span-3">
                                    <label className="block text-xs font-bold uppercase text-slate-500 mb-1.5 font-mono">
                                      Identifier Code (ID)
                                    </label>
                                    <input
                                      type="text"
                                      value={q.id}
                                      onChange={(e) => updateQuestionField(q.id, "id", e.target.value)}
                                      className="w-full px-3.5 py-2 text-sm border border-slate-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-indigo-500 focus:border-indigo-500 font-mono"
                                    />
                                  </div>

                                  {/* Type Select */}
                                  <div className="md:col-span-4">
                                    <label className="block text-xs font-bold uppercase text-slate-500 mb-1.5 font-mono">
                                      Question Type
                                    </label>
                                    <select
                                      value={q.type}
                                      onChange={(e) => updateQuestionField(q.id, "type", e.target.value)}
                                      className="w-full px-3.5 py-2 text-sm border border-slate-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-indigo-500 focus:border-indigo-500"
                                    >
                                      {QUESTION_TYPES.map((qt) => (
                                        <option key={qt.value} value={qt.value}>
                                          {qt.label}
                                        </option>
                                      ))}
                                    </select>
                                  </div>

                                  {/* Active properties / isIgnored switch */}
                                  <div className="md:col-span-5 flex items-end gap-3 pb-0.5">
                                    <button
                                      onClick={() => updateQuestionField(q.id, "isIgnored", !q.isIgnored)}
                                      className={`flex-1 flex items-center justify-center gap-1.5 py-2 px-3 border rounded-lg text-xs font-bold transition-all cursor-pointer ${
                                        q.isIgnored
                                          ? "bg-rose-50 border-rose-300 text-rose-700"
                                          : "bg-slate-50 border-slate-200 text-slate-600 hover:bg-slate-100"
                                      }`}
                                    >
                                      {q.isIgnored ? <EyeOff className="w-4 h-4" /> : <Eye className="w-4 h-4" />}
                                      <span>{q.isIgnored ? "Ignored in Export" : "Export enabled"}</span>
                                    </button>

                                    {/* Randomize tag switch */}
                                    <button
                                      onClick={() => {
                                        const currentTags = q.tags || [];
                                        const hasRand = currentTags.includes("randomize");
                                        const nextTags = hasRand
                                          ? currentTags.filter((t) => t !== "randomize")
                                          : [...currentTags, "randomize"];
                                        updateQuestionField(q.id, "tags", nextTags);
                                      }}
                                      className={`flex-1 flex items-center justify-center gap-1.5 py-2 px-3 border rounded-lg text-xs font-bold transition-all cursor-pointer ${
                                        q.tags?.includes("randomize")
                                          ? "bg-amber-50 border-amber-300 text-amber-700 font-extrabold shadow-sm ring-1 ring-amber-400/30"
                                          : "bg-slate-50 border-slate-200 text-slate-600 hover:bg-slate-100"
                                      }`}
                                    >
                                      <Sparkles className="w-4 h-4" />
                                      <span>Randomize options</span>
                                    </button>
                                  </div>
                                </div>

                                {/* Question Label Text Input */}
                                <div>
                                  <label className="block text-xs font-bold uppercase text-slate-500 mb-1.5 font-mono">
                                    Question Text Sentence
                                  </label>
                                  <textarea
                                    value={q.questionText}
                                    rows={2}
                                    onChange={(e) => updateQuestionField(q.id, "questionText", e.target.value)}
                                    className="w-full px-3.5 py-2 text-sm border border-slate-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-indigo-500 focus:border-indigo-500 font-sans"
                                  />
                                </div>

                                {/* Custom Program Instruction Input */}
                                <div>
                                  <label className="block text-xs font-bold uppercase text-slate-500 mb-1.5 font-mono">
                                    Predefined Instruction Text (Forsta Target)
                                  </label>
                                  <input
                                    type="text"
                                    value={q.instruction}
                                    onChange={(e) => updateQuestionField(q.id, "instruction", e.target.value)}
                                    className="w-full px-3.5 py-2 text-sm border border-slate-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-indigo-500 focus:border-indigo-500"
                                  />
                                </div>

                                {/* If Grid: special scale columns/prompts builder */}
                                {q.type === "grid" ? (
                                  <div className="grid grid-cols-1 md:grid-cols-2 gap-5 p-4 bg-slate-50 rounded-xl border border-slate-150">
                                    
                                    {/* Grid scale columns */}
                                    <div>
                                      <div className="flex items-center justify-between mb-2">
                                        <label className="text-xs font-bold uppercase text-slate-500 font-mono">
                                          Grid Columns (Scales)
                                        </label>
                                        <button
                                          onClick={() => {
                                            const cols = q.gridHeaders || [];
                                            updateQuestionField(q.id, "gridHeaders", [...cols, `Scale ${cols.length + 1}`]);
                                          }}
                                          className="text-xs text-indigo-650 hover:text-indigo-850 font-bold cursor-pointer"
                                        >
                                          + Add Column
                                        </button>
                                      </div>
                                      <div className="space-y-2">
                                        {(q.gridHeaders || []).map((header, hIdx) => (
                                          <div key={hIdx} className="flex items-center gap-2">
                                            <input
                                              type="text"
                                              value={header}
                                              onChange={(e) => {
                                                const newHeaders = [...(q.gridHeaders || [])];
                                                newHeaders[hIdx] = e.target.value;
                                                updateQuestionField(q.id, "gridHeaders", newHeaders);
                                              }}
                                              className="flex-1 px-3 py-1.5 text-xs border border-slate-300 bg-white rounded-md focus:outline-none"
                                            />
                                            <button
                                              onClick={() => {
                                                const newHeaders = (q.gridHeaders || []).filter((_, idx) => idx !== hIdx);
                                                updateQuestionField(q.id, "gridHeaders", newHeaders);
                                              }}
                                              className="text-slate-400 hover:text-rose-600 cursor-pointer p-1"
                                            >
                                              <Trash2 className="w-4 h-4" />
                                            </button>
                                          </div>
                                        ))}
                                      </div>
                                    </div>

                                    {/* Grid prompts / rows */}
                                    <div>
                                      <div className="flex items-center justify-between mb-2">
                                        <label className="text-xs font-bold uppercase text-slate-500 font-mono">
                                          Grid Rows (Prompt Items)
                                        </label>
                                        <button
                                          onClick={() => {
                                            const r = q.gridRows || [];
                                            updateQuestionField(q.id, "gridRows", [...r, `Brand Prompt ${r.length + 1}`]);
                                          }}
                                          className="text-xs text-indigo-650 hover:text-indigo-850 font-bold cursor-pointer"
                                        >
                                          + Add Row
                                        </button>
                                      </div>
                                      <div className="space-y-2">
                                        {(q.gridRows || []).map((rowVal, rIdx) => (
                                          <div key={rIdx} className="flex items-center gap-2">
                                            <input
                                              type="text"
                                              value={rowVal}
                                              onChange={(e) => {
                                                const newRows = [...(q.gridRows || [])];
                                                newRows[rIdx] = e.target.value;
                                                updateQuestionField(q.id, "gridRows", newRows);
                                              }}
                                              className="flex-1 px-3 py-1.5 text-xs border border-slate-300 bg-white rounded-md focus:outline-none"
                                            />
                                            <button
                                              onClick={() => {
                                                const newRows = (q.gridRows || []).filter((_, idx) => idx !== rIdx);
                                                updateQuestionField(q.id, "gridRows", newRows);
                                              }}
                                              className="text-slate-400 hover:text-rose-600 cursor-pointer p-1"
                                            >
                                              <Trash2 className="w-4 h-4" />
                                            </button>
                                          </div>
                                        ))}
                                      </div>
                                    </div>
                                  </div>
                                ) : (
                                  /* Categorical Choice Option List Builder */
                                  <div>
                                    <div className="flex items-center justify-between border-b border-slate-100 pb-2 mb-3">
                                      <label className="text-xs font-bold uppercase text-slate-500 font-mono">
                                        Answer Categories / Options list
                                      </label>
                                      <button
                                        onClick={() => addOption(q.id)}
                                        className="inline-flex items-center gap-1 px-3 py-1 bg-indigo-50 text-indigo-750 text-xs font-bold rounded-md hover:bg-indigo-150 transition-colors cursor-pointer"
                                      >
                                        <Plus className="w-3.5 h-3.5" />
                                        <span>Add Answer Choice</span>
                                      </button>
                                    </div>

                                    <div className="space-y-2.5 max-h-[300px] overflow-y-auto pr-1">
                                      {q.options.map((opt, oIdx) => (
                                        <div key={oIdx} className="flex items-center gap-2">
                                          <span className="text-slate-400 font-mono text-xs select-none">
                                            •
                                          </span>
                                          <input
                                            type="text"
                                            value={opt}
                                            onChange={(e) => handleOptionChange(q.id, oIdx, e.target.value)}
                                            className="flex-1 px-3.5 py-2 text-sm border border-slate-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-indigo-500 bg-white"
                                            placeholder="Write customized answer option text with code label tags..."
                                          />
                                          <button
                                            onClick={() => deleteOption(q.id, oIdx)}
                                            className="p-2 text-slate-400 hover:text-rose-600 hover:bg-rose-50 rounded-lg transition-colors cursor-pointer"
                                            title="Delete Option"
                                          >
                                            <Trash2 className="w-4 h-4" />
                                          </button>
                                        </div>
                                      ))}

                                      {!q.options.length && (
                                        <p className="text-xs italic text-slate-400 py-3 text-center">No custom answer categories defined yet.</p>
                                      )}
                                    </div>
                                  </div>
                                )}

                                {/* Delete row Question Button */}
                                <div className="border-t border-slate-100 pt-3.5 mt-2 flex justify-between items-center">
                                  <p className="text-xs text-slate-400 italic">
                                    {q.originalText ? "Identified from original questionnaire." : "Added custom survey element."}
                                  </p>
                                  <button
                                    onClick={() => {
                                      if (confirm("Are you sure you want to remove this question?")) {
                                        setQuestions(prev => prev.filter(item => item.id !== q.id));
                                      }
                                    }}
                                    className="inline-flex items-center gap-1 text-xs text-rose-500 hover:text-rose-700 font-bold cursor-pointer"
                                  >
                                    <Trash2 className="w-3.5 h-3.5" />
                                    <span>Delete Question Block</span>
                                  </button>
                                </div>

                              </div>
                            </motion.div>
                          )}
                        </AnimatePresence>
                      </div>
                    );
                  })}
                </div>

                {/* Right Side: Quick Segment Summary Panel */}
                <div className="lg:col-span-4 flex flex-col gap-6">
                  
                  {/* Status Card & Overview */}
                  <div className="bg-white rounded-2xl p-6 border border-slate-200 shadow-sm">
                    <h3 className="font-bold text-slate-900 border-b border-slate-100 pb-3 mb-4 text-base">
                      Interactive Workspace Summary
                    </h3>

                    <div className="space-y-4">
                      
                      <div className="flex items-center justify-between pb-3 border-b border-dotted border-slate-100">
                        <span className="text-sm font-semibold text-slate-500">Total Segments detected</span>
                        <span className="font-mono text-sm font-bold px-2 py-0.5 bg-slate-100 rounded-md">
                          {questions.length}
                        </span>
                      </div>

                      <div className="flex items-center justify-between pb-3 border-b border-dotted border-slate-100">
                        <span className="text-sm font-semibold text-slate-500">Screeners (Group S)</span>
                        <span className="font-mono text-sm font-bold text-indigo-600 bg-indigo-50 px-2.5 py-0.5 rounded-md">
                          {questions.filter((q) => q.id.toLowerCase().startsWith("s")).length}
                        </span>
                      </div>

                      <div className="flex items-center justify-between pb-3 border-b border-dotted border-slate-100">
                        <span className="text-sm font-semibold text-slate-500">Main Questions (Group Q)</span>
                        <span className="font-mono text-sm font-bold text-emerald-600 bg-emerald-50 px-2.5 py-0.5 rounded-md">
                          {questions.filter((q) => q.id.toLowerCase().startsWith("q")).length}
                        </span>
                      </div>

                      <div className="flex items-center justify-between pb-3 border-b border-dotted border-slate-100">
                        <span className="text-sm font-semibold text-slate-500">Grid Elements (Matrix)</span>
                        <span className="font-mono text-sm font-bold text-amber-600 bg-amber-50 px-2.5 py-0.5 rounded-md">
                          {questions.filter((q) => q.type === "grid").length}
                        </span>
                      </div>

                      <div className="flex items-center justify-between">
                        <span className="text-sm font-semibold text-slate-500">Custom Manual Additions</span>
                        <span className="font-mono text-sm font-bold text-purple-600 bg-purple-50 px-2.5 py-0.5 rounded-md">
                          {questions.filter((q) => !q.originalText).length}
                        </span>
                      </div>
                    </div>

                    <div className="mt-6 pt-5 border-t border-slate-150">
                      <button
                        onClick={downloadDocx}
                        className="w-full flex items-center justify-center gap-2 py-3 px-4 bg-indigo-600 hover:bg-indigo-700 text-white font-bold rounded-xl shadow-lg transition-all cursor-pointer"
                      >
                        <Download className="w-5 h-5" />
                        <span>Export & Download Document</span>
                      </button>
                    </div>
                  </div>

                  {/* Predefined rules reminder */}
                  <div className="bg-slate-900 text-slate-300 rounded-2xl p-5 border border-slate-800 text-xs space-y-3 leading-relaxed">
                    <h4 className="text-white font-bold mb-1 uppercase tracking-wider text-[10px]">
                      Quick Guide on Special Codes
                    </h4>
                    <p>The parser dynamically evaluates categories and binds the correct numbering labels:</p>
                    <ul className="space-y-1 bg-slate-950/40 p-2.5 rounded-lg border border-slate-800 font-mono">
                      <li>• Other -&gt; <span className="text-amber-300 font-bold">[[97, other]]</span></li>
                      <li>• None of the above -&gt; <span className="text-cyan-330 font-bold">[[#98, exclusive]]</span></li>
                      <li>• Don't know -&gt; <span className="text-emerald-300 font-bold">[[#99, exclusive]]</span></li>
                    </ul>
                    <p>
                      If a question relies on <strong>"answer per column"</strong> rules, Forsta rules demand it is completely ignored to prevent matrix validation crashes.
                    </p>
                  </div>

                </div>
              </div>
            ) : (
              /* Raw Program Markup Preview Tab */
              <div className="bg-slate-900 rounded-2xl border border-slate-800 overflow-hidden shadow-xl flex flex-col">
                
                {/* Header Actions */}
                <div className="bg-slate-950 px-5 py-3.5 border-b border-slate-805 flex items-center justify-between">
                  <div className="flex items-center gap-2 text-xs font-mono text-slate-400 uppercase">
                    <FileText className="w-4 h-4 text-indigo-450" />
                    <span>Forsta Import Raw Text block preview</span>
                  </div>
                  
                  <div className="flex items-center gap-2">
                    <button
                      onClick={copyToClipboard}
                      className="flex items-center gap-1.5 px-3 py-1.5 text-xs font-semibold rounded-md transition-all text-slate-300 hover:text-white bg-slate-900 border border-slate-800 hover:bg-slate-800 cursor-pointer"
                    >
                      <Copy className="w-3.5 h-3.5" />
                      <span>{copiedIndex ? "Copied!" : "Copy Raw Text"}</span>
                    </button>
                    <button
                      onClick={downloadDocx}
                      className="flex items-center gap-1.5 px-3.5 py-1.5 text-xs font-bold rounded-md bg-emerald-600 hover:bg-emerald-700 text-white shadow-md cursor-pointer"
                    >
                      <Download className="w-3.5 h-3.5" />
                      <span>Download .docx</span>
                    </button>
                  </div>
                </div>

                {/* Preformatted Markup code layout */}
                <div className="p-6 overflow-x-auto min-h-[400px] max-h-[600px] overflow-y-auto bg-slate-950 text-slate-200">
                  <pre className="font-mono text-sm leading-relaxed antialiased">
                    {getCombinedTextMarkup() || "// No valid questions extracted"}
                  </pre>
                </div>
              </div>
            )}

          </div>
        )}

      </main>

      {/* Slide-out Tag Rules Cheatsheet Drawer */}
      <AnimatePresence>
        {showCheatsheet && (
          <>
            {/* Backdrop */}
            <motion.div
              initial={{ opacity: 0 }}
              animate={{ opacity: 0.4 }}
              exit={{ opacity: 0 }}
              onClick={() => setShowCheatsheet(false)}
              className="fixed inset-0 bg-black z-50 pointer-events-auto"
            />
            {/* Drawer */}
            <motion.div
              initial={{ x: "100%" }}
              animate={{ x: 0 }}
              exit={{ x: "100%" }}
              transition={{ type: "spring", damping: 25, stiffness: 220 }}
              className="fixed top-0 right-0 h-full w-full max-w-lg bg-white shadow-2xl z-55 flex flex-col border-l border-slate-200 overflow-y-auto scrollbar-none"
            >
              <div className="p-6 border-b border-slate-100 flex items-center justify-between bg-slate-50">
                <div className="flex items-center gap-2">
                  <HelpCircle className="w-5 h-5 text-indigo-600" />
                  <h3 className="font-bold text-slate-900 text-lg">Forsta Programming Tags Rules</h3>
                </div>
                <button
                  onClick={() => setShowCheatsheet(false)}
                  className="p-1 px-3 text-sm font-semibold rounded-md border border-slate-200 hover:bg-slate-100 cursor-pointer"
                >
                  Close
                </button>
              </div>

              {/* Scrollable rules list */}
              <div className="p-6 overflow-y-auto flex-1 space-y-6 text-sm text-slate-700 leading-normal">
                
                <div>
                  <h4 className="font-bold text-slate-950 mb-2.5 pb-1.5 border-b border-slate-100 flex items-center gap-1 font-mono text-xs uppercase tracking-wider text-indigo-650">
                    General Purpose Tags
                  </h4>
                  <div className="space-y-2 font-mono text-xs bg-slate-50 p-3 rounded-lg border border-slate-100">
                    <p><strong>[[begin note]] &amp; [[end note]]</strong>: Tags section as informational note</p>
                    <p><strong>[[begin ignore]] &amp; [[end ignore]]</strong>: Excludes custom paragraphs</p>
                    <p><strong>[[required]] / [[not required]]</strong>: Import with required mode (on by default)</p>
                    <p><strong>[[title]]</strong>: Applied directly underneath as title (e.g. Q1 [[title]])</p>
                    <p><strong>[[instruction]]</strong>: Tags a sentence/paragraph as question info</p>
                  </div>
                </div>

                <div>
                  <h4 className="font-bold text-slate-950 mb-2.5 pb-1.5 border-b border-slate-100 flex items-center gap-1 font-mono text-xs uppercase tracking-wider text-indigo-650">
                    Question Type Tags
                  </h4>
                  <div className="grid grid-cols-2 gap-2 text-xs font-mono bg-slate-50 p-3 rounded-lg border border-slate-100">
                    <div>• [[single]]</div>
                    <div>• [[multi]]</div>
                    <div>• [[grid]]</div>
                    <div>• [[ranking]]</div>
                    <div>• [[numeric]]</div>
                    <div>• [[numeric list]]</div>
                    <div>• [[text]]</div>
                    <div>• [[text list]]</div>
                    <div>• [[date]]</div>
                    <div>• [[info]]</div>
                  </div>
                </div>

                <div>
                  <h4 className="font-bold text-slate-950 mb-2.5 pb-1.5 border-b border-slate-100 flex items-center gap-1 font-mono text-xs uppercase tracking-wider text-indigo-650">
                    Special Codes Applied
                  </h4>
                  <table className="w-full text-left text-xs font-mono bg-slate-50 rounded-lg overflow-hidden border border-slate-100">
                    <thead>
                      <tr className="bg-slate-100 font-bold border-b border-slate-200">
                        <th className="p-2">Keyword Option</th>
                        <th className="p-2">Assigned Forsta Format</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-150">
                      <tr>
                        <td className="p-2">Other (please specify)</td>
                        <td className="p-2 text-indigo-600 font-bold">[[97, other]]</td>
                      </tr>
                      <tr>
                        <td className="p-2">None of the above / NA</td>
                        <td className="p-2 text-indigo-600 font-bold">[[#98, exclusive]]</td>
                      </tr>
                      <tr>
                        <td className="p-2">Don't know / I don't know</td>
                        <td className="p-2 text-indigo-600 font-bold">[[#99, exclusive]]</td>
                      </tr>
                    </tbody>
                  </table>
                </div>

                <div>
                  <h4 className="font-bold text-slate-950 mb-2.5 pb-1.5 border-b border-slate-100 flex items-center gap-1 font-mono text-xs uppercase tracking-wider text-indigo-650">
                    Special Conditions
                  </h4>
                  <div className="space-y-2 text-slate-600 leading-relaxed text-xs">
                    <p>• Randomization mentioned in instructions results in the option <strong>[[randomize]]</strong> tag applied beneath titles.</p>
                    <p>• Multi questions may have restrictions mapped automatically, e.g. <strong>[[min=1]] [[max=3]]</strong>.</p>
                    <p>• If a question relies on <strong>"answer per column"</strong> rules, Forsta rules demand it is completely ignored to prevent matrix validation crashes.</p>
                  </div>
                </div>

              </div>

              <div className="p-5 border-t border-slate-100 bg-slate-50">
                <button
                  onClick={() => setShowCheatsheet(false)}
                  className="w-full py-2.5 bg-indigo-600 hover:bg-indigo-700 text-white font-bold rounded-xl transition-all cursor-pointer shadow-lg"
                >
                  Got It
                </button>
              </div>
            </motion.div>
          </>
        )}
      </AnimatePresence>

      {/* Humble styling footer */}
      <footer className="bg-white border-t border-slate-200 py-5 text-center mt-auto">
        <p className="text-xs text-slate-400 font-mono">
          Forsta Survey Tagging Studio &bull; Powered by Gemini-3.5-Flash
        </p>
      </footer>

    </div>
  );
}
