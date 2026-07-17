import React, { useState, useRef } from "react";
import {
  Upload,
  FileText,
  Download,
  CheckCircle,
  AlertCircle,
  RefreshCw,
  Sparkles,
} from "lucide-react";
import { motion } from "motion/react";

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

export default function App() {
  const [file, setFile] = useState<File | null>(null);
  const [isDragging, setIsDragging] = useState(false);
  const [loading, setLoading] = useState(false);
  const [loadingStep, setLoadingStep] = useState(0);
  const [error, setError] = useState<string | null>(null);
  const [originalName, setOriginalName] = useState<string>("");
  const [questions, setQuestions] = useState<Question[]>([]);

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
        setQuestions([]); // Clear any previous results
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
        setQuestions([]); // Clear any previous results
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
        
        const lowerText = textResponse.toLowerCase();
        if (
          lowerText.includes("cookie check") || 
          lowerText.includes("action required to load your app") || 
          lowerText.includes("redirecttoreturnurl") ||
          lowerText.includes("grantstorageaccess") ||
          lowerText.includes("authinseparatewindowbutton") ||
          lowerText.includes("__secure-aistudio")
        ) {
          throw new Error("COOKIE_BLOCKED");
        }
        
        if (response.status === 413) {
          throw new Error("The uploaded file is too large. Please select a smaller Word Document.");
        }
        
        throw new Error(`The server returned an invalid response (Status ${response.status}). This can occur if your session expired, the backend is starting up, or the Gemini API key is missing. Please try reloading the page, opening the app in a new tab, or verify your API keys in the Settings.`);
      }

      const data = await response.json();

      if (!response.ok || !data.success) {
        throw new Error(data.error || "Failed to process the document.");
      }

      setQuestions(data.questions);
      setOriginalName(data.originalName);
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
        const contentType = response.headers.get("content-type");
        if (contentType && contentType.includes("text/html")) {
          const text = await response.text();
          if (text.includes("Cookie check") || text.includes("Action required to load your app")) {
            alert("Your session has expired or browser cookies are blocked. Please open this app in a new tab or authenticate using the link in the error panel.");
            return;
          }
        }
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
            <Sparkles className="w-6 h-6" />
          </div>
          <div>
            <span className="text-xs uppercase font-bold tracking-widest text-indigo-500">Forsta Assistant</span>
            <h1 className="font-sans font-bold text-lg text-slate-900 tracking-tight flex items-center gap-2">
              Survey Questionnaire Tagging App
            </h1>
          </div>
        </div>

        {(file || questions.length > 0) && (
          <button
            onClick={resetUpload}
            className="flex items-center gap-1.5 px-3.5 py-2 text-sm font-medium text-slate-600 hover:text-indigo-600 hover:bg-indigo-50 rounded-lg transition-all cursor-pointer"
          >
            <RefreshCw className="w-4 h-4" />
            <span>Start Over</span>
          </button>
        )}
      </header>

      {/* Primary Workspace */}
      <main className="flex-1 max-w-4xl w-full mx-auto p-4 md:p-8 flex flex-col justify-center gap-8">
        
        {/* Intro Hero Box (when no results and not loading) */}
        {!questions.length && !loading && (
          <div className="text-center max-w-2xl mx-auto space-y-4">
            <span className="inline-block px-3.5 py-1 bg-indigo-50 border border-indigo-100 text-indigo-600 rounded-full text-xs font-semibold">
              Powered by Gemini-3.5-Flash
            </span>
            <h2 className="text-3xl font-extrabold text-slate-900 tracking-tight leading-tight">
              Automatically tag and format survey questionnaires
            </h2>
            <p className="text-base text-slate-500 leading-relaxed">
              Upload your standard Word Document questionnaire (.docx). The assistant will instantly scan, organize, and structure Screener & Main questions with official Forsta tagging parameters.
            </p>
          </div>
        )}

        {/* Core Card Section */}
        <div className="max-w-2xl w-full mx-auto">
          {/* Case 1: Uploading & Ready Stage */}
          {!questions.length && !loading && (
            <div className="space-y-6">
              {/* Drag Drop Area */}
              <div
                onDragOver={handleDragOver}
                onDragLeave={handleDragLeave}
                onDrop={handleDrop}
                onClick={() => fileInputRef.current?.click()}
                className={`border-2 border-dashed rounded-2xl p-12 text-center cursor-pointer transition-all duration-300 min-h-[300px] flex flex-col items-center justify-center bg-white shadow-sm ${
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
                    <p className="text-xs uppercase tracking-wider font-bold text-indigo-500">Selected File</p>
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
                      Please upload only .docx files.
                    </p>
                    <span className="inline-flex py-1.5 px-4 bg-indigo-50 text-indigo-600 rounded-lg text-sm font-semibold hover:bg-indigo-100 transition-colors">
                      Browse Files
                    </span>
                  </div>
                )}
              </div>

              {/* Error Box */}
              {error && (
                <div className="p-5 bg-rose-50 border border-rose-150 rounded-xl text-rose-800 flex flex-col gap-3 shadow-sm">
                  <div className="flex items-start gap-3">
                    <AlertCircle className="w-5 h-5 text-rose-500 shrink-0 mt-0.5" />
                    <div className="flex-1">
                      {error === "COOKIE_BLOCKED" ? (
                        <div>
                          <h4 className="font-bold text-sm text-rose-900">Browser Security Cookie Blocked</h4>
                          <p className="text-xs text-rose-700 leading-relaxed mt-1">
                            Your browser's privacy settings are blocking cookies in the preview panel (common in Safari, iOS, or with strict privacy rules). Please open the app in a new tab or authenticate to restore the session.
                          </p>
                          <div className="mt-3.5 flex flex-wrap gap-2">
                            <a
                              href={window.location.href}
                              target="_blank"
                              rel="noreferrer"
                              className="inline-flex items-center justify-center gap-1.5 px-4 py-2 bg-indigo-650 hover:bg-indigo-700 text-white text-xs font-bold rounded-lg shadow-sm transition-colors text-center cursor-pointer decoration-none"
                            >
                              Open in New Tab (Recommended)
                            </a>
                            <button
                              onClick={() => {
                                window.open(window.location.href, "_blank");
                              }}
                              className="inline-flex items-center justify-center gap-1.5 px-4 py-2 bg-white hover:bg-slate-100 text-slate-800 border border-slate-200 text-xs font-semibold rounded-lg transition-colors cursor-pointer text-center"
                            >
                              Authenticate Cookie
                            </button>
                          </div>
                        </div>
                      ) : (
                        <div>
                          <h4 className="font-bold text-sm text-rose-900">Processing Error</h4>
                          <p className="text-xs text-rose-700 leading-relaxed mt-0.5">{error}</p>
                        </div>
                      )}
                    </div>
                  </div>
                </div>
              )}

              {/* Action Button */}
              {file && (
                <button
                  onClick={parseDocument}
                  className="flex items-center gap-2 px-6 py-4 bg-indigo-600 hover:bg-indigo-700 text-white font-bold rounded-xl shadow-lg hover:shadow-indigo-500/20 hover:-translate-y-0.5 transition-all w-full justify-center cursor-pointer"
                >
                  <Sparkles className="w-5 h-5" />
                  <span>Run AI Tagging Engine</span>
                </button>
              )}
            </div>
          )}

          {/* Case 2: Loading & Progress Stage */}
          {loading && (
            <div className="bg-white border border-slate-200 rounded-3xl p-16 shadow-lg flex flex-col items-center justify-center text-center relative overflow-hidden">
              <div className="absolute top-0 right-0 w-32 h-32 bg-indigo-50/50 rounded-full blur-2xl -z-10" />
              <div className="relative mb-6">
                <div className="w-16 h-16 border-4 border-slate-100 border-t-indigo-600 rounded-full animate-spin" />
                <div className="absolute top-1/2 left-1/2 -translate-x-1/2 -translate-y-1/2 text-indigo-650">
                  <Sparkles className="w-6 h-6 animate-pulse" />
                </div>
              </div>
              
              <h3 className="font-bold text-slate-900 text-xl tracking-tight mb-2">Analyzing Questionnaire</h3>
              <p className="text-sm text-slate-500 max-w-xs mb-6 text-center leading-normal">
                Applying official Forsta parsing and tag structure to your document.
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

          {/* Case 3: Success & Download Stage */}
          {questions.length > 0 && !loading && (
            <div className="bg-white border border-slate-200 rounded-3xl p-12 shadow-lg text-center space-y-6">
              <div className="inline-flex p-4 bg-emerald-50 text-emerald-600 rounded-full border border-emerald-100">
                <CheckCircle className="w-12 h-12" />
              </div>

              <div className="space-y-2">
                <h3 className="font-sans font-bold text-slate-900 text-2xl tracking-tight">
                  Questionnaire Tagged Successfully!
                </h3>
                <p className="text-sm text-slate-500 max-w-md mx-auto">
                  Our AI successfully parsed and formatted your document, identifying <span className="font-bold text-slate-800">{questions.length}</span> survey questions with full Forsta tag configurations.
                </p>
              </div>

              <div className="p-4 bg-slate-50 rounded-2xl border border-slate-100 text-left max-w-sm mx-auto flex items-center gap-3">
                <FileText className="w-8 h-8 text-indigo-600 shrink-0" />
                <div className="min-w-0 flex-1">
                  <p className="text-xs font-semibold text-slate-400 uppercase tracking-wider">Processed File</p>
                  <p className="text-sm font-bold text-slate-800 truncate">{originalName || file?.name || "tagged_questionnaire.docx"}</p>
                </div>
              </div>

              <div className="pt-4 flex flex-col sm:flex-row gap-3 justify-center">
                <button
                  onClick={downloadDocx}
                  className="flex items-center justify-center gap-2 px-6 py-4 bg-emerald-600 hover:bg-emerald-700 text-white font-bold rounded-xl shadow-lg hover:shadow-emerald-500/10 hover:-translate-y-0.5 transition-all cursor-pointer flex-1"
                >
                  <Download className="w-5 h-5" />
                  <span>Download Tagged Document</span>
                </button>
                <button
                  onClick={resetUpload}
                  className="flex items-center justify-center gap-2 px-6 py-4 bg-slate-100 hover:bg-slate-200 text-slate-700 font-semibold rounded-xl transition-all cursor-pointer"
                >
                  <span>Upload Another</span>
                </button>
              </div>
            </div>
          )}
        </div>

      </main>

      {/* Humble footer */}
      <footer className="bg-white border-t border-slate-200 py-5 text-center mt-auto">
        <p className="text-xs text-slate-400 font-mono">
          Survey Questionnaire Tagging App &bull; Powered by Gemini-3.5-Flash
        </p>
      </footer>
    </div>
  );
}
