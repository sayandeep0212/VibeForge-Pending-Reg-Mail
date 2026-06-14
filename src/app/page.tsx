"use client";

import React, { useState } from "react";
import * as xlsx from "xlsx";
import { UploadCloud, FileSpreadsheet, Play, Copy, CheckCircle2, AlertCircle, Send, Loader2, X } from "lucide-react";
import { emailTemplateHtml } from "@/lib/emailTemplate";

type Participant = {
  UID: string;
  Name: string;
  Email: string;
  Phone: string;
  College: string;
  Year: string;
  Program: string;
  [key: string]: any;
};

type Registration = {
  "Leader UID"?: string;
  Members?: string;
  [key: string]: any;
};

export default function Home() {
  const [participantsFile, setParticipantsFile] = useState<File | null>(null);
  const [registrationsFile, setRegistrationsFile] = useState<File | null>(null);
  const [results, setResults] = useState<Participant[] | null>(null);
  const [isAnalyzing, setIsAnalyzing] = useState(false);
  const [copied, setCopied] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [isSendingMail, setIsSendingMail] = useState(false);
  const [mailSent, setMailSent] = useState(false);
  const [showPreview, setShowPreview] = useState(false);

  const handleFileUpload = (e: React.ChangeEvent<HTMLInputElement>, type: "participants" | "registrations") => {
    const file = e.target.files?.[0];
    if (file) {
      if (type === "participants") setParticipantsFile(file);
      else setRegistrationsFile(file);
      setError(null);
    }
  };

  const readExcelFile = (file: File): Promise<any[]> => {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = (e) => {
        try {
          const data = new Uint8Array(e.target?.result as ArrayBuffer);
          const workbook = xlsx.read(data, { type: "array" });
          const sheetName = workbook.SheetNames[0];
          const sheet = workbook.Sheets[sheetName];
          const json = xlsx.utils.sheet_to_json(sheet);
          resolve(json);
        } catch (error) {
          reject(error);
        }
      };
      reader.onerror = (error) => reject(error);
      reader.readAsArrayBuffer(file);
    });
  };

  const handleAnalyze = async () => {
    if (!participantsFile || !registrationsFile) {
      setError("Please upload both the Participants and Registrations files.");
      return;
    }

    setIsAnalyzing(true);
    setError(null);

    try {
      const participantsData = (await readExcelFile(participantsFile)) as Participant[];
      const registrationsData = (await readExcelFile(registrationsFile)) as Registration[];

      const registeredUIDs = new Set<string>();

      registrationsData.forEach((reg) => {
        if (reg["Leader UID"]) {
          registeredUIDs.add(String(reg["Leader UID"]).trim());
        }

        if (reg["Members"]) {
          const membersStr = String(reg["Members"]);
          const uidRegex = /\((GL-[A-Z0-9_-]+)\)/gi;
          let match;
          while ((match = uidRegex.exec(membersStr)) !== null) {
            registeredUIDs.add(match[1].trim());
          }
        }
      });

      const missing = participantsData.filter((p) => p.UID && !registeredUIDs.has(String(p.UID).trim()));
      setResults(missing);
    } catch (err) {
      console.error(err);
      setError("An error occurred while analyzing the files. Please make sure they are valid Excel files.");
    } finally {
      setIsAnalyzing(false);
    }
  };

  const handleCopyEmails = () => {
    if (!results) return;
    const emails = results.map((r) => r.Email).filter(Boolean).join(", ");
    navigator.clipboard.writeText(emails);
    setCopied(true);
    setTimeout(() => setCopied(false), 2000);
  };

  const handleSendEmails = async () => {
    if (!results || results.length === 0) return;

    setIsSendingMail(true);
    setError(null);

    const emails = results.map((r) => r.Email).filter(Boolean);

    try {
      const response = await fetch('/api/send-email', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({ recipients: emails }),
      });

      const data = await response.json();

      if (!response.ok) {
        throw new Error(data.error || 'Failed to send emails');
      }

      setMailSent(true);
      setShowPreview(false);
      setTimeout(() => setMailSent(false), 5000);
    } catch (err: any) {
      console.error(err);
      setError(err.message || 'An error occurred while sending emails.');
    } finally {
      setIsSendingMail(false);
    }
  };

  return (
    <main className="min-h-screen bg-slate-950 text-slate-100 p-6 md:p-12 selection:bg-indigo-500/30">
      <div className="max-w-6xl mx-auto space-y-12">
        {/* Header */}
        <header className="text-center space-y-4">
          <div className="inline-flex items-center justify-center p-3 rounded-2xl bg-indigo-500/10 text-indigo-400 mb-2">
            <UploadCloud size={32} />
          </div>
          <h1 className="text-4xl md:text-5xl font-extrabold tracking-tight bg-gradient-to-br from-white to-slate-400 bg-clip-text text-transparent">
            Data Refiner Pro
          </h1>
          <p className="text-slate-400 max-w-2xl mx-auto text-lg">
            Upload your weekly Participation and Registration reports. We&apos;ll instantly identify who set up a profile but hasn&apos;t registered for a team yet.
          </p>
        </header>

        {/* Upload Section */}
        <div className="grid md:grid-cols-2 gap-6">
          <div className={`relative group rounded-3xl border-2 border-dashed transition-all duration-300 p-8 flex flex-col items-center justify-center text-center ${participantsFile ? 'border-emerald-500/50 bg-emerald-500/5' : 'border-slate-800 hover:border-indigo-500/50 hover:bg-indigo-500/5 bg-slate-900/50'}`}>
            <input 
              type="file" 
              accept=".xlsx,.xls" 
              onChange={(e) => handleFileUpload(e, "participants")}
              className="absolute inset-0 w-full h-full opacity-0 cursor-pointer"
            />
            <FileSpreadsheet className={`w-12 h-12 mb-4 transition-colors ${participantsFile ? 'text-emerald-400' : 'text-slate-500 group-hover:text-indigo-400'}`} />
            <h3 className="text-xl font-semibold mb-2">1. Participants Report</h3>
            {participantsFile ? (
              <p className="text-emerald-400 font-medium">{participantsFile.name}</p>
            ) : (
              <p className="text-slate-400">Drag & drop or click to upload</p>
            )}
          </div>

          <div className={`relative group rounded-3xl border-2 border-dashed transition-all duration-300 p-8 flex flex-col items-center justify-center text-center ${registrationsFile ? 'border-emerald-500/50 bg-emerald-500/5' : 'border-slate-800 hover:border-indigo-500/50 hover:bg-indigo-500/5 bg-slate-900/50'}`}>
            <input 
              type="file" 
              accept=".xlsx,.xls" 
              onChange={(e) => handleFileUpload(e, "registrations")}
              className="absolute inset-0 w-full h-full opacity-0 cursor-pointer"
            />
            <FileSpreadsheet className={`w-12 h-12 mb-4 transition-colors ${registrationsFile ? 'text-emerald-400' : 'text-slate-500 group-hover:text-indigo-400'}`} />
            <h3 className="text-xl font-semibold mb-2">2. Registrations Report</h3>
            {registrationsFile ? (
              <p className="text-emerald-400 font-medium">{registrationsFile.name}</p>
            ) : (
              <p className="text-slate-400">Drag & drop or click to upload</p>
            )}
          </div>
        </div>

        {/* Action Button */}
        <div className="flex flex-col items-center space-y-4">
          <button
            onClick={handleAnalyze}
            disabled={!participantsFile || !registrationsFile || isAnalyzing}
            className="group relative inline-flex items-center justify-center px-8 py-4 font-bold text-white transition-all duration-300 bg-indigo-600 rounded-full hover:bg-indigo-500 focus:outline-none focus:ring-4 focus:ring-indigo-500/50 disabled:opacity-50 disabled:cursor-not-allowed overflow-hidden"
          >
            <div className="absolute inset-0 w-full h-full -mt-1 rounded-lg opacity-30 bg-gradient-to-b from-transparent via-transparent to-black pointer-events-none"></div>
            <span className="relative flex items-center gap-2">
              {isAnalyzing ? (
                <div className="w-5 h-5 border-2 border-white/30 border-t-white rounded-full animate-spin" />
              ) : (
                <Play className="w-5 h-5 group-hover:scale-110 transition-transform" />
              )}
              {isAnalyzing ? "Analyzing Data..." : "Run Analysis"}
            </span>
          </button>
          
          {error && (
            <div className="flex items-center gap-2 text-rose-400 bg-rose-400/10 px-4 py-3 rounded-xl">
              <AlertCircle className="w-5 h-5" />
              <p>{error}</p>
            </div>
          )}
        </div>

        {/* Results Section */}
        {results && (
          <div className="animate-in fade-in slide-in-from-bottom-8 duration-700 space-y-6">
            <div className="flex flex-col md:flex-row items-center justify-between gap-4 p-6 rounded-3xl bg-slate-900/50 border border-slate-800 backdrop-blur-sm">
              <div>
                <h2 className="text-2xl font-bold">Analysis Complete</h2>
                <p className="text-slate-400">
                  Found <span className="text-indigo-400 font-bold text-xl">{results.length}</span> participants who haven&apos;t registered.
                </p>
              </div>
              <div className="flex gap-3">
                <button
                  onClick={() => setShowPreview(true)}
                  disabled={isSendingMail || results.length === 0}
                  className={`flex items-center gap-2 px-6 py-3 rounded-xl font-semibold transition-all disabled:opacity-50 disabled:cursor-not-allowed ${
                    mailSent
                      ? 'bg-emerald-500/20 text-emerald-400 border border-emerald-500/50'
                      : 'bg-indigo-600 text-white hover:bg-indigo-500 border border-indigo-500/50'
                  }`}
                >
                  {isSendingMail ? <Loader2 className="w-5 h-5 animate-spin" /> : (mailSent ? <CheckCircle2 className="w-5 h-5" /> : <Send className="w-5 h-5" />)}
                  {isSendingMail ? "Sending..." : (mailSent ? "Mail Sent!" : "Send Mail")}
                </button>
                <button
                  onClick={handleCopyEmails}
                  className={`flex items-center gap-2 px-6 py-3 rounded-xl font-semibold transition-all ${
                    copied 
                      ? 'bg-emerald-500/20 text-emerald-400 border border-emerald-500/50' 
                      : 'bg-slate-800 text-slate-200 hover:bg-slate-700 border border-slate-700'
                  }`}
                >
                  {copied ? <CheckCircle2 className="w-5 h-5" /> : <Copy className="w-5 h-5" />}
                  {copied ? "Copied!" : "Copy All Emails"}
                </button>
              </div>
            </div>

            <div className="rounded-3xl border border-slate-800 bg-slate-900/50 backdrop-blur-sm overflow-hidden">
              <div className="overflow-x-auto">
                <table className="w-full text-left border-collapse">
                  <thead>
                    <tr className="bg-slate-800/50 text-slate-400 text-sm uppercase tracking-wider">
                      <th className="p-4 font-medium">Name</th>
                      <th className="p-4 font-medium">Email</th>
                      <th className="p-4 font-medium">College</th>
                      <th className="p-4 font-medium">Phone</th>
                      <th className="p-4 font-medium">Year</th>
                      <th className="p-4 font-medium">Program</th>
                      <th className="p-4 font-medium">UID</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-800/50">
                    {results.map((r, i) => (
                      <tr key={i} className="hover:bg-slate-800/30 transition-colors">
                        <td className="p-4 font-medium text-slate-200">{r.Name || "-"}</td>
                        <td className="p-4 text-slate-400">{r.Email || "-"}</td>
                        <td className="p-4 text-slate-400">{r.College || "-"}</td>
                        <td className="p-4 text-slate-400">{r.Phone || "-"}</td>
                        <td className="p-4 text-slate-400 whitespace-nowrap">{r.Year || "-"}</td>
                        <td className="p-4 text-slate-400">{r.Program || "-"}</td>
                        <td className="p-4 text-slate-500 font-mono text-xs">{r.UID || "-"}</td>
                      </tr>
                    ))}
                    {results.length === 0 && (
                      <tr>
                        <td colSpan={7} className="p-8 text-center text-slate-500">
                          Great job! All participants have registered.
                        </td>
                      </tr>
                    )}
                  </tbody>
                </table>
              </div>
            </div>
          </div>
        )}
      </div>

      {/* Preview Modal */}
      {showPreview && (
        <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-black/60 backdrop-blur-sm animate-in fade-in duration-200">
          <div className="bg-slate-900 border border-slate-800 rounded-3xl w-full max-w-4xl max-h-[90vh] flex flex-col overflow-hidden shadow-2xl">
            <div className="flex items-center justify-between p-6 border-b border-slate-800">
              <h2 className="text-xl font-bold text-white">Email Preview</h2>
              <button 
                onClick={() => setShowPreview(false)}
                className="p-2 rounded-full hover:bg-slate-800 text-slate-400 hover:text-white transition-colors"
              >
                <X className="w-5 h-5" />
              </button>
            </div>
            
            <div className="flex-1 overflow-hidden bg-white p-0">
              <iframe 
                srcDoc={emailTemplateHtml}
                title="Email Preview"
                className="w-full h-full min-h-[500px] border-none"
                sandbox="allow-same-origin"
              />
            </div>

            <div className="p-6 border-t border-slate-800 bg-slate-900/50 flex justify-end gap-4">
              <button
                onClick={() => setShowPreview(false)}
                disabled={isSendingMail}
                className="px-6 py-3 rounded-xl font-semibold text-slate-300 hover:text-white hover:bg-slate-800 transition-colors disabled:opacity-50"
              >
                Cancel
              </button>
              <button
                onClick={handleSendEmails}
                disabled={isSendingMail}
                className="flex items-center gap-2 px-6 py-3 rounded-xl font-semibold bg-indigo-600 text-white hover:bg-indigo-500 transition-all disabled:opacity-50 disabled:cursor-not-allowed"
              >
                {isSendingMail ? <Loader2 className="w-5 h-5 animate-spin" /> : <Send className="w-5 h-5" />}
                {isSendingMail ? "Sending..." : "Confirm & Send Mail"}
              </button>
            </div>
          </div>
        </div>
      )}
    </main>
  );
}
