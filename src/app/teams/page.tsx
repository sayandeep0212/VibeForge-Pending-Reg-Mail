"use client";

import React, { useState, useMemo } from "react";
import * as xlsx from "xlsx";
import {
  UploadCloud,
  FileSpreadsheet,
  Play,
  Download,
  Search,
  Users,
  Crown,
  User,
  ChevronDown,
  ChevronUp,
  AlertCircle,
  Hash,
  FileText,
  Table,
} from "lucide-react";

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
  "Team Name"?: string;
  "Leader UID"?: string;
  Members?: string;
  "Member Emails"?: string;
  "Member Count"?: number;
  Type?: string;
  University?: string;
  [key: string]: any;
};

type TeamMember = {
  name: string;
  uid: string;
  email: string;
  phone: string;
  college: string;
  year: string;
  program: string;
  isLeader: boolean;
};

type TeamData = {
  teamName: string;
  type: string;
  university: string;
  memberCount: number;
  leader: TeamMember;
  members: TeamMember[];
};

export default function TeamsPage() {
  const [participantsFile, setParticipantsFile] = useState<File | null>(null);
  const [registrationsFile, setRegistrationsFile] = useState<File | null>(null);
  const [teams, setTeams] = useState<TeamData[] | null>(null);
  const [isAnalyzing, setIsAnalyzing] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [searchQuery, setSearchQuery] = useState("");
  const [expandedTeams, setExpandedTeams] = useState<Set<number>>(new Set());
  const [filterType, setFilterType] = useState<"all" | "team" | "individual">("all");
  const [showExportMenu, setShowExportMenu] = useState(false);

  const filteredTeams = useMemo(() => {
    if (!teams) return null;
    let filtered = teams;

    if (filterType !== "all") {
      filtered = filtered.filter((t) => t.type === filterType);
    }

    if (searchQuery.trim()) {
      const q = searchQuery.toLowerCase();
      filtered = filtered.filter(
        (t) =>
          t.teamName.toLowerCase().includes(q) ||
          t.leader.name.toLowerCase().includes(q) ||
          t.leader.uid.toLowerCase().includes(q) ||
          t.members.some(
            (m) =>
              m.name.toLowerCase().includes(q) ||
              m.uid.toLowerCase().includes(q)
          )
      );
    }

    return filtered;
  }, [teams, searchQuery, filterType]);

  const handleFileUpload = (
    e: React.ChangeEvent<HTMLInputElement>,
    type: "participants" | "registrations"
  ) => {
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
          const sheet = workbook.Sheets[workbook.SheetNames[0]];
          resolve(xlsx.utils.sheet_to_json(sheet));
        } catch (err) {
          reject(err);
        }
      };
      reader.onerror = reject;
      reader.readAsArrayBuffer(file);
    });
  };

  const handleAnalyze = async () => {
    if (!participantsFile || !registrationsFile) {
      setError("Please upload both files.");
      return;
    }
    setIsAnalyzing(true);
    setError(null);

    try {
      const participants = (await readExcelFile(participantsFile)) as Participant[];
      const registrations = (await readExcelFile(registrationsFile)) as Registration[];

      // Build a lookup map: UID -> Participant
      const participantMap = new Map<string, Participant>();
      participants.forEach((p) => {
        if (p.UID) participantMap.set(String(p.UID).trim(), p);
      });

      const teamResults: TeamData[] = registrations.map((reg) => {
        const leaderUID = String(reg["Leader UID"] || "").trim();
        const leaderData = participantMap.get(leaderUID);

        const leader: TeamMember = {
          name: leaderData?.Name || "Unknown",
          uid: leaderUID,
          email: leaderData?.Email || "",
          phone: leaderData?.Phone || "",
          college: leaderData?.College || "",
          year: leaderData?.Year || "",
          program: leaderData?.Program || "",
          isLeader: true,
        };

        // Parse members from "Name (UID), Name (UID)" format
        const membersStr = String(reg["Members"] || "");
        const memberRegex = /([^,]+?)\s*\(([^)]+)\)/g;
        const allMembers: TeamMember[] = [];
        let match;
        while ((match = memberRegex.exec(membersStr)) !== null) {
          const memberName = match[1].trim();
          const memberUID = match[2].trim();
          const memberData = participantMap.get(memberUID);

          allMembers.push({
            name: memberData?.Name || memberName,
            uid: memberUID,
            email: memberData?.Email || "",
            phone: memberData?.Phone || "",
            college: memberData?.College || "",
            year: memberData?.Year || "",
            program: memberData?.Program || "",
            isLeader: memberUID === leaderUID,
          });
        }

        // Non-leader members
        const nonLeaderMembers = allMembers.filter((m) => m.uid !== leaderUID);

        return {
          teamName: String(reg["Team Name"] || "Unnamed Team"),
          type: String(reg["Type"] || "team"),
          university: String(reg["University"] || ""),
          memberCount: reg["Member Count"] || allMembers.length,
          leader,
          members: nonLeaderMembers,
        };
      });

      setTeams(teamResults);
      // Expand all teams by default
      setExpandedTeams(new Set(teamResults.map((_, i) => i)));
    } catch (err) {
      console.error(err);
      setError("Error analyzing files. Please check the Excel format.");
    } finally {
      setIsAnalyzing(false);
    }
  };

  const buildExportRows = () => {
    if (!filteredTeams || filteredTeams.length === 0) return [];
    return filteredTeams.map((team) => ({
      "Team Name": team.teamName,
      "Leader Name": team.leader.name,
      "Leader Phone": team.leader.phone,
      "Leader Email": team.leader.email,
    }));
  };

  const handleExportExcel = () => {
    const rows = buildExportRows();
    if (rows.length === 0) return;
    const ws = xlsx.utils.json_to_sheet(rows);
    const wb = xlsx.utils.book_new();
    xlsx.utils.book_append_sheet(wb, ws, "Teams");
    xlsx.writeFile(wb, "team_analysis.xlsx");
    setShowExportMenu(false);
  };

  const handleExportCSV = () => {
    const rows = buildExportRows();
    if (rows.length === 0) return;
    const headers = Object.keys(rows[0]);
    const csvLines = [
      headers.map((h) => `"${h}"`).join(","),
      ...rows.map((row) =>
        headers.map((h) => {
          const val = String((row as Record<string, string>)[h] ?? "").replace(/"/g, '""');
          return `"${val}"`;
        }).join(",")
      ),
    ];
    const blob = new Blob([csvLines.join("\n")], { type: "text/csv;charset=utf-8;" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = "team_analysis.csv";
    a.click();
    URL.revokeObjectURL(url);
    setShowExportMenu(false);
  };

  const toggleTeam = (index: number) => {
    setExpandedTeams((prev) => {
      const next = new Set(prev);
      if (next.has(index)) next.delete(index);
      else next.add(index);
      return next;
    });
  };

  const toggleAll = () => {
    if (!filteredTeams) return;
    if (expandedTeams.size === filteredTeams.length) {
      setExpandedTeams(new Set());
    } else {
      setExpandedTeams(new Set(filteredTeams.map((_, i) => i)));
    }
  };

  return (
    <main className="min-h-screen bg-slate-950 text-slate-100 p-4 md:p-10 selection:bg-violet-500/30">
      <div className="max-w-7xl mx-auto space-y-10">
        {/* Header */}
        <header className="text-center space-y-3 pt-4">
          <div className="inline-flex items-center justify-center p-3.5 rounded-2xl bg-violet-500/10 text-violet-400 mb-1">
            <Users size={30} />
          </div>
          <h1 className="text-4xl md:text-5xl font-extrabold tracking-tight bg-gradient-to-br from-white via-violet-200 to-violet-400 bg-clip-text text-transparent">
            Team Analyzer
          </h1>
          <p className="text-slate-400 max-w-2xl mx-auto text-lg">
            Upload Participants &amp; Registrations reports to view all teams with their leaders and members.
          </p>
        </header>

        {/* Upload Cards */}
        <div className="grid md:grid-cols-2 gap-5">
          {(["participants", "registrations"] as const).map((type) => {
            const file = type === "participants" ? participantsFile : registrationsFile;
            const label = type === "participants" ? "1. Participants Report" : "2. Registrations Report";
            return (
              <div
                key={type}
                className={`relative group rounded-2xl border-2 border-dashed transition-all duration-300 p-7 flex flex-col items-center justify-center text-center ${
                  file
                    ? "border-emerald-500/50 bg-emerald-500/5"
                    : "border-slate-800 hover:border-violet-500/50 hover:bg-violet-500/5 bg-slate-900/50"
                }`}
              >
                <input
                  type="file"
                  accept=".xlsx,.xls"
                  onChange={(e) => handleFileUpload(e, type)}
                  className="absolute inset-0 w-full h-full opacity-0 cursor-pointer"
                />
                <FileSpreadsheet
                  className={`w-10 h-10 mb-3 transition-colors ${
                    file ? "text-emerald-400" : "text-slate-500 group-hover:text-violet-400"
                  }`}
                />
                <h3 className="text-lg font-semibold mb-1">{label}</h3>
                {file ? (
                  <p className="text-emerald-400 font-medium text-sm">{file.name}</p>
                ) : (
                  <p className="text-slate-500 text-sm">Click or drag to upload</p>
                )}
              </div>
            );
          })}
        </div>

        {/* Analyze Button */}
        <div className="flex flex-col items-center space-y-3">
          <button
            onClick={handleAnalyze}
            disabled={!participantsFile || !registrationsFile || isAnalyzing}
            className="group inline-flex items-center gap-2 px-8 py-3.5 font-bold text-white bg-violet-600 rounded-full hover:bg-violet-500 transition-all focus:outline-none focus:ring-4 focus:ring-violet-500/40 disabled:opacity-50 disabled:cursor-not-allowed"
          >
            {isAnalyzing ? (
              <div className="w-5 h-5 border-2 border-white/30 border-t-white rounded-full animate-spin" />
            ) : (
              <Play className="w-5 h-5 group-hover:scale-110 transition-transform" />
            )}
            {isAnalyzing ? "Analyzing..." : "Run Analysis"}
          </button>
          {error && (
            <div className="flex items-center gap-2 text-rose-400 bg-rose-400/10 px-4 py-2.5 rounded-xl text-sm">
              <AlertCircle className="w-4 h-4 shrink-0" />
              <p>{error}</p>
            </div>
          )}
        </div>

        {/* Results */}
        {teams && filteredTeams && (
          <div className="space-y-5">
            {/* Toolbar */}
            <div className="flex flex-col lg:flex-row items-stretch lg:items-center justify-between gap-4 p-5 rounded-2xl bg-slate-900/60 border border-slate-800 backdrop-blur-sm overflow-visible">
              <div className="flex items-center gap-3 flex-wrap">
                <h2 className="text-xl font-bold whitespace-nowrap">
                  <span className="text-violet-400">{filteredTeams.length}</span>{" "}
                  <span className="text-slate-300">
                    {filteredTeams.length === 1 ? "Team" : "Teams"}
                  </span>
                </h2>
                <div className="flex bg-slate-800/60 p-0.5 rounded-lg">
                  {(["all", "team", "individual"] as const).map((t) => (
                    <button
                      key={t}
                      onClick={() => setFilterType(t)}
                      className={`px-3 py-1.5 rounded-md text-sm font-medium transition-all capitalize ${
                        filterType === t
                          ? "bg-violet-600 text-white shadow"
                          : "text-slate-400 hover:text-white"
                      }`}
                    >
                      {t}
                    </button>
                  ))}
                </div>
              </div>
              <div className="flex items-center gap-3 flex-1 lg:flex-initial lg:max-w-md overflow-visible">
                <div className="relative flex-1">
                  <Search className="absolute left-3 top-1/2 -translate-y-1/2 w-4 h-4 text-slate-500" />
                  <input
                    type="text"
                    placeholder="Search team, member, or UID..."
                    value={searchQuery}
                    onChange={(e) => setSearchQuery(e.target.value)}
                    className="w-full pl-9 pr-3 py-2.5 rounded-xl bg-slate-800/60 border border-slate-700 text-sm text-slate-200 placeholder:text-slate-500 focus:outline-none focus:ring-2 focus:ring-violet-500/40 focus:border-violet-500/50 transition-all"
                  />
                </div>
                <button
                  onClick={toggleAll}
                  className="px-3 py-2.5 rounded-xl bg-slate-800/60 border border-slate-700 text-slate-400 hover:text-white text-sm font-medium transition-colors whitespace-nowrap"
                >
                  {expandedTeams.size === filteredTeams.length ? "Collapse" : "Expand"} All
                </button>
                <div className="relative">
                  <button
                    onClick={() => setShowExportMenu((p) => !p)}
                    className="flex items-center gap-1.5 px-4 py-2.5 rounded-xl bg-emerald-600/20 text-emerald-400 hover:bg-emerald-600/30 border border-emerald-500/30 text-sm font-semibold transition-all whitespace-nowrap"
                  >
                    <Download className="w-4 h-4" />
                    Export
                    <ChevronDown className="w-3.5 h-3.5" />
                  </button>
                  {showExportMenu && (
                    <>
                      <div className="fixed inset-0 z-40" onClick={() => setShowExportMenu(false)} />
                      <div className="absolute right-0 bottom-full mb-1.5 w-44 rounded-xl bg-slate-800 border border-slate-700 shadow-2xl overflow-hidden z-50">
                        <button
                          onClick={handleExportExcel}
                          className="w-full flex items-center gap-2.5 px-4 py-3 text-sm text-slate-200 hover:bg-slate-700/60 transition-colors"
                        >
                          <Table className="w-4 h-4 text-emerald-400" />
                          Excel (.xlsx)
                        </button>
                        <button
                          onClick={handleExportCSV}
                          className="w-full flex items-center gap-2.5 px-4 py-3 text-sm text-slate-200 hover:bg-slate-700/60 transition-colors border-t border-slate-700/60"
                        >
                          <FileText className="w-4 h-4 text-blue-400" />
                          CSV (.csv)
                        </button>
                      </div>
                    </>
                  )}
                </div>
              </div>
            </div>

            {/* Team Cards */}
            <div className="space-y-3">
              {filteredTeams.map((team, idx) => {
                const isExpanded = expandedTeams.has(idx);
                return (
                  <div
                    key={idx}
                    className="rounded-2xl border border-slate-800 bg-slate-900/50 backdrop-blur-sm overflow-hidden transition-all hover:border-slate-700"
                  >
                    {/* Team Header */}
                    <button
                      onClick={() => toggleTeam(idx)}
                      className="w-full flex items-center justify-between px-5 py-4 text-left hover:bg-slate-800/30 transition-colors"
                    >
                      <div className="flex items-center gap-3 min-w-0">
                        <div className="w-10 h-10 rounded-xl bg-violet-500/15 text-violet-400 flex items-center justify-center shrink-0">
                          <Users className="w-5 h-5" />
                        </div>
                        <div className="min-w-0">
                          <h3 className="font-bold text-base text-white truncate">
                            {team.teamName}
                          </h3>
                          <div className="flex items-center gap-2 text-xs text-slate-500 mt-0.5 flex-wrap">
                            <span className={`px-2 py-0.5 rounded-full font-medium ${
                              team.type === "team"
                                ? "bg-blue-500/10 text-blue-400"
                                : "bg-amber-500/10 text-amber-400"
                            }`}>
                              {team.type}
                            </span>
                            <span className="flex items-center gap-1">
                              <Hash className="w-3 h-3" />
                              {team.memberCount} member{team.memberCount !== 1 ? "s" : ""}
                            </span>
                            {team.university && (
                              <span className="text-slate-600">• {team.university}</span>
                            )}
                          </div>
                        </div>
                      </div>
                      {isExpanded ? (
                        <ChevronUp className="w-5 h-5 text-slate-500 shrink-0" />
                      ) : (
                        <ChevronDown className="w-5 h-5 text-slate-500 shrink-0" />
                      )}
                    </button>

                    {/* Expanded Content */}
                    {isExpanded && (
                      <div className="px-5 pb-5 pt-1 space-y-3 border-t border-slate-800/60">
                        {/* Leader Card */}
                        <div className="rounded-xl bg-gradient-to-r from-amber-500/5 via-amber-500/10 to-amber-500/5 border border-amber-500/20 p-4">
                          <div className="flex items-center gap-2 mb-2.5">
                            <Crown className="w-4 h-4 text-amber-400" />
                            <span className="text-xs font-bold uppercase tracking-wider text-amber-400">
                              Team Leader
                            </span>
                          </div>
                          <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-x-6 gap-y-1.5 text-sm">
                            <div>
                              <span className="text-slate-500 text-xs">Name</span>
                              <p className="text-white font-medium">{team.leader.name}</p>
                            </div>
                            <div>
                              <span className="text-slate-500 text-xs">UID</span>
                              <p className="text-violet-400 font-mono text-xs mt-0.5">{team.leader.uid}</p>
                            </div>
                            <div>
                              <span className="text-slate-500 text-xs">Phone</span>
                              <p className="text-slate-300">{team.leader.phone || "—"}</p>
                            </div>
                            <div>
                              <span className="text-slate-500 text-xs">Email</span>
                              <p className="text-slate-300 truncate">{team.leader.email || "—"}</p>
                            </div>
                            <div>
                              <span className="text-slate-500 text-xs">College</span>
                              <p className="text-slate-300 truncate">{team.leader.college || "—"}</p>
                            </div>
                            <div>
                              <span className="text-slate-500 text-xs">Year / Program</span>
                              <p className="text-slate-300 truncate">
                                {team.leader.year || "—"} • {team.leader.program || "—"}
                              </p>
                            </div>
                          </div>
                        </div>

                        {/* Members */}
                        {team.members.length > 0 && (
                          <div className="space-y-2">
                            <div className="flex items-center gap-2 px-1">
                              <User className="w-3.5 h-3.5 text-slate-500" />
                              <span className="text-xs font-bold uppercase tracking-wider text-slate-500">
                                Members ({team.members.length})
                              </span>
                            </div>
                            {team.members.map((m, mi) => (
                              <div
                                key={mi}
                                className="rounded-xl bg-slate-800/30 border border-slate-800 p-4"
                              >
                                <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-x-6 gap-y-1.5 text-sm">
                                  <div>
                                    <span className="text-slate-500 text-xs">Name</span>
                                    <p className="text-white font-medium">{m.name}</p>
                                  </div>
                                  <div>
                                    <span className="text-slate-500 text-xs">UID</span>
                                    <p className="text-violet-400 font-mono text-xs mt-0.5">{m.uid}</p>
                                  </div>
                                  <div>
                                    <span className="text-slate-500 text-xs">Phone</span>
                                    <p className="text-slate-300">{m.phone || "—"}</p>
                                  </div>
                                  <div>
                                    <span className="text-slate-500 text-xs">Email</span>
                                    <p className="text-slate-300 truncate">{m.email || "—"}</p>
                                  </div>
                                  <div>
                                    <span className="text-slate-500 text-xs">College</span>
                                    <p className="text-slate-300 truncate">{m.college || "—"}</p>
                                  </div>
                                  <div>
                                    <span className="text-slate-500 text-xs">Year / Program</span>
                                    <p className="text-slate-300 truncate">
                                      {m.year || "—"} • {m.program || "—"}
                                    </p>
                                  </div>
                                </div>
                              </div>
                            ))}
                          </div>
                        )}
                      </div>
                    )}
                  </div>
                );
              })}

              {filteredTeams.length === 0 && (
                <div className="text-center py-16 text-slate-500">
                  <Users className="w-12 h-12 mx-auto mb-3 opacity-30" />
                  <p className="text-lg font-medium">No teams match your search.</p>
                </div>
              )}
            </div>
          </div>
        )}
      </div>
    </main>
  );
}
