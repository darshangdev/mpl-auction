import { useState, useRef } from "react";
import * as XLSX from "xlsx";

const B = {
  bg: "#000",
  surface: "#0a0a0a",
  surface2: "#111",
  surface3: "#1a1a1a",
  border: "#222",
  border2: "#333",
  text: "#fff",
  muted: "#666",
  dim: "#333",
  accent: "#fff",
};

const SAMPLE_PLAYERS = [
  { name: "Rohit Sharma", role: "Batsman", dept: "Engineering", avg: 48.2, sr: 138, matches: 42, specialty: "Opener", tier: "A" },
  { name: "Jasprit Bumrah", role: "Bowler", dept: "Design", avg: 12.1, sr: 98, matches: 38, specialty: "Death Overs", tier: "A" },
  { name: "Hardik Pandya", role: "All-Rounder", dept: "Product", avg: 35.6, sr: 145, matches: 40, specialty: "Finisher", tier: "B" },
  { name: "KL Rahul", role: "Batsman", dept: "Sales", avg: 44.1, sr: 130, matches: 35, specialty: "Anchor", tier: "A" },
  { name: "Ravindra Jadeja", role: "All-Rounder", dept: "HR", avg: 28.3, sr: 120, matches: 45, specialty: "Left-arm Spin", tier: "B" },
  { name: "Shubman Gill", role: "Batsman", dept: "Finance", avg: 51.0, sr: 142, matches: 30, specialty: "Top Order", tier: "B" },
  { name: "Mohammed Siraj", role: "Bowler", dept: "Marketing", avg: 14.2, sr: 90, matches: 28, specialty: "Swing Bowling", tier: "C" },
  { name: "Suryakumar Yadav", role: "Batsman", dept: "Ops", avg: 38.7, sr: 178, matches: 50, specialty: "360° Batting", tier: "A" },
  { name: "Virat Kohli", role: "Batsman", dept: "Leadership", avg: 59.3, sr: 140, matches: 60, specialty: "Chase Master", tier: "A" },
  { name: "Yuzvendra Chahal", role: "Bowler", dept: "Analytics", avg: 18.4, sr: 85, matches: 35, specialty: "Leg Spin", tier: "B" },
];

const ROLE_ICONS = { Batsman: "🏏", Bowler: "⚡", "All-Rounder": "◈", "Wicket-Keeper": "◉" };
const TIER_LABELS = { A: "ELITE", B: "STRONG", C: "RISING" };

function parseExcel(file, cb) {
  const reader = new FileReader();
  reader.onload = (e) => {
    const wb = XLSX.read(e.target.result, { type: "binary" });
    const ws = wb.Sheets[wb.SheetNames[0]];
    cb(XLSX.utils.sheet_to_json(ws).map(row => ({
      name: row["Name"] || row["name"] || row["Player"] || "Unknown",
      role: row["Role"] || row["role"] || "Batsman",
      dept: row["Department"] || row["dept"] || row["Team"] || "",
      avg: row["Average"] || row["avg"] || "—",
      sr: row["Strike Rate"] || row["SR"] || "—",
      matches: row["Matches"] || row["matches"] || "—",
      specialty: row["Specialty"] || row["specialty"] || "",
      tier: row["Tier"] || row["tier"] || "C",
      basePrice: row["Base Price"] || row["basePrice"] || "",
    })));
  };
  reader.readAsBinaryString(file);
}

function Logo({ small }) {
  return (
    <div style={{ display: "flex", alignItems: "center", gap: small ? 8 : 10 }}>
      <img src="https://www.metropolis.io/favicon.ico" alt=""
        style={{ width: small ? 22 : 28, height: small ? 22 : 28, borderRadius: 4, filter: "grayscale(1) brightness(2)", objectFit: "contain" }}
        onError={e => e.target.style.display = "none"} />
      <div>
        <div style={{ fontWeight: 900, fontSize: small ? 12 : 14, letterSpacing: 2, color: "#fff" }}>METROPOLIS</div>
        <div style={{ fontSize: small ? 8 : 9, color: B.muted, letterSpacing: 3 }}>PREMIER LEAGUE</div>
      </div>
    </div>
  );
}

function Divider() {
  return <div style={{ height: 1, background: B.border, width: "100%" }} />;
}

function StatBox({ label, value }) {
  return (
    <div style={{ textAlign: "center", padding: "12px 8px", border: `1px solid ${B.border}` }}>
      <div style={{ fontSize: 22, fontWeight: 900, color: "#fff", letterSpacing: -1 }}>{value ?? "—"}</div>
      <div style={{ fontSize: 9, color: B.muted, marginTop: 4, textTransform: "uppercase", letterSpacing: 2 }}>{label}</div>
    </div>
  );
}

function PlayerCard({ player, revealed, flipping }) {
  const role = player?.role || "Batsman";
  const icon = ROLE_ICONS[role] || "🏏";

  return (
    <div style={{ perspective: 1000, width: 320, height: 460, margin: "0 auto" }}>
      <div style={{
        width: "100%", height: "100%", position: "relative",
        transformStyle: "preserve-3d",
        transition: "transform 0.7s cubic-bezier(0.4,0,0.2,1)",
        transform: revealed && !flipping ? "rotateY(0deg)" : "rotateY(180deg)",
      }}>
        {/* BACK */}
        <div style={{
          position: "absolute", width: "100%", height: "100%", backfaceVisibility: "hidden",
          transform: "rotateY(180deg)",
          background: B.surface,
          border: `1px solid ${B.border2}`,
          display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center", gap: 16,
        }}>
          <img src="https://www.metropolis.io/favicon.ico" alt=""
            style={{ width: 36, height: 36, borderRadius: 6, filter: "grayscale(1) brightness(2)", opacity: 0.6 }}
            onError={e => e.target.style.display = "none"} />
          <div style={{ textAlign: "center" }}>
            <div style={{ fontWeight: 900, fontSize: 16, letterSpacing: 4, color: "#fff" }}>METROPOLIS</div>
            <div style={{ fontSize: 9, color: B.muted, letterSpacing: 4, marginTop: 4 }}>PREMIER LEAGUE</div>
          </div>
          <Divider />
          <div style={{ fontSize: 9, color: B.dim, letterSpacing: 3 }}>PLAYER AUCTION</div>
        </div>

        {/* FRONT */}
        <div style={{
          position: "absolute", width: "100%", height: "100%", backfaceVisibility: "hidden",
          background: B.surface, border: `1px solid ${B.border2}`,
          display: "flex", flexDirection: "column", boxSizing: "border-box",
        }}>
          {player && <>
            {/* Top bar */}
            <div style={{ padding: "14px 18px", borderBottom: `1px solid ${B.border}`, display: "flex", justifyContent: "space-between", alignItems: "center" }}>
              <div style={{ fontSize: 9, color: B.muted, letterSpacing: 3 }}>{TIER_LABELS[player.tier] || "PLAYER"}</div>
              <div style={{ fontSize: 9, color: B.muted, letterSpacing: 3 }}>{role.toUpperCase()}</div>
            </div>

            {/* Name block */}
            <div style={{ padding: "22px 18px 18px", flex: 1 }}>
              <div style={{ fontSize: 28, fontWeight: 900, color: "#fff", lineHeight: 1.05, letterSpacing: -0.5 }}>{player.name}</div>
              <div style={{ fontSize: 11, color: B.muted, marginTop: 6, letterSpacing: 1 }}>{player.dept || "—"}</div>
              {player.specialty && (
                <div style={{ marginTop: 12, fontSize: 10, color: B.muted, letterSpacing: 2, borderLeft: `2px solid ${B.border2}`, paddingLeft: 10 }}>
                  {player.specialty.toUpperCase()}
                </div>
              )}
            </div>

            {/* Stats */}
            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", borderTop: `1px solid ${B.border}` }}>
              <StatBox label="Matches" value={player.matches} />
              <div style={{ borderLeft: `1px solid ${B.border}`, borderRight: `1px solid ${B.border}` }}>
                <StatBox label={role === "Bowler" ? "Economy" : "Average"} value={player.avg} />
              </div>
              <StatBox label={role === "Bowler" ? "Wickets" : "Strike Rate"} value={player.sr} />
            </div>

            {/* Price */}
            <div style={{ padding: "14px 18px", borderTop: `1px solid ${B.border}`, display: "flex", justifyContent: "space-between", alignItems: "center" }}>
              <div style={{ fontSize: 9, color: B.muted, letterSpacing: 3 }}>BASE PRICE</div>
              <div style={{ fontSize: 20, fontWeight: 900, color: "#fff" }}>
                ₹{player.basePrice || (player.tier === "A" ? "2,000" : player.tier === "B" ? "1,000" : "500")}
              </div>
            </div>
          </>}
        </div>
      </div>
    </div>
  );
}

// ── SETUP ─────────────────────────────────────────────────────────────────────
function SetupPage({ onStart }) {
  const [players, setPlayers] = useState(SAMPLE_PLAYERS);
  const [teams, setTeams] = useState([{ name: "Team Alpha" }, { name: "Team Beta" }]);
  const [newTeam, setNewTeam] = useState("");
  const [budget, setBudget] = useState("10000");
  const [usingSample, setUsingSample] = useState(true);
  const fileRef = useRef();

  const addTeam = () => {
    if (!newTeam.trim() || teams.length >= 8) return;
    setTeams(t => [...t, { name: newTeam.trim() }]);
    setNewTeam("");
  };
  const removeTeam = (i) => setTeams(t => t.filter((_, idx) => idx !== i));
  const updateTeamName = (i, val) => setTeams(t => t.map((tm, idx) => idx === i ? { ...tm, name: val } : tm));
  const handleFile = (e) => {
    const file = e.target.files[0];
    if (!file) return;
    parseExcel(file, (data) => { setPlayers(data); setUsingSample(false); });
  };
  const canStart = teams.length >= 2;

  const inp = { background: B.surface2, border: `1px solid ${B.border2}`, color: "#fff", padding: "10px 14px", borderRadius: 0, fontSize: 14, outline: "none", width: "100%", boxSizing: "border-box", fontFamily: "inherit" };
  const btn = (active = true) => ({ background: active ? "#fff" : B.surface2, color: active ? "#000" : B.muted, border: `1px solid ${active ? "#fff" : B.border2}`, padding: "10px 20px", cursor: active ? "pointer" : "default", fontWeight: 700, fontSize: 13, letterSpacing: 1, fontFamily: "inherit" });

  return (
    <div style={{ minHeight: "100vh", background: B.bg, color: "#fff", fontFamily: "'Segoe UI',system-ui,sans-serif" }}>
      {/* Header */}
      <div style={{ padding: "18px 28px", borderBottom: `1px solid ${B.border}`, display: "flex", justifyContent: "space-between", alignItems: "center" }}>
        <Logo />
        <div style={{ fontSize: 9, color: B.muted, letterSpacing: 4 }}>AUCTION SETUP</div>
      </div>

      {/* Title */}
      <div style={{ padding: "48px 28px 36px", borderBottom: `1px solid ${B.border}` }}>
        <div style={{ fontSize: 9, color: B.muted, letterSpacing: 4, marginBottom: 10 }}>SEASON 01</div>
        <div style={{ fontSize: 36, fontWeight: 900, letterSpacing: -1, lineHeight: 1 }}>PLAYER<br />AUCTION</div>
      </div>

      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", borderBottom: `1px solid ${B.border}` }}>
        {/* LEFT — Teams */}
        <div style={{ borderRight: `1px solid ${B.border}`, padding: "28px" }}>
          <div style={{ fontSize: 9, color: B.muted, letterSpacing: 4, marginBottom: 20 }}>TEAMS — {teams.length}/8</div>

          <div style={{ display: "flex", flexDirection: "column", gap: 1, marginBottom: 16 }}>
            {teams.map((team, i) => (
              <div key={i} style={{ display: "flex", alignItems: "center", gap: 0, border: `1px solid ${B.border}` }}>
                <div style={{ width: 40, height: 40, display: "flex", alignItems: "center", justifyContent: "center", borderRight: `1px solid ${B.border}`, color: B.muted, fontSize: 11, fontWeight: 700, flexShrink: 0 }}>{i + 1}</div>
                <input value={team.name} onChange={e => updateTeamName(i, e.target.value)}
                  style={{ ...inp, border: "none", flex: 1, padding: "10px 14px" }} />
                <button onClick={() => removeTeam(i)}
                  style={{ width: 40, height: 40, background: "none", border: "none", borderLeft: `1px solid ${B.border}`, color: B.dim, cursor: "pointer", fontSize: 16, flexShrink: 0 }}>×</button>
              </div>
            ))}
          </div>

          <div style={{ display: "flex", gap: 0 }}>
            <input placeholder="Team name" value={newTeam} onChange={e => setNewTeam(e.target.value)}
              onKeyDown={e => e.key === "Enter" && addTeam()}
              style={{ ...inp, flex: 1, borderRight: "none" }} />
            <button onClick={addTeam} style={{ ...btn(), padding: "10px 16px", flexShrink: 0 }}>+</button>
          </div>

          <div style={{ marginTop: 28, borderTop: `1px solid ${B.border}`, paddingTop: 24 }}>
            <div style={{ fontSize: 9, color: B.muted, letterSpacing: 4, marginBottom: 12 }}>BUDGET PER TEAM (₹)</div>
            <input type="number" value={budget} onChange={e => setBudget(e.target.value)}
              style={{ ...inp, fontSize: 22, fontWeight: 900, padding: "12px 14px" }} />
          </div>
        </div>

        {/* RIGHT — Players */}
        <div style={{ padding: "28px" }}>
          <div style={{ fontSize: 9, color: B.muted, letterSpacing: 4, marginBottom: 20 }}>PLAYERS — {players.length}</div>

          <input ref={fileRef} type="file" accept=".xlsx,.xls" onChange={handleFile} style={{ display: "none" }} />
          <div style={{ border: `1px solid ${B.border}`, padding: "24px", textAlign: "center", marginBottom: 16, cursor: "pointer" }} onClick={() => fileRef.current.click()}>
            <div style={{ fontSize: 9, color: B.muted, letterSpacing: 4, marginBottom: 8 }}>UPLOAD EXCEL</div>
            <div style={{ fontSize: 11, color: B.dim, marginBottom: 14, lineHeight: 1.8 }}>Name · Role · Department<br />Average · Strike Rate · Matches<br />Specialty · Tier · Base Price</div>
            <button style={{ ...btn(), fontSize: 11, letterSpacing: 2 }}>CHOOSE FILE</button>
            <div style={{ fontSize: 10, color: B.dim, marginTop: 12 }}>
              {usingSample ? `${players.length} sample players loaded` : `✓ ${players.length} players from Excel`}
            </div>
          </div>

          <div style={{ maxHeight: 240, overflowY: "auto" }}>
            {players.map((p, i) => (
              <div key={i} style={{ display: "flex", alignItems: "center", gap: 12, padding: "8px 0", borderBottom: `1px solid ${B.border}` }}>
                <div style={{ fontSize: 10, color: B.dim, width: 20, textAlign: "right", flexShrink: 0 }}>{i + 1}</div>
                <div style={{ flex: 1 }}>
                  <div style={{ fontSize: 13, fontWeight: 700 }}>{p.name}</div>
                  <div style={{ fontSize: 10, color: B.muted }}>{p.role} · {p.dept}</div>
                </div>
                <div style={{ fontSize: 9, color: B.muted, letterSpacing: 1 }}>{TIER_LABELS[p.tier] || "C"}</div>
              </div>
            ))}
          </div>
        </div>
      </div>

      {/* Start */}
      <div style={{ padding: "32px 28px", display: "flex", alignItems: "center", justifyContent: "space-between" }}>
        <div style={{ fontSize: 10, color: B.dim }}>
          {canStart ? `${teams.length} teams · ${players.length} players · ₹${Number(budget).toLocaleString()} budget` : "Add at least 2 teams to continue"}
        </div>
        <button onClick={() => canStart && onStart(teams, players, Number(budget))} style={{ ...btn(canStart), padding: "14px 40px", fontSize: 12, letterSpacing: 3 }}>
          START AUCTION →
        </button>
      </div>
    </div>
  );
}

// ── TEAMS VIEW ────────────────────────────────────────────────────────────────
function TeamsView({ teams, soldLog, budget, onClose }) {
  const [activeTab, setActiveTab] = useState(0);
  const roleGroups = ["Batsman", "All-Rounder", "Bowler", "Wicket-Keeper"];

  return (
    <div style={{ position: "fixed", inset: 0, background: B.bg, zIndex: 200, display: "flex", flexDirection: "column", fontFamily: "'Segoe UI',system-ui,sans-serif", color: "#fff" }}>
      {/* Header */}
      <div style={{ padding: "14px 22px", borderBottom: `1px solid ${B.border}`, display: "flex", alignItems: "center", justifyContent: "space-between" }}>
        <Logo small />
        <div style={{ fontSize: 9, color: B.muted, letterSpacing: 4 }}>TEAM SQUADS</div>
        <button onClick={onClose} style={{ background: "none", border: `1px solid ${B.border2}`, color: B.muted, padding: "6px 16px", cursor: "pointer", fontSize: 12, fontFamily: "inherit", letterSpacing: 1 }}>CLOSE</button>
      </div>

      {/* Tabs */}
      <div style={{ display: "flex", overflowX: "auto", borderBottom: `1px solid ${B.border}` }}>
        {teams.map((t, i) => (
          <button key={i} onClick={() => setActiveTab(i)}
            style={{ padding: "12px 20px", background: activeTab === i ? "#fff" : "none", border: "none", borderRight: `1px solid ${B.border}`, color: activeTab === i ? "#000" : B.muted, cursor: "pointer", fontWeight: activeTab === i ? 800 : 400, fontSize: 12, whiteSpace: "nowrap", fontFamily: "inherit", letterSpacing: 1, transition: "all 0.15s" }}>
            {t.name}
          </button>
        ))}
      </div>

      {/* Content */}
      {teams.map((team, ti) => {
        if (ti !== activeTab) return null;
        const tp = soldLog.filter(p => p.soldTo === team.name);
        const spent = tp.reduce((s, p) => s + (Number(p.soldFor) || 0), 0);
        return (
          <div key={ti} style={{ flex: 1, overflowY: "auto" }}>
            {/* Budget row */}
            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", borderBottom: `1px solid ${B.border}` }}>
              {[["PLAYERS", tp.length], ["SPENT", `₹${spent.toLocaleString()}`], ["REMAINING", `₹${(budget - spent).toLocaleString()}`]].map(([label, val]) => (
                <div key={label} style={{ padding: "20px 24px", borderRight: `1px solid ${B.border}` }}>
                  <div style={{ fontSize: 9, color: B.muted, letterSpacing: 3, marginBottom: 6 }}>{label}</div>
                  <div style={{ fontSize: 22, fontWeight: 900 }}>{val}</div>
                </div>
              ))}
            </div>
            {/* Budget bar */}
            <div style={{ height: 2, background: B.border }}>
              <div style={{ height: "100%", width: `${Math.min(100, (spent / budget) * 100)}%`, background: "#fff", transition: "width 0.5s" }} />
            </div>

            {tp.length === 0 ? (
              <div style={{ textAlign: "center", color: B.dim, padding: "80px 0", fontSize: 11, letterSpacing: 2 }}>NO PLAYERS YET</div>
            ) : (
              <div style={{ padding: "28px" }}>
                {roleGroups.map(role => {
                  const rp = tp.filter(p => p.role === role);
                  if (!rp.length) return null;
                  return (
                    <div key={role} style={{ marginBottom: 28 }}>
                      <div style={{ fontSize: 9, color: B.muted, letterSpacing: 4, marginBottom: 14, paddingBottom: 10, borderBottom: `1px solid ${B.border}` }}>{role.toUpperCase()}S — {rp.length}</div>
                      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill,minmax(210px,1fr))", gap: 1 }}>
                        {rp.map((p, pi) => (
                          <div key={pi} style={{ border: `1px solid ${B.border}`, padding: "14px 16px" }}>
                            <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 8 }}>
                              <div style={{ fontWeight: 800, fontSize: 14 }}>{p.name}</div>
                              <div style={{ fontSize: 12, fontWeight: 900 }}>₹{Number(p.soldFor || 0).toLocaleString()}</div>
                            </div>
                            <div style={{ fontSize: 10, color: B.muted }}>{p.dept}</div>
                            {p.specialty && <div style={{ fontSize: 10, color: B.dim, marginTop: 4 }}>{p.specialty}</div>}
                            <div style={{ display: "flex", gap: 16, marginTop: 10 }}>
                              <div style={{ fontSize: 10, color: B.dim }}>AVG <span style={{ color: B.muted }}>{p.avg}</span></div>
                              <div style={{ fontSize: 10, color: B.dim }}>SR <span style={{ color: B.muted }}>{p.sr}</span></div>
                              <div style={{ fontSize: 10, color: B.dim }}>M <span style={{ color: B.muted }}>{p.matches}</span></div>
                            </div>
                          </div>
                        ))}
                      </div>
                    </div>
                  );
                })}
              </div>
            )}
          </div>
        );
      })}
    </div>
  );
}

// ── AUCTION PAGE ──────────────────────────────────────────────────────────────
function AuctionPage({ teams, players, budget, onReset }) {
  const shuffled = [...players].sort(() => Math.random() - 0.5);
  const [queue, setQueue] = useState(shuffled.slice(1));
  const [current, setCurrent] = useState(shuffled[0]);
  const [revealed, setRevealed] = useState(false);
  const [flipping, setFlipping] = useState(false);
  const [done, setDone] = useState(false);
  const [soldLog, setSoldLog] = useState([]);
  const [bidAmount, setBidAmount] = useState("");
  const [soldTo, setSoldTo] = useState("");
  const [showSold, setShowSold] = useState(false);
  const [showTeams, setShowTeams] = useState(false);

  const revealPlayer = () => { setFlipping(true); setTimeout(() => { setRevealed(true); setFlipping(false); }, 50); };

  const markSold = () => {
    if (!current || !soldTo) return;
    setSoldLog(prev => [...prev, { ...current, soldTo, soldFor: bidAmount }]);
    setShowSold(true);
    setTimeout(() => { setShowSold(false); nextPlayer(); }, 1800);
  };

  const nextPlayer = () => {
    setBidAmount(""); setSoldTo(""); setShowSold(false);
    setFlipping(true); setRevealed(false);
    setTimeout(() => {
      if (queue.length === 0) { setDone(true); setCurrent(null); }
      else { setCurrent(queue[0]); setQueue(q => q.slice(1)); }
      setFlipping(false);
    }, 500);
  };

  const inp = { background: B.surface2, border: `1px solid ${B.border2}`, color: "#fff", padding: "10px 14px", fontSize: 14, outline: "none", fontFamily: "inherit" };
  const getSpent = (name) => soldLog.filter(p => p.soldTo === name).reduce((s, p) => s + (Number(p.soldFor) || 0), 0);

  return (
    <div style={{ minHeight: "100vh", background: B.bg, color: "#fff", fontFamily: "'Segoe UI',system-ui,sans-serif", display: "flex", flexDirection: "column" }}>
      {showTeams && <TeamsView teams={teams} soldLog={soldLog} budget={budget} onClose={() => setShowTeams(false)} />}

      {/* Header */}
      <div style={{ padding: "12px 22px", borderBottom: `1px solid ${B.border}`, display: "flex", alignItems: "center", justifyContent: "space-between" }}>
        <Logo small />
        <div style={{ display: "flex", gap: 10, alignItems: "center" }}>
          {!done && <div style={{ fontSize: 10, color: B.muted, letterSpacing: 2 }}>{queue.length} REMAINING</div>}
          <button onClick={() => setShowTeams(true)}
            style={{ background: "none", border: `1px solid ${B.border2}`, color: "#fff", padding: "6px 14px", cursor: "pointer", fontSize: 11, fontFamily: "inherit", letterSpacing: 2 }}>
            SQUADS ({soldLog.length})
          </button>
          <button onClick={onReset}
            style={{ background: "none", border: `1px solid ${B.border}`, color: B.muted, padding: "6px 12px", cursor: "pointer", fontSize: 11, fontFamily: "inherit" }}>
            ⚙
          </button>
        </div>
      </div>

      {/* Budget strip */}
      <div style={{ display: "flex", overflowX: "auto", borderBottom: `1px solid ${B.border}` }}>
        {teams.map((t, i) => {
          const spent = getSpent(t.name);
          const pct = Math.min(100, (spent / budget) * 100);
          return (
            <div key={i} style={{ padding: "8px 16px", minWidth: 120, borderRight: `1px solid ${B.border}` }}>
              <div style={{ fontSize: 10, color: "#fff", fontWeight: 700, whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis", maxWidth: 100 }}>{t.name}</div>
              <div style={{ height: 1, background: B.border2, margin: "5px 0", overflow: "hidden" }}>
                <div style={{ height: "100%", width: `${pct}%`, background: "#fff" }} />
              </div>
              <div style={{ fontSize: 10, color: B.muted }}>₹{(budget - spent).toLocaleString()}</div>
            </div>
          );
        })}
      </div>

      {/* Main */}
      <div style={{ flex: 1, display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center", padding: "32px 16px" }}>
        {done ? (
          <div style={{ textAlign: "center", maxWidth: 400 }}>
            <div style={{ fontSize: 9, color: B.muted, letterSpacing: 4, marginBottom: 16 }}>AUCTION COMPLETE</div>
            <div style={{ fontSize: 48, fontWeight: 900, letterSpacing: -2, lineHeight: 1, marginBottom: 8 }}>{soldLog.length}</div>
            <div style={{ fontSize: 9, color: B.muted, letterSpacing: 4, marginBottom: 36 }}>PLAYERS SOLD</div>
            <Divider />
            <div style={{ display: "flex", gap: 0, marginTop: 24, justifyContent: "center" }}>
              <button onClick={() => setShowTeams(true)}
                style={{ background: "#fff", color: "#000", border: "none", padding: "12px 28px", cursor: "pointer", fontWeight: 800, fontSize: 12, letterSpacing: 2, fontFamily: "inherit" }}>
                VIEW SQUADS
              </button>
              <button onClick={onReset}
                style={{ background: "none", border: `1px solid ${B.border2}`, color: B.muted, padding: "12px 24px", cursor: "pointer", fontSize: 12, letterSpacing: 2, fontFamily: "inherit", marginLeft: 8 }}>
                NEW AUCTION
              </button>
            </div>
          </div>
        ) : (
          <div style={{ width: "100%", display: "flex", flexDirection: "column", alignItems: "center", gap: 24 }}>
            <PlayerCard player={current} revealed={revealed} flipping={flipping} />

            {/* SOLD overlay */}
            {showSold && (
              <div style={{ position: "fixed", inset: 0, display: "flex", alignItems: "center", justifyContent: "center", background: "rgba(0,0,0,0.92)", zIndex: 999 }}>
                <div style={{ border: `1px solid ${B.border2}`, padding: "52px 72px", textAlign: "center", background: B.surface }}>
                  <div style={{ fontSize: 9, color: B.muted, letterSpacing: 6, marginBottom: 16 }}>HAMMER DOWN</div>
                  <div style={{ fontSize: 48, fontWeight: 900, color: "#fff", letterSpacing: 4 }}>SOLD</div>
                  {soldTo && <div style={{ color: B.muted, fontSize: 13, marginTop: 12, letterSpacing: 2 }}>{soldTo.toUpperCase()}</div>}
                  {bidAmount && <div style={{ color: "#fff", fontSize: 28, fontWeight: 900, marginTop: 6 }}>₹{Number(bidAmount).toLocaleString()}</div>}
                </div>
              </div>
            )}

            {!revealed ? (
              <button onClick={revealPlayer}
                style={{ background: "#fff", color: "#000", border: "none", padding: "14px 52px", fontSize: 12, fontWeight: 900, cursor: "pointer", letterSpacing: 4, fontFamily: "inherit" }}>
                REVEAL PLAYER
              </button>
            ) : (
              <div style={{ display: "flex", flexDirection: "column", alignItems: "center", gap: 10, width: "100%", maxWidth: 400 }}>
                <div style={{ display: "flex", gap: 0, width: "100%" }}>
                  <select value={soldTo} onChange={e => setSoldTo(e.target.value)}
                    style={{ ...inp, flex: 1, borderRight: "none", cursor: "pointer" }}>
                    <option value="">Select team →</option>
                    {teams.map((t, i) => <option key={i} value={t.name}>{t.name} — ₹{(budget - getSpent(t.name)).toLocaleString()} left</option>)}
                  </select>
                  <input placeholder="₹" value={bidAmount} onChange={e => setBidAmount(e.target.value)} type="number"
                    style={{ ...inp, width: 90, textAlign: "right" }} />
                </div>
                <div style={{ display: "flex", gap: 8, width: "100%" }}>
                  <button onClick={markSold} disabled={!soldTo}
                    style={{ flex: 1, background: soldTo ? "#fff" : B.surface2, color: soldTo ? "#000" : B.dim, border: `1px solid ${soldTo ? "#fff" : B.border}`, padding: "12px", cursor: soldTo ? "pointer" : "default", fontWeight: 900, fontSize: 13, letterSpacing: 3, fontFamily: "inherit", transition: "all 0.15s" }}>
                    SOLD
                  </button>
                  <button onClick={nextPlayer}
                    style={{ background: "none", color: B.muted, border: `1px solid ${B.border}`, padding: "12px 20px", cursor: "pointer", fontSize: 12, letterSpacing: 2, fontFamily: "inherit" }}>
                    UNSOLD
                  </button>
                </div>
              </div>
            )}

            {/* Recent sales */}
            {soldLog.length > 0 && (
              <div style={{ width: "100%", maxWidth: 400, borderTop: `1px solid ${B.border}`, paddingTop: 16 }}>
                <div style={{ fontSize: 9, color: B.dim, letterSpacing: 3, marginBottom: 10 }}>RECENT</div>
                <div style={{ display: "flex", flexDirection: "column", gap: 1 }}>
                  {soldLog.slice(-4).reverse().map((p, i) => (
                    <div key={i} style={{ display: "flex", justifyContent: "space-between", padding: "6px 0", borderBottom: `1px solid ${B.border}`, fontSize: 12 }}>
                      <span style={{ color: i === 0 ? "#fff" : B.muted }}>{p.name}</span>
                      <span style={{ color: B.dim }}>{p.soldTo}</span>
                      <span style={{ color: i === 0 ? "#fff" : B.dim }}>₹{Number(p.soldFor || 0).toLocaleString()}</span>
                    </div>
                  ))}
                </div>
              </div>
            )}
          </div>
        )}
      </div>
    </div>
  );
}

// ── ROOT ──────────────────────────────────────────────────────────────────────
export default function App() {
  const [page, setPage] = useState("setup");
  const [config, setConfig] = useState(null);
  if (page === "auction" && config) {
    return <AuctionPage teams={config.teams} players={config.players} budget={config.budget} onReset={() => setPage("setup")} />;
  }
  return <SetupPage onStart={(t, p, b) => { setConfig({ teams: t, players: p, budget: b }); setPage("auction"); }} />;
}
