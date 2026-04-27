"use client";

import { useEffect, useRef, useState } from "react";

const W = 800;
const H = 450;

// Track is defined as a sequence of [x, y] control points (loop)
const TRACK_POINTS = [
  [400, 60],[640, 60],[730, 130],[740, 220],[720, 330],[630, 400],
  [500, 420],[400, 410],[270, 420],[150, 390],[80, 300],[70, 180],
  [130, 100],[230, 60],
];

// Build a smooth closed curve from the track points
function catmullRom(p0, p1, p2, p3, t) {
  const t2 = t * t, t3 = t2 * t;
  return {
    x: 0.5 * ((2 * p1.x) + (-p0.x + p2.x) * t + (2 * p0.x - 5 * p1.x + 4 * p2.x - p3.x) * t2 + (-p0.x + 3 * p1.x - 3 * p2.x + p3.x) * t3),
    y: 0.5 * ((2 * p1.y) + (-p0.y + p2.y) * t + (2 * p0.y - 5 * p1.y + 4 * p2.y - p3.y) * t2 + (-p0.y + 3 * p1.y - 3 * p2.y + p3.y) * t3),
  };
}

function buildCurve(pts, steps = 10) {
  const n = pts.length;
  const result = [];
  for (let i = 0; i < n; i++) {
    const p0 = pts[(i - 1 + n) % n];
    const p1 = pts[i];
    const p2 = pts[(i + 1) % n];
    const p3 = pts[(i + 2) % n];
    for (let t = 0; t < 1; t += 1 / steps) {
      result.push(catmullRom(
        { x: p0[0], y: p0[1] }, { x: p1[0], y: p1[1] },
        { x: p2[0], y: p2[1] }, { x: p3[0], y: p3[1] }, t
      ));
    }
  }
  return result;
}

const CURVE = buildCurve(TRACK_POINTS, 18);
const CURVE_LEN = CURVE.length;

function getPosOnCurve(t) {
  const i = ((Math.floor(t * CURVE_LEN) % CURVE_LEN) + CURVE_LEN) % CURVE_LEN;
  return CURVE[i];
}

function getAngleOnCurve(t) {
  const a = getPosOnCurve(t);
  const b = getPosOnCurve(t + 1 / CURVE_LEN);
  return Math.atan2(b.y - a.y, b.x - a.x);
}

// Precompute normals for each curve point to define track width
function buildTrackBounds(curve, hw = 34) {
  const n = curve.length;
  const inner = [], outer = [];
  for (let i = 0; i < n; i++) {
    const next = curve[(i + 1) % n];
    const dx = next.x - curve[i].x, dy = next.y - curve[i].y;
    const len = Math.sqrt(dx * dx + dy * dy) || 1;
    const nx = -dy / len, ny = dx / len;
    inner.push({ x: curve[i].x + nx * hw, y: curve[i].y + ny * hw });
    outer.push({ x: curve[i].x - nx * hw, y: curve[i].y - ny * hw });
  }
  return { inner, outer };
}

const BOUNDS = buildTrackBounds(CURVE, 34);

function isOnTrack(x, y) {
  let minDist = Infinity;
  for (const pt of CURVE) {
    const d = (x - pt.x) ** 2 + (y - pt.y) ** 2;
    if (d < minDist) minDist = d;
  }
  return minDist < 34 * 34;
}

function makeState() {
  const startPos = CURVE[0];
  const startAngle = getAngleOnCurve(0);
  return {
    phase: "idle",
    x: startPos.x, y: startPos.y,
    angle: startAngle,
    speed: 0,
    lap: 0,
    lapProgress: 0,
    bestLap: null,
    lapStartTime: null,
    lastProgress: 0,
    lapCrossed: false,
    time: 0,
    score: 0,
    offTrackTimer: 0,
    finishPos: 1,
  };
}

// AI cars: they follow the curve at a fixed speed
const AI_COLORS = ["#f55", "#55f", "#fa0"];
const AI_OFFSETS = [0.33, 0.66, 0.5];
const DIFFICULTY_SPEEDS = {
  easy:   [0.00090, 0.00080, 0.00085],
  medium: [0.00138, 0.00122, 0.00130],
  hard:   [0.00178, 0.00170, 0.00184],
};

function makeAIState(speeds = DIFFICULTY_SPEEDS.medium) {
  return AI_OFFSETS.map((offset, i) => ({
    t: offset,
    speed: speeds[i],
    color: AI_COLORS[i],
    lap: 0,
    finished: false,
  }));
}

export default function RacingGame() {
  const [screen, setScreen] = useState("setup");
  const [lapCount, setLapCount] = useState(3);
  const [difficulty, setDifficulty] = useState("medium");
  const settingsRef = useRef({ lapCount: 3, difficulty: "medium" });
  const canvasRef = useRef(null);
  const stateRef = useRef(makeState());
  const aiRef = useRef(makeAIState());
  const keysRef = useRef({});
  const animRef = useRef(null);

  const startGame = (laps, diff) => {
    settingsRef.current = { lapCount: laps, difficulty: diff };
    const ns = makeState();
    ns.phase = "running";
    ns.lapStartTime = 0;
    stateRef.current = ns;
    aiRef.current = makeAIState(DIFFICULTY_SPEEDS[diff]);
    setScreen("game");
  };

  useEffect(() => {
    const canvas = canvasRef.current;
    const ctx = canvas.getContext("2d");

    const drawCar = (x, y, angle, color, shadow = false) => {
      ctx.save();
      ctx.translate(x, y);
      ctx.rotate(angle + Math.PI / 2);
      if (shadow) {
        ctx.fillStyle = "rgba(0,0,0,0.22)";
        ctx.fillRect(-9, -14, 18, 28);
      }
      ctx.fillStyle = color;
      ctx.fillRect(-8, -13, 16, 26);
      ctx.fillStyle = "rgba(255,255,255,0.6)";
      ctx.fillRect(-6, -13, 12, 7);
      ctx.fillRect(-6, 6, 12, 6);
      ctx.fillStyle = "#222";
      ctx.fillRect(-8, -13, 3, 7);
      ctx.fillRect(5, -13, 3, 7);
      ctx.fillRect(-8, 8, 3, 7);
      ctx.fillRect(5, 8, 3, 7);
      ctx.restore();
    };

    const drawTrack = () => {
      const n = CURVE.length;
      // Outer border (grass)
      ctx.fillStyle = "#2d7a2d";
      ctx.fillRect(0, 0, W, H);

      // Road shadow
      ctx.beginPath();
      for (let i = 0; i <= n; i++) {
        const o = BOUNDS.outer[i % n];
        i === 0 ? ctx.moveTo(o.x + 3, o.y + 3) : ctx.lineTo(o.x + 3, o.y + 3);
      }
      for (let i = n - 1; i >= 0; i--) {
        const b = BOUNDS.inner[i];
        ctx.lineTo(b.x + 3, b.y + 3);
      }
      ctx.closePath();
      ctx.fillStyle = "rgba(0,0,0,0.18)";
      ctx.fill();

      // Road surface
      ctx.beginPath();
      for (let i = 0; i <= n; i++) {
        const o = BOUNDS.outer[i % n];
        i === 0 ? ctx.moveTo(o.x, o.y) : ctx.lineTo(o.x, o.y);
      }
      for (let i = n - 1; i >= 0; i--) {
        ctx.lineTo(BOUNDS.inner[i].x, BOUNDS.inner[i].y);
      }
      ctx.closePath();
      ctx.fillStyle = "#333";
      ctx.fill();

      // Curbs (red-white alternating)
      for (let i = 0; i < n; i += 4) {
        const c = i % 8 === 0 ? "#e33" : "#fff";
        const o = BOUNDS.outer[i];
        const o2 = BOUNDS.outer[(i + 4) % n];
        ctx.beginPath();
        ctx.moveTo(o.x, o.y); ctx.lineTo(o2.x, o2.y);
        ctx.strokeStyle = c; ctx.lineWidth = 5; ctx.stroke();
        const ii = BOUNDS.inner[i];
        const ii2 = BOUNDS.inner[(i + 4) % n];
        ctx.beginPath();
        ctx.moveTo(ii.x, ii.y); ctx.lineTo(ii2.x, ii2.y);
        ctx.stroke();
      }

      // Center dashes
      ctx.setLineDash([18, 18]);
      ctx.strokeStyle = "#fff8";
      ctx.lineWidth = 2;
      ctx.beginPath();
      for (let i = 0; i <= n; i++) {
        const c = CURVE[i % n];
        i === 0 ? ctx.moveTo(c.x, c.y) : ctx.lineTo(c.x, c.y);
      }
      ctx.stroke();
      ctx.setLineDash([]);

      // Start/finish line
      const sf = CURVE[0];
      const sf2 = CURVE[5];
      const ang = Math.atan2(sf2.y - sf.y, sf2.x - sf.x);
      ctx.save();
      ctx.translate(sf.x, sf.y);
      ctx.rotate(ang + Math.PI / 2);
      for (let bx = -34; bx <= 34; bx += 8) {
        for (let by = -10; by <= 10; by += 10) {
          ctx.fillStyle = (bx + by) % 16 === 0 ? "#fff" : "#000";
          ctx.fillRect(bx, by, 8, 10);
        }
      }
      ctx.restore();
    };

    const loop = () => {
      const g = stateRef.current;

      if (g.phase === "running") {
        g.time++;

        // Player physics
        const MAX_SPEED = isOnTrack(g.x, g.y) ? 3.8 : 0.9;
        const ACCEL = 0.13;
        const BRAKE = 0.22;
        const FRICTION = 0.03;
        const STEER = 0.038;

        if (keysRef.current["ArrowUp"] || keysRef.current["KeyW"]) g.speed = Math.min(g.speed + ACCEL, MAX_SPEED);
        else if (keysRef.current["ArrowDown"] || keysRef.current["KeyS"]) g.speed = Math.max(g.speed - BRAKE, -1.2);
        else g.speed *= (1 - FRICTION);
        if (Math.abs(g.speed) > 0.1) {
          if (keysRef.current["ArrowLeft"] || keysRef.current["KeyA"]) g.angle -= STEER * (g.speed > 0 ? 1 : -1);
          if (keysRef.current["ArrowRight"] || keysRef.current["KeyD"]) g.angle += STEER * (g.speed > 0 ? 1 : -1);
        }

        const newX = g.x + Math.cos(g.angle) * g.speed;
        const newY = g.y + Math.sin(g.angle) * g.speed;
        g.x = Math.max(10, Math.min(W - 10, newX));
        g.y = Math.max(10, Math.min(H - 10, newY));

        if (!isOnTrack(g.x, g.y)) {
          g.offTrackTimer++;
          g.speed *= 0.80;
        } else {
          g.offTrackTimer = 0;
        }

        // Lap tracking: find closest curve t value
        let minD = Infinity, closestT = g.lastProgress;
        for (let i = 0; i < CURVE_LEN; i++) {
          const dx = g.x - CURVE[i].x, dy = g.y - CURVE[i].y;
          const d = dx * dx + dy * dy;
          if (d < minD) { minD = d; closestT = i / CURVE_LEN; }
        }
        const prevProgress = g.lastProgress;
        g.lastProgress = closestT;

        // Detect lap cross (t wraps 0.95 → 0.05)
        if (prevProgress > 0.9 && closestT < 0.1 && !g.lapCrossed) {
          g.lap++;
          g.lapCrossed = true;
          const elapsed = g.lapStartTime ? g.time - g.lapStartTime : null;
          if (elapsed && (g.bestLap === null || elapsed < g.bestLap)) g.bestLap = elapsed;
          g.lapStartTime = g.time;
          g.score += 500;
          if (g.lap >= settingsRef.current.lapCount) {
            g.finishPos = 1 + aiRef.current.filter(ai => ai.lap >= settingsRef.current.lapCount).length;
            g.phase = "win";
          }
        }
        if (closestT > 0.1 && closestT < 0.9) g.lapCrossed = false;

        // AI
        for (const ai of aiRef.current) {
          if (ai.finished) continue;
          const prevAiT = ai.t;
          ai.t = (ai.t + ai.speed) % 1;
          if (prevAiT > 0.9 && ai.t < 0.1) {
            ai.lap++;
            if (ai.lap >= settingsRef.current.lapCount) ai.finished = true;
          }
          const pos = getPosOnCurve(ai.t);
          const pos2 = getPosOnCurve(ai.t + 1 / CURVE_LEN);
          ai.x = pos.x; ai.y = pos.y;
          ai.angle = Math.atan2(pos2.y - pos.y, pos2.x - pos.x);
        }
      }

      // ---- DRAW ----
      drawTrack();

      // AI cars
      for (const ai of aiRef.current) {
        drawCar(ai.x, ai.y, ai.angle, ai.color, true);
      }

      // Player car
      drawCar(g.x, g.y, g.angle, "#ffe000", true);

      // Speed particles
      if (g.phase === "running" && Math.abs(g.speed) > 2 && g.offTrackTimer === 0) {
        for (let i = 0; i < 3; i++) {
          ctx.fillStyle = `rgba(255,180,0,${0.1 + Math.random() * 0.2})`;
          ctx.beginPath();
          ctx.arc(
            g.x - Math.cos(g.angle) * (14 + i * 5) + (Math.random() - 0.5) * 6,
            g.y - Math.sin(g.angle) * (14 + i * 5) + (Math.random() - 0.5) * 6,
            2 + Math.random() * 3, 0, Math.PI * 2
          );
          ctx.fill();
        }
      }
      // Dust when off track
      if (g.offTrackTimer > 0) {
        ctx.fillStyle = `rgba(180,140,80,0.25)`;
        ctx.beginPath();
        ctx.arc(g.x + (Math.random() - 0.5) * 14, g.y + (Math.random() - 0.5) * 14, 8 + Math.random() * 8, 0, Math.PI * 2);
        ctx.fill();
      }

      // HUD
      const totalLaps = settingsRef.current.lapCount;
      const currentLap = Math.min(g.lap + 1, totalLaps);
      ctx.fillStyle = "rgba(0,0,0,0.55)";
      ctx.fillRect(10, 10, 205, 64);
      ctx.font = "bold 14px monospace";
      ctx.fillStyle = "#ffe000";
      ctx.textAlign = "left";
      ctx.fillText(`LAP  ${currentLap}/${totalLaps}`, 22, 30);
      ctx.fillStyle = "#fff";
      ctx.fillText(`SPEED  ${Math.abs(g.speed * 60 | 0)} km/h`, 22, 48);
      const secs = g.bestLap ? (g.bestLap / 60).toFixed(2) : "--";
      ctx.fillText(`BEST   ${secs}s`, 22, 66);

      // Positions panel
      const racerEntries = [
        { label: "YOU", prog: g.lap + g.lastProgress, col: "#ffe000", isPlayer: true },
        ...aiRef.current.map((ai, j) => ({
          label: ["RED", "BLU", "ORG"][j], prog: ai.lap + ai.t, col: ai.color, isPlayer: false,
        })),
      ].sort((a, b) => b.prog - a.prog);
      ctx.fillStyle = "rgba(0,0,0,0.58)";
      ctx.fillRect(W - 120, 10, 110, 18 + racerEntries.length * 19);
      ctx.font = "bold 11px monospace";
      ctx.textAlign = "left";
      for (let ri = 0; ri < racerEntries.length; ri++) {
        const rc = racerEntries[ri];
        ctx.fillStyle = rc.isPlayer ? "#ffe000" : "#999";
        ctx.fillText(`P${ri + 1}  ${rc.label}${rc.isPlayer ? " ◄" : ""}`, W - 113, 26 + ri * 19);
      }

      // Off-track warning
      if (g.offTrackTimer > 20) {
        ctx.fillStyle = `rgba(255,60,0,${Math.min(0.7, g.offTrackTimer / 60)})`;
        ctx.font = "bold 22px monospace";
        ctx.textAlign = "center";
        ctx.fillText("OFF TRACK!", W / 2, 52);
      }

      // Overlays
      ctx.textAlign = "center";
      if (g.phase === "idle") {
        ctx.fillStyle = "rgba(0,0,0,0.6)";
        ctx.fillRect(0, 0, W, H);
        ctx.fillStyle = "#ffe000";
        ctx.font = "bold 38px monospace";
        ctx.fillText("NEON RACER", W / 2, H / 2 - 52);
        ctx.fillStyle = "#aaa";
        ctx.font = "13px monospace";
        ctx.fillText("WASD / ↑↓←→ to steer", W / 2, H / 2 - 8);
        ctx.fillText("Complete 3 laps  •  Off-road slows you down", W / 2, H / 2 + 14);
        ctx.fillStyle = "#ffe000";
        ctx.font = "bold 18px monospace";
        ctx.fillText("▶  Press SPACE to Race  ◀", W / 2, H / 2 + 58);
      } else if (g.phase === "win") {
        ctx.fillStyle = "rgba(0,0,0,0.7)";
        ctx.fillRect(0, 0, W, H);
        ctx.fillStyle = "#ffe000";
        ctx.font = "bold 40px monospace";
        ctx.fillText("FINISH! 🏁", W / 2, H / 2 - 28);
        ctx.fillStyle = "#fff";
        ctx.font = "18px monospace";
        const best = g.bestLap ? (g.bestLap / 60).toFixed(2) + "s" : "--";
        ctx.fillText(`Score: ${g.score}  •  Best Lap: ${best}`, W / 2, H / 2 + 18);
        ctx.fillStyle = "#aaa";
        ctx.font = "14px monospace";
        ctx.fillText("Press SPACE to Race Again", W / 2, H / 2 + 56);
      }

      animRef.current = requestAnimationFrame(loop);
    };

    const onKeyDown = (e) => {
      keysRef.current[e.code] = true;
      if (["Space", "ArrowUp", "ArrowDown", "ArrowLeft", "ArrowRight"].includes(e.code)) e.preventDefault();
      const g = stateRef.current;
      if (e.code === "Space" && g.phase === "win") {
        const ns = makeState();
        ns.phase = "running";
        ns.lapStartTime = 0;
        stateRef.current = ns;
        aiRef.current = makeAIState(DIFFICULTY_SPEEDS[settingsRef.current.difficulty]);
      }
      if (e.code === "Escape" && (g.phase === "win" || g.phase === "running")) {
        stateRef.current = makeState();
        aiRef.current = makeAIState();
        setScreen("setup");
      }
    };
    const onKeyUp = (e) => { keysRef.current[e.code] = false; };

    window.addEventListener("keydown", onKeyDown);
    window.addEventListener("keyup", onKeyUp);
    animRef.current = requestAnimationFrame(loop);

    return () => {
      cancelAnimationFrame(animRef.current);
      window.removeEventListener("keydown", onKeyDown);
      window.removeEventListener("keyup", onKeyUp);
    };
  }, []);

  return (
    <div className="relative flex flex-col items-center gap-3">
      <canvas ref={canvasRef} width={W} height={H}
        className="border-2 border-[#ffe000] rounded-xl max-w-full cursor-none"
      />
      {screen === "setup" && (
        <div className="absolute inset-0 flex items-center justify-center rounded-xl"
          style={{ background: "rgba(6,6,15,0.93)", fontFamily: "'Courier New', monospace" }}>
          <div className="flex flex-col items-center gap-5 p-8 max-w-[400px] w-full">
            <div className="text-5xl">🏁</div>
            <h2 className="text-3xl font-black tracking-widest" style={{ color: "#ffe000" }}>NEON RACER</h2>

            {/* Lap picker */}
            <div className="w-full">
              <p className="text-xs tracking-[0.4em] mb-2 text-center" style={{ color: "#ffe00088" }}>NUMBER OF LAPS</p>
              <div className="flex gap-2 justify-center">
                {[1, 3, 5, 10].map(n => (
                  <button key={n} onClick={() => setLapCount(n)}
                    className="w-14 h-10 rounded-lg font-black text-sm tracking-wider border-2 transition-all"
                    style={lapCount === n
                      ? { background: "#ffe000", color: "#000", borderColor: "#ffe000" }
                      : { background: "transparent", color: "#ffe000", borderColor: "#ffe00050" }}>
                    {n}
                  </button>
                ))}
              </div>
            </div>

            {/* Difficulty picker */}
            <div className="w-full">
              <p className="text-xs tracking-[0.4em] mb-2 text-center" style={{ color: "#ffe00088" }}>DIFFICULTY</p>
              <div className="flex gap-2">
                {[
                  { id: "easy",   label: "EASY", desc: "Slow AI",  col: "#00e676" },
                  { id: "medium", label: "MED",  desc: "Fair",     col: "#ffe000" },
                  { id: "hard",   label: "HARD", desc: "Fast AI",  col: "#f55" },
                ].map(d => (
                  <button key={d.id} onClick={() => setDifficulty(d.id)}
                    className="flex-1 py-2 rounded-lg font-black text-xs tracking-wider border-2 transition-all flex flex-col items-center gap-0.5"
                    style={difficulty === d.id
                      ? { background: d.col, color: "#000", borderColor: d.col }
                      : { background: "transparent", color: d.col, borderColor: d.col + "55" }}>
                    <span>{d.label}</span>
                    <span className="font-normal text-[9px] opacity-80">{d.desc}</span>
                  </button>
                ))}
              </div>
            </div>

            <button
              onClick={() => startGame(lapCount, difficulty)}
              className="w-full py-3 rounded-xl font-black text-lg tracking-widest border-2 transition-all hover:scale-105 active:scale-95"
              style={{ background: "#ffe000", color: "#000", borderColor: "#ffe000" }}>
              ▶ START RACE
            </button>
            <p className="text-[10px] tracking-widest" style={{ color: "#333" }}>WASD / ↑↓←→ Steer  •  ESC Change Settings</p>
          </div>
        </div>
      )}
      {screen === "game" && (
        <p className="text-xs tracking-wider" style={{ color: "#555" }}>WASD / ↑↓←→ Steer  •  SPACE Race Again  •  ESC Settings</p>
      )}
    </div>
  );
}
