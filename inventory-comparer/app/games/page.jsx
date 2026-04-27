"use client";

import { useState, useEffect, useRef } from "react";
import Link from "next/link";
import DinoGame from "./DinoGame";
import SpaceShooter from "./SpaceShooter";
import FPSGame from "./FPSGame";
import RacingGame from "./RacingGame";

const GAMES = [
  {
    id: "dino",
    title: "DINO RUN",
    emoji: "🦕",
    desc: "Dodge the cacti. Survive forever.",
    genre: "ENDLESS RUNNER",
    color: "#c8c8c8",
    glow: "rgba(200,200,200,0.5)",
    bg: "linear-gradient(135deg,#181818,#2e2e2e)",
    keys: "SPACE / ↑ to jump",
  },
  {
    id: "shooter",
    title: "SPACE INVADERS",
    emoji: "🚀",
    desc: "Blast the alien swarm before they land.",
    genre: "SHOOT EM UP",
    color: "#00e676",
    glow: "rgba(0,230,118,0.6)",
    bg: "linear-gradient(135deg,#001a0d,#003320)",
    keys: "← → Move · SPACE Shoot",
  },
  {
    id: "fps",
    title: "DUNGEON FPS",
    emoji: "🔫",
    desc: "Raycasted first-person dungeon. Kill them all.",
    genre: "FIRST PERSON SHOOTER",
    color: "#69ff69",
    glow: "rgba(105,255,105,0.5)",
    bg: "linear-gradient(135deg,#001400,#002000)",
    keys: "WASD Move · SPACE Shoot",
  },
  {
    id: "racing",
    title: "NEON RACER",
    emoji: "🏎️",
    desc: "3 laps. Beat the AI. Own the track.",
    genre: "TOP-DOWN RACER",
    color: "#ffe000",
    glow: "rgba(255,224,0,0.6)",
    bg: "linear-gradient(135deg,#1a1400,#2a2000)",
    keys: "WASD / ↑↓←→ Steer",
  },
];

const scanlineStyle = {
  backgroundImage:
    "repeating-linear-gradient(0deg,transparent,transparent 2px,rgba(0,0,0,0.13) 2px,rgba(0,0,0,0.13) 4px)",
  pointerEvents: "none",
};

export default function GamesPage() {
  const [selected, setSelected] = useState(null);
  const [fading, setFading] = useState(false);
  const [tick, setTick] = useState(0);
  const [particles] = useState(() =>
    Array.from({ length: 26 }, (_, i) => ({
      x: (i * 137.5) % 100,
      y: (i * 83.7) % 100,
      size: 2 + (i % 5),
      speed: 0.009 + (i % 7) * 0.004,
      hue: (i * 37) % 360,
      opacity: 0.2 + (i % 4) * 0.1,
    }))
  );
  const [ptPos, setPtPos] = useState(() => particles.map((p) => p.y));
  const animRef = useRef(null);
  const frameRef = useRef(0);

  useEffect(() => {
    const loop = () => {
      frameRef.current++;
      if (frameRef.current % 3 === 0) {
        setPtPos((prev) => prev.map((y, i) => (y - particles[i].speed + 100) % 100));
        setTick((t) => t + 1);
      }
      animRef.current = requestAnimationFrame(loop);
    };
    animRef.current = requestAnimationFrame(loop);
    return () => cancelAnimationFrame(animRef.current);
  }, [particles]);

  const navigate = (id) => {
    setFading(true);
    setTimeout(() => { setSelected(id); setFading(false); }, 300);
  };

  const activeGame = GAMES.find((g) => g.id === selected);

  return (
    <div
      className="min-h-screen flex flex-col overflow-hidden relative"
      style={{ background: "#06060f", fontFamily: "'Courier New', monospace" }}
    >
      {/* Scanlines overlay */}
      <div className="absolute inset-0 z-0" style={scanlineStyle} />

      {/* Floating particles */}
      <div className="absolute inset-0 z-0 overflow-hidden pointer-events-none">
        {particles.map((p, i) => (
          <div
            key={i}
            className="absolute rounded-sm"
            style={{
              left: `${p.x}%`,
              top: `${ptPos[i]}%`,
              width: p.size,
              height: p.size,
              background: `hsla(${p.hue},100%,70%,${p.opacity})`,
              boxShadow: `0 0 ${p.size * 2}px hsla(${p.hue},100%,70%,0.45)`,
            }}
          />
        ))}
      </div>

      {/* Corner brackets */}
      {[["top-0 left-0","top-3 left-3","#4a6dff"],["top-0 right-0","top-3 right-3","#e040fb"],["bottom-0 left-0","bottom-3 left-3","#00e676"],["bottom-0 right-0","bottom-3 right-3","#ffe000"]].map(([pos, inner, col], i) => (
        <div key={i} className={`absolute ${pos} w-24 h-24 z-0 pointer-events-none`}>
          <div className={`absolute ${inner} w-16 h-0.5 opacity-50`} style={{ background: col }} />
          <div className={`absolute ${inner} w-0.5 h-16 opacity-50`} style={{ background: col }} />
        </div>
      ))}

      {/* Main content */}
      <div
        className="relative z-10 flex flex-col items-center px-6 py-8 transition-opacity duration-300"
        style={{ opacity: fading ? 0 : 1 }}
      >
        {!selected ? (
          <>
            <Link href="/" className="self-start text-[#4a6dff] hover:text-[#7a9dff] text-sm mb-6 tracking-widest transition-colors">
              ← DASHBOARD
            </Link>

            <div className="text-center mb-2">
              <div className="text-6xl mb-3" style={{ animation: "bounce 1s infinite" }}>🕹️</div>
              <h1
                className="text-5xl font-black tracking-[0.2em] mb-1"
                style={{
                  background: `linear-gradient(90deg,#4a6dff,#e040fb,#00e676,#ffe000,#4a6dff)`,
                  backgroundSize: "300% 100%",
                  backgroundPosition: `${(tick * 1.5) % 300}% 0`,
                  WebkitBackgroundClip: "text",
                  WebkitTextFillColor: "transparent",
                }}
              >
                ARCADE
              </h1>
              <p className="text-xs tracking-[0.4em] mt-1" style={{ color: "#444" }}>
                ✦ SECRET GAME ROOM ✦
              </p>
            </div>

            {/* Blinking insert coin */}
            <p
              className="text-xs tracking-[0.5em] mt-4 mb-10 font-bold"
              style={{ color: "#ffe000", opacity: Math.floor(tick / 22) % 2 === 0 ? 1 : 0 }}
            >
              ▶ INSERT COIN · SELECT GAME ◀
            </p>

            {/* Game select grid */}
            <div className="grid grid-cols-1 md:grid-cols-2 gap-5 w-full max-w-[900px]">
              {GAMES.map((game) => (
                <button
                  key={game.id}
                  onClick={() => navigate(game.id)}
                  className="group relative text-left p-5 rounded-2xl border-2 transition-all duration-200 hover:-translate-y-1 overflow-hidden"
                  style={{ background: game.bg, borderColor: game.color + "80" }}
                  onMouseEnter={(e) => { e.currentTarget.style.borderColor = game.color; e.currentTarget.style.boxShadow = `0 0 28px ${game.glow}`; }}
                  onMouseLeave={(e) => { e.currentTarget.style.borderColor = game.color + "80"; e.currentTarget.style.boxShadow = "none"; }}
                >
                  <div className="absolute inset-0 rounded-2xl opacity-30 pointer-events-none" style={scanlineStyle} />
                  <div className="relative z-10 flex items-start gap-4">
                    <div className="text-4xl p-2 rounded-xl flex-shrink-0" style={{ background: "rgba(0,0,0,0.45)", border: `1px solid ${game.color}30` }}>
                      {game.emoji}
                    </div>
                    <div className="flex-1 min-w-0">
                      <p className="text-[9px] tracking-[0.4em] mb-1 uppercase" style={{ color: game.color, opacity: 0.7 }}>{game.genre}</p>
                      <h2 className="text-lg font-black tracking-wider mb-1 truncate" style={{ color: game.color }}>{game.title}</h2>
                      <p className="text-xs leading-relaxed" style={{ color: "#888" }}>{game.desc}</p>
                      <p className="text-[10px] mt-2 tracking-wider" style={{ color: game.color + "88" }}>🎮 {game.keys}</p>
                    </div>
                    <div className="self-center text-xl font-black transition-transform duration-150 group-hover:translate-x-1" style={{ color: game.color }}>▶</div>
                  </div>
                </button>
              ))}
            </div>

            <p className="text-[#1a1a1a] text-[10px] mt-14 tracking-[0.4em] select-none">↑↑↓↓←→←→BA</p>
          </>
        ) : (
          <div className="w-full max-w-[860px] flex flex-col items-center">
            {/* Top bar */}
            <div className="flex items-center justify-between w-full mb-5">
              <button onClick={() => navigate(null)} className="text-xs tracking-widest transition-opacity hover:opacity-60" style={{ color: activeGame.color }}>
                ← ARCADE
              </button>
              <span
                className="text-xs font-black tracking-[0.3em] px-4 py-1.5 rounded-full"
                style={{ color: activeGame.color, background: "rgba(0,0,0,0.5)", border: `1px solid ${activeGame.color}`, boxShadow: `0 0 14px ${activeGame.glow}` }}
              >
                {activeGame.emoji} {activeGame.title}
              </span>
              <div className="w-20" />
            </div>

            {/* Quick-switch tabs */}
            <div className="flex gap-2 mb-5 flex-wrap justify-center">
              {GAMES.map((g) => (
                <button
                  key={g.id}
                  onClick={() => navigate(g.id)}
                  className="text-[10px] px-3 py-1 rounded-full border tracking-wider transition-all"
                  style={
                    g.id === selected
                      ? { color: g.color, borderColor: g.color, background: "rgba(0,0,0,0.5)", boxShadow: `0 0 8px ${g.glow}` }
                      : { color: "#444", borderColor: "#2a2a2a", background: "transparent" }
                  }
                >
                  {g.emoji} {g.title}
                </button>
              ))}
            </div>

            {/* Game canvas wrapper with glow */}
            <div style={{ boxShadow: `0 0 48px ${activeGame.glow}`, borderRadius: 14 }}>
              {selected === "dino" && <DinoGame />}
              {selected === "shooter" && <SpaceShooter />}
              {selected === "fps" && <FPSGame />}
              {selected === "racing" && <RacingGame />}
            </div>
          </div>
        )}
      </div>
    </div>
  );
}
