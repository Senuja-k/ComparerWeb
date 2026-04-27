"use client";

import { useEffect, useRef } from "react";

const W = 800;
const H = 300;
const GROUND_Y = 245;
const GRAVITY = 0.9;
const JUMP_FORCE = -15;

const makeState = (prev) => ({
  phase: "idle",
  score: 0,
  speed: 5,
  frame: 0,
  hiScore: prev?.hiScore ?? 0,
  dino: { x: 80, y: GROUND_Y - 52, w: 44, h: 52, vy: 0, onGround: true },
  obstacles: [],
  clouds: [
    { x: 180, y: 38 },
    { x: 490, y: 26 },
    { x: 710, y: 52 },
  ],
});

export default function DinoGame() {
  const canvasRef = useRef(null);
  const stateRef = useRef(makeState());
  const animIdRef = useRef(null);

  useEffect(() => {
    const canvas = canvasRef.current;
    const ctx = canvas.getContext("2d");

    const reset = () => {
      stateRef.current = { ...makeState(stateRef.current), phase: "running" };
    };

    const jump = () => {
      const s = stateRef.current;
      if (s.phase === "idle" || s.phase === "over") { reset(); return; }
      if (s.dino.onGround) { s.dino.vy = JUMP_FORCE; s.dino.onGround = false; }
    };

    const drawDino = (d, frame) => {
      ctx.fillStyle = "#535353";
      ctx.fillRect(d.x, d.y + 14, 38, 28);           // body
      ctx.fillRect(d.x + 18, d.y, 24, 22);            // head
      ctx.fillRect(d.x - 12, d.y + 20, 14, 8);        // tail tip
      ctx.fillRect(d.x - 8, d.y + 14, 10, 10);        // tail base
      ctx.fillRect(d.x + 38, d.y + 16, 6, 4);         // mouth
      ctx.fillStyle = "#fff";
      ctx.fillRect(d.x + 30, d.y + 4, 7, 7);          // eye white
      ctx.fillStyle = "#1a1a1a";
      ctx.fillRect(d.x + 33, d.y + 6, 4, 4);          // pupil
      ctx.fillStyle = "#535353";
      const leg = d.onGround ? Math.floor(frame / 5) % 2 : 0;
      ctx.fillRect(d.x + 8,  d.y + 40, 9, 12 + leg * 3);
      ctx.fillRect(d.x + 22, d.y + 40, 9, 14 - leg * 3 + 1);
    };

    const drawCactus = (o) => {
      ctx.fillStyle = "#2d6a2d";
      const mx = o.x + Math.floor(o.w / 2) - 6;
      ctx.fillRect(mx, o.y, 12, o.h);                                                // trunk
      ctx.fillRect(o.x, o.y + Math.floor(o.h * 0.30), Math.floor(o.w / 2) - 6, 8); // L arm
      ctx.fillRect(o.x, o.y + Math.floor(o.h * 0.10), 10, Math.floor(o.h * 0.25)); // L stalk
      ctx.fillRect(o.x + Math.floor(o.w / 2) + 6, o.y + Math.floor(o.h * 0.35), Math.floor(o.w / 2) - 6, 8); // R arm
      ctx.fillRect(o.x + o.w - 10, o.y + Math.floor(o.h * 0.15), 10, Math.floor(o.h * 0.30)); // R stalk
    };

    const loop = () => {
      const s = stateRef.current;

      if (s.phase === "running") {
        s.frame++;
        s.score += 0.1 + s.speed * 0.004;
        s.speed = 5 + Math.floor(s.score / 150) * 0.5;

        // Dino physics
        const d = s.dino;
        if (!d.onGround) {
          d.vy += GRAVITY;
          d.y += d.vy;
          if (d.y >= GROUND_Y - d.h) {
            d.y = GROUND_Y - d.h;
            d.vy = 0;
            d.onGround = true;
          }
        }

        // Clouds
        for (const c of s.clouds) {
          c.x -= 0.8;
          if (c.x < -90) c.x = W + 50 + Math.random() * 120;
        }

        // Spawn obstacles
        const last = s.obstacles[s.obstacles.length - 1];
        if (!last || last.x < W - 220 - Math.random() * 260) {
          const tall = Math.random() > 0.4;
          const h = tall ? 65 : 40;
          s.obstacles.push({ x: W + 10, y: GROUND_Y - h, w: 44, h });
        }
        for (const o of s.obstacles) o.x -= s.speed;
        s.obstacles = s.obstacles.filter((o) => o.x > -60);

        // Collision (shrunk hitbox for fairness)
        for (const o of s.obstacles) {
          if (d.x + 8 < o.x + o.w - 4 && d.x + d.w - 10 > o.x + 4 && d.y + 8 < o.y + o.h && d.y + d.h - 4 > o.y) {
            if (s.score > s.hiScore) s.hiScore = s.score;
            s.phase = "over";
          }
        }
      } else {
        // Animate clouds even on idle/over
        for (const c of s.clouds) {
          c.x -= 0.5;
          if (c.x < -90) c.x = W + 50 + Math.random() * 120;
        }
      }

      // ---- DRAW ----
      ctx.fillStyle = "#f7f7f7";
      ctx.fillRect(0, 0, W, H);

      // Ground
      ctx.fillStyle = "#535353";
      ctx.fillRect(0, GROUND_Y, W, 3);

      // Ground texture dots
      ctx.fillStyle = "#bbb";
      const gOff = s.phase === "running" ? (s.frame * 2) % 60 : 0;
      for (let gx = (gOff % 60) - 60; gx < W; gx += 60) {
        ctx.fillRect(gx, GROUND_Y + 8, 14, 2);
        ctx.fillRect(gx + 34, GROUND_Y + 14, 7, 2);
      }

      // Clouds
      ctx.fillStyle = "#ddd";
      for (const cl of s.clouds) {
        ctx.fillRect(cl.x, cl.y, 72, 14);
        ctx.fillRect(cl.x + 14, cl.y - 10, 46, 14);
        ctx.fillRect(cl.x + 26, cl.y - 18, 24, 12);
      }

      for (const o of s.obstacles) drawCactus(o);
      drawDino(s.dino, s.frame);

      // Score
      ctx.fillStyle = "#535353";
      ctx.font = "bold 16px monospace";
      if (s.hiScore > 0) {
        ctx.textAlign = "right";
        ctx.fillStyle = "#bbb";
        ctx.fillText("HI " + String(Math.floor(s.hiScore)).padStart(5, "0"), W - 100, 30);
        ctx.fillStyle = "#535353";
      }
      ctx.textAlign = "right";
      ctx.fillText(String(Math.floor(s.score)).padStart(5, "0"), W - 20, 30);

      // Overlay
      ctx.textAlign = "center";
      if (s.phase === "idle") {
        ctx.fillStyle = "#535353";
        ctx.font = "bold 20px monospace";
        ctx.fillText("PRESS SPACE / TAP TO START", W / 2, H / 2 + 60);
      } else if (s.phase === "over") {
        ctx.fillStyle = "#535353";
        ctx.font = "bold 22px monospace";
        ctx.fillText("G A M E   O V E R", W / 2, H / 2 + 38);
        ctx.font = "14px monospace";
        ctx.fillText("Press SPACE or Tap to Restart", W / 2, H / 2 + 64);
      }

      animIdRef.current = requestAnimationFrame(loop);
    };

    const onKey = (e) => {
      if (e.code === "Space" || e.code === "ArrowUp") {
        e.preventDefault();
        jump();
      }
    };

    window.addEventListener("keydown", onKey);
    canvas.addEventListener("click", jump);
    canvas.addEventListener("touchstart", (e) => { e.preventDefault(); jump(); }, { passive: false });

    animIdRef.current = requestAnimationFrame(loop);

    return () => {
      cancelAnimationFrame(animIdRef.current);
      window.removeEventListener("keydown", onKey);
    };
  }, []);

  return (
    <div className="flex flex-col items-center gap-3">
      <canvas
        ref={canvasRef}
        width={800}
        height={300}
        className="border-2 border-[#535353] rounded-xl cursor-pointer max-w-full"
      />
      <p className="text-sm text-[#999]">SPACE / ↑ / Click to jump &nbsp;•&nbsp; Speed increases over time</p>
    </div>
  );
}
