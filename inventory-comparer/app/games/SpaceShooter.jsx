"use client";

import { useEffect, useRef } from "react";

const W = 800;
const H = 500;
const COLS = 8;
const ROWS = 4;
const EW = 36;
const EH = 26;
const EPAD_X = 14;
const EPAD_Y = 12;

const makeState = () => ({
  phase: "idle", // idle | running | over | win
  score: 0,
  lives: 3,
  player: { x: W / 2 - 20, y: H - 55, w: 40, h: 28, shootCooldown: 0 },
  bullets: [],
  eBullets: [],
  enemies: (() => {
    const arr = [];
    for (let r = 0; r < ROWS; r++)
      for (let c = 0; c < COLS; c++)
        arr.push({ x: 56 + c * (EW + EPAD_X), y: 50 + r * (EH + EPAD_Y), w: EW, h: EH, alive: true, type: r });
    return arr;
  })(),
  dir: 1,
  moveTimer: 0,
  shootTimer: 0,
  explosions: [],
  invincTimer: 0,
  frame: 0,
});

// Deterministic star positions
const STARS = Array.from({ length: 90 }, (_, i) => ({
  x: (i * 173 + 47) % W,
  y: (i * 97 + 23) % H,
  size: i % 5 === 0 ? 2 : 1,
  twinkle: i % 3 === 0,
}));

const ENEMY_COLORS = [
  ["#ff5252", "#ff8a80"],
  ["#ff9800", "#ffcc80"],
  ["#e040fb", "#ea80fc"],
  ["#29b6f6", "#81d4fa"],
];

export default function SpaceShooter() {
  const canvasRef = useRef(null);
  const animIdRef = useRef(null);
  const keysRef = useRef({});
  const gameRef = useRef(makeState());

  useEffect(() => {
    const canvas = canvasRef.current;
    const ctx = canvas.getContext("2d");

    const drawShip = (p, invincTimer) => {
      if (invincTimer > 0 && Math.floor(invincTimer / 5) % 2 === 0) return;
      ctx.fillStyle = "#00e676";
      ctx.beginPath();
      ctx.moveTo(p.x + p.w / 2, p.y);
      ctx.lineTo(p.x + p.w, p.y + p.h);
      ctx.lineTo(p.x + 6, p.y + p.h);
      ctx.closePath();
      ctx.fill();
      ctx.fillStyle = "#b9f6ca";
      ctx.fillRect(p.x + p.w / 2 - 3, p.y - 7, 6, 9);     // cannon
      ctx.fillStyle = "rgba(0,230,118,0.2)";
      ctx.fillRect(p.x + 4, p.y + p.h, p.w - 8, 8);        // engine glow
    };

    const drawEnemy = (e, frame) => {
      if (!e.alive) return;
      const [main, light] = ENEMY_COLORS[e.type % ENEMY_COLORS.length];
      ctx.fillStyle = main;
      const { x, y, w, h } = e;
      const anim = Math.floor(frame / 20) % 2;
      ctx.fillRect(x + 6, y + 4, w - 12, h - 8);       // body center
      ctx.fillRect(x + 2, y + 10, w - 4, h - 16);      // body wide
      ctx.fillRect(x + 4, y, 5, 7 + anim);              // left antenna
      ctx.fillRect(x + w - 9, y, 5, 7 + anim);         // right antenna
      ctx.fillRect(x + 4, y + h - 6, 7, 5 + anim);     // left leg
      ctx.fillRect(x + w - 11, y + h - 6, 7, 5 + anim);// right leg
      ctx.fillStyle = light;
      ctx.fillRect(x + 8, y + 8, 5, 5);                 // left eye
      ctx.fillRect(x + w - 13, y + 8, 5, 5);            // right eye
    };

    const loop = () => {
      const g = gameRef.current;

      if (g.phase === "running") {
        g.frame++;
        const p = g.player;

        // Player movement
        if (keysRef.current["ArrowLeft"]) p.x = Math.max(0, p.x - 5);
        if (keysRef.current["ArrowRight"]) p.x = Math.min(W - p.w, p.x + 5);

        // Player shoot
        if (keysRef.current["Space"] && p.shootCooldown <= 0) {
          g.bullets.push({ x: p.x + p.w / 2 - 2, y: p.y - 8, w: 4, h: 14 });
          p.shootCooldown = 18;
        }
        if (p.shootCooldown > 0) p.shootCooldown--;
        if (g.invincTimer > 0) g.invincTimer--;

        // Move bullets
        for (const b of g.bullets) b.y -= 8;
        g.bullets = g.bullets.filter((b) => b.y > -20);
        for (const b of g.eBullets) b.y += 4;
        g.eBullets = g.eBullets.filter((b) => b.y < H + 20);

        // Enemy movement
        const alive = g.enemies.filter((e) => e.alive);
        if (alive.length === 0) {
          g.phase = "win";
        } else {
          const speedBoost = Math.floor((COLS * ROWS - alive.length) / 3);
          g.moveTimer++;
          if (g.moveTimer >= Math.max(4, 28 - speedBoost * 3)) {
            g.moveTimer = 0;
            let hitWall = false;
            for (const e of alive) {
              if ((g.dir > 0 && e.x + e.w + 12 > W) || (g.dir < 0 && e.x - 12 < 0)) {
                hitWall = true;
                break;
              }
            }
            if (hitWall) {
              g.dir *= -1;
              for (const e of alive) e.y += 18;
            }
            for (const e of alive) e.x += g.dir * 10;
          }

          // Enemy shoot
          g.shootTimer++;
          if (g.shootTimer >= 50) {
            g.shootTimer = 0;
            const shooter = alive[Math.floor(Math.random() * alive.length)];
            g.eBullets.push({ x: shooter.x + shooter.w / 2 - 2, y: shooter.y + shooter.h, w: 4, h: 10 });
          }

          // Player bullet vs enemy
          for (const b of g.bullets) {
            for (const e of g.enemies) {
              if (!e.alive) continue;
              if (b.x < e.x + e.w && b.x + b.w > e.x && b.y < e.y + e.h && b.y + b.h > e.y) {
                e.alive = false;
                b.y = -200;
                g.score += (ROWS - e.type) * 10;
                g.explosions.push({ x: e.x + e.w / 2, y: e.y + e.h / 2, r: 0, maxR: 22, cr: 255, cg: 180, cb: 50 });
              }
            }
          }

          // Enemy bullet vs player
          if (g.invincTimer === 0) {
            for (const b of g.eBullets) {
              if (b.x < p.x + p.w && b.x + b.w > p.x && b.y < p.y + p.h && b.y + b.h > p.y) {
                b.y = H + 100;
                g.lives--;
                g.invincTimer = 90;
                g.explosions.push({ x: p.x + p.w / 2, y: p.y + p.h / 2, r: 0, maxR: 32, cr: 255, cg: 80, cb: 80 });
                if (g.lives <= 0) g.phase = "over";
              }
            }
          }

          // Enemies reach bottom
          for (const e of alive) {
            if (e.y + e.h >= g.player.y + 10) g.phase = "over";
          }
        }

        // Explosions
        for (const ex of g.explosions) ex.r += 1.6;
        g.explosions = g.explosions.filter((ex) => ex.r < ex.maxR);
      }

      // ---- DRAW ----
      ctx.fillStyle = "#080818";
      ctx.fillRect(0, 0, W, H);

      // Stars with twinkling
      for (const star of STARS) {
        const alpha = star.twinkle ? 0.4 + 0.3 * Math.sin(g.frame * 0.05 + star.x * 0.01) : 0.6;
        ctx.fillStyle = `rgba(255,255,255,${alpha.toFixed(2)})`;
        ctx.fillRect(star.x, star.y, star.size, star.size);
      }

      // Ground line
      ctx.fillStyle = "#00e676";
      ctx.fillRect(0, H - 18, W, 2);

      for (const e of g.enemies) drawEnemy(e, g.frame);
      drawShip(g.player, g.invincTimer);

      // Player bullets with glow
      for (const b of g.bullets) {
        ctx.fillStyle = "rgba(0,230,118,0.25)";
        ctx.fillRect(b.x - 3, b.y, b.w + 6, b.h);
        ctx.fillStyle = "#00e676";
        ctx.fillRect(b.x, b.y, b.w, b.h);
      }

      // Enemy bullets
      ctx.fillStyle = "#ff5252";
      for (const b of g.eBullets) ctx.fillRect(b.x, b.y, b.w, b.h);

      // Explosions
      for (const ex of g.explosions) {
        const alpha = (1 - ex.r / ex.maxR).toFixed(2);
        ctx.fillStyle = `rgba(${ex.cr},${ex.cg},${ex.cb},${alpha})`;
        ctx.beginPath();
        ctx.arc(ex.x, ex.y, ex.r, 0, Math.PI * 2);
        ctx.fill();
      }

      // HUD
      ctx.font = "bold 16px monospace";
      ctx.fillStyle = "#fff";
      ctx.textAlign = "left";
      ctx.fillText(`SCORE: ${g.score}`, 16, 26);
      ctx.textAlign = "right";
      ctx.fillStyle = "#ff5252";
      ctx.fillText(Array.from({ length: Math.max(0, g.lives) }, () => "♥").join(" "), W - 16, 26);

      // Score legend (top right of enemy grid)
      ctx.font = "11px monospace";
      ctx.fillStyle = "#666";
      ctx.textAlign = "left";
      ctx.fillText("= 40 pts", 66, 20);
      ctx.fillStyle = ENEMY_COLORS[0][0];
      ctx.fillRect(56, 10, 8, 8);
      ctx.fillStyle = "#666";
      ctx.fillText("= 30 pts", 66, 34);
      ctx.fillStyle = ENEMY_COLORS[1][0];
      ctx.fillRect(56, 24, 8, 8);

      // Phase overlays
      ctx.textAlign = "center";
      if (g.phase === "idle") {
        ctx.fillStyle = "rgba(0,0,0,0.55)";
        ctx.fillRect(0, 0, W, H);
        ctx.fillStyle = "#00e676";
        ctx.font = "bold 34px monospace";
        ctx.fillText("SPACE INVADERS", W / 2, H / 2 - 44);
        ctx.fillStyle = "#aaa";
        ctx.font = "15px monospace";
        ctx.fillText("← → to move   •   SPACE to shoot", W / 2, H / 2 + 2);
        ctx.fillStyle = "#00e676";
        ctx.font = "bold 18px monospace";
        ctx.fillText("▶  Press SPACE to Start  ◀", W / 2, H / 2 + 48);
      } else if (g.phase === "over") {
        ctx.fillStyle = "rgba(0,0,0,0.65)";
        ctx.fillRect(0, 0, W, H);
        ctx.fillStyle = "#ff5252";
        ctx.font = "bold 38px monospace";
        ctx.fillText("GAME OVER", W / 2, H / 2 - 24);
        ctx.fillStyle = "#fff";
        ctx.font = "20px monospace";
        ctx.fillText(`SCORE: ${g.score}`, W / 2, H / 2 + 18);
        ctx.fillStyle = "#aaa";
        ctx.font = "14px monospace";
        ctx.fillText("Press SPACE to Play Again", W / 2, H / 2 + 58);
      } else if (g.phase === "win") {
        ctx.fillStyle = "rgba(0,0,0,0.65)";
        ctx.fillRect(0, 0, W, H);
        ctx.fillStyle = "#00e676";
        ctx.font = "bold 38px monospace";
        ctx.fillText("YOU WIN! 🎉", W / 2, H / 2 - 24);
        ctx.fillStyle = "#fff";
        ctx.font = "20px monospace";
        ctx.fillText(`SCORE: ${g.score}`, W / 2, H / 2 + 18);
        ctx.fillStyle = "#aaa";
        ctx.font = "14px monospace";
        ctx.fillText("Press SPACE to Play Again", W / 2, H / 2 + 58);
      }

      animIdRef.current = requestAnimationFrame(loop);
    };

    const onKeyDown = (e) => {
      keysRef.current[e.code] = true;
      if (["Space", "ArrowLeft", "ArrowRight", "ArrowUp", "ArrowDown"].includes(e.code)) {
        e.preventDefault();
      }
      if (e.code === "Space") {
        const g = gameRef.current;
        if (g.phase === "idle") {
          g.phase = "running";
        } else if (g.phase === "over" || g.phase === "win") {
          gameRef.current = { ...makeState(), phase: "running" };
          keysRef.current["Space"] = false; // prevent instant shot on restart
        }
      }
    };

    const onKeyUp = (e) => { keysRef.current[e.code] = false; };

    window.addEventListener("keydown", onKeyDown);
    window.addEventListener("keyup", onKeyUp);
    animIdRef.current = requestAnimationFrame(loop);

    return () => {
      cancelAnimationFrame(animIdRef.current);
      window.removeEventListener("keydown", onKeyDown);
      window.removeEventListener("keyup", onKeyUp);
    };
  }, []);

  return (
    <div className="flex flex-col items-center gap-3">
      <canvas
        ref={canvasRef}
        width={800}
        height={500}
        className="border-2 border-[#00e676] rounded-xl max-w-full"
      />
      <p className="text-sm text-[#999]">← → to move &nbsp;•&nbsp; SPACE to shoot &nbsp;•&nbsp; Enemies speed up as you kill them</p>
    </div>
  );
}
