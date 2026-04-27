"use client";

import { useEffect, useRef } from "react";

const W = 800;
const H = 450;
const FOV = Math.PI / 3;
const HALF_FOV = FOV / 2;
const NUM_RAYS = 400;
const MAX_DEPTH = 30;
const PLAYER_SPEED = 0.038;
const ROT_SPEED = 0.024;
const MOUSE_SENS = 0.0022;
const MAP_SCALE = 9;

// prettier-ignore
const MAP_DATA = [
  [1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1],
  [1,0,0,0,0,0,0,0,0,0,0,0,0,1,0,0,0,0,0,0,0,0,0,0,0,0,0,1],
  [1,0,0,0,2,2,0,0,0,2,2,2,0,1,0,0,2,2,0,0,0,0,3,3,0,0,0,1],
  [1,0,0,0,2,0,0,0,0,0,0,2,0,0,0,0,2,0,0,0,0,0,0,3,0,0,0,1],
  [1,0,2,0,2,0,0,3,0,0,0,0,0,0,0,0,2,0,0,0,0,0,0,0,0,0,0,1],
  [1,0,2,0,0,0,0,3,0,0,0,0,0,1,0,0,0,0,0,3,0,3,0,0,0,0,0,1],
  [1,0,2,2,0,0,0,0,0,0,0,0,0,1,0,0,0,0,0,3,0,3,0,0,2,2,0,1],
  [1,0,0,0,0,0,0,0,0,0,0,0,0,1,1,0,0,0,0,0,0,0,0,0,2,0,0,1],
  [1,0,0,0,0,3,3,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,1],
  [1,0,0,0,0,3,0,0,0,0,2,2,2,0,0,0,1,1,0,0,0,3,0,0,0,0,0,1],
  [1,0,0,0,0,0,0,0,0,0,2,0,0,0,0,0,1,0,0,0,0,3,0,0,0,0,0,1],
  [1,0,0,2,2,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,1],
  [1,0,0,2,0,0,0,0,0,0,0,0,0,0,0,0,0,0,2,2,0,0,0,0,0,0,0,1],
  [1,1,0,0,0,0,0,0,3,3,0,0,0,0,0,0,0,0,2,0,0,0,0,0,0,0,0,1],
  [1,0,0,0,0,0,0,0,3,0,0,0,0,0,0,3,3,0,0,0,0,0,0,2,2,2,0,1],
  [1,0,0,0,0,0,0,0,0,0,0,0,0,0,0,3,0,0,0,0,0,0,0,2,0,2,0,1],
  [1,0,2,2,2,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,1],
  [1,0,2,0,0,0,0,0,0,0,1,1,1,0,0,0,0,0,0,0,0,0,0,0,0,3,0,1],
  [1,0,0,0,0,0,3,0,0,0,1,0,0,0,0,0,0,0,2,2,0,0,0,0,0,3,0,1],
  [1,0,0,0,0,0,3,0,0,0,0,0,0,0,0,0,0,0,2,0,0,0,0,0,0,0,0,1],
  [1,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,1],
  [1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1],
];
const MAP_H = MAP_DATA.length;
const MAP_W = MAP_DATA[0].length;

const WALL_COLORS = {
  1: { h: ["#b03040", "#7a1f2a"], v: ["#d04050", "#a03040"] },
  2: { h: ["#3060b0", "#1a3a6a"], v: ["#4070c0", "#2050a0"] },
  3: { h: ["#308050", "#1a5030"], v: ["#40a060", "#2a7040"] },
};

// Valid open-floor spawn positions [x, y] spread across the new larger map
const SPAWN_POOL = [
  // top half
  [2.5,1.5],[6.5,1.5],[10.5,1.5],[15.5,1.5],[19.5,1.5],[23.5,1.5],[26.5,1.5],
  [2.5,3.5],[6.5,3.5],[9.5,3.5],[15.5,3.5],[20.5,3.5],[25.5,3.5],
  [5.5,5.5],[8.5,5.5],[15.5,5.5],[20.5,5.5],[22.5,5.5],[25.5,5.5],
  [2.5,7.5],[7.5,7.5],[12.5,7.5],[18.5,7.5],[22.5,7.5],[26.5,7.5],
  [2.5,9.5],[7.5,9.5],[12.5,9.5],[18.5,9.5],[22.5,9.5],[26.5,9.5],
  // bottom half
  [5.5,11.5],[8.5,11.5],[14.5,11.5],[19.5,11.5],[23.5,11.5],[26.5,11.5],
  [2.5,13.5],[6.5,13.5],[12.5,13.5],[17.5,13.5],[21.5,13.5],[26.5,13.5],
  [3.5,15.5],[7.5,15.5],[11.5,15.5],[18.5,15.5],[22.5,15.5],[25.5,15.5],
  [3.5,17.5],[7.5,17.5],[13.5,17.5],[18.5,17.5],[22.5,17.5],[24.5,17.5],
  [3.5,19.5],[7.5,19.5],[11.5,19.5],[16.5,19.5],[21.5,19.5],[26.5,19.5],
  [3.5,20.5],[10.5,20.5],[15.5,20.5],[20.5,20.5],[25.5,20.5],
];

function spawnWave(g) {
  g.wave++;
  const count = Math.min(3 + g.wave * 2, 16);
  const hp = 2 + Math.floor(g.wave / 2);
  const shootDelay = Math.max(22, 70 - g.wave * 4);
  const moveSpeed = 0.022 + g.wave * 0.003;
  // Shuffle pool and take count positions, all far enough from player
  const pool = [...SPAWN_POOL]
    .filter(([x, y]) => Math.hypot(x - g.x, y - g.y) > 4)
    .sort(() => Math.random() - 0.5);
  g.enemies = pool.slice(0, count).map(([x, y]) => ({
    x, y, health: hp, maxHealth: hp, alive: true,
    shootTimer: shootDelay + Math.random() * 40,
    alert: false, speed: moveSpeed,
  }));
  // Reset health each wave and reduce incoming damage scaling
  g.health = 100;
  g.dmgReduction = Math.min(0.5, (g.wave - 1) * 0.04); // 4% less damage per wave, cap at 50%
  // Top up ammo each wave
  g.ammo = Math.min(g.ammo + 15 + g.wave * 3, 60);
  g.waveFlash = 160;
}

function makeState() {
  return {
    x: 2.5, y: 2.5, angle: 0.3,
    phase: "idle",
    score: 0,
    health: 100,
    ammo: 30,
    enemyKills: 0,
    wave: 0,
    waveFlash: 0,
    shootCooldown: 0,
    shootFlash: 0,
    bobTime: 0,
    enemies: [],
    hitFlash: 0,
  };
}

function raycast(px, py, angle) {
  const rays = [];
  for (let i = 0; i < NUM_RAYS; i++) {
    const rayAngle = angle - HALF_FOV + (i / NUM_RAYS) * FOV;
    const cos = Math.cos(rayAngle);
    const sin = Math.sin(rayAngle);
    let dist = 0;
    let wallType = 0;
    let side = 0;
    for (let d = 0; d < MAX_DEPTH; d += 0.05) {
      const tx = Math.floor(px + cos * d);
      const ty = Math.floor(py + sin * d);
      if (tx < 0 || tx >= MAP_W || ty < 0 || ty >= MAP_H) { dist = MAX_DEPTH; break; }
      const cell = MAP_DATA[ty][tx];
      if (cell > 0) {
        dist = d;
        wallType = cell;
        // Determine horizontal vs vertical hit for shading
        const px2 = px + cos * (d - 0.05);
        const py2 = py + sin * (d - 0.05);
        side = Math.floor(px2) !== tx ? 1 : 0;
        break;
      }
    }
    const corrected = dist * Math.cos(rayAngle - angle); // fish-eye fix
    rays.push({ dist: corrected, wallType, side });
  }
  return rays;
}

export default function FPSGame() {
  const canvasRef = useRef(null);
  const stateRef = useRef(makeState());
  const keysRef = useRef({});
  const animRef = useRef(null);
  const lockedRef = useRef(false);

  useEffect(() => {
    const canvas = canvasRef.current;
    const ctx = canvas.getContext("2d");

    const isSolid = (x, y) => {
      const tx = Math.floor(x); const ty = Math.floor(y);
      if (tx < 0 || tx >= MAP_W || ty < 0 || ty >= MAP_H) return true;
      return MAP_DATA[ty][tx] > 0;
    };

    const shoot = () => {
      const g = stateRef.current;
      if (g.phase !== "running" || g.ammo <= 0 || g.shootCooldown > 0) return;
      g.ammo--;
      g.shootCooldown = 22;
      g.shootFlash = 10;
      // Hit check: find closest enemy in crosshair that isn't wall-occluded
      for (const e of g.enemies) {
        if (!e.alive) continue;
        const dx = e.x - g.x; const dy = e.y - g.y;
        const dist = Math.sqrt(dx * dx + dy * dy);
        if (dist > 18) continue;
        const angleToEnemy = Math.atan2(dy, dx);
        let diff = angleToEnemy - g.angle;
        while (diff > Math.PI) diff -= Math.PI * 2;
        while (diff < -Math.PI) diff += Math.PI * 2;
        if (Math.abs(diff) < 0.12 + 0.05 / dist) {
          // Wall-occlusion check: step ray toward enemy
          const steps = Math.ceil(dist / 0.1);
          let blocked = false;
          for (let s = 1; s < steps; s++) {
            const fx = g.x + (dx / dist) * s * 0.1;
            const fy = g.y + (dy / dist) * s * 0.1;
            if (isSolid(fx, fy)) { blocked = true; break; }
          }
          if (blocked) continue;
          e.health--;
          if (e.health <= 0) { e.alive = false; g.enemyKills++; g.score += 100; }
          break;
        }
      }
    };

    const loop = () => {
      const g = stateRef.current;

      if (g.phase === "running") {
        const cos = Math.cos(g.angle);
        const sin = Math.sin(g.angle);
        let nx = g.x, ny = g.y;
        if (keysRef.current["ArrowUp"] || keysRef.current["KeyW"]) { nx += cos * PLAYER_SPEED; ny += sin * PLAYER_SPEED; }
        if (keysRef.current["ArrowDown"] || keysRef.current["KeyS"]) { nx -= cos * PLAYER_SPEED; ny -= sin * PLAYER_SPEED; }
        if (keysRef.current["ArrowLeft"] || keysRef.current["KeyA"]) g.angle -= ROT_SPEED;
        if (keysRef.current["ArrowRight"] || keysRef.current["KeyD"]) g.angle += ROT_SPEED;
        if (!isSolid(nx, g.y)) g.x = nx;
        if (!isSolid(g.x, ny)) g.y = ny;
        if (g.shootCooldown > 0) g.shootCooldown--;
        if (g.shootFlash > 0) g.shootFlash--;
        if (g.hitFlash > 0) g.hitFlash--;
        g.bobTime += 0.1;
        g.score += 0.01;

        // Enemy AI
        for (const e of g.enemies) {
          if (!e.alive) continue;
          const dx = g.x - e.x; const dy = g.y - e.y;
          const dist = Math.sqrt(dx * dx + dy * dy);
          if (dist < 8) e.alert = true;
          if (e.alert) {
            const moveStep = e.speed || 0.025;
            const nx2 = e.x + (dx / dist) * moveStep;
            const ny2 = e.y + (dy / dist) * moveStep;
            if (!isSolid(nx2, e.y)) e.x = nx2;
            if (!isSolid(e.x, ny2)) e.y = ny2;
            e.shootTimer--;
            if (e.shootTimer <= 0 && dist < 6) {
              e.shootTimer = Math.max(22, 70 - g.wave * 4) + Math.random() * 40;
              // Line-of-sight: don't shoot through walls
              const steps = Math.ceil(dist / 0.1);
              let wallBlocked = false;
              for (let s = 1; s < steps; s++) {
                const fx = e.x + (dx / dist) * s * 0.1;
                const fy = e.y + (dy / dist) * s * 0.1;
                if (isSolid(fx, fy)) { wallBlocked = true; break; }
              }
              if (!wallBlocked) {
                const dmg = Math.round(8 * (1 - (g.dmgReduction || 0)));
                g.health -= dmg;
                g.hitFlash = 20;
                if (g.health <= 0) { g.health = 0; g.phase = "over"; }
              }
            }
          }
        }
        // Wave complete — spawn next wave after brief flash
        if (g.waveFlash > 0) {
          g.waveFlash--;
        } else if (g.enemies.length > 0 && g.enemies.every((e) => !e.alive)) {
          spawnWave(g);
        }
      }

      // ---- DRAW ----
      // Sky gradient
      const sky = ctx.createLinearGradient(0, 0, 0, H / 2);
      sky.addColorStop(0, "#0a0a1a");
      sky.addColorStop(1, "#1a1a3a");
      ctx.fillStyle = sky;
      ctx.fillRect(0, 0, W, H / 2);

      // Floor gradient
      const floor = ctx.createLinearGradient(0, H / 2, 0, H);
      floor.addColorStop(0, "#1a1008");
      floor.addColorStop(1, "#0a0804");
      ctx.fillStyle = floor;
      ctx.fillRect(0, H / 2, W, H / 2);

      // Raycasted walls
      const rays = raycast(stateRef.current.x, stateRef.current.y, stateRef.current.angle);
      const sliceW = W / NUM_RAYS;
      for (let i = 0; i < NUM_RAYS; i++) {
        const { dist, wallType, side } = rays[i];
        const wc = WALL_COLORS[wallType] || WALL_COLORS[1];
        const colors = side === 1 ? wc.v : wc.h;
        const sliceH = Math.min((H / (dist || 0.1)) * 1.2, H);
        const top = (H - sliceH) / 2;
        const darkFactor = Math.max(0, 1 - dist / MAX_DEPTH);
        const col = dist < 3 ? colors[0] : (dist < 7 ? colors[1] : "#111");
        ctx.fillStyle = col;
        ctx.globalAlpha = darkFactor * 0.9 + 0.1;
        ctx.fillRect(i * sliceW, top, sliceW + 1, sliceH);
      }
      ctx.globalAlpha = 1;

      // Enemy sprites (simple billboard)
      for (const e of g.enemies) {
        if (!e.alive) continue;
        const dx = e.x - g.x; const dy = e.y - g.y;
        const dist = Math.sqrt(dx * dx + dy * dy);
        let spriteAngle = Math.atan2(dy, dx) - g.angle;
        while (spriteAngle > Math.PI) spriteAngle -= Math.PI * 2;
        while (spriteAngle < -Math.PI) spriteAngle += Math.PI * 2;
        if (Math.abs(spriteAngle) > HALF_FOV + 0.1) continue;
        const screenX = (0.5 + spriteAngle / FOV) * W;
        const size = Math.min((H / (dist || 0.1)) * 0.9, H);
        const sx = screenX - size / 2;
        const sy = (H - size) / 2;
        // Body
        ctx.fillStyle = e.health > 2 ? "#cc3333" : (e.health > 1 ? "#cc6633" : "#996633");
        ctx.globalAlpha = Math.max(0, 1 - dist / 12);
        ctx.fillRect(sx + size * 0.25, sy + size * 0.15, size * 0.5, size * 0.55);
        // Head
        ctx.fillStyle = "#f5c58a";
        ctx.fillRect(sx + size * 0.32, sy, size * 0.36, size * 0.22);
        // Eyes
        ctx.fillStyle = "#200000";
        ctx.fillRect(sx + size * 0.36, sy + size * 0.06, size * 0.08, size * 0.08);
        ctx.fillRect(sx + size * 0.56, sy + size * 0.06, size * 0.08, size * 0.08);
        // Health bar
        ctx.fillStyle = "#300";
        ctx.fillRect(sx + size * 0.2, sy - size * 0.08, size * 0.6, size * 0.06);
        ctx.fillStyle = "#0f0";
        ctx.fillRect(sx + size * 0.2, sy - size * 0.08, size * 0.6 * (e.health / 3), size * 0.06);
        ctx.globalAlpha = 1;
      }

      // Hit flash
      if (g.hitFlash > 0) {
        ctx.fillStyle = `rgba(255,0,0,${(g.hitFlash / 20) * 0.35})`;
        ctx.fillRect(0, 0, W, H);
      }

      // Muzzle flash
      if (g.shootFlash > 0) {
        const grd = ctx.createRadialGradient(W / 2, H / 2 + 60, 0, W / 2, H / 2 + 60, 55);
        grd.addColorStop(0, `rgba(255,220,80,${(g.shootFlash / 10) * 0.7})`);
        grd.addColorStop(1, "rgba(255,100,0,0)");
        ctx.fillStyle = grd;
        ctx.fillRect(W / 2 - 60, H / 2, 120, 120);
      }

      // Gun sprite
      const bobY = g.phase === "running" ? Math.sin(g.bobTime * 2) * 4 : 0;
      ctx.fillStyle = "#555";
      ctx.fillRect(W / 2 + 40, H - 95 + bobY, 60, 80);
      ctx.fillStyle = "#444";
      ctx.fillRect(W / 2 + 44, H - 112 + bobY, 18, 22);
      ctx.fillStyle = "#333";
      ctx.fillRect(W / 2 + 44, H - 116 + bobY, 14, 8);

      // Crosshair
      ctx.strokeStyle = "rgba(255,255,255,0.75)";
      ctx.lineWidth = 1.5;
      ctx.beginPath();
      ctx.moveTo(W / 2 - 12, H / 2); ctx.lineTo(W / 2 + 12, H / 2);
      ctx.moveTo(W / 2, H / 2 - 12); ctx.lineTo(W / 2, H / 2 + 12);
      ctx.stroke();
      ctx.beginPath();
      ctx.arc(W / 2, H / 2, 4, 0, Math.PI * 2);
      ctx.stroke();

      // HUD
      ctx.font = "bold 14px monospace";
      ctx.fillStyle = "#0f0";
      ctx.textAlign = "left";
      ctx.fillText(`❤ ${Math.max(0, g.health)}%  🔫 ${g.ammo}  💀 ${g.enemyKills}  W${g.wave}`, 14, 28);
      ctx.textAlign = "right";
      ctx.fillStyle = "#ff0";
      ctx.fillText(`SCORE ${Math.floor(g.score)}`, W - 14, 28);

      // Wave banner
      if (g.waveFlash > 60) {
        const t = Math.min(1, (g.waveFlash - 60) / 25);
        ctx.globalAlpha = t;
        ctx.fillStyle = "rgba(0,0,0,0.65)";
        ctx.fillRect(W / 2 - 210, H / 2 - 46, 420, 82);
        ctx.fillStyle = "#ffe000";
        ctx.font = "bold 34px monospace";
        ctx.textAlign = "center";
        ctx.fillText(`◆  WAVE ${g.wave}  ◆`, W / 2, H / 2 - 6);
        ctx.fillStyle = "#ccc";
        ctx.font = "13px monospace";
        ctx.fillText(`${g.enemies.length} enemies incoming`, W / 2, H / 2 + 22);
        ctx.globalAlpha = 1;
      }

      // Minimap
      const mox = 14, moy = H - MAP_H * MAP_SCALE - 14;
      ctx.fillStyle = "rgba(0,0,0,0.55)";
      ctx.fillRect(mox - 2, moy - 2, MAP_W * MAP_SCALE + 4, MAP_H * MAP_SCALE + 4);
      for (let r = 0; r < MAP_H; r++) {
        for (let c = 0; c < MAP_W; c++) {
          const cell = MAP_DATA[r][c];
          ctx.fillStyle = cell > 0 ? WALL_COLORS[cell]?.h[0] ?? "#888" : "#111";
          ctx.fillRect(mox + c * MAP_SCALE, moy + r * MAP_SCALE, MAP_SCALE - 1, MAP_SCALE - 1);
        }
      }
      for (const e of g.enemies) {
        if (!e.alive) continue;
        ctx.fillStyle = "#f44";
        ctx.fillRect(mox + e.x * MAP_SCALE - 3, moy + e.y * MAP_SCALE - 3, 6, 6);
      }
      ctx.fillStyle = "#0f0";
      ctx.beginPath();
      ctx.arc(mox + g.x * MAP_SCALE, moy + g.y * MAP_SCALE, 4, 0, Math.PI * 2);
      ctx.fill();
      ctx.strokeStyle = "#0f0";
      ctx.lineWidth = 1.5;
      ctx.beginPath();
      ctx.moveTo(mox + g.x * MAP_SCALE, moy + g.y * MAP_SCALE);
      ctx.lineTo(mox + g.x * MAP_SCALE + Math.cos(g.angle) * 10, moy + g.y * MAP_SCALE + Math.sin(g.angle) * 10);
      ctx.stroke();

      // Phase overlays
      ctx.textAlign = "center";
      if (g.phase === "idle") {
        ctx.fillStyle = "rgba(0,0,0,0.7)";
        ctx.fillRect(0, 0, W, H);
        ctx.fillStyle = "#0f0";
        ctx.font = "bold 32px monospace";
        ctx.fillText("DUNGEON FPS", W / 2, H / 2 - 50);
        ctx.fillStyle = "#aaa";
        ctx.font = "13px monospace";
        ctx.fillText("WASD / ↑↓←→ — Move    Mouse / A/D — Turn    Click / SPACE — Shoot", W / 2, H / 2 - 8);
        ctx.fillText("Survive endless waves. Click canvas to lock mouse.", W / 2, H / 2 + 16);
        ctx.fillStyle = "#0f0";
        ctx.font = "bold 18px monospace";
        ctx.fillText("▶  Press SPACE to Start  ◀", W / 2, H / 2 + 58);
      } else if (g.phase === "over") {
        ctx.fillStyle = "rgba(0,0,0,0.75)";
        ctx.fillRect(0, 0, W, H);
        ctx.fillStyle = "#f44";
        ctx.font = "bold 44px monospace";
        ctx.fillText("YOU DIED", W / 2, H / 2 - 20);
        ctx.fillStyle = "#fff";
        ctx.font = "18px monospace";
        ctx.fillText(`Wave: ${g.wave}  •  Score: ${Math.floor(g.score)}  •  Kills: ${g.enemyKills}`, W / 2, H / 2 + 24);
        ctx.fillStyle = "#aaa";
        ctx.font = "14px monospace";
        ctx.fillText("Press SPACE to Restart", W / 2, H / 2 + 60);
      }

      animRef.current = requestAnimationFrame(loop);
    };

    const onKeyDown = (e) => {
      keysRef.current[e.code] = true;
      if (["Space", "ArrowUp", "ArrowDown", "ArrowLeft", "ArrowRight"].includes(e.code)) e.preventDefault();
      const g = stateRef.current;
      if (e.code === "Space") {
        if (g.phase === "idle") { g.phase = "running"; spawnWave(g); return; }
        if (g.phase === "over") { const ns = makeState(); ns.phase = "running"; spawnWave(ns); stateRef.current = ns; return; }
        shoot();
      }
    };
    const onKeyUp = (e) => { keysRef.current[e.code] = false; };

    // Pointer lock mouse look
    const onMouseMove = (e) => {
      if (!lockedRef.current) return;
      const g = stateRef.current;
      if (g.phase === "running") g.angle += e.movementX * MOUSE_SENS;
    };
    const onMouseDown = (e) => {
      if (e.button !== 0) return;
      if (!lockedRef.current) {
        canvas.requestPointerLock();
        return;
      }
      shoot();
    };
    const onLockChange = () => {
      lockedRef.current = document.pointerLockElement === canvas;
    };

    window.addEventListener("keydown", onKeyDown);
    window.addEventListener("keyup", onKeyUp);
    document.addEventListener("mousemove", onMouseMove);
    canvas.addEventListener("mousedown", onMouseDown);
    document.addEventListener("pointerlockchange", onLockChange);
    animRef.current = requestAnimationFrame(loop);

    return () => {
      cancelAnimationFrame(animRef.current);
      window.removeEventListener("keydown", onKeyDown);
      window.removeEventListener("keyup", onKeyUp);
      document.removeEventListener("mousemove", onMouseMove);
      canvas.removeEventListener("mousedown", onMouseDown);
      document.removeEventListener("pointerlockchange", onLockChange);
      if (document.pointerLockElement === canvas) document.exitPointerLock();
    };
  }, []);

  return (
    <div className="flex flex-col items-center gap-3">
      <canvas ref={canvasRef} width={W} height={H}
        className="border-2 border-[#0f0] rounded-xl max-w-full cursor-crosshair"
      />
      <p className="text-sm text-[#999]">WASD Move &nbsp;•&nbsp; Mouse / A&amp;D Turn &nbsp;•&nbsp; Click / SPACE Shoot &nbsp;•&nbsp; Click canvas to lock mouse</p>
    </div>
  );
}
