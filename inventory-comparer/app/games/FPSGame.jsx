"use client";

import { useEffect, useRef, useState } from "react";

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

// Daytime wall colours — warm stone/brick tones lit by sunlight
const WALL_COLORS = {
  1: { h: ["#c8a46e", "#a07848"], v: ["#e0bc88", "#b89060"] },
  2: { h: ["#7094b8", "#4a6888"], v: ["#88acd0", "#5a7ea0"] },
  3: { h: ["#60a060", "#3a7840"], v: ["#78b878", "#509858"] },
};

// GUNS
const GUNS = {
  pistol: { name: "Pistol",     cost: 0,   dmg: 1, cooldown: 22, bulletsPerShot: 1, auto: false, spread: 0.00 },
  semi:   { name: "Semi-Auto",  cost: 150, dmg: 1, cooldown: 10, bulletsPerShot: 3, auto: false, spread: 0.06 },
  rifle:  { name: "Auto Rifle", cost: 350, dmg: 2, cooldown: 6,  bulletsPerShot: 1, auto: true,  spread: 0.03 },
  laser:  { name: "Laser Gun",  cost: 600, dmg: 3, cooldown: 3,  bulletsPerShot: 1, auto: true,  spread: 0.00 },
};

// Difficulty
const DIFFICULTY = {
  easy:   { label: "Easy",   enemyHpMult: 0.5, enemyDmgMult: 0.5, enemySpeedMult: 0.7, killMoney: 40 },
  normal: { label: "Normal", enemyHpMult: 1.0, enemyDmgMult: 1.0, enemySpeedMult: 1.0, killMoney: 25 },
  hard:   { label: "Hard",   enemyHpMult: 1.8, enemyDmgMult: 1.5, enemySpeedMult: 1.3, killMoney: 15 },
};

const BOSS_EVERY = 5;

// Valid open-floor spawn positions
const SPAWN_POOL = [
  [2.5,1.5],[6.5,1.5],[10.5,1.5],[15.5,1.5],[19.5,1.5],[23.5,1.5],[26.5,1.5],
  [2.5,3.5],[6.5,3.5],[9.5,3.5],[15.5,3.5],[20.5,3.5],[25.5,3.5],
  [5.5,5.5],[8.5,5.5],[15.5,5.5],[20.5,5.5],[22.5,5.5],[25.5,5.5],
  [2.5,7.5],[7.5,7.5],[12.5,7.5],[18.5,7.5],[22.5,7.5],[26.5,7.5],
  [2.5,9.5],[7.5,9.5],[12.5,9.5],[18.5,9.5],[22.5,9.5],[26.5,9.5],
  [5.5,11.5],[8.5,11.5],[14.5,11.5],[19.5,11.5],[23.5,11.5],[26.5,11.5],
  [2.5,13.5],[6.5,13.5],[12.5,13.5],[17.5,13.5],[21.5,13.5],[26.5,13.5],
  [3.5,15.5],[7.5,15.5],[11.5,15.5],[18.5,15.5],[22.5,15.5],[25.5,15.5],
  [3.5,17.5],[7.5,17.5],[13.5,17.5],[18.5,17.5],[22.5,17.5],[24.5,17.5],
  [3.5,19.5],[7.5,19.5],[11.5,19.5],[16.5,19.5],[21.5,19.5],[26.5,19.5],
  [3.5,20.5],[10.5,20.5],[15.5,20.5],[20.5,20.5],[25.5,20.5],
];

function spawnWave(g, setShopStateFn) {
  g.wave++;
  const diff = DIFFICULTY[g.difficulty];

  if (g.wave % BOSS_EVERY === 0) {
    g.phase = "boss";
    const bossHp = Math.floor(80 * diff.enemyHpMult * (1 + g.wave * 0.2));
    g.boss = {
      x: 13.5, y: 10.5,
      health: bossHp, maxHealth: bossHp,
      alive: true,
      fireballTimer: 120,
      moveTimer: 0,
      fireballs: [],
      speed: 0.012 * diff.enemySpeedMult,
    };
    g.enemies = [];
    g.waveFlash = 200;
    g.shopOpen = false;
    return;
  }

  // open shop before starting the wave
  g.phase = "shop";
  g.shopOpen = true;
}

function startNormalWave(g) {
  const diff = DIFFICULTY[g.difficulty];
  const count = Math.min(3 + g.wave * 2, 16);
  const baseHp = Math.ceil((2 + Math.floor(g.wave / 2)) * diff.enemyHpMult);
  const shootDelay = Math.max(22, 70 - g.wave * 4);
  const moveSpeed = (0.022 + g.wave * 0.003) * diff.enemySpeedMult;

  const pool = [...SPAWN_POOL]
    .filter(([x, y]) => Math.hypot(x - g.x, y - g.y) > 4)
    .sort(() => Math.random() - 0.5);

  g.enemies = pool.slice(0, count).map(([x, y]) => ({
    x, y, health: baseHp, maxHealth: baseHp, alive: true,
    shootTimer: shootDelay + Math.random() * 40,
    alert: false, speed: moveSpeed,
  }));

  g.health = 100;
  g.shopOpen = false;
  g.phase = "running";
  g.waveFlash = 160;
}

function makeState(difficulty = "normal") {
  return {
    x: 2.5, y: 2.5, angle: 0.3,
    phase: "idle",
    difficulty,
    score: 0,
    money: 0,
    health: 100,
    enemyKills: 0,
    wave: 0,
    waveFlash: 0,
    shootCooldown: 0,
    shootFlash: 0,
    bobTime: 0,
    enemies: [],
    hitFlash: 0,
    gun: "pistol",
    boss: null,
    shopOpen: false,
    moneyPopups: [],
    frameCount: 0,
  };
}

function raycast(px, py, angle) {
  const rays = [];
  for (let i = 0; i < NUM_RAYS; i++) {
    const rayAngle = angle - HALF_FOV + (i / NUM_RAYS) * FOV;
    const cos = Math.cos(rayAngle);
    const sin = Math.sin(rayAngle);
    let dist = 0; let wallType = 0; let side = 0;
    for (let d = 0; d < MAX_DEPTH; d += 0.05) {
      const tx = Math.floor(px + cos * d);
      const ty = Math.floor(py + sin * d);
      if (tx < 0 || tx >= MAP_W || ty < 0 || ty >= MAP_H) { dist = MAX_DEPTH; break; }
      const cell = MAP_DATA[ty][tx];
      if (cell > 0) {
        dist = d; wallType = cell;
        const px2 = px + cos * (d - 0.05);
        side = Math.floor(px2) !== tx ? 1 : 0;
        break;
      }
    }
    rays.push({ dist: dist * Math.cos(rayAngle - angle), wallType, side });
  }
  return rays;
}

export default function FPSGame() {
  const canvasRef = useRef(null);
  const stateRef = useRef(null);
  const keysRef = useRef({});
  const animRef = useRef(null);
  const lockedRef = useRef(false);
  const mouseDownRef = useRef(false);

  const [uiPhase, setUiPhase] = useState("diffpick");
  const [shopTick, setShopTick] = useState(0); // force shop re-render
  const [shopCursor, setShopCursor] = useState(0); // 0-3 = gun cards, 4 = Start Wave button
  const shopCursorRef = useRef(0);

  const setShopCursorBoth = (v) => { shopCursorRef.current = v; setShopCursor(v); };

  const startGame = (diff) => {
    stateRef.current = makeState(diff);
    setUiPhase("game");
  };

  const openShop = () => { setShopCursorBoth(0); setShopTick(t => t + 1); };

  useEffect(() => {
    if (uiPhase !== "game") return;
    const canvas = canvasRef.current;
    const ctx = canvas.getContext("2d");

    const isSolid = (x, y) => {
      const tx = Math.floor(x); const ty = Math.floor(y);
      if (tx < 0 || tx >= MAP_W || ty < 0 || ty >= MAP_H) return true;
      return MAP_DATA[ty][tx] > 0;
    };

    const doHitCheck = (g, spreadAngle = 0) => {
      const shotAngle = g.angle + (Math.random() - 0.5) * spreadAngle;
      const allTargets = [
        ...g.enemies,
        ...(g.boss && g.boss.alive ? [{ ...g.boss, _isBoss: true }] : []),
      ];
      for (const e of allTargets) {
        if (!e.alive) continue;
        const dx = e.x - g.x; const dy = e.y - g.y;
        const dist = Math.sqrt(dx * dx + dy * dy);
        if (dist > 18) continue;
        const angleToEnemy = Math.atan2(dy, dx);
        let diff = angleToEnemy - shotAngle;
        while (diff > Math.PI) diff -= Math.PI * 2;
        while (diff < -Math.PI) diff += Math.PI * 2;
        if (Math.abs(diff) < 0.12 + 0.05 / dist) {
          const steps = Math.ceil(dist / 0.1);
          let blocked = false;
          for (let s = 1; s < steps; s++) {
            const fx = g.x + Math.cos(shotAngle) * s * 0.1;
            const fy = g.y + Math.sin(shotAngle) * s * 0.1;
            if (isSolid(fx, fy)) { blocked = true; break; }
          }
          if (blocked) continue;
          const dmg = GUNS[g.gun].dmg;
          if (e._isBoss) {
            g.boss.health -= dmg;
            if (g.boss.health <= 0) {
              g.boss.health = 0; g.boss.alive = false;
              g.enemyKills++; g.score += 500;
              const reward = Math.round(200 * DIFFICULTY[g.difficulty].killMoney / 25);
              g.money += reward;
              g.moneyPopups.push({ x: W / 2, y: H / 2 - 40, text: `+$${reward} BOSS SLAIN!`, timer: 90, color: "#f80" });
            }
          } else {
            e.health -= dmg;
            if (e.health <= 0) {
              e.alive = false; e.health = 0;
              g.enemyKills++; g.score += 100;
              const reward = DIFFICULTY[g.difficulty].killMoney;
              g.money += reward;
              g.moneyPopups.push({ x: W * 0.75 + Math.random() * 60, y: 50 + Math.random() * 40, text: `+$${reward}`, timer: 50, color: "#ff0" });
            }
          }
          return;
        }
      }
    };

    const shoot = () => {
      const g = stateRef.current;
      if (!["running","boss"].includes(g.phase)) return;
      if (g.shootCooldown > 0) return;
      const gun = GUNS[g.gun];
      g.shootCooldown = gun.cooldown;
      g.shootFlash = 8;
      for (let i = 0; i < gun.bulletsPerShot; i++) doHitCheck(g, gun.spread);
    };

    const loop = () => {
      const g = stateRef.current;
      if (!g) { animRef.current = requestAnimationFrame(loop); return; }

      g.frameCount = (g.frameCount || 0) + 1;
      g.moneyPopups = (g.moneyPopups || []).filter(p => p.timer > 0);
      for (const p of g.moneyPopups) { p.timer--; p.y -= 0.5; }

      const running = g.phase === "running" || g.phase === "boss";
      if (running) {
        const cos = Math.cos(g.angle); const sin = Math.sin(g.angle);
        // Perpendicular (strafe) direction
        const scos = Math.cos(g.angle + Math.PI / 2); const ssin = Math.sin(g.angle + Math.PI / 2);
        let nx = g.x, ny = g.y;
        if (keysRef.current["ArrowUp"]    || keysRef.current["KeyW"]) { nx += cos * PLAYER_SPEED; ny += sin * PLAYER_SPEED; }
        if (keysRef.current["ArrowDown"]  || keysRef.current["KeyS"]) { nx -= cos * PLAYER_SPEED; ny -= sin * PLAYER_SPEED; }
        if (keysRef.current["ArrowLeft"]  || keysRef.current["KeyA"]) { nx -= scos * PLAYER_SPEED; ny -= ssin * PLAYER_SPEED; }
        if (keysRef.current["ArrowRight"] || keysRef.current["KeyD"]) { nx += scos * PLAYER_SPEED; ny += ssin * PLAYER_SPEED; }
        if (!isSolid(nx, g.y)) g.x = nx;
        if (!isSolid(g.x, ny)) g.y = ny;
        if (g.shootCooldown > 0) g.shootCooldown--;
        if (g.shootFlash > 0) g.shootFlash--;
        if (g.hitFlash > 0) g.hitFlash--;
        g.bobTime += 0.1;
        g.score += 0.01;

        // Auto-fire
        if (GUNS[g.gun].auto && mouseDownRef.current) shoot();

        if (g.phase === "running") {
          // Enemy AI
          for (const e of g.enemies) {
            if (!e.alive) continue;
            const dx = g.x - e.x; const dy = g.y - e.y;
            const dist = Math.sqrt(dx * dx + dy * dy);
            if (dist < 8) e.alert = true;
            if (e.alert) {
              const nx2 = e.x + (dx / dist) * e.speed;
              const ny2 = e.y + (dy / dist) * e.speed;
              if (!isSolid(nx2, e.y)) e.x = nx2;
              if (!isSolid(e.x, ny2)) e.y = ny2;
              e.shootTimer--;
              if (e.shootTimer <= 0 && dist < 6) {
                e.shootTimer = Math.max(22, 70 - g.wave * 4) + Math.random() * 40;
                let blocked = false;
                const steps = Math.ceil(dist / 0.1);
                for (let s = 1; s < steps; s++) {
                  const fx = e.x + (dx / dist) * s * 0.1;
                  const fy = e.y + (dy / dist) * s * 0.1;
                  if (isSolid(fx, fy)) { blocked = true; break; }
                }
                if (!blocked) {
                  const dmg = Math.round(8 * DIFFICULTY[g.difficulty].enemyDmgMult);
                  g.health -= dmg; g.hitFlash = 20;
                  if (g.health <= 0) { g.health = 0; g.phase = "over"; }
                }
              }
            }
          }
          if (g.waveFlash > 0) {
            g.waveFlash--;
          } else if (g.enemies.length > 0 && g.enemies.every(e => !e.alive)) {
            spawnWave(g);
            if (g.phase === "shop") openShop();
          }
        } else if (g.phase === "boss") {
          const b = g.boss;
          if (!b || !b.alive) {
            g.phase = "running";
            spawnWave(g);
            if (g.phase === "shop") openShop();
          } else {
            const dx = g.x - b.x; const dy = g.y - b.y;
            const dist = Math.sqrt(dx * dx + dy * dy);
            b.moveTimer++;
            if (b.moveTimer >= 3) {
              b.moveTimer = 0;
              const nx2 = b.x + (dx / dist) * b.speed;
              const ny2 = b.y + (dy / dist) * b.speed;
              if (!isSolid(nx2, b.y)) b.x = nx2;
              if (!isSolid(b.x, ny2)) b.y = ny2;
            }
            b.fireballTimer--;
            if (b.fireballTimer <= 0) {
              b.fireballTimer = Math.max(80, 140 - g.wave * 3);
              const angle = Math.atan2(dy, dx);
              for (const sp of [-0.22, 0, 0.22]) {
                b.fireballs.push({ x: b.x, y: b.y, vx: Math.cos(angle + sp) * 0.07, vy: Math.sin(angle + sp) * 0.07, alive: true });
              }
            }
            b.fireballs = b.fireballs.filter(f => f.alive);
            for (const f of b.fireballs) {
              f.x += f.vx; f.y += f.vy;
              if (isSolid(f.x, f.y)) { f.alive = false; continue; }
              if (Math.hypot(g.x - f.x, g.y - f.y) < 0.5) {
                f.alive = false;
                const dmg = Math.round(20 * DIFFICULTY[g.difficulty].enemyDmgMult);
                g.health -= dmg; g.hitFlash = 30;
                if (g.health <= 0) { g.health = 0; g.phase = "over"; }
              }
            }
          }
          if (g.waveFlash > 0) g.waveFlash--;
        }
      }

      // ---- DRAW ----
      // Daytime sky
      const sky = ctx.createLinearGradient(0, 0, 0, H / 2);
      sky.addColorStop(0, "#4a90d9");
      sky.addColorStop(1, "#a8d8f0");
      ctx.fillStyle = sky; ctx.fillRect(0, 0, W, H / 2);

      // Daytime floor (grass)
      const floor = ctx.createLinearGradient(0, H / 2, 0, H);
      floor.addColorStop(0, "#6ab040");
      floor.addColorStop(1, "#4a8020");
      ctx.fillStyle = floor; ctx.fillRect(0, H / 2, W, H / 2);

      // Sun
      ctx.fillStyle = "#fff8a0";
      ctx.beginPath(); ctx.arc(W - 80, 44, 28, 0, Math.PI * 2); ctx.fill();
      ctx.fillStyle = "rgba(255,248,160,0.22)";
      ctx.beginPath(); ctx.arc(W - 80, 44, 44, 0, Math.PI * 2); ctx.fill();

      // Walls
      const rays = raycast(g.x, g.y, g.angle);
      const sliceW = W / NUM_RAYS;
      for (let i = 0; i < NUM_RAYS; i++) {
        const { dist, wallType, side } = rays[i];
        const wc = WALL_COLORS[wallType] || WALL_COLORS[1];
        const colors = side === 1 ? wc.v : wc.h;
        const sliceH = Math.min((H / (dist || 0.1)) * 1.2, H);
        const top = (H - sliceH) / 2;
        const darkFactor = Math.max(0, 1 - dist / MAX_DEPTH);
        const col = dist < 3 ? colors[0] : (dist < 7 ? colors[1] : "#c8a46e44");
        ctx.fillStyle = col;
        ctx.globalAlpha = darkFactor * 0.9 + 0.1;
        ctx.fillRect(i * sliceW, top, sliceW + 1, sliceH);
      }
      ctx.globalAlpha = 1;

      // Sprites: enemies + boss sorted back-to-front
      const sprites = [
        ...g.enemies.filter(e => e.alive).map(e => ({ ...e, _isBoss: false })),
        ...(g.boss && g.boss.alive ? [{ ...g.boss, _isBoss: true }] : []),
      ].sort((a, b) => Math.hypot(b.x - g.x, b.y - g.y) - Math.hypot(a.x - g.x, a.y - g.y));

      for (const e of sprites) {
        const dx = e.x - g.x; const dy = e.y - g.y;
        const dist = Math.sqrt(dx * dx + dy * dy);
        let sa = Math.atan2(dy, dx) - g.angle;
        while (sa > Math.PI) sa -= Math.PI * 2;
        while (sa < -Math.PI) sa += Math.PI * 2;
        if (Math.abs(sa) > HALF_FOV + 0.1) continue;
        const screenX = (0.5 + sa / FOV) * W;
        const scale = e._isBoss ? 1.6 : 0.9;
        const size = Math.min((H / (dist || 0.1)) * scale, H);
        const sx = screenX - size / 2;
        const sy = (H - size) / 2;
        ctx.globalAlpha = Math.max(0, 1 - dist / (e._isBoss ? 22 : 12));

        if (e._isBoss) {
          // Big demon boss
          ctx.fillStyle = "#b02800"; ctx.fillRect(sx + size * 0.2, sy + size * 0.15, size * 0.6, size * 0.6);
          ctx.fillStyle = "#d03800"; ctx.fillRect(sx + size * 0.28, sy, size * 0.44, size * 0.22);
          ctx.fillStyle = "#700";
          ctx.fillRect(sx + size * 0.27, sy - size * 0.14, size * 0.09, size * 0.16);
          ctx.fillRect(sx + size * 0.64, sy - size * 0.14, size * 0.09, size * 0.16);
          ctx.fillStyle = "#ff0";
          ctx.fillRect(sx + size * 0.33, sy + size * 0.05, size * 0.10, size * 0.09);
          ctx.fillRect(sx + size * 0.57, sy + size * 0.05, size * 0.10, size * 0.09);
          ctx.fillStyle = "#901800";
          ctx.fillRect(sx + size * 0.04, sy + size * 0.18, size * 0.17, size * 0.42);
          ctx.fillRect(sx + size * 0.79, sy + size * 0.18, size * 0.17, size * 0.42);
          ctx.fillRect(sx + size * 0.24, sy + size * 0.73, size * 0.22, size * 0.22);
          ctx.fillRect(sx + size * 0.54, sy + size * 0.73, size * 0.22, size * 0.22);
          // HP bar
          ctx.fillStyle = "#400"; ctx.fillRect(sx + size * 0.1, sy - size * 0.1, size * 0.8, size * 0.07);
          ctx.fillStyle = "#f40"; ctx.fillRect(sx + size * 0.1, sy - size * 0.1, size * 0.8 * (e.health / e.maxHealth), size * 0.07);
          ctx.fillStyle = "#ff0"; ctx.font = `bold ${Math.max(10, size * 0.1)}px monospace`;
          ctx.textAlign = "center"; ctx.fillText("BOSS", sx + size / 2, sy - size * 0.14);
        } else {
          ctx.fillStyle = e.health > 2 ? "#cc3333" : (e.health > 1 ? "#cc6633" : "#996633");
          ctx.fillRect(sx + size * 0.25, sy + size * 0.15, size * 0.5, size * 0.55);
          ctx.fillStyle = "#f5c58a"; ctx.fillRect(sx + size * 0.32, sy, size * 0.36, size * 0.22);
          ctx.fillStyle = "#200000";
          ctx.fillRect(sx + size * 0.36, sy + size * 0.06, size * 0.08, size * 0.08);
          ctx.fillRect(sx + size * 0.56, sy + size * 0.06, size * 0.08, size * 0.08);
          ctx.fillStyle = "#300"; ctx.fillRect(sx + size * 0.2, sy - size * 0.08, size * 0.6, size * 0.06);
          ctx.fillStyle = "#0f0"; ctx.fillRect(sx + size * 0.2, sy - size * 0.08, size * 0.6 * (e.health / e.maxHealth), size * 0.06);
        }
        ctx.globalAlpha = 1;
      }

      // Boss fireballs
      if (g.boss && g.boss.fireballs) {
        for (const f of g.boss.fireballs) {
          if (!f.alive) continue;
          const dx = f.x - g.x; const dy = f.y - g.y;
          let sa = Math.atan2(dy, dx) - g.angle;
          while (sa > Math.PI) sa -= Math.PI * 2;
          while (sa < -Math.PI) sa += Math.PI * 2;
          if (Math.abs(sa) > HALF_FOV + 0.2) continue;
          const dist = Math.sqrt(dx * dx + dy * dy);
          const screenX = (0.5 + sa / FOV) * W;
          const fSize = Math.min((H / (dist || 0.1)) * 0.35, 80);
          const pulse = 0.82 + 0.18 * Math.sin(g.frameCount * 0.3 + f.x * 3);
          const gr = ctx.createRadialGradient(screenX, H / 2, 0, screenX, H / 2, fSize * pulse);
          gr.addColorStop(0, "rgba(255,240,80,0.95)");
          gr.addColorStop(0.4, "rgba(255,110,0,0.85)");
          gr.addColorStop(1, "rgba(200,20,0,0)");
          ctx.fillStyle = gr;
          ctx.beginPath(); ctx.arc(screenX, H / 2, fSize * pulse, 0, Math.PI * 2); ctx.fill();
        }
      }

      // Hit flash
      if (g.hitFlash > 0) { ctx.fillStyle = `rgba(255,0,0,${(g.hitFlash / 30) * 0.42})`; ctx.fillRect(0, 0, W, H); }

      // Muzzle flash
      if (g.shootFlash > 0) {
        if (g.gun === "laser") {
          ctx.strokeStyle = `rgba(255,0,255,${(g.shootFlash / 8) * 0.9})`;
          ctx.lineWidth = 3; ctx.beginPath(); ctx.moveTo(W / 2, H / 2); ctx.lineTo(W, H / 2 + 5); ctx.stroke();
          ctx.strokeStyle = "rgba(255,180,255,0.35)"; ctx.lineWidth = 10; ctx.stroke();
        } else {
          const grd = ctx.createRadialGradient(W / 2, H / 2 + 60, 0, W / 2, H / 2 + 60, 55);
          grd.addColorStop(0, `rgba(255,220,80,${(g.shootFlash / 8) * 0.7})`);
          grd.addColorStop(1, "rgba(255,100,0,0)");
          ctx.fillStyle = grd; ctx.fillRect(W / 2 - 60, H / 2, 120, 120);
        }
      }

      // Gun sprite
      const bobY = running ? Math.sin(g.bobTime * 2) * 4 : 0;
      if (g.gun === "pistol") {
        ctx.fillStyle = "#aaaaaa"; ctx.fillRect(W / 2 + 48, H - 85 + bobY, 44, 60);
        ctx.fillStyle = "#888"; ctx.fillRect(W / 2 + 52, H - 98 + bobY, 12, 18);
        ctx.fillStyle = "#666"; ctx.fillRect(W / 2 + 52, H - 102 + bobY, 10, 6);
      } else if (g.gun === "semi") {
        ctx.fillStyle = "#5aaa88"; ctx.fillRect(W / 2 + 38, H - 88 + bobY, 64, 62);
        ctx.fillStyle = "#3a8866"; ctx.fillRect(W / 2 + 42, H - 104 + bobY, 14, 22);
        ctx.fillStyle = "#287050"; ctx.fillRect(W / 2 + 58, H - 82 + bobY, 12, 20);
        // 3-barrel hint
        for (let i = 0; i < 3; i++) ctx.fillRect(W / 2 + 42 + i * 4, H - 108 + bobY, 3, 6);
      } else if (g.gun === "rifle") {
        ctx.fillStyle = "#5599cc"; ctx.fillRect(W / 2 + 22, H - 92 + bobY, 92, 58);
        ctx.fillStyle = "#3877aa"; ctx.fillRect(W / 2 + 22, H - 108 + bobY, 16, 24);
        ctx.fillStyle = "#2a5080"; ctx.fillRect(W / 2 + 60, H - 86 + bobY, 18, 22);
        ctx.fillStyle = "#1a3060"; ctx.fillRect(W / 2 + 82, H - 78 + bobY, 10, 8);
        // stock
        ctx.fillStyle = "#446688"; ctx.fillRect(W / 2 + 100, H - 72 + bobY, 16, 40);
      } else if (g.gun === "laser") {
        ctx.fillStyle = "#cc44cc"; ctx.fillRect(W / 2 + 28, H - 94 + bobY, 86, 52);
        ctx.fillStyle = "#aa00aa"; ctx.fillRect(W / 2 + 28, H - 112 + bobY, 10, 24);
        // Glowing tip
        const grd2 = ctx.createRadialGradient(W / 2 + 33, H - 112 + bobY, 0, W / 2 + 33, H - 112 + bobY, 14);
        grd2.addColorStop(0, "rgba(255,100,255,0.9)"); grd2.addColorStop(1, "rgba(200,0,200,0)");
        ctx.fillStyle = grd2; ctx.fillRect(W / 2 + 20, H - 124 + bobY, 26, 26);
        ctx.fillStyle = "#880088"; ctx.fillRect(W / 2 + 70, H - 88 + bobY, 16, 20);
        ctx.fillStyle = "#550055"; ctx.fillRect(W / 2 + 90, H - 78 + bobY, 12, 10);
      }

      // Crosshair
      const chColor = g.gun === "laser" ? "rgba(255,80,255,0.9)" : "rgba(255,255,255,0.8)";
      ctx.strokeStyle = chColor; ctx.lineWidth = 1.5;
      ctx.beginPath();
      ctx.moveTo(W / 2 - 12, H / 2); ctx.lineTo(W / 2 + 12, H / 2);
      ctx.moveTo(W / 2, H / 2 - 12); ctx.lineTo(W / 2, H / 2 + 12);
      ctx.stroke();
      ctx.beginPath(); ctx.arc(W / 2, H / 2, 4, 0, Math.PI * 2); ctx.stroke();

      // HUD
      ctx.font = "bold 14px monospace"; ctx.textAlign = "left";
      ctx.fillStyle = "#fff";
      ctx.fillText(`❤ ${Math.max(0, g.health)}%   💀 ${g.enemyKills}   W${g.wave}   [${GUNS[g.gun].name}]`, 14, 28);
      ctx.fillStyle = "#ffe000"; ctx.fillText(`$${g.money}`, 14, 48);
      ctx.textAlign = "right"; ctx.fillStyle = "#ff0";
      ctx.fillText(`SCORE ${Math.floor(g.score)}`, W - 14, 28);

      // Money popups
      for (const p of g.moneyPopups) {
        ctx.globalAlpha = Math.min(1, p.timer / 30);
        ctx.fillStyle = p.color; ctx.font = "bold 15px monospace"; ctx.textAlign = "center";
        ctx.fillText(p.text, p.x, p.y);
      }
      ctx.globalAlpha = 1;

      // Wave banner
      if (g.waveFlash > 60) {
        const t = Math.min(1, (g.waveFlash - 60) / 25);
        ctx.globalAlpha = t;
        ctx.fillStyle = "rgba(0,0,0,0.65)"; ctx.fillRect(W / 2 - 220, H / 2 - 46, 440, 82);
        const isBossWave = g.phase === "boss";
        ctx.fillStyle = isBossWave ? "#f84" : "#ffe000";
        ctx.font = "bold 34px monospace"; ctx.textAlign = "center";
        ctx.fillText(isBossWave ? `⚠  BOSS — WAVE ${g.wave}  ⚠` : `◆  WAVE ${g.wave}  ◆`, W / 2, H / 2 - 6);
        ctx.fillStyle = "#ccc"; ctx.font = "13px monospace";
        ctx.fillText(isBossWave ? "A fire demon approaches — dodge the fireballs!" : `${g.enemies.length} enemies incoming`, W / 2, H / 2 + 22);
        ctx.globalAlpha = 1;
      }

      // Minimap
      const mox = 14, moy = H - MAP_H * MAP_SCALE - 14;
      ctx.fillStyle = "rgba(0,0,0,0.45)"; ctx.fillRect(mox - 2, moy - 2, MAP_W * MAP_SCALE + 4, MAP_H * MAP_SCALE + 4);
      for (let r = 0; r < MAP_H; r++) for (let c = 0; c < MAP_W; c++) {
        const cell = MAP_DATA[r][c];
        ctx.fillStyle = cell > 0 ? WALL_COLORS[cell]?.h[0] ?? "#888" : "#222";
        ctx.fillRect(mox + c * MAP_SCALE, moy + r * MAP_SCALE, MAP_SCALE - 1, MAP_SCALE - 1);
      }
      for (const e of g.enemies) {
        if (!e.alive) continue;
        ctx.fillStyle = "#f44"; ctx.fillRect(mox + e.x * MAP_SCALE - 3, moy + e.y * MAP_SCALE - 3, 6, 6);
      }
      if (g.boss && g.boss.alive) {
        ctx.fillStyle = "#f80"; ctx.fillRect(mox + g.boss.x * MAP_SCALE - 5, moy + g.boss.y * MAP_SCALE - 5, 10, 10);
      }
      ctx.fillStyle = "#0f0"; ctx.beginPath(); ctx.arc(mox + g.x * MAP_SCALE, moy + g.y * MAP_SCALE, 4, 0, Math.PI * 2); ctx.fill();
      ctx.strokeStyle = "#0f0"; ctx.lineWidth = 1.5;
      ctx.beginPath();
      ctx.moveTo(mox + g.x * MAP_SCALE, moy + g.y * MAP_SCALE);
      ctx.lineTo(mox + g.x * MAP_SCALE + Math.cos(g.angle) * 10, moy + g.y * MAP_SCALE + Math.sin(g.angle) * 10);
      ctx.stroke();

      // Phase overlays
      ctx.textAlign = "center";
      if (g.phase === "idle") {
        ctx.fillStyle = "rgba(0,0,0,0.72)"; ctx.fillRect(0, 0, W, H);
        ctx.fillStyle = "#ffe000"; ctx.font = "bold 32px monospace"; ctx.fillText("DUNGEON FPS", W / 2, H / 2 - 50);
        ctx.fillStyle = "#aaa"; ctx.font = "13px monospace";
        ctx.fillText("WASD / ↑↓←→ — Move    Mouse — Turn    Click / SPACE — Shoot", W / 2, H / 2 - 8);
        ctx.fillText("Kill enemies → earn money → buy guns → survive boss every 5 waves", W / 2, H / 2 + 16);
        ctx.fillStyle = "#0f0"; ctx.font = "bold 18px monospace";
        ctx.fillText("▶  Press SPACE to Start  ◀", W / 2, H / 2 + 58);
      } else if (g.phase === "over") {
        ctx.fillStyle = "rgba(0,0,0,0.78)"; ctx.fillRect(0, 0, W, H);
        ctx.fillStyle = "#f44"; ctx.font = "bold 44px monospace"; ctx.fillText("YOU DIED", W / 2, H / 2 - 20);
        ctx.fillStyle = "#fff"; ctx.font = "18px monospace";
        ctx.fillText(`Wave: ${g.wave}  •  Score: ${Math.floor(g.score)}  •  Kills: ${g.enemyKills}`, W / 2, H / 2 + 24);
        ctx.fillStyle = "#aaa"; ctx.font = "14px monospace"; ctx.fillText("Press SPACE to Restart", W / 2, H / 2 + 60);
      }

      animRef.current = requestAnimationFrame(loop);
    };

    const onKeyDown = (e) => {
      const g = stateRef.current;

      // Shop keyboard navigation — intercept arrows + Enter before game loop
      if (g?.phase === "shop") {
        const navKeys = ["ArrowLeft","ArrowRight","ArrowUp","ArrowDown","Enter","Space"];
        if (navKeys.includes(e.code)) e.preventDefault();
        const gunKeys = Object.keys(GUNS); // ["pistol","semi","rifle","laser"]
        const total = gunKeys.length + 1; // +1 for Start Wave button
        const cur = shopCursorRef.current;
        if (e.code === "ArrowRight" || e.code === "ArrowDown") {
          setShopCursorBoth((cur + 1) % total);
        } else if (e.code === "ArrowLeft" || e.code === "ArrowUp") {
          setShopCursorBoth((cur - 1 + total) % total);
        } else if (e.code === "Enter" || e.code === "Space") {
          if (cur < gunKeys.length) {
            // Select a gun card
            const key = gunKeys[cur];
            const gun = GUNS[key];
            if (g.gun !== key && g.money >= gun.cost) {
              g.money -= gun.cost;
              g.gun = key;
              setShopTick(t => t + 1);
            }
          } else {
            // Start Wave button
            startNormalWave(g);
            setShopTick(t => t + 1);
          }
        }
        return; // don't pass shop keys to game loop
      }

      keysRef.current[e.code] = true;
      if (["Space","ArrowUp","ArrowDown","ArrowLeft","ArrowRight"].includes(e.code)) e.preventDefault();
      if (!g) return;
      if (e.code === "Space") {
        if (g.phase === "idle") {
          g.phase = "running"; spawnWave(g); if (g.phase === "shop") openShop(); return;
        }
        if (g.phase === "over") {
          stateRef.current = makeState(g.difficulty);
          stateRef.current.phase = "running"; spawnWave(stateRef.current);
          if (stateRef.current.phase === "shop") openShop(); return;
        }
        if (!GUNS[g.gun].auto) shoot();
      }
    };
    const onKeyUp = (e) => { keysRef.current[e.code] = false; };
    const onMouseMove = (e) => {
      if (!lockedRef.current) return;
      const g = stateRef.current;
      if (g && ["running","boss"].includes(g.phase)) g.angle += e.movementX * MOUSE_SENS;
    };
    const onMouseDown = (e) => {
      if (e.button !== 0) return;
      mouseDownRef.current = true;
      if (!lockedRef.current) { canvas.requestPointerLock(); return; }
      const g = stateRef.current; if (!g) return;
      if (!GUNS[g.gun].auto) shoot();
    };
    const onMouseUp = () => { mouseDownRef.current = false; };
    const onLockChange = () => { lockedRef.current = document.pointerLockElement === canvas; };

    window.addEventListener("keydown", onKeyDown);
    window.addEventListener("keyup", onKeyUp);
    document.addEventListener("mousemove", onMouseMove);
    canvas.addEventListener("mousedown", onMouseDown);
    window.addEventListener("mouseup", onMouseUp);
    document.addEventListener("pointerlockchange", onLockChange);
    animRef.current = requestAnimationFrame(loop);

    return () => {
      cancelAnimationFrame(animRef.current);
      window.removeEventListener("keydown", onKeyDown);
      window.removeEventListener("keyup", onKeyUp);
      document.removeEventListener("mousemove", onMouseMove);
      canvas.removeEventListener("mousedown", onMouseDown);
      window.removeEventListener("mouseup", onMouseUp);
      document.removeEventListener("pointerlockchange", onLockChange);
      if (document.pointerLockElement === canvas) document.exitPointerLock();
    };
  }, [uiPhase]);

  const buyGun = (key) => {
    const g = stateRef.current; if (!g) return;
    const gun = GUNS[key];
    if (g.money >= gun.cost && g.gun !== key) { g.money -= gun.cost; g.gun = key; }
    setShopTick(t => t + 1);
  };

  const closeShop = () => {
    const g = stateRef.current; if (!g) return;
    startNormalWave(g);
    setShopTick(t => t + 1);
  };

  const shopOpen = stateRef.current?.phase === "shop";

  return (
    <div className="flex flex-col items-center gap-3">
      {uiPhase === "diffpick" && (
        <div className="flex flex-col items-center gap-6 py-10">
          <h2 className="text-3xl font-bold text-yellow-300 font-mono">DUNGEON FPS</h2>
          <p className="text-gray-400 font-mono text-sm">Select Difficulty</p>
          <div className="flex gap-4">
            {Object.entries(DIFFICULTY).map(([key, d]) => (
              <button key={key} onClick={() => startGame(key)}
                className={`px-6 py-3 rounded-lg font-mono font-bold text-lg border-2 transition-all
                  ${key === "easy"   ? "border-green-500 text-green-400 hover:bg-green-900" : ""}
                  ${key === "normal" ? "border-yellow-400 text-yellow-300 hover:bg-yellow-900" : ""}
                  ${key === "hard"   ? "border-red-500 text-red-400 hover:bg-red-900" : ""}
                `}>{d.label}</button>
            ))}
          </div>
          <div className="text-gray-500 font-mono text-xs text-center space-y-1">
            <p>Easy — less damage, more money per kill ($40)</p>
            <p>Normal — balanced ($25/kill)</p>
            <p>Hard — fast enemies, high damage ($15/kill)</p>
            <p className="text-gray-400 mt-2">Boss wave every 5 waves • 4 guns to unlock</p>
          </div>
        </div>
      )}

      {uiPhase === "game" && (
        <div className="relative flex flex-col items-center gap-3">
          {/* Shop overlay */}
          {shopOpen && (
            <div className="absolute inset-0 flex items-center justify-center z-50 bg-black/75 rounded-xl">
              <div className="bg-gray-900 border-2 border-yellow-400 rounded-2xl p-6 w-[480px] font-mono shadow-2xl">
                <h3 className="text-yellow-300 text-xl font-bold text-center mb-1">🔫 GUN SHOP</h3>
                <p className="text-green-400 text-center text-sm mb-4">
                  Cash: <span className="text-yellow-200 font-bold text-lg">${stateRef.current?.money ?? 0}</span>
                </p>
                <div className="grid grid-cols-2 gap-3 mb-5">
                  {Object.entries(GUNS).map(([key, gun], idx) => {
                    const g = stateRef.current;
                    const owned = g?.gun === key;
                    const canAfford = (g?.money ?? 0) >= gun.cost;
                    const highlighted = shopCursor === idx;
                    return (
                      <div key={key}
                        onClick={() => !owned && canAfford && buyGun(key)}
                        className={`rounded-xl border-2 p-3 flex flex-col gap-1 transition-all
                          ${owned ? "border-green-400 bg-green-950" :
                            canAfford ? "border-gray-500 bg-gray-800 hover:border-yellow-400 cursor-pointer" :
                            "border-gray-700 bg-gray-900 opacity-40 cursor-not-allowed"}
                          ${highlighted ? "ring-2 ring-yellow-300 ring-offset-1 ring-offset-gray-900 scale-[1.03]" : ""}
                        `}
                      >
                        <span className="font-bold text-white text-sm">{gun.name}</span>
                        <span className="text-xs text-gray-400">
                          {gun.auto ? "🔄 Full-Auto" : `💥 ${gun.bulletsPerShot} bullet${gun.bulletsPerShot > 1 ? "s" : ""}/click`}
                        </span>
                        <span className="text-xs text-gray-400">DMG: {gun.dmg} · Rate: {Math.round(60 / gun.cooldown)}/s</span>
                        <span className={`text-sm font-bold mt-1 ${gun.cost === 0 ? "text-green-400" : owned ? "text-green-300" : canAfford ? "text-yellow-300" : "text-red-400"}`}>
                          {gun.cost === 0 ? "FREE (default)" : owned ? "✓ EQUIPPED" : `$${gun.cost}`}
                        </span>
                      </div>
                    );
                  })}
                </div>
                <button onClick={closeShop}
                  className={`w-full py-3 bg-yellow-400 text-gray-900 font-bold rounded-lg text-lg hover:bg-yellow-300 transition-all
                    ${shopCursor === 4 ? "ring-2 ring-white ring-offset-1 ring-offset-gray-900 scale-[1.02]" : ""}`}>
                  ▶ Start Wave {stateRef.current?.wave ?? ""}
                </button>
              </div>
            </div>
          )}

          <canvas ref={canvasRef} width={W} height={H}
            className="border-2 border-yellow-400 rounded-xl max-w-full cursor-crosshair"
          />
          <p className="text-sm text-[#999] font-mono">
            WASD / Arrows Move &amp; Strafe &nbsp;•&nbsp; Mouse Turn &nbsp;•&nbsp; Click / SPACE Shoot &nbsp;•&nbsp; Click canvas to lock mouse
            {shopOpen && <span className="text-yellow-400"> &nbsp;|&nbsp; Shop: ←→ Navigate &nbsp;•&nbsp; Enter / Space Select</span>}
          </p>
        </div>
      )}
    </div>
  );
}
