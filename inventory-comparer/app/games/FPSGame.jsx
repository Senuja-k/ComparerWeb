"use client";

import { useEffect, useRef, useState } from "react";
import * as THREE from "three";

const W = 800;
const H = 450;
const PLAYER_SPEED = 0.038;
const MOUSE_SENS   = 0.0022;
const MAP_SCALE    = 9;
const MAX_ENEMIES  = 16;
const MAX_FIREBALLS = 20;

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

// Three.js wall hex colours
const WALL_HEX = { 1: 0xc8a46e, 2: 0x7094b8, 3: 0x60a060 };

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

// Derive valid open-floor spawn positions directly from MAP_DATA so no
// position ever lands inside a wall, regardless of future map edits.
const SPAWN_POOL = MAP_DATA.flatMap((row, r) =>
  row.flatMap((cell, c) => (cell === 0 ? [[c + 0.5, r + 0.5]] : []))
);

function spawnWave(g) {
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

// â”€â”€ Component â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€

export default function FPSGame() {
  const mountRef     = useRef(null); // Three.js renderer mounts here
  const minimapRef   = useRef(null); // 2D minimap canvas
  const overlayRef   = useRef(null); // 2D overlay: gun sprite, crosshair, flash, banners
  const stateRef     = useRef(null);
  const keysRef      = useRef({});
  const animRef      = useRef(null);
  const lockedRef    = useRef(false);
  const mouseDownRef = useRef(false);
  const threeRef     = useRef(null);

  const [uiPhase, setUiPhase]       = useState("diffpick");
  const [shopTick, setShopTick]     = useState(0);
  const [shopCursor, setShopCursor] = useState(0);
  const [hudState, setHudState]     = useState({
    health: 100, score: 0, money: 0, kills: 0, wave: 0, gun: "pistol",
  });
  const shopCursorRef = useRef(0);

  const setShopCursorBoth = (v) => { shopCursorRef.current = v; setShopCursor(v); };
  const openShop = () => { setShopCursorBoth(0); setShopTick(t => t + 1); };

  const startGame = (diff) => {
    stateRef.current = makeState(diff);
    setUiPhase("game");
  };

  // â”€â”€ Three.js + game loop â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
  useEffect(() => {
    if (uiPhase !== "game") return;
    const mount = mountRef.current;

    // Scene
    const scene = new THREE.Scene();
    scene.background = new THREE.Color(0x4a90d9);
    scene.fog = new THREE.Fog(0x88c0e0, 10, 28);

    // Lights
    scene.add(new THREE.AmbientLight(0xffffff, 0.55));
    const sun = new THREE.DirectionalLight(0xfff8c0, 0.95);
    sun.position.set(12, 20, 8);
    scene.add(sun);

    // Floor
    const floorMesh = new THREE.Mesh(
      new THREE.PlaneGeometry(MAP_W, MAP_H),
      new THREE.MeshLambertMaterial({ color: 0x5a9030 })
    );
    floorMesh.rotation.x = -Math.PI / 2;
    floorMesh.position.set(MAP_W / 2, 0, MAP_H / 2);
    scene.add(floorMesh);

    // Ceiling
    const ceilMesh = new THREE.Mesh(
      new THREE.PlaneGeometry(MAP_W, MAP_H),
      new THREE.MeshLambertMaterial({ color: 0x7ab8d0, side: THREE.BackSide })
    );
    ceilMesh.rotation.x = -Math.PI / 2;
    ceilMesh.position.set(MAP_W / 2, 1, MAP_H / 2);
    scene.add(ceilMesh);

    // Static walls
    for (let r = 0; r < MAP_H; r++) {
      for (let c = 0; c < MAP_W; c++) {
        const cell = MAP_DATA[r][c];
        if (cell > 0) {
          const wall = new THREE.Mesh(
            new THREE.BoxGeometry(1, 1, 1),
            new THREE.MeshLambertMaterial({ color: WALL_HEX[cell] ?? 0xc8a46e })
          );
          wall.position.set(c + 0.5, 0.5, r + 0.5);
          scene.add(wall);
        }
      }
    }

    // Camera
    const camera = new THREE.PerspectiveCamera(75, W / H, 0.05, 50);
    camera.position.set(2.5, 0.5, 2.5);

    // Renderer
    const renderer = new THREE.WebGLRenderer({ antialias: true });
    renderer.setSize(W, H);
    renderer.setPixelRatio(Math.min(window.devicePixelRatio, 2));
    mount.appendChild(renderer.domElement);

    // â”€â”€ Enemy mesh pool â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
    const enemyPool = [];
    for (let i = 0; i < MAX_ENEMIES; i++) {
      const group = new THREE.Group();

      const bodyMat = new THREE.MeshLambertMaterial({ color: 0xcc3333 });
      const body = new THREE.Mesh(new THREE.BoxGeometry(0.5, 0.65, 0.3), bodyMat);
      body.position.y = 0.32;
      group.add(body);

      const head = new THREE.Mesh(
        new THREE.BoxGeometry(0.36, 0.36, 0.32),
        new THREE.MeshLambertMaterial({ color: 0xf0b070 })
      );
      head.position.y = 0.82;
      group.add(head);

      const eyeMat = new THREE.MeshBasicMaterial({ color: 0x200000 });
      const eyeL = new THREE.Mesh(new THREE.BoxGeometry(0.1, 0.1, 0.02), eyeMat);
      eyeL.position.set(-0.1, 0.86, 0.17);
      group.add(eyeL);
      const eyeR = eyeL.clone();
      eyeR.position.set(0.1, 0.86, 0.17);
      group.add(eyeR);

      const hpBg = new THREE.Mesh(
        new THREE.PlaneGeometry(0.6, 0.07),
        new THREE.MeshBasicMaterial({ color: 0x400000, side: THREE.DoubleSide })
      );
      hpBg.position.set(0, 1.18, 0);
      group.add(hpBg);

      const hpBar = new THREE.Mesh(
        new THREE.PlaneGeometry(0.6, 0.07),
        new THREE.MeshBasicMaterial({ color: 0x00ee00, side: THREE.DoubleSide })
      );
      hpBar.position.set(0, 1.18, 0.001);
      group.add(hpBar);

      group.visible = false;
      scene.add(group);
      enemyPool.push({ group, bodyMat, hpBar });
    }

    // â”€â”€ Boss mesh â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
    const bossGroup  = new THREE.Group();
    const bossBodyMat = new THREE.MeshLambertMaterial({ color: 0xb02800 });

    const bossBody = new THREE.Mesh(new THREE.BoxGeometry(1.2, 1.5, 0.9), bossBodyMat);
    bossBody.position.set(0, 0.75, 0);
    bossGroup.add(bossBody);

    const bossHead = new THREE.Mesh(
      new THREE.BoxGeometry(0.9, 0.75, 0.75),
      new THREE.MeshLambertMaterial({ color: 0xd03800 })
    );
    bossHead.position.set(0, 1.625, 0);
    bossGroup.add(bossHead);

    const hornMat = new THREE.MeshLambertMaterial({ color: 0x700000 });
    const hornL = new THREE.Mesh(new THREE.ConeGeometry(0.1, 0.5, 6), hornMat);
    hornL.position.set(-0.32, 2.15, 0);
    bossGroup.add(hornL);
    const hornR = hornL.clone();
    hornR.position.set(0.32, 2.15, 0);
    bossGroup.add(hornR);

    const bEyeMat = new THREE.MeshBasicMaterial({ color: 0xffff00 });
    const bEyeL = new THREE.Mesh(new THREE.BoxGeometry(0.18, 0.18, 0.02), bEyeMat);
    bEyeL.position.set(-0.24, 1.68, 0.38);
    bossGroup.add(bEyeL);
    const bEyeR = bEyeL.clone();
    bEyeR.position.set(0.24, 1.68, 0.38);
    bossGroup.add(bEyeR);

    const armMat = new THREE.MeshLambertMaterial({ color: 0x901800 });
    const armL = new THREE.Mesh(new THREE.BoxGeometry(0.25, 1.0, 0.25), armMat);
    armL.position.set(-0.73, 0.8, 0);
    bossGroup.add(armL);
    const armR = armL.clone();
    armR.position.set(0.73, 0.8, 0);
    bossGroup.add(armR);

    const bossHpBg = new THREE.Mesh(
      new THREE.PlaneGeometry(1.4, 0.12),
      new THREE.MeshBasicMaterial({ color: 0x400000, side: THREE.DoubleSide })
    );
    bossHpBg.position.set(0, 2.65, 0);
    bossGroup.add(bossHpBg);

    const bossHpBar = new THREE.Mesh(
      new THREE.PlaneGeometry(1.4, 0.12),
      new THREE.MeshBasicMaterial({ color: 0xff4400, side: THREE.DoubleSide })
    );
    bossHpBar.position.set(0, 2.65, 0.001);
    bossGroup.add(bossHpBar);

    bossGroup.visible = false;
    scene.add(bossGroup);

    // â”€â”€ Fireball pool â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
    const fireballPool = [];
    for (let i = 0; i < MAX_FIREBALLS; i++) {
      const fb = new THREE.Mesh(
        new THREE.SphereGeometry(0.22, 8, 8),
        new THREE.MeshBasicMaterial({ color: 0xff6600 })
      );
      fb.visible = false;
      scene.add(fb);
      fireballPool.push(fb);
    }

    threeRef.current = { renderer, enemyPool, bossGroup, bossHpBar, bossBodyMat, fireballPool };

    // â”€â”€ Minimap draw â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
    const drawMinimap = (g) => {
      const mc = minimapRef.current;
      if (!mc) return;
      const ctx = mc.getContext("2d");
      ctx.clearRect(0, 0, mc.width, mc.height);
      ctx.fillStyle = "rgba(0,0,0,0.50)";
      ctx.fillRect(0, 0, mc.width, mc.height);
      for (let r = 0; r < MAP_H; r++) {
        for (let c = 0; c < MAP_W; c++) {
          const cell = MAP_DATA[r][c];
          ctx.fillStyle = cell === 1 ? "#c8a46e" : cell === 2 ? "#7094b8" : cell === 3 ? "#60a060" : "#222";
          ctx.fillRect(c * MAP_SCALE, r * MAP_SCALE, MAP_SCALE - 1, MAP_SCALE - 1);
        }
      }
      for (const e of g.enemies) {
        if (!e.alive) continue;
        ctx.fillStyle = "#f44";
        ctx.fillRect(e.x * MAP_SCALE - 3, e.y * MAP_SCALE - 3, 6, 6);
      }
      if (g.boss?.alive) {
        ctx.fillStyle = "#f80";
        ctx.fillRect(g.boss.x * MAP_SCALE - 5, g.boss.y * MAP_SCALE - 5, 10, 10);
      }
      ctx.fillStyle = "#0f0";
      ctx.beginPath(); ctx.arc(g.x * MAP_SCALE, g.y * MAP_SCALE, 4, 0, Math.PI * 2); ctx.fill();
      ctx.strokeStyle = "#0f0"; ctx.lineWidth = 1.5;
      ctx.beginPath();
      ctx.moveTo(g.x * MAP_SCALE, g.y * MAP_SCALE);
      ctx.lineTo(g.x * MAP_SCALE + Math.cos(g.angle) * 10, g.y * MAP_SCALE + Math.sin(g.angle) * 10);
      ctx.stroke();
    };

    // â”€â”€ Overlay canvas â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
    const drawOverlay = (g) => {
      const oc = overlayRef.current;
      if (!oc) return;
      const ctx = oc.getContext("2d");
      ctx.clearRect(0, 0, W, H);

      if (g.hitFlash > 0) {
        ctx.fillStyle = `rgba(255,0,0,${(g.hitFlash / 30) * 0.42})`;
        ctx.fillRect(0, 0, W, H);
      }

      if (g.shootFlash > 0) {
        if (g.gun === "laser") {
          ctx.strokeStyle = `rgba(255,0,255,${(g.shootFlash / 8) * 0.9})`;
          ctx.lineWidth = 3; ctx.beginPath();
          ctx.moveTo(W / 2, H / 2); ctx.lineTo(W, H / 2 + 5); ctx.stroke();
          ctx.strokeStyle = "rgba(255,180,255,0.35)"; ctx.lineWidth = 10; ctx.stroke();
        } else {
          const grd = ctx.createRadialGradient(W / 2, H / 2 + 60, 0, W / 2, H / 2 + 60, 55);
          grd.addColorStop(0, `rgba(255,220,80,${(g.shootFlash / 8) * 0.7})`);
          grd.addColorStop(1, "rgba(255,100,0,0)");
          ctx.fillStyle = grd; ctx.fillRect(W / 2 - 60, H / 2, 120, 120);
        }
      }

      const running = ["running", "boss"].includes(g.phase);
      const bobY = running ? Math.sin(g.bobTime * 2) * 4 : 0;
      if (g.gun === "pistol") {
        ctx.fillStyle = "#aaaaaa"; ctx.fillRect(W / 2 + 48, H - 85 + bobY, 44, 60);
        ctx.fillStyle = "#888";    ctx.fillRect(W / 2 + 52, H - 98 + bobY, 12, 18);
        ctx.fillStyle = "#666";    ctx.fillRect(W / 2 + 52, H - 102 + bobY, 10, 6);
      } else if (g.gun === "semi") {
        ctx.fillStyle = "#5aaa88"; ctx.fillRect(W / 2 + 38, H - 88 + bobY, 64, 62);
        ctx.fillStyle = "#3a8866"; ctx.fillRect(W / 2 + 42, H - 104 + bobY, 14, 22);
        ctx.fillStyle = "#287050"; ctx.fillRect(W / 2 + 58, H - 82 + bobY, 12, 20);
        for (let i = 0; i < 3; i++) ctx.fillRect(W / 2 + 42 + i * 4, H - 108 + bobY, 3, 6);
      } else if (g.gun === "rifle") {
        ctx.fillStyle = "#5599cc"; ctx.fillRect(W / 2 + 22, H - 92 + bobY, 92, 58);
        ctx.fillStyle = "#3877aa"; ctx.fillRect(W / 2 + 22, H - 108 + bobY, 16, 24);
        ctx.fillStyle = "#2a5080"; ctx.fillRect(W / 2 + 60, H - 86 + bobY, 18, 22);
        ctx.fillStyle = "#1a3060"; ctx.fillRect(W / 2 + 82, H - 78 + bobY, 10, 8);
        ctx.fillStyle = "#446688"; ctx.fillRect(W / 2 + 100, H - 72 + bobY, 16, 40);
      } else if (g.gun === "laser") {
        ctx.fillStyle = "#cc44cc"; ctx.fillRect(W / 2 + 28, H - 94 + bobY, 86, 52);
        ctx.fillStyle = "#aa00aa"; ctx.fillRect(W / 2 + 28, H - 112 + bobY, 10, 24);
        const grd2 = ctx.createRadialGradient(W / 2 + 33, H - 112 + bobY, 0, W / 2 + 33, H - 112 + bobY, 14);
        grd2.addColorStop(0, "rgba(255,100,255,0.9)"); grd2.addColorStop(1, "rgba(200,0,200,0)");
        ctx.fillStyle = grd2; ctx.fillRect(W / 2 + 20, H - 124 + bobY, 26, 26);
        ctx.fillStyle = "#880088"; ctx.fillRect(W / 2 + 70, H - 88 + bobY, 16, 20);
        ctx.fillStyle = "#550055"; ctx.fillRect(W / 2 + 90, H - 78 + bobY, 12, 10);
      }

      const chColor = g.gun === "laser" ? "rgba(255,80,255,0.9)" : "rgba(255,255,255,0.8)";
      ctx.strokeStyle = chColor; ctx.lineWidth = 1.5;
      ctx.beginPath();
      ctx.moveTo(W / 2 - 12, H / 2); ctx.lineTo(W / 2 + 12, H / 2);
      ctx.moveTo(W / 2, H / 2 - 12); ctx.lineTo(W / 2, H / 2 + 12);
      ctx.stroke();
      ctx.beginPath(); ctx.arc(W / 2, H / 2, 4, 0, Math.PI * 2); ctx.stroke();

      for (const p of g.moneyPopups) {
        ctx.globalAlpha = Math.min(1, p.timer / 30);
        ctx.fillStyle = p.color; ctx.font = "bold 15px monospace"; ctx.textAlign = "center";
        ctx.fillText(p.text, p.x, p.y);
      }
      ctx.globalAlpha = 1;

      if (g.waveFlash > 60) {
        const t = Math.min(1, (g.waveFlash - 60) / 25);
        ctx.globalAlpha = t;
        ctx.fillStyle = "rgba(0,0,0,0.65)"; ctx.fillRect(W / 2 - 220, H / 2 - 46, 440, 82);
        const isBossWave = g.phase === "boss";
        ctx.fillStyle = isBossWave ? "#f84" : "#ffe000";
        ctx.font = "bold 34px monospace"; ctx.textAlign = "center";
        ctx.fillText(isBossWave ? `âš   BOSS â€” WAVE ${g.wave}  âš ` : `â—†  WAVE ${g.wave}  â—†`, W / 2, H / 2 - 6);
        ctx.fillStyle = "#ccc"; ctx.font = "13px monospace";
        ctx.fillText(
          isBossWave ? "A fire demon approaches â€” dodge the fireballs!" : `${g.enemies.length} enemies incoming`,
          W / 2, H / 2 + 22
        );
        ctx.globalAlpha = 1;
      }

      ctx.textAlign = "center";
      if (g.phase === "idle") {
        ctx.fillStyle = "rgba(0,0,0,0.72)"; ctx.fillRect(0, 0, W, H);
        ctx.fillStyle = "#ffe000"; ctx.font = "bold 32px monospace"; ctx.fillText("DUNGEON FPS", W / 2, H / 2 - 50);
        ctx.fillStyle = "#aaa"; ctx.font = "13px monospace";
        ctx.fillText("WASD / â†‘â†“â†â†’ â€” Move    Mouse â€” Turn    Click / SPACE â€” Shoot", W / 2, H / 2 - 8);
        ctx.fillText("Kill enemies â†’ earn money â†’ buy guns â†’ survive boss every 5 waves", W / 2, H / 2 + 16);
        ctx.fillStyle = "#0f0"; ctx.font = "bold 18px monospace";
        ctx.fillText("â–¶  Press SPACE to Start  â—€", W / 2, H / 2 + 58);
      } else if (g.phase === "over") {
        ctx.fillStyle = "rgba(0,0,0,0.78)"; ctx.fillRect(0, 0, W, H);
        ctx.fillStyle = "#f44"; ctx.font = "bold 44px monospace"; ctx.fillText("YOU DIED", W / 2, H / 2 - 20);
        ctx.fillStyle = "#fff"; ctx.font = "18px monospace";
        ctx.fillText(`Wave: ${g.wave}  â€¢  Score: ${Math.floor(g.score)}  â€¢  Kills: ${g.enemyKills}`, W / 2, H / 2 + 24);
        ctx.fillStyle = "#aaa"; ctx.font = "14px monospace"; ctx.fillText("Press SPACE to Restart", W / 2, H / 2 + 60);
      }
    };

    // â”€â”€ Game logic â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
    const isSolid = (x, y) => {
      const tx = Math.floor(x); const ty = Math.floor(y);
      if (tx < 0 || tx >= MAP_W || ty < 0 || ty >= MAP_H) return true;
      return MAP_DATA[ty][tx] > 0;
    };

    const doHitCheck = (g, spreadAngle = 0) => {
      const shotAngle = g.angle + (Math.random() - 0.5) * spreadAngle;
      const allTargets = [
        ...g.enemies,
        ...(g.boss?.alive ? [{ ...g.boss, _isBoss: true }] : []),
      ];
      for (const e of allTargets) {
        if (!e.alive) continue;
        const dx = e.x - g.x; const dy = e.y - g.y;
        const dist = Math.sqrt(dx * dx + dy * dy);
        if (dist > 18) continue;
        const angleToEnemy = Math.atan2(dy, dx);
        let diff = angleToEnemy - shotAngle;
        while (diff >  Math.PI) diff -= Math.PI * 2;
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

    // â”€â”€ Game + render loop â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
    const loop = () => {
      const g = stateRef.current;
      if (!g) { animRef.current = requestAnimationFrame(loop); return; }

      g.frameCount = (g.frameCount || 0) + 1;
      g.moneyPopups = (g.moneyPopups || []).filter(p => p.timer > 0);
      for (const p of g.moneyPopups) { p.timer--; p.y -= 0.5; }

      const running = g.phase === "running" || g.phase === "boss";
      if (running) {
        const cos  = Math.cos(g.angle); const sin  = Math.sin(g.angle);
        const scos = Math.cos(g.angle + Math.PI / 2); const ssin = Math.sin(g.angle + Math.PI / 2);
        let nx = g.x, ny = g.y;
        if (keysRef.current["ArrowUp"]    || keysRef.current["KeyW"]) { nx += cos  * PLAYER_SPEED; ny += sin  * PLAYER_SPEED; }
        if (keysRef.current["ArrowDown"]  || keysRef.current["KeyS"]) { nx -= cos  * PLAYER_SPEED; ny -= sin  * PLAYER_SPEED; }
        if (keysRef.current["ArrowLeft"]  || keysRef.current["KeyA"]) { nx -= scos * PLAYER_SPEED; ny -= ssin * PLAYER_SPEED; }
        if (keysRef.current["ArrowRight"] || keysRef.current["KeyD"]) { nx += scos * PLAYER_SPEED; ny += ssin * PLAYER_SPEED; }
        if (!isSolid(nx, g.y)) g.x = nx;
        if (!isSolid(g.x, ny)) g.y = ny;
        if (g.shootCooldown > 0) g.shootCooldown--;
        if (g.shootFlash  > 0)  g.shootFlash--;
        if (g.hitFlash    > 0)  g.hitFlash--;
        g.bobTime += 0.1;
        g.score   += 0.01;
        if (GUNS[g.gun].auto && mouseDownRef.current) shoot();

        if (g.phase === "running") {
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
            g.phase = "running"; spawnWave(g); if (g.phase === "shop") openShop();
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

      // â”€â”€ Update Three.js scene â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
      // Camera: position at player, look along game angle
      camera.position.set(g.x, 0.5, g.y);
      camera.lookAt(g.x + Math.cos(g.angle), 0.5, g.y + Math.sin(g.angle));

      // Enemies â€” Y-axis billboard rotation to always face the player
      for (let i = 0; i < MAX_ENEMIES; i++) {
        const em = enemyPool[i];
        const e  = g.enemies[i];
        if (e && e.alive) {
          em.group.visible = true;
          em.group.position.set(e.x, 0, e.y);
          em.group.rotation.y = Math.atan2(g.x - e.x, g.y - e.y);
          const hpFrac = Math.max(0, e.health / e.maxHealth);
          em.hpBar.scale.x = hpFrac;
          em.hpBar.position.x = 0.3 * (hpFrac - 1); // shrink from right
          em.bodyMat.color.setHex(e.health > 2 ? 0xcc3333 : e.health > 1 ? 0xcc6633 : 0x996633);
        } else {
          em.group.visible = false;
        }
      }

      // Boss
      if (g.boss?.alive) {
        bossGroup.visible = true;
        bossGroup.position.set(g.boss.x, 0, g.boss.y);
        bossGroup.rotation.y = Math.atan2(g.x - g.boss.x, g.y - g.boss.y);
        const bpFrac = Math.max(0, g.boss.health / g.boss.maxHealth);
        bossHpBar.scale.x = bpFrac;
        bossHpBar.position.x = 0.7 * (bpFrac - 1);
        const pulse = 0.85 + 0.15 * Math.sin(g.frameCount * 0.08);
        bossBodyMat.color.setRGB(0.69 * pulse, 0.16 * pulse, 0);
      } else {
        bossGroup.visible = false;
      }

      // Fireballs
      let fbIdx = 0;
      if (g.boss?.fireballs) {
        for (const f of g.boss.fireballs) {
          if (!f.alive || fbIdx >= MAX_FIREBALLS) continue;
          const fm = fireballPool[fbIdx++];
          fm.visible = true;
          fm.position.set(f.x, 0.5, f.y);
          const t = 0.7 + 0.3 * Math.sin(g.frameCount * 0.3 + f.x * 3);
          fm.material.color.setRGB(1, 0.4 * t, 0);
        }
      }
      for (let i = fbIdx; i < MAX_FIREBALLS; i++) fireballPool[i].visible = false;

      // 2-D overlays
      drawOverlay(g);
      drawMinimap(g);

      // WebGL render
      renderer.render(scene, camera);

      // Throttled HUD update
      if (g.frameCount % 4 === 0) {
        setHudState({
          health: Math.max(0, g.health),
          score:  Math.floor(g.score),
          money:  g.money,
          kills:  g.enemyKills,
          wave:   g.wave,
          gun:    g.gun,
        });
      }

      animRef.current = requestAnimationFrame(loop);
    };

    // â”€â”€ Input handlers â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
    const onKeyDown = (e) => {
      const g = stateRef.current;
      if (g?.phase === "shop") {
        const navKeys = ["ArrowLeft","ArrowRight","ArrowUp","ArrowDown","Enter","Space"];
        if (navKeys.includes(e.code)) e.preventDefault();
        const gunKeys = Object.keys(GUNS);
        const total = gunKeys.length + 1;
        const cur = shopCursorRef.current;
        if (e.code === "ArrowRight" || e.code === "ArrowDown") {
          setShopCursorBoth((cur + 1) % total);
        } else if (e.code === "ArrowLeft" || e.code === "ArrowUp") {
          setShopCursorBoth((cur - 1 + total) % total);
        } else if (e.code === "Enter" || e.code === "Space") {
          if (cur < gunKeys.length) {
            const key = gunKeys[cur];
            const gun = GUNS[key];
            if (g.gun !== key && g.money >= gun.cost) {
              g.money -= gun.cost; g.gun = key; setShopTick(t => t + 1);
            }
          } else {
            startNormalWave(g); setShopTick(t => t + 1);
          }
        }
        return;
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

    const onKeyUp      = (e)  => { keysRef.current[e.code] = false; };
    const onMouseMove  = (e)  => {
      if (!lockedRef.current) return;
      const g = stateRef.current;
      if (g && ["running","boss"].includes(g.phase)) g.angle += e.movementX * MOUSE_SENS;
    };
    const onMouseDown  = (e)  => {
      if (e.button !== 0) return;
      mouseDownRef.current = true;
      if (!lockedRef.current) { renderer.domElement.requestPointerLock(); return; }
      const g = stateRef.current; if (!g) return;
      if (!GUNS[g.gun].auto) shoot();
    };
    const onMouseUp    = ()   => { mouseDownRef.current = false; };
    const onLockChange = ()   => { lockedRef.current = document.pointerLockElement === renderer.domElement; };

    window.addEventListener("keydown", onKeyDown);
    window.addEventListener("keyup",   onKeyUp);
    document.addEventListener("mousemove", onMouseMove);
    renderer.domElement.addEventListener("mousedown", onMouseDown);
    window.addEventListener("mouseup", onMouseUp);
    document.addEventListener("pointerlockchange", onLockChange);
    animRef.current = requestAnimationFrame(loop);

    return () => {
      cancelAnimationFrame(animRef.current);
      window.removeEventListener("keydown", onKeyDown);
      window.removeEventListener("keyup",   onKeyUp);
      document.removeEventListener("mousemove", onMouseMove);
      renderer.domElement.removeEventListener("mousedown", onMouseDown);
      window.removeEventListener("mouseup", onMouseUp);
      document.removeEventListener("pointerlockchange", onLockChange);
      if (document.pointerLockElement === renderer.domElement) document.exitPointerLock();
      renderer.dispose();
      if (mount?.contains(renderer.domElement)) mount.removeChild(renderer.domElement);
    };
  }, [uiPhase]);

  // â”€â”€ Shop helpers â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
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

  // â”€â”€ Render â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
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
            <p>Easy â€” less damage, more money per kill ($40)</p>
            <p>Normal â€” balanced ($25/kill)</p>
            <p>Hard â€” fast enemies, high damage ($15/kill)</p>
            <p className="text-gray-400 mt-2">Boss wave every 5 waves â€¢ 4 guns to unlock</p>
          </div>
        </div>
      )}

      {uiPhase === "game" && (
        <div className="relative flex flex-col items-center gap-3">
          {/* Shop overlay */}
          {shopOpen && (
            <div className="absolute inset-0 flex items-center justify-center z-50 bg-black/75 rounded-xl">
              <div className="bg-gray-900 border-2 border-yellow-400 rounded-2xl p-6 w-[480px] font-mono shadow-2xl">
                <h3 className="text-yellow-300 text-xl font-bold text-center mb-1">ðŸ”« GUN SHOP</h3>
                <p className="text-green-400 text-center text-sm mb-4">
                  Cash: <span className="text-yellow-200 font-bold text-lg">${stateRef.current?.money ?? 0}</span>
                </p>
                <div className="grid grid-cols-2 gap-3 mb-5">
                  {Object.entries(GUNS).map(([key, gun], idx) => {
                    const g = stateRef.current;
                    const owned      = g?.gun === key;
                    const canAfford  = (g?.money ?? 0) >= gun.cost;
                    const highlighted = shopCursor === idx;
                    return (
                      <div key={key}
                        onClick={() => !owned && canAfford && buyGun(key)}
                        className={`rounded-xl border-2 p-3 flex flex-col gap-1 transition-all
                          ${owned ? "border-green-400 bg-green-950" :
                            canAfford ? "border-gray-500 bg-gray-800 hover:border-yellow-400 cursor-pointer" :
                            "border-gray-700 bg-gray-900 opacity-40 cursor-not-allowed"}
                          ${highlighted ? "ring-2 ring-yellow-300 ring-offset-1 ring-offset-gray-900 scale-[1.03]" : ""}
                        `}>
                        <span className="font-bold text-white text-sm">{gun.name}</span>
                        <span className="text-xs text-gray-400">
                          {gun.auto ? "ðŸ”„ Full-Auto" : `ðŸ’¥ ${gun.bulletsPerShot} bullet${gun.bulletsPerShot > 1 ? "s" : ""}/click`}
                        </span>
                        <span className="text-xs text-gray-400">DMG: {gun.dmg} Â· Rate: {Math.round(60 / gun.cooldown)}/s</span>
                        <span className={`text-sm font-bold mt-1 ${gun.cost === 0 ? "text-green-400" : owned ? "text-green-300" : canAfford ? "text-yellow-300" : "text-red-400"}`}>
                          {gun.cost === 0 ? "FREE (default)" : owned ? "âœ“ EQUIPPED" : `$${gun.cost}`}
                        </span>
                      </div>
                    );
                  })}
                </div>
                <button onClick={closeShop}
                  className={`w-full py-3 bg-yellow-400 text-gray-900 font-bold rounded-lg text-lg hover:bg-yellow-300 transition-all
                    ${shopCursor === 4 ? "ring-2 ring-white ring-offset-1 ring-offset-gray-900 scale-[1.02]" : ""}`}>
                  â–¶ Start Wave {stateRef.current?.wave ?? ""}
                </button>
              </div>
            </div>
          )}

          {/* 3-D viewport */}
          <div
            className="relative border-2 border-yellow-400 rounded-xl overflow-hidden cursor-crosshair"
            style={{ width: W, height: H }}
          >
            {/* Three.js WebGL canvas mounts here */}
            <div ref={mountRef} className="absolute inset-0" />

            {/* Overlay: gun sprite, crosshair, hit flash, wave banners */}
            <canvas ref={overlayRef} width={W} height={H}
              className="absolute inset-0 pointer-events-none" />

            {/* Minimap */}
            <canvas ref={minimapRef}
              width={MAP_W * MAP_SCALE} height={MAP_H * MAP_SCALE}
              className="absolute bottom-2 left-2 opacity-75 rounded pointer-events-none" />

            {/* HUD */}
            <div className="absolute top-2 left-3 font-mono text-white text-sm leading-5 pointer-events-none drop-shadow-[0_1px_2px_rgba(0,0,0,0.9)]">
              <div>â¤ {hudState.health}%&nbsp;&nbsp;ðŸ’€ {hudState.kills}&nbsp;&nbsp;W{hudState.wave}&nbsp;&nbsp;[{GUNS[hudState.gun]?.name}]</div>
              <div className="text-yellow-300">${hudState.money}</div>
            </div>
            <div className="absolute top-2 right-3 font-mono text-yellow-400 text-sm pointer-events-none drop-shadow-[0_1px_2px_rgba(0,0,0,0.9)]">
              SCORE {hudState.score}
            </div>
          </div>

          <p className="text-sm text-[#999] font-mono">
            WASD / Arrows Move &amp; Strafe&nbsp;â€¢&nbsp;Mouse Turn&nbsp;â€¢&nbsp;Click / SPACE Shoot&nbsp;â€¢&nbsp;Click canvas to lock mouse
            {shopOpen && <span className="text-yellow-400">&nbsp;|&nbsp;Shop: â†â†’ Navigate&nbsp;â€¢&nbsp;Enter / Space Select</span>}
          </p>
        </div>
      )}
    </div>
  );
}
