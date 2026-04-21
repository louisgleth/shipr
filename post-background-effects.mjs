import * as THREE from "https://cdn.jsdelivr.net/npm/three@0.183.2/build/three.module.js";

const PIXEL_BLAST_MAX_CLICKS = 10;

const PIXEL_BLAST_SHAPE_MAP = Object.freeze({
  square: 0,
  circle: 1,
  triangle: 2,
  diamond: 3,
});

const PARTICLES_VERTEX = `
attribute vec4 random;
attribute vec3 color;

uniform float uTime;
uniform float uSpread;
uniform float uBaseSize;
uniform float uSizeRandomness;

varying vec4 vRandom;
varying vec3 vColor;

void main() {
  vRandom = random;
  vColor = color;

  vec3 pos = position * uSpread;
  pos.z *= 10.0;

  vec4 mPos = modelMatrix * vec4(pos, 1.0);
  float t = uTime;
  mPos.x += sin(t * random.z + 6.28 * random.w) * mix(0.1, 1.5, random.x);
  mPos.y += sin(t * random.y + 6.28 * random.x) * mix(0.1, 1.5, random.w);
  mPos.z += sin(t * random.w + 6.28 * random.y) * mix(0.1, 1.5, random.z);

  vec4 mvPos = viewMatrix * mPos;

  if (uSizeRandomness == 0.0) {
    gl_PointSize = uBaseSize;
  } else {
    gl_PointSize = (uBaseSize * (1.0 + uSizeRandomness * (random.x - 0.5))) / length(mvPos.xyz);
  }

  gl_Position = projectionMatrix * mvPos;
}
`;

const PARTICLES_FRAGMENT = `
precision highp float;

uniform float uTime;
uniform float uAlphaParticles;
varying vec4 vRandom;
varying vec3 vColor;

void main() {
  vec2 uv = gl_PointCoord.xy;
  float d = length(uv - vec2(0.5));

  if (uAlphaParticles < 0.5) {
    if (d > 0.5) {
      discard;
    }
    gl_FragColor = vec4(vColor + 0.2 * sin(uv.yxx + uTime + vRandom.y * 6.28), 1.0);
  } else {
    float circle = smoothstep(0.5, 0.4, d) * 0.8;
    gl_FragColor = vec4(vColor + 0.2 * sin(uv.yxx + uTime + vRandom.y * 6.28), circle);
  }
}
`;

const DITHER_VERTEX = `
void main() {
  gl_Position = vec4(position, 1.0);
}
`;

const DITHER_FRAGMENT = `
precision highp float;

uniform vec2 uResolution;
uniform float uTime;
uniform float uWaveSpeed;
uniform float uWaveFrequency;
uniform float uWaveAmplitude;
uniform vec3 uWaveColor;
uniform vec3 uBackgroundColor;
uniform vec2 uMousePos;
uniform float uMouseActive;
uniform float uMouseRadius;
uniform float uColorNum;
uniform float uPixelSize;

vec4 mod289(vec4 x) { return x - floor(x * (1.0 / 289.0)) * 289.0; }
vec4 permute(vec4 x) { return mod289(((x * 34.0) + 1.0) * x); }
vec4 taylorInvSqrt(vec4 r) { return 1.79284291400159 - 0.85373472095314 * r; }
vec2 fade(vec2 t) { return t * t * t * (t * (t * 6.0 - 15.0) + 10.0); }

float cnoise(vec2 P) {
  vec4 Pi = floor(P.xyxy) + vec4(0.0, 0.0, 1.0, 1.0);
  vec4 Pf = fract(P.xyxy) - vec4(0.0, 0.0, 1.0, 1.0);
  Pi = mod289(Pi);
  vec4 ix = Pi.xzxz;
  vec4 iy = Pi.yyww;
  vec4 fx = Pf.xzxz;
  vec4 fy = Pf.yyww;
  vec4 i = permute(permute(ix) + iy);
  vec4 gx = fract(i * (1.0 / 41.0)) * 2.0 - 1.0;
  vec4 gy = abs(gx) - 0.5;
  vec4 tx = floor(gx + 0.5);
  gx = gx - tx;
  vec2 g00 = vec2(gx.x, gy.x);
  vec2 g10 = vec2(gx.y, gy.y);
  vec2 g01 = vec2(gx.z, gy.z);
  vec2 g11 = vec2(gx.w, gy.w);
  vec4 norm = taylorInvSqrt(vec4(dot(g00, g00), dot(g01, g01), dot(g10, g10), dot(g11, g11)));
  g00 *= norm.x;
  g01 *= norm.y;
  g10 *= norm.z;
  g11 *= norm.w;
  float n00 = dot(g00, vec2(fx.x, fy.x));
  float n10 = dot(g10, vec2(fx.y, fy.y));
  float n01 = dot(g01, vec2(fx.z, fy.z));
  float n11 = dot(g11, vec2(fx.w, fy.w));
  vec2 fadeXY = fade(Pf.xy);
  vec2 nX = mix(vec2(n00, n01), vec2(n10, n11), fadeXY.x);
  return 2.3 * mix(nX.x, nX.y, fadeXY.y);
}

const int OCTAVES = 4;
float fbm(vec2 p) {
  float value = 0.0;
  float amp = 1.0;
  float freq = uWaveFrequency;
  for (int i = 0; i < OCTAVES; i++) {
    value += amp * abs(cnoise(p));
    p *= freq;
    amp *= uWaveAmplitude;
  }
  return value;
}

float pattern(vec2 p) {
  vec2 p2 = p - uTime * uWaveSpeed;
  return fbm(p + fbm(p2));
}

const float bayerMatrix8x8[64] = float[64](
  0.0/64.0, 48.0/64.0, 12.0/64.0, 60.0/64.0,  3.0/64.0, 51.0/64.0, 15.0/64.0, 63.0/64.0,
  32.0/64.0, 16.0/64.0, 44.0/64.0, 28.0/64.0, 35.0/64.0, 19.0/64.0, 47.0/64.0, 31.0/64.0,
  8.0/64.0, 56.0/64.0,  4.0/64.0, 52.0/64.0, 11.0/64.0, 59.0/64.0,  7.0/64.0, 55.0/64.0,
  40.0/64.0, 24.0/64.0, 36.0/64.0, 20.0/64.0, 43.0/64.0, 27.0/64.0, 39.0/64.0, 23.0/64.0,
  2.0/64.0, 50.0/64.0, 14.0/64.0, 62.0/64.0,  1.0/64.0, 49.0/64.0, 13.0/64.0, 61.0/64.0,
  34.0/64.0, 18.0/64.0, 46.0/64.0, 30.0/64.0, 33.0/64.0, 17.0/64.0, 45.0/64.0, 29.0/64.0,
  10.0/64.0, 58.0/64.0,  6.0/64.0, 54.0/64.0,  9.0/64.0, 57.0/64.0,  5.0/64.0, 53.0/64.0,
  42.0/64.0, 26.0/64.0, 38.0/64.0, 22.0/64.0, 41.0/64.0, 25.0/64.0, 37.0/64.0, 21.0/64.0
);

void main() {
  vec2 uv = gl_FragCoord.xy / uResolution.xy;
  vec2 waveUv = uv - 0.5;
  waveUv.x *= uResolution.x / uResolution.y;

  float f = pattern(waveUv);
  if (uMouseActive > 0.5) {
    vec2 mouseNDC = (uMousePos / uResolution - 0.5) * vec2(1.0, -1.0);
    mouseNDC.x *= uResolution.x / uResolution.y;
    float dist = length(waveUv - mouseNDC);
    float effect = 1.0 - smoothstep(0.0, uMouseRadius, dist);
    f -= 0.5 * effect;
  }

  vec3 baseColor = mix(uBackgroundColor, uWaveColor, clamp(f, 0.0, 1.0));

  vec2 normalizedPixelSize = uPixelSize / uResolution;
  vec2 uvPixel = normalizedPixelSize * floor(uv / normalizedPixelSize);
  vec2 scaledCoord = floor(uvPixel * uResolution / uPixelSize);
  int x = int(mod(scaledCoord.x, 8.0));
  int y = int(mod(scaledCoord.y, 8.0));
  float threshold = bayerMatrix8x8[y * 8 + x] - 0.25;
  float stepValue = 1.0 / max(1.0, uColorNum - 1.0);

  vec3 color = baseColor + threshold * stepValue;
  color = clamp(color - 0.2, 0.0, 1.0);
  color = floor(color * max(1.0, uColorNum - 1.0) + 0.5) / max(1.0, uColorNum - 1.0);

  gl_FragColor = vec4(color, 1.0);
}
`;

const PIXEL_BLAST_VERTEX = `
void main() {
  gl_Position = vec4(position, 1.0);
}
`;

const PIXEL_BLAST_FRAGMENT = `
precision highp float;

uniform vec3  uColor;
uniform vec2  uResolution;
uniform float uTime;
uniform float uPixelSize;
uniform float uScale;
uniform float uDensity;
uniform float uPixelJitter;
uniform int   uEnableRipples;
uniform float uRippleSpeed;
uniform float uRippleThickness;
uniform float uRippleIntensity;
uniform float uEdgeFade;

uniform int   uShapeType;
const int SHAPE_SQUARE   = 0;
const int SHAPE_CIRCLE   = 1;
const int SHAPE_TRIANGLE = 2;
const int SHAPE_DIAMOND  = 3;

const int MAX_CLICKS = 10;

uniform vec2  uClickPos[MAX_CLICKS];
uniform float uClickTimes[MAX_CLICKS];

out vec4 fragColor;

float Bayer2(vec2 a) {
  a = floor(a);
  return fract(a.x / 2. + a.y * a.y * .75);
}
#define Bayer4(a) (Bayer2(.5*(a))*0.25 + Bayer2(a))
#define Bayer8(a) (Bayer4(.5*(a))*0.25 + Bayer2(a))

#define FBM_OCTAVES     5
#define FBM_LACUNARITY  1.25
#define FBM_GAIN        1.0

float hash11(float n) { return fract(sin(n) * 43758.5453); }

float vnoise(vec3 p){
  vec3 ip = floor(p);
  vec3 fp = fract(p);
  float n000 = hash11(dot(ip + vec3(0.0,0.0,0.0), vec3(1.0,57.0,113.0)));
  float n100 = hash11(dot(ip + vec3(1.0,0.0,0.0), vec3(1.0,57.0,113.0)));
  float n010 = hash11(dot(ip + vec3(0.0,1.0,0.0), vec3(1.0,57.0,113.0)));
  float n110 = hash11(dot(ip + vec3(1.0,1.0,0.0), vec3(1.0,57.0,113.0)));
  float n001 = hash11(dot(ip + vec3(0.0,0.0,1.0), vec3(1.0,57.0,113.0)));
  float n101 = hash11(dot(ip + vec3(1.0,0.0,1.0), vec3(1.0,57.0,113.0)));
  float n011 = hash11(dot(ip + vec3(0.0,1.0,1.0), vec3(1.0,57.0,113.0)));
  float n111 = hash11(dot(ip + vec3(1.0,1.0,1.0), vec3(1.0,57.0,113.0)));
  vec3 w = fp*fp*fp*(fp*(fp*6.0-15.0)+10.0);
  float x00 = mix(n000, n100, w.x);
  float x10 = mix(n010, n110, w.x);
  float x01 = mix(n001, n101, w.x);
  float x11 = mix(n011, n111, w.x);
  float y0  = mix(x00, x10, w.y);
  float y1  = mix(x01, x11, w.y);
  return mix(y0, y1, w.z) * 2.0 - 1.0;
}

float fbm2(vec2 uv, float t){
  vec3 p = vec3(uv * uScale, t);
  float amp = 1.0;
  float freq = 1.0;
  float sum = 1.0;
  for (int i = 0; i < FBM_OCTAVES; ++i){
    sum += amp * vnoise(p * freq);
    freq *= FBM_LACUNARITY;
    amp *= FBM_GAIN;
  }
  return sum * 0.5 + 0.5;
}

float maskCircle(vec2 p, float cov){
  float r = sqrt(cov) * .25;
  float d = length(p - 0.5) - r;
  float aa = 0.5 * fwidth(d);
  return cov * (1.0 - smoothstep(-aa, aa, d * 2.0));
}

float maskTriangle(vec2 p, vec2 id, float cov){
  bool flip = mod(id.x + id.y, 2.0) > 0.5;
  if (flip) p.x = 1.0 - p.x;
  float r = sqrt(cov);
  float d = p.y - r * (1.0 - p.x);
  float aa = fwidth(d);
  return cov * clamp(0.5 - d / aa, 0.0, 1.0);
}

float maskDiamond(vec2 p, float cov){
  float r = sqrt(cov) * 0.564;
  return step(abs(p.x - 0.49) + abs(p.y - 0.49), r);
}

void main() {
  float pixelSize = uPixelSize;
  vec2 fragCoord = gl_FragCoord.xy - uResolution * .5;
  float aspectRatio = uResolution.x / uResolution.y;

  vec2 pixelId = floor(fragCoord / pixelSize);
  vec2 pixelUV = fract(fragCoord / pixelSize);

  float cellPixelSize = 8.0 * pixelSize;
  vec2 cellId = floor(fragCoord / cellPixelSize);
  vec2 cellCoord = cellId * cellPixelSize;
  vec2 uv = cellCoord / uResolution * vec2(aspectRatio, 1.0);

  float base = fbm2(uv, uTime * 0.05);
  base = base * 0.5 - 0.65;
  float feed = base + (uDensity - 0.5) * 0.3;

  if (uEnableRipples == 1) {
    for (int i = 0; i < MAX_CLICKS; ++i){
      vec2 pos = uClickPos[i];
      if (pos.x < 0.0) continue;
      float localCellPixelSize = 8.0 * pixelSize;
      vec2 cuv = (((pos - uResolution * .5 - localCellPixelSize * .5) / uResolution)) * vec2(aspectRatio, 1.0);
      float t = max(uTime - uClickTimes[i], 0.0);
      float r = distance(uv, cuv);
      float waveR = uRippleSpeed * t;
      float ring = exp(-pow((r - waveR) / uRippleThickness, 2.0));
      float atten = exp(-1.0 * t) * exp(-10.0 * r);
      feed = max(feed, ring * atten * uRippleIntensity);
    }
  }

  float bayer = Bayer8(fragCoord / uPixelSize) - 0.5;
  float bw = step(0.5, feed + bayer);

  float h = fract(sin(dot(floor(fragCoord / uPixelSize), vec2(127.1, 311.7))) * 43758.5453);
  float jitterScale = 1.0 + (h - 0.5) * uPixelJitter;
  float coverage = bw * jitterScale;
  float maskValue;
  if (uShapeType == SHAPE_CIRCLE) maskValue = maskCircle(pixelUV, coverage);
  else if (uShapeType == SHAPE_TRIANGLE) maskValue = maskTriangle(pixelUV, pixelId, coverage);
  else if (uShapeType == SHAPE_DIAMOND) maskValue = maskDiamond(pixelUV, coverage);
  else maskValue = coverage;

  if (uEdgeFade > 0.0) {
    vec2 norm = gl_FragCoord.xy / uResolution;
    float edge = min(min(norm.x, norm.y), min(1.0 - norm.x, 1.0 - norm.y));
    float fade = smoothstep(0.0, uEdgeFade, edge);
    maskValue *= fade;
  }

  vec3 srgbColor = mix(
    uColor * 12.92,
    1.055 * pow(uColor, vec3(1.0 / 2.4)) - 0.055,
    step(0.0031308, uColor)
  );

  fragColor = vec4(srgbColor, maskValue);
}
`;

function clamp01(value) {
  return Math.max(0, Math.min(1, value));
}

function parseCssRgbColor(value) {
  const match = String(value || "")
    .trim()
    .match(/^rgba?\(([^)]+)\)$/i);
  if (!match) return null;
  const parts = match[1].split(/[\s,\/]+/).filter(Boolean);
  if (parts.length < 3) return null;
  return {
    r: clamp01(Number(parts[0]) / 255),
    g: clamp01(Number(parts[1]) / 255),
    b: clamp01(Number(parts[2]) / 255),
  };
}

function parseHexColor(value, fallback = "#ffffff") {
  const cssRgb = parseCssRgbColor(value);
  if (cssRgb) {
    return cssRgb;
  }
  const normalized = String(value || fallback).trim().replace(/^#/, "");
  const hex = normalized.length === 3
    ? normalized.split("").map((part) => `${part}${part}`).join("")
    : normalized;
  const safeHex = hex.length >= 6 ? hex.slice(0, 6) : String(fallback).replace(/^#/, "");
  return {
    r: Number.parseInt(safeHex.slice(0, 2), 16) / 255,
    g: Number.parseInt(safeHex.slice(2, 4), 16) / 255,
    b: Number.parseInt(safeHex.slice(4, 6), 16) / 255,
  };
}

function parseHexColor255(value, fallback = "#ffffff") {
  const color = parseHexColor(value, fallback);
  return {
    r: Math.round(color.r * 255),
    g: Math.round(color.g * 255),
    b: Math.round(color.b * 255),
  };
}

function rgbToCss(color, alpha = 1) {
  const safe = parseHexColor255(color);
  return `rgba(${safe.r}, ${safe.g}, ${safe.b}, ${clamp01(alpha)})`;
}

function mixHexColors(source, target, amount) {
  const from = parseHexColor255(source);
  const to = parseHexColor255(target);
  const t = clamp01(amount);
  return `rgb(${Math.round(from.r + (to.r - from.r) * t)}, ${Math.round(from.g + (to.g - from.g) * t)}, ${Math.round(from.b + (to.b - from.b) * t)})`;
}

function createRandomFloat() {
  if (typeof window !== "undefined" && window.crypto?.getRandomValues) {
    const buffer = new Uint32Array(1);
    window.crypto.getRandomValues(buffer);
    return buffer[0] / 0xffffffff;
  }
  return Math.random();
}

class BaseEffect {
  constructor(width, height, options = {}) {
    this.width = width;
    this.height = height;
    this.options = { ...options };
    this.canvas = document.createElement("canvas");
    this.pointer = { x: 0.5, y: 0.5, active: false };
    this.paused = false;
    this.animationFrame = 0;
    this.lastFrameTime = 0;
  }

  setSize(width, height) {
    this.width = Math.max(1, Math.floor(width));
    this.height = Math.max(1, Math.floor(height));
  }

  update(options = {}) {
    this.options = { ...this.options, ...options };
  }

  setPointer(pointer = this.pointer) {
    this.pointer = {
      x: clamp01(pointer?.x ?? 0.5),
      y: clamp01(pointer?.y ?? 0.5),
      active: Boolean(pointer?.active),
    };
  }

  pointerDown(pointer = this.pointer) {
    this.setPointer({ ...pointer, active: true });
  }

  setPaused(paused) {
    const nextPaused = Boolean(paused);
    if (this.paused === nextPaused) return;
    this.paused = nextPaused;
    if (nextPaused) {
      if (this.animationFrame) cancelAnimationFrame(this.animationFrame);
      this.animationFrame = 0;
      return;
    }
    this.start();
  }

  start() {}
  renderOnce() {}
  dispose() {}
}

class ShapeGridEffect extends BaseEffect {
  constructor(width, height, options = {}) {
    super(width, height, options);
    this.ctx = this.canvas.getContext("2d");
    this.gridOffset = { x: 0, y: 0 };
    this.hoveredSquare = null;
    this.trailCells = [];
    this.cellOpacities = new Map();
    this.setSize(width, height);
    this.start();
  }

  setSize(width, height) {
    super.setSize(width, height);
    this.canvas.width = this.width;
    this.canvas.height = this.height;
  }

  update(options = {}) {
    super.update(options);
  }

  setPointer(pointer = this.pointer) {
    super.setPointer(pointer);
    if (!this.pointer.active) {
      if (this.hoveredSquare && (this.options.hoverTrailAmount || 0) > 0) {
        this.trailCells.unshift({ ...this.hoveredSquare });
        this.trailCells.length = Math.min(this.trailCells.length, this.options.hoverTrailAmount || 0);
      }
      this.hoveredSquare = null;
      return;
    }
    this.updateHoveredCell();
  }

  updateHoveredCell() {
    const squareSize = Math.max(10, Number(this.options.squareSize) || 40);
    const shape = String(this.options.shape || "square");
    const isHex = shape === "hexagon";
    const isTri = shape === "triangle";
    const mouseX = this.pointer.x * this.canvas.width;
    const mouseY = this.pointer.y * this.canvas.height;
    if (isHex) {
      const hexHoriz = squareSize * 1.5;
      const hexVert = squareSize * Math.sqrt(3);
      const colShift = Math.floor(this.gridOffset.x / hexHoriz);
      const offsetX = ((this.gridOffset.x % hexHoriz) + hexHoriz) % hexHoriz;
      const offsetY = ((this.gridOffset.y % hexVert) + hexVert) % hexVert;
      const adjustedX = mouseX - offsetX;
      const adjustedY = mouseY - offsetY;
      const col = Math.round(adjustedX / hexHoriz);
      const rowOffset = (col + colShift) % 2 !== 0 ? hexVert / 2 : 0;
      const row = Math.round((adjustedY - rowOffset) / hexVert);
      this.commitHoveredSquare(col, row);
      return;
    }
    if (isTri) {
      const halfW = squareSize / 2;
      const offsetX = ((this.gridOffset.x % halfW) + halfW) % halfW;
      const offsetY = ((this.gridOffset.y % squareSize) + squareSize) % squareSize;
      const col = Math.round((mouseX - offsetX) / halfW);
      const row = Math.floor((mouseY - offsetY) / squareSize);
      this.commitHoveredSquare(col, row);
      return;
    }
    const offsetX = ((this.gridOffset.x % squareSize) + squareSize) % squareSize;
    const offsetY = ((this.gridOffset.y % squareSize) + squareSize) % squareSize;
    const measure = shape === "circle" ? Math.round : Math.floor;
    this.commitHoveredSquare(measure((mouseX - offsetX) / squareSize), measure((mouseY - offsetY) / squareSize));
  }

  commitHoveredSquare(x, y) {
    if (this.hoveredSquare && this.hoveredSquare.x === x && this.hoveredSquare.y === y) return;
    if (this.hoveredSquare && (this.options.hoverTrailAmount || 0) > 0) {
      this.trailCells.unshift({ ...this.hoveredSquare });
      this.trailCells.length = Math.min(this.trailCells.length, this.options.hoverTrailAmount || 0);
    }
    this.hoveredSquare = { x, y };
  }

  updateCellOpacities() {
    const targets = new Map();
    if (this.hoveredSquare) {
      targets.set(`${this.hoveredSquare.x},${this.hoveredSquare.y}`, 1);
    }
    const trailCount = this.options.hoverTrailAmount || 0;
    if (trailCount > 0) {
      for (let index = 0; index < this.trailCells.length; index += 1) {
        const cell = this.trailCells[index];
        const key = `${cell.x},${cell.y}`;
        if (!targets.has(key)) {
          targets.set(key, (this.trailCells.length - index) / (this.trailCells.length + 1));
        }
      }
    }
    for (const [key] of targets) {
      if (!this.cellOpacities.has(key)) {
        this.cellOpacities.set(key, 0);
      }
    }
    for (const [key, opacity] of this.cellOpacities) {
      const target = targets.get(key) || 0;
      const next = opacity + (target - opacity) * 0.15;
      if (next < 0.005) {
        this.cellOpacities.delete(key);
      } else {
        this.cellOpacities.set(key, next);
      }
    }
  }

  drawShape(cx, cy, size, flip = false) {
    const ctx = this.ctx;
    const shape = String(this.options.shape || "square");
    if (!ctx) return;
    ctx.beginPath();
    if (shape === "hexagon") {
      for (let index = 0; index < 6; index += 1) {
        const angle = (Math.PI / 3) * index;
        const vx = cx + size * Math.cos(angle);
        const vy = cy + size * Math.sin(angle);
        if (index === 0) ctx.moveTo(vx, vy);
        else ctx.lineTo(vx, vy);
      }
      ctx.closePath();
      return;
    }
    if (shape === "circle") {
      ctx.arc(cx, cy, size / 2, 0, Math.PI * 2);
      ctx.closePath();
      return;
    }
    if (shape === "triangle") {
      if (flip) {
        ctx.moveTo(cx, cy + size / 2);
        ctx.lineTo(cx + size / 2, cy - size / 2);
        ctx.lineTo(cx - size / 2, cy - size / 2);
      } else {
        ctx.moveTo(cx, cy - size / 2);
        ctx.lineTo(cx + size / 2, cy + size / 2);
        ctx.lineTo(cx - size / 2, cy + size / 2);
      }
      ctx.closePath();
      return;
    }
    ctx.rect(cx - size / 2, cy - size / 2, size, size);
  }

  drawFrame() {
    const ctx = this.ctx;
    if (!ctx) return;
    const squareSize = Math.max(12, Number(this.options.squareSize) || 40);
    const borderColor = String(this.options.borderColor || "#ffffff");
    const hoverFillColor = String(this.options.hoverFillColor || "#222222");
    const shape = String(this.options.shape || "square");
    const isHex = shape === "hexagon";
    const isTri = shape === "triangle";
    const hexHoriz = squareSize * 1.5;
    const hexVert = squareSize * Math.sqrt(3);

    ctx.clearRect(0, 0, this.canvas.width, this.canvas.height);
    ctx.lineWidth = Math.max(0.5, Number(this.options.lineWidth) || 1);

    if (isHex) {
      const colShift = Math.floor(this.gridOffset.x / hexHoriz);
      const offsetX = ((this.gridOffset.x % hexHoriz) + hexHoriz) % hexHoriz;
      const offsetY = ((this.gridOffset.y % hexVert) + hexVert) % hexVert;
      const cols = Math.ceil(this.canvas.width / hexHoriz) + 3;
      const rows = Math.ceil(this.canvas.height / hexVert) + 3;
      for (let col = -2; col < cols; col += 1) {
        for (let row = -2; row < rows; row += 1) {
          const cx = col * hexHoriz + offsetX;
          const cy = row * hexVert + ((col + colShift) % 2 !== 0 ? hexVert / 2 : 0) + offsetY;
          const key = `${col},${row}`;
          const alpha = this.cellOpacities.get(key);
          if (alpha) {
            ctx.globalAlpha = alpha;
            this.drawShape(cx, cy, squareSize);
            ctx.fillStyle = hoverFillColor;
            ctx.fill();
            ctx.globalAlpha = 1;
          }
          this.drawShape(cx, cy, squareSize);
          ctx.strokeStyle = borderColor;
          ctx.stroke();
        }
      }
    } else if (isTri) {
      const halfW = squareSize / 2;
      const colShift = Math.floor(this.gridOffset.x / halfW);
      const rowShift = Math.floor(this.gridOffset.y / squareSize);
      const offsetX = ((this.gridOffset.x % halfW) + halfW) % halfW;
      const offsetY = ((this.gridOffset.y % squareSize) + squareSize) % squareSize;
      const cols = Math.ceil(this.canvas.width / halfW) + 4;
      const rows = Math.ceil(this.canvas.height / squareSize) + 4;
      for (let col = -2; col < cols; col += 1) {
        for (let row = -2; row < rows; row += 1) {
          const cx = col * halfW + offsetX;
          const cy = row * squareSize + squareSize / 2 + offsetY;
          const flip = ((col + colShift + row + rowShift) % 2 + 2) % 2 !== 0;
          const key = `${col},${row}`;
          const alpha = this.cellOpacities.get(key);
          if (alpha) {
            ctx.globalAlpha = alpha;
            this.drawShape(cx, cy, squareSize, flip);
            ctx.fillStyle = hoverFillColor;
            ctx.fill();
            ctx.globalAlpha = 1;
          }
          this.drawShape(cx, cy, squareSize, flip);
          ctx.strokeStyle = borderColor;
          ctx.stroke();
        }
      }
    } else {
      const offsetX = ((this.gridOffset.x % squareSize) + squareSize) % squareSize;
      const offsetY = ((this.gridOffset.y % squareSize) + squareSize) % squareSize;
      const cols = Math.ceil(this.canvas.width / squareSize) + 3;
      const rows = Math.ceil(this.canvas.height / squareSize) + 3;
      for (let col = -2; col < cols; col += 1) {
        for (let row = -2; row < rows; row += 1) {
          const cx = col * squareSize + squareSize / 2 + offsetX;
          const cy = row * squareSize + squareSize / 2 + offsetY;
          const key = `${col},${row}`;
          const alpha = this.cellOpacities.get(key);
          if (alpha) {
            ctx.globalAlpha = alpha;
            this.drawShape(cx, cy, squareSize);
            ctx.fillStyle = hoverFillColor;
            ctx.fill();
            ctx.globalAlpha = 1;
          }
          this.drawShape(cx, cy, squareSize);
          ctx.strokeStyle = borderColor;
          ctx.stroke();
        }
      }
    }
  }

  start() {
    if (this.animationFrame) cancelAnimationFrame(this.animationFrame);
    this.lastFrameTime = performance.now();
    const tick = () => {
      if (this.paused) return;
      const now = performance.now();
      const delta = Math.min(40, Math.max(0, now - this.lastFrameTime));
      this.lastFrameTime = now;
      const speed = Math.max(Number(this.options.speed) || 0.1, 0.1);
      const direction = String(this.options.direction || "right");
      const squareSize = Math.max(12, Number(this.options.squareSize) || 40);
      const wrapX = this.options.shape === "hexagon" ? squareSize * 3 : squareSize;
      const wrapY = this.options.shape === "hexagon" ? squareSize * Math.sqrt(3) : this.options.shape === "triangle" ? squareSize * 2 : squareSize;
      const step = (delta / 16.6667) * speed;
      if (direction === "right") this.gridOffset.x = (this.gridOffset.x - step + wrapX) % wrapX;
      else if (direction === "left") this.gridOffset.x = (this.gridOffset.x + step + wrapX) % wrapX;
      else if (direction === "up") this.gridOffset.y = (this.gridOffset.y + step + wrapY) % wrapY;
      else if (direction === "down") this.gridOffset.y = (this.gridOffset.y - step + wrapY) % wrapY;
      else {
        this.gridOffset.x = (this.gridOffset.x - step + wrapX) % wrapX;
        this.gridOffset.y = (this.gridOffset.y - step + wrapY) % wrapY;
      }
      this.updateHoveredCell();
      this.updateCellOpacities();
      this.drawFrame();
      this.animationFrame = requestAnimationFrame(tick);
    };
    this.animationFrame = requestAnimationFrame(tick);
  }

  renderOnce() {
    this.updateHoveredCell();
    this.updateCellOpacities();
    this.drawFrame();
  }

  dispose() {
    if (this.animationFrame) cancelAnimationFrame(this.animationFrame);
    this.animationFrame = 0;
  }
}

class ParticlesEffect extends BaseEffect {
  constructor(width, height, options = {}) {
    super(width, height, options);
    this.canvas.width = width;
    this.canvas.height = height;
    this.renderer = new THREE.WebGLRenderer({
      canvas: this.canvas,
      alpha: true,
      antialias: true,
      powerPreference: "high-performance",
      preserveDrawingBuffer: true,
    });
    this.renderer.setClearColor(0x000000, 0);
    this.renderer.setPixelRatio(Math.max(1, Number(options.pixelRatio) || 1));
    this.scene = new THREE.Scene();
    this.camera = new THREE.PerspectiveCamera(15, width / height, 0.1, 1000);
    this.camera.position.set(0, 0, 20);
    this.uniforms = {
      uTime: { value: 0 },
      uSpread: { value: 10 },
      uBaseSize: { value: 100 },
      uSizeRandomness: { value: 1 },
      uAlphaParticles: { value: 0 },
    };
    this.material = new THREE.ShaderMaterial({
      vertexShader: PARTICLES_VERTEX,
      fragmentShader: PARTICLES_FRAGMENT,
      uniforms: this.uniforms,
      transparent: true,
      depthTest: false,
      blending: THREE.NormalBlending,
    });
    this.points = null;
    this.elapsed = 0;
    this.lastFrameTime = performance.now();
    this.setSize(width, height);
    this.update(options);
    this.start();
  }

  buildGeometry() {
    const count = Math.max(20, Math.round(this.options.particleCount || 200));
    const positions = new Float32Array(count * 3);
    const randoms = new Float32Array(count * 4);
    const colors = new Float32Array(count * 3);
    const palette = Array.isArray(this.options.particleColors) && this.options.particleColors.length
      ? this.options.particleColors
      : ["#ffffff", "#ffffff", "#ffffff"];
    for (let index = 0; index < count; index += 1) {
      let x;
      let y;
      let z;
      let len;
      do {
        x = Math.random() * 2 - 1;
        y = Math.random() * 2 - 1;
        z = Math.random() * 2 - 1;
        len = x * x + y * y + z * z;
      } while (len > 1 || len === 0);
      const radius = Math.cbrt(Math.random());
      positions.set([x * radius, y * radius, z * radius], index * 3);
      randoms.set([Math.random(), Math.random(), Math.random(), Math.random()], index * 4);
      const color = parseHexColor(palette[Math.floor(Math.random() * palette.length)], "#ffffff");
      colors.set([color.r, color.g, color.b], index * 3);
    }
    const geometry = new THREE.BufferGeometry();
    geometry.setAttribute("position", new THREE.BufferAttribute(positions, 3));
    geometry.setAttribute("random", new THREE.BufferAttribute(randoms, 4));
    geometry.setAttribute("color", new THREE.BufferAttribute(colors, 3));
    return geometry;
  }

  update(options = {}) {
    const previousCount = this.options.particleCount;
    const previousPalette = JSON.stringify(this.options.particleColors || []);
    super.update(options);
    if (
      !this.points ||
      previousCount !== this.options.particleCount ||
      previousPalette !== JSON.stringify(this.options.particleColors || [])
    ) {
      if (this.points) {
        this.scene.remove(this.points);
        this.points.geometry.dispose();
      }
      this.points = new THREE.Points(this.buildGeometry(), this.material);
      this.scene.add(this.points);
    }
    this.renderer.setPixelRatio(Math.max(1, Number(this.options.pixelRatio) || 1));
    this.uniforms.uSpread.value = Number(this.options.particleSpread || 10);
    this.uniforms.uBaseSize.value =
      Number(this.options.particleBaseSize || 100) * this.renderer.getPixelRatio();
    this.uniforms.uSizeRandomness.value = Number(this.options.sizeRandomness ?? 1);
    this.uniforms.uAlphaParticles.value = this.options.alphaParticles ? 1 : 0;
    this.camera.position.z = Number(this.options.cameraDistance || 20);
  }

  setSize(width, height) {
    super.setSize(width, height);
    this.renderer.setSize(this.width, this.height, false);
    this.camera.aspect = this.width / this.height;
    this.camera.updateProjectionMatrix();
  }

  renderFrame(now) {
    const delta = Math.min(40, Math.max(0, now - this.lastFrameTime));
    this.lastFrameTime = now;
    this.elapsed += delta * Number(this.options.speed || 0.1);
    this.uniforms.uTime.value = this.elapsed * 0.001;
    this.renderScene();
  }

  renderScene() {
    if (this.points) {
      if (this.options.moveParticlesOnHover && this.pointer.active) {
        const pointerX = this.pointer.x * 2 - 1;
        const pointerY = -((this.pointer.y * 2) - 1);
        this.points.position.x = -pointerX * Number(this.options.particleHoverFactor || 1);
        this.points.position.y = -pointerY * Number(this.options.particleHoverFactor || 1);
      } else {
        this.points.position.x = 0;
        this.points.position.y = 0;
      }
      if (!this.options.disableRotation) {
        this.points.rotation.x = Math.sin(this.elapsed * 0.0002) * 0.1;
        this.points.rotation.y = Math.cos(this.elapsed * 0.0005) * 0.15;
        this.points.rotation.z += 0.01 * Number(this.options.speed || 0.1);
      }
    }
    this.renderer.render(this.scene, this.camera);
  }

  start() {
    if (this.animationFrame) cancelAnimationFrame(this.animationFrame);
    this.lastFrameTime = performance.now();
    const tick = (now) => {
      if (this.paused) return;
      this.renderFrame(now);
      this.animationFrame = requestAnimationFrame(tick);
    };
    this.animationFrame = requestAnimationFrame(tick);
  }

  renderOnce() {
    this.renderScene();
  }

  dispose() {
    if (this.animationFrame) cancelAnimationFrame(this.animationFrame);
    if (this.points) {
      this.points.geometry.dispose();
      this.scene.remove(this.points);
    }
    this.material.dispose();
    this.renderer.dispose();
    this.renderer.forceContextLoss();
  }
}

class DitherEffect extends BaseEffect {
  constructor(width, height, options = {}) {
    super(width, height, options);
    this.canvas.width = width;
    this.canvas.height = height;
    this.renderer = new THREE.WebGLRenderer({
      canvas: this.canvas,
      alpha: true,
      antialias: true,
      powerPreference: "high-performance",
      preserveDrawingBuffer: true,
    });
    this.renderer.setClearColor(0x000000, 0);
    this.renderer.setPixelRatio(1);
    this.scene = new THREE.Scene();
    this.camera = new THREE.OrthographicCamera(-1, 1, 1, -1, 0, 1);
    this.uniforms = {
      uResolution: { value: new THREE.Vector2(width, height) },
      uTime: { value: 0 },
      uWaveSpeed: { value: 0.05 },
      uWaveFrequency: { value: 3 },
      uWaveAmplitude: { value: 0.3 },
      uWaveColor: { value: new THREE.Vector3(0.5, 0.5, 0.5) },
      uBackgroundColor: { value: new THREE.Vector3(0, 0, 0) },
      uMousePos: { value: new THREE.Vector2(0, 0) },
      uMouseActive: { value: 0 },
      uMouseRadius: { value: 1 },
      uColorNum: { value: 4 },
      uPixelSize: { value: 2 },
    };
    this.material = new THREE.ShaderMaterial({
      vertexShader: DITHER_VERTEX,
      fragmentShader: DITHER_FRAGMENT,
      uniforms: this.uniforms,
      transparent: true,
      depthTest: false,
      depthWrite: false,
    });
    this.quad = new THREE.Mesh(new THREE.PlaneGeometry(2, 2), this.material);
    this.scene.add(this.quad);
    this.clock = new THREE.Clock();
    this.setSize(width, height);
    this.update(options);
    this.start();
  }

  update(options = {}) {
    super.update(options);
    this.uniforms.uWaveSpeed.value = Number(this.options.waveSpeed || 0.05);
    this.uniforms.uWaveFrequency.value = Number(this.options.waveFrequency || 3);
    this.uniforms.uWaveAmplitude.value = Number(this.options.waveAmplitude || 0.3);
    const waveColor = parseHexColor(this.options.waveColor || "#8ec5ff", "#8ec5ff");
    const backgroundColor = parseHexColor(this.options.backgroundColor || "#000000", "#000000");
    this.uniforms.uWaveColor.value.set(waveColor.r, waveColor.g, waveColor.b);
    this.uniforms.uBackgroundColor.value.set(backgroundColor.r, backgroundColor.g, backgroundColor.b);
    this.uniforms.uMouseRadius.value = Number(this.options.mouseRadius || 1);
    this.uniforms.uColorNum.value = Number(this.options.colorNum || 4);
    this.uniforms.uPixelSize.value = Number(this.options.pixelSize || 2);
  }

  setPointer(pointer = this.pointer) {
    super.setPointer(pointer);
    this.uniforms.uMouseActive.value = this.pointer.active ? 1 : 0;
    this.uniforms.uMousePos.value.set(this.pointer.x * this.width, (1 - this.pointer.y) * this.height);
  }

  setSize(width, height) {
    super.setSize(width, height);
    this.renderer.setSize(this.width, this.height, false);
    this.uniforms.uResolution.value.set(this.width, this.height);
  }

  renderFrame() {
    if (!this.options.disableAnimation) {
      this.uniforms.uTime.value = this.clock.getElapsedTime();
    }
    this.renderer.render(this.scene, this.camera);
  }

  start() {
    if (this.animationFrame) cancelAnimationFrame(this.animationFrame);
    const tick = () => {
      if (this.paused) return;
      this.renderFrame();
      this.animationFrame = requestAnimationFrame(tick);
    };
    this.animationFrame = requestAnimationFrame(tick);
  }

  renderOnce() {
    this.renderer.render(this.scene, this.camera);
  }

  dispose() {
    if (this.animationFrame) cancelAnimationFrame(this.animationFrame);
    this.quad.geometry.dispose();
    this.material.dispose();
    this.renderer.dispose();
    this.renderer.forceContextLoss();
  }
}

class PixelBlastEffect extends BaseEffect {
  constructor(width, height, options = {}) {
    super(width, height, options);
    this.canvas.width = width;
    this.canvas.height = height;
    this.renderer = new THREE.WebGLRenderer({
      canvas: this.canvas,
      antialias: true,
      alpha: true,
      powerPreference: "high-performance",
      preserveDrawingBuffer: true,
    });
    this.renderer.setPixelRatio(Math.max(1, Number(options.pixelRatio) || 1));
    this.renderer.setClearAlpha(0);
    this.scene = new THREE.Scene();
    this.camera = new THREE.OrthographicCamera(-1, 1, 1, -1, 0, 1);
    this.uniforms = {
      uResolution: { value: new THREE.Vector2(width, height) },
      uTime: { value: 0 },
      uColor: { value: new THREE.Color(this.options.color || "#B497CF") },
      uClickPos: {
        value: Array.from({ length: PIXEL_BLAST_MAX_CLICKS }, () => new THREE.Vector2(-1, -1)),
      },
      uClickTimes: { value: new Float32Array(PIXEL_BLAST_MAX_CLICKS) },
      uShapeType: { value: PIXEL_BLAST_SHAPE_MAP.square },
      uPixelSize: { value: 3 },
      uScale: { value: 2 },
      uDensity: { value: 1 },
      uPixelJitter: { value: 0 },
      uEnableRipples: { value: 1 },
      uRippleSpeed: { value: 0.3 },
      uRippleThickness: { value: 0.1 },
      uRippleIntensity: { value: 1 },
      uEdgeFade: { value: 0.25 },
    };
    this.material = new THREE.ShaderMaterial({
      vertexShader: PIXEL_BLAST_VERTEX,
      fragmentShader: PIXEL_BLAST_FRAGMENT,
      uniforms: this.uniforms,
      transparent: true,
      depthTest: false,
      depthWrite: false,
      glslVersion: THREE.GLSL3,
    });
    this.quad = new THREE.Mesh(new THREE.PlaneGeometry(2, 2), this.material);
    this.scene.add(this.quad);
    this.clock = new THREE.Clock();
    this.clickIndex = 0;
    this.timeOffset = createRandomFloat() * 1000;
    this.setSize(width, height);
    this.update(options);
    this.start();
  }

  setSize(width, height) {
    super.setSize(width, height);
    this.renderer.setSize(this.width, this.height, false);
    this.uniforms.uResolution.value.set(this.canvas.width, this.canvas.height);
    this.uniforms.uPixelSize.value = Number(this.options.pixelSize || 3) * this.renderer.getPixelRatio();
  }

  update(options = {}) {
    super.update(options);
    this.renderer.setPixelRatio(Math.max(1, Number(this.options.pixelRatio) || 1));
    this.uniforms.uShapeType.value = PIXEL_BLAST_SHAPE_MAP[this.options.variant] ?? PIXEL_BLAST_SHAPE_MAP.square;
    this.uniforms.uPixelSize.value =
      Number(this.options.pixelSize || 3) * this.renderer.getPixelRatio();
    this.uniforms.uColor.value.set(this.options.color || "#B497CF");
    this.uniforms.uScale.value = Number(this.options.patternScale || 2);
    this.uniforms.uDensity.value = Number(this.options.patternDensity || 1);
    this.uniforms.uPixelJitter.value = Number(this.options.pixelSizeJitter || 0);
    this.uniforms.uEnableRipples.value = this.options.enableRipples === false ? 0 : 1;
    this.uniforms.uRippleSpeed.value = Number(this.options.rippleSpeed || 0.3);
    this.uniforms.uRippleThickness.value = Number(this.options.rippleThickness || 0.1);
    this.uniforms.uRippleIntensity.value = Number(this.options.rippleIntensityScale || 1);
    this.uniforms.uEdgeFade.value = Number(this.options.edgeFade || 0.25);
  }

  pointerDown(pointer = this.pointer) {
    super.pointerDown(pointer);
    const fx = this.pointer.x * this.canvas.width;
    const fy = (1 - this.pointer.y) * this.canvas.height;
    this.uniforms.uClickPos.value[this.clickIndex].set(fx, fy);
    this.uniforms.uClickTimes.value[this.clickIndex] = this.uniforms.uTime.value;
    this.clickIndex = (this.clickIndex + 1) % PIXEL_BLAST_MAX_CLICKS;
  }

  renderFrame() {
    this.uniforms.uTime.value = this.timeOffset + this.clock.getElapsedTime() * Number(this.options.speed || 0.5);
    this.renderer.render(this.scene, this.camera);
  }

  start() {
    if (this.animationFrame) cancelAnimationFrame(this.animationFrame);
    const tick = () => {
      if (this.paused) return;
      this.renderFrame();
      this.animationFrame = requestAnimationFrame(tick);
    };
    this.animationFrame = requestAnimationFrame(tick);
  }

  renderOnce() {
    this.renderer.render(this.scene, this.camera);
  }

  dispose() {
    if (this.animationFrame) cancelAnimationFrame(this.animationFrame);
    this.quad.geometry.dispose();
    this.material.dispose();
    this.renderer.dispose();
    this.renderer.forceContextLoss();
  }
}

export class PostBackgroundRenderer {
  constructor({ width, height }) {
    this.width = Math.max(1, Math.floor(width));
    this.height = Math.max(1, Math.floor(height));
    this.mode = "";
    this.effect = null;
    this.paused = false;
    this.pointer = { x: 0.5, y: 0.5, active: false };
  }

  getCanvas() {
    return this.effect?.canvas || null;
  }

  async setMode(mode, options = {}) {
    const nextMode = String(mode || "").trim();
    if (this.effect && this.mode === nextMode) {
      this.effect.update(options);
      this.effect.setPointer(this.pointer);
      this.effect.setPaused(this.paused);
      return this.effect;
    }
    if (this.effect) {
      this.effect.dispose();
      this.effect = null;
    }
    this.mode = nextMode;
    if (nextMode === "particles") {
      this.effect = new ParticlesEffect(this.width, this.height, options);
    } else if (nextMode === "dither") {
      this.effect = new DitherEffect(this.width, this.height, options);
    } else if (nextMode === "shapeGrid") {
      this.effect = new ShapeGridEffect(this.width, this.height, options);
    } else if (nextMode === "pixelBlast") {
      this.effect = new PixelBlastEffect(this.width, this.height, options);
    } else {
      this.effect = null;
      return null;
    }
    this.effect.setPointer(this.pointer);
    this.effect.setPaused(this.paused);
    return this.effect;
  }

  update(options = {}) {
    this.effect?.update(options);
  }

  resize(width, height) {
    this.width = Math.max(1, Math.floor(width));
    this.height = Math.max(1, Math.floor(height));
    this.effect?.setSize(this.width, this.height);
  }

  setPaused(paused) {
    this.paused = Boolean(paused);
    this.effect?.setPaused(this.paused);
  }

  setPointerNormalized(pointer = this.pointer) {
    this.pointer = {
      x: clamp01(pointer?.x ?? 0.5),
      y: clamp01(pointer?.y ?? 0.5),
      active: Boolean(pointer?.active),
    };
    this.effect?.setPointer(this.pointer);
  }

  pointerDown(pointer = this.pointer) {
    this.setPointerNormalized({ ...pointer, active: true });
    this.effect?.pointerDown(this.pointer);
  }

  renderOnce() {
    this.effect?.renderOnce?.();
  }

  dispose() {
    this.effect?.dispose();
    this.effect = null;
  }
}

export function createPostBackgroundRenderer(options) {
  return new PostBackgroundRenderer(options);
}

export function buildPostBackgroundOptions(mode, settings, palette) {
  const background = palette?.background || "#0f172a";
  const accent = palette?.accent || "#8ec5ff";
  const text = palette?.text || "#f8fafc";
  if (mode === "particles") {
    return {
      particleCount: settings.count,
      particleSpread: Math.max(2, Number(settings.spread || 0.8) * 10),
      speed: settings.speed,
      particleBaseSize: Math.max(28, Number(settings.baseSize || 2.8) * 36),
      moveParticlesOnHover: true,
      particleHoverFactor: 0.8,
      alphaParticles: settings.alpha,
      disableRotation: !settings.rotation,
      sizeRandomness: 1,
      cameraDistance: 20,
      pixelRatio: 1,
      particleColors: [accent, text, mixHexColors(accent, text, 0.35)],
    };
  }
  if (mode === "dither") {
    return {
      waveSpeed: settings.waveSpeed,
      waveFrequency: settings.waveFrequency,
      waveAmplitude: settings.waveAmplitude,
      waveColor: accent,
      backgroundColor: background,
      colorNum: settings.colorNum,
      pixelSize: settings.pixelSize,
      mouseRadius: settings.mouseRadius,
      disableAnimation: false,
    };
  }
  if (mode === "shapeGrid") {
    return {
      direction: settings.direction,
      speed: settings.speed,
      squareSize: settings.squareSize,
      borderColor: rgbToCss(accent, 0.38),
      hoverFillColor: rgbToCss(accent, clamp01(settings.fillStrength)),
      shape: settings.shape,
      hoverTrailAmount: 5,
      lineWidth: settings.lineWidth,
    };
  }
  if (mode === "pixelBlast") {
    return {
      variant: settings.variant,
      pixelSize: settings.pixelSize,
      color: accent,
      patternScale: settings.patternScale,
      patternDensity: settings.density,
      pixelSizeJitter: settings.jitter,
      enableRipples: true,
      rippleSpeed: settings.rippleSpeed,
      rippleThickness: settings.rippleThickness,
      rippleIntensityScale: settings.rippleIntensity,
      speed: 0.5,
      edgeFade: settings.edgeFade,
      pixelRatio: 1,
      transparent: true,
    };
  }
  return {};
}
