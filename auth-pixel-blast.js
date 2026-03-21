import * as THREE from "https://cdn.jsdelivr.net/npm/three@0.183.2/build/three.module.js";

const MAX_CLICKS = 10;

const SHAPE_MAP = Object.freeze({
  square: 0,
  circle: 1,
  triangle: 2,
  diamond: 3,
});

const VERTEX_SRC = `
void main() {
  gl_Position = vec4(position, 1.0);
}
`;

const FRAGMENT_SRC = `
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

float hash11(float n){ return fract(sin(n)*43758.5453); }

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
    sum  += amp * vnoise(p * freq);
    freq *= FBM_LACUNARITY;
    amp  *= FBM_GAIN;
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
  float d  = p.y - r*(1.0 - p.x);
  float aa = fwidth(d);
  return cov * clamp(0.5 - d/aa, 0.0, 1.0);
}

float maskDiamond(vec2 p, float cov){
  float r = sqrt(cov) * 0.564;
  return step(abs(p.x - 0.49) + abs(p.y - 0.49), r);
}

void main(){
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
      vec2 cuv = ((pos - uResolution * .5) / uResolution) * vec2(aspectRatio, 1.0);
      float t = max(uTime - uClickTimes[i], 0.0);
      float r = distance(uv, cuv);
      float waveR = uRippleSpeed * t;
      float ring  = exp(-pow((r - waveR) / uRippleThickness, 2.0));
      float atten = exp(-1.0 * t) * exp(-10.0 * r);
      feed = max(feed, ring * atten * uRippleIntensity);
    }
  }

  float bayer = Bayer8(fragCoord / uPixelSize) - 0.5;
  float bw = step(0.5, feed + bayer);

  float h = fract(sin(dot(floor(fragCoord / uPixelSize), vec2(127.1, 311.7))) * 43758.5453);
  float jitterScale = 1.0 + (h - 0.5) * uPixelJitter;
  float coverage = bw * jitterScale;
  float M;
  if      (uShapeType == SHAPE_CIRCLE)   M = maskCircle (pixelUV, coverage);
  else if (uShapeType == SHAPE_TRIANGLE) M = maskTriangle(pixelUV, pixelId, coverage);
  else if (uShapeType == SHAPE_DIAMOND)  M = maskDiamond(pixelUV, coverage);
  else                                   M = coverage;

  if (uEdgeFade > 0.0) {
    vec2 norm = gl_FragCoord.xy / uResolution;
    float edge = min(min(norm.x, norm.y), min(1.0 - norm.x, 1.0 - norm.y));
    float fade = smoothstep(0.0, uEdgeFade, edge);
    M *= fade;
  }

  vec3 srgbColor = mix(
    uColor * 12.92,
    1.055 * pow(uColor, vec3(1.0 / 2.4)) - 0.055,
    step(0.0031308, uColor)
  );

  fragColor = vec4(srgbColor, M);
}
`;

class AuthPixelBlastBackground {
  constructor(container, options = {}) {
    this.container = container;
    this.options = {
      variant: "square",
      pixelSize: 2,
      color: "#7747e3",
      patternScale: 1,
      patternDensity: 1.4,
      pixelSizeJitter: 1.6,
      enableRipples: true,
      rippleSpeed: 0.4,
      rippleThickness: 0.12,
      rippleIntensityScale: 1.5,
      speed: 0.95,
      edgeFade: 0.31,
      transparent: true,
      interactionTarget: options.interactionTarget || container,
      ...options,
    };
    this.clickIndex = 0;
    this.isPaused = false;
    this.lastPointerRippleAt = 0;

    this.renderer = new THREE.WebGLRenderer({
      alpha: true,
      antialias: true,
      powerPreference: "high-performance",
    });
    this.renderer.setPixelRatio(Math.min(window.devicePixelRatio || 1, 2));
    this.renderer.domElement.style.width = "100%";
    this.renderer.domElement.style.height = "100%";
    this.renderer.domElement.style.display = "block";
    if (this.options.transparent) {
      this.renderer.setClearAlpha(0);
    } else {
      this.renderer.setClearColor(0x000000, 1);
    }
    this.container.appendChild(this.renderer.domElement);

    this.scene = new THREE.Scene();
    this.camera = new THREE.OrthographicCamera(-1, 1, 1, -1, 0, 1);
    this.uniforms = {
      uResolution: { value: new THREE.Vector2(1, 1) },
      uTime: { value: 0 },
      uColor: { value: new THREE.Color(this.options.color) },
      uClickPos: {
        value: Array.from({ length: MAX_CLICKS }, () => new THREE.Vector2(-1, -1)),
      },
      uClickTimes: { value: new Float32Array(MAX_CLICKS) },
      uShapeType: { value: SHAPE_MAP[this.options.variant] ?? 0 },
      uPixelSize: { value: this.options.pixelSize * this.renderer.getPixelRatio() },
      uScale: { value: this.options.patternScale },
      uDensity: { value: this.options.patternDensity },
      uPixelJitter: { value: this.options.pixelSizeJitter },
      uEnableRipples: { value: this.options.enableRipples ? 1 : 0 },
      uRippleSpeed: { value: this.options.rippleSpeed },
      uRippleThickness: { value: this.options.rippleThickness },
      uRippleIntensity: { value: this.options.rippleIntensityScale },
      uEdgeFade: { value: this.options.edgeFade },
    };

    this.material = new THREE.ShaderMaterial({
      vertexShader: VERTEX_SRC,
      fragmentShader: FRAGMENT_SRC,
      uniforms: this.uniforms,
      transparent: true,
      depthTest: false,
      depthWrite: false,
      glslVersion: THREE.GLSL3,
    });
    this.quad = new THREE.Mesh(new THREE.PlaneGeometry(2, 2), this.material);
    this.scene.add(this.quad);

    this.clock = new THREE.Clock();
    this.handleResize = this.handleResize.bind(this);
    this.onPointerDown = this.onPointerDown.bind(this);
    this.onPointerMove = this.onPointerMove.bind(this);
    this.animate = this.animate.bind(this);

    this.resizeObserver = new ResizeObserver(this.handleResize);
    this.resizeObserver.observe(this.container);
    this.handleResize();

    this.options.interactionTarget?.addEventListener("pointerdown", this.onPointerDown, {
      passive: true,
    });
    this.options.interactionTarget?.addEventListener("pointermove", this.onPointerMove, {
      passive: true,
    });

    this.raf = requestAnimationFrame(this.animate);
  }

  handleResize() {
    const width = Math.max(1, this.container.clientWidth || 1);
    const height = Math.max(1, this.container.clientHeight || 1);
    this.renderer.setSize(width, height, false);
    this.uniforms.uResolution.value.set(
      this.renderer.domElement.width,
      this.renderer.domElement.height
    );
    this.uniforms.uPixelSize.value = this.options.pixelSize * this.renderer.getPixelRatio();
  }

  mapPointerEvent(event) {
    const rect = this.container.getBoundingClientRect();
    if (!rect.width || !rect.height) return null;
    const scaleX = this.renderer.domElement.width / rect.width;
    const scaleY = this.renderer.domElement.height / rect.height;
    const x = (event.clientX - rect.left) * scaleX;
    const y = (rect.height - (event.clientY - rect.top)) * scaleY;
    return { x, y };
  }

  addRipple(x, y) {
    const index = this.clickIndex;
    this.uniforms.uClickPos.value[index].set(x, y);
    this.uniforms.uClickTimes.value[index] = this.uniforms.uTime.value;
    this.clickIndex = (index + 1) % MAX_CLICKS;
  }

  onPointerDown(event) {
    const point = this.mapPointerEvent(event);
    if (!point) return;
    this.addRipple(point.x, point.y);
  }

  onPointerMove(event) {
    const point = this.mapPointerEvent(event);
    if (!point) return;
    const now = performance.now();
    if (now - this.lastPointerRippleAt < 110) return;
    this.lastPointerRippleAt = now;
    this.addRipple(point.x, point.y);
  }

  animate() {
    if (this.isPaused) return;
    this.uniforms.uTime.value = this.clock.getElapsedTime() * this.options.speed;
    this.renderer.render(this.scene, this.camera);
    this.raf = requestAnimationFrame(this.animate);
  }

  setPaused(paused) {
    const nextPaused = Boolean(paused);
    if (nextPaused === this.isPaused) return;
    this.isPaused = nextPaused;
    if (nextPaused) {
      cancelAnimationFrame(this.raf);
      return;
    }
    this.clock.getElapsedTime();
    this.raf = requestAnimationFrame(this.animate);
  }

  destroy() {
    cancelAnimationFrame(this.raf);
    this.resizeObserver?.disconnect();
    this.options.interactionTarget?.removeEventListener("pointerdown", this.onPointerDown);
    this.options.interactionTarget?.removeEventListener("pointermove", this.onPointerMove);
    this.quad?.geometry.dispose();
    this.material?.dispose();
    this.renderer?.dispose();
    if (this.renderer?.domElement?.parentElement === this.container) {
      this.container.removeChild(this.renderer.domElement);
    }
  }
}

export function createAuthPixelBlastBackground(container, options = {}) {
  return new AuthPixelBlastBackground(container, options);
}
