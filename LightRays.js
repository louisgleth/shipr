(function () {
  const DEFAULT_COLOR = "#ffffff";

  function hexToRgb(hex) {
    const match = /^#?([a-f\d]{2})([a-f\d]{2})([a-f\d]{2})$/i.exec(hex || DEFAULT_COLOR);
    return match
      ? [
          parseInt(match[1], 16) / 255,
          parseInt(match[2], 16) / 255,
          parseInt(match[3], 16) / 255,
        ]
      : [1, 1, 1];
  }

  function getAnchorAndDir(origin, w, h) {
    const outside = 0.2;
    switch (origin) {
      case "top-left":
        return { anchor: [0, -outside * h], dir: [0, 1] };
      case "top-right":
        return { anchor: [w, -outside * h], dir: [0, 1] };
      case "left":
        return { anchor: [-outside * w, 0.5 * h], dir: [1, 0] };
      case "right":
        return { anchor: [(1 + outside) * w, 0.5 * h], dir: [-1, 0] };
      case "bottom-left":
        return { anchor: [0, (1 + outside) * h], dir: [0, -1] };
      case "bottom-center":
        return { anchor: [0.5 * w, (1 + outside) * h], dir: [0, -1] };
      case "bottom-right":
        return { anchor: [w, (1 + outside) * h], dir: [0, -1] };
      default:
        return { anchor: [0.5 * w, -outside * h], dir: [0, 1] };
    }
  }

  function compileShader(gl, type, source) {
    const shader = gl.createShader(type);
    gl.shaderSource(shader, source);
    gl.compileShader(shader);
    if (!gl.getShaderParameter(shader, gl.COMPILE_STATUS)) {
      const message = gl.getShaderInfoLog(shader) || "Unknown shader compile error";
      gl.deleteShader(shader);
      throw new Error(message);
    }
    return shader;
  }

  function createProgram(gl, vertexSource, fragmentSource) {
    const vertexShader = compileShader(gl, gl.VERTEX_SHADER, vertexSource);
    const fragmentShader = compileShader(gl, gl.FRAGMENT_SHADER, fragmentSource);
    const program = gl.createProgram();
    gl.attachShader(program, vertexShader);
    gl.attachShader(program, fragmentShader);
    gl.linkProgram(program);
    gl.deleteShader(vertexShader);
    gl.deleteShader(fragmentShader);
    if (!gl.getProgramParameter(program, gl.LINK_STATUS)) {
      const message = gl.getProgramInfoLog(program) || "Unknown program link error";
      gl.deleteProgram(program);
      throw new Error(message);
    }
    return program;
  }

  window.LightRays = function LightRays(container, props = {}) {
    if (!container) return () => {};

    const options = {
      raysOrigin: "top-center",
      raysColor: DEFAULT_COLOR,
      raysSpeed: 1,
      lightSpread: 0.5,
      rayLength: 3,
      followMouse: true,
      mouseInfluence: 0.1,
      noiseAmount: 0,
      distortion: 0,
      pulsating: false,
      fadeDistance: 1,
      saturation: 1,
      className: "",
    };
    Object.assign(options, props);

    container.classList.add("light-rays-container");
    if (options.className) container.classList.add(options.className);
    while (container.firstChild) container.removeChild(container.firstChild);

    const canvas = document.createElement("canvas");
    canvas.setAttribute("aria-hidden", "true");
    container.appendChild(canvas);

    const gl = canvas.getContext("webgl", {
      alpha: true,
      antialias: false,
      premultipliedAlpha: false,
      depth: false,
      stencil: false,
    });

    if (!gl) {
      return () => {
        canvas.remove();
      };
    }

    const vertexSource = `
      attribute vec2 position;
      varying vec2 vUv;
      void main() {
        vUv = position * 0.5 + 0.5;
        gl_Position = vec4(position, 0.0, 1.0);
      }
    `;

    const fragmentSource = `
      precision highp float;

      uniform float iTime;
      uniform vec2 iResolution;
      uniform vec2 rayPos;
      uniform vec2 rayDir;
      uniform vec3 raysColor;
      uniform float raysSpeed;
      uniform float lightSpread;
      uniform float rayLength;
      uniform float pulsating;
      uniform float fadeDistance;
      uniform float saturation;
      uniform vec2 mousePos;
      uniform float mouseInfluence;
      uniform float noiseAmount;
      uniform float distortion;

      varying vec2 vUv;

      float noise(vec2 st) {
        return fract(sin(dot(st.xy, vec2(12.9898,78.233))) * 43758.5453123);
      }

      float rayStrength(vec2 raySource, vec2 rayRefDirection, vec2 coord, float seedA, float seedB, float speed) {
        vec2 sourceToCoord = coord - raySource;
        vec2 dirNorm = normalize(sourceToCoord);
        float cosAngle = dot(dirNorm, rayRefDirection);
        float distortedAngle = cosAngle + distortion * sin(iTime * 2.0 + length(sourceToCoord) * 0.01) * 0.2;
        float spreadFactor = pow(max(distortedAngle, 0.0), 1.0 / max(lightSpread, 0.001));
        float distance = length(sourceToCoord);
        float maxDistance = iResolution.x * rayLength;
        float lengthFalloff = clamp((maxDistance - distance) / maxDistance, 0.0, 1.0);
        float fadeFalloff = clamp((iResolution.x * fadeDistance - distance) / (iResolution.x * fadeDistance), 0.5, 1.0);
        float pulse = pulsating > 0.5 ? (0.8 + 0.2 * sin(iTime * speed * 3.0)) : 1.0;
        float baseStrength = clamp(
          (0.45 + 0.15 * sin(distortedAngle * seedA + iTime * speed)) +
          (0.3 + 0.2 * cos(-distortedAngle * seedB + iTime * speed)),
          0.0,
          1.0
        );

        return baseStrength * lengthFalloff * fadeFalloff * spreadFactor * pulse;
      }

      void mainImage(out vec4 fragColor, in vec2 fragCoord) {
        vec2 coord = vec2(fragCoord.x, iResolution.y - fragCoord.y);
        vec2 finalRayDir = rayDir;
        if (mouseInfluence > 0.0) {
          vec2 mouseScreenPos = mousePos * iResolution.xy;
          vec2 mouseDirection = normalize(mouseScreenPos - rayPos);
          finalRayDir = normalize(mix(rayDir, mouseDirection, mouseInfluence));
        }
        vec4 rays1 = vec4(1.0) * rayStrength(rayPos, finalRayDir, coord, 36.2214, 21.11349, 1.5 * raysSpeed);
        vec4 rays2 = vec4(1.0) * rayStrength(rayPos, finalRayDir, coord, 22.3991, 18.0234, 1.1 * raysSpeed);

        fragColor = rays1 * 0.5 + rays2 * 0.4;
        if (noiseAmount > 0.0) {
          float n = noise(coord * 0.01 + iTime * 0.1);
          fragColor.rgb *= (1.0 - noiseAmount + noiseAmount * n);
        }

        float brightness = 1.0 - (coord.y / iResolution.y);
        fragColor.x *= 0.1 + brightness * 0.8;
        fragColor.y *= 0.3 + brightness * 0.6;
        fragColor.z *= 0.5 + brightness * 0.5;
        if (saturation != 1.0) {
          float gray = dot(fragColor.rgb, vec3(0.299, 0.587, 0.114));
          fragColor.rgb = mix(vec3(gray), fragColor.rgb, saturation);
        }

        fragColor.rgb *= raysColor;
      }

      void main() {
        vec4 color;
        mainImage(color, gl_FragCoord.xy);
        gl_FragColor = color;
      }
    `;

    const program = createProgram(gl, vertexSource, fragmentSource);
    const positionLocation = gl.getAttribLocation(program, "position");
    const locations = {
      iTime: gl.getUniformLocation(program, "iTime"),
      iResolution: gl.getUniformLocation(program, "iResolution"),
      rayPos: gl.getUniformLocation(program, "rayPos"),
      rayDir: gl.getUniformLocation(program, "rayDir"),
      raysColor: gl.getUniformLocation(program, "raysColor"),
      raysSpeed: gl.getUniformLocation(program, "raysSpeed"),
      lightSpread: gl.getUniformLocation(program, "lightSpread"),
      rayLength: gl.getUniformLocation(program, "rayLength"),
      pulsating: gl.getUniformLocation(program, "pulsating"),
      fadeDistance: gl.getUniformLocation(program, "fadeDistance"),
      saturation: gl.getUniformLocation(program, "saturation"),
      mousePos: gl.getUniformLocation(program, "mousePos"),
      mouseInfluence: gl.getUniformLocation(program, "mouseInfluence"),
      noiseAmount: gl.getUniformLocation(program, "noiseAmount"),
      distortion: gl.getUniformLocation(program, "distortion"),
    };

    const buffer = gl.createBuffer();
    gl.bindBuffer(gl.ARRAY_BUFFER, buffer);
    gl.bufferData(gl.ARRAY_BUFFER, new Float32Array([-1, -1, 3, -1, -1, 3]), gl.STATIC_DRAW);

    const mouse = { x: 0.5, y: 0.5 };
    const smoothMouse = { x: 0.5, y: 0.5 };
    let animationId = 0;
    let isVisible = true;
    let width = 1;
    let height = 1;

    function applyPlacement() {
      const dpr = Math.min(window.devicePixelRatio || 1, 2);
      const nextWidth = Math.max(1, Math.floor((container.clientWidth || 1) * dpr));
      const nextHeight = Math.max(1, Math.floor((container.clientHeight || 1) * dpr));
      if (canvas.width !== nextWidth || canvas.height !== nextHeight) {
        canvas.width = nextWidth;
        canvas.height = nextHeight;
        gl.viewport(0, 0, nextWidth, nextHeight);
      }
      width = nextWidth;
      height = nextHeight;
      const { anchor, dir } = getAnchorAndDir(options.raysOrigin, width, height);
      gl.useProgram(program);
      gl.uniform2f(locations.iResolution, width, height);
      gl.uniform2f(locations.rayPos, anchor[0], anchor[1]);
      gl.uniform2f(locations.rayDir, dir[0], dir[1]);
    }

    function render(time) {
      animationId = window.requestAnimationFrame(render);
      if (!isVisible) return;

      smoothMouse.x = smoothMouse.x * 0.92 + mouse.x * 0.08;
      smoothMouse.y = smoothMouse.y * 0.92 + mouse.y * 0.08;

      gl.useProgram(program);
      gl.bindBuffer(gl.ARRAY_BUFFER, buffer);
      gl.enableVertexAttribArray(positionLocation);
      gl.vertexAttribPointer(positionLocation, 2, gl.FLOAT, false, 0, 0);
      gl.enable(gl.BLEND);
      gl.blendFunc(gl.SRC_ALPHA, gl.ONE_MINUS_SRC_ALPHA);
      gl.clearColor(0, 0, 0, 0);
      gl.clear(gl.COLOR_BUFFER_BIT);

      const color = hexToRgb(options.raysColor);
      gl.uniform1f(locations.iTime, time * 0.001);
      gl.uniform3f(locations.raysColor, color[0], color[1], color[2]);
      gl.uniform1f(locations.raysSpeed, options.raysSpeed);
      gl.uniform1f(locations.lightSpread, options.lightSpread);
      gl.uniform1f(locations.rayLength, options.rayLength);
      gl.uniform1f(locations.pulsating, options.pulsating ? 1 : 0);
      gl.uniform1f(locations.fadeDistance, options.fadeDistance);
      gl.uniform1f(locations.saturation, options.saturation);
      gl.uniform2f(locations.mousePos, smoothMouse.x, smoothMouse.y);
      gl.uniform1f(locations.mouseInfluence, options.mouseInfluence);
      gl.uniform1f(locations.noiseAmount, options.noiseAmount);
      gl.uniform1f(locations.distortion, options.distortion);
      gl.drawArrays(gl.TRIANGLES, 0, 3);
    }

    function handleMouseMove(event) {
      if (!options.followMouse) return;
      const rect = container.getBoundingClientRect();
      mouse.x = rect.width ? (event.clientX - rect.left) / rect.width : 0.5;
      mouse.y = rect.height ? (event.clientY - rect.top) / rect.height : 0.5;
    }

    const resizeObserver = new ResizeObserver(applyPlacement);
    const visibilityObserver = new IntersectionObserver(
      ([entry]) => {
        isVisible = Boolean(entry?.isIntersecting);
      },
      { threshold: 0.05 }
    );

    resizeObserver.observe(container);
    visibilityObserver.observe(container);
    window.addEventListener("resize", applyPlacement);
    window.addEventListener("mousemove", handleMouseMove);
    applyPlacement();
    animationId = window.requestAnimationFrame(render);

    const api = {
      update(nextProps = {}) {
        Object.assign(options, nextProps);
        applyPlacement();
      },
      destroy() {
        window.cancelAnimationFrame(animationId);
        resizeObserver.disconnect();
        visibilityObserver.disconnect();
        window.removeEventListener("resize", applyPlacement);
        window.removeEventListener("mousemove", handleMouseMove);
        gl.deleteBuffer(buffer);
        gl.deleteProgram(program);
        canvas.remove();
      },
      getOptions() {
        return { ...options };
      },
    };

    return api;
  };
})();
