"use client";

import { ShaderGradient, ShaderGradientCanvas } from "@shadergradient/react";

const shaderProps = {
  animate: "on",
  axesHelper: "on",
  bgColor1: "#000000",
  bgColor2: "#000000",
  brightness: 1.2,
  cAzimuthAngle: 180,
  cDistance: 2.4,
  cPolarAngle: 95,
  cameraZoom: 1,
  color1: "#ffe524",
  color2: "#dd5602",
  color3: "#fdf219",
  destination: "localFile",
  embedMode: "off",
  envPreset: "city",
  format: "gif",
  frameRate: 10,
  gizmoHelper: "hide",
  grain: "off",
  lightType: "3d",
  loop: "on",
  loopDuration: 10,
  positionX: 0,
  positionY: -2.1,
  positionZ: 0,
  range: "enabled",
  rangeEnd: 40,
  rangeStart: 0,
  reflection: 0.1,
  rotationX: 0,
  rotationY: 0,
  rotationZ: 225,
  shader: "defaults",
  toggleAxis: false,
  type: "waterPlane",
  uAmplitude: 0,
  uDensity: 1.8,
  uFrequency: 5.5,
  uSpeed: 0.2,
  uStrength: 3,
  uTime: 7.8,
  wireframe: false,
  zoomOut: false,
} as const;

export function LoginShaderGradient() {
  return (
    <ShaderGradientCanvas
      className="login-shader-canvas"
      fov={45}
      pixelDensity={1}
      pointerEvents="none"
      powerPreference="high-performance"
    >
      <ShaderGradient control="props" {...shaderProps} />
    </ShaderGradientCanvas>
  );
}
