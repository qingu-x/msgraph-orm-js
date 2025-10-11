import { createDefaultPreset } from "ts-jest";

const tsJestTransformCfg = createDefaultPreset().transform;

/** @type {import("jest").Config} **/
export default {
  preset: 'ts-jest/presets/js-with-ts',
  testEnvironment: 'node',
  transform: {
    ...tsJestTransformCfg,
  },
  transformIgnorePatterns: [
    '/node_modules/?!ky/distribution'
  ]
};