// Minimal Jest setup. Tests live under test/ (outside src/, so the
// src-scoped ESLint/Prettier/tsc CI globs don't pick them up) and are
// transpiled by babel-jest via the existing babel.config.json
// (@babel/preset-typescript + @babel/preset-env). The unit-tested modules
// must stay free of Office.js / pg-js runtime imports so they run in plain
// Node without browser globals.
module.exports = {
  testEnvironment: "node",
  testMatch: ["<rootDir>/test/**/*.test.ts"],
};
