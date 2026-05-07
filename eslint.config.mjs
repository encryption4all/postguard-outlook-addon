import officeAddins from "eslint-plugin-office-addins";
import tsParser from "@typescript-eslint/parser";
import globals from "globals";

export default [
  ...officeAddins.configs.recommended,
  {
    plugins: {
      "office-addins": officeAddins,
    },
    languageOptions: {
      parser: tsParser,
      globals: {
        ...globals.browser,
        Office: "readonly",
        OfficeRuntime: "readonly",
      },
    },
    rules: {
      // TypeScript already checks this via tsc; ESLint's no-undef doesn't
      // understand TS ambient types (RequestInit, BlobPart, etc.) or
      // build-time-replaced identifiers like webpack DefinePlugin's `process`.
      "no-undef": "off",
      "@typescript-eslint/no-unused-vars": [
        "error",
        { argsIgnorePattern: "^_", varsIgnorePattern: "^_", caughtErrorsIgnorePattern: "^_" },
      ],
    },
  },
];
