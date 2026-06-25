import officeAddins from "eslint-plugin-office-addins";
import tsParser from "@typescript-eslint/parser";
import licenseHeader from "eslint-plugin-license-header";
import { fileURLToPath } from "url";
import path from "path";

const __dirname = path.dirname(fileURLToPath(import.meta.url));

const recommended = officeAddins.configs.recommended.map((config) => {
  if (config.files?.some((f) => f === "**/*.{js,mjs,cjs,ts,cts,mts}")) {
    return { ...config, files: ["**/*.{js,mjs,cjs,jsx,ts,cts,mts,tsx}"] };
  }
  return config;
});

export default [
  ...recommended,
  {
    plugins: {
      "office-addins": officeAddins,
      "license-header": licenseHeader,
    },
    languageOptions: {
      parser: tsParser,
    },
    rules: {
      "license-header/header": [
        "error",
        path.resolve(__dirname, ".config/source-license-header.js"),
      ],
    },
  },
];
