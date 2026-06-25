import officeAddins from "eslint-plugin-office-addins";
import tsParser from "@typescript-eslint/parser";

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
    },
    languageOptions: {
      parser: tsParser,
    },
  },
];
