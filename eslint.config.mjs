import { dirname } from "node:path";
import { fileURLToPath } from "node:url";
import { createNodeResolver, importX } from "eslint-plugin-import-x";
import simpleImportSort from "eslint-plugin-simple-import-sort";
import tseslint from "typescript-eslint";
import prettier from "eslint-config-prettier";

const repoRoot = dirname(fileURLToPath(import.meta.url));
const documentSource = fileURLToPath(new URL("./packages/document/src/index.ts", import.meta.url));
const documentBoundaryRestrictedImportPattern = {
  group: [
    "@pptx-glimpse/renderer",
    "@pptx-glimpse/renderer/*",
    "pptx-glimpse",
    "pptx-glimpse/*",
  ],
  message:
    "@pptx-glimpse/document is the lower-level OOXML/PptxSourceModel foundation and must not import renderer or the public core package.",
};

export default tseslint.config(
  {
    files: [
      "packages/*/src/**/*.ts",
      "vrt/**/*.ts",
      "scripts/**/*.ts",
      "bench/**/*.ts",
      "e2e/**/*.ts",
    ],
    extends: tseslint.configs.recommendedTypeChecked,
    plugins: {
      "import-x": importX,
      "simple-import-sort": simpleImportSort,
    },
    languageOptions: {
      parserOptions: {
        project: "./tsconfig.eslint.json",
        tsconfigRootDir: repoRoot,
      },
    },
    settings: {
      "import-x/extensions": [".ts", ".js"],
      "import-x/resolver-next": [
        createNodeResolver({
          extensions: [".ts", ".js", ".json", ".node"],
          extensionAlias: {
            ".js": [".ts", ".js"],
          },
          alias: {
            "@pptx-glimpse/document": [documentSource],
          },
          mainFields: ["module", "main"],
        }),
      ],
    },
    rules: {
      "simple-import-sort/imports": "error",
      "simple-import-sort/exports": "error",
      "import-x/no-restricted-paths": [
        "error",
        {
          basePath: repoRoot,
          zones: [
            {
              target: "./packages/document/src",
              from: [
                "./packages/core/src",
                "./packages/renderer/src",
                "./packages/cli/src",
                "./demo",
                "./scripts",
              ],
              message:
                "@pptx-glimpse/document is the lower-level OOXML/PptxSourceModel foundation and must not import higher-level packages or app/script code.",
            },
          ],
        },
      ],
      "@typescript-eslint/no-unused-vars": [
        "error",
        { argsIgnorePattern: "^_" },
      ],
      "@typescript-eslint/switch-exhaustiveness-check": [
        "error",
        {
          allowDefaultCaseForExhaustiveSwitch: false,
          considerDefaultExhaustiveForUnions: false,
          requireDefaultForNonUnion: false,
        },
      ],
      "@typescript-eslint/no-unsafe-argument": "error",
      "@typescript-eslint/no-unsafe-assignment": "error",
      "@typescript-eslint/no-unsafe-call": "error",
      "@typescript-eslint/no-unsafe-member-access": "error",
      "@typescript-eslint/no-unsafe-return": "error",
      // Type assertions are allowed at explicit boundaries only. Unsafe
      // narrowing must stay inside boundary helpers or a reasoned local
      // eslint-disable-next-line directive.
      "@typescript-eslint/no-unnecessary-type-assertion": "error",
      "@typescript-eslint/no-unsafe-type-assertion": "error",
      "@typescript-eslint/consistent-type-assertions": [
        "error",
        {
          assertionStyle: "as",
          objectLiteralTypeAssertions: "never",
          arrayLiteralTypeAssertions: "never",
        },
      ],
    },
  },
  {
    files: ["packages/*/src/**/*.ts"],
    rules: {
      "import-x/no-relative-packages": "error",
    },
  },
  {
    files: ["packages/*/src/**/*.ts"],
    ignores: [
      "packages/*/src/**/*.test.ts",
      "packages/*/src/**/*.e2e.test.ts",
    ],
    rules: {
      "no-restricted-imports": [
        "error",
        {
          patterns: [
            {
              group: ["**/unsafe-type-assertion.js"],
              importNames: ["unsafeFixtureAssertion"],
              message:
                "unsafeFixtureAssertion is test-only; production code must use a boundary-specific helper.",
            },
          ],
        },
      ],
    },
  },
  {
    files: ["scripts/**/*.ts", "vrt/**/*.ts", "bench/**/*.ts", "e2e/**/*.ts"],
    rules: {
      "@typescript-eslint/no-unsafe-argument": "off",
      "@typescript-eslint/no-unsafe-assignment": "off",
      "@typescript-eslint/no-unsafe-call": "off",
      "@typescript-eslint/no-unsafe-member-access": "off",
      "@typescript-eslint/no-unsafe-return": "off",
    },
  },
  {
    files: ["packages/core/src/**/*.ts"],
    ignores: [
      "packages/core/src/**/*.test.ts",
      "packages/core/src/**/*.e2e.test.ts",
    ],
    rules: {
      "import-x/no-extraneous-dependencies": [
        "error",
        {
          packageDir: ["packages/core"],
          devDependencies: true,
          includeInternal: true,
        },
      ],
    },
  },
  {
    files: [
      "packages/core/src/**/*.test.ts",
      "packages/core/src/**/*.e2e.test.ts",
    ],
    rules: {
      "import-x/no-extraneous-dependencies": [
        "error",
        {
          packageDir: ["packages/core", "."],
          devDependencies: true,
          includeInternal: true,
        },
      ],
    },
  },
  {
    files: ["packages/document/src/**/*.ts"],
    ignores: [
      "packages/document/src/**/*.test.ts",
      "packages/document/src/**/*.e2e.test.ts",
    ],
    rules: {
      "import-x/no-extraneous-dependencies": [
        "error",
        {
          packageDir: ["packages/document"],
          devDependencies: false,
          includeInternal: true,
        },
      ],
      "no-restricted-imports": [
        "error",
        {
          patterns: [
            documentBoundaryRestrictedImportPattern,
            {
              group: ["**/unsafe-type-assertion.js"],
              importNames: ["unsafeFixtureAssertion"],
              message:
                "unsafeFixtureAssertion is test-only; production code must use a boundary-specific helper.",
            },
          ],
        },
      ],
    },
  },
  {
    files: [
      "packages/document/src/**/*.test.ts",
      "packages/document/src/**/*.e2e.test.ts",
    ],
    rules: {
      "import-x/no-extraneous-dependencies": [
        "error",
        {
          packageDir: ["packages/document", "."],
          devDependencies: true,
          includeInternal: true,
        },
      ],
      "no-restricted-imports": [
        "error",
        {
          patterns: [documentBoundaryRestrictedImportPattern],
        },
      ],
    },
  },
  {
    files: ["packages/renderer/src/**/*.ts"],
    ignores: [
      "packages/renderer/src/**/*.test.ts",
      "packages/renderer/src/**/*.e2e.test.ts",
    ],
    rules: {
      "import-x/no-extraneous-dependencies": [
        "error",
        {
          packageDir: ["packages/renderer"],
          devDependencies: false,
          includeInternal: true,
        },
      ],
    },
  },
  {
    files: [
      "packages/renderer/src/**/*.test.ts",
      "packages/renderer/src/**/*.e2e.test.ts",
    ],
    rules: {
      "import-x/no-extraneous-dependencies": [
        "error",
        {
          packageDir: ["packages/renderer", "."],
          devDependencies: true,
          includeInternal: true,
        },
      ],
    },
  },
  {
    files: ["packages/cli/src/**/*.ts"],
    ignores: [
      "packages/cli/src/**/*.test.ts",
      "packages/cli/src/**/*.e2e.test.ts",
    ],
    rules: {
      "import-x/no-extraneous-dependencies": [
        "error",
        {
          packageDir: ["packages/cli"],
          devDependencies: false,
          includeInternal: true,
        },
      ],
    },
  },
  {
    files: [
      "packages/cli/src/**/*.test.ts",
      "packages/cli/src/**/*.e2e.test.ts",
    ],
    rules: {
      "import-x/no-extraneous-dependencies": [
        "error",
        {
          packageDir: ["packages/cli", "."],
          devDependencies: true,
          includeInternal: true,
        },
      ],
    },
  },
  prettier,
);
