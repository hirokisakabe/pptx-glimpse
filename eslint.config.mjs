import { dirname } from "node:path";
import { fileURLToPath } from "node:url";
import { createNodeResolver, importX } from "eslint-plugin-import-x";
import simpleImportSort from "eslint-plugin-simple-import-sort";
import tseslint from "typescript-eslint";
import prettier from "eslint-config-prettier";

const repoRoot = dirname(fileURLToPath(import.meta.url));
const documentExperimentalSource = fileURLToPath(
  new URL("./packages/pptx-glimpse-document/src/experimental.ts", import.meta.url),
);

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
            "@pptx-glimpse/document/experimental": [documentExperimentalSource],
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
              target: "./packages/pptx-glimpse-document/src",
              from: [
                "./packages/pptx-glimpse/src",
                "./packages/pptx-glimpse-renderer/src",
                "./packages/pptx-glimpse-cli/src",
                "./demo",
                "./scripts",
              ],
              message:
                "@pptx-glimpse/document is the lower-level OOXML/CleanDoc foundation and must not import higher-level packages or app/script code.",
            },
          ],
        },
      ],
      "@typescript-eslint/no-unused-vars": [
        "error",
        { argsIgnorePattern: "^_" },
      ],
      "@typescript-eslint/no-unsafe-argument": "error",
      "@typescript-eslint/no-unsafe-assignment": "error",
      "@typescript-eslint/no-unsafe-call": "error",
      "@typescript-eslint/no-unsafe-member-access": "error",
      "@typescript-eslint/no-unsafe-return": "error",
    },
  },
  {
    files: ["packages/*/src/**/*.ts"],
    rules: {
      "import-x/no-relative-packages": "error",
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
    files: ["packages/pptx-glimpse/src/**/*.ts"],
    ignores: [
      "packages/pptx-glimpse/src/**/*.test.ts",
      "packages/pptx-glimpse/src/**/*.e2e.test.ts",
    ],
    rules: {
      "import-x/no-extraneous-dependencies": [
        "error",
        {
          packageDir: ["packages/pptx-glimpse"],
          devDependencies: false,
          includeInternal: true,
        },
      ],
    },
  },
  {
    files: [
      "packages/pptx-glimpse/src/**/*.test.ts",
      "packages/pptx-glimpse/src/**/*.e2e.test.ts",
    ],
    rules: {
      "import-x/no-extraneous-dependencies": [
        "error",
        {
          packageDir: ["packages/pptx-glimpse", "."],
          devDependencies: true,
          includeInternal: true,
        },
      ],
    },
  },
  {
    files: ["packages/pptx-glimpse-document/src/**/*.ts"],
    ignores: [
      "packages/pptx-glimpse-document/src/**/*.test.ts",
      "packages/pptx-glimpse-document/src/**/*.e2e.test.ts",
    ],
    rules: {
      "import-x/no-extraneous-dependencies": [
        "error",
        {
          packageDir: ["packages/pptx-glimpse-document"],
          devDependencies: false,
          includeInternal: true,
        },
      ],
    },
  },
  {
    files: [
      "packages/pptx-glimpse-document/src/**/*.test.ts",
      "packages/pptx-glimpse-document/src/**/*.e2e.test.ts",
    ],
    rules: {
      "import-x/no-extraneous-dependencies": [
        "error",
        {
          packageDir: ["packages/pptx-glimpse-document", "."],
          devDependencies: true,
          includeInternal: true,
        },
      ],
    },
  },
  {
    files: ["packages/pptx-glimpse-renderer/src/**/*.ts"],
    ignores: [
      "packages/pptx-glimpse-renderer/src/**/*.test.ts",
      "packages/pptx-glimpse-renderer/src/**/*.e2e.test.ts",
    ],
    rules: {
      "import-x/no-extraneous-dependencies": [
        "error",
        {
          packageDir: ["packages/pptx-glimpse-renderer"],
          devDependencies: false,
          includeInternal: true,
        },
      ],
    },
  },
  {
    files: [
      "packages/pptx-glimpse-renderer/src/**/*.test.ts",
      "packages/pptx-glimpse-renderer/src/**/*.e2e.test.ts",
    ],
    rules: {
      "import-x/no-extraneous-dependencies": [
        "error",
        {
          packageDir: ["packages/pptx-glimpse-renderer", "."],
          devDependencies: true,
          includeInternal: true,
        },
      ],
    },
  },
  {
    files: ["packages/pptx-glimpse-cli/src/**/*.ts"],
    ignores: [
      "packages/pptx-glimpse-cli/src/**/*.test.ts",
      "packages/pptx-glimpse-cli/src/**/*.e2e.test.ts",
    ],
    rules: {
      "import-x/no-extraneous-dependencies": [
        "error",
        {
          packageDir: ["packages/pptx-glimpse-cli"],
          devDependencies: false,
          includeInternal: true,
        },
      ],
    },
  },
  {
    files: [
      "packages/pptx-glimpse-cli/src/**/*.test.ts",
      "packages/pptx-glimpse-cli/src/**/*.e2e.test.ts",
    ],
    rules: {
      "import-x/no-extraneous-dependencies": [
        "error",
        {
          packageDir: ["packages/pptx-glimpse-cli", "."],
          devDependencies: true,
          includeInternal: true,
        },
      ],
    },
  },
  prettier,
);
