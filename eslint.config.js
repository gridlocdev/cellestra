import js from '@eslint/js';
import tseslint from 'typescript-eslint';

/** Packages are used through their entry points (`src/index.ts`), never through their other source files. */
const DEEP_IMPORTS = {
  regex: '^(@cellestra/[^/]+|(\\.\\./)+(packages/)?(core|charts|ui))/src(?!/index\\.ts$)(/|$)',
  message: "Import from the package's entry point (@cellestra/core, @cellestra/ui, @cellestra/charts).",
};

export default tseslint.config(
  { ignores: ['**/dist/**', '**/out/**', '**/release/**', '**/.types/**', '**/node_modules/**'] },
  js.configs.recommended,
  ...tseslint.configs.recommended,
  {
    rules: {
      '@typescript-eslint/no-unused-vars': ['error', { argsIgnorePattern: '^_', varsIgnorePattern: '^_' }],
      '@typescript-eslint/consistent-type-imports': 'error',
      'no-restricted-imports': ['error', { patterns: [DEEP_IMPORTS] }],
    },
  },
  {
    // Build scripts run in Node.
    files: ['**/scripts/**/*.mjs'],
    languageOptions: { globals: { console: 'readonly', process: 'readonly' } },
  },
  {
    // The engine must stay portable: browser, worker, Node, Electron renderer. Its tests run in Node.
    files: ['packages/core/src/**/*.ts', 'packages/charts/src/**/*.ts'],
    rules: {
      'no-restricted-globals': ['error', 'window', 'document', 'navigator', 'localStorage', 'process'],
      'no-restricted-imports': [
        'error',
        { patterns: [DEEP_IMPORTS, { group: ['node:*', 'preact', 'preact/*', '@preact/*'] }] },
      ],
    },
  },
);
