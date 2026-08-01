import js from '@eslint/js'
import globals from 'globals'
import reactHooks from 'eslint-plugin-react-hooks'
import reactRefresh from 'eslint-plugin-react-refresh'

// TypeScript 7 no longer exposes the compiler internals consumed by
// typescript-eslint 8.x. TypeScript sources are checked by `tsc -b`; ESLint
// remains the JavaScript/JSX lint pass and deliberately ignores TS files.
export default [
  { ignores: ['**/dist/**', '**/*.ts', '**/*.tsx'] },
  {
    ...js.configs.recommended,
    files: ['**/*.{js,jsx,mjs,cjs}'],
    languageOptions: {
      ecmaVersion: 'latest',
      sourceType: 'module',
      globals: { ...globals.browser, ...globals.node },
    },
    plugins: {
      'react-hooks': reactHooks,
      'react-refresh': reactRefresh,
    },
    rules: {
      ...reactHooks.configs.recommended.rules,
      'react-refresh/only-export-components': [
        'warn',
        { allowConstantExport: true },
      ],
    },
  },
]
