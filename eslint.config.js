import js from '@eslint/js'
import globals from 'globals'
import appsScript from 'eslint-plugin-googleappsscript'
import html from 'eslint-plugin-html'

export default [
  { ignores: ['node_modules/**'] },
  {
    files: ['**/*.{js,cjs,gs,html}'],
    ...js.configs.recommended,
    rules: {
      ...js.configs.recommended.rules,
      semi: ['warn', 'never'],
      'no-unused-vars': ['error', { caughtErrors: 'none' }]
    }
  },
  {
    files: ['src/**/*.js', '**/*.gs'],
    ignores: ['src/test/**'],
    rules: {
      // Apps Script invokes these through menus and the HTML service bridge.
      'no-unused-vars': ['error', {
        caughtErrors: 'none',
        varsIgnorePattern: '^(onOpen|showPasswordDialog|createInventory|setupConfigSheet)$'
      }]
    },
    languageOptions: {
      sourceType: 'script',
      globals: appsScript.environments.googleappsscript.globals
    }
  },
  {
    files: ['**/*.html'],
    plugins: { html },
    languageOptions: {
      sourceType: 'script',
      globals: { ...globals.browser, google: 'readonly' }
    }
  },
  {
    files: ['eslint.config.js', '**/test/**/*.cjs'],
    languageOptions: { globals: globals.node }
  }
]
